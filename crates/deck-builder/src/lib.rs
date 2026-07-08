pub mod sources;
pub mod textproc;

use anyhow::{anyhow, Context};
use async_trait::async_trait;
use chrono::NaiveDate;
use once_cell::sync::Lazy;
use pptx_template::{Presentation, Run};
use regex::Regex;
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

const TEMPLATE: &[u8] = include_bytes!("../../../template.pptx");

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ServiceOrder {
    pub date: NaiveDate,
    pub components: Vec<Component>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum Component {
    Psalm {
        reference: String,
    },
    Hymn {
        url: String,
    },
    Scripture {
        reference: String,
        #[serde(default)]
        title: Option<String>,
    },
    Catechism {
        question: u16,
    },
    Fixed {
        key: String,
        #[serde(default)]
        title: Option<String>,
    },
    Sermon {
        title: String,
        text: String,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Scripture {
    pub reference: String,
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hymn {
    pub title: String,
    pub stanzas: Vec<String>,
    pub author: String,
    pub composer: String,
    pub tune: String,
    pub copyright: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Psalm {
    pub title: String,
    pub meter: String,
    pub stanzas: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Catechism {
    pub number: u16,
    pub question: String,
    pub answer: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FixedComponent {
    pub key: String,
    pub speaker: String,
    pub slides: Vec<String>,
}

#[async_trait]
pub trait Sources: Send + Sync {
    async fn scripture(&self, reference: &str) -> anyhow::Result<Scripture>;
    async fn hymn(&self, url: &str) -> anyhow::Result<Hymn>;

    fn psalm(&self, reference: &str) -> anyhow::Result<Psalm> {
        Psalm::find(reference)
    }

    fn catechism(&self, question: u16) -> anyhow::Result<Catechism> {
        Catechism::find(question)
    }

    fn fixed_component(&self, key: &str) -> anyhow::Result<FixedComponent> {
        FixedComponent::find(key)
    }
}

#[derive(Clone)]
pub struct LiveSources {
    esv: sources::esv::EsvClient,
    hymnary: sources::hymnary::HymnaryClient,
}

impl LiveSources {
    pub fn new(esv_api_key: impl Into<String>) -> anyhow::Result<Self> {
        Ok(Self {
            esv: sources::esv::EsvClient::new(esv_api_key)?,
            hymnary: sources::hymnary::HymnaryClient::new()?,
        })
    }
}

#[async_trait]
impl Sources for LiveSources {
    async fn scripture(&self, reference: &str) -> anyhow::Result<Scripture> {
        self.esv.passage(reference).await
    }

    async fn hymn(&self, url: &str) -> anyhow::Result<Hymn> {
        self.hymnary.hymn(url).await
    }
}

pub async fn build_deck(
    order: &ServiceOrder,
    sources: &(impl Sources + ?Sized),
) -> anyhow::Result<Vec<u8>> {
    let mut pres = Presentation::open_bytes(TEMPLATE).context("open embedded pptx template")?;

    for component in &order.components {
        match component {
            Component::Psalm { reference } => {
                let psalm = sources.psalm(reference)?;
                for stanza in psalm.stanzas {
                    let idx = pres.add_slide_from_layout(3).context("add psalm slide")?;
                    set_text(&mut pres, idx, 0, &psalm.title)?;
                    set_text(&mut pres, idx, 1, &stanza)?;
                    set_text(
                        &mut pres,
                        idx,
                        2,
                        "Words: Sing Psalms! (c) 2003 Free Church of Scotland\nCCLI: 522221",
                    )?;
                    set_text(&mut pres, idx, 3, &format!("Meter: {}", psalm.meter))?;
                }
            }
            Component::Hymn { url } => {
                let hymn = sources.hymn(url).await?;
                for stanza in hymn.stanzas {
                    let idx = pres.add_slide_from_layout(0).context("add hymn slide")?;
                    set_text(&mut pres, idx, 0, &hymn.title)?;
                    set_text(&mut pres, idx, 1, &stanza)?;
                    set_text(
                        &mut pres,
                        idx,
                        2,
                        &format!(
                            "Words: {}\nComposer: {}\nTune: {}\n(c): {}\nCCLI: 522221",
                            hymn.author, hymn.composer, hymn.tune, hymn.copyright
                        ),
                    )?;
                }
            }
            Component::Scripture { reference, title } => {
                let scripture = sources.scripture(reference).await?;
                let title = title.as_deref().unwrap_or(&scripture.reference);
                for page in
                    textproc::split_lines(&textproc::british_spellings(&scripture.text), 14, 900)
                {
                    let idx = pres
                        .add_slide_from_layout(4)
                        .context("add scripture slide")?;
                    set_text(&mut pres, idx, 0, title)?;
                    set_rich_text(&mut pres, idx, 1, &textproc::scripture_runs(&page))?;
                }
            }
            Component::Catechism { question } => {
                let catechism = sources.catechism(*question)?;
                let idx = pres
                    .add_slide_from_layout(5)
                    .context("add catechism slide")?;
                set_text(
                    &mut pres,
                    idx,
                    0,
                    &format!("Westminster Shorter Catechism {}", catechism.number),
                )?;
                set_text(
                    &mut pres,
                    idx,
                    1,
                    &format!("{}\n\n{}", catechism.question, catechism.answer),
                )?;
            }
            Component::Fixed { key, title } => {
                let fixed = sources.fixed_component(key)?;
                let title = title.as_deref().unwrap_or(key);
                for body in fixed.slides {
                    let idx = pres
                        .add_slide_from_layout(2)
                        .context("add fixed component slide")?;
                    set_text(&mut pres, idx, 0, title)?;
                    set_rich_text(
                        &mut pres,
                        idx,
                        1,
                        &[
                            Run {
                                text: format!("{} ", fixed.speaker),
                                superscript: false,
                                bold: true,
                                italic: false,
                            },
                            Run::plain(body),
                        ],
                    )?;
                }
            }
            Component::Sermon { title, text } => {
                let idx = pres.add_slide_from_layout(4).context("add sermon slide")?;
                set_text(&mut pres, idx, 0, title)?;
                set_text(&mut pres, idx, 1, text)?;
            }
        }
    }

    if pres.slide_count() > 1 {
        pres.delete_slide(0)
            .context("remove blank template slide")?;
    }

    pres.validate().context("validate generated pptx")?;
    pres.save_bytes().context("save generated pptx")
}

impl Psalm {
    pub fn find(reference: &str) -> anyhow::Result<Self> {
        static RE: Lazy<Regex> = Lazy::new(|| {
            Regex::new(r"^(?:Psalm )?(\d{1,3})(?::([1-9]\d{0,2})-([1-9]\d{0,2}))?(?:\s\(([a-zA-Z])\))?(?:\s\((\d{1,2})\))?$")
                .expect("valid psalm reference regex")
        });

        let caps = RE
            .captures(reference)
            .ok_or_else(|| anyhow!("invalid psalm reference: {reference}"))?;
        let number = caps.get(1).expect("psalm number").as_str();
        let version = caps.get(4).map(|m| m.as_str()).unwrap_or("a");
        let section = caps.get(5).map(|m| m.as_str());
        let start = caps
            .get(2)
            .and_then(|m| m.as_str().parse::<u16>().ok())
            .unwrap_or(1);
        let end = caps
            .get(3)
            .and_then(|m| m.as_str().parse::<u16>().ok())
            .unwrap_or(if section.is_some() { 300 } else { 30 });

        let entry = PSALMS
            .iter()
            .find(|entry| {
                entry.psalm == number
                    && section
                        .map(|wanted| entry.content.section == wanted)
                        .unwrap_or(entry.content.version == version)
            })
            .ok_or_else(|| anyhow!("psalm not found: {reference}"))?;

        let mut stanzas = Vec::new();
        for verse in start..=end {
            if let Some(text) = entry.content.body.get(&verse.to_string()) {
                stanzas.push(textproc::psalm_superscripts(text));
            }
        }

        if stanzas.is_empty() {
            return Err(anyhow!("no stanzas found for {reference}"));
        }

        Ok(Self {
            title: reference.to_string(),
            meter: entry.content.meter.clone(),
            stanzas,
        })
    }
}

impl Catechism {
    pub fn find(question: u16) -> anyhow::Result<Self> {
        let item = WSC
            .data
            .iter()
            .find(|item| item.number == question)
            .ok_or_else(|| anyhow!("catechism question not found: {question}"))?;
        Ok(Self {
            number: item.number,
            question: item.question.clone(),
            answer: item.answer.clone(),
        })
    }
}

impl FixedComponent {
    pub fn find(key: &str) -> anyhow::Result<Self> {
        let component = COMPONENTS
            .iter()
            .find(|component| component.component == key)
            .ok_or_else(|| anyhow!("fixed component not found: {key}"))?;
        let content = component
            .content
            .as_object()
            .ok_or_else(|| anyhow!("fixed component {key} has unsupported content"))?;
        let speaker = content
            .get("speaker")
            .and_then(|value| value.as_str())
            .ok_or_else(|| anyhow!("fixed component {key} missing speaker"))?
            .to_string();
        let mut slides = Vec::new();
        for i in 1..=20 {
            if let Some(value) = content.get(&i.to_string()) {
                if let Some(text) = value.as_str() {
                    slides.push(text.to_string());
                } else if let Some(lines) = value.as_array() {
                    slides.push(
                        lines
                            .iter()
                            .filter_map(|line| line.as_str())
                            .collect::<Vec<_>>()
                            .join("\n"),
                    );
                }
            }
        }
        Ok(Self {
            key: key.to_string(),
            speaker,
            slides,
        })
    }
}

fn set_text(
    pres: &mut Presentation,
    slide: usize,
    placeholder: usize,
    text: &str,
) -> anyhow::Result<()> {
    pres.slide_mut(slide)?
        .placeholder(placeholder)?
        .set_text(text)?;
    Ok(())
}

fn set_rich_text(
    pres: &mut Presentation,
    slide: usize,
    placeholder: usize,
    runs: &[Run],
) -> anyhow::Result<()> {
    pres.slide_mut(slide)?
        .placeholder(placeholder)?
        .set_rich_text(runs)?;
    Ok(())
}

#[derive(Debug, Deserialize)]
struct PsalmEntry {
    #[serde(rename = "Psalm")]
    psalm: String,
    #[serde(rename = "Content")]
    content: PsalmContent,
}

#[derive(Debug, Deserialize)]
struct PsalmContent {
    #[serde(rename = "Version")]
    version: String,
    #[serde(rename = "Section")]
    section: String,
    #[serde(rename = "Meter")]
    meter: String,
    #[serde(rename = "Body")]
    body: BTreeMap<String, String>,
}

#[derive(Debug, Deserialize)]
struct WscFile {
    #[serde(rename = "Data")]
    data: Vec<WscItem>,
}

#[derive(Debug, Deserialize)]
struct WscItem {
    #[serde(rename = "Number")]
    number: u16,
    #[serde(rename = "Question")]
    question: String,
    #[serde(rename = "Answer")]
    answer: String,
}

#[derive(Debug, Deserialize)]
struct ComponentEntry {
    #[serde(rename = "Component")]
    component: String,
    #[serde(rename = "Content")]
    content: serde_json::Value,
}

static PSALMS: Lazy<Vec<PsalmEntry>> = Lazy::new(|| {
    serde_json::from_str(include_str!("../../../psalms.json")).expect("valid embedded psalms.json")
});

static WSC: Lazy<WscFile> = Lazy::new(|| {
    serde_json::from_str(include_str!("../../../wsc.json")).expect("valid embedded wsc.json")
});

static COMPONENTS: Lazy<Vec<ComponentEntry>> = Lazy::new(|| {
    serde_json::from_str(include_str!("../../../components.json"))
        .expect("valid embedded components.json")
});
