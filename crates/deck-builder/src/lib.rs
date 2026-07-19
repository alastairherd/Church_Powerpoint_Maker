pub mod domain;
pub mod presets;
pub mod sources;
pub mod textproc;

pub use domain::*;

use anyhow::{anyhow, Context};
use async_trait::async_trait;
use once_cell::sync::Lazy;
use pptx_template::{Presentation, Run};
use regex::Regex;
use serde::Deserialize;
use std::collections::BTreeMap;

const TEMPLATE: &[u8] = include_bytes!("../assets/template.pptx");

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Scripture {
    pub reference: String,
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoredSong {
    pub title: String,
    pub slides: Vec<String>,
    pub credits: String,
    pub source_pptx: Option<Vec<u8>>,
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

    async fn song(&self, id: &str, version: u64) -> anyhow::Result<StoredSong> {
        Err(anyhow!("song {id} version {version} is not available"))
    }

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
}

impl LiveSources {
    pub fn new(esv_api_key: impl Into<String>) -> anyhow::Result<Self> {
        Ok(Self {
            esv: sources::esv::EsvClient::new(esv_api_key)?,
        })
    }
}

#[async_trait]
impl Sources for LiveSources {
    async fn scripture(&self, reference: &str) -> anyhow::Result<Scripture> {
        self.esv.passage(reference).await
    }
}

pub async fn build_deck(
    service: &ServiceRecord,
    sources: &(impl Sources + ?Sized),
    ccli_licence_number: &str,
) -> anyhow::Result<Vec<u8>> {
    let mut pres = Presentation::open_bytes(TEMPLATE).context("open embedded TWPC template")?;

    for component in &service.components {
        match component {
            ServiceComponent::Welcome { heading, .. } => {
                add_text_slide(&mut pres, heading, "")?;
            }
            ServiceComponent::Notices { heading, rows, .. } => {
                let pages = paginate_notices(rows, 5);
                for page in pages {
                    add_text_slide(&mut pres, heading, &page)?;
                }
            }
            ServiceComponent::CallToWorship {
                heading,
                reference,
                text,
                ..
            } => {
                let body = if reference.trim().is_empty() {
                    text.clone()
                } else {
                    format!("{}\n\n{}", text.trim(), reference.trim())
                };
                add_text_slide(&mut pres, heading, body.trim())?;
            }
            ServiceComponent::CuePrayer {
                heading, cue, text, ..
            } => {
                let body = [cue.trim(), text.trim()]
                    .into_iter()
                    .filter(|part| !part.is_empty())
                    .collect::<Vec<_>>()
                    .join("\n\n");
                add_text_slide(&mut pres, heading, &body)?;
            }
            ServiceComponent::Song {
                title,
                song,
                lyric_slides,
                credits,
                ..
            } => {
                let stored = if lyric_slides.is_empty() {
                    if let Some(pin) = song {
                        Some(sources.song(&pin.entity_id, pin.version).await?)
                    } else {
                        None
                    }
                } else {
                    None
                };
                let (resolved_title, slides, resolved_credits) = match stored {
                    Some(stored) => (stored.title, stored.slides, stored.credits),
                    None => (title.clone(), lyric_slides.clone(), credits.clone()),
                };
                let slides = if slides.is_empty() {
                    vec!["Song selected in the service editor".to_string()]
                } else {
                    slides
                };
                let slide_count = slides.len();
                for (index, lyrics) in slides.into_iter().enumerate() {
                    let idx = pres.add_slide_from_layout(0).context("add song slide")?;
                    set_text(&mut pres, idx, 0, &resolved_title)?;
                    set_text(&mut pres, idx, 1, &lyrics)?;
                    if index + 1 == slide_count {
                        let footer = if resolved_credits.trim().is_empty() {
                            format!("CCLI: {ccli_licence_number}")
                        } else {
                            format!("{}\nCCLI: {ccli_licence_number}", resolved_credits.trim())
                        };
                        set_text(&mut pres, idx, 2, &footer)?;
                    }
                }
            }
            ServiceComponent::Psalm {
                heading,
                reference,
                slide_breaks,
                tune,
                ..
            } => {
                let (slides, meter) = if slide_breaks.is_empty() && !reference.trim().is_empty() {
                    let psalm = sources.psalm(reference)?;
                    (propose_psalm_groups(&psalm.stanzas), psalm.meter)
                } else {
                    (slide_breaks.clone(), String::new())
                };
                let slides = if slides.is_empty() {
                    vec!["Choose a psalm passage".to_string()]
                } else {
                    slides
                };
                let count = slides.len();
                for (index, stanza) in slides.into_iter().enumerate() {
                    let idx = pres.add_slide_from_layout(3).context("add psalm slide")?;
                    let title = if reference.trim().is_empty() {
                        heading
                    } else {
                        reference
                    };
                    set_text(&mut pres, idx, 0, title)?;
                    set_text(&mut pres, idx, 1, &stanza)?;
                    if index + 1 == count {
                        set_text(
                            &mut pres,
                            idx,
                            2,
                            &format!(
                                "Words: Sing Psalms! © 2003 Free Church of Scotland\nCCLI: {ccli_licence_number}"
                            ),
                        )?;
                        let tune_credit = tune
                            .as_ref()
                            .map(|pin| {
                                format!("Tune catalogue: {} v{}", pin.entity_id, pin.version)
                            })
                            .unwrap_or_else(|| format!("Meter: {meter}"));
                        set_text(&mut pres, idx, 3, &tune_credit)?;
                    }
                }
            }
            ServiceComponent::Reading {
                heading,
                reference,
                bible_page,
                ..
            } => {
                let page = bible_page
                    .map(|page| format!("\nPage {page}"))
                    .unwrap_or_default();
                add_reading_slide(&mut pres, heading, &format!("{reference}{page}"))?;
            }
            ServiceComponent::Teaching {
                heading,
                source,
                selection,
                text,
                ..
            } => {
                let resolved = if text.trim().is_empty()
                    && *source == TeachingSource::WestminsterShorterCatechism
                {
                    selection
                        .parse::<u16>()
                        .ok()
                        .and_then(|number| sources.catechism(number).ok())
                        .map(|item| format!("{}\n\n{}", item.question, item.answer))
                        .unwrap_or_default()
                } else {
                    text.clone()
                };
                add_reading_slide(&mut pres, heading, &resolved)?;
            }
            ServiceComponent::LiturgyBlock {
                heading, key, text, ..
            } => {
                let pages = if text.trim().is_empty() {
                    sources
                        .fixed_component(key)
                        .map(|component| component.slides)
                        .unwrap_or_else(|_| vec![String::new()])
                } else {
                    text.split("\n\n").map(str::to_string).collect()
                };
                for page in pages {
                    add_text_slide(&mut pres, heading, &page)?;
                }
            }
            ServiceComponent::CustomTextImage {
                heading, slides, ..
            } => {
                for page in slides
                    .iter()
                    .map(String::as_str)
                    .chain(slides.is_empty().then_some(""))
                {
                    add_reading_slide(&mut pres, heading, page)?;
                }
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

pub fn paginate_notices(rows: &[NoticeRow], rows_per_slide: usize) -> Vec<String> {
    if rows.is_empty() {
        return vec![String::new()];
    }
    rows.chunks(rows_per_slide.max(1))
        .map(|page| {
            page.iter()
                .map(|row| {
                    let lead = [row.when.trim(), row.title.trim()]
                        .into_iter()
                        .filter(|part| !part.is_empty())
                        .collect::<Vec<_>>()
                        .join(" · ");
                    if row.details.trim().is_empty() {
                        lead
                    } else {
                        format!("{lead}\n{}", row.details.trim())
                    }
                })
                .collect::<Vec<_>>()
                .join("\n\n")
        })
        .collect()
}

pub fn propose_psalm_groups(stanzas: &[String]) -> Vec<String> {
    let mut groups = Vec::new();
    let mut current = String::new();
    for stanza in stanzas {
        let separator = usize::from(!current.is_empty()) * 2;
        if !current.is_empty()
            && (current.lines().count() >= 8 || current.len() + separator + stanza.len() > 620)
        {
            groups.push(current);
            current = String::new();
        }
        if !current.is_empty() {
            current.push_str("\n\n");
        }
        current.push_str(stanza);
    }
    if !current.is_empty() {
        groups.push(current);
    }
    groups
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
        let stanzas = (start..=end)
            .filter_map(|verse| entry.content.body.get(&verse.to_string()))
            .map(|text| textproc::psalm_superscripts(text))
            .collect::<Vec<_>>();
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
        let canonical = if key == "grace" { "the_grace" } else { key };
        let component = COMPONENTS
            .iter()
            .find(|component| component.component == canonical)
            .ok_or_else(|| anyhow!("fixed component not found: {key}"))?;
        let content = component
            .content
            .as_object()
            .ok_or_else(|| anyhow!("fixed component {key} has unsupported content"))?;
        let speaker = content
            .get("speaker")
            .and_then(|value| value.as_str())
            .unwrap_or_default()
            .to_string();
        let slides = (1..=30)
            .filter_map(|i| content.get(&i.to_string()))
            .filter_map(|value| {
                value.as_str().map(str::to_string).or_else(|| {
                    value.as_array().map(|lines| {
                        lines
                            .iter()
                            .filter_map(|line| line.as_str())
                            .collect::<Vec<_>>()
                            .join("\n")
                    })
                })
            })
            .collect();
        Ok(Self {
            key: key.to_string(),
            speaker,
            slides,
        })
    }
}

fn add_text_slide(pres: &mut Presentation, title: &str, text: &str) -> anyhow::Result<()> {
    let idx = pres.add_slide_from_layout(2).context("add text slide")?;
    set_text(pres, idx, 0, title)?;
    set_text(pres, idx, 1, text)
}

fn add_reading_slide(pres: &mut Presentation, title: &str, text: &str) -> anyhow::Result<()> {
    let idx = pres.add_slide_from_layout(4).context("add reading slide")?;
    set_text(pres, idx, 0, title)?;
    set_text(pres, idx, 1, text)
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

#[allow(dead_code)]
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
    serde_json::from_str(include_str!("../assets/psalms.json")).expect("valid embedded psalms.json")
});
static WSC: Lazy<WscFile> = Lazy::new(|| {
    serde_json::from_str(include_str!("../assets/wsc.json")).expect("valid embedded wsc.json")
});
static COMPONENTS: Lazy<Vec<ComponentEntry>> = Lazy::new(|| {
    serde_json::from_str(include_str!("../assets/components.json"))
        .expect("valid embedded components.json")
});

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn complete_notice_rows_paginate_without_dropping_content() {
        let rows: Vec<_> = (1..=12)
            .map(|i| NoticeRow {
                when: format!("Day {i}"),
                title: format!("Notice {i}"),
                details: format!("Details {i}"),
                emphasis: i == 12,
            })
            .collect();
        let pages = paginate_notices(&rows, 5);
        assert_eq!(pages.len(), 3);
        assert!(pages[2].contains("Notice 12"));
    }

    #[test]
    fn psalm_grouping_preserves_every_stanza() {
        let stanzas = vec![
            "a\nb\nc\nd".to_string(),
            "e\nf\ng\nh".to_string(),
            "i".to_string(),
        ];
        let pages = propose_psalm_groups(&stanzas);
        assert_eq!(pages.join("\n\n"), stanzas.join("\n\n"));
    }
}
