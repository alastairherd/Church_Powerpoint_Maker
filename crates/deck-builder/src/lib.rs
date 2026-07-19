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
const SEED_WELCOME: usize = 0;
const SEED_NOTICES: usize = 1;
const SEED_CALL_TO_WORSHIP: usize = 2;
const SEED_PRAYER: usize = 8;
const SEED_READING: usize = 15;
const SEED_PSALM: usize = 16;
const SEED_TEACHING: usize = 19;
const SEED_LITURGY: usize = 23;
const SEED_SONG: usize = 39;
const SEED_SONG_FINAL: usize = 46;
const SEED_REFRESHMENTS: usize = 48;

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
    let seed_count = pres.slide_count();

    for component in &service.components {
        match component {
            ServiceComponent::Welcome { heading, .. } => {
                let slide = clone_seed(&mut pres, SEED_WELCOME, "welcome")?;
                set_shape_text(&mut pres, slide, "TextShape 1", heading)?;
            }
            ServiceComponent::Notices { heading, rows, .. } => {
                let pages = paginate_notices(rows, 5);
                for page in pages {
                    let slide = clone_seed(&mut pres, SEED_NOTICES, "notices")?;
                    set_shape_text(&mut pres, slide, "Title 1", heading)?;
                    set_shape_text(&mut pres, slide, "TextBox 2", &page)?;
                }
            }
            ServiceComponent::CallToWorship {
                heading,
                reference,
                text,
                ..
            } => {
                let slide = clone_seed(&mut pres, SEED_CALL_TO_WORSHIP, "call to worship")?;
                set_shape_text(&mut pres, slide, "Title 1", heading)?;
                let mut runs = textproc::scripture_runs(text.trim());
                if !reference.trim().is_empty() {
                    let separator = if text.trim().is_empty() { "" } else { "\n\n" };
                    runs.push(Run::plain(format!("{separator}{}", reference.trim())));
                }
                set_shape_runs(&mut pres, slide, "Text Placeholder 2", &runs)?;
            }
            ServiceComponent::CuePrayer {
                heading, cue, text, ..
            } => {
                let body = [cue.trim(), text.trim()]
                    .into_iter()
                    .filter(|part| !part.is_empty())
                    .collect::<Vec<_>>()
                    .join("\n\n");
                if body.is_empty() {
                    let (seed, shape) = if heading == "Join us for refreshments" {
                        (SEED_REFRESHMENTS, "TextShape 1")
                    } else {
                        (SEED_PRAYER, "Title 1")
                    };
                    let slide = clone_seed(&mut pres, seed, "prayer cue")?;
                    set_shape_text(&mut pres, slide, shape, heading)?;
                } else {
                    let slide = clone_seed(&mut pres, SEED_LITURGY, "prayer text")?;
                    set_shape_text(&mut pres, slide, "TextShape 1", heading)?;
                    set_shape_text(&mut pres, slide, "TextShape 2", &body)?;
                }
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
                let (resolved_title, slides, resolved_credits, source_pptx) = match stored {
                    Some(stored) => (
                        stored.title,
                        stored.slides,
                        stored.credits,
                        stored.source_pptx,
                    ),
                    None => (title.clone(), lyric_slides.clone(), credits.clone(), None),
                };
                if let Some(source_pptx) = source_pptx {
                    pres.import_slides(&source_pptx)
                        .context("import original song slides")?;
                } else {
                    let slides = if slides.is_empty() {
                        vec!["Song selected in the service editor".to_string()]
                    } else {
                        slides
                    };
                    let slide_count = slides.len();
                    for (index, lyrics) in slides.into_iter().enumerate() {
                        let is_final = index + 1 == slide_count;
                        let seed = if is_final { SEED_SONG_FINAL } else { SEED_SONG };
                        let slide = clone_seed(&mut pres, seed, "lyric slide")?;
                        set_shape_text(&mut pres, slide, "TextShape 1", &resolved_title)?;
                        set_shape_text(&mut pres, slide, "TextBox 1", &lyrics)?;
                        if is_final {
                            let footer = if resolved_credits.trim().is_empty() {
                                format!("CCLI: {ccli_licence_number}")
                            } else {
                                format!("{}\nCCLI: {ccli_licence_number}", resolved_credits.trim())
                            };
                            set_shape_text(&mut pres, slide, "TextBox 4", &footer)?;
                        }
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
                let (slides, meter) = if !reference.trim().is_empty() {
                    let psalm = sources.psalm(reference)?;
                    let slides = if slide_breaks.is_empty() {
                        propose_psalm_groups(&psalm.stanzas)
                    } else {
                        slide_breaks.clone()
                    };
                    (slides, psalm.meter)
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
                    let slide = clone_seed(&mut pres, SEED_PSALM, "psalm")?;
                    let title = if reference.trim().is_empty() {
                        heading
                    } else {
                        reference
                    };
                    set_shape_text(&mut pres, slide, "TextShape 1", title)?;
                    set_shape_runs(&mut pres, slide, "TextShape 2", &psalm_runs(&stanza))?;
                    if index + 1 == count {
                        pres.copy_shape(SEED_SONG_FINAL, "TextBox 4", slide, "Psalm Credits")?;
                        let tune_credit = tune
                            .as_ref()
                            .map(|pin| {
                                format!("Tune catalogue: {} v{}", pin.entity_id, pin.version)
                            })
                            .unwrap_or_else(|| format!("Meter: {meter}"));
                        set_shape_text(
                            &mut pres,
                            slide,
                            "Psalm Credits",
                            &format!(
                                "Words: Sing Psalms! © 2003\nFree Church of Scotland\n{tune_credit}\nCCLI: {ccli_licence_number}"
                            ),
                        )?;
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
                let slide = clone_seed(&mut pres, SEED_READING, "reading")?;
                set_shape_text(&mut pres, slide, "TextShape 1", heading)?;
                set_shape_text(
                    &mut pres,
                    slide,
                    "TextShape 3",
                    &format!("{reference}{page}"),
                )?;
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
                let slide = clone_seed(&mut pres, SEED_TEACHING, "teaching")?;
                set_shape_text(&mut pres, slide, "TextShape 1", heading)?;
                set_shape_text(&mut pres, slide, "TextShape 3", &resolved)?;
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
                    let slide = clone_seed(&mut pres, SEED_LITURGY, "liturgy")?;
                    set_shape_text(&mut pres, slide, "TextShape 1", heading)?;
                    set_shape_text(&mut pres, slide, "TextShape 2", &page)?;
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
                    let slide = clone_seed(&mut pres, SEED_READING, "custom text")?;
                    set_shape_text(&mut pres, slide, "TextShape 1", heading)?;
                    set_shape_text(&mut pres, slide, "TextShape 3", page)?;
                }
            }
        }
    }

    for _ in 0..seed_count {
        pres.delete_slide(0)
            .context("remove canonical template seed slide")?;
    }

    pres.remove_auxiliary_content()
        .context("remove notes and non-visible relationships")?;
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
    const MAX_RENDERED_LINES: usize = 8;
    const MAX_CHARACTERS: usize = 360;

    let mut groups = Vec::new();
    let mut current = String::new();
    for stanza in stanzas {
        let separator = usize::from(!current.is_empty()) * 2;
        let rendered_lines = current
            .lines()
            .filter(|line| !line.trim().is_empty())
            .count()
            + stanza
                .lines()
                .filter(|line| !line.trim().is_empty())
                .count();
        if !current.is_empty()
            && (rendered_lines > MAX_RENDERED_LINES
                || current.len() + separator + stanza.len() > MAX_CHARACTERS)
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
            Regex::new(r"(?i)^(?:Psalm\s+)?(\d{1,3})(?::([1-9]\d{0,2})-([1-9]\d{0,2}))?(?:\s\(([a-z])\))?(?:\s\((\d{1,2})\))?$")
                .expect("valid psalm reference regex")
        });

        let reference = reference
            .trim()
            .replace(['\u{2013}', '\u{2014}', '\u{2212}'], "-");

        let caps = RE
            .captures(&reference)
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
            .unwrap_or(300);

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
            title: reference,
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

fn clone_seed(
    pres: &mut Presentation,
    seed_slide: usize,
    description: &str,
) -> anyhow::Result<usize> {
    pres.clone_slide(seed_slide)
        .with_context(|| format!("clone TWPC {description} seed slide"))
}

fn set_shape_text(
    pres: &mut Presentation,
    slide: usize,
    shape: &str,
    text: &str,
) -> anyhow::Result<()> {
    pres.slide_mut(slide)?.shape(shape)?.set_text(text)?;
    Ok(())
}

fn set_shape_runs(
    pres: &mut Presentation,
    slide: usize,
    shape: &str,
    runs: &[Run],
) -> anyhow::Result<()> {
    pres.slide_mut(slide)?.shape(shape)?.set_rich_text(runs)?;
    Ok(())
}

fn psalm_runs(text: &str) -> Vec<Run> {
    let mut marked = Vec::new();
    let mut remaining = text;
    let mut underlined = false;
    while !remaining.is_empty() {
        let next_open = remaining.find("<underline>");
        let next_close = remaining.find("</underline>");
        let next = match (next_open, next_close) {
            (Some(open), Some(close)) => Some(open.min(close)),
            (Some(open), None) => Some(open),
            (None, Some(close)) => Some(close),
            (None, None) => None,
        };
        let Some(index) = next else {
            let mut run = Run::plain(remaining);
            run.underline = underlined;
            marked.push(run);
            break;
        };
        if index > 0 {
            let mut run = Run::plain(&remaining[..index]);
            run.underline = underlined;
            marked.push(run);
        }
        if remaining[index..].starts_with("<underline>") {
            underlined = true;
            remaining = &remaining[index + "<underline>".len()..];
        } else {
            underlined = false;
            remaining = &remaining[index + "</underline>".len()..];
        }
    }

    let mut runs = Vec::new();
    for run in marked {
        let mut current = String::new();
        let mut current_superscript = None;
        for character in run.text.chars() {
            let superscript = normalise_superscript(character);
            let is_superscript = superscript.is_some();
            if current_superscript.is_some_and(|value| value != is_superscript) {
                let mut split = Run::plain(std::mem::take(&mut current));
                split.underline = run.underline;
                split.superscript = current_superscript.unwrap_or(false);
                runs.push(split);
            }
            current_superscript = Some(is_superscript);
            current.push(superscript.unwrap_or(character));
        }
        if !current.is_empty() {
            let mut split = Run::plain(current);
            split.underline = run.underline;
            split.superscript = current_superscript.unwrap_or(false);
            runs.push(split);
        }
    }
    runs
}

fn normalise_superscript(character: char) -> Option<char> {
    Some(match character {
        '⁰' => '0',
        '¹' => '1',
        '²' => '2',
        '³' => '3',
        '⁴' => '4',
        '⁵' => '5',
        '⁶' => '6',
        '⁷' => '7',
        '⁸' => '8',
        '⁹' => '9',
        '⁻' => '-',
        _ => return None,
    })
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
        assert_eq!(pages.len(), 2);
    }

    #[test]
    fn psalm_23_uses_readable_two_stanza_groups() {
        let psalm = Psalm::find("Psalm 23:1-6").expect("Psalm 23 exists");
        let pages = propose_psalm_groups(&psalm.stanzas);

        assert_eq!(pages.len(), 3);
        assert!(pages.iter().all(|page| page.lines().count() <= 7));
    }

    #[test]
    fn psalm_runs_preserve_superscript_verse_numbers_and_underlining() {
        let runs = psalm_runs("⁴<underline>Though I</underline> walk through the valley");
        assert!(runs.iter().any(|run| run.superscript && run.text == "4"));
        assert!(runs
            .iter()
            .any(|run| run.underline && run.text == "Though I"));
        assert!(!runs.iter().any(|run| run.text.contains("underline")));
    }
}
