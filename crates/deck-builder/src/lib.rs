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
const NOTICES_BODY_X: u64 = 765_544;
const NOTICES_BODY_Y: u64 = 1_100_000;
const NOTICES_BODY_WIDTH: u64 = 8_893_463;
// Keep the seed textbox's bottom edge unchanged while adding space below the title.
const NOTICES_BODY_BOTTOM: u64 = 5_279_036;
const NOTICE_FONT_SIZE: u32 = 2_800;
const NOTICE_DETAIL_FONT_SIZE: u32 = 2_400;

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
pub struct Teaching {
    pub selection: String,
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

    fn teaching(&self, source: TeachingSource, selection: &str) -> anyhow::Result<Teaching> {
        Teaching::find(source, selection)
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
                let pages = paginate_notice_rows(rows, 5);
                for page in pages {
                    let slide = clone_seed(&mut pres, SEED_NOTICES, "notices")?;
                    set_shape_text(&mut pres, slide, "Title 1", heading)?;
                    pres.slide_mut(slide)?.shape("TextBox 2")?.set_position(
                        NOTICES_BODY_X,
                        NOTICES_BODY_Y,
                        NOTICES_BODY_WIDTH,
                        NOTICES_BODY_BOTTOM - NOTICES_BODY_Y,
                    )?;
                    set_shape_runs(&mut pres, slide, "TextBox 2", &notice_runs(&page))?;
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
                let body = textproc::normalise_scripture_lines(text.trim());
                let mut runs = textproc::scripture_runs(&body);
                if !reference.trim().is_empty() {
                    let separator = if body.is_empty() { "" } else { "\n\n" };
                    let mut reference_run = Run::plain(format!("{separator}{}", reference.trim()));
                    reference_run.italic = true;
                    runs.push(reference_run);
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
                show_verse_numbers,
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
                    set_shape_runs(
                        &mut pres,
                        slide,
                        "TextShape 2",
                        &psalm_runs(&stanza, *show_verse_numbers),
                    )?;
                    if index + 1 == count {
                        pres.copy_shape(SEED_SONG_FINAL, "TextBox 4", slide, "Psalm Credits")?;
                        pres.slide_mut(slide)?
                            .shape("Psalm Credits")?
                            .set_position(6724800, 6080400, 4146550, 1754326)?;
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
                let resolved = if text.trim().is_empty() {
                    if selection.trim().is_empty() {
                        "Choose a teaching question or enter teaching text".to_string()
                    } else {
                        let item = sources.teaching(*source, selection)?;
                        format!("{}\n\n{}", item.question, item.answer)
                    }
                } else {
                    text.clone()
                };
                if resolved.trim().is_empty() {
                    return Err(anyhow!("teaching text cannot be empty"));
                }
                let slide = clone_seed(&mut pres, SEED_TEACHING, "teaching")?;
                set_shape_text(&mut pres, slide, "TextShape 1", heading)?;
                set_shape_runs(&mut pres, slide, "TextShape 3", &teaching_runs(&resolved))?;
            }
            ServiceComponent::LiturgyBlock {
                heading, key, text, ..
            } => {
                if text.trim().is_empty() {
                    if let Some(source_slides) = service
                        .preset
                        .is_lords_supper()
                        .then(|| lords_supper_canonical_slides(key))
                        .flatten()
                    {
                        for &source_slide in source_slides {
                            clone_seed(
                                &mut pres,
                                source_slide,
                                &format!("Lord's Supper {key} canonical liturgy"),
                            )?;
                        }
                        continue;
                    }
                }
                let (speaker, pages) = if text.trim().is_empty() {
                    let component = sources.fixed_component(key)?;
                    (component.speaker, component.slides)
                } else {
                    (
                        sources
                            .fixed_component(key)
                            .map(|component| component.speaker)
                            .unwrap_or_default(),
                        text.split("\n\n").map(str::to_string).collect(),
                    )
                };
                for (index, page) in pages.into_iter().enumerate() {
                    let slide = clone_seed(&mut pres, SEED_LITURGY, "liturgy")?;
                    set_shape_text(&mut pres, slide, "TextShape 1", heading)?;
                    set_shape_runs(
                        &mut pres,
                        slide,
                        "TextShape 2",
                        &liturgy_runs(&speaker, &page, index == 0),
                    )?;
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
                        .join(" – ");
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

fn paginate_notice_rows(rows: &[NoticeRow], rows_per_slide: usize) -> Vec<Vec<NoticeRow>> {
    if rows.is_empty() {
        return vec![Vec::new()];
    }
    rows.chunks(rows_per_slide.max(1))
        .map(|page| page.to_vec())
        .collect()
}

fn notice_runs(rows: &[NoticeRow]) -> Vec<Run> {
    let mut runs = Vec::new();
    for (index, row) in rows.iter().enumerate() {
        let when = row.when.trim();
        let title = row.title.trim();
        let details = row.details.trim();
        let lead_separator = if when.is_empty() || title.is_empty() {
            ""
        } else {
            " – "
        };
        if !when.is_empty() {
            runs.push(
                Run::plain(when)
                    .with_font_size(NOTICE_FONT_SIZE)
                    .with_text_style("Arial Black", "accent1"),
            );
        }
        if !lead_separator.is_empty() {
            runs.push(
                Run::plain(" ")
                    .with_font_size(NOTICE_FONT_SIZE)
                    .with_text_style("Arial Black", "accent1"),
            );
            runs.push(
                Run::plain("–")
                    .with_font_size(NOTICE_FONT_SIZE)
                    .with_text_style("Arial Black", "000000"),
            );
            runs.push(
                Run::plain(" ")
                    .with_font_size(NOTICE_FONT_SIZE)
                    .with_text_style("Arial Black", "accent1"),
            );
        }
        if !title.is_empty() {
            let mut title_run = Run::plain(title)
                .with_font_size(NOTICE_FONT_SIZE)
                .with_text_style(
                    "Arial Black",
                    if row.emphasis { "FF0000" } else { "000000" },
                );
            title_run.bold = row.emphasis;
            runs.push(title_run);
        }
        if !details.is_empty() {
            runs.push(Run::plain("\n"));
            let mut detail_run = Run::plain(details)
                .with_font_size(NOTICE_DETAIL_FONT_SIZE)
                .with_text_style("Arial Black", "000000");
            detail_run.italic = true;
            runs.push(detail_run);
        }
        if index + 1 < rows.len() {
            runs.push(Run::plain("\n\n"));
        }
    }
    runs
}

pub fn propose_psalm_groups(stanzas: &[String]) -> Vec<String> {
    let (max_lines, characters_per_line) = psalm_layout_capacity();

    let mut groups = Vec::new();
    let mut current = String::new();
    let mut current_lines = 0;
    for stanza in stanzas {
        let stanza_lines = estimated_psalm_lines(stanza, characters_per_line);
        let separator_lines = usize::from(!current.is_empty());
        if !current.is_empty() && current_lines + separator_lines + stanza_lines > max_lines {
            groups.push(current);
            current = String::new();
            current_lines = 0;
        }
        if !current.is_empty() {
            current.push_str("\n\n");
        }
        current.push_str(stanza);
        current_lines += separator_lines + stanza_lines;
    }
    if !current.is_empty() {
        groups.push(current);
    }
    groups
}

// TextShape 2 in the Psalm seed is 9,683,583 × 4,506,298 EMU. Generated
// runs use 28pt Arial Black, with 100% paragraph line spacing. These limits
// are derived from that box rather than from an arbitrary stanza count. The
// 50% average glyph-width and 85/90% safety factors leave room for wide words,
// verse markers, and PowerPoint's font metrics while still grouping ordinary
// three- or four-line Psalm paragraphs.
const PSALM_TEXT_BOX_WIDTH_EMU: u64 = 9_683_583;
const PSALM_TEXT_BOX_HEIGHT_EMU: u64 = 4_506_298;
const PSALM_FONT_SIZE_HUNDREDTHS_PT: u64 = 2_800;
const EMU_PER_POINT: u64 = 12_700;
const PSALM_AVERAGE_GLYPH_WIDTH_PERCENT: u64 = 50;
const PSALM_VERTICAL_SAFETY_PERCENT: u64 = 85;
const PSALM_HORIZONTAL_SAFETY_PERCENT: u64 = 90;

fn psalm_layout_capacity() -> (usize, usize) {
    let font_height_emu = PSALM_FONT_SIZE_HUNDREDTHS_PT * EMU_PER_POINT / 100;
    let natural_lines = PSALM_TEXT_BOX_HEIGHT_EMU / font_height_emu;
    let max_lines = (natural_lines * PSALM_VERTICAL_SAFETY_PERCENT / 100).max(1) as usize;

    let average_glyph_width_emu = font_height_emu * PSALM_AVERAGE_GLYPH_WIDTH_PERCENT / 100;
    let natural_characters = PSALM_TEXT_BOX_WIDTH_EMU / average_glyph_width_emu;
    let characters_per_line =
        (natural_characters * PSALM_HORIZONTAL_SAFETY_PERCENT / 100).max(1) as usize;

    (max_lines, characters_per_line)
}

fn estimated_psalm_lines(text: &str, characters_per_line: usize) -> usize {
    text.lines()
        .map(|line| {
            if line.trim().is_empty() {
                return 1;
            }
            let visible_characters = line
                .replace("<underline>", "")
                .replace("</underline>", "")
                .chars()
                .count();
            visible_characters.div_ceil(characters_per_line).max(1)
        })
        .sum::<usize>()
        .max(1)
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

impl Teaching {
    pub fn find(source: TeachingSource, selection: &str) -> anyhow::Result<Self> {
        match source {
            TeachingSource::WestminsterShorterCatechism => Self::catechism_entry(&WSC, selection),
            TeachingSource::Heidelberg1891 => Self::catechism_entry(&HEIDELBERG, selection),
            TeachingSource::WestminsterConfessionOriginalBritish => {
                Self::confession_entry(selection)
            }
        }
    }

    fn catechism_entry(file: &CatechismFile, selection: &str) -> anyhow::Result<Self> {
        let question = parse_catechism_selection(selection)?;
        let item = file
            .data
            .iter()
            .find(|item| item.number == question)
            .ok_or_else(|| anyhow!("catechism question not found: {question}"))?;
        Ok(Self {
            selection: question.to_string(),
            question: item.question.clone(),
            answer: item.answer.clone(),
        })
    }

    fn confession_entry(selection: &str) -> anyhow::Result<Self> {
        let (chapter, section) = parse_confession_selection(selection)?;
        let entry = WCF
            .data
            .iter()
            .find(|item| item.chapter == chapter)
            .ok_or_else(|| anyhow!("confession chapter not found: {chapter}"))?;
        let question = format!("Chapter {chapter}: {}", entry.title);
        match section {
            Some(wanted) => {
                let found = entry
                    .sections
                    .iter()
                    .find(|item| item.section == wanted)
                    .ok_or_else(|| anyhow!("confession section not found: {chapter}.{wanted}"))?;
                Ok(Self {
                    selection: format!("{chapter}.{wanted}"),
                    question,
                    answer: found.content.clone(),
                })
            }
            None => Ok(Self {
                selection: chapter.to_string(),
                question,
                answer: entry
                    .sections
                    .iter()
                    .map(|item| format!("{}. {}", item.section, item.content))
                    .collect::<Vec<_>>()
                    .join("\n\n"),
            }),
        }
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

fn lords_supper_canonical_slides(key: &str) -> Option<&'static [usize]> {
    match key {
        "prayer_for_purity" => Some(&[9]),
        "ten_commandments" => Some(&[10, 11, 12, 13]),
        "lords_prayer" => Some(&[14]),
        "confession" => Some(&[23, 24, 25]),
        "assurance" => Some(&[26]),
        "comfortable_words" => Some(&[27, 28, 29, 30]),
        "humble_access" => Some(&[31, 32]),
        "consecration" => Some(&[33, 34, 35, 36, 37]),
        "final_blessing" => Some(&[47]),
        _ => None,
    }
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

fn psalm_runs(text: &str, show_verse_numbers: bool) -> Vec<Run> {
    let text = if show_verse_numbers {
        text.to_string()
    } else {
        strip_leading_verse_numbers(text)
    };
    let mut marked = Vec::new();
    let mut remaining = text.as_str();
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
            let mut run = Run::plain(remaining)
                .with_font_size(2800)
                .with_text_style("Arial Black", "000000");
            run.underline = underlined;
            marked.push(run);
            break;
        };
        if index > 0 {
            let mut run = Run::plain(&remaining[..index])
                .with_font_size(2800)
                .with_text_style("Arial Black", "000000");
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
                let mut split = Run::plain(std::mem::take(&mut current))
                    .with_font_size(2800)
                    .with_text_style("Arial Black", "000000");
                split.underline = run.underline;
                split.superscript = current_superscript.unwrap_or(false);
                runs.push(split);
            }
            current_superscript = Some(is_superscript);
            current.push(superscript.unwrap_or(character));
        }
        if !current.is_empty() {
            let mut split = Run::plain(current)
                .with_font_size(2800)
                .with_text_style("Arial Black", "000000");
            split.underline = run.underline;
            split.superscript = current_superscript.unwrap_or(false);
            runs.push(split);
        }
    }
    runs
}

fn strip_leading_verse_numbers(text: &str) -> String {
    text.split_inclusive('\n')
        .map(|line| {
            let (content, newline) = line
                .strip_suffix('\n')
                .map_or((line, ""), |content| (content, "\n"));
            let mut output = String::with_capacity(content.len() + newline.len());
            let mut rest = content;
            let leading_whitespace = rest.len() - rest.trim_start().len();
            output.push_str(&rest[..leading_whitespace]);
            rest = &rest[leading_whitespace..];
            if rest.starts_with("<underline>") {
                output.push_str("<underline>");
                rest = &rest["<underline>".len()..];
                let number_end = rest
                    .char_indices()
                    .take_while(|(_, character)| is_verse_number_character(*character))
                    .last()
                    .map(|(index, character)| index + character.len_utf8())
                    .unwrap_or(0);
                rest = &rest[number_end..];
            } else {
                let number_end = rest
                    .char_indices()
                    .take_while(|(_, character)| is_verse_number_character(*character))
                    .last()
                    .map(|(index, character)| index + character.len_utf8())
                    .unwrap_or(0);
                rest = &rest[number_end..];
            }
            output.push_str(rest);
            output.push_str(newline);
            output
        })
        .collect()
}

fn is_verse_number_character(character: char) -> bool {
    matches!(
        character,
        '⁰' | '¹' | '²' | '³' | '⁴' | '⁵' | '⁶' | '⁷' | '⁸' | '⁹' | '⁻'
    )
}

fn teaching_runs(text: &str) -> Vec<Run> {
    let text = text.replace("\r\n", "\n");
    let paragraphs = text
        .trim()
        .split("\n\n")
        .map(str::trim)
        .filter(|paragraph| !paragraph.is_empty())
        .collect::<Vec<_>>();

    let (question, answer) = if paragraphs.len() >= 2 {
        (paragraphs[0].to_string(), paragraphs[1..].join("\n\n"))
    } else {
        split_teaching_answer(paragraphs.first().copied().unwrap_or_default())
    };

    let mut runs = Vec::new();
    if !question.is_empty() {
        runs.push(teaching_run(&question, "accent1"));
    }
    if !question.is_empty() && !answer.is_empty() {
        runs.push(Run::plain("\n\n"));
    }
    if !answer.is_empty() {
        runs.push(teaching_run(answer, "000000"));
    }
    runs
}

fn split_teaching_answer(text: &str) -> (String, String) {
    static ANSWER_RE: Lazy<Regex> = Lazy::new(|| {
        Regex::new(r"(?i)^\s*(?:a(?:nswer)?\s*[.:]|answer\s+)")
            .expect("valid teaching answer regex")
    });
    let lines = text.lines().collect::<Vec<_>>();
    let answer_line = lines
        .iter()
        .position(|line| ANSWER_RE.is_match(line))
        .filter(|index| *index > 0);
    match answer_line {
        Some(index) => (
            lines[..index].join("\n").trim().to_string(),
            lines[index..].join("\n").trim().to_string(),
        ),
        None => (text.trim().to_string(), String::new()),
    }
}

fn teaching_run(text: impl Into<String>, color: &str) -> Run {
    let mut run = Run::plain(text)
        .with_font_size(3000)
        .with_text_style("Arial Black", color);
    run.bold = true;
    run
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

pub fn parse_catechism_selection(selection: &str) -> anyhow::Result<u16> {
    static RE: Lazy<Regex> = Lazy::new(|| {
        Regex::new(r"(?i)^\s*(?:q(?:uestion)?\s*\.?\s*)?(\d{1,3})\s*$")
            .expect("valid catechism selection regex")
    });
    RE.captures(selection.trim())
        .and_then(|captures| captures.get(1))
        .and_then(|number| number.as_str().parse::<u16>().ok())
        .filter(|number| *number > 0)
        .ok_or_else(|| anyhow!("enter a catechism question such as 1, Q1, or Q. 1"))
}

pub fn parse_confession_selection(selection: &str) -> anyhow::Result<(u16, Option<u16>)> {
    static RE: Lazy<Regex> = Lazy::new(|| {
        Regex::new(r"(?i)^\s*(?:ch(?:apter)?\s*\.?\s*)?(\d{1,2})(?:\s*[.:]\s*(\d{1,2}))?\s*$")
            .expect("valid confession selection regex")
    });
    RE.captures(selection.trim())
        .map(|captures| {
            (
                captures
                    .get(1)
                    .and_then(|number| number.as_str().parse::<u16>().ok()),
                captures
                    .get(2)
                    .and_then(|number| number.as_str().parse::<u16>().ok()),
            )
        })
        .and_then(|(chapter, section)| chapter.map(|chapter| (chapter, section)))
        .filter(|(chapter, section)| {
            *chapter > 0 && section.map(|number| number > 0) != Some(false)
        })
        .ok_or_else(|| anyhow!("enter a confession chapter such as 21, 21.8, or Chapter 21"))
}

fn liturgy_runs(speaker: &str, text: &str, include_speaker: bool) -> Vec<Run> {
    let (existing_cue, body) = leading_liturgy_cue(text, speaker);
    let cue = existing_cue.or_else(|| {
        include_speaker
            .then(|| speaker.trim().to_string())
            .filter(|cue| !cue.is_empty())
    });
    let mut runs = Vec::new();
    if let Some(cue) = cue {
        runs.push(
            Run::plain(cue)
                .with_font_size(4000)
                .with_text_style("Liberation Serif", "accent1"),
        );
        runs.last_mut().expect("cue run").italic = true;
        runs.last_mut().expect("cue run").bold = true;
        if !body.trim().is_empty() {
            runs.push(
                Run::plain("  ")
                    .with_font_size(4000)
                    .with_text_style("Arial Black", "accent1"),
            );
        }
    }

    let body = body.trim_start();
    if body.is_empty() {
        return runs;
    }
    static AMEN_RE: Lazy<Regex> =
        Lazy::new(|| Regex::new(r"(?i)\bAmen[.!?…]*$").expect("valid Amen regex"));
    if let Some(amen) = AMEN_RE.find(body.trim_end()) {
        let body_end = body.trim_end().len();
        let prefix = &body[..amen.start()];
        if !prefix.is_empty() {
            runs.push(
                Run::plain(prefix)
                    .with_font_size(4000)
                    .with_text_style("Arial Black", "000000"),
            );
        }
        let mut amen_run = Run::plain(&body[amen.start()..body_end])
            .with_font_size(4000)
            .with_text_style("Arial Black", "accent1");
        amen_run.bold = true;
        runs.push(amen_run);
    } else {
        runs.push(
            Run::plain(body)
                .with_font_size(4000)
                .with_text_style("Arial Black", "000000"),
        );
    }
    runs
}

fn leading_liturgy_cue<'a>(text: &'a str, speaker: &str) -> (Option<String>, &'a str) {
    let text = text.trim_start();
    let mut candidates = vec![speaker.trim()];
    candidates.extend(["Minister.", "All."]);
    for candidate in candidates {
        if candidate.is_empty()
            || !text
                .get(..candidate.len())
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case(candidate))
            || !text.get(candidate.len()..).is_none_or(|rest| {
                rest.is_empty() || rest.chars().next().is_some_and(char::is_whitespace)
            })
        {
            continue;
        }
        return (
            Some(text[..candidate.len()].to_string()),
            text[candidate.len()..].trim_start(),
        );
    }
    (None, text)
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
struct CatechismFile {
    #[serde(rename = "Data")]
    data: Vec<CatechismItem>,
}

#[derive(Debug, Deserialize)]
struct CatechismItem {
    #[serde(rename = "Number")]
    number: u16,
    #[serde(rename = "Question")]
    question: String,
    #[serde(rename = "Answer")]
    answer: String,
}

#[derive(Debug, Deserialize)]
struct ConfessionFile {
    #[serde(rename = "Data")]
    data: Vec<ConfessionChapter>,
}

#[derive(Debug, Deserialize)]
struct ConfessionChapter {
    #[serde(rename = "Chapter")]
    chapter: u16,
    #[serde(rename = "Title")]
    title: String,
    #[serde(rename = "Sections")]
    sections: Vec<ConfessionSection>,
}

#[derive(Debug, Deserialize)]
struct ConfessionSection {
    #[serde(rename = "Section")]
    section: u16,
    #[serde(rename = "Content")]
    content: String,
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
static WSC: Lazy<CatechismFile> = Lazy::new(|| {
    serde_json::from_str(include_str!("../assets/wsc.json")).expect("valid embedded wsc.json")
});
static HEIDELBERG: Lazy<CatechismFile> = Lazy::new(|| {
    serde_json::from_str(include_str!("../assets/heidelberg.json"))
        .expect("valid embedded heidelberg.json")
});
static WCF: Lazy<ConfessionFile> = Lazy::new(|| {
    serde_json::from_str(include_str!("../assets/wcf.json")).expect("valid embedded wcf.json")
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
    fn psalm_grouping_preserves_every_stanza_and_fits_two_common_stanzas() {
        let stanzas = vec![
            "1Blessed is the one who turns away\nfrom where the wicked walk,\nWho does not stand in sinners' paths\nor sit with those who mock."
                .to_string(),
            "2Instead he finds God's holy law\nhis joy and great delight;\nHe makes the precepts of the LORD\nhis study day and night."
                .to_string(),
            "i".to_string(),
        ];
        let pages = propose_psalm_groups(&stanzas);
        assert_eq!(pages.join("\n\n"), stanzas.join("\n\n"));
        assert_eq!(pages[0], format!("{}\n\n{}", stanzas[0], stanzas[1]));
        assert_eq!(pages.len(), 2);
    }

    #[test]
    fn psalm_grouping_never_splits_an_oversized_stanza() {
        let oversized = (1..=9)
            .map(|line| format!("line {line}"))
            .collect::<Vec<_>>()
            .join("\n");
        let stanzas = vec![oversized.clone(), "final stanza".to_string()];

        let pages = propose_psalm_groups(&stanzas);

        assert_eq!(pages, vec![oversized, "final stanza".to_string()]);
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
        let runs = psalm_runs(
            "⁴<underline>Though I</underline> walk through the valley",
            true,
        );
        assert!(runs.iter().all(
            |run| run.font_size == Some(2800) && run.typeface.as_deref() == Some("Arial Black")
        ));
        assert!(runs.iter().any(|run| run.superscript && run.text == "4"));
        assert!(runs
            .iter()
            .any(|run| run.underline && run.text == "Though I"));
        assert!(!runs.iter().any(|run| run.text.contains("underline")));
    }

    #[test]
    fn psalm_runs_can_hide_only_leading_verse_numbers() {
        let runs = psalm_runs(
            "⁴<underline>Though I</underline> walk\nThe word ⁴ remains",
            false,
        );

        assert!(runs
            .iter()
            .any(|run| run.underline && run.text == "Though I"));
        assert!(runs.iter().any(|run| run.superscript && run.text == "4"));
        assert!(runs.iter().any(|run| run.text.contains("The word ")));
        assert!(runs.iter().any(|run| run.text.contains("\nThe word ")));
    }

    #[test]
    fn old_psalm_records_default_to_showing_verse_numbers() {
        let component: ServiceComponent = serde_json::from_str(
            r#"{"type":"psalm","id":"psalm","heading":"Psalm","reference":"Psalm 23","slide_breaks":[]}"#,
        )
        .expect("old Psalm record deserialises");

        assert!(matches!(
            component,
            ServiceComponent::Psalm {
                show_verse_numbers: true,
                ..
            }
        ));
    }
}
