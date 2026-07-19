use regex::Regex;
use std::collections::{BTreeMap, HashSet};
use std::io::{Cursor, Read, Write};
use thiserror::Error;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

const CONTENT_TYPES: &str = "[Content_Types].xml";
const PRESENTATION: &str = "ppt/presentation.xml";
const PRESENTATION_RELS: &str = "ppt/_rels/presentation.xml.rels";
const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SLIDE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const LAYOUT_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const MAX_COMPRESSED_PACKAGE_BYTES: usize = 25 * 1024 * 1024;
const MAX_DECOMPRESSED_PACKAGE_BYTES: u64 = 200 * 1024 * 1024;
const MAX_PACKAGE_PARTS: usize = 10_000;

#[derive(Debug, Error)]
pub enum Error {
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("missing package part {0}")]
    MissingPart(String),
    #[error("invalid pptx package: {0}")]
    InvalidPackage(String),
    #[error("slide index {0} is out of range")]
    SlideIndex(usize),
    #[error("placeholder index {0} is out of range")]
    PlaceholderIndex(usize),
}

pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Run {
    pub text: String,
    pub superscript: bool,
    pub bold: bool,
    pub italic: bool,
}

impl Run {
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            superscript: false,
            bold: false,
            italic: false,
        }
    }

    pub fn superscript(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            superscript: true,
            bold: false,
            italic: false,
        }
    }
}

#[derive(Debug, Clone)]
struct SlideMeta {
    id: u32,
    rid: String,
    part: String,
}

#[derive(Debug, Clone)]
pub struct Presentation {
    files: BTreeMap<String, Vec<u8>>,
    slides: Vec<SlideMeta>,
    next_slide_number: u32,
    next_slide_id: u32,
    next_rel_id: u32,
}

impl Presentation {
    pub fn open_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > MAX_COMPRESSED_PACKAGE_BYTES {
            return Err(Error::InvalidPackage(format!(
                "compressed package exceeds {} MiB",
                MAX_COMPRESSED_PACKAGE_BYTES / 1024 / 1024
            )));
        }
        let mut archive = ZipArchive::new(Cursor::new(bytes))?;
        if archive.len() > MAX_PACKAGE_PARTS {
            return Err(Error::InvalidPackage(format!(
                "package contains more than {MAX_PACKAGE_PARTS} parts"
            )));
        }
        let mut files = BTreeMap::new();
        let mut decompressed_bytes = 0_u64;
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            if file.is_dir() {
                continue;
            }
            decompressed_bytes = decompressed_bytes.saturating_add(file.size());
            if decompressed_bytes > MAX_DECOMPRESSED_PACKAGE_BYTES {
                return Err(Error::InvalidPackage(format!(
                    "decompressed package exceeds {} MiB",
                    MAX_DECOMPRESSED_PACKAGE_BYTES / 1024 / 1024
                )));
            }
            let mut data = Vec::new();
            file.read_to_end(&mut data)?;
            files.insert(file.name().to_string(), data);
        }

        let rels = utf8_part(&files, PRESENTATION_RELS)?;
        let presentation = utf8_part(&files, PRESENTATION)?;
        let mut rel_map = BTreeMap::new();
        let mut max_rel_id = 0;
        for tag in relationship_tags(&rels) {
            if let Some(id) = attr(&tag, "Id") {
                max_rel_id = max_rel_id.max(rid_number(&id));
                if attr(&tag, "Type").is_some_and(|value| value.ends_with("/slide")) {
                    if let Some(target) = attr(&tag, "Target") {
                        rel_map.insert(id, target_to_slide_part(&target));
                    }
                }
            }
        }

        let mut slides = Vec::new();
        let mut max_slide_id = 255;
        let sld_re = Regex::new(r#"<p:sldId\b[^>]*/>"#).expect("valid regex");
        for m in sld_re.find_iter(&presentation) {
            let tag = m.as_str();
            let id = attr(tag, "id")
                .ok_or_else(|| Error::InvalidPackage("slide id missing id".into()))?
                .parse::<u32>()
                .map_err(|_| Error::InvalidPackage("invalid slide id".into()))?;
            let rid = attr(tag, "r:id")
                .or_else(|| attr(tag, "id"))
                .ok_or_else(|| Error::InvalidPackage("slide id missing relationship".into()))?;
            let part = rel_map
                .get(&rid)
                .ok_or_else(|| Error::InvalidPackage(format!("missing relationship {rid}")))?
                .clone();
            max_slide_id = max_slide_id.max(id);
            slides.push(SlideMeta { id, rid, part });
        }

        let next_slide_number = files
            .keys()
            .filter_map(|name| slide_number(name))
            .max()
            .unwrap_or(0)
            + 1;

        Ok(Self {
            files,
            slides,
            next_slide_number,
            next_slide_id: max_slide_id + 1,
            next_rel_id: max_rel_id + 1,
        })
    }

    pub fn save_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        for (name, data) in &self.files {
            writer.start_file(name, options)?;
            writer.write_all(data)?;
        }
        Ok(writer.finish()?.into_inner())
    }

    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    pub fn slide_size(&self) -> Result<(u64, u64)> {
        let presentation = self.part_string(PRESENTATION)?;
        let tag = Regex::new(r#"<p:sldSz\b[^>]*/>"#)
            .expect("valid slide size regex")
            .find(&presentation)
            .ok_or_else(|| Error::InvalidPackage("presentation has no slide size".into()))?
            .as_str();
        let cx = attr(tag, "cx")
            .and_then(|value| value.parse().ok())
            .ok_or_else(|| Error::InvalidPackage("slide width is invalid".into()))?;
        let cy = attr(tag, "cy")
            .and_then(|value| value.parse().ok())
            .ok_or_else(|| Error::InvalidPackage("slide height is invalid".into()))?;
        Ok((cx, cy))
    }

    pub fn validate_song_source(&self, expected_size: (u64, u64)) -> Result<()> {
        let size = self.slide_size()?;
        if size != expected_size {
            return Err(Error::InvalidPackage(format!(
                "incompatible slide dimensions {} × {}; expected {} × {}",
                size.0, size.1, expected_size.0, expected_size.1
            )));
        }
        if self.slides.is_empty() {
            return Err(Error::InvalidPackage("song deck contains no slides".into()));
        }
        for name in self.files.keys() {
            let lowercase = name.to_ascii_lowercase();
            if lowercase.contains("vbaproject")
                || lowercase.contains("/activex/")
                || lowercase.contains("/embeddings/")
                || lowercase.contains("oleobject")
            {
                return Err(Error::InvalidPackage(format!(
                    "unsupported active or embedded content in {name}"
                )));
            }
        }
        let content_types = self.part_string(CONTENT_TYPES)?;
        if content_types.contains("macroEnabled") || content_types.contains("vnd.ms-office") {
            return Err(Error::InvalidPackage(
                "macro-enabled PowerPoint packages are not accepted".into(),
            ));
        }
        for (name, bytes) in &self.files {
            if !name.ends_with(".rels") {
                continue;
            }
            let xml = String::from_utf8(bytes.clone())
                .map_err(|_| Error::InvalidPackage(format!("{name} is not utf-8")))?;
            for relationship in relationship_tags(&xml) {
                let external = attr(&relationship, "TargetMode").as_deref() == Some("External");
                let relationship_type = attr(&relationship, "Type").unwrap_or_default();
                if external
                    && ["image", "audio", "video", "media"]
                        .iter()
                        .any(|kind| relationship_type.ends_with(kind))
                {
                    return Err(Error::InvalidPackage(format!(
                        "externally linked media is not accepted ({name})"
                    )));
                }
            }
        }
        self.validate()
    }

    pub fn add_slide_from_layout(&mut self, layout_index: usize) -> Result<usize> {
        let layout_number = layout_index + 1;
        let layout_part = format!("ppt/slideLayouts/slideLayout{layout_number}.xml");
        let layout_xml = self.part_string(&layout_part)?;
        let layout_start = layout_xml
            .find("<p:sldLayout")
            .ok_or_else(|| Error::InvalidPackage(format!("invalid {layout_part}")))?;
        let inner_start = layout_xml[layout_start..]
            .find('>')
            .map(|idx| layout_start + idx)
            .ok_or_else(|| Error::InvalidPackage(format!("invalid {layout_part}")))?
            + 1;
        let inner_end = layout_xml
            .rfind("</p:sldLayout>")
            .ok_or_else(|| Error::InvalidPackage(format!("invalid {layout_part}")))?;
        let slide_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{}</p:sld>"#,
            &layout_xml[inner_start..inner_end]
        );

        let slide_number = self.next_slide_number;
        self.next_slide_number += 1;
        let part = format!("ppt/slides/slide{slide_number}.xml");
        self.files.insert(part.clone(), slide_xml.into_bytes());

        let rels = self.slide_rels_from_layout(layout_number)?;
        self.files.insert(slide_rels_part(&part), rels.into_bytes());

        self.add_slide_to_presentation(part)
    }

    pub fn clone_slide(&mut self, index: usize) -> Result<usize> {
        let source = self
            .slides
            .get(index)
            .ok_or(Error::SlideIndex(index))?
            .part
            .clone();
        let slide_number = self.next_slide_number;
        self.next_slide_number += 1;
        let part = format!("ppt/slides/slide{slide_number}.xml");
        let xml = self
            .files
            .get(&source)
            .ok_or_else(|| Error::MissingPart(source.clone()))?
            .clone();
        self.files.insert(part.clone(), xml);

        let source_rels = slide_rels_part(&source);
        if let Some(rels) = self.files.get(&source_rels).cloned() {
            self.files.insert(slide_rels_part(&part), rels);
        }

        self.add_slide_to_presentation(part)
    }

    pub fn delete_slide(&mut self, index: usize) -> Result<()> {
        let slide = self
            .slides
            .get(index)
            .ok_or(Error::SlideIndex(index))?
            .clone();
        self.slides.remove(index);
        self.files.remove(&slide.part);
        self.files.remove(&slide_rels_part(&slide.part));
        self.remove_content_type(&slide.part)?;
        self.remove_presentation_relationship(&slide.rid)?;
        self.sync_slide_order()
    }

    pub fn reorder(&mut self, order: &[usize]) -> Result<()> {
        if order.len() != self.slides.len() {
            return Err(Error::InvalidPackage(
                "slide reorder length mismatch".into(),
            ));
        }
        let mut seen = HashSet::new();
        let mut reordered = Vec::with_capacity(order.len());
        for &idx in order {
            if !seen.insert(idx) {
                return Err(Error::InvalidPackage(
                    "slide reorder contains duplicate index".into(),
                ));
            }
            reordered.push(self.slides.get(idx).ok_or(Error::SlideIndex(idx))?.clone());
        }
        self.slides = reordered;
        self.sync_slide_order()
    }

    pub fn slide_mut(&mut self, index: usize) -> Result<SlideMut<'_>> {
        if index >= self.slides.len() {
            return Err(Error::SlideIndex(index));
        }
        Ok(SlideMut {
            presentation: self,
            index,
        })
    }

    pub fn slide_xml(&self, index: usize) -> Result<String> {
        let slide = self.slides.get(index).ok_or(Error::SlideIndex(index))?;
        self.part_string(&slide.part)
    }

    pub fn slide_text(&self, index: usize) -> Result<String> {
        let xml = self.slide_xml(index)?;
        let text_re = Regex::new(r#"<a:t>(.*?)</a:t>"#).expect("valid regex");
        Ok(text_re
            .captures_iter(&xml)
            .map(|cap| xml_unescape(&cap[1]))
            .collect::<Vec<_>>()
            .join("\n"))
    }

    pub fn validate(&self) -> Result<()> {
        let content_types = self.part_string(CONTENT_TYPES)?;
        let pres_rels = self.part_string(PRESENTATION_RELS)?;
        for slide in &self.slides {
            if !self.files.contains_key(&slide.part) {
                return Err(Error::MissingPart(slide.part.clone()));
            }
            let content_name = format!("/{}", slide.part);
            if !content_types.contains(&format!("PartName=\"{content_name}\"")) {
                return Err(Error::InvalidPackage(format!(
                    "missing content type override for {}",
                    slide.part
                )));
            }
            if !pres_rels.contains(&format!("Id=\"{}\"", slide.rid)) {
                return Err(Error::InvalidPackage(format!(
                    "missing presentation relationship {}",
                    slide.rid
                )));
            }
            self.validate_slide_relationships(slide)?;
        }
        Ok(())
    }

    fn add_slide_to_presentation(&mut self, part: String) -> Result<usize> {
        self.add_content_type(&part)?;
        let rid = format!("rId{}", self.next_rel_id);
        self.next_rel_id += 1;
        self.add_presentation_relationship(&rid, &part)?;
        let id = self.next_slide_id.max(256);
        self.next_slide_id = id + 1;
        self.slides.push(SlideMeta { id, rid, part });
        self.sync_slide_order()?;
        Ok(self.slides.len() - 1)
    }

    fn part_string(&self, part: &str) -> Result<String> {
        utf8_part(&self.files, part)
    }

    fn add_content_type(&mut self, part: &str) -> Result<()> {
        let part_name = format!("/{part}");
        let mut xml = self.part_string(CONTENT_TYPES)?;
        if xml.contains(&format!("PartName=\"{part_name}\"")) {
            return Ok(());
        }
        let override_xml =
            format!(r#"<Override PartName="{part_name}" ContentType="{SLIDE_CONTENT_TYPE}"/>"#);
        insert_before(&mut xml, "</Types>", &override_xml)?;
        self.files.insert(CONTENT_TYPES.into(), xml.into_bytes());
        Ok(())
    }

    fn remove_content_type(&mut self, part: &str) -> Result<()> {
        let part_name = regex::escape(&format!("/{part}"));
        let re = Regex::new(&format!(r#"<Override\b[^>]*PartName="{part_name}"[^>]*/>"#))
            .expect("valid regex");
        let xml = self.part_string(CONTENT_TYPES)?;
        self.files.insert(
            CONTENT_TYPES.into(),
            re.replace_all(&xml, "").as_bytes().to_vec(),
        );
        Ok(())
    }

    fn add_presentation_relationship(&mut self, rid: &str, part: &str) -> Result<()> {
        let target = part.strip_prefix("ppt/").unwrap_or(part);
        let mut xml = self.part_string(PRESENTATION_RELS)?;
        let relationship =
            format!(r#"<Relationship Id="{rid}" Type="{SLIDE_REL_TYPE}" Target="{target}"/>"#);
        insert_before(&mut xml, "</Relationships>", &relationship)?;
        self.files
            .insert(PRESENTATION_RELS.into(), xml.into_bytes());
        Ok(())
    }

    fn remove_presentation_relationship(&mut self, rid: &str) -> Result<()> {
        let re = Regex::new(&format!(
            r#"<Relationship\b[^>]*Id="{}"[^>]*/>"#,
            regex::escape(rid)
        ))
        .expect("valid regex");
        let xml = self.part_string(PRESENTATION_RELS)?;
        self.files.insert(
            PRESENTATION_RELS.into(),
            re.replace_all(&xml, "").as_bytes().to_vec(),
        );
        Ok(())
    }

    fn sync_slide_order(&mut self) -> Result<()> {
        let mut xml = self.part_string(PRESENTATION)?;
        let start_tag_end = xml
            .find("<p:sldIdLst>")
            .ok_or_else(|| Error::InvalidPackage("missing p:sldIdLst".into()))?
            + "<p:sldIdLst>".len();
        let end_tag_start = xml
            .find("</p:sldIdLst>")
            .ok_or_else(|| Error::InvalidPackage("missing /p:sldIdLst".into()))?;
        let entries = self
            .slides
            .iter()
            .map(|slide| format!(r#"<p:sldId id="{}" r:id="{}"/>"#, slide.id, slide.rid))
            .collect::<String>();
        xml.replace_range(start_tag_end..end_tag_start, &entries);
        self.files.insert(PRESENTATION.into(), xml.into_bytes());
        Ok(())
    }

    fn slide_rels_from_layout(&self, layout_number: usize) -> Result<String> {
        let layout_rels_part =
            format!("ppt/slideLayouts/_rels/slideLayout{layout_number}.xml.rels");
        let layout_rels = self.part_string(&layout_rels_part)?;
        let mut copied = Vec::new();
        let mut max_rid = 0;
        for tag in relationship_tags(&layout_rels) {
            if attr(&tag, "Type").is_some_and(|value| value.ends_with("/slideMaster")) {
                continue;
            }
            if let Some(id) = attr(&tag, "Id") {
                max_rid = max_rid.max(rid_number(&id));
            }
            copied.push(tag);
        }
        let layout_rid = format!("rId{}", max_rid + 1);
        copied.push(format!(
            r#"<Relationship Id="{layout_rid}" Type="{LAYOUT_REL_TYPE}" Target="../slideLayouts/slideLayout{layout_number}.xml"/>"#
        ));
        Ok(format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
            copied.join("")
        ))
    }

    fn validate_slide_relationships(&self, slide: &SlideMeta) -> Result<()> {
        let rels_part = slide_rels_part(&slide.part);
        let rels = self.part_string(&rels_part)?;
        let rel_ids = relationship_tags(&rels)
            .into_iter()
            .filter_map(|tag| attr(&tag, "Id"))
            .collect::<HashSet<_>>();
        let xml = self.part_string(&slide.part)?;
        let refs = Regex::new(r#"r:(?:embed|id)="([^"]+)""#).expect("valid regex");
        for cap in refs.captures_iter(&xml) {
            if !rel_ids.contains(&cap[1]) {
                return Err(Error::InvalidPackage(format!(
                    "{} references missing relationship {}",
                    slide.part, &cap[1]
                )));
            }
        }
        Ok(())
    }
}

pub struct SlideMut<'a> {
    presentation: &'a mut Presentation,
    index: usize,
}

impl<'a> SlideMut<'a> {
    pub fn placeholder(self, index: usize) -> Result<PlaceholderMut<'a>> {
        let xml = self.presentation.slide_xml(self.index)?;
        find_placeholder_block(&xml, index).map(|_| PlaceholderMut {
            presentation: self.presentation,
            slide_index: self.index,
            placeholder_index: index,
        })
    }
}

pub struct PlaceholderMut<'a> {
    presentation: &'a mut Presentation,
    slide_index: usize,
    placeholder_index: usize,
}

impl PlaceholderMut<'_> {
    pub fn set_text(self, text: &str) -> Result<()> {
        self.set_rich_text(&[Run::plain(text)])
    }

    pub fn set_rich_text(self, runs: &[Run]) -> Result<()> {
        let part = self.presentation.slides[self.slide_index].part.clone();
        let mut xml = self.presentation.part_string(&part)?;
        let (start, end, block) = find_placeholder_block(&xml, self.placeholder_index)?;
        let updated = replace_text_body(&block, runs)?;
        xml.replace_range(start..end, &updated);
        self.presentation.files.insert(part, xml.into_bytes());
        Ok(())
    }
}

fn utf8_part(files: &BTreeMap<String, Vec<u8>>, part: &str) -> Result<String> {
    let bytes = files
        .get(part)
        .ok_or_else(|| Error::MissingPart(part.to_string()))?;
    String::from_utf8(bytes.clone())
        .map_err(|_| Error::InvalidPackage(format!("{part} is not utf-8")))
}

fn relationship_tags(xml: &str) -> Vec<String> {
    let re = Regex::new(r#"<Relationship\b[^>]*/>"#).expect("valid regex");
    re.find_iter(xml).map(|m| m.as_str().to_string()).collect()
}

fn attr(tag: &str, name: &str) -> Option<String> {
    let re = Regex::new(&format!(r#"\b{}="([^"]*)""#, regex::escape(name))).ok()?;
    re.captures(tag).map(|cap| cap[1].to_string())
}

fn rid_number(rid: &str) -> u32 {
    rid.strip_prefix("rId")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(0)
}

fn target_to_slide_part(target: &str) -> String {
    if target.starts_with("ppt/") {
        target.to_string()
    } else if target.starts_with("slides/") {
        format!("ppt/{target}")
    } else {
        target.trim_start_matches("../").to_string()
    }
}

fn slide_number(name: &str) -> Option<u32> {
    let re = Regex::new(r#"^ppt/slides/slide(\d+)\.xml$"#).expect("valid regex");
    re.captures(name)?.get(1)?.as_str().parse().ok()
}

fn slide_rels_part(part: &str) -> String {
    let file = part.trim_start_matches("ppt/slides/");
    format!("ppt/slides/_rels/{file}.rels")
}

fn insert_before(xml: &mut String, needle: &str, insertion: &str) -> Result<()> {
    let idx = xml
        .find(needle)
        .ok_or_else(|| Error::InvalidPackage(format!("missing {needle}")))?;
    xml.insert_str(idx, insertion);
    Ok(())
}

fn find_placeholder_block(xml: &str, index: usize) -> Result<(usize, usize, String)> {
    let re = Regex::new(r#"(?s)<p:sp>.*?</p:sp>"#).expect("valid regex");
    let mut current = 0;
    for m in re.find_iter(xml) {
        let block = m.as_str();
        if block.contains("<p:ph") {
            if current == index {
                return Ok((m.start(), m.end(), block.to_string()));
            }
            current += 1;
        }
    }
    Err(Error::PlaceholderIndex(index))
}

fn replace_text_body(shape: &str, runs: &[Run]) -> Result<String> {
    let body_re = Regex::new(r#"(?s)<p:txBody>.*?</p:txBody>"#).expect("valid regex");
    let tx_body = if let Some(m) = body_re.find(shape) {
        let existing = m.as_str();
        let prefix_end = existing
            .find("<a:p")
            .unwrap_or(existing.len() - "</p:txBody>".len());
        let prefix = &existing[..prefix_end];
        let paragraph_properties = extract_xml_tag(existing, "a:pPr").unwrap_or_default();
        let run_properties = extract_xml_tag(existing, "a:rPr")
            .or_else(|| extract_xml_tag(existing, "a:defRPr").map(def_rpr_to_rpr));
        format!(
            "{prefix}<a:p>{paragraph_properties}{}</a:p></p:txBody>",
            runs_xml(runs, run_properties.as_deref())
        )
    } else {
        format!(
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p>{}</a:p></p:txBody>",
            runs_xml(runs, None)
        )
    };

    if body_re.is_match(shape) {
        Ok(body_re.replace(shape, tx_body.as_str()).to_string())
    } else {
        let mut updated = shape.to_string();
        insert_before(&mut updated, "</p:sp>", &tx_body)?;
        Ok(updated)
    }
}

fn runs_xml(runs: &[Run], default_rpr: Option<&str>) -> String {
    runs.iter()
        .map(|run| {
            let attrs = run_attrs(run);
            let rpr = default_rpr
                .map(|default| merge_run_attrs(default, &attrs))
                .unwrap_or_else(|| format!("<a:rPr lang=\"en-GB\"{attrs}/>"));
            format!("<a:r>{rpr}<a:t>{}</a:t></a:r>", xml_escape(&run.text))
        })
        .collect::<String>()
}

fn run_attrs(run: &Run) -> String {
    let mut attrs = String::new();
    if run.superscript {
        attrs.push_str(" baseline=\"30000\"");
    }
    if run.bold {
        attrs.push_str(" b=\"1\"");
    }
    if run.italic {
        attrs.push_str(" i=\"1\"");
    }
    attrs
}

fn merge_run_attrs(default_rpr: &str, attrs: &str) -> String {
    if attrs.is_empty() {
        return default_rpr.to_string();
    }
    if let Some(idx) = default_rpr.find('>') {
        let mut out = default_rpr.to_string();
        let insert_at = idx.saturating_sub(usize::from(default_rpr[..idx].ends_with('/')));
        out.insert_str(insert_at, attrs);
        out
    } else {
        format!("<a:rPr lang=\"en-GB\"{attrs}/>")
    }
}

fn extract_xml_tag(xml: &str, tag: &str) -> Option<String> {
    let escaped = regex::escape(tag);
    let paired = Regex::new(&format!(r#"(?s)<{escaped}\b[^>]*>.*?</{escaped}>"#)).ok()?;
    if let Some(m) = paired.find(xml) {
        return Some(m.as_str().to_string());
    }
    let empty = Regex::new(&format!(r#"<{escaped}\b[^>]*/>"#)).ok()?;
    empty.find(xml).map(|m| m.as_str().to_string())
}

fn def_rpr_to_rpr(xml: String) -> String {
    xml.replacen("<a:defRPr", "<a:rPr", 1)
        .replace("</a:defRPr>", "</a:rPr>")
}

fn xml_escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn xml_unescape(text: &str) -> String {
    text.replace("&apos;", "'")
        .replace("&quot;", "\"")
        .replace("&gt;", ">")
        .replace("&lt;", "<")
        .replace("&amp;", "&")
}
