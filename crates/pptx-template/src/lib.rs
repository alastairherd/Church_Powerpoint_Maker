use regex::Regex;
use std::collections::{BTreeMap, HashSet};
use std::io::{Cursor, Read, Write};
use thiserror::Error;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

const CONTENT_TYPES: &str = "[Content_Types].xml";
const PRESENTATION: &str = "ppt/presentation.xml";
const PRESENTATION_RELS: &str = "ppt/_rels/presentation.xml.rels";
const REVISION_INFO: &str = "ppt/revisionInfo.xml";
const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SLIDE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_MASTER_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const LAYOUT_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const MIN_SLIDE_MASTER_ID: u32 = 2_147_483_648;
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
    pub underline: bool,
    pub font_size: Option<u32>,
    pub typeface: Option<String>,
    pub color: Option<String>,
}

impl Run {
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            superscript: false,
            bold: false,
            italic: false,
            underline: false,
            font_size: None,
            typeface: None,
            color: None,
        }
    }

    pub fn superscript(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            superscript: true,
            bold: false,
            italic: false,
            underline: false,
            font_size: None,
            typeface: None,
            color: None,
        }
    }

    pub fn underlined(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            superscript: false,
            bold: false,
            italic: false,
            underline: true,
            font_size: None,
            typeface: None,
            color: None,
        }
    }

    pub fn with_text_style(
        mut self,
        typeface: impl Into<String>,
        color: impl Into<String>,
    ) -> Self {
        self.typeface = Some(typeface.into());
        self.color = Some(color.into());
        self
    }

    pub fn with_font_size(mut self, font_size: u32) -> Self {
        self.font_size = Some(font_size);
        self
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
    next_master_id: u32,
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

        let mut max_master_id = MIN_SLIDE_MASTER_ID - 1;
        let master_re = Regex::new(r#"<p:sldMasterId\b[^>]*/>"#).expect("valid regex");
        for m in master_re.find_iter(&presentation) {
            if let Some(id) = attr(m.as_str(), "id").and_then(|value| value.parse::<u32>().ok()) {
                max_master_id = max_master_id.max(id);
            }
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
            next_master_id: max_master_id.saturating_add(1),
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

    pub fn import_slides(&mut self, source_bytes: &[u8]) -> Result<Vec<usize>> {
        let source = Self::open_bytes(source_bytes)?;
        source.validate_song_source(self.slide_size()?)?;

        let import_number = self
            .files
            .keys()
            .filter(|name| name.contains("_import"))
            .count() as u32
            + 1;
        let mut mapping = BTreeMap::new();
        for slide in &source.slides {
            let destination = format!("ppt/slides/slide{}.xml", self.next_slide_number);
            self.next_slide_number += 1;
            mapping.insert(slide.part.clone(), destination);
        }

        let mut processed = HashSet::new();
        let mut imported = Vec::with_capacity(source.slides.len());
        for slide in &source.slides {
            let destination = mapping
                .get(&slide.part)
                .expect("source slides are mapped before import")
                .clone();
            self.copy_import_part(
                &source,
                &slide.part,
                &destination,
                import_number,
                &mut mapping,
                &mut processed,
            )?;
            imported.push(self.add_slide_to_presentation(destination)?);
        }
        Ok(imported)
    }

    pub fn remove_auxiliary_content(&mut self) -> Result<()> {
        let removed: HashSet<String> = self
            .files
            .keys()
            .filter(|name| auxiliary_part(name))
            .cloned()
            .collect();
        let relationship_files: Vec<String> = self
            .files
            .keys()
            .filter(|name| name.ends_with(".rels") && !removed.contains(*name))
            .cloned()
            .collect();

        for relationships_name in relationship_files {
            let Some(owner_part) = owner_part_for_relationships(&relationships_name) else {
                continue;
            };
            let relationships = self.part_string(&relationships_name)?;
            let mut owner_xml = self
                .files
                .get(&owner_part)
                .and_then(|bytes| String::from_utf8(bytes.clone()).ok());
            let mut kept = Vec::new();
            for relationship in relationship_tags(&relationships) {
                let relationship_type = attr(&relationship, "Type").unwrap_or_default();
                let external = attr(&relationship, "TargetMode").as_deref() == Some("External");
                let target_removed = attr(&relationship, "Target")
                    .and_then(|target| resolve_part_target(&owner_part, &target).ok())
                    .is_some_and(|target| removed.contains(&target));
                if external || drop_auxiliary_relationship(&relationship_type) || target_removed {
                    if let (Some(xml), Some(id)) = (owner_xml.as_mut(), attr(&relationship, "Id")) {
                        *xml = strip_relationship_reference(xml, &id);
                    }
                } else {
                    kept.push(relationship);
                }
            }
            self.files.insert(
                relationships_name,
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
                    kept.join("")
                )
                .into_bytes(),
            );
            if let Some(xml) = owner_xml {
                self.files.insert(owner_part, xml.into_bytes());
            }
        }

        self.files.retain(|name, _| !removed.contains(name));
        let mut content_types = self.part_string(CONTENT_TYPES)?;
        let overrides = Regex::new(r#"<Override\b[^>]*/>"#).expect("valid override regex");
        content_types = overrides
            .replace_all(&content_types, |captures: &regex::Captures<'_>| {
                let entry = captures.get(0).expect("override match").as_str();
                let removed_entry = attr(entry, "PartName")
                    .map(|name| name.trim_start_matches('/').to_string())
                    .is_some_and(|name| removed.contains(&name));
                if removed_entry {
                    String::new()
                } else {
                    entry.to_string()
                }
            })
            .to_string();
        self.files
            .insert(CONTENT_TYPES.into(), content_types.into_bytes());
        Ok(())
    }

    fn copy_import_part(
        &mut self,
        source: &Presentation,
        source_part: &str,
        destination_part: &str,
        import_number: u32,
        mapping: &mut BTreeMap<String, String>,
        processed: &mut HashSet<String>,
    ) -> Result<()> {
        if !processed.insert(source_part.to_string()) {
            return Ok(());
        }
        let source_bytes = source
            .files
            .get(source_part)
            .ok_or_else(|| Error::MissingPart(source_part.to_string()))?
            .clone();
        self.files
            .insert(destination_part.to_string(), source_bytes.clone());
        self.copy_import_content_type(source, source_part, destination_part)?;

        let source_relationships_part = relationships_part(source_part);
        let Some(source_relationships) = source.files.get(&source_relationships_part) else {
            return Ok(());
        };
        let relationships_xml = String::from_utf8(source_relationships.clone()).map_err(|_| {
            Error::InvalidPackage(format!("{source_relationships_part} is not utf-8"))
        })?;
        let mut owner_xml = String::from_utf8(source_bytes).ok();
        let mut copied_relationships = Vec::new();

        for relationship in relationship_tags(&relationships_xml) {
            let relationship_type = attr(&relationship, "Type").unwrap_or_default();
            let relationship_id = attr(&relationship, "Id").unwrap_or_default();
            let external = attr(&relationship, "TargetMode").as_deref() == Some("External");
            if external || drop_import_relationship(&relationship_type) {
                if let Some(xml) = owner_xml.as_mut() {
                    *xml = strip_relationship_reference(xml, &relationship_id);
                }
                continue;
            }

            let target = attr(&relationship, "Target").ok_or_else(|| {
                Error::InvalidPackage(format!(
                    "relationship {relationship_id} in {source_relationships_part} has no target"
                ))
            })?;
            let source_target = resolve_part_target(source_part, &target)?;
            if !source.files.contains_key(&source_target) {
                return Err(Error::MissingPart(source_target));
            }

            let destination_target = if let Some(mapped) = mapping.get(&source_target) {
                mapped.clone()
            } else if let Some(existing) =
                self.identical_import_part(source, &source_target, mapping)?
            {
                mapping.insert(source_target.clone(), existing.clone());
                processed.insert(source_target.clone());
                existing
            } else {
                let mapped = self.allocate_import_part(&source_target, import_number);
                mapping.insert(source_target.clone(), mapped.clone());
                mapped
            };

            self.copy_import_part(
                source,
                &source_target,
                &destination_target,
                import_number,
                mapping,
                processed,
            )?;
            let rewritten_target = relative_part_target(destination_part, &destination_target);
            copied_relationships.push(replace_xml_attr(&relationship, "Target", &rewritten_target));
        }

        if let Some(xml) = owner_xml {
            self.files
                .insert(destination_part.to_string(), xml.into_bytes());
        }
        let destination_relationships_part = relationships_part(destination_part);
        self.files.insert(
            destination_relationships_part,
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
                copied_relationships.join("")
            )
            .into_bytes(),
        );
        Ok(())
    }

    fn identical_import_part(
        &self,
        source: &Presentation,
        source_part: &str,
        mapping: &BTreeMap<String, String>,
    ) -> Result<Option<String>> {
        // PowerPoint requires each slide master to own its theme part; reusing
        // an existing identical theme would make two masters share one.
        if source_part.starts_with("ppt/theme/") {
            return Ok(None);
        }
        let Some(bytes) = source.files.get(source_part) else {
            return Ok(None);
        };
        let Some((_, extension)) = source_part.rsplit_once('.') else {
            return Ok(None);
        };
        let candidate = self
            .files
            .iter()
            .find(|(name, candidate)| {
                name.rsplit_once('.')
                    .is_some_and(|(_, candidate_extension)| candidate_extension == extension)
                    && import_part_contents_match(source_part, bytes, name, candidate)
                    && self
                        .import_relationships_match(source, source_part, name, mapping)
                        .unwrap_or(false)
            })
            .map(|(name, _)| name.clone());
        Ok(candidate)
    }

    fn import_relationships_match(
        &self,
        source: &Presentation,
        source_part: &str,
        destination_part: &str,
        mapping: &BTreeMap<String, String>,
    ) -> Result<bool> {
        let source_relationships = source
            .files
            .get(&relationships_part(source_part))
            .map(|bytes| String::from_utf8(bytes.clone()))
            .transpose()
            .map_err(|_| {
                Error::InvalidPackage(format!("{} is not utf-8", relationships_part(source_part)))
            })?;
        let destination_relationships = self
            .files
            .get(&relationships_part(destination_part))
            .map(|bytes| String::from_utf8(bytes.clone()))
            .transpose()
            .map_err(|_| {
                Error::InvalidPackage(format!(
                    "{} is not utf-8",
                    relationships_part(destination_part)
                ))
            })?;
        let Some(source_relationships) = source_relationships else {
            return Ok(destination_relationships.is_none());
        };
        let Some(destination_relationships) = destination_relationships else {
            return Ok(false);
        };
        let source_tags = relationship_tags(&source_relationships);
        let mut destination_tags = relationship_tags(&destination_relationships);
        if source_tags.len() != destination_tags.len() {
            return Ok(false);
        }
        for source_tag in source_tags {
            let Some(destination_index) = destination_tags.iter().position(|destination_tag| {
                relationship_shape(&source_tag) == relationship_shape(destination_tag)
                    && self
                        .import_relationship_target_matches(
                            source,
                            source_part,
                            &source_tag,
                            destination_part,
                            destination_tag,
                            mapping,
                        )
                        .unwrap_or(false)
            }) else {
                return Ok(false);
            };
            destination_tags.remove(destination_index);
        }
        Ok(destination_tags.is_empty())
    }

    fn import_relationship_target_matches(
        &self,
        source: &Presentation,
        source_part: &str,
        source_relationship: &str,
        destination_part: &str,
        destination_relationship: &str,
        mapping: &BTreeMap<String, String>,
    ) -> Result<bool> {
        let Some(source_target) = attr(source_relationship, "Target") else {
            return Ok(false);
        };
        let Some(destination_target) = attr(destination_relationship, "Target") else {
            return Ok(false);
        };
        if attr(source_relationship, "TargetMode").as_deref() == Some("External") {
            return Ok(source_target == destination_target);
        }
        let source_target = resolve_part_target(source_part, &source_target)?;
        let destination_target = resolve_part_target(destination_part, &destination_target)?;
        if mapping
            .get(&source_target)
            .is_some_and(|mapped| mapped == &destination_target)
            || source_target == destination_target
        {
            return Ok(true);
        }
        let Some(source_bytes) = source.files.get(&source_target) else {
            return Ok(false);
        };
        let Some(destination_bytes) = self.files.get(&destination_target) else {
            return Ok(false);
        };
        if !import_part_contents_match(
            &source_target,
            source_bytes,
            &destination_target,
            destination_bytes,
        ) {
            return Ok(false);
        }
        Ok(relationship_shapes_match(
            source.files.get(&relationships_part(&source_target)),
            self.files.get(&relationships_part(&destination_target)),
        ))
    }

    fn allocate_import_part(&self, source_part: &str, import_number: u32) -> String {
        if !self.files.contains_key(source_part) {
            return source_part.to_string();
        }
        let (directory, filename) = source_part.rsplit_once('/').unwrap_or(("", source_part));
        let (stem, extension) = filename
            .rsplit_once('.')
            .map(|(stem, extension)| (stem, format!(".{extension}")))
            .unwrap_or((filename, String::new()));
        let prefix = stem.trim_end_matches(|character: char| character.is_ascii_digit());
        if prefix.len() < stem.len() {
            for number in 1..100_000 {
                let candidate = if directory.is_empty() {
                    format!("{prefix}{number}{extension}")
                } else {
                    format!("{directory}/{prefix}{number}{extension}")
                };
                if !self.files.contains_key(&candidate) {
                    return candidate;
                }
            }
        }
        let candidate = if directory.is_empty() {
            format!("{stem}_import{import_number}{extension}")
        } else {
            format!("{directory}/{stem}_import{import_number}{extension}")
        };
        if !self.files.contains_key(&candidate) {
            return candidate;
        }
        for suffix in 2..100_000 {
            let candidate = if directory.is_empty() {
                format!("{stem}_import{import_number}_{suffix}{extension}")
            } else {
                format!("{directory}/{stem}_import{import_number}_{suffix}{extension}")
            };
            if !self.files.contains_key(&candidate) {
                return candidate;
            }
        }
        unreachable!("package part limit prevents exhausting import names")
    }

    fn copy_import_content_type(
        &mut self,
        source: &Presentation,
        source_part: &str,
        destination_part: &str,
    ) -> Result<()> {
        let source_types = source.part_string(CONTENT_TYPES)?;
        let mut destination_types = self.part_string(CONTENT_TYPES)?;
        let source_name = format!("/{source_part}");
        let destination_name = format!("/{destination_part}");
        let overrides = Regex::new(r#"<Override\b[^>]*/>"#).expect("valid override regex");
        if let Some(source_override) = overrides
            .find_iter(&source_types)
            .find(|entry| attr(entry.as_str(), "PartName").as_deref() == Some(source_name.as_str()))
        {
            if !destination_types.contains(&format!("PartName=\"{destination_name}\"")) {
                let copied =
                    replace_xml_attr(source_override.as_str(), "PartName", &destination_name);
                insert_before(&mut destination_types, "</Types>", &copied)?;
            }
        } else if let Some(extension) = source_part.rsplit_once('.').map(|(_, value)| value) {
            let defaults = Regex::new(r#"<Default\b[^>]*/>"#).expect("valid default regex");
            let source_default = defaults
                .find_iter(&source_types)
                .find(|entry| attr(entry.as_str(), "Extension").as_deref() == Some(extension))
                .map(|entry| entry.as_str().to_string());
            if let Some(source_default) = source_default {
                let has_default = defaults
                    .find_iter(&destination_types)
                    .any(|entry| attr(entry.as_str(), "Extension").as_deref() == Some(extension));
                if !has_default {
                    insert_before(&mut destination_types, "</Types>", &source_default)?;
                }
            }
        }
        self.files
            .insert(CONTENT_TYPES.into(), destination_types.into_bytes());
        Ok(())
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
        let mut xml = self.part_string(&source)?;

        let source_rels = slide_rels_part(&source);
        if let Some(rels) = self.files.get(&source_rels).cloned() {
            let relationships = String::from_utf8(rels)
                .map_err(|_| Error::InvalidPackage(format!("{source_rels} is not utf-8")))?;
            let mut copied = Vec::new();
            for relationship in relationship_tags(&relationships) {
                let relationship_type = attr(&relationship, "Type").unwrap_or_default();
                let external = attr(&relationship, "TargetMode").as_deref() == Some("External");
                if external || drop_import_relationship(&relationship_type) {
                    if let Some(id) = attr(&relationship, "Id") {
                        xml = strip_relationship_reference(&xml, &id);
                    }
                } else {
                    copied.push(relationship);
                }
            }
            self.files.insert(
                slide_rels_part(&part),
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
                    copied.join("")
                )
                .into_bytes(),
            );
        }
        self.files.insert(part.clone(), xml.into_bytes());

        self.add_slide_to_presentation(part)
    }

    pub fn copy_shape(
        &mut self,
        source_slide: usize,
        source_name: &str,
        target_slide: usize,
        target_name: &str,
    ) -> Result<()> {
        let source_xml = self.slide_xml(source_slide)?;
        let (_, _, source_shape) = find_shape_block(&source_xml, source_name)?;
        let target_part = self
            .slides
            .get(target_slide)
            .ok_or(Error::SlideIndex(target_slide))?
            .part
            .clone();
        let mut target_xml = self.part_string(&target_part)?;
        let ids = Regex::new(r#"<p:cNvPr\b[^>]*\bid="(\d+)""#).expect("valid shape id regex");
        let next_id = ids
            .captures_iter(&target_xml)
            .filter_map(|captures| captures[1].parse::<u32>().ok())
            .max()
            .unwrap_or(1)
            + 1;
        let properties = Regex::new(r#"<p:cNvPr\b[^>]*/?>"#).expect("valid shape properties regex");
        let original = properties.find(&source_shape).ok_or_else(|| {
            Error::InvalidPackage(format!("shape {source_name} has no properties"))
        })?;
        let renamed = replace_xml_attr(
            &replace_xml_attr(original.as_str(), "id", &next_id.to_string()),
            "name",
            target_name,
        );
        let copied = source_shape.replacen(original.as_str(), &renamed, 1);
        insert_before(&mut target_xml, "</p:spTree>", &copied)?;
        self.files.insert(target_part, target_xml.into_bytes());
        Ok(())
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
        self.validate_notes_master_reference(&pres_rels)?;
        self.validate_slide_master_references(&pres_rels)?;
        let mut visited = HashSet::new();
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
            self.validate_part_graph(&slide.part, &content_types, &mut visited)?;
        }
        Ok(())
    }

    fn validate_notes_master_reference(&self, pres_rels: &str) -> Result<()> {
        let relationship_ids = relationship_tags(pres_rels)
            .into_iter()
            .filter_map(|tag| attr(&tag, "Id"))
            .collect::<HashSet<_>>();
        let presentation = self.part_string(PRESENTATION)?;
        let notes_master =
            Regex::new(r#"<p:notesMasterId\b[^>]*>"#).expect("valid notes master regex");
        for tag in notes_master.find_iter(&presentation) {
            let tag = tag.as_str();
            let relationship_id = attr(tag, "r:id").ok_or_else(|| {
                Error::InvalidPackage("notes master declaration missing relationship".into())
            })?;
            if !relationship_ids.contains(&relationship_id) {
                return Err(Error::InvalidPackage(format!(
                    "notes master declaration references missing relationship {relationship_id}"
                )));
            }
        }
        Ok(())
    }

    fn validate_slide_master_references(&self, pres_rels: &str) -> Result<()> {
        let presentation = self.part_string(PRESENTATION)?;
        let mut master_relationships = BTreeMap::new();
        let mut master_targets = HashSet::new();
        for relationship in relationship_tags(pres_rels) {
            if !attr(&relationship, "Type").is_some_and(|value| value.ends_with("/slideMaster")) {
                continue;
            }
            let rid = attr(&relationship, "Id").ok_or_else(|| {
                Error::InvalidPackage("slide master relationship is missing an id".into())
            })?;
            let target = attr(&relationship, "Target").ok_or_else(|| {
                Error::InvalidPackage(format!("slide master relationship {rid} has no target"))
            })?;
            let target = resolve_part_target(PRESENTATION, &target)?;
            if !master_targets.insert(target.clone()) {
                return Err(Error::InvalidPackage(format!(
                    "duplicate slide master relationship target {target}"
                )));
            }
            master_relationships.insert(rid, target);
        }

        let master_entries =
            Regex::new(r#"<p:sldMasterId\b[^>]*/>"#).expect("valid slide master entry regex");
        let mut entry_ids = HashSet::new();
        let mut entry_targets = HashSet::new();
        for entry in master_entries.find_iter(&presentation) {
            let entry = entry.as_str();
            let id = attr(entry, "id")
                .ok_or_else(|| Error::InvalidPackage("slide master entry is missing id".into()))?;
            let id_number = parse_slide_master_id(&id)?;
            if !entry_ids.insert(id_number) {
                return Err(Error::InvalidPackage(format!(
                    "duplicate slide master id {id}"
                )));
            }
            let rid = attr(entry, "r:id").ok_or_else(|| {
                Error::InvalidPackage(format!("slide master entry {id} is missing relationship"))
            })?;
            let target = master_relationships.get(&rid).ok_or_else(|| {
                Error::InvalidPackage(format!(
                    "slide master entry {id} references missing relationship {rid}"
                ))
            })?;
            if !entry_targets.insert(target.clone()) {
                return Err(Error::InvalidPackage(format!(
                    "duplicate slide master entry for {target}"
                )));
            }
        }

        for slide in &self.slides {
            for master in self.reachable_slide_masters(&slide.part)? {
                if !entry_targets.contains(&master) {
                    return Err(Error::InvalidPackage(format!(
                        "{} reaches unregistered slide master {master}",
                        slide.part
                    )));
                }
            }
        }
        let mut theme_owners: BTreeMap<String, String> = BTreeMap::new();
        for master in &entry_targets {
            let relationships = self.part_string(&relationships_part(master))?;
            for relationship in relationship_tags(&relationships) {
                if !attr(&relationship, "Type").is_some_and(|value| value.ends_with("/theme")) {
                    continue;
                }
                let target = attr(&relationship, "Target").ok_or_else(|| {
                    Error::InvalidPackage(format!("{master} theme relationship has no target"))
                })?;
                let theme = resolve_part_target(master, &target)?;
                if let Some(other) = theme_owners.insert(theme.clone(), master.clone()) {
                    return Err(Error::InvalidPackage(format!(
                        "slide masters {other} and {master} share theme part {theme}"
                    )));
                }
            }
        }
        self.registered_master_and_layout_ids()?;
        Ok(())
    }

    fn registered_master_and_layout_ids(&self) -> Result<HashSet<u32>> {
        let presentation = self.part_string(PRESENTATION)?;
        let presentation_relationships = self.part_string(PRESENTATION_RELS)?;
        let master_relationships = relationship_tags(&presentation_relationships)
            .into_iter()
            .filter(|relationship| {
                attr(relationship, "Type").is_some_and(|value| value.ends_with("/slideMaster"))
            })
            .map(|relationship| {
                let rid = attr(&relationship, "Id").ok_or_else(|| {
                    Error::InvalidPackage("slide master relationship is missing an id".into())
                })?;
                let target = attr(&relationship, "Target").ok_or_else(|| {
                    Error::InvalidPackage(format!("slide master relationship {rid} has no target"))
                })?;
                Ok((rid, resolve_part_target(PRESENTATION, &target)?))
            })
            .collect::<Result<BTreeMap<_, _>>>()?;
        let master_entry =
            Regex::new(r#"<p:sldMasterId\b[^>]*/>"#).expect("valid slide master entry regex");
        let mut ids = HashSet::new();
        let mut master_parts = HashSet::new();
        for entry in master_entry.find_iter(&presentation) {
            let entry = entry.as_str();
            let id = attr(entry, "id")
                .ok_or_else(|| Error::InvalidPackage("slide master entry is missing id".into()))?;
            let id = parse_slide_master_id(&id)?;
            if !ids.insert(id) {
                return Err(Error::InvalidPackage(format!(
                    "duplicate slide master id {id}"
                )));
            }
            let rid = attr(entry, "r:id").ok_or_else(|| {
                Error::InvalidPackage(format!("slide master entry {id} is missing relationship"))
            })?;
            let master_part = master_relationships.get(&rid).ok_or_else(|| {
                Error::InvalidPackage(format!(
                    "slide master entry {id} references missing relationship {rid}"
                ))
            })?;
            master_parts.insert(master_part.clone());
        }

        let layout_entry =
            Regex::new(r#"<p:sldLayoutId\b[^>]*/>"#).expect("valid slide layout entry regex");
        for part in master_parts {
            let xml = self.part_string(&part)?;
            for entry in layout_entry.find_iter(&xml) {
                let id = attr(entry.as_str(), "id").ok_or_else(|| {
                    Error::InvalidPackage(format!("{part} slide layout entry is missing id"))
                })?;
                let id = parse_slide_layout_id(&id)?;
                if !ids.insert(id) {
                    return Err(Error::InvalidPackage(format!(
                        "slide layout id {id} in {part} duplicates a master or layout id"
                    )));
                }
            }
        }
        Ok(ids)
    }

    fn reachable_slide_masters(&self, slide_part: &str) -> Result<Vec<String>> {
        let mut pending = vec![slide_part.to_string()];
        let mut visited = HashSet::new();
        let mut masters = BTreeMap::new();
        while let Some(part) = pending.pop() {
            if !visited.insert(part.clone()) {
                continue;
            }
            let relationships_name = relationships_part(&part);
            let Some(relationships) = self.files.get(&relationships_name) else {
                continue;
            };
            let relationships = String::from_utf8(relationships.clone())
                .map_err(|_| Error::InvalidPackage(format!("{relationships_name} is not utf-8")))?;
            for relationship in relationship_tags(&relationships) {
                if attr(&relationship, "TargetMode").as_deref() == Some("External") {
                    continue;
                }
                let relationship_type = attr(&relationship, "Type").unwrap_or_default();
                let is_master = relationship_type.ends_with("/slideMaster");
                let is_layout = relationship_type.ends_with("/slideLayout");
                if !is_master && !is_layout {
                    continue;
                }
                let target = attr(&relationship, "Target").ok_or_else(|| {
                    Error::InvalidPackage(format!(
                        "relationship in {relationships_name} has no target"
                    ))
                })?;
                let target = resolve_part_target(&part, &target)?;
                if !self.files.contains_key(&target) {
                    return Err(Error::MissingPart(target));
                }
                if is_master {
                    masters.insert(target.clone(), ());
                } else {
                    pending.push(target);
                }
            }
        }
        Ok(masters.into_keys().collect())
    }

    fn validate_part_graph(
        &self,
        part: &str,
        content_types: &str,
        visited: &mut HashSet<String>,
    ) -> Result<()> {
        if !visited.insert(part.to_string()) {
            return Ok(());
        }
        let has_override = content_types.contains(&format!("PartName=\"/{part}\""));
        let extension = part.rsplit_once('.').map(|(_, value)| value);
        let has_default = extension.is_some_and(|extension| {
            Regex::new(r#"<Default\b[^>]*/>"#)
                .expect("valid default regex")
                .find_iter(content_types)
                .any(|entry| attr(entry.as_str(), "Extension").as_deref() == Some(extension))
        });
        if !has_override && !has_default {
            return Err(Error::InvalidPackage(format!(
                "missing content type for {part}"
            )));
        }

        let relationships_name = relationships_part(part);
        let Some(relationships) = self.files.get(&relationships_name) else {
            return Ok(());
        };
        let relationships = String::from_utf8(relationships.clone())
            .map_err(|_| Error::InvalidPackage(format!("{relationships_name} is not utf-8")))?;
        for relationship in relationship_tags(&relationships) {
            if attr(&relationship, "TargetMode").as_deref() == Some("External") {
                continue;
            }
            let target = attr(&relationship, "Target").ok_or_else(|| {
                Error::InvalidPackage(format!(
                    "relationship in {relationships_name} has no target"
                ))
            })?;
            let target = resolve_part_target(part, &target)?;
            if !self.files.contains_key(&target) {
                return Err(Error::MissingPart(target));
            }
            self.validate_part_graph(&target, content_types, visited)?;
        }
        Ok(())
    }

    fn add_slide_to_presentation(&mut self, part: String) -> Result<usize> {
        for master in self.reachable_slide_masters(&part)? {
            self.register_slide_master(&master)?;
        }
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

    fn register_slide_master(&mut self, master_part: &str) -> Result<()> {
        let mut relationships_xml = self.part_string(PRESENTATION_RELS)?;
        let mut master_rid = None;
        for relationship in relationship_tags(&relationships_xml) {
            if !attr(&relationship, "Type").is_some_and(|value| value.ends_with("/slideMaster")) {
                continue;
            }
            let target = attr(&relationship, "Target").ok_or_else(|| {
                Error::InvalidPackage("slide master relationship has no target".into())
            })?;
            if resolve_part_target(PRESENTATION, &target)? == master_part {
                master_rid = attr(&relationship, "Id");
                break;
            }
        }

        let mut presentation = self.part_string(PRESENTATION)?;
        let entries =
            Regex::new(r#"<p:sldMasterId\b[^>]*/>"#).expect("valid slide master entry regex");
        let already_registered = entries.find_iter(&presentation).any(|entry| {
            attr(entry.as_str(), "r:id")
                .and_then(|rid| {
                    relationship_tags(&relationships_xml)
                        .into_iter()
                        .find(|relationship| {
                            attr(relationship, "Id").as_deref() == Some(rid.as_str())
                        })
                })
                .and_then(|relationship| attr(&relationship, "Target"))
                .and_then(|target| resolve_part_target(PRESENTATION, &target).ok())
                .as_deref()
                == Some(master_part)
        });
        if already_registered {
            self.registered_master_and_layout_ids()?;
            return Ok(());
        }

        let mut used_ids = self.normalize_slide_master_layout_ids(master_part)?;
        let id = self.allocate_master_or_layout_id(&mut used_ids)?;
        let master_rid = if let Some(rid) = master_rid {
            rid
        } else {
            let rid = self.allocate_presentation_relationship_id(&relationships_xml)?;
            let target = master_part.strip_prefix("ppt/").unwrap_or(master_part);
            let relationship = format!(
                r#"<Relationship Id="{rid}" Type="{SLIDE_MASTER_REL_TYPE}" Target="{target}"/>"#
            );
            insert_before(&mut relationships_xml, "</Relationships>", &relationship)?;
            self.files
                .insert(PRESENTATION_RELS.into(), relationships_xml.into_bytes());
            rid
        };
        let entry = format!(r#"<p:sldMasterId id="{id}" r:id="{master_rid}"/>"#);
        insert_before(&mut presentation, "</p:sldMasterIdLst>", &entry)?;
        self.files
            .insert(PRESENTATION.into(), presentation.into_bytes());
        Ok(())
    }

    fn normalize_slide_master_layout_ids(&mut self, master_part: &str) -> Result<HashSet<u32>> {
        let mut used_ids = self.registered_master_and_layout_ids()?;
        let xml = self.part_string(master_part)?;
        let layout_entries =
            Regex::new(r#"<p:sldLayoutId\b[^>]*/>"#).expect("valid slide layout entry regex");
        let previously_used_ids = used_ids.clone();
        let mut incoming_ids = HashSet::new();
        let mut replacements = Vec::new();
        for entry in layout_entries.find_iter(&xml) {
            let id = attr(entry.as_str(), "id").ok_or_else(|| {
                Error::InvalidPackage(format!("{master_part} slide layout entry is missing id"))
            })?;
            let id = parse_slide_layout_id(&id)?;
            let retain = !previously_used_ids.contains(&id) && incoming_ids.insert(id);
            if retain {
                used_ids.insert(id);
            }
            replacements.push((entry.start(), entry.end(), !retain));
        }

        if replacements.iter().all(|(_, _, replace)| !replace) {
            return Ok(used_ids);
        }

        let mut normalized = String::with_capacity(xml.len());
        let mut previous_end = 0;
        for (start, end, replace) in replacements {
            normalized.push_str(&xml[previous_end..start]);
            let entry = &xml[start..end];
            if replace {
                let id = self.allocate_master_or_layout_id(&mut used_ids)?;
                normalized.push_str(&replace_unqualified_id_attr(entry, id));
            } else {
                normalized.push_str(entry);
            }
            previous_end = end;
        }
        normalized.push_str(&xml[previous_end..]);
        self.files
            .insert(master_part.to_string(), normalized.into_bytes());
        Ok(used_ids)
    }

    fn allocate_master_or_layout_id(&mut self, used_ids: &mut HashSet<u32>) -> Result<u32> {
        let mut candidate = self.next_master_id.max(MIN_SLIDE_MASTER_ID);
        loop {
            if used_ids.insert(candidate) {
                self.next_master_id = candidate.saturating_add(1);
                return Ok(candidate);
            }
            if candidate == u32::MAX {
                return Err(Error::InvalidPackage(
                    "slide master/layout id space exhausted".into(),
                ));
            }
            candidate += 1;
        }
    }

    fn allocate_presentation_relationship_id(&mut self, relationships_xml: &str) -> Result<String> {
        let existing = relationship_tags(relationships_xml);
        loop {
            let rid = format!("rId{}", self.next_rel_id);
            if !existing
                .iter()
                .any(|relationship| attr(relationship, "Id").as_deref() == Some(rid.as_str()))
            {
                self.next_rel_id = self.next_rel_id.saturating_add(1);
                return Ok(rid);
            }
            if self.next_rel_id == u32::MAX {
                return Err(Error::InvalidPackage(
                    "presentation relationship id space exhausted".into(),
                ));
            }
            self.next_rel_id += 1;
        }
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

    pub fn shape(self, name: &str) -> Result<ShapeMut<'a>> {
        let xml = self.presentation.slide_xml(self.index)?;
        find_shape_block(&xml, name).map(|_| ShapeMut {
            presentation: self.presentation,
            slide_index: self.index,
            shape_name: name.to_string(),
        })
    }

    pub fn hide_master_graphics(self) -> Result<()> {
        let part = self.presentation.slides[self.index].part.clone();
        let mut xml = self.presentation.part_string(&part)?;
        let flag_re =
            Regex::new(r#"showMasterSp="[^"]*""#).expect("valid master graphics flag regex");
        if let Some(existing) = flag_re.find(&xml) {
            let range = existing.range();
            xml.replace_range(range, r#"showMasterSp="0""#);
        } else if let Some(position) = xml.find("<p:sld ") {
            xml.insert_str(position + "<p:sld ".len(), r#"showMasterSp="0" "#);
        } else {
            return Err(Error::InvalidPackage(format!(
                "{part} has no p:sld root element"
            )));
        }
        self.presentation.files.insert(part, xml.into_bytes());
        Ok(())
    }
}

pub struct ShapeMut<'a> {
    presentation: &'a mut Presentation,
    slide_index: usize,
    shape_name: String,
}

impl ShapeMut<'_> {
    pub fn position(&self) -> Result<(u64, u64, u64, u64)> {
        let part = &self.presentation.slides[self.slide_index].part;
        let xml = self.presentation.part_string(part)?;
        let (_, _, block) = find_shape_block(&xml, &self.shape_name)?;
        let off_re = Regex::new(r#"<a:off x="(\d+)" y="(\d+)"/>"#).expect("valid offset regex");
        let ext_re = Regex::new(r#"<a:ext cx="(\d+)" cy="(\d+)"/>"#).expect("valid extent regex");
        let off = off_re.captures(&block).ok_or_else(|| {
            Error::InvalidPackage(format!("shape {} has no offset", self.shape_name))
        })?;
        let ext = ext_re.captures(&block).ok_or_else(|| {
            Error::InvalidPackage(format!("shape {} has no extent", self.shape_name))
        })?;
        Ok((
            off[1].parse().unwrap_or(0),
            off[2].parse().unwrap_or(0),
            ext[1].parse().unwrap_or(0),
            ext[2].parse().unwrap_or(0),
        ))
    }

    pub fn set_text(self, text: &str) -> Result<()> {
        self.set_rich_text(&[Run::plain(text)])
    }

    pub fn set_rich_text(self, runs: &[Run]) -> Result<()> {
        let part = self.presentation.slides[self.slide_index].part.clone();
        let mut xml = self.presentation.part_string(&part)?;
        let (start, end, block) = find_shape_block(&xml, &self.shape_name)?;
        let updated = replace_text_body(&block, runs)?;
        xml.replace_range(start..end, &updated);
        self.presentation.files.insert(part, xml.into_bytes());
        Ok(())
    }

    pub fn set_position(self, x: u64, y: u64, cx: u64, cy: u64) -> Result<()> {
        let part = self.presentation.slides[self.slide_index].part.clone();
        let mut xml = self.presentation.part_string(&part)?;
        let (start, end, block) = find_shape_block(&xml, &self.shape_name)?;
        let xfrm_re =
            Regex::new(r#"(?s)<a:xfrm>.*?</a:xfrm>"#).expect("valid shape transform regex");
        let xfrm = format!(
            "<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
        );
        let updated = xfrm_re.replace(&block, xfrm.as_str()).to_string();
        xml.replace_range(start..end, &updated);
        self.presentation.files.insert(part, xml.into_bytes());
        Ok(())
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

fn relationship_shape(tag: &str) -> (Option<String>, Option<String>, Option<String>) {
    (attr(tag, "Id"), attr(tag, "Type"), attr(tag, "TargetMode"))
}

fn relationship_shapes_match(source: Option<&Vec<u8>>, destination: Option<&Vec<u8>>) -> bool {
    match (source, destination) {
        (None, None) => true,
        (Some(source), Some(destination)) => {
            let Ok(source) = String::from_utf8(source.clone()) else {
                return false;
            };
            let Ok(destination) = String::from_utf8(destination.clone()) else {
                return false;
            };
            let mut source_shapes = relationship_tags(&source)
                .iter()
                .map(|tag| relationship_shape(tag))
                .collect::<Vec<_>>();
            let mut destination_shapes = relationship_tags(&destination)
                .iter()
                .map(|tag| relationship_shape(tag))
                .collect::<Vec<_>>();
            source_shapes.sort();
            destination_shapes.sort();
            source_shapes == destination_shapes
        }
        _ => false,
    }
}

fn import_part_contents_match(
    source_part: &str,
    source: &[u8],
    destination_part: &str,
    destination: &[u8],
) -> bool {
    if source == destination {
        return true;
    }
    if !source_part.starts_with("ppt/slideMasters/slideMaster")
        || !destination_part.starts_with("ppt/slideMasters/slideMaster")
        || !source_part.ends_with(".xml")
        || !destination_part.ends_with(".xml")
    {
        return false;
    }
    normalized_slide_master_for_comparison(source)
        .zip(normalized_slide_master_for_comparison(destination))
        .is_some_and(|(source, destination)| source == destination)
}

fn normalized_slide_master_for_comparison(bytes: &[u8]) -> Option<String> {
    let xml = String::from_utf8(bytes.to_vec()).ok()?;
    let layout_entries =
        Regex::new(r#"<p:sldLayoutId\b[^>]*/>"#).expect("valid slide layout entry regex");
    Some(
        layout_entries
            .replace_all(&xml, |captures: &regex::Captures<'_>| {
                replace_unqualified_id_attr(
                    captures.get(0).expect("layout entry match").as_str(),
                    MIN_SLIDE_MASTER_ID,
                )
            })
            .to_string(),
    )
}

fn attr(tag: &str, name: &str) -> Option<String> {
    let re = Regex::new(&format!(r#"\b{}="([^"]*)""#, regex::escape(name))).ok()?;
    re.captures(tag).map(|cap| cap[1].to_string())
}

fn parse_slide_master_id(value: &str) -> Result<u32> {
    let id = value
        .parse::<u32>()
        .map_err(|_| Error::InvalidPackage(format!("invalid slide master id {value}")))?;
    if id < MIN_SLIDE_MASTER_ID {
        return Err(Error::InvalidPackage(format!(
            "slide master id {value} is outside the schema range"
        )));
    }
    Ok(id)
}

fn parse_slide_layout_id(value: &str) -> Result<u32> {
    let id = value
        .parse::<u32>()
        .map_err(|_| Error::InvalidPackage(format!("invalid slide layout id {value}")))?;
    if id < MIN_SLIDE_MASTER_ID {
        return Err(Error::InvalidPackage(format!(
            "slide layout id {value} is outside the schema range"
        )));
    }
    Ok(id)
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

fn relationships_part(part: &str) -> String {
    let (directory, filename) = part.rsplit_once('/').unwrap_or(("", part));
    if directory.is_empty() {
        format!("_rels/{filename}.rels")
    } else {
        format!("{directory}/_rels/{filename}.rels")
    }
}

fn owner_part_for_relationships(relationships: &str) -> Option<String> {
    let without_suffix = relationships.strip_suffix(".rels")?;
    Some(without_suffix.replacen("/_rels/", "/", 1))
}

fn auxiliary_part(name: &str) -> bool {
    [
        "/notesSlides/",
        "/notesMasters/",
        "/comments/",
        "/threadedComments/",
        "/people/",
        "/tags/",
    ]
    .iter()
    .any(|segment| name.contains(segment))
        || name.ends_with("/commentAuthors.xml")
        || name == REVISION_INFO
}

fn resolve_part_target(owner_part: &str, target: &str) -> Result<String> {
    let owner_directory = owner_part
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    let joined = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else if owner_directory.is_empty() {
        target.to_string()
    } else {
        format!("{owner_directory}/{target}")
    };
    let mut normalised = Vec::new();
    for component in joined.split('/') {
        match component {
            "" | "." => {}
            ".." => {
                if normalised.pop().is_none() {
                    return Err(Error::InvalidPackage(format!(
                        "relationship target escapes the package: {target}"
                    )));
                }
            }
            value => normalised.push(value),
        }
    }
    Ok(normalised.join("/"))
}

fn relative_part_target(owner_part: &str, target_part: &str) -> String {
    let owner_directory = owner_part
        .rsplit_once('/')
        .map(|(directory, _)| directory)
        .unwrap_or("");
    let owner: Vec<_> = owner_directory
        .split('/')
        .filter(|component| !component.is_empty())
        .collect();
    let target: Vec<_> = target_part
        .split('/')
        .filter(|component| !component.is_empty())
        .collect();
    let common = owner
        .iter()
        .zip(&target)
        .take_while(|(left, right)| left == right)
        .count();
    let mut relative = vec![".."; owner.len().saturating_sub(common)];
    relative.extend(target[common..].iter().copied());
    relative.join("/")
}

fn drop_import_relationship(relationship_type: &str) -> bool {
    [
        "/notesSlide",
        "/notesMaster",
        "/comments",
        "/commentAuthors",
        "/tags",
        "/customXml",
        "/slide",
        "/oleObject",
        "/activeXControl",
        "/vbaProject",
    ]
    .iter()
    .any(|suffix| relationship_type.ends_with(suffix))
}

fn drop_auxiliary_relationship(relationship_type: &str) -> bool {
    [
        "/notesSlide",
        "/notesMaster",
        "/comments",
        "/commentAuthors",
        "/tags",
        "/customXml",
        "/oleObject",
        "/activeXControl",
        "/vbaProject",
        "/revisionInfo",
    ]
    .iter()
    .any(|suffix| relationship_type.ends_with(suffix))
}

fn strip_relationship_reference(xml: &str, relationship_id: &str) -> String {
    let notes_master = Regex::new(&format!(
        r#"<p:notesMasterId\b[^>]*r:id="{}"[^>]*/>\s*"#,
        regex::escape(relationship_id)
    ))
    .expect("valid notes master relationship regex");
    let empty_notes_master_list = Regex::new(r#"(?s)<p:notesMasterIdLst>\s*</p:notesMasterIdLst>"#)
        .expect("valid empty notes master list regex");
    let reference = Regex::new(&format!(
        r#"\s+r:(?:id|embed|link)="{}""#,
        regex::escape(relationship_id)
    ))
    .expect("valid relationship reference regex");
    let xml = notes_master.replace_all(xml, "");
    let xml = empty_notes_master_list.replace_all(&xml, "");
    reference.replace_all(&xml, "").to_string()
}

fn replace_xml_attr(tag: &str, name: &str, value: &str) -> String {
    let attribute = Regex::new(&format!(r#"\b{}="[^"]*""#, regex::escape(name)))
        .expect("valid XML attribute regex");
    attribute
        .replace(tag, format!(r#"{name}="{}""#, xml_escape(value)))
        .to_string()
}

fn replace_unqualified_id_attr(tag: &str, value: u32) -> String {
    let attribute = Regex::new(r#"\sid="[^"]*""#).expect("valid unqualified id regex");
    attribute
        .replace(tag, format!(r#" id="{value}""#))
        .to_string()
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

fn find_shape_block(xml: &str, name: &str) -> Result<(usize, usize, String)> {
    let re = Regex::new(r#"(?s)<p:sp>.*?</p:sp>"#).expect("valid regex");
    let properties_re = Regex::new(r#"<p:cNvPr\b[^>]*/?>"#).expect("valid shape properties regex");
    for matched in re.find_iter(xml) {
        let block = matched.as_str();
        let Some(properties) = properties_re.find(block) else {
            continue;
        };
        if attr(properties.as_str(), "name").as_deref() == Some(name) {
            return Ok((matched.start(), matched.end(), block.to_string()));
        }
    }
    Err(Error::InvalidPackage(format!(
        "slide does not contain shape {name}"
    )))
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
            "{prefix}{}</p:txBody>",
            paragraphs_xml(runs, &paragraph_properties, run_properties.as_deref())
        )
    } else {
        format!(
            "<p:txBody><a:bodyPr/><a:lstStyle/>{}</p:txBody>",
            paragraphs_xml(runs, "", None)
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

fn paragraphs_xml(runs: &[Run], paragraph_properties: &str, default_rpr: Option<&str>) -> String {
    let mut paragraphs: Vec<Vec<Run>> = vec![Vec::new()];
    for run in runs {
        let mut lines = run.text.split('\n').peekable();
        while let Some(line) = lines.next() {
            if !line.is_empty() {
                let mut fragment = run.clone();
                fragment.text = line.to_string();
                paragraphs.last_mut().expect("one paragraph").push(fragment);
            }
            if lines.peek().is_some() {
                paragraphs.push(Vec::new());
            }
        }
    }
    paragraphs
        .into_iter()
        .map(|paragraph| {
            format!(
                "<a:p>{paragraph_properties}{}</a:p>",
                runs_xml(&paragraph, default_rpr)
            )
        })
        .collect()
}

fn runs_xml(runs: &[Run], default_rpr: Option<&str>) -> String {
    runs.iter()
        .map(|run| {
            let attrs = run_attrs(run);
            let mut rpr = default_rpr
                .map(str::to_string)
                .unwrap_or_else(|| "<a:rPr lang=\"en-GB\"/>".to_string());
            if let Some(typeface) = &run.typeface {
                rpr = merge_run_style(&rpr, typeface, run.color.as_deref().unwrap_or("000000"));
            }
            rpr = merge_run_attrs(&rpr, &attrs);
            format!("<a:r>{rpr}<a:t>{}</a:t></a:r>", xml_escape(&run.text))
        })
        .collect::<String>()
}

fn merge_run_style(rpr: &str, typeface: &str, color: &str) -> String {
    let direct_style =
        Regex::new(r#"\s(?:baseline|b|i|u)=\"[^\"]*\""#).expect("valid direct run style regex");
    let mut out = direct_style.replace_all(rpr, "").to_string();
    if out.ends_with("/>") {
        out.truncate(out.len() - 2);
        out.push('>');
        out.push_str("</a:rPr>");
    }
    out = remove_rpr_children(
        &out,
        &[
            "solidFill",
            "noFill",
            "gradFill",
            "blipFill",
            "pattFill",
            "grpFill",
        ],
    );
    out = remove_rpr_children(&out, &["latin", "ea", "cs"]);
    let fill = if color.eq_ignore_ascii_case("accent1") {
        "<a:schemeClr val=\"accent1\"/>".to_string()
    } else {
        format!("<a:srgbClr val=\"{}\"/>", xml_escape(color))
    };
    let fill_style = format!("<a:solidFill>{fill}</a:solidFill>");
    let script_fonts = format!(
        "<a:latin typeface=\"{typeface}\"/>\
         <a:ea typeface=\"{typeface}\"/>\
         <a:cs typeface=\"{typeface}\"/>",
        typeface = xml_escape(typeface)
    );
    insert_before_rpr_children(
        &mut out,
        &fill_style,
        &[
            "effectLst",
            "effectDag",
            "highlight",
            "uLn",
            "uLnTx",
            "uFill",
            "uFillTx",
            "latin",
            "ea",
            "cs",
            "sym",
            "hlinkClick",
            "hlinkMouseOver",
            "rtl",
            "extLst",
        ],
    );
    insert_before_rpr_children(
        &mut out,
        &script_fonts,
        &["sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"],
    );
    out
}

fn remove_rpr_children(xml: &str, names: &[&str]) -> String {
    let children = rpr_children(xml);
    let mut out = String::with_capacity(xml.len());
    let mut cursor = 0;
    for (name, start, end) in children {
        if names.contains(&name.as_str()) {
            out.push_str(&xml[cursor..start]);
            cursor = end;
        }
    }
    out.push_str(&xml[cursor..]);
    out
}

fn insert_before_rpr_children(xml: &mut String, insertion: &str, names: &[&str]) {
    let index = rpr_children(xml)
        .into_iter()
        .find(|(name, _, _)| names.contains(&name.as_str()))
        .map(|(_, start, _)| start)
        .or_else(|| xml.rfind("</a:rPr>"));
    if let Some(index) = index {
        xml.insert_str(index, insertion);
    }
}

fn rpr_children(xml: &str) -> Vec<(String, usize, usize)> {
    let Some(open_end) = xml.find('>') else {
        return Vec::new();
    };
    let Some(close_start) = xml.rfind("</a:rPr>") else {
        return Vec::new();
    };
    let inner = &xml[open_end + 1..close_start];
    let mut children = Vec::new();
    let mut cursor = 0;
    let mut depth = 0_usize;
    let mut current = None;

    while let Some(relative_start) = inner[cursor..].find('<') {
        let start = cursor + relative_start;
        let Some(relative_end) = xml_tag_end(&inner[start..]) else {
            break;
        };
        let end = start + relative_end;
        let tag = &inner[start..end];
        if tag.starts_with("<!--") || tag.starts_with("<?") || tag.starts_with("<!") {
            cursor = end;
            continue;
        }
        if tag.starts_with("</") {
            depth = depth.saturating_sub(1);
            if depth == 0 {
                if let Some((name, child_start)) = current.take() {
                    children.push((name, child_start + open_end + 1, end + open_end + 1));
                }
            }
        } else {
            if depth == 0 {
                current = Some((xml_tag_name(tag).unwrap_or_default(), start));
            }
            if !tag.trim_end().ends_with("/>") {
                depth += 1;
            } else if depth == 0 {
                if let Some((name, child_start)) = current.take() {
                    children.push((name, child_start + open_end + 1, end + open_end + 1));
                }
            }
        }
        cursor = end;
    }
    children
}

fn xml_tag_end(xml: &str) -> Option<usize> {
    let mut quote = None;
    for (index, character) in xml.char_indices() {
        match (quote, character) {
            (Some(expected), value) if value == expected => quote = None,
            (None, '\"') | (None, '\'') => quote = Some(character),
            (None, '>') => return Some(index + 1),
            _ => {}
        }
    }
    None
}

fn xml_tag_name(tag: &str) -> Option<String> {
    tag.strip_prefix('<')?
        .trim_start_matches('/')
        .split(|character: char| {
            character.is_ascii_whitespace() || character == '>' || character == '/'
        })
        .next()
        .filter(|name| !name.is_empty())
        .map(|name| name.strip_prefix("a:").unwrap_or(name).to_string())
}

fn run_attrs(run: &Run) -> String {
    let mut attrs = String::new();
    if let Some(font_size) = run.font_size {
        attrs.push_str(&format!(" sz=\"{font_size}\""));
    }
    if run.superscript {
        attrs.push_str(" baseline=\"30000\"");
    }
    if run.bold {
        attrs.push_str(" b=\"1\"");
    }
    if run.italic {
        attrs.push_str(" i=\"1\"");
    }
    if run.underline {
        attrs.push_str(" u=\"sng\"");
    }
    attrs
}

fn merge_run_attrs(default_rpr: &str, attrs: &str) -> String {
    if attrs.is_empty() {
        return default_rpr.to_string();
    }
    let mut out = default_rpr.to_string();
    let attributes =
        Regex::new(r#"\s([A-Za-z][A-Za-z0-9]*)="([^"]*)""#).expect("valid run attribute regex");
    for captures in attributes.captures_iter(attrs) {
        let name = &captures[1];
        let value = &captures[2];
        let existing = Regex::new(&format!(r#"\s{}="[^"]*""#, regex::escape(name)))
            .expect("valid existing run attribute regex");
        let replacement = format!(r#" {name}="{value}""#);
        if existing.is_match(&out) {
            out = existing.replace(&out, replacement.as_str()).to_string();
        } else if let Some(idx) = out.find('>') {
            let insert_at = idx.saturating_sub(usize::from(out[..idx].ends_with('/')));
            out.insert_str(insert_at, &replacement);
        }
    }
    out
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

#[cfg(test)]
mod tests {
    use super::merge_run_style;

    #[test]
    fn styled_runs_replace_all_script_font_fallbacks() {
        let rpr = concat!(
            r#"<a:rPr lang="en-GB"><a:latin typeface="Calibri"/>"#,
            r#"<a:ea typeface="+mn-ea"/><a:cs typeface="Calibri"/>"#,
            r#"<a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr>"#,
        );

        let styled = merge_run_style(rpr, "Arial Black", "000000");

        assert_eq!(styled.matches(r#"typeface="Arial Black""#).count(), 3);
        assert!(!styled.contains("Calibri"));
        assert!(!styled.contains("+mn-ea"));
    }

    #[test]
    fn styled_runs_keep_drawingml_character_property_order() {
        let rpr = concat!(
            r#"<a:rPr lang="en-GB"><a:ln><a:noFill/></a:ln>"#,
            r#"<a:effectLst/><a:uLnTx/><a:uFillTx/>"#,
            r#"<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>"#,
            r#"<a:latin typeface="Calibri"/><a:ea typeface="Calibri"/><a:cs typeface="Calibri"/>"#,
            r#"<a:hlinkClick r:id="rId1"/><a:rtl/><a:extLst/></a:rPr>"#,
        );

        let styled = merge_run_style(rpr, "Arial Black", "000000");
        let positions = [
            styled.find("<a:ln").unwrap(),
            styled.find("<a:solidFill").unwrap(),
            styled.find("<a:effectLst").unwrap(),
            styled.find("<a:uLnTx").unwrap(),
            styled.find("<a:uFillTx").unwrap(),
            styled.find("<a:latin").unwrap(),
            styled.find("<a:ea").unwrap(),
            styled.find("<a:cs").unwrap(),
            styled.find("<a:hlinkClick").unwrap(),
            styled.find("<a:rtl").unwrap(),
            styled.find("<a:extLst").unwrap(),
        ];
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
        assert_eq!(styled.matches("<a:solidFill>").count(), 1);
        assert!(styled.contains("<a:srgbClr val=\"000000\"/>"));
        assert!(styled.contains("<a:ln><a:noFill/></a:ln>"));
        assert_eq!(styled.matches(r#"typeface="Arial Black""#).count(), 3);
    }
}
