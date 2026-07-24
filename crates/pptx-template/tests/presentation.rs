use pptx_template::{Presentation, Run};
use regex::Regex;
use std::collections::{BTreeMap, HashSet};
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

const TEMPLATE: &[u8] = include_bytes!("../../deck-builder/assets/template.pptx");
const DISTINCT_MASTER_ID: u32 = 2_147_484_000;

#[test]
fn round_trip_preserves_slide_relationship_integrity() {
    let pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    let bytes = pres.save_bytes().expect("save template");
    let reopened = Presentation::open_bytes(&bytes).expect("reopen saved template");

    assert_eq!(reopened.slide_count(), 49);
    reopened.validate().expect("structural validation");
}

#[test]
fn shape_position_reads_geometry_and_master_graphics_can_be_hidden() {
    let mut pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    let position = pres
        .slide_mut(16)
        .unwrap()
        .shape("TextShape 2")
        .unwrap()
        .position()
        .unwrap();
    assert_eq!(position, (397_041, 1_267_778, 9_683_583, 4_506_298));

    assert!(!pres.slide_xml(16).unwrap().contains("showMasterSp"));
    pres.slide_mut(16).unwrap().hide_master_graphics().unwrap();
    let xml = pres.slide_xml(16).unwrap();
    assert!(xml.contains("showMasterSp=\"0\""));
    pres.validate().expect("package still validates");
    // Hiding twice keeps a single, still-disabled flag.
    pres.slide_mut(16).unwrap().hide_master_graphics().unwrap();
    assert_eq!(
        pres.slide_xml(16).unwrap().matches("showMasterSp").count(),
        1
    );
}

#[test]
fn canonical_template_has_exact_twpc_dimensions() {
    let pres = Presentation::open_bytes(TEMPLATE).expect("template opens");
    assert_eq!(pres.slide_size().unwrap(), (10_080_625, 7_559_675));
    pres.validate_song_source((10_080_625, 7_559_675))
        .expect("canonical template is a valid source package");
}

#[test]
fn removing_auxiliary_content_removes_notes_master_declaration() {
    let mut pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    pres.remove_auxiliary_content()
        .expect("remove auxiliary content");
    let bytes = pres.save_bytes().expect("save cleaned template");

    let mut archive = ZipArchive::new(std::io::Cursor::new(&bytes)).expect("open saved package");
    let mut presentation_xml = String::new();
    archive
        .by_name("ppt/presentation.xml")
        .expect("presentation XML exists")
        .read_to_string(&mut presentation_xml)
        .expect("presentation XML is UTF-8");
    assert!(!presentation_xml.contains("<p:notesMasterId/>"));
    assert!(!presentation_xml.contains("<p:notesMasterId"));

    let reopened = Presentation::open_bytes(&bytes).expect("reopen cleaned package");
    reopened.validate().expect("cleaned package validates");
}

#[test]
fn removing_auxiliary_content_removes_revision_info() {
    let mut pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    pres.remove_auxiliary_content()
        .expect("remove auxiliary content");
    let bytes = pres.save_bytes().expect("save cleaned template");

    let mut archive = ZipArchive::new(std::io::Cursor::new(&bytes)).expect("open saved package");
    let names = (0..archive.len())
        .map(|index| {
            archive
                .by_index(index)
                .expect("package part")
                .name()
                .to_string()
        })
        .collect::<Vec<_>>();
    assert!(!names.contains(&"ppt/revisionInfo.xml".to_string()));

    let mut presentation_rels = String::new();
    archive
        .by_name("ppt/_rels/presentation.xml.rels")
        .expect("presentation relationships exist")
        .read_to_string(&mut presentation_rels)
        .expect("presentation relationships are UTF-8");
    assert!(!presentation_rels.contains("revisionInfo"));

    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")
        .expect("content types exist")
        .read_to_string(&mut content_types)
        .expect("content types are UTF-8");
    assert!(!content_types.contains("revisionInfo"));

    let reopened = Presentation::open_bytes(&bytes).expect("reopen cleaned package");
    reopened.validate().expect("cleaned package validates");
}

#[test]
fn can_clone_named_twpc_shapes_set_text_and_reorder() {
    let mut pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    let first = pres.clone_slide(0).expect("clone first slide");
    let second = pres.clone_slide(0).expect("clone second slide");

    pres.slide_mut(first)
        .expect("first slide")
        .shape("TextShape 1")
        .expect("first title")
        .set_text("First cloned title")
        .expect("set first title");
    pres.slide_mut(second)
        .expect("second slide")
        .shape("TextShape 1")
        .expect("second title")
        .set_text("Second cloned title")
        .expect("set second title");

    let mut order: Vec<_> = (0..pres.slide_count()).collect();
    order.swap(first, second);
    pres.reorder(&order).expect("reorder");
    let bytes = pres.save_bytes().expect("save");
    let reopened = Presentation::open_bytes(&bytes).expect("reopen");

    assert_eq!(reopened.slide_count(), 51);
    assert!(reopened
        .slide_text(first)
        .expect("text")
        .starts_with("Second cloned title"));
    assert!(reopened
        .slide_text(second)
        .expect("text")
        .starts_with("First cloned title"));
    reopened.validate().expect("structural validation");
}

#[test]
fn rich_text_writes_superscript_runs() {
    let mut pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    let idx = pres.clone_slide(2).expect("clone call-to-worship slide");

    pres.slide_mut(idx)
        .expect("slide")
        .shape("Text Placeholder 2")
        .expect("body placeholder")
        .set_rich_text(&[
            Run::plain("In the beginning "),
            Run::superscript("1"),
            Run::plain(" God created"),
        ])
        .expect("set rich text");

    let xml = pres.slide_xml(idx).expect("slide xml");
    assert!(xml.contains("baseline=\"30000\""));
    assert!(xml.contains("<a:t>1</a:t>"));
    assert_eq!(xml.matches("<a:p>").count(), 2, "title plus scripture body");
}

#[test]
fn multiline_text_creates_real_powerpoint_paragraphs() {
    let mut pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    let idx = pres.clone_slide(16).expect("clone psalm slide");
    pres.slide_mut(idx)
        .expect("slide")
        .shape("TextShape 2")
        .expect("psalm body")
        .set_text("First line\nSecond line\nThird line")
        .expect("set multiline body");

    let xml = pres.slide_xml(idx).expect("slide xml");
    assert_eq!(
        xml.matches("<a:p>").count(),
        4,
        "three body lines plus title"
    );
    assert!(!xml.contains("First line\nSecond line"));
}

#[test]
fn imports_complete_slides_from_another_presentation_in_order() {
    let source = Presentation::open_bytes(TEMPLATE).expect("open source presentation");
    let mut destination = Presentation::open_bytes(TEMPLATE).expect("open destination");
    let original_count = destination.slide_count();
    let imported = destination
        .import_slides(TEMPLATE)
        .expect("import source slides");

    assert_eq!(imported.len(), source.slide_count());
    assert_eq!(
        destination.slide_count(),
        original_count + source.slide_count()
    );
    for (source_index, destination_index) in imported.into_iter().enumerate() {
        assert_eq!(
            destination.slide_text(destination_index).unwrap(),
            source.slide_text(source_index).unwrap(),
            "slide {source_index} text and order"
        );
    }

    let bytes = destination
        .save_bytes()
        .expect("save imported presentation");
    let reopened = Presentation::open_bytes(&bytes).expect("reopen imported presentation");
    reopened
        .validate()
        .expect("imported relationship graph validates");
}

#[test]
fn imports_and_registers_distinct_slide_master_once() {
    let source_bytes = source_with_distinct_master(DISTINCT_MASTER_ID);
    let mut destination = Presentation::open_bytes(TEMPLATE).expect("open destination");

    destination
        .import_slides(&source_bytes)
        .expect("import source with distinct master");
    destination
        .import_slides(&source_bytes)
        .expect("import source again");

    let bytes = destination
        .save_bytes()
        .expect("save imported presentation");
    let mut archive = ZipArchive::new(Cursor::new(&bytes)).expect("open saved package");
    let layout_part_count = (0..archive.len())
        .filter(|index| {
            let name = archive
                .by_index(*index)
                .expect("package part")
                .name()
                .to_string();
            name.starts_with("ppt/slideLayouts/slideLayout") && name.ends_with(".xml")
        })
        .count();
    let mut presentation = String::new();
    archive
        .by_name("ppt/presentation.xml")
        .expect("presentation XML exists")
        .read_to_string(&mut presentation)
        .expect("presentation XML is UTF-8");
    let mut presentation_rels = String::new();
    archive
        .by_name("ppt/_rels/presentation.xml.rels")
        .expect("presentation relationships exist")
        .read_to_string(&mut presentation_rels)
        .expect("presentation relationships are UTF-8");
    assert_eq!(presentation.matches("<p:sldMasterId ").count(), 2);
    assert_eq!(
        presentation_rels
            .matches("/relationships/slideMaster\"")
            .count(),
        2
    );
    assert!(presentation_rels.contains("Target=\"slideMasters/slideMaster2.xml\""));

    let master_entries = Regex::new(r#"<p:sldMasterId\b[^>]*/>"#).unwrap();
    let master_ids = master_entries
        .find_iter(&presentation)
        .filter_map(|entry| xml_attr(entry.as_str(), "id"))
        .collect::<Vec<_>>();
    let layout_entries = Regex::new(r#"<p:sldLayoutId\b[^>]*/>"#).unwrap();
    let mut layout_ids = Vec::new();
    let mut imported_master = String::new();
    for index in 0..archive.len() {
        let mut part = archive.by_index(index).expect("package part");
        if part.name().starts_with("ppt/slideMasters/") && part.name().ends_with(".xml") {
            let mut xml = String::new();
            part.read_to_string(&mut xml)
                .expect("slide master XML is UTF-8");
            if part.name() == "ppt/slideMasters/slideMaster2.xml" {
                imported_master = xml.clone();
            }
            layout_ids.extend(
                layout_entries
                    .find_iter(&xml)
                    .filter_map(|entry| xml_attr(entry.as_str(), "id")),
            );
        }
    }
    let all_ids = master_ids.iter().chain(&layout_ids).collect::<HashSet<_>>();
    assert_eq!(master_ids.len(), 2);
    assert_eq!(layout_ids.len(), 24);
    assert_eq!(all_ids.len(), master_ids.len() + layout_ids.len());
    assert_eq!(layout_part_count, 24, "repeated import must reuse layouts");
    let imported_layout_rids = layout_entries
        .find_iter(&imported_master)
        .filter_map(|entry| xml_attr(entry.as_str(), "r:id"))
        .collect::<Vec<_>>();
    assert_eq!(
        imported_layout_rids,
        (1..=11).map(|id| format!("rId{id}")).collect::<Vec<_>>()
    );

    let reopened = Presentation::open_bytes(&bytes).expect("reopen imported presentation");
    reopened
        .validate()
        .expect("all slide-reachable masters are registered");
}

#[test]
fn validation_rejects_unregistered_reachable_slide_master() {
    let bytes = source_without_distinct_master_registration();
    let presentation = Presentation::open_bytes(&bytes).expect("open invalid source package");
    let error = presentation
        .validate()
        .expect_err("unregistered reachable master must fail validation");
    assert!(error
        .to_string()
        .contains("reaches unregistered slide master"));
}

#[test]
fn validation_rejects_slide_master_layout_id_collision() {
    let presentation =
        Presentation::open_bytes(&source_with_distinct_master(2_147_483_661)).expect("open");
    let error = presentation
        .validate()
        .expect_err("master/layout ID collision must fail validation");
    assert!(error
        .to_string()
        .contains("duplicates a master or layout id"));
}

#[test]
fn validation_rejects_duplicate_layout_ids() {
    let source = source_with_distinct_master(DISTINCT_MASTER_ID);
    let bytes = rewrite_zip_part(source, "ppt/slideMasters/slideMaster2.xml", |xml| {
        xml.replacen("id=\"2147483662\"", "id=\"2147483661\"", 1)
    });
    let presentation = Presentation::open_bytes(&bytes).expect("open malformed package");
    let error = presentation
        .validate()
        .expect_err("duplicate layout IDs must fail validation");
    assert!(error
        .to_string()
        .contains("duplicates a master or layout id"));
}

#[test]
fn validation_rejects_r15_layout_ids_overlapping_across_registered_masters() {
    let source = source_with_distinct_master(DISTINCT_MASTER_ID);
    let mut destination = Presentation::open_bytes(TEMPLATE).expect("open destination");
    destination
        .import_slides(&source)
        .expect("import source with distinct master");
    let generated = destination.save_bytes().expect("save generated package");
    let mut overlapping_ids = 2_147_483_661_u32..=2_147_483_671;
    let r15_like = rewrite_zip_part(generated, "ppt/slideMasters/slideMaster2.xml", |xml| {
        let layout_entry =
            Regex::new(r#"<p:sldLayoutId\b[^>]*/>"#).expect("valid layout entry regex");
        layout_entry
            .replace_all(&xml, |captures: &regex::Captures<'_>| {
                let id = overlapping_ids.next().expect("eleven imported layouts");
                replace_xml_id(captures.get(0).unwrap().as_str(), id)
            })
            .into_owned()
    });
    assert!(overlapping_ids.next().is_none());

    let presentation = Presentation::open_bytes(&r15_like).expect("open r15-like package");
    let error = presentation
        .validate()
        .expect_err("cross-master duplicate layout IDs must fail validation");
    assert!(
        error
            .to_string()
            .contains("duplicates a master or layout id"),
        "unexpected validation error: {error}"
    );
}

#[test]
fn import_reports_master_layout_id_exhaustion() {
    let destination = rewrite_zip_part(TEMPLATE.to_vec(), "ppt/presentation.xml", |xml| {
        xml.replacen("id=\"2147483660\"", "id=\"4294967295\"", 1)
    });
    let source = source_with_distinct_master(DISTINCT_MASTER_ID);
    let mut presentation = Presentation::open_bytes(&destination).expect("open destination");
    let error = presentation
        .import_slides(&source)
        .expect_err("combined ID space must be exhausted deterministically");
    assert!(error
        .to_string()
        .contains("slide master/layout id space exhausted"));
}

#[test]
fn normalization_preserves_noncolliding_incoming_layout_ids() {
    let source = rewrite_zip_part(
        source_with_distinct_master(DISTINCT_MASTER_ID),
        "ppt/slideMasters/slideMaster2.xml",
        |xml| xml.replacen("id=\"2147483671\"", "id=\"2147483672\"", 1),
    );
    let mut destination = Presentation::open_bytes(TEMPLATE).expect("open destination");
    destination
        .import_slides(&source)
        .expect("normalize imported master IDs");
    let generated = destination.save_bytes().expect("save generated package");
    let mut archive = ZipArchive::new(Cursor::new(generated)).expect("open generated package");
    let mut imported_master = String::new();
    archive
        .by_name("ppt/slideMasters/slideMaster2.xml")
        .expect("imported master exists")
        .read_to_string(&mut imported_master)
        .expect("imported master is UTF-8");
    let layout_entry = Regex::new(r#"<p:sldLayoutId\b[^>]*/>"#).unwrap();
    let r_id_11 = layout_entry
        .find_iter(&imported_master)
        .find(|entry| xml_attr(entry.as_str(), "r:id").as_deref() == Some("rId11"))
        .expect("rId11 layout entry");
    assert_eq!(
        xml_attr(r_id_11.as_str(), "id").as_deref(),
        Some("2147483672")
    );
}

#[test]
fn imported_master_receives_its_own_theme_part() {
    let source = source_with_distinct_master(DISTINCT_MASTER_ID);
    let mut destination = Presentation::open_bytes(TEMPLATE).expect("open destination");
    destination
        .import_slides(&source)
        .expect("import source with distinct master");
    destination
        .import_slides(&source)
        .expect("import source again");
    let generated = destination.save_bytes().expect("save generated package");
    let mut archive = ZipArchive::new(Cursor::new(generated)).expect("open generated package");

    let theme_relationship =
        Regex::new(r#"<Relationship\b[^>]*Type="[^"]*/theme"[^>]*/>"#).unwrap();
    let mut theme_targets = Vec::new();
    for master in ["slideMaster1", "slideMaster2"] {
        let mut rels = String::new();
        archive
            .by_name(&format!("ppt/slideMasters/_rels/{master}.xml.rels"))
            .expect("master relationships exist")
            .read_to_string(&mut rels)
            .expect("master relationships are UTF-8");
        let theme = theme_relationship
            .find(&rels)
            .expect("master references a theme");
        theme_targets.push(
            xml_attr(theme.as_str(), "Target")
                .expect("theme relationship has target")
                .trim_start_matches("../")
                .to_string(),
        );
    }
    assert_ne!(
        theme_targets[0], theme_targets[1],
        "each registered master must own a distinct theme part"
    );

    let imported_theme = format!("ppt/{}", theme_targets[1]);
    archive
        .by_name(&imported_theme)
        .expect("imported theme part exists");
    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")
        .expect("content types exist")
        .read_to_string(&mut content_types)
        .expect("content types are UTF-8");
    assert!(
        content_types.contains(&format!("PartName=\"/{imported_theme}\"")),
        "imported theme part must have a content-type override"
    );
}

#[test]
fn validation_rejects_registered_masters_sharing_a_theme_part() {
    let source = source_with_distinct_master(DISTINCT_MASTER_ID);
    let mut destination = Presentation::open_bytes(TEMPLATE).expect("open destination");
    destination
        .import_slides(&source)
        .expect("import source with distinct master");
    let generated = destination.save_bytes().expect("save generated package");
    let r16_like = rewrite_zip_part(
        generated,
        "ppt/slideMasters/_rels/slideMaster2.xml.rels",
        |xml| {
            Regex::new(r#"(Type="[^"]*/theme"[^>]*Target=")[^"]*(")"#)
                .expect("valid theme relationship regex")
                .replace(&xml, "${1}../theme/theme1.xml${2}")
                .into_owned()
        },
    );

    let presentation = Presentation::open_bytes(&r16_like).expect("open r16-like package");
    let error = presentation
        .validate()
        .expect_err("masters sharing one theme part must fail validation");
    assert!(
        error.to_string().contains("share theme part"),
        "unexpected validation error: {error}"
    );
}

fn source_without_distinct_master_registration() -> Vec<u8> {
    let source = source_with_distinct_master(DISTINCT_MASTER_ID);
    let mut input = ZipArchive::new(Cursor::new(source)).expect("open source package");
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for index in 0..input.len() {
        let mut file = input.by_index(index).expect("read source part");
        let name = file.name().to_string();
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).expect("read source bytes");
        if name == "ppt/presentation.xml" {
            let xml = String::from_utf8(bytes).expect("presentation XML is UTF-8");
            let entry = Regex::new(&format!(
                r#"<p:sldMasterId\b id="{DISTINCT_MASTER_ID}"[^>]*/>"#
            ))
            .expect("valid master entry regex");
            bytes = entry.replace(&xml, "").into_owned().into_bytes();
        }
        writer.start_file(name, options).expect("write source part");
        writer.write_all(&bytes).expect("write source bytes");
    }
    writer.finish().expect("finish source package").into_inner()
}

fn source_with_distinct_master(master_id: u32) -> Vec<u8> {
    let mut input = ZipArchive::new(Cursor::new(TEMPLATE)).expect("open template package");
    let mut files = BTreeMap::new();
    for index in 0..input.len() {
        let mut file = input.by_index(index).expect("read template part");
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)
            .expect("read template part bytes");
        files.insert(file.name().to_string(), bytes);
    }

    let mut master = String::from_utf8(
        files
            .get("ppt/slideMasters/slideMaster1.xml")
            .expect("template master")
            .clone(),
    )
    .expect("master XML is UTF-8")
    .replacen("preserve=\"1\"", "preserve=\"0\"", 1);
    let unused_layout_entries = Regex::new(r#"<p:sldLayoutId\b[^>]*r:id=\"rId(?:12|13)\"[^>]*/>"#)
        .expect("valid unused layout entry regex");
    master = unused_layout_entries.replace_all(&master, "").into_owned();
    files.insert(
        "ppt/slideMasters/slideMaster2.xml".into(),
        master.into_bytes(),
    );
    let master_rels = String::from_utf8(
        files
            .get("ppt/slideMasters/_rels/slideMaster1.xml.rels")
            .expect("template master relationships")
            .clone(),
    )
    .expect("master relationships are UTF-8");
    let unused_layout_relationships =
        Regex::new(r#"<Relationship\b[^>]*Id=\"rId(?:12|13)\"[^>]*/>"#)
            .expect("valid unused layout relationship regex");
    files.insert(
        "ppt/slideMasters/_rels/slideMaster2.xml.rels".into(),
        unused_layout_relationships
            .replace_all(&master_rels, "")
            .into_owned()
            .into_bytes(),
    );

    for layout_number in 1..=11 {
        let part = format!("ppt/slideLayouts/_rels/slideLayout{layout_number}.xml.rels");
        let layout_rels = String::from_utf8(
            files
                .get(&part)
                .expect("template layout relationships")
                .clone(),
        )
        .expect("layout relationships are UTF-8")
        .replace(
            "../slideMasters/slideMaster1.xml",
            "../slideMasters/slideMaster2.xml",
        );
        files.insert(part, layout_rels.into_bytes());
    }

    let mut presentation = String::from_utf8(
        files
            .get("ppt/presentation.xml")
            .expect("template presentation")
            .clone(),
    )
    .expect("presentation XML is UTF-8");
    let slide_list_start =
        presentation.find("<p:sldIdLst>").expect("slide list start") + "<p:sldIdLst>".len();
    let slide_list_end = presentation.find("</p:sldIdLst>").expect("slide list end");
    let slide_entry = Regex::new(r#"<p:sldId\b[^>]*/>"#).expect("valid slide entry regex");
    let second_slide = slide_entry
        .find_iter(&presentation[slide_list_start..slide_list_end])
        .nth(1)
        .expect("second slide entry")
        .as_str()
        .to_string();
    presentation.replace_range(slide_list_start..slide_list_end, &second_slide);
    let master_list = Regex::new(r#"(?s)<p:sldMasterIdLst>.*?</p:sldMasterIdLst>"#)
        .expect("valid master list regex");
    presentation = master_list
        .replace(
            &presentation,
            format!(
                "<p:sldMasterIdLst><p:sldMasterId id=\"{master_id}\" r:id=\"rId61\"/></p:sldMasterIdLst>"
            ),
        )
        .into_owned();
    files.insert("ppt/presentation.xml".into(), presentation.into_bytes());

    let mut presentation_rels = String::from_utf8(
        files
            .get("ppt/_rels/presentation.xml.rels")
            .expect("template presentation relationships")
            .clone(),
    )
    .expect("presentation relationships are UTF-8");
    let master_relationship = Regex::new(r#"<Relationship\b[^>]*Type="[^"]*/slideMaster"[^>]*/>"#)
        .expect("valid master relationship regex");
    presentation_rels = master_relationship
        .replace_all(&presentation_rels, "")
        .into_owned();
    presentation_rels = presentation_rels.replacen(
        "</Relationships>",
        "<Relationship Id=\"rId61\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster2.xml\"/></Relationships>",
        1,
    );
    files.insert(
        "ppt/_rels/presentation.xml.rels".into(),
        presentation_rels.into_bytes(),
    );

    let mut content_types = String::from_utf8(
        files
            .get("[Content_Types].xml")
            .expect("template content types")
            .clone(),
    )
    .expect("content types are UTF-8");
    content_types = content_types.replacen(
        "</Types>",
        "<Override PartName=\"/ppt/slideMasters/slideMaster2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/></Types>",
        1,
    );
    files.insert("[Content_Types].xml".into(), content_types.into_bytes());

    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for (name, bytes) in files {
        writer.start_file(name, options).expect("write source part");
        writer.write_all(&bytes).expect("write source bytes");
    }
    writer.finish().expect("finish source package").into_inner()
}

fn rewrite_zip_part(
    source: Vec<u8>,
    part_name: &str,
    rewrite: impl FnOnce(String) -> String,
) -> Vec<u8> {
    let mut input = ZipArchive::new(Cursor::new(source)).expect("open source package");
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    let mut rewrite = Some(rewrite);
    for index in 0..input.len() {
        let mut file = input.by_index(index).expect("read source part");
        let name = file.name().to_string();
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes).expect("read source bytes");
        if name == part_name {
            let xml = String::from_utf8(bytes).expect("rewritten part is UTF-8");
            bytes = rewrite.take().expect("rewrite part once")(xml).into_bytes();
        }
        writer.start_file(name, options).expect("write source part");
        writer.write_all(&bytes).expect("write source bytes");
    }
    assert!(rewrite.is_none(), "part to rewrite was present");
    writer.finish().expect("finish source package").into_inner()
}

fn xml_attr(tag: &str, name: &str) -> Option<String> {
    let marker = format!(r#"{name}="#);
    let start = tag.find(&marker)? + marker.len();
    let value = &tag[start..];
    let value = value.strip_prefix('"')?;
    Some(value.split('"').next()?.to_string())
}

fn replace_xml_id(tag: &str, id: u32) -> String {
    Regex::new(r#"\sid="[^"]*""#)
        .expect("valid id attribute regex")
        .replace(tag, format!(r#" id="{id}""#))
        .into_owned()
}
