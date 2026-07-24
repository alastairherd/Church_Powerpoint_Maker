use pptx_template::{Presentation, Run};
use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

const TEMPLATE: &[u8] = include_bytes!("../../deck-builder/assets/template.pptx");

#[test]
fn round_trip_preserves_slide_relationship_integrity() {
    let pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    let bytes = pres.save_bytes().expect("save template");
    let reopened = Presentation::open_bytes(&bytes).expect("reopen saved template");

    assert_eq!(reopened.slide_count(), 49);
    reopened.validate().expect("structural validation");
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
    let source_bytes = source_with_distinct_master();
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
    assert_eq!(presentation.matches("r:id=\"rId61\"").count(), 1);
    assert!(presentation_rels.contains("Target=\"slideMasters/slideMaster2.xml\""));

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

fn source_without_distinct_master_registration() -> Vec<u8> {
    let source = source_with_distinct_master();
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
            bytes = xml
                .replace("<p:sldMasterId id=\"2147483661\" r:id=\"rId61\"/>", "")
                .into_bytes();
        }
        writer.start_file(name, options).expect("write source part");
        writer.write_all(&bytes).expect("write source bytes");
    }
    writer.finish().expect("finish source package").into_inner()
}

fn source_with_distinct_master() -> Vec<u8> {
    let mut input = ZipArchive::new(Cursor::new(TEMPLATE)).expect("open template package");
    let mut files = BTreeMap::new();
    for index in 0..input.len() {
        let mut file = input.by_index(index).expect("read template part");
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)
            .expect("read template part bytes");
        files.insert(file.name().to_string(), bytes);
    }

    let master = String::from_utf8(
        files
            .get("ppt/slideMasters/slideMaster1.xml")
            .expect("template master")
            .clone(),
    )
    .expect("master XML is UTF-8")
    .replacen("preserve=\"1\"", "preserve=\"0\"", 1);
    files.insert(
        "ppt/slideMasters/slideMaster2.xml".into(),
        master.into_bytes(),
    );
    let master_rels = files
        .get("ppt/slideMasters/_rels/slideMaster1.xml.rels")
        .expect("template master relationships")
        .clone();
    files.insert(
        "ppt/slideMasters/_rels/slideMaster2.xml.rels".into(),
        master_rels,
    );

    let layout_rels = String::from_utf8(
        files
            .get("ppt/slideLayouts/_rels/slideLayout12.xml.rels")
            .expect("template layout relationships")
            .clone(),
    )
    .expect("layout relationships are UTF-8")
    .replace(
        "../slideMasters/slideMaster1.xml",
        "../slideMasters/slideMaster2.xml",
    );
    files.insert(
        "ppt/slideLayouts/_rels/slideLayout12.xml.rels".into(),
        layout_rels.into_bytes(),
    );

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
    let first_slide_end = presentation[slide_list_start..slide_list_end]
        .find("/>")
        .expect("first slide entry")
        + slide_list_start
        + 2;
    let first_slide = presentation[slide_list_start..first_slide_end].to_string();
    presentation.replace_range(slide_list_start..slide_list_end, &first_slide);
    presentation = presentation.replacen(
        "</p:sldMasterIdLst>",
        "<p:sldMasterId id=\"2147483661\" r:id=\"rId61\"/></p:sldMasterIdLst>",
        1,
    );
    files.insert("ppt/presentation.xml".into(), presentation.into_bytes());

    let mut presentation_rels = String::from_utf8(
        files
            .get("ppt/_rels/presentation.xml.rels")
            .expect("template presentation relationships")
            .clone(),
    )
    .expect("presentation relationships are UTF-8");
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
