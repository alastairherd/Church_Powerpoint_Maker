use pptx_template::{Presentation, Run};

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
