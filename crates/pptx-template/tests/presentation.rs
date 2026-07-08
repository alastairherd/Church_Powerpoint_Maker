use pptx_template::{Presentation, Run};

const TEMPLATE: &[u8] = include_bytes!("../../../template.pptx");

#[test]
fn round_trip_preserves_slide_relationship_integrity() {
    let pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    let bytes = pres.save_bytes().expect("save template");
    let reopened = Presentation::open_bytes(&bytes).expect("reopen saved template");

    assert_eq!(reopened.slide_count(), 1);
    reopened.validate().expect("structural validation");
}

#[test]
fn can_add_layout_slides_set_text_and_reorder() {
    let mut pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    let first = pres.add_slide_from_layout(0).expect("add first slide");
    let second = pres.add_slide_from_layout(0).expect("add second slide");

    pres.slide_mut(first)
        .expect("first slide")
        .placeholder(0)
        .expect("first title")
        .set_text("First cloned title")
        .expect("set first title");
    pres.slide_mut(second)
        .expect("second slide")
        .placeholder(0)
        .expect("second title")
        .set_text("Second cloned title")
        .expect("set second title");

    pres.reorder(&[0, second, first]).expect("reorder");
    let bytes = pres.save_bytes().expect("save");
    let reopened = Presentation::open_bytes(&bytes).expect("reopen");

    assert_eq!(reopened.slide_count(), 3);
    assert!(reopened
        .slide_text(1)
        .expect("text")
        .starts_with("Second cloned title"));
    assert!(reopened
        .slide_text(2)
        .expect("text")
        .starts_with("First cloned title"));
    reopened.validate().expect("structural validation");
}

#[test]
fn rich_text_writes_superscript_runs() {
    let mut pres = Presentation::open_bytes(TEMPLATE).expect("open template");
    let idx = pres.add_slide_from_layout(0).expect("add slide");

    pres.slide_mut(idx)
        .expect("slide")
        .placeholder(1)
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
}
