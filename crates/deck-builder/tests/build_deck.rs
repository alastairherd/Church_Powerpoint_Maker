use async_trait::async_trait;
use chrono::NaiveDate;
use deck_builder::{build_deck, Component, FixedComponent, Hymn, Scripture, ServiceOrder, Sources};
use pptx_template::Presentation;
use std::io::{Cursor, Read};
use zip::ZipArchive;

struct MockSources;

#[async_trait]
impl Sources for MockSources {
    async fn scripture(&self, reference: &str) -> anyhow::Result<Scripture> {
        Ok(Scripture {
            reference: reference.to_string(),
            text: "[1] In the beginning God created the heavens and the earth.".to_string(),
        })
    }

    async fn hymn(&self, _url: &str) -> anyhow::Result<Hymn> {
        Ok(Hymn {
            title: "Amazing Grace".to_string(),
            stanzas: vec!["Amazing grace! how sweet the sound".to_string()],
            author: "John Newton".to_string(),
            composer: "Unknown".to_string(),
            tune: "NEW BRITAIN".to_string(),
            copyright: "Public Domain".to_string(),
        })
    }
}

#[tokio::test]
async fn builds_valid_pptx_from_sample_service_order() {
    let order = ServiceOrder {
        date: NaiveDate::from_ymd_opt(2026, 7, 12).unwrap(),
        components: vec![
            Component::Fixed {
                key: "confession".to_string(),
                title: Some("Confession".to_string()),
            },
            Component::Psalm {
                reference: "Psalm 1:1-3 (a)".to_string(),
            },
            Component::Scripture {
                reference: "Genesis 1:1".to_string(),
                title: Some("First Reading".to_string()),
            },
            Component::Catechism { question: 1 },
            Component::Hymn {
                url: "https://hymnary.org/text/amazing_grace_how_sweet_the_sound".to_string(),
            },
        ],
    };

    let bytes = build_deck(&order, &MockSources).await.expect("deck builds");
    let pres = Presentation::open_bytes(&bytes).expect("opens generated deck");

    assert!(bytes.starts_with(b"PK"));
    assert!(pres.slide_count() >= 5);
    pres.validate()
        .expect("generated deck is structurally valid");
}

#[tokio::test]
async fn generated_deck_contains_parseable_slide_xml() {
    let order = ServiceOrder {
        date: NaiveDate::from_ymd_opt(2026, 7, 12).unwrap(),
        components: vec![Component::Fixed {
            key: "confession".to_string(),
            title: Some("Confession".to_string()),
        }],
    };

    let bytes = build_deck(&order, &MockSources).await.expect("deck builds");
    let mut zip = ZipArchive::new(Cursor::new(bytes)).expect("pptx zip opens");

    for i in 0..zip.len() {
        let mut file = zip.by_index(i).expect("zip entry");
        let name = file.name().to_string();
        if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
            let mut xml = String::new();
            file.read_to_string(&mut xml).expect("slide xml is utf-8");
            assert!(
                xml_is_parseable(&xml),
                "generated slide XML should parse: {name}\n{xml}"
            );
            assert!(
                !xml.contains("charset=\"0\" b=\"1\""),
                "bold should be applied to a:rPr, not nested font tags: {name}\n{xml}"
            );
        }
    }
}

#[test]
fn embedded_sources_resolve_catechism_psalm_and_fixed_component() {
    let fixed = FixedComponent::find("confession").expect("confession exists");
    let catechism = deck_builder::Catechism::find(1).expect("wsc q1 exists");
    let psalm = deck_builder::Psalm::find("Psalm 1:1-3 (a)").expect("psalm exists");

    assert_eq!(fixed.speaker, "All.");
    assert_eq!(catechism.question, "What is the chief end of man?");
    assert_eq!(psalm.title, "Psalm 1:1-3 (a)");
    assert_eq!(psalm.stanzas.len(), 3);
}

fn xml_is_parseable(xml: &str) -> bool {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => return true,
            Ok(_) => buf.clear(),
            Err(_) => return false,
        }
    }
}
