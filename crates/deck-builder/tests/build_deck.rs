use async_trait::async_trait;
use chrono::NaiveDate;
use deck_builder::{
    build_deck, FixedComponent, Scripture, ServiceComponent, ServicePreset, ServiceRecord, Sources,
    StoredSong, VersionPin,
};
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

    async fn song(&self, _id: &str, _version: u64) -> anyhow::Result<StoredSong> {
        Ok(StoredSong {
            title: "Source presentation".into(),
            slides: Vec::new(),
            credits: String::new(),
            source_pptx: Some(include_bytes!("../assets/template.pptx").to_vec()),
        })
    }
}

#[tokio::test]
async fn builds_valid_pptx_from_service_record() {
    let mut service = ServiceRecord::new(
        "service-one",
        "Morning service",
        NaiveDate::from_ymd_opt(2026, 7, 12).unwrap(),
        ServicePreset::Am,
        "Alastair",
    );
    service.components = vec![
        ServiceComponent::LiturgyBlock {
            id: "confession".into(),
            heading: "Confession".into(),
            key: "confession".into(),
            version: Some(1),
            text: String::new(),
        },
        ServiceComponent::Psalm {
            id: "psalm".into(),
            heading: "Psalm".into(),
            reference: "Psalm 1:1-3 (a)".into(),
            tune: None,
            slide_breaks: Vec::new(),
        },
        ServiceComponent::Reading {
            id: "reading".into(),
            heading: "First Reading".into(),
            reference: "Genesis 1:1".into(),
            bible_page: Some(1),
        },
        ServiceComponent::Song {
            id: "song".into(),
            title: "Amazing Grace".into(),
            song: None,
            lyric_slides: vec!["Amazing grace! how sweet the sound".into()],
            credits: "Words: John Newton · Public Domain".into(),
        },
    ];

    let bytes = build_deck(&service, &MockSources, "522221")
        .await
        .expect("deck builds");
    let pres = Presentation::open_bytes(&bytes).expect("opens generated deck");
    assert!(bytes.starts_with(b"PK"));
    assert!(pres.slide_count() >= 4);
    pres.validate()
        .expect("generated deck is structurally valid");
    let mut zip = ZipArchive::new(Cursor::new(bytes)).expect("generated zip opens");
    let mut liturgy_body_is_black_arial = false;
    for index in 0..zip.len() {
        let mut part = zip.by_index(index).expect("zip part");
        let name = part.name().to_string();
        assert!(!name.contains("notesSlide"), "notes are removed: {name}");
        assert!(!name.contains("notesMaster"), "notes are removed: {name}");
        assert!(
            !name.contains("revisionInfo"),
            "revision metadata is removed: {name}"
        );
        if name == "ppt/_rels/presentation.xml.rels" || name == "[Content_Types].xml" {
            let mut xml = String::new();
            part.read_to_string(&mut xml).expect("package XML is UTF-8");
            assert!(
                !xml.contains("revisionInfo"),
                "revision metadata is removed: {name}"
            );
        }
        if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
            let mut xml = String::new();
            part.read_to_string(&mut xml).expect("slide xml is utf-8");
            if xml.contains("typeface=\"Arial\"") && xml.contains("<a:srgbClr val=\"000000\"/>") {
                liturgy_body_is_black_arial = true;
            }
        }
    }
    assert!(
        liturgy_body_is_black_arial,
        "liturgy body runs explicitly use black Arial"
    );
}

#[tokio::test]
async fn generated_deck_contains_parseable_slide_xml() {
    let mut service = ServiceRecord::new(
        "service-two",
        "Evening service",
        NaiveDate::from_ymd_opt(2026, 7, 12).unwrap(),
        ServicePreset::TraditionalPm,
        "Alastair",
    );
    service.components.truncate(6);
    let bytes = build_deck(&service, &MockSources, "522221")
        .await
        .expect("deck builds");
    let mut zip = ZipArchive::new(Cursor::new(bytes)).expect("pptx zip opens");

    for i in 0..zip.len() {
        let mut file = zip.by_index(i).expect("zip entry");
        let name = file.name().to_string();
        if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
            let mut xml = String::new();
            file.read_to_string(&mut xml).expect("slide xml is utf-8");
            assert!(
                xml_is_parseable(&xml),
                "generated slide XML should parse: {name}"
            );
        }
    }
}

#[test]
fn embedded_sources_resolve_catechism_psalm_and_fixed_component() {
    let fixed = FixedComponent::find("confession").expect("confession exists");
    let catechism = deck_builder::Catechism::find(1).expect("wsc q1 exists");
    let psalm = deck_builder::Psalm::find("Psalm 1:1-3 (a)").expect("psalm exists");
    let psalm_with_typographic_dash =
        deck_builder::Psalm::find("psalm 23:1–6").expect("friendly psalm reference works");
    assert_eq!(fixed.speaker, "All.");
    assert_eq!(catechism.question, "What is the chief end of man?");
    assert_eq!(psalm.stanzas.len(), 3);
    assert_eq!(psalm_with_typographic_dash.meter, "11 11 11");
    assert_eq!(psalm_with_typographic_dash.stanzas.len(), 5);
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

#[tokio::test]
async fn imports_original_song_slides_instead_of_rebuilding_their_text() {
    let mut service = ServiceRecord::new(
        "service-song-import",
        "Song import",
        NaiveDate::from_ymd_opt(2026, 7, 19).unwrap(),
        ServicePreset::Am,
        "Alastair",
    );
    service.components = vec![ServiceComponent::Song {
        id: "song".into(),
        title: "Source presentation".into(),
        song: Some(VersionPin {
            entity_id: "source-song".into(),
            version: 1,
            slide_count: 28,
        }),
        lyric_slides: Vec::new(),
        credits: String::new(),
    }];

    let source =
        Presentation::open_bytes(include_bytes!("../assets/template.pptx")).expect("source opens");
    let bytes = build_deck(&service, &MockSources, "522221")
        .await
        .expect("deck builds");
    let generated = Presentation::open_bytes(&bytes).expect("generated deck opens");

    assert_eq!(generated.slide_count(), source.slide_count());
    for index in 0..source.slide_count() {
        assert_eq!(
            generated.slide_text(index).unwrap(),
            source.slide_text(index).unwrap(),
            "source slide {index} is preserved"
        );
    }
}
