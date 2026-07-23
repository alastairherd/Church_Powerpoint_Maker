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
    let mut liturgy_body_is_black_arial_black = false;
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
            if xml.contains("typeface=\"Arial Black\"")
                && xml.contains("<a:srgbClr val=\"000000\"/>")
            {
                liturgy_body_is_black_arial_black = true;
            }
        }
    }
    assert!(
        liturgy_body_is_black_arial_black,
        "liturgy body runs explicitly use black Arial Black"
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

#[tokio::test]
async fn generated_content_keeps_template_hierarchy_and_safe_sizing() {
    let mut service = ServiceRecord::new(
        "service-content",
        "Content regression",
        NaiveDate::from_ymd_opt(2026, 7, 19).unwrap(),
        ServicePreset::Am,
        "Alastair",
    );
    service.components = vec![
        ServiceComponent::Notices {
            id: "notices".into(),
            heading: "Notices".into(),
            rows: vec![deck_builder::NoticeRow {
                when: "Today".into(),
                title: "Notice title".into(),
                details: "Details remain readable".into(),
                emphasis: true,
            }],
        },
        ServiceComponent::CallToWorship {
            id: "call".into(),
            heading: "Call to Worship".into(),
            reference: "Psalm 96:2".into(),
            text: "[1] Sing to the LORD; [2] bless his name.".into(),
            external_source_failed: false,
        },
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
            reference: String::new(),
            tune: None,
            slide_breaks: vec!["one\ntwo\nthree\nfour\nfive\nsix\nseven".into()],
        },
        ServiceComponent::Teaching {
            id: "teaching".into(),
            heading: "Teaching".into(),
            source: deck_builder::TeachingSource::WestminsterShorterCatechism,
            selection: "Q. 1".into(),
            text: String::new(),
        },
    ];

    let bytes = build_deck(&service, &MockSources, "522221")
        .await
        .expect("content deck builds");
    let pres = Presentation::open_bytes(&bytes).expect("generated deck opens");
    let xml: Vec<_> = (0..pres.slide_count())
        .map(|index| pres.slide_xml(index).unwrap())
        .collect();
    let notices = xml
        .iter()
        .find(|slide| slide.contains("Notice title"))
        .unwrap();
    assert!(notices.contains("typeface=\"Arial Black\""));
    assert!(!notices.contains("typeface=\"Arial\""));
    assert!(!notices.contains("sz=\"2200\""));
    assert!(notices.contains("sz=\"2800\""));
    assert!(notices.contains("b=\"1\""));
    assert!(notices.contains("<a:t>Today</a:t>") || notices.contains("<a:t>Today </a:t>"));
    assert!(notices.contains("<a:t>–</a:t>"));
    assert!(!notices.contains(" · "));
    assert!(notices.contains("<a:t>Details remain readable</a:t>"));

    let call = xml
        .iter()
        .find(|slide| slide.contains("Sing to the LORD"))
        .unwrap();
    assert!(!call.contains("[1]") && !call.contains("[2]"));
    assert!(!call.contains("baseline=\"30000\""));
    assert!(call.contains("Psalm 96:2"));

    let liturgy = xml.iter().find(|slide| slide.contains("All.")).unwrap();
    assert!(liturgy.contains("typeface=\"Liberation Serif\""));
    assert!(liturgy.contains("schemeClr val=\"accent1\""));
    assert!(liturgy.contains("typeface=\"Arial Black\""));
    assert!(!liturgy.contains("typeface=\"Arial\""));
    assert!(xml.iter().any(|slide| slide.contains("Amen.")));
    let amen = xml.iter().find(|slide| slide.contains("Amen.")).unwrap();
    assert!(amen.contains("<a:t>Amen.</a:t>"));
    assert!(amen.contains("<a:schemeClr val=\"accent1\"/>") && amen.contains("b=\"1\""));

    let psalm = xml.iter().find(|slide| slide.contains("seven")).unwrap();
    assert!(psalm.contains("sz=\"3200\""));
    assert!(!psalm.contains("sz=\"2600\""));
    assert!(psalm.contains("typeface=\"Arial Black\""));
    assert!(psalm.contains("<a:off x=\"6724800\" y=\"6080400\"/>"));
    let teaching = xml
        .iter()
        .find(|slide| slide.contains("What is the chief end of man?"))
        .unwrap();
    assert!(teaching.contains("sz=\"3000\" b=\"1\"") || teaching.contains("b=\"1\" sz=\"3000\""));
    assert!(teaching.contains("<a:t>What is the chief end of man?</a:t>"));
    assert!(teaching
        .contains("<a:t>Man&apos;s chief end is to glorify God, and to enjoy him forever.</a:t>"));
    assert!(teaching.contains("<a:schemeClr val=\"accent1\"/>"));
    assert!(teaching.contains("<a:srgbClr val=\"000000\"/>"));
    assert!(
        teaching.matches("<a:p>").count() >= 4,
        "teaching keeps question/answer paragraphs"
    );
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
    for selection in ["1", "Q1", "Q. 1", "Question 1"] {
        assert_eq!(
            deck_builder::parse_catechism_selection(selection).unwrap(),
            1
        );
    }
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
