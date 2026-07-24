use async_trait::async_trait;
use chrono::NaiveDate;
use deck_builder::{
    build_deck, FixedComponent, Scripture, ServiceComponent, ServicePreset, ServiceRecord, Sources,
    StoredSong, VersionPin,
};
use pptx_template::Presentation;
use std::io::{Cursor, Read};
use std::path::PathBuf;
use std::sync::{Arc, Mutex};
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

struct FailingFixedSources;

#[async_trait]
impl Sources for FailingFixedSources {
    async fn scripture(&self, reference: &str) -> anyhow::Result<Scripture> {
        Ok(Scripture {
            reference: reference.to_string(),
            text: "[1] In the beginning God created the heavens and the earth.".to_string(),
        })
    }

    fn fixed_component(&self, key: &str) -> anyhow::Result<FixedComponent> {
        Err(anyhow::anyhow!(
            "fixed component lookup should not run for {key}"
        ))
    }
}

struct TrackingFixedSources {
    calls: Arc<Mutex<Vec<String>>>,
}

#[async_trait]
impl Sources for TrackingFixedSources {
    async fn scripture(&self, reference: &str) -> anyhow::Result<Scripture> {
        Ok(Scripture {
            reference: reference.to_string(),
            text: "[1] In the beginning God created the heavens and the earth.".to_string(),
        })
    }

    fn fixed_component(&self, key: &str) -> anyhow::Result<FixedComponent> {
        self.calls.lock().unwrap().push(key.to_string());
        FixedComponent::find(key)
    }
}

fn canonical_liturgy_components() -> Vec<ServiceComponent> {
    [
        ("Prayer for Purity", "prayer_for_purity"),
        ("The Ten Commandments", "ten_commandments"),
        ("Lord's Prayer", "lords_prayer"),
        ("Confession", "confession"),
        ("Assurance of Forgiveness", "assurance"),
        ("Comfortable Words", "comfortable_words"),
        ("Prayer of Humble Access", "humble_access"),
        ("Prayer of Consecration", "consecration"),
        ("Final Blessing", "final_blessing"),
    ]
    .into_iter()
    .enumerate()
    .map(|(index, (heading, key))| ServiceComponent::LiturgyBlock {
        id: format!("liturgy-{}", index + 1),
        heading: heading.into(),
        key: key.into(),
        version: None,
        text: String::new(),
    })
    .collect()
}

fn named_shape_xml<'a>(xml: &'a str, name: &str) -> &'a str {
    let marker = format!("name=\"{name}\"");
    let name_start = xml.find(&marker).expect("named shape exists");
    let start = xml[..name_start].rfind("<p:sp>").expect("shape starts");
    let end = start + xml[start..].find("</p:sp>").expect("shape ends") + "</p:sp>".len();
    &xml[start..end]
}

fn write_openxml_validation_output(bytes: &[u8]) {
    let Some(path) = std::env::var_os("OPENXML_VALIDATOR_OUTPUT") else {
        return;
    };
    let path = PathBuf::from(path);
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent).expect("Open XML validation output directory creates");
    }
    std::fs::write(&path, bytes).expect("generated deck writes for Open XML validation");
}

const CANONICAL_LITURGY_SLIDES: &[usize] = &[
    9, 10, 11, 12, 13, 14, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 47,
];

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
            show_verse_numbers: true,
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
    write_openxml_validation_output(&bytes);
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
async fn generated_psalm_runs_specify_arial_black_for_all_script_ranges() {
    let mut service = ServiceRecord::new(
        "psalm-fonts",
        "Psalm font regression",
        NaiveDate::from_ymd_opt(2026, 7, 12).unwrap(),
        ServicePreset::Am,
        "Alastair",
    );
    service.components = vec![ServiceComponent::Psalm {
        id: "psalm".into(),
        heading: "Psalm".into(),
        reference: "Psalm 1:1-3 (a)".into(),
        show_verse_numbers: true,
        tune: None,
        slide_breaks: Vec::new(),
    }];

    let bytes = build_deck(&service, &MockSources, "522221")
        .await
        .expect("deck builds");
    let pres = Presentation::open_bytes(&bytes).expect("generated deck opens");
    let first_psalm_slide = pres.slide_xml(0).expect("first Psalm slide XML");
    let first_psalm_body = named_shape_xml(&first_psalm_slide, "TextShape 2");

    for script in ["latin", "ea", "cs"] {
        assert!(
            first_psalm_body.contains(&format!("<a:{script} typeface=\"Arial Black\"/>")),
            "Psalm body should specify Arial Black for {script} script"
        );
    }
    assert!(!first_psalm_body.contains("Calibri"));
    assert!(!first_psalm_body.contains("+mn-"));

    let rpr_start = first_psalm_body
        .find("<a:rPr")
        .expect("Psalm body has run properties");
    let rpr_end = first_psalm_body[rpr_start..]
        .find("</a:rPr>")
        .map(|offset| rpr_start + offset + "</a:rPr>".len())
        .expect("Psalm run properties are closed");
    let rpr = &first_psalm_body[rpr_start..rpr_end];
    let positions = [
        rpr.find("<a:ln").expect("run line properties"),
        rpr.find("<a:solidFill").expect("run fill properties"),
        rpr.find("<a:effectLst").expect("run effects"),
        rpr.find("<a:uLnTx").expect("run underline line properties"),
        rpr.find("<a:uFillTx")
            .expect("run underline fill properties"),
        rpr.find("<a:latin").expect("Latin script font"),
        rpr.find("<a:ea").expect("East Asian script font"),
        rpr.find("<a:cs").expect("complex script font"),
    ];
    assert!(
        positions.windows(2).all(|pair| pair[0] < pair[1]),
        "Psalm run properties follow the DrawingML schema order: {rpr}"
    );
}

#[tokio::test]
async fn generated_psalm_can_hide_or_retain_leading_verse_numbers() {
    let mut service = ServiceRecord::new(
        "psalm-verse-numbers",
        "Psalm verse number setting",
        NaiveDate::from_ymd_opt(2026, 7, 12).unwrap(),
        ServicePreset::Am,
        "Alastair",
    );
    service.components = vec![ServiceComponent::Psalm {
        id: "psalm".into(),
        heading: "Psalm".into(),
        reference: "Psalm 23:1-1".into(),
        show_verse_numbers: false,
        tune: None,
        slide_breaks: vec!["⁴<underline>Though I</underline> walk through the valley".into()],
    }];

    let hidden = build_deck(&service, &MockSources, "522221")
        .await
        .expect("deck without verse numbers builds");
    let hidden_xml = Presentation::open_bytes(&hidden)
        .expect("hidden-number deck opens")
        .slide_xml(0)
        .expect("hidden-number Psalm slide XML");
    assert!(hidden_xml.contains("Though I"));
    assert!(!hidden_xml.contains("<a:t>4</a:t>"));

    if let ServiceComponent::Psalm {
        show_verse_numbers, ..
    } = &mut service.components[0]
    {
        *show_verse_numbers = true;
    }
    let shown = build_deck(&service, &MockSources, "522221")
        .await
        .expect("deck with verse numbers builds");
    let shown_xml = Presentation::open_bytes(&shown)
        .expect("shown-number deck opens")
        .slide_xml(0)
        .expect("shown-number Psalm slide XML");
    assert!(shown_xml.contains("<a:t>4</a:t>"));
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
            text: "[1] Sing to the LORD;\n    [2] bless his name.\n2 Tell of his salvation.".into(),
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
            show_verse_numbers: true,
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
    assert!(notices.contains("sz=\"2400\" i=\"1\""));
    assert!(notices.contains("b=\"1\""));
    assert!(notices.contains("<a:off x=\"765544\" y=\"1100000\"/>"));
    assert!(notices.contains("<a:ext cx=\"8893463\" cy=\"4179036\"/>"));
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
    assert!(call.contains("<a:t>bless his name.</a:t>"));
    assert!(call.contains("<a:t>Tell of his salvation.</a:t>"));
    let reference_run = call
        .split("<a:r>")
        .find(|run| run.contains("<a:t>Psalm 96:2</a:t>"))
        .expect("reference run present");
    assert!(reference_run.contains("i=\"1\""));

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
    assert!(psalm.contains("sz=\"2800\""));
    assert!(!psalm.contains("sz=\"3200\"") && !psalm.contains("sz=\"2600\""));
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

#[tokio::test]
async fn blank_teaching_renders_a_placeholder_but_invalid_automatic_loading_still_fails() {
    let mut blank = ServiceRecord::new(
        "blank-teaching",
        "Blank teaching",
        NaiveDate::from_ymd_opt(2026, 7, 19).unwrap(),
        ServicePreset::Am,
        "Alastair",
    );
    blank.components = vec![ServiceComponent::Teaching {
        id: "teaching".into(),
        heading: "Teaching".into(),
        source: deck_builder::TeachingSource::WestminsterShorterCatechism,
        selection: String::new(),
        text: String::new(),
    }];

    let bytes = build_deck(&blank, &MockSources, "522221")
        .await
        .expect("blank teaching still builds");
    let pres = Presentation::open_bytes(&bytes).expect("generated deck opens");
    assert!(pres
        .slide_text(0)
        .unwrap()
        .contains("Choose a teaching question or enter teaching text"));

    let mut invalid = blank.clone();
    if let ServiceComponent::Teaching { selection, .. } = &mut invalid.components[0] {
        *selection = "not a question".into();
    }
    let error = build_deck(&invalid, &MockSources, "522221")
        .await
        .expect_err("invalid catechism selections still fail");
    assert!(error
        .to_string()
        .contains("enter a catechism question such as 1, Q1, or Q. 1"));

    let mut heidelberg = blank.clone();
    if let ServiceComponent::Teaching {
        source, selection, ..
    } = &mut heidelberg.components[0]
    {
        *source = deck_builder::TeachingSource::Heidelberg1891;
        *selection = "Q1".into();
    }
    let bytes = build_deck(&heidelberg, &MockSources, "522221")
        .await
        .expect("heidelberg questions load automatically");
    let pres = Presentation::open_bytes(&bytes).expect("generated deck opens");
    assert!(pres
        .slide_text(0)
        .unwrap()
        .contains("What is your only comfort in life and death?"));

    let mut confession = blank;
    if let ServiceComponent::Teaching {
        source, selection, ..
    } = &mut confession.components[0]
    {
        *source = deck_builder::TeachingSource::WestminsterConfessionOriginalBritish;
        *selection = "21.8".into();
    }
    let bytes = build_deck(&confession, &MockSources, "522221")
        .await
        .expect("confession sections load automatically");
    let pres = Presentation::open_bytes(&bytes).expect("generated deck opens");
    let text = pres.slide_text(0).unwrap();
    assert!(text.contains("Chapter 21"));
    assert!(text.contains("This Sabbath is then kept holy unto the Lord"));
}

#[tokio::test]
async fn blank_psalm_slide_breaks_are_skipped() {
    let mut service = ServiceRecord::new(
        "service-blank-psalm",
        "Blank psalm break",
        NaiveDate::from_ymd_opt(2026, 7, 26).unwrap(),
        ServicePreset::Am,
        "Alastair",
    );
    service.components = vec![ServiceComponent::Psalm {
        id: "psalm".into(),
        heading: "Psalm".into(),
        reference: "Psalm 23:1–6".into(),
        show_verse_numbers: true,
        tune: None,
        slide_breaks: vec!["The LORD's my shepherd".into(), String::new(), "  ".into()],
    }];

    let bytes = build_deck(&service, &MockSources, "522221")
        .await
        .expect("psalm with blank breaks builds");
    let pres = Presentation::open_bytes(&bytes).expect("generated deck opens");
    assert_eq!(pres.slide_count(), 1);
    let text = pres.slide_text(0).unwrap();
    assert!(text.contains("The LORD's my shepherd"));
    assert!(text.contains("Sing Psalms"));
}

#[tokio::test]
async fn crowded_slides_hide_the_master_logo() {
    let mut service = ServiceRecord::new(
        "service-crowded",
        "Crowded slide",
        NaiveDate::from_ymd_opt(2026, 7, 26).unwrap(),
        ServicePreset::Am,
        "Alastair",
    );
    let long_verse = (1..=14)
        .map(|line| format!("Line {line} of a very long lyric verse"))
        .collect::<Vec<_>>()
        .join("\n");
    service.components = vec![ServiceComponent::Song {
        id: "song".into(),
        title: "Crowded song".into(),
        song: None,
        lyric_slides: vec![long_verse, "A short verse\nof two lines".into()],
        credits: String::new(),
    }];

    let bytes = build_deck(&service, &MockSources, "522221")
        .await
        .expect("crowded song builds");
    let pres = Presentation::open_bytes(&bytes).expect("generated deck opens");
    assert!(pres.slide_xml(0).unwrap().contains("showMasterSp=\"0\""));
    assert!(!pres.slide_xml(1).unwrap().contains("showMasterSp=\"0\""));
}

#[test]
fn embedded_sources_resolve_catechism_psalm_and_fixed_component() {
    let fixed = FixedComponent::find("confession").expect("confession exists");
    let catechism = deck_builder::Teaching::find(
        deck_builder::TeachingSource::WestminsterShorterCatechism,
        "Q. 1",
    )
    .expect("wsc q1 exists");
    let heidelberg =
        deck_builder::Teaching::find(deck_builder::TeachingSource::Heidelberg1891, "1")
            .expect("heidelberg q1 exists");
    let confession = deck_builder::Teaching::find(
        deck_builder::TeachingSource::WestminsterConfessionOriginalBritish,
        "1.2",
    )
    .expect("wcf 1.2 exists");
    let whole_chapter = deck_builder::Teaching::find(
        deck_builder::TeachingSource::WestminsterConfessionOriginalBritish,
        "Chapter 1",
    )
    .expect("wcf chapter 1 exists");
    let psalm = deck_builder::Psalm::find("Psalm 1:1-3 (a)").expect("psalm exists");
    let psalm_with_typographic_dash =
        deck_builder::Psalm::find("psalm 23:1–6").expect("friendly psalm reference works");
    assert_eq!(fixed.speaker, "All.");
    assert_eq!(catechism.question, "What is the chief end of man?");
    assert_eq!(catechism.selection, "1");
    assert_eq!(
        heidelberg.question,
        "What is your only comfort in life and death?"
    );
    assert_eq!(confession.selection, "1.2");
    assert_eq!(confession.question, "Chapter 1: Of the Holy Scripture");
    assert!(confession.answer.contains("Word of God written"));
    assert_eq!(whole_chapter.selection, "1");
    assert!(whole_chapter.answer.starts_with("1. "));
    assert!(whole_chapter.answer.contains("\n\n10. "));
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

#[tokio::test]
async fn lords_supper_presets_clone_canonical_liturgy_slides_without_sources() {
    let source =
        Presentation::open_bytes(include_bytes!("../assets/template.pptx")).expect("source opens");

    for preset in [ServicePreset::AmLordsSupper, ServicePreset::PmLordsSupper] {
        let mut service = ServiceRecord::new(
            format!("service-{preset:?}"),
            "Lord's Supper service",
            NaiveDate::from_ymd_opt(2026, 7, 19).unwrap(),
            preset,
            "Alastair",
        );
        service.lords_supper = false;
        service.components = canonical_liturgy_components();

        let bytes = build_deck(&service, &FailingFixedSources, "522221")
            .await
            .expect("canonical Lord's Supper deck builds without fixed components");
        let generated = Presentation::open_bytes(&bytes).expect("generated deck opens");

        assert_eq!(generated.slide_count(), CANONICAL_LITURGY_SLIDES.len());
        for (generated_index, &source_index) in CANONICAL_LITURGY_SLIDES.iter().enumerate() {
            assert_eq!(
                generated.slide_text(generated_index).unwrap(),
                source.slide_text(source_index).unwrap(),
                "canonical slide text at generated index {generated_index}"
            );
            assert_eq!(
                generated.slide_xml(generated_index).unwrap(),
                source.slide_xml(source_index).unwrap(),
                "canonical slide XML at generated index {generated_index}"
            );
        }

        let ten_commandments_minister = generated.slide_xml(1).unwrap();
        assert!(ten_commandments_minister.contains("typeface=\"Arial Black\""));
        assert!(ten_commandments_minister.contains("<a:schemeClr val=\"accent1\"/>"));
        let ten_commandments_all = generated.slide_xml(2).unwrap();
        assert!(ten_commandments_all.contains("<a:srgbClr val=\"000000\"/>"));
        assert!(generated.slide_text(2).unwrap().contains("All:"));
        let communion_all = generated.slide_xml(6).unwrap();
        assert!(generated.slide_text(6).unwrap().contains("All."));
        assert!(communion_all.contains("typeface=\"Arial Black\""));
        let communion_minister = generated.slide_xml(9).unwrap();
        assert!(communion_minister.contains("<a:t>Minister.</a:t>"));
        assert!(communion_minister.contains("<a:schemeClr val=\"accent1\"/>"));
    }
}

#[tokio::test]
async fn ordinary_blank_liturgy_uses_fixed_component_renderer() {
    let calls = Arc::new(Mutex::new(Vec::new()));
    let source = TrackingFixedSources {
        calls: Arc::clone(&calls),
    };
    let mut service = ServiceRecord::new(
        "ordinary-service",
        "Morning service",
        NaiveDate::from_ymd_opt(2026, 7, 19).unwrap(),
        ServicePreset::Am,
        "Alastair",
    );
    service.lords_supper = true;
    service.components = vec![ServiceComponent::LiturgyBlock {
        id: "confession".into(),
        heading: "Confession".into(),
        key: "confession".into(),
        version: None,
        text: String::new(),
    }];

    let bytes = build_deck(&service, &source, "522221")
        .await
        .expect("ordinary deck builds");
    let generated = Presentation::open_bytes(&bytes).expect("generated deck opens");

    assert_eq!(generated.slide_count(), 2);
    assert_eq!(
        calls.lock().unwrap().clone(),
        vec!["confession".to_string()]
    );
}

#[tokio::test]
async fn lords_supper_manual_liturgy_uses_generic_renderer() {
    let mut service = ServiceRecord::new(
        "manual-communion-service",
        "Lord's Supper service",
        NaiveDate::from_ymd_opt(2026, 7, 19).unwrap(),
        ServicePreset::PmLordsSupper,
        "Alastair",
    );
    service.components = vec![ServiceComponent::LiturgyBlock {
        id: "confession".into(),
        heading: "Confession".into(),
        key: "confession".into(),
        version: None,
        text: "Minister. Manual confession text.".into(),
    }];

    let bytes = build_deck(&service, &FailingFixedSources, "522221")
        .await
        .expect("manual communion text uses generic rendering");
    let generated = Presentation::open_bytes(&bytes).expect("generated deck opens");

    assert_eq!(generated.slide_count(), 1);
    assert!(generated
        .slide_text(0)
        .unwrap()
        .contains("Manual confession text."));
}
