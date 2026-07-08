use deck_builder::textproc::{british_spellings, scripture_runs, split_lines};

#[test]
fn converts_original_american_spellings_to_british() {
    let text = "The Savior showed favor to his neighbor with honor and judgment.";
    assert_eq!(
        british_spellings(text),
        "The Saviour showed favour to his neighbour with honour and judgement."
    );
}

#[test]
fn converts_esv_verse_markers_to_superscript_runs() {
    let runs = scripture_runs("[1] In the beginning [2] And the earth");

    assert_eq!(runs[0].text, "1");
    assert!(runs[0].superscript);
    assert_eq!(runs[1].text, "In the beginning ");
    assert!(!runs[1].superscript);
    assert_eq!(runs[2].text, "2");
    assert!(runs[2].superscript);
}

#[test]
fn splits_long_text_without_losing_lines() {
    let text = "one\ntwo\nthree\nfour\nfive";
    let pages = split_lines(text, 2, 100);

    assert_eq!(pages, vec!["one\ntwo", "three\nfour", "five"]);
}
