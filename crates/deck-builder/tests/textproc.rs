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
fn removes_esv_verse_markers_from_plain_runs() {
    let runs = scripture_runs("[1] In the beginning [2] And the earth");

    assert_eq!(runs.iter().map(|run| run.text.as_str()).collect::<String>(), "In the beginning And the earth");
    assert!(runs.iter().all(|run| !run.superscript));
}

#[test]
fn splits_long_text_without_losing_lines() {
    let text = "one\ntwo\nthree\nfour\nfive";
    let pages = split_lines(text, 2, 100);

    assert_eq!(pages, vec!["one\ntwo", "three\nfour", "five"]);
}
