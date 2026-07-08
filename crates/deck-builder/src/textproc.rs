use once_cell::sync::Lazy;
use pptx_template::Run;
use regex::Regex;

static BRITISH_REPLACEMENTS: Lazy<Vec<(Regex, &'static str)>> = Lazy::new(|| {
    [
        (r"(\W)([Ss])avior(s)?(\W)", r"${1}${2}aviour${3}${4}"),
        (r"(\W)neighbor(s|ing)?(\W)", r"${1}neighbour${2}${3}"),
        (r"(\W)favor(able|ite|s|ed|itism)?(\W)", r"${1}favour${2}${3}"),
        (r"(\W)Favor(\W)", r"${1}Favour${2}"),
        (r"(\W)labor(ed|s)?(\W)", r"${1}labour${2}${3}"),
        (r"(\W)(vap|vig)or(\W)", r"${1}${2}our${3}"),
        (r"(\W)clamor(\W)", r"${1}clamour${2}"),
        (r"(\W)([Ss])plendor(\W)", r"${1}${2}plendour${3}"),
        (r"(\W)color(s|ed)?(\W)", r"${1}colour${2}${3}"),
        (r"(\W)([Hh])onor(s|able|ing|ed)?(\W)", r"${1}${2}onour${3}${4}"),
        (r"(\W)dishonor(s|able|ing|ed)?(\W)", r"${1}dishonour${2}${3}"),
        (r"(\W)travel(ed|er|ers|ing)(\W)", r"${1}travell${2}${3}"),
        (r"(\W)marvel(ous|ously|ed|ing)(\W)", r"${1}marvell${2}${3}"),
        (r"(\W)([Cc])ounsel(or|ors|ed)(\W)", r"${1}${2}ounsell${3}${4}"),
        (r"(\W)plow(s|ed|ers|ing|man|men|share|shares)?(\W)", r"${1}plough${2}${3}"),
        (r"(\W)judgment(s)?(\W)", r"${1}judgement${2}${3}"),
        (
            r"(\W)(recogn|[Rr]eal|[Oo]rgan|[Ss]ymbol|bapt|critic|apolog|sympath)iz(e|ed|es|ing)?(\W)",
            r"${1}${2}is${3}${4}",
        ),
        (r"(\W)(un)?authorized(\W)", r"${1}${2}authorised${3}"),
        (r"(\W)(centi)?meters(\W)", r"${1}${2}metres${3}"),
        (r"(\W)liter(s)?(\W)", r"${1}litre${2}${3}"),
        (r"(\W)scepter(s)?(\W)", r"${1}sceptre${2}${3}"),
        (r"(\W)worship(ed|er|ers|ing)(\W)", r"${1}worshipp${2}${3}"),
        (r"(\W)quarrel(ed|ing)(\W)", r"${1}quarrell${2}${3}"),
        (r"(\W)benefited(\W)", r"${1}benefitted${2}"),
        (r"(\W)signaled(\W)", r"${1}signalled${2}"),
        (r"(\W)paralyzed(\W)", r"${1}paralysed${2}"),
        (r"(\W)fulfill(s|ment)?(\W)", r"${1}fulfil${2}${3}"),
        (r"(\W)skillful(ly)?(\W)", r"${1}skilful${2}${3}"),
        (r"(\W)jewelry(\W)", r"${1}jewellery${2}"),
        (r"(\W)(De|de|of)fense(s|less)?(\W)", r"${1}${2}fence${3}${4}"),
        (r"(\W)([Ss])ulfur?(\W)", r"${1}${2}ulphur${3}"),
    ]
    .into_iter()
    .map(|(pattern, replacement)| (Regex::new(pattern).expect("valid replacement regex"), replacement))
    .collect()
});

pub fn british_spellings(text: &str) -> String {
    let mut out = format!(" {text} ");
    for (pattern, replacement) in BRITISH_REPLACEMENTS.iter() {
        out = pattern.replace_all(&out, *replacement).to_string();
    }
    out.trim().to_string()
}

pub fn scripture_runs(text: &str) -> Vec<Run> {
    let marker = Regex::new(r"\[([^\[\]]+)\]\s*").expect("valid scripture marker regex");
    let mut runs = Vec::new();
    let mut last = 0;
    for cap in marker.captures_iter(text) {
        let full = cap.get(0).expect("full match");
        if full.start() > last {
            runs.push(Run::plain(&text[last..full.start()]));
        }
        runs.push(Run::superscript(cap[1].to_string()));
        last = full.end();
    }
    if last < text.len() {
        runs.push(Run::plain(&text[last..]));
    }
    runs
}

pub fn split_lines(text: &str, max_lines: usize, max_chars: usize) -> Vec<String> {
    if text.is_empty() {
        return Vec::new();
    }

    let mut pages = Vec::new();
    let mut current = Vec::new();
    let mut current_chars = 0;

    for line in text.lines() {
        let would_exceed_lines = !current.is_empty() && current.len() >= max_lines;
        let would_exceed_chars = !current.is_empty() && current_chars + line.len() + 1 > max_chars;
        if would_exceed_lines || would_exceed_chars {
            pages.push(current.join("\n"));
            current.clear();
            current_chars = 0;
        }
        current_chars += line.len() + usize::from(!current.is_empty());
        current.push(line.to_string());
    }

    if !current.is_empty() {
        pages.push(current.join("\n"));
    }
    pages
}

pub fn psalm_superscripts(text: &str) -> String {
    static RE: Lazy<Regex> =
        Lazy::new(|| Regex::new(r"(\d+-?\d*)([^\s\d])").expect("valid psalm superscript regex"));
    RE.replace_all(text, |caps: &regex::Captures<'_>| {
        format!("{}{}", superscript_digits(&caps[1]), &caps[2])
    })
    .to_string()
}

fn superscript_digits(text: &str) -> String {
    text.chars()
        .map(|ch| match ch {
            '0' => '⁰',
            '1' => '¹',
            '2' => '²',
            '3' => '³',
            '4' => '⁴',
            '5' => '⁵',
            '6' => '⁶',
            '7' => '⁷',
            '8' => '⁸',
            '9' => '⁹',
            '-' => '⁻',
            other => other,
        })
        .collect()
}
