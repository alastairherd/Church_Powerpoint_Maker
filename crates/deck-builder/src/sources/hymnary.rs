use crate::Hymn;
use anyhow::{anyhow, Context};
use reqwest::Client;
use scraper::{Html, Selector};

#[derive(Clone)]
pub struct HymnaryClient {
    client: Client,
}

impl HymnaryClient {
    pub fn new() -> anyhow::Result<Self> {
        Ok(Self {
            client: Client::builder()
                .user_agent("church-deck/0.1 (+https://github.com/alastairherd)")
                .build()
                .context("build hymnary HTTP client")?,
        })
    }

    pub async fn hymn(&self, url: &str) -> anyhow::Result<Hymn> {
        validate_hymnary_url(url)?;
        let html = self
            .client
            .get(url)
            .send()
            .await
            .context("request hymnary page")?
            .error_for_status()
            .context("hymnary returned an error")?
            .text()
            .await
            .context("read hymnary page")?;
        parse_hymnary_page(&html)
    }
}

pub fn validate_hymnary_url(url: &str) -> anyhow::Result<()> {
    let parsed = reqwest::Url::parse(url).context("invalid hymn URL")?;
    let host = parsed.host_str().unwrap_or_default();
    if parsed.scheme() != "https" || !(host == "hymnary.org" || host.ends_with(".hymnary.org")) {
        return Err(anyhow!("hymn URLs must be https://hymnary.org/..."));
    }
    let path = parsed.path();
    if !(path.starts_with("/text/") || path.starts_with("/hymn/")) {
        return Err(anyhow!(
            "hymn URLs must point to hymnary.org text or hymn pages"
        ));
    }
    Ok(())
}

pub fn parse_hymnary_page(html: &str) -> anyhow::Result<Hymn> {
    let doc = Html::parse_document(html);
    let title = info_label(&doc, "Title:").unwrap_or_else(|| "Unable to find".to_string());
    let author = info_label(&doc, "Author:")
        .or_else(|| {
            info_label(&doc, "Author (attributed to):").map(|value| format!("{value} (atrb)"))
        })
        .unwrap_or_else(|| "Unable to find".to_string());
    let copyright = info_label(&doc, "Copyright:").unwrap_or_else(|| "Unable to find".to_string());

    let section = Selector::parse("div#at_fulltext.authority_section div div.authority_columns p")
        .expect("valid selector");
    let mut stanzas = Vec::new();
    let mut refrain: Option<String> = None;
    for p in doc.select(&section) {
        let text = p.text().collect::<Vec<_>>().join("").replace('\r', "");
        if let Some(rest) = text.strip_prefix("Refrain:") {
            let value = rest.trim().to_string();
            refrain = Some(value.clone());
            stanzas.push(value);
            continue;
        }
        if text.chars().next().is_some_and(|ch| ch.is_ascii_digit()) {
            let mut words = text.split_whitespace();
            words.next();
            let mut stanza = words.collect::<Vec<_>>().join(" ");
            if stanza.ends_with(" [Refrain]") {
                stanza = stanza.trim_end_matches(" [Refrain]").to_string();
                stanzas.push(stanza);
                if let Some(refrain) = &refrain {
                    stanzas.push(refrain.clone());
                }
            } else {
                stanzas.push(stanza);
            }
        }
    }

    if stanzas.is_empty() {
        stanzas.push("Unable to find".to_string());
    }

    Ok(Hymn {
        title,
        stanzas,
        author,
        composer: "Unknown".to_string(),
        tune: "Unable to find".to_string(),
        copyright,
    })
}

fn info_label(doc: &Html, label: &str) -> Option<String> {
    let selector = Selector::parse("span.hy_infoLabel").expect("valid selector");
    for span in doc.select(&selector) {
        if span.text().collect::<Vec<_>>().join("").trim() == label {
            if let Some(next) = span.next_sibling() {
                if let Some(element) = scraper::ElementRef::wrap(next) {
                    let value = element
                        .text()
                        .collect::<Vec<_>>()
                        .join("")
                        .trim()
                        .to_string();
                    if !value.is_empty() {
                        return Some(value);
                    }
                }
            }
        }
    }
    None
}
