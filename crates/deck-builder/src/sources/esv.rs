use crate::{textproc, Scripture};
use anyhow::{anyhow, Context};
use reqwest::header::{HeaderMap, HeaderValue, AUTHORIZATION};
use reqwest::Client;
use serde::Deserialize;

#[derive(Clone)]
pub struct EsvClient {
    client: Client,
    api_key: String,
}

impl EsvClient {
    pub fn new(api_key: impl Into<String>) -> anyhow::Result<Self> {
        let api_key = api_key.into();
        if api_key.trim().is_empty() {
            return Err(anyhow!("ESV_API_KEY is required"));
        }
        Ok(Self {
            client: Client::new(),
            api_key,
        })
    }

    pub async fn passage(&self, reference: &str) -> anyhow::Result<Scripture> {
        let mut headers = HeaderMap::new();
        headers.insert(
            AUTHORIZATION,
            HeaderValue::from_str(&format!("Token {}", self.api_key))
                .context("build ESV authorization header")?,
        );

        let response = self
            .client
            .get("https://api.esv.org/v3/passage/text/")
            .headers(headers)
            .query(&[
                ("q", reference),
                ("include-headings", "false"),
                ("include-footnotes", "false"),
                ("include-verse-numbers", "true"),
                ("include-short-copyright", "false"),
                ("include-passage-references", "false"),
            ])
            .send()
            .await
            .context("request ESV passage")?;

        if response.status() == reqwest::StatusCode::UNAUTHORIZED {
            return Err(anyhow!("ESV API rejected the configured API key"));
        }
        if !response.status().is_success() {
            return Err(anyhow!("ESV API returned {}", response.status()));
        }

        let body: EsvResponse = response.json().await.context("parse ESV response")?;
        let text = body
            .passages
            .first()
            .ok_or_else(|| anyhow!("ESV API returned no passages for {reference}"))?
            .trim()
            .to_string();
        Ok(Scripture {
            reference: reference.to_string(),
            text: textproc::british_spellings(&text),
        })
    }
}

#[derive(Debug, Deserialize)]
struct EsvResponse {
    passages: Vec<String>,
}
