use chrono::Utc;
use deck_builder::{AuditMetadata, RightsStatus, SongRecord, SongVersion};
use pptx_template::Presentation;
use server::store::{ObjectStore, PutCondition, R2ObjectStore, StoreError};
use sha2::{Digest, Sha256};
use std::collections::HashMap;
use std::io::{Cursor, Read};
use zip::ZipArchive;

const MARKER: &str = "imports/supplied-song-library-2026-07-18.json";

#[derive(serde::Serialize)]
struct ImportMarker {
    imported_at: chrono::DateTime<Utc>,
    song_count: usize,
    slide_count: usize,
}

struct Candidate {
    filename: String,
    title: String,
    bytes: Vec<u8>,
    slides: Vec<String>,
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let archive_path = std::env::args().nth(1).ok_or_else(|| {
        anyhow::anyhow!("usage: import-song-library <supplied-archive.zip> [--dry-run]")
    })?;
    let dry_run = std::env::args().any(|argument| argument == "--dry-run");
    let store = if dry_run {
        None
    } else {
        Some(R2ObjectStore::from_env()?)
    };
    if let Some(store) = &store {
        if store.get(MARKER).await.is_ok() {
            println!("Song library import has already completed.");
            return Ok(());
        }
    }

    let bytes = std::fs::read(&archive_path)?;
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut candidates = Vec::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index)?;
        if !file.name().contains("song slides/") || !file.name().ends_with(".pptx") {
            continue;
        }
        let filename = file
            .name()
            .rsplit('/')
            .next()
            .unwrap_or("song.pptx")
            .to_string();
        let mut deck_bytes = Vec::new();
        file.read_to_end(&mut deck_bytes)?;
        let presentation = Presentation::open_bytes(&deck_bytes)?;
        presentation.validate_song_source((10_080_625, 7_559_675))?;
        let slides = (0..presentation.slide_count())
            .map(|slide| presentation.slide_text(slide))
            .collect::<Result<Vec<_>, _>>()?;
        let title = title_from_filename(&filename);
        candidates.push(Candidate {
            filename,
            title,
            bytes: deck_bytes,
            slides,
        });
    }
    candidates.sort_by_key(|candidate| candidate.filename.to_lowercase());
    let slide_count: usize = candidates
        .iter()
        .map(|candidate| candidate.slides.len())
        .sum();
    if candidates.len() != 86 || slide_count != 365 {
        anyhow::bail!(
            "archive validation failed: found {} songs and {} slides; expected 86 and 365",
            candidates.len(),
            slide_count
        );
    }

    let mut occurrences: HashMap<String, usize> = HashMap::new();
    for candidate in &candidates {
        *occurrences
            .entry(normalise_title(&candidate.title))
            .or_default() += 1;
    }
    let mut variants: HashMap<String, usize> = HashMap::new();
    for candidate in candidates {
        let hash = Sha256::digest(&candidate.bytes);
        let hash_hex = hex(&hash);
        let id = format!("song-{}", &hash_hex[..16]);
        let normalised = normalise_title(&candidate.title);
        let variant_number = variants.entry(normalised.clone()).or_default();
        *variant_number += 1;
        let variant_label = if occurrences[&normalised] > 1 {
            format!("Imported variant {}", *variant_number)
        } else {
            String::new()
        };
        let asset_key = format!("sources/songs/{id}/versions/1.pptx");
        let metadata_key = format!("entities/songs/{id}.json");
        let version_key = format!("entities/songs/{id}/versions/1-pptx.json");
        let now = Utc::now();
        let song = SongRecord {
            id: id.clone(),
            title: candidate.title,
            aliases: Vec::new(),
            variant_label,
            author_owner: String::new(),
            rights_status: RightsStatus::Unknown,
            ccli_song_number: None,
            source_filename: Some(candidate.filename),
            slide_count: candidate.slides.len(),
            current_version: 1,
            archived: false,
            review_warnings: vec![
                "Author or rights owner needs review.".into(),
                "Rights status needs review.".into(),
                "CCLI song number needs review where applicable.".into(),
            ],
            audit: AuditMetadata::new("one-time importer"),
        };
        let version = SongVersion {
            song_id: id,
            version: 1,
            object_key: asset_key.clone(),
            sha256: hash_hex,
            slide_count: candidate.slides.len(),
            extracted_text: candidate.slides,
            created_at: now,
            created_by: "one-time importer".into(),
        };
        if !dry_run {
            let store = store.as_ref().expect("R2 store exists outside dry run");
            put_if_missing(
                store,
                &asset_key,
                candidate.bytes,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            .await?;
            put_if_missing(
                store,
                &version_key,
                serde_json::to_vec_pretty(&version)?,
                "application/json",
            )
            .await?;
            put_if_missing(
                store,
                &metadata_key,
                serde_json::to_vec_pretty(&song)?,
                "application/json",
            )
            .await?;
        }
    }
    if !dry_run {
        let marker = ImportMarker {
            imported_at: Utc::now(),
            song_count: 86,
            slide_count: 365,
        };
        put_if_missing(
            store.as_ref().expect("R2 store exists outside dry run"),
            MARKER,
            serde_json::to_vec_pretty(&marker)?,
            "application/json",
        )
        .await?;
    }
    println!(
        "Validated 86 songs and 365 slides{}.",
        if dry_run { " (dry run)" } else { "" }
    );
    Ok(())
}

async fn put_if_missing(
    store: &dyn ObjectStore,
    key: &str,
    bytes: Vec<u8>,
    content_type: &str,
) -> Result<(), StoreError> {
    match store
        .put(key, bytes, content_type, PutCondition::IfNoneMatch)
        .await
    {
        Ok(_) | Err(StoreError::PreconditionFailed) => Ok(()),
        Err(error) => Err(error),
    }
}

fn title_from_filename(filename: &str) -> String {
    let stem = filename.strip_suffix(".pptx").unwrap_or(filename).trim();
    stem.strip_prefix("song - ")
        .or_else(|| stem.strip_prefix("Song - "))
        .unwrap_or(stem)
        .trim()
        .to_string()
}

fn normalise_title(title: &str) -> String {
    title
        .chars()
        .filter(|character| character.is_alphanumeric())
        .flat_map(char::to_lowercase)
        .collect()
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| format!("{byte:02x}")).collect()
}
