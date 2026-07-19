use crate::store::{PutCondition, StoredObject};
use crate::{hex, new_id, put_json, AppError, AppState, StaffSession, PPTX_CONTENT_TYPE};
use axum::body::{to_bytes, Body};
use axum::extract::{Path, Query, State};
use axum::{Extension, Json};
use chrono::Utc;
use deck_builder::{AuditMetadata, RightsStatus, SongRecord, SongVersion};
use http::HeaderMap;
use pptx_template::Presentation;
use serde::{Deserialize, Serialize};
use serde_json::json;
use sha2::{Digest, Sha256};

#[derive(Deserialize)]
pub(crate) struct SongQuery {
    #[serde(default)]
    q: String,
    #[serde(default)]
    archived: bool,
}

pub(crate) async fn list(
    State(state): State<AppState>,
    Query(query): Query<SongQuery>,
) -> Result<Json<Vec<SongRecord>>, AppError> {
    let search = query.q.trim().to_lowercase();
    let mut songs = Vec::new();
    for key in state.store.list("entities/songs/").await? {
        if key.contains("/versions/") {
            continue;
        }
        let object = state.store.get(&key).await?;
        if let Ok(song) = serde_json::from_slice::<SongRecord>(&object.bytes) {
            let matches = search.is_empty()
                || song.title.to_lowercase().contains(&search)
                || song
                    .aliases
                    .iter()
                    .any(|alias| alias.to_lowercase().contains(&search));
            if matches && (query.archived || !song.archived) {
                songs.push(song);
            }
        }
    }
    songs.sort_by_key(|song| song.title.to_lowercase());
    Ok(Json(songs))
}

#[derive(Deserialize)]
pub(crate) struct CreateSong {
    title: String,
    #[serde(default)]
    aliases: Vec<String>,
    #[serde(default)]
    variant_label: String,
    #[serde(default)]
    author_owner: String,
    rights_status: RightsStatus,
    #[serde(default)]
    ccli_song_number: Option<String>,
    #[serde(default)]
    lyric_slides: Vec<String>,
    #[serde(default)]
    credits: String,
}

#[derive(Serialize, Deserialize)]
struct LyricVersion {
    song_id: String,
    version: u64,
    title: String,
    slides: Vec<String>,
    credits: String,
    created_at: chrono::DateTime<Utc>,
    created_by: String,
}

pub(crate) async fn create(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Json(input): Json<CreateSong>,
) -> Result<Json<SongRecord>, AppError> {
    validate_title(&input.title)?;
    let id = new_id(&state, "song");
    let record = SongRecord {
        id: id.clone(),
        title: input.title.clone(),
        aliases: input.aliases,
        variant_label: input.variant_label,
        author_owner: input.author_owner,
        rights_status: input.rights_status,
        ccli_song_number: input.ccli_song_number,
        source_filename: None,
        slide_count: input.lyric_slides.len(),
        current_version: 1,
        archived: false,
        review_warnings: warnings(input.rights_status, &input.lyric_slides),
        audit: AuditMetadata::new(&session.display_name),
    };
    let version = LyricVersion {
        song_id: id.clone(),
        version: 1,
        title: input.title,
        slides: input.lyric_slides,
        credits: input.credits,
        created_at: Utc::now(),
        created_by: session.display_name,
    };
    put_json(
        state.store.as_ref(),
        &lyric_version_key(&id, 1),
        &version,
        PutCondition::IfNoneMatch,
    )
    .await?;
    put_json(
        state.store.as_ref(),
        &song_key(&id),
        &record,
        PutCondition::IfNoneMatch,
    )
    .await?;
    Ok(Json(record))
}

pub(crate) async fn get(
    State(state): State<AppState>,
    Path(id): Path<String>,
) -> Result<Json<SongRecord>, AppError> {
    Ok(Json(load(&state, &id).await?.0))
}

pub(crate) async fn archive(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Path(id): Path<String>,
) -> Result<Json<SongRecord>, AppError> {
    set_archived(&state, &id, true, &session.display_name).await
}

pub(crate) async fn restore(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Path(id): Path<String>,
) -> Result<Json<SongRecord>, AppError> {
    set_archived(&state, &id, false, &session.display_name).await
}

async fn set_archived(
    state: &AppState,
    id: &str,
    archived: bool,
    staff: &str,
) -> Result<Json<SongRecord>, AppError> {
    let (mut song, object) = load(state, id).await?;
    song.archived = archived;
    song.audit.touch(staff);
    put_json(
        state.store.as_ref(),
        &song_key(id),
        &song,
        PutCondition::IfMatch(object.etag),
    )
    .await?;
    Ok(Json(song))
}

pub(crate) async fn upload(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Path(id): Path<String>,
    headers: HeaderMap,
    body: Body,
) -> Result<Json<SongRecord>, AppError> {
    let (mut song, metadata) = load(&state, &id).await?;
    let bytes = to_bytes(body, 25 * 1024 * 1024)
        .await
        .map_err(|_| AppError::bad_request("PowerPoint upload exceeds 25 MiB"))?;
    let presentation = Presentation::open_bytes(&bytes)
        .map_err(|error| AppError::bad_request(format!("PowerPoint was rejected: {error}")))?;
    presentation
        .validate_song_source((10_080_625, 7_559_675))
        .map_err(|error| AppError::bad_request(format!("PowerPoint was rejected: {error}")))?;
    let extracted_text = (0..presentation.slide_count())
        .map(|index| presentation.slide_text(index))
        .collect::<Result<Vec<_>, _>>()
        .map_err(|error| AppError::bad_request(error.to_string()))?;
    let version = song.current_version.saturating_add(1);
    let object_key = format!("sources/songs/{id}/versions/{version}.pptx");
    state
        .store
        .put(
            &object_key,
            bytes.to_vec(),
            PPTX_CONTENT_TYPE,
            PutCondition::IfNoneMatch,
        )
        .await?;
    let version_record = SongVersion {
        song_id: id.clone(),
        version,
        object_key,
        sha256: hex(Sha256::digest(&bytes).as_slice()),
        slide_count: presentation.slide_count(),
        extracted_text,
        created_at: Utc::now(),
        created_by: session.display_name.clone(),
    };
    put_json(
        state.store.as_ref(),
        &pptx_version_key(&id, version),
        &version_record,
        PutCondition::IfNoneMatch,
    )
    .await?;
    song.current_version = version;
    song.slide_count = presentation.slide_count();
    song.source_filename = headers
        .get("x-source-filename")
        .and_then(|value| value.to_str().ok())
        .map(safe_filename);
    song.audit.touch(session.display_name);
    put_json(
        state.store.as_ref(),
        &song_key(&id),
        &song,
        PutCondition::IfMatch(metadata.etag),
    )
    .await?;
    Ok(Json(song))
}

pub(crate) async fn preview(
    State(state): State<AppState>,
    Path(id): Path<String>,
) -> Result<Json<serde_json::Value>, AppError> {
    let (song, _) = load(&state, &id).await?;
    if let Ok(object) = state
        .store
        .get(&pptx_version_key(&id, song.current_version))
        .await
    {
        let version: SongVersion = serde_json::from_slice(&object.bytes)?;
        return Ok(Json(
            json!({ "song": song, "slides": version.extracted_text }),
        ));
    }
    let object = state
        .store
        .get(&lyric_version_key(&id, song.current_version))
        .await?;
    let version: LyricVersion = serde_json::from_slice(&object.bytes)?;
    Ok(Json(json!({ "song": song, "slides": version.slides })))
}

async fn load(state: &AppState, id: &str) -> Result<(SongRecord, StoredObject), AppError> {
    let object = state.store.get(&song_key(id)).await?;
    Ok((serde_json::from_slice(&object.bytes)?, object))
}

fn warnings(rights: RightsStatus, slides: &[String]) -> Vec<String> {
    let mut result = Vec::new();
    if rights == RightsStatus::Unknown {
        result.push("Rights status is unknown.".into());
    }
    if slides.iter().any(|slide| slide.len() > 620) {
        result.push("One or more lyric slides may be too dense.".into());
    }
    result
}

fn validate_title(title: &str) -> Result<(), AppError> {
    if title.trim().is_empty() || title.len() > 160 {
        return Err(AppError::bad_request(
            "song title must contain between 1 and 160 characters",
        ));
    }
    Ok(())
}

fn safe_filename(filename: &str) -> String {
    filename
        .chars()
        .filter(|character| {
            character.is_alphanumeric() || matches!(character, ' ' | '-' | '_' | '.')
        })
        .take(180)
        .collect()
}

fn song_key(id: &str) -> String {
    format!("entities/songs/{id}.json")
}
fn lyric_version_key(id: &str, version: u64) -> String {
    format!("entities/songs/{id}/versions/{version}-lyrics.json")
}
fn pptx_version_key(id: &str, version: u64) -> String {
    format!("entities/songs/{id}/versions/{version}-pptx.json")
}
