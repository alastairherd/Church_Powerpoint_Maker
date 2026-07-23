mod songs;
pub mod store;

use argon2::{Argon2, PasswordHash, PasswordVerifier};
use askama::Template;
use async_trait::async_trait;
use axum::body::Body;
use axum::extract::{Form, Path, Query, Request, State};
use axum::middleware::{self, Next};
use axum::response::{Html, IntoResponse, Redirect, Response};
use axum::routing::{get, post, put};
use axum::{Extension, Json, Router};
use chrono::{Duration as ChronoDuration, NaiveDate, Utc};
use deck_builder::{
    build_deck, propose_psalm_groups, Catechism, FixedComponent, GeneratedDeckVersion,
    GlobalSettingsVersion, Psalm, ServicePreset, ServiceRecord, ServiceStatus,
    Sources, StoredSong, TeachingSource,
};
use hmac::{Hmac, Mac};
use http::header::{ACCEPT, CONTENT_DISPOSITION, CONTENT_TYPE, COOKIE, ETAG, SET_COOKIE};
use http::{HeaderMap, HeaderValue, Method, StatusCode};
use serde::{Deserialize, Serialize};
use serde_json::json;
use sha2::Sha256;
use std::collections::HashMap;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use store::{MemoryObjectStore, ObjectStore, PutCondition, StoreError, StoredObject};

pub(crate) const PPTX_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const SESSION_COOKIE: &str = "twpc_session";
const JSON_CONTENT_TYPE: &str = "application/json";
const LOGIN_WINDOW: Duration = Duration::from_secs(15 * 60);
const LOGIN_ATTEMPTS: usize = 8;
const GENERATED_DECK_RETENTION_DAYS: i64 = 730;

#[derive(Clone)]
pub struct AppConfig {
    pub password_hash: String,
    pub session_signing_secret: String,
    pub secure_cookies: bool,
    pub session_ttl: Duration,
}

impl AppConfig {
    pub fn from_env() -> anyhow::Result<Self> {
        let password_hash = std::env::var("STAFF_PASSWORD_HASH")
            .map_err(|_| anyhow::anyhow!("STAFF_PASSWORD_HASH is required"))?;
        PasswordHash::new(&password_hash)
            .map_err(|_| anyhow::anyhow!("STAFF_PASSWORD_HASH is not a valid Argon2 PHC hash"))?;
        let session_signing_secret = std::env::var("SESSION_SIGNING_SECRET")
            .map_err(|_| anyhow::anyhow!("SESSION_SIGNING_SECRET is required"))?;
        if session_signing_secret.len() < 32 {
            return Err(anyhow::anyhow!(
                "SESSION_SIGNING_SECRET must contain at least 32 characters"
            ));
        }
        Ok(Self {
            password_hash,
            session_signing_secret,
            secure_cookies: std::env::var("COOKIE_SECURE")
                .map(|value| value != "false")
                .unwrap_or(true),
            session_ttl: Duration::from_secs(8 * 60 * 60),
        })
    }
}

#[derive(Clone)]
pub(crate) struct AppState {
    sources: Arc<dyn Sources>,
    pub(crate) store: Arc<dyn ObjectStore>,
    config: AppConfig,
    login_limiter: Arc<Mutex<HashMap<String, Vec<Instant>>>>,
    next_id: Arc<AtomicU64>,
}

struct ServiceSources {
    upstream: Arc<dyn Sources>,
    store: Arc<dyn ObjectStore>,
}

#[async_trait]
impl Sources for ServiceSources {
    async fn scripture(&self, reference: &str) -> anyhow::Result<deck_builder::Scripture> {
        self.upstream.scripture(reference).await
    }

    async fn song(&self, id: &str, version: u64) -> anyhow::Result<StoredSong> {
        songs::resolve(self.store.as_ref(), id, version).await
    }

    fn psalm(&self, reference: &str) -> anyhow::Result<Psalm> {
        self.upstream.psalm(reference)
    }

    fn catechism(&self, question: u16) -> anyhow::Result<Catechism> {
        self.upstream.catechism(question)
    }

    fn fixed_component(&self, key: &str) -> anyhow::Result<FixedComponent> {
        self.upstream.fixed_component(key)
    }
}

pub fn app(sources: Arc<dyn Sources>, store: Arc<dyn ObjectStore>, config: AppConfig) -> Router {
    let state = AppState {
        sources,
        store,
        config,
        login_limiter: Arc::new(Mutex::new(HashMap::new())),
        next_id: Arc::new(AtomicU64::new(1)),
    };

    let protected = Router::new()
        .route("/", get(builder_page))
        .route("/library", get(library_page))
        .route("/admin", get(admin_page))
        .route("/generated", get(generated_page))
        .route("/api/session", get(current_session))
        .route("/api/logout", post(logout))
        .route("/api/presets", get(list_presets))
        .route("/api/scripture", get(fetch_scripture))
        .route("/api/psalm", get(fetch_psalm))
        .route("/api/teaching", get(fetch_teaching))
        .route("/api/songs", get(songs::list).post(songs::create))
        .route("/api/songs/:id", get(songs::get).delete(songs::archive))
        .route("/api/songs/:id/restore", post(songs::restore))
        .route("/api/songs/:id/upload", post(songs::upload))
        .route("/api/songs/:id/preview", get(songs::preview))
        .route("/api/settings", get(get_settings).put(update_settings))
        .route("/api/services", get(list_services).post(create_service))
        .route(
            "/api/services/:id",
            get(get_service).put(update_service).delete(archive_service),
        )
        .route("/api/services/:id/restore", post(restore_service))
        .route("/api/services/:id/autosave", put(update_service))
        .route("/api/services/:id/generate", post(generate_service))
        .route("/api/services/:id/history", get(service_history))
        .route("/api/generated", get(generated_decks))
        .route(
            "/api/services/:id/revisions/:revision/download",
            get(download_service_revision),
        )
        .route_layer(middleware::from_fn_with_state(state.clone(), require_staff))
        .with_state(state.clone());

    Router::new()
        .route("/login", get(login_page).post(login))
        .route("/healthz", get(healthz))
        .route("/static/app.css", get(stylesheet))
        .route("/static/app.js", get(javascript))
        .route(
            "/static/editor-controller.js",
            get(editor_controller_javascript),
        )
        .route("/static/library.js", get(library_javascript))
        .route("/static/admin.js", get(admin_javascript))
        .route("/static/generated.js", get(generated_javascript))
        .route("/favicon.svg", get(favicon))
        .merge(protected)
        .with_state(state)
}

pub fn app_with_sources(sources: Arc<dyn Sources>, config: AppConfig) -> Router {
    app(sources, Arc::new(MemoryObjectStore::default()), config)
}

async fn healthz() -> &'static str {
    "ok"
}

async fn stylesheet() -> impl IntoResponse {
    (
        [(CONTENT_TYPE, "text/css; charset=utf-8")],
        include_str!("../static/app.css"),
    )
}

async fn javascript() -> impl IntoResponse {
    (
        [(CONTENT_TYPE, "text/javascript; charset=utf-8")],
        include_str!("../static/app.js"),
    )
}

async fn editor_controller_javascript() -> impl IntoResponse {
    (
        [(CONTENT_TYPE, "text/javascript; charset=utf-8")],
        include_str!("../static/editor-controller.js"),
    )
}

async fn library_javascript() -> impl IntoResponse {
    (
        [(CONTENT_TYPE, "text/javascript; charset=utf-8")],
        include_str!("../static/library.js"),
    )
}

async fn admin_javascript() -> impl IntoResponse {
    (
        [(CONTENT_TYPE, "text/javascript; charset=utf-8")],
        include_str!("../static/admin.js"),
    )
}

async fn generated_javascript() -> impl IntoResponse {
    (
        [(CONTENT_TYPE, "text/javascript; charset=utf-8")],
        include_str!("../static/generated.js"),
    )
}

async fn favicon() -> impl IntoResponse {
    (
        [(CONTENT_TYPE, "image/svg+xml; charset=utf-8")],
        include_str!("../static/favicon.svg"),
    )
}

#[derive(Template)]
#[template(path = "login.html")]
struct LoginTemplate {
    has_error: bool,
    error: String,
}

#[derive(Template)]
#[template(path = "builder.html")]
struct BuilderTemplate {
    staff_name: String,
    staff_initial: String,
    csrf: String,
}

#[derive(Template)]
#[template(path = "library.html")]
struct LibraryTemplate {
    staff_name: String,
    staff_initial: String,
    csrf: String,
}

#[derive(Template)]
#[template(path = "admin.html")]
struct AdminTemplate {
    staff_name: String,
    staff_initial: String,
    csrf: String,
}

#[derive(Template)]
#[template(path = "generated.html")]
struct GeneratedTemplate {
    staff_name: String,
    staff_initial: String,
    csrf: String,
}

async fn login_page() -> Result<Html<String>, AppError> {
    render(LoginTemplate {
        has_error: false,
        error: String::new(),
    })
}

#[derive(Deserialize)]
struct LoginForm {
    display_name: String,
    password: String,
}

async fn login(
    State(state): State<AppState>,
    headers: HeaderMap,
    Form(form): Form<LoginForm>,
) -> Result<Response, AppError> {
    check_login_rate_limit(&state, &headers)?;
    let display_name = validate_display_name(&form.display_name)?;
    let hash = PasswordHash::new(&state.config.password_hash)
        .map_err(|_| AppError::internal("staff password configuration is invalid"))?;
    if Argon2::default()
        .verify_password(form.password.as_bytes(), &hash)
        .is_err()
    {
        let html = LoginTemplate {
            has_error: true,
            error: "The shared password was not recognised.".into(),
        }
        .render()
        .map_err(|err| AppError::internal(err.to_string()))?;
        return Ok((StatusCode::UNAUTHORIZED, Html(html)).into_response());
    }

    let token = issue_session(&state, &display_name)?;
    let cookie = session_cookie(
        &token,
        state.config.secure_cookies,
        state.config.session_ttl,
    );
    let mut response = Redirect::to("/").into_response();
    response.headers_mut().insert(
        SET_COOKIE,
        HeaderValue::from_str(&cookie).map_err(|err| AppError::internal(err.to_string()))?,
    );
    Ok(response)
}

async fn logout(State(state): State<AppState>) -> Response {
    let secure = if state.config.secure_cookies {
        "; Secure"
    } else {
        ""
    };
    let mut response = StatusCode::NO_CONTENT.into_response();
    response.headers_mut().insert(
        SET_COOKIE,
        HeaderValue::from_str(&format!(
            "{SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0{secure}"
        ))
        .expect("static cookie header"),
    );
    response
}

async fn builder_page(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
) -> Result<Html<String>, AppError> {
    let initial = staff_initial(&session.display_name);
    render(BuilderTemplate {
        staff_name: session.display_name,
        staff_initial: initial,
        csrf: csrf_for(&state, &session.token)?,
    })
}

async fn library_page(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
) -> Result<Html<String>, AppError> {
    render(LibraryTemplate {
        staff_initial: staff_initial(&session.display_name),
        staff_name: session.display_name,
        csrf: csrf_for(&state, &session.token)?,
    })
}

async fn admin_page(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
) -> Result<Html<String>, AppError> {
    render(AdminTemplate {
        staff_initial: staff_initial(&session.display_name),
        staff_name: session.display_name,
        csrf: csrf_for(&state, &session.token)?,
    })
}

async fn generated_page(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
) -> Result<Html<String>, AppError> {
    render(GeneratedTemplate {
        staff_initial: staff_initial(&session.display_name),
        staff_name: session.display_name,
        csrf: csrf_for(&state, &session.token)?,
    })
}

fn staff_initial(name: &str) -> String {
    name.chars()
        .next()
        .unwrap_or('S')
        .to_uppercase()
        .to_string()
}

async fn current_session(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
) -> Result<Json<serde_json::Value>, AppError> {
    Ok(Json(json!({
        "display_name": session.display_name,
        "expires_at": session.expires_at,
        "csrf": csrf_for(&state, &session.token)?,
    })))
}

#[derive(Serialize)]
struct PresetResponse {
    id: ServicePreset,
    label: &'static str,
    components: Vec<deck_builder::ServiceComponent>,
}

async fn list_presets() -> Json<Vec<PresetResponse>> {
    Json(
        ServicePreset::all()
            .into_iter()
            .map(|preset| PresetResponse {
                id: preset,
                label: preset.label(),
                components: preset.components(),
            })
            .collect(),
    )
}

#[derive(Deserialize)]
struct ScriptureQuery {
    reference: String,
}

async fn fetch_scripture(
    State(state): State<AppState>,
    Query(query): Query<ScriptureQuery>,
) -> Json<serde_json::Value> {
    match state.sources.scripture(query.reference.trim()).await {
        Ok(scripture) => Json(json!({
            "ok": true,
            "reference": scripture.reference,
            "text": scripture.text,
        })),
        Err(error) => Json(json!({
            "ok": false,
            "reference": query.reference,
            "text": "",
            "warning": format!("ESV text could not be fetched. Enter the text manually. {error}"),
        })),
    }
}

#[derive(Deserialize)]
struct PsalmQuery {
    reference: String,
}

async fn fetch_psalm(
    State(state): State<AppState>,
    Query(query): Query<PsalmQuery>,
) -> Result<Json<serde_json::Value>, AppError> {
    let reference = query.reference.trim();
    if reference.is_empty() {
        return Err(AppError::bad_request("enter a Psalm reference first"));
    }
    let psalm = state
        .sources
        .psalm(reference)
        .map_err(|error| AppError::bad_request(error.to_string()))?;
    Ok(Json(json!({
        "reference": psalm.title,
        "meter": psalm.meter,
        "slides": propose_psalm_groups(&psalm.stanzas),
    })))
}

#[derive(Deserialize)]
struct TeachingQuery {
    source: TeachingSource,
    selection: String,
}

async fn fetch_teaching(
    State(state): State<AppState>,
    Query(query): Query<TeachingQuery>,
) -> Result<Json<serde_json::Value>, AppError> {
    if query.source != TeachingSource::WestminsterShorterCatechism {
        return Err(AppError::bad_request(
            "automatic loading is only available for the Westminster Shorter Catechism; enter this source manually",
        ));
    }
    let number = deck_builder::parse_catechism_selection(&query.selection)
        .map_err(|error| AppError::bad_request(error.to_string()))?;
    let item = state
        .sources
        .catechism(number)
        .map_err(|error| AppError::bad_request(error.to_string()))?;
    Ok(Json(json!({
        "source": query.source,
        "selection": number,
        "question": item.question,
        "answer": item.answer,
    })))
}

#[derive(Deserialize)]
struct CreateService {
    name: String,
    date: NaiveDate,
    preset: ServicePreset,
}

async fn create_service(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Json(input): Json<CreateService>,
) -> Result<Response, AppError> {
    validate_service_name(&input.name)?;
    let id = new_id(&state, "svc");
    let service = ServiceRecord::new(
        id,
        input.name,
        input.date,
        input.preset,
        session.display_name,
    );
    let stored = put_json(
        state.store.as_ref(),
        &service_key(&service.id),
        &service,
        PutCondition::IfNoneMatch,
    )
    .await?;
    json_with_etag(StatusCode::CREATED, &service, &stored.etag)
}

async fn list_services(
    State(state): State<AppState>,
) -> Result<Json<Vec<ServiceRecord>>, AppError> {
    let mut services = Vec::new();
    for key in state.store.list("entities/services/").await? {
        if let Ok(object) = state.store.get(&key).await {
            if let Ok(service) = serde_json::from_slice::<ServiceRecord>(&object.bytes) {
                services.push(service);
            }
        }
    }
    services.sort_by(|a, b| b.date.cmp(&a.date).then_with(|| a.name.cmp(&b.name)));
    Ok(Json(services))
}

async fn get_service(
    State(state): State<AppState>,
    Path(id): Path<String>,
) -> Result<Response, AppError> {
    let (service, object) = load_service(&state, &id).await?;
    json_with_etag(StatusCode::OK, &service, &object.etag)
}

async fn update_service(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Path(id): Path<String>,
    Json(mut incoming): Json<ServiceRecord>,
) -> Result<Response, AppError> {
    if incoming.id != id {
        return Err(AppError::bad_request("service ID does not match the route"));
    }
    validate_service_name(&incoming.name)?;
    let (current, object) = load_service(&state, &id).await?;
    if incoming.revision != current.revision {
        return Err(AppError::conflict(
            "this service changed in another browser; reload before saving",
        ));
    }
    incoming.audit = current.audit.clone();
    incoming.status = if current.status == ServiceStatus::Completed {
        ServiceStatus::Draft
    } else {
        incoming.status
    };
    incoming.mark_edited(&session.display_name);
    let stored = put_json(
        state.store.as_ref(),
        &service_key(&id),
        &incoming,
        PutCondition::IfMatch(object.etag),
    )
    .await?;
    json_with_etag(StatusCode::OK, &incoming, &stored.etag)
}

async fn archive_service(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Path(id): Path<String>,
) -> Result<Json<ServiceRecord>, AppError> {
    change_service_status(&state, &id, ServiceStatus::Archived, &session.display_name).await
}

async fn restore_service(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Path(id): Path<String>,
) -> Result<Json<ServiceRecord>, AppError> {
    change_service_status(&state, &id, ServiceStatus::Draft, &session.display_name).await
}

async fn change_service_status(
    state: &AppState,
    id: &str,
    status: ServiceStatus,
    staff_name: &str,
) -> Result<Json<ServiceRecord>, AppError> {
    let (mut service, object) = load_service(state, id).await?;
    service.status = status;
    service.mark_edited(staff_name);
    put_json(
        state.store.as_ref(),
        &service_key(id),
        &service,
        PutCondition::IfMatch(object.etag),
    )
    .await?;
    Ok(Json(service))
}

async fn generate_service(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Path(id): Path<String>,
) -> Result<Response, AppError> {
    let (mut service, object) = load_service(&state, &id).await?;
    let settings = load_settings(&state).await?;
    let sources = ServiceSources {
        upstream: state.sources.clone(),
        store: state.store.clone(),
    };
    let bytes = build_deck(&service, &sources, &settings.ccli_licence_number)
        .await
        .map_err(|error| AppError::internal(format!("could not build service deck: {error}")))?;
    let revision = state
        .store
        .list(&format!("generated/services/{id}/revisions/"))
        .await?
        .into_iter()
        .filter(|key| key.ends_with(".pptx"))
        .count() as u64
        + 1;
    let deck_key = format!("generated/services/{id}/revisions/{revision}.pptx");
    state
        .store
        .put(
            &deck_key,
            bytes.clone(),
            PPTX_CONTENT_TYPE,
            PutCondition::IfNoneMatch,
        )
        .await?;
    let generated_at = Utc::now();
    let record = GeneratedDeckVersion {
        service_id: id.clone(),
        revision,
        object_key: deck_key,
        generated_at,
        generated_by: session.display_name.clone(),
        expires_at: generated_at + ChronoDuration::days(GENERATED_DECK_RETENTION_DAYS),
        source_revision: service.revision,
    };
    put_json(
        state.store.as_ref(),
        &format!("entities/services/{id}/revisions/{revision}.json"),
        &record,
        PutCondition::IfNoneMatch,
    )
    .await?;
    service.status = ServiceStatus::Completed;
    service.audit.touch(&session.display_name);
    put_json(
        state.store.as_ref(),
        &service_key(&id),
        &service,
        PutCondition::IfMatch(object.etag),
    )
    .await?;

    let mut response = Response::new(Body::from(bytes));
    response
        .headers_mut()
        .insert(CONTENT_TYPE, HeaderValue::from_static(PPTX_CONTENT_TYPE));
    response.headers_mut().insert(
        CONTENT_DISPOSITION,
        HeaderValue::from_str(&format!(
            "attachment; filename=\"{}\"",
            deck_filename(&service, revision)
        ))
        .map_err(|error| AppError::internal(error.to_string()))?,
    );
    Ok(response)
}

async fn service_history(
    State(state): State<AppState>,
    Path(id): Path<String>,
) -> Result<Json<Vec<GeneratedDeckVersion>>, AppError> {
    load_service(&state, &id).await?;
    let mut revisions = Vec::new();
    for key in state
        .store
        .list(&format!("entities/services/{id}/revisions/"))
        .await?
    {
        let object = state.store.get(&key).await?;
        revisions.push(serde_json::from_slice(&object.bytes)?);
    }
    revisions.sort_by_key(|revision: &GeneratedDeckVersion| revision.revision);
    revisions.reverse();
    Ok(Json(revisions))
}

#[derive(Debug, Serialize)]
struct GeneratedDeckListing {
    service_id: String,
    service_name: String,
    service_date: NaiveDate,
    revision: u64,
    generated_at: chrono::DateTime<Utc>,
    generated_by: String,
    expires_at: chrono::DateTime<Utc>,
    source_revision: u64,
    download_url: String,
}

async fn generated_decks(
    State(state): State<AppState>,
) -> Result<Json<Vec<GeneratedDeckListing>>, AppError> {
    let mut generated = Vec::new();
    for key in state.store.list("entities/services/").await? {
        let parts: Vec<_> = key.split('/').collect();
        if parts.len() != 5 || parts[1] != "services" || parts[3] != "revisions" {
            continue;
        }
        let Some(file) = parts.last() else { continue };
        let Some(revision_text) = file.strip_suffix(".json") else {
            continue;
        };
        let Ok(revision) = revision_text.parse::<u64>() else {
            continue;
        };
        let object = state.store.get(&key).await?;
        let metadata: GeneratedDeckVersion = serde_json::from_slice(&object.bytes)?;
        if metadata.service_id != parts[2] || metadata.revision != revision {
            continue;
        }
        let (service, _) = load_service(&state, &metadata.service_id).await?;
        generated.push(GeneratedDeckListing {
            service_id: metadata.service_id.clone(),
            service_name: service.name,
            service_date: service.date,
            revision: metadata.revision,
            generated_at: metadata.generated_at,
            generated_by: metadata.generated_by,
            expires_at: metadata.expires_at,
            source_revision: metadata.source_revision,
            download_url: format!(
                "/api/services/{}/revisions/{}/download",
                metadata.service_id, metadata.revision
            ),
        });
    }
    generated.sort_by(|left, right| {
        right
            .generated_at
            .cmp(&left.generated_at)
            .then_with(|| right.service_id.cmp(&left.service_id))
            .then_with(|| right.revision.cmp(&left.revision))
    });
    Ok(Json(generated))
}

async fn download_service_revision(
    State(state): State<AppState>,
    Path((id, revision)): Path<(String, u64)>,
) -> Result<Response, AppError> {
    let (service, _) = load_service(&state, &id).await?;
    let metadata_object = state
        .store
        .get(&format!("entities/services/{id}/revisions/{revision}.json"))
        .await?;
    let metadata: GeneratedDeckVersion = serde_json::from_slice(&metadata_object.bytes)?;
    if metadata.service_id != id || metadata.revision != revision {
        return Err(AppError::new(StatusCode::NOT_FOUND, "record not found"));
    }
    let deck = state.store.get(&metadata.object_key).await?;

    let mut response = Response::new(Body::from(deck.bytes));
    response
        .headers_mut()
        .insert(CONTENT_TYPE, HeaderValue::from_static(PPTX_CONTENT_TYPE));
    response.headers_mut().insert(
        CONTENT_DISPOSITION,
        HeaderValue::from_str(&format!(
            "attachment; filename=\"{}\"",
            deck_filename(&service, revision)
        ))
        .map_err(|error| AppError::internal(error.to_string()))?,
    );
    Ok(response)
}

fn deck_filename(service: &ServiceRecord, revision: u64) -> String {
    let safe_name = service
        .name
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() || character == '-' {
                character
            } else {
                '-'
            }
        })
        .collect::<String>();
    format!("{}-{}-r{}.pptx", safe_name, service.date, revision)
}

async fn get_settings(
    State(state): State<AppState>,
) -> Result<Json<GlobalSettingsVersion>, AppError> {
    Ok(Json(load_settings(&state).await?))
}

#[derive(Deserialize)]
struct SettingsInput {
    ccli_licence_number: String,
}

async fn update_settings(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Json(input): Json<SettingsInput>,
) -> Result<Json<GlobalSettingsVersion>, AppError> {
    let licence = input.ccli_licence_number.trim();
    if licence.is_empty()
        || licence.len() > 32
        || !licence
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || matches!(character, '-' | ' '))
    {
        return Err(AppError::bad_request(
            "CCLI licence number must contain 1 to 32 letters, numbers, spaces or hyphens",
        ));
    }
    let (current, object) = load_settings_object(&state).await?;
    let settings = GlobalSettingsVersion {
        version: current.version.saturating_add(1),
        ccli_licence_number: licence.to_string(),
        created_at: Utc::now(),
        created_by: session.display_name,
    };
    put_json(
        state.store.as_ref(),
        &format!("entities/settings/versions/{}.json", settings.version),
        &settings,
        PutCondition::IfNoneMatch,
    )
    .await?;
    put_json(
        state.store.as_ref(),
        "entities/settings/current.json",
        &settings,
        PutCondition::IfMatch(object.etag),
    )
    .await?;
    Ok(Json(settings))
}

async fn load_settings(state: &AppState) -> Result<GlobalSettingsVersion, AppError> {
    Ok(load_settings_object(state).await?.0)
}

async fn load_settings_object(
    state: &AppState,
) -> Result<(GlobalSettingsVersion, StoredObject), AppError> {
    match state.store.get("entities/settings/current.json").await {
        Ok(object) => Ok((serde_json::from_slice(&object.bytes)?, object)),
        Err(StoreError::NotFound(_)) => {
            let settings = GlobalSettingsVersion::default();
            let object = put_json(
                state.store.as_ref(),
                "entities/settings/current.json",
                &settings,
                PutCondition::IfNoneMatch,
            )
            .await?;
            Ok((settings, object))
        }
        Err(error) => Err(error.into()),
    }
}

#[derive(Debug, Clone)]
pub(crate) struct StaffSession {
    pub(crate) display_name: String,
    expires_at: u64,
    token: String,
}

async fn require_staff(
    State(state): State<AppState>,
    mut request: Request,
    next: Next,
) -> Response {
    let session = request
        .headers()
        .get(COOKIE)
        .and_then(|value| value.to_str().ok())
        .and_then(|cookies| cookie_value(cookies, SESSION_COOKIE))
        .and_then(|token| verify_session(&state, token).ok());
    let Some(session) = session else {
        let wants_html = request
            .headers()
            .get(ACCEPT)
            .and_then(|value| value.to_str().ok())
            .is_some_and(|value| value.contains("text/html"));
        return if wants_html {
            Redirect::to("/login").into_response()
        } else {
            AppError::unauthorised("staff sign-in required").into_response()
        };
    };

    if matches!(
        *request.method(),
        Method::POST | Method::PUT | Method::PATCH | Method::DELETE
    ) {
        let expected = csrf_for(&state, &session.token).ok();
        let supplied = request
            .headers()
            .get("x-csrf-token")
            .and_then(|value| value.to_str().ok());
        if expected.as_deref() != supplied {
            return AppError::forbidden("CSRF token is missing or invalid").into_response();
        }
    }
    request.extensions_mut().insert(session);
    next.run(request).await
}

fn issue_session(state: &AppState, display_name: &str) -> Result<String, AppError> {
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|error| AppError::internal(error.to_string()))?;
    let expires_at = now.as_secs() + state.config.session_ttl.as_secs();
    let nonce = state.next_id.fetch_add(1, Ordering::Relaxed);
    let body = format!("{expires_at}:{nonce}:{display_name}");
    let signature = sign(&state.config.session_signing_secret, &body)?;
    Ok(format!("{body}:{signature}"))
}

fn verify_session(state: &AppState, token: &str) -> Result<StaffSession, AppError> {
    let (body, supplied_signature) = token
        .rsplit_once(':')
        .ok_or_else(|| AppError::unauthorised("invalid session"))?;
    let expected_signature = sign(&state.config.session_signing_secret, body)?;
    if !constant_time_equal(supplied_signature.as_bytes(), expected_signature.as_bytes()) {
        return Err(AppError::unauthorised("invalid session"));
    }
    let mut parts = body.splitn(3, ':');
    let expires_at = parts
        .next()
        .and_then(|value| value.parse::<u64>().ok())
        .ok_or_else(|| AppError::unauthorised("invalid session"))?;
    let _nonce = parts
        .next()
        .ok_or_else(|| AppError::unauthorised("invalid session"))?;
    let display_name = parts
        .next()
        .filter(|name| !name.is_empty())
        .ok_or_else(|| AppError::unauthorised("invalid session"))?;
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|error| AppError::internal(error.to_string()))?
        .as_secs();
    if expires_at <= now {
        return Err(AppError::unauthorised("session expired"));
    }
    Ok(StaffSession {
        display_name: display_name.to_string(),
        expires_at,
        token: token.to_string(),
    })
}

fn csrf_for(state: &AppState, token: &str) -> Result<String, AppError> {
    sign(
        &state.config.session_signing_secret,
        &format!("csrf:{token}"),
    )
}

fn sign(secret: &str, value: &str) -> Result<String, AppError> {
    let mut mac = Hmac::<Sha256>::new_from_slice(secret.as_bytes())
        .map_err(|_| AppError::internal("session signing secret is invalid"))?;
    mac.update(value.as_bytes());
    Ok(hex(mac.finalize().into_bytes().as_slice()))
}

pub(crate) fn hex(bytes: &[u8]) -> String {
    const DIGITS: &[u8; 16] = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        output.push(DIGITS[(byte >> 4) as usize] as char);
        output.push(DIGITS[(byte & 0x0f) as usize] as char);
    }
    output
}

fn constant_time_equal(left: &[u8], right: &[u8]) -> bool {
    if left.len() != right.len() {
        return false;
    }
    left.iter()
        .zip(right)
        .fold(0_u8, |difference, (left, right)| {
            difference | (left ^ right)
        })
        == 0
}

fn cookie_value<'a>(cookies: &'a str, name: &str) -> Option<&'a str> {
    cookies.split(';').find_map(|cookie| {
        let (key, value) = cookie.trim().split_once('=')?;
        (key == name).then_some(value)
    })
}

fn session_cookie(token: &str, secure: bool, ttl: Duration) -> String {
    let secure = if secure { "; Secure" } else { "" };
    format!(
        "{SESSION_COOKIE}={token}; Path=/; HttpOnly; SameSite=Strict; Max-Age={}{}",
        ttl.as_secs(),
        secure
    )
}

fn validate_display_name(name: &str) -> Result<String, AppError> {
    let name = name.trim();
    if name.len() < 2 || name.len() > 60 {
        return Err(AppError::bad_request(
            "your name must contain between 2 and 60 characters",
        ));
    }
    if !name.chars().all(|character| {
        character.is_alphanumeric()
            || character.is_whitespace()
            || matches!(character, '-' | '\'' | '.')
    }) {
        return Err(AppError::bad_request(
            "your name contains an unsupported character",
        ));
    }
    Ok(name.to_string())
}

fn check_login_rate_limit(state: &AppState, headers: &HeaderMap) -> Result<(), AppError> {
    let client = client_address(headers);
    let now = Instant::now();
    let mut limiter = state
        .login_limiter
        .lock()
        .map_err(|_| AppError::internal("login limiter unavailable"))?;
    let attempts = limiter.entry(client).or_default();
    attempts.retain(|attempt| now.duration_since(*attempt) < LOGIN_WINDOW);
    if attempts.len() >= LOGIN_ATTEMPTS {
        return Err(AppError::new(
            StatusCode::TOO_MANY_REQUESTS,
            "too many sign-in attempts; wait before trying again",
        ));
    }
    attempts.push(now);
    Ok(())
}

fn client_address(headers: &HeaderMap) -> String {
    headers
        .get("x-forwarded-for")
        .and_then(|value| value.to_str().ok())
        .and_then(|value| value.split(',').next())
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .unwrap_or("unknown")
        .to_string()
}

fn validate_service_name(name: &str) -> Result<(), AppError> {
    if name.trim().is_empty() || name.len() > 100 {
        return Err(AppError::bad_request(
            "service name must contain between 1 and 100 characters",
        ));
    }
    Ok(())
}

pub(crate) fn new_id(state: &AppState, prefix: &str) -> String {
    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis())
        .unwrap_or_default();
    let counter = state.next_id.fetch_add(1, Ordering::Relaxed);
    format!("{prefix}-{timestamp:x}-{counter:x}")
}

fn service_key(id: &str) -> String {
    format!("entities/services/{id}.json")
}

async fn load_service(
    state: &AppState,
    id: &str,
) -> Result<(ServiceRecord, StoredObject), AppError> {
    let object = state.store.get(&service_key(id)).await?;
    let service = serde_json::from_slice(&object.bytes)?;
    Ok((service, object))
}

pub(crate) async fn put_json<T: Serialize + ?Sized>(
    store: &dyn ObjectStore,
    key: &str,
    value: &T,
    condition: PutCondition,
) -> Result<StoredObject, AppError> {
    let bytes = serde_json::to_vec_pretty(value)?;
    Ok(store.put(key, bytes, JSON_CONTENT_TYPE, condition).await?)
}

fn json_with_etag<T: Serialize>(
    status: StatusCode,
    value: &T,
    etag: &str,
) -> Result<Response, AppError> {
    let mut response = (status, Json(value)).into_response();
    response.headers_mut().insert(
        ETAG,
        HeaderValue::from_str(etag).map_err(|error| AppError::internal(error.to_string()))?,
    );
    Ok(response)
}

fn render(template: impl Template) -> Result<Html<String>, AppError> {
    Ok(Html(
        template
            .render()
            .map_err(|error| AppError::internal(error.to_string()))?,
    ))
}

#[derive(Debug)]
pub(crate) struct AppError {
    status: StatusCode,
    message: String,
}

impl AppError {
    fn new(status: StatusCode, message: impl Into<String>) -> Self {
        Self {
            status,
            message: message.into(),
        }
    }
    pub(crate) fn bad_request(message: impl Into<String>) -> Self {
        Self::new(StatusCode::BAD_REQUEST, message)
    }
    fn unauthorised(message: impl Into<String>) -> Self {
        Self::new(StatusCode::UNAUTHORIZED, message)
    }
    fn forbidden(message: impl Into<String>) -> Self {
        Self::new(StatusCode::FORBIDDEN, message)
    }
    fn conflict(message: impl Into<String>) -> Self {
        Self::new(StatusCode::CONFLICT, message)
    }
    fn internal(message: impl Into<String>) -> Self {
        Self::new(StatusCode::INTERNAL_SERVER_ERROR, message)
    }
}

impl IntoResponse for AppError {
    fn into_response(self) -> Response {
        (self.status, Json(json!({ "error": self.message }))).into_response()
    }
}

impl From<serde_json::Error> for AppError {
    fn from(error: serde_json::Error) -> Self {
        Self::internal(error.to_string())
    }
}

impl From<StoreError> for AppError {
    fn from(error: StoreError) -> Self {
        match error {
            StoreError::NotFound(_) => Self::new(StatusCode::NOT_FOUND, "record not found"),
            StoreError::PreconditionFailed => Self::conflict(
                "this record was changed by another staff member; reload and try again",
            ),
            StoreError::Unavailable(message) => Self::internal(message),
        }
    }
}
