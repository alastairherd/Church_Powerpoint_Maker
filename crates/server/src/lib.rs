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
    build_deck, propose_psalm_groups, AuditMetadata, FixedComponent, GeneratedDeckVersion,
    GlobalSettingsVersion, Psalm, ServicePreset, ServiceRecord, ServiceStatus, Sources, StoredSong,
    Teaching, TeachingSource,
};
use hmac::{Hmac, Mac};
use http::header::{ACCEPT, CONTENT_DISPOSITION, CONTENT_TYPE, COOKIE, ETAG, SET_COOKIE};
use http::{HeaderMap, HeaderValue, Method, StatusCode};
use pptx_template::Presentation;
use serde::{Deserialize, Serialize};
use serde_json::json;
use sha2::Sha256;
use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::mpsc::{self, Receiver, SyncSender};
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
const PREPARED_DECK_TTL: Duration = Duration::from_secs(20 * 60);
const PREPARED_DECK_LIMIT: usize = 4;
const PREPARED_DECK_BYTE_LIMIT: usize = 64 * 1024 * 1024;
const PREPARE_QUEUE_LIMIT: usize = 4;
// The template is embedded in the process. Bump this if a future deployment can swap templates
// without restarting, so bytes produced under the old template can never be reused.
const TEMPLATE_GENERATION: u32 = 1;

#[derive(Clone)]
pub struct AppConfig {
    pub password_hash: String,
    pub session_signing_secret: String,
    pub secure_cookies: bool,
    pub session_ttl: Duration,
    pub background_deck_preparation: bool,
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
            background_deck_preparation: matches!(
                std::env::var("BACKGROUND_DECK_PREPARATION").as_deref(),
                Ok("true" | "1")
            ),
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
    prepared_decks: Arc<Mutex<PreparedDeckCache>>,
    prepare_tx: SyncSender<PrepareJob>,
    background_deck_preparation: bool,
}

#[derive(Clone, Debug, Eq, Hash, PartialEq)]
struct PreparedDeckKey {
    service_id: String,
    service_revision: u64,
    settings_version: u64,
    template_generation: u32,
}

impl PreparedDeckKey {
    fn new(service: &ServiceRecord, settings: &GlobalSettingsVersion) -> Self {
        Self {
            service_id: service.id.clone(),
            service_revision: service.revision,
            settings_version: settings.version,
            template_generation: TEMPLATE_GENERATION,
        }
    }
}

struct PreparedDeckEntry {
    bytes: Arc<[u8]>,
    created_at: Instant,
}

#[derive(Default)]
struct PreparedDeckCache {
    ready: HashMap<PreparedDeckKey, PreparedDeckEntry>,
    in_flight: HashSet<PreparedDeckKey>,
    latest_for_service: HashMap<String, PreparedDeckKey>,
    total_bytes: usize,
}

impl PreparedDeckCache {
    fn prune(&mut self, now: Instant) {
        let expired = self
            .ready
            .iter()
            .filter(|(_, entry)| {
                now.saturating_duration_since(entry.created_at) > PREPARED_DECK_TTL
            })
            .map(|(key, _)| key.clone())
            .collect::<Vec<_>>();
        for key in expired {
            self.remove_ready(&key);
        }
    }

    fn get_ready(&mut self, key: &PreparedDeckKey) -> Option<Arc<[u8]>> {
        self.prune(Instant::now());
        self.ready.get(key).map(|entry| entry.bytes.clone())
    }

    fn insert_ready(&mut self, key: PreparedDeckKey, bytes: Arc<[u8]>) {
        self.prune(Instant::now());
        let older_for_service = self
            .ready
            .keys()
            .filter(|candidate| candidate.service_id == key.service_id && **candidate != key)
            .cloned()
            .collect::<Vec<_>>();
        for candidate in older_for_service {
            self.remove_ready(&candidate);
        }
        if bytes.len() > PREPARED_DECK_BYTE_LIMIT {
            return;
        }
        if let Some(previous) = self.ready.remove(&key) {
            self.total_bytes = self.total_bytes.saturating_sub(previous.bytes.len());
        }
        while self.ready.len() >= PREPARED_DECK_LIMIT
            || self.total_bytes.saturating_add(bytes.len()) > PREPARED_DECK_BYTE_LIMIT
        {
            let Some(oldest) = self
                .ready
                .iter()
                .min_by_key(|(_, entry)| entry.created_at)
                .map(|(candidate, _)| candidate.clone())
            else {
                break;
            };
            self.remove_ready(&oldest);
        }
        self.total_bytes = self.total_bytes.saturating_add(bytes.len());
        self.ready.insert(
            key,
            PreparedDeckEntry {
                bytes,
                created_at: Instant::now(),
            },
        );
    }

    fn remove_ready(&mut self, key: &PreparedDeckKey) {
        if let Some(entry) = self.ready.remove(key) {
            self.total_bytes = self.total_bytes.saturating_sub(entry.bytes.len());
        }
    }

    fn is_latest(&self, key: &PreparedDeckKey) -> bool {
        self.latest_for_service.get(&key.service_id) == Some(key)
    }

    fn finish(&mut self, key: &PreparedDeckKey) {
        self.in_flight.remove(key);
        if self.latest_for_service.get(&key.service_id) == Some(key) {
            self.latest_for_service.remove(&key.service_id);
        }
    }
}

struct PrepareJob {
    key: PreparedDeckKey,
    service: ServiceRecord,
    settings: GlobalSettingsVersion,
}

struct PreparationWorkerState {
    sources: Arc<dyn Sources>,
    store: Arc<dyn ObjectStore>,
    prepared_decks: Arc<Mutex<PreparedDeckCache>>,
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

    fn teaching(&self, source: TeachingSource, selection: &str) -> anyhow::Result<Teaching> {
        self.upstream.teaching(source, selection)
    }

    fn fixed_component(&self, key: &str) -> anyhow::Result<FixedComponent> {
        self.upstream.fixed_component(key)
    }
}

pub fn app(sources: Arc<dyn Sources>, store: Arc<dyn ObjectStore>, config: AppConfig) -> Router {
    let prepared_decks = Arc::new(Mutex::new(PreparedDeckCache::default()));
    let (prepare_tx, prepare_rx) = mpsc::sync_channel(PREPARE_QUEUE_LIMIT);
    let background_deck_preparation = config.background_deck_preparation;
    let state = AppState {
        sources,
        store,
        config,
        login_limiter: Arc::new(Mutex::new(HashMap::new())),
        next_id: Arc::new(AtomicU64::new(1)),
        prepared_decks,
        prepare_tx,
        background_deck_preparation,
    };
    let worker_state = PreparationWorkerState {
        sources: state.sources.clone(),
        store: state.store.clone(),
        prepared_decks: state.prepared_decks.clone(),
    };
    if background_deck_preparation {
        std::thread::Builder::new()
            .name("deck-preparation".to_string())
            .spawn(move || prepare_worker(worker_state, prepare_rx))
            .expect("could not start background deck preparation worker");
    }

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
        .route("/api/services/:id/prepare", post(prepare_service))
        .route("/api/services/:id/generate", post(generate_service))
        .route("/api/services/:id/history", get(service_history))
        .route("/api/generated", get(generated_decks))
        .route(
            "/api/services/:id/revisions/:revision/download",
            get(download_service_revision),
        )
        .route(
            "/api/services/:id/revisions/:revision/snapshot",
            get(service_revision_snapshot),
        )
        .route(
            "/api/services/:id/revisions/:revision/restore",
            post(restore_service_revision),
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
    let item = state
        .sources
        .teaching(query.source, &query.selection)
        .map_err(|error| AppError::bad_request(error.to_string()))?;
    Ok(Json(json!({
        "source": query.source,
        "selection": item.selection,
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
        // The same prefix also contains every generated revision. Fetching those one by one from
        // R2 only to fail ServiceRecord deserialization made builder startup slower over time.
        if !is_current_service_key(&key) {
            continue;
        }
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
    invalidate_prepared_service(&state, &id);
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
    invalidate_prepared_service(state, id);
    Ok(Json(service))
}

#[derive(Deserialize)]
struct PrepareDeckRequest {
    revision: u64,
}

#[derive(Serialize)]
struct PrepareDeckResponse {
    status: &'static str,
    revision: u64,
}

async fn prepare_service(
    State(state): State<AppState>,
    Path(id): Path<String>,
    Json(request): Json<PrepareDeckRequest>,
) -> Result<impl IntoResponse, AppError> {
    if !state.background_deck_preparation {
        return Ok((
            StatusCode::OK,
            Json(PrepareDeckResponse {
                status: "disabled",
                revision: request.revision,
            }),
        ));
    }
    let (service, _) = load_service(&state, &id).await?;
    if request.revision != service.revision {
        return Err(AppError::conflict(
            "this service changed before background preparation could begin",
        ));
    }
    let settings = load_settings(&state).await?;
    let key = PreparedDeckKey::new(&service, &settings);
    let previous_latest = {
        let mut cache = prepared_cache(&state.prepared_decks)?;
        if cache.get_ready(&key).is_some() {
            return Ok((
                StatusCode::OK,
                Json(PrepareDeckResponse {
                    status: "ready",
                    revision: request.revision,
                }),
            ));
        }
        if cache.in_flight.contains(&key) {
            return Ok((
                StatusCode::ACCEPTED,
                Json(PrepareDeckResponse {
                    status: "preparing",
                    revision: request.revision,
                }),
            ));
        }
        cache.in_flight.insert(key.clone());
        cache
            .latest_for_service
            .insert(key.service_id.clone(), key.clone())
    };

    let job = PrepareJob {
        key: key.clone(),
        service,
        settings,
    };
    if state.prepare_tx.try_send(job).is_err() {
        let mut cache = prepared_cache(&state.prepared_decks)?;
        cache.in_flight.remove(&key);
        if cache.latest_for_service.get(&key.service_id) == Some(&key) {
            match previous_latest {
                Some(previous) if cache.in_flight.contains(&previous) => {
                    cache
                        .latest_for_service
                        .insert(key.service_id.clone(), previous);
                }
                _ => {
                    cache.latest_for_service.remove(&key.service_id);
                }
            }
        }
        return Ok((
            StatusCode::ACCEPTED,
            Json(PrepareDeckResponse {
                status: "busy",
                revision: request.revision,
            }),
        ));
    }

    Ok((
        StatusCode::ACCEPTED,
        Json(PrepareDeckResponse {
            status: "preparing",
            revision: request.revision,
        }),
    ))
}

fn prepare_worker(state: PreparationWorkerState, jobs: Receiver<PrepareJob>) {
    let runtime = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .expect("could not start the deck preparation runtime");
    while let Ok(job) = jobs.recv() {
        runtime.block_on(process_prepare_job(&state, job));
    }

    // A closed worker means no queued receiver should leave a Generate request waiting forever.
    if let Ok(mut cache) = prepared_cache(&state.prepared_decks) {
        cache.in_flight.clear();
        cache.latest_for_service.clear();
    }
}

async fn process_prepare_job(state: &PreparationWorkerState, job: PrepareJob) {
    let should_build = prepared_cache(&state.prepared_decks)
        .map(|cache| cache.is_latest(&job.key))
        .unwrap_or(false);
    if !should_build {
        if let Ok(mut cache) = prepared_cache(&state.prepared_decks) {
            cache.finish(&job.key);
        }
        return;
    }

    let result =
        render_service_deck(&state.sources, &state.store, &job.service, &job.settings).await;
    let still_current = if result.is_ok() {
        preparation_is_current(&state.store, &state.prepared_decks, &job.key).await
    } else {
        false
    };

    match result {
        Ok(bytes) if still_current => {
            let bytes = Arc::<[u8]>::from(bytes);
            if let Ok(mut cache) = prepared_cache(&state.prepared_decks) {
                if cache.is_latest(&job.key) {
                    cache.insert_ready(job.key.clone(), bytes);
                }
                cache.finish(&job.key);
            }
        }
        _ => {
            if let Ok(mut cache) = prepared_cache(&state.prepared_decks) {
                cache.finish(&job.key);
            }
        }
    }
}

async fn preparation_is_current(
    store: &Arc<dyn ObjectStore>,
    prepared_decks: &Arc<Mutex<PreparedDeckCache>>,
    key: &PreparedDeckKey,
) -> bool {
    let still_latest = prepared_cache(prepared_decks)
        .map(|cache| cache.is_latest(key))
        .unwrap_or(false);
    if !still_latest {
        return false;
    }
    let Ok(service_object) = store.get(&service_key(&key.service_id)).await else {
        return false;
    };
    let Ok(service) = serde_json::from_slice::<ServiceRecord>(&service_object.bytes) else {
        return false;
    };
    let Ok(settings_object) = store.get("entities/settings/current.json").await else {
        return false;
    };
    let Ok(settings) = serde_json::from_slice::<GlobalSettingsVersion>(&settings_object.bytes)
    else {
        return false;
    };
    service.revision == key.service_revision && settings.version == key.settings_version
}

fn prepared_cache(
    prepared_decks: &Arc<Mutex<PreparedDeckCache>>,
) -> Result<std::sync::MutexGuard<'_, PreparedDeckCache>, AppError> {
    prepared_decks
        .lock()
        .map_err(|_| AppError::internal("background deck cache became unavailable"))
}

fn invalidate_prepared_service(state: &AppState, id: &str) {
    let Ok(mut cache) = prepared_cache(&state.prepared_decks) else {
        return;
    };
    cache.latest_for_service.remove(id);
    let ready = cache
        .ready
        .keys()
        .filter(|key| key.service_id == id)
        .cloned()
        .collect::<Vec<_>>();
    for key in ready {
        cache.remove_ready(&key);
    }
}

fn invalidate_all_prepared_decks(state: &AppState) {
    let Ok(mut cache) = prepared_cache(&state.prepared_decks) else {
        return;
    };
    cache.ready.clear();
    cache.latest_for_service.clear();
    cache.total_bytes = 0;
}

async fn render_service_deck(
    upstream: &Arc<dyn Sources>,
    store: &Arc<dyn ObjectStore>,
    service: &ServiceRecord,
    settings: &GlobalSettingsVersion,
) -> Result<Vec<u8>, AppError> {
    let sources = ServiceSources {
        upstream: upstream.clone(),
        store: store.clone(),
    };
    build_deck(service, &sources, &settings.ccli_licence_number)
        .await
        .map_err(|error| AppError::internal(format!("could not build service deck: {error}")))
}

fn prepared_deck_for_generation(state: &AppState, key: &PreparedDeckKey) -> Option<Arc<[u8]>> {
    let mut cache = prepared_cache(&state.prepared_decks).ok()?;
    let ready = cache.get_ready(key);
    if ready.is_none() && cache.latest_for_service.get(&key.service_id) == Some(key) {
        // Prevent a queued speculative job from starting after foreground generation begins.
        // An already-running build cannot be interrupted safely, which is why this feature is
        // disabled by default on constrained hosts.
        cache.latest_for_service.remove(&key.service_id);
    }
    ready
}

async fn generate_service(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Path(id): Path<String>,
) -> Result<Response, AppError> {
    let (mut service, object) = load_service(&state, &id).await?;
    let settings = load_settings(&state).await?;
    let cache_key = PreparedDeckKey::new(&service, &settings);
    let prepared = prepared_deck_for_generation(&state, &cache_key);
    let (bytes, preparation_status) = if let Some(prepared) = prepared {
        (prepared.as_ref().to_vec(), "hit")
    } else {
        (
            render_service_deck(&state.sources, &state.store, &service, &settings).await?,
            "miss",
        )
    };
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
        // Captured before the status change below, so the snapshot is the service as it was
        // when its slides were built.
        service: Some(service.clone()),
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
    response.headers_mut().insert(
        "x-deck-preparation",
        HeaderValue::from_static(preparation_status),
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
    /// Decks generated before the service snapshot was recorded have nothing to show, and the
    /// history page hides the control rather than offering a link that cannot work.
    snapshot_url: Option<String>,
    restore_url: Option<String>,
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
        // The name and date the deck was built with, so renaming a service later does not
        // rewrite the history of what was already handed out.
        let snapshot = metadata.service.as_ref();
        generated.push(GeneratedDeckListing {
            service_id: metadata.service_id.clone(),
            service_name: snapshot.map_or(service.name, |snapshot| snapshot.name.clone()),
            service_date: snapshot.map_or(service.date, |snapshot| snapshot.date),
            revision: metadata.revision,
            generated_at: metadata.generated_at,
            generated_by: metadata.generated_by,
            expires_at: metadata.expires_at,
            source_revision: metadata.source_revision,
            download_url: format!(
                "/api/services/{}/revisions/{}/download",
                metadata.service_id, metadata.revision
            ),
            snapshot_url: snapshot.map(|_| {
                format!(
                    "/api/services/{}/revisions/{}/snapshot",
                    metadata.service_id, metadata.revision
                )
            }),
            restore_url: snapshot.map(|_| {
                format!(
                    "/api/services/{}/revisions/{}/restore",
                    metadata.service_id, metadata.revision
                )
            }),
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

/// The order of service a generated deck was built from. The live service keeps being edited
/// after a deck is handed out, so this is the only record of what actually went on the screen.
async fn service_revision_snapshot(
    State(state): State<AppState>,
    Path((id, revision)): Path<(String, u64)>,
) -> Result<Json<ServiceRecord>, AppError> {
    let metadata = load_generated_deck(&state, &id, revision).await?;
    metadata.service.map(Json).ok_or_else(|| {
        AppError::new(
            StatusCode::NOT_FOUND,
            "this PowerPoint was generated before the service was recorded with it",
        )
    })
}

/// Start a new editable service from the immutable settings saved with a generated deck. The
/// original service and its history stay untouched, even if the live service has since changed.
async fn restore_service_revision(
    State(state): State<AppState>,
    Extension(session): Extension<StaffSession>,
    Path((id, revision)): Path<(String, u64)>,
) -> Result<Response, AppError> {
    let metadata = load_generated_deck(&state, &id, revision).await?;
    let mut service = metadata.service.ok_or_else(|| {
        AppError::new(
            StatusCode::NOT_FOUND,
            "this PowerPoint was generated before the service was recorded with it",
        )
    })?;

    service.id = new_id(&state, "svc");
    service.name = copy_service_name(&service.name);
    service.status = ServiceStatus::Draft;
    service.revision = 0;
    service.audit = AuditMetadata::new(session.display_name);
    validate_service_name(&service.name)?;

    let stored = put_json(
        state.store.as_ref(),
        &service_key(&service.id),
        &service,
        PutCondition::IfNoneMatch,
    )
    .await?;
    json_with_etag(StatusCode::CREATED, &service, &stored.etag)
}

fn copy_service_name(source: &str) -> String {
    let mut name = format!("Copy of {source}");
    while name.len() > 100 {
        name.pop();
    }
    name
}

async fn load_generated_deck(
    state: &AppState,
    id: &str,
    revision: u64,
) -> Result<GeneratedDeckVersion, AppError> {
    let object = state
        .store
        .get(&format!("entities/services/{id}/revisions/{revision}.json"))
        .await?;
    let metadata: GeneratedDeckVersion = serde_json::from_slice(&object.bytes)?;
    if metadata.service_id != id || metadata.revision != revision {
        return Err(AppError::new(StatusCode::NOT_FOUND, "record not found"));
    }
    Ok(metadata)
}

/// Decks generated before slide layout ownership was enforced are still in storage, and
/// PowerPoint refuses them outright rather than offering to repair them. Mend one on its way out
/// and write the healed deck back, so the file is fixed for good rather than on every download.
///
/// A deck that cannot be mended is served exactly as stored: the staff member gets the same
/// broken file they had before rather than an error, and nothing in storage is disturbed.
async fn heal_stored_deck(state: &AppState, object_key: &str, bytes: Vec<u8>) -> Vec<u8> {
    let Ok(mut deck) = Presentation::open_bytes(&bytes) else {
        return bytes;
    };
    if deck.validate().is_ok() {
        return bytes;
    }
    match deck.repair() {
        Ok(true) => {}
        Ok(false) => return bytes,
        Err(error) => {
            eprintln!("could not repair stored deck {object_key}: {error}");
            return bytes;
        }
    }
    if let Err(error) = deck.validate() {
        eprintln!("stored deck {object_key} is still invalid after repair: {error}");
        return bytes;
    }
    let healed = match deck.save_bytes() {
        Ok(healed) => healed,
        Err(error) => {
            eprintln!("repaired deck {object_key} could not be saved: {error}");
            return bytes;
        }
    };

    // Writing back is a convenience, not a precondition for serving the deck.
    if let Err(error) = state
        .store
        .put(
            object_key,
            healed.clone(),
            PPTX_CONTENT_TYPE,
            PutCondition::Any,
        )
        .await
    {
        eprintln!("healed deck {object_key} could not be written back: {error}");
    }
    healed
}

async fn download_service_revision(
    State(state): State<AppState>,
    Path((id, revision)): Path<(String, u64)>,
) -> Result<Response, AppError> {
    let (service, _) = load_service(&state, &id).await?;
    let metadata = load_generated_deck(&state, &id, revision).await?;
    let deck = state.store.get(&metadata.object_key).await?;
    let bytes = heal_stored_deck(&state, &metadata.object_key, deck.bytes).await;

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
    invalidate_all_prepared_decks(&state);
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

fn is_current_service_key(key: &str) -> bool {
    key.strip_prefix("entities/services/")
        .is_some_and(|filename| !filename.contains('/') && filename.ends_with(".json"))
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

#[cfg(test)]
mod tests {
    use super::{
        is_current_service_key, PreparedDeckCache, PreparedDeckEntry, PreparedDeckKey,
        PREPARED_DECK_LIMIT, PREPARED_DECK_TTL, TEMPLATE_GENERATION,
    };
    use std::sync::Arc;
    use std::time::{Duration, Instant};

    #[test]
    fn service_listing_skips_generated_revision_metadata() {
        assert!(is_current_service_key("entities/services/svc-123.json"));
        assert!(!is_current_service_key(
            "entities/services/svc-123/revisions/7.json"
        ));
        assert!(!is_current_service_key("entities/services/svc-123.pptx"));
    }

    #[test]
    fn prepared_deck_cache_is_bounded_and_prunes_expired_entries() {
        let mut cache = PreparedDeckCache::default();
        for revision in 0..=PREPARED_DECK_LIMIT as u64 {
            cache.insert_ready(
                PreparedDeckKey {
                    service_id: format!("service-{revision}"),
                    service_revision: revision,
                    settings_version: 1,
                    template_generation: TEMPLATE_GENERATION,
                },
                Arc::<[u8]>::from(vec![revision as u8; 8]),
            );
        }
        assert_eq!(cache.ready.len(), PREPARED_DECK_LIMIT);
        assert_eq!(cache.total_bytes, PREPARED_DECK_LIMIT * 8);

        let expired_key = cache.ready.keys().next().unwrap().clone();
        cache.ready.insert(
            expired_key.clone(),
            PreparedDeckEntry {
                bytes: Arc::<[u8]>::from(vec![0; 8]),
                created_at: Instant::now() - PREPARED_DECK_TTL - Duration::from_secs(1),
            },
        );
        assert!(cache.get_ready(&expired_key).is_none());
        assert_eq!(cache.ready.len(), PREPARED_DECK_LIMIT - 1);
        assert_eq!(cache.total_bytes, (PREPARED_DECK_LIMIT - 1) * 8);
    }
}
