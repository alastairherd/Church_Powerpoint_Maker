use axum::body::Body;
use axum::extract::State;
use axum::response::{Html, IntoResponse, Response};
use axum::routing::{get, post};
use axum::{Json, Router};
use deck_builder::sources::hymnary::validate_hymnary_url;
use deck_builder::{
    build_deck, Catechism, Component, FixedComponent, Psalm, ServiceOrder, Sources,
};
use http::header::{CONTENT_DISPOSITION, CONTENT_TYPE};
use http::{HeaderMap, HeaderValue, StatusCode};
use serde_json::json;
use std::collections::HashMap;
use std::sync::Arc;
use std::sync::Mutex;
use std::time::{Duration, Instant};

const PPTX_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_COMPONENTS: usize = 30;
const MAX_FIELD_LEN: usize = 300;
const RATE_LIMIT_WINDOW: Duration = Duration::from_secs(60);
const RATE_LIMIT_REQUESTS: usize = 5;

#[derive(Clone)]
pub struct AppState {
    sources: Arc<dyn Sources>,
    sheet_csv_url: Option<String>,
    rate_limiter: Arc<Mutex<HashMap<String, Vec<Instant>>>>,
}

pub fn app_with_sources(sources: Arc<dyn Sources>, sheet_csv_url: Option<String>) -> Router {
    let state = AppState {
        sources,
        sheet_csv_url,
        rate_limiter: Arc::new(Mutex::new(HashMap::new())),
    };

    Router::new()
        .route("/", get(index))
        .route("/healthz", get(healthz))
        .route("/api/prefill", get(prefill))
        .route("/api/generate", post(generate))
        .with_state(state)
}

async fn index() -> Html<&'static str> {
    Html(INDEX_HTML)
}

async fn healthz() -> &'static str {
    "ok"
}

async fn generate(
    State(state): State<AppState>,
    headers: HeaderMap,
    Json(order): Json<ServiceOrder>,
) -> Result<Response, AppError> {
    check_rate_limit(&state, &headers)?;
    validate_order(&order)?;
    let filename = format!("service-{}.pptx", order.date);
    let bytes = build_deck(&order, state.sources.as_ref())
        .await
        .map_err(|err| AppError::internal(err.to_string()))?;

    let mut response = Response::new(Body::from(bytes));
    response
        .headers_mut()
        .insert(CONTENT_TYPE, HeaderValue::from_static(PPTX_CONTENT_TYPE));
    response.headers_mut().insert(
        CONTENT_DISPOSITION,
        HeaderValue::from_str(&format!("attachment; filename=\"{filename}\""))
            .map_err(|err| AppError::internal(err.to_string()))?,
    );
    Ok(response)
}

fn check_rate_limit(state: &AppState, headers: &HeaderMap) -> Result<(), AppError> {
    let client = headers
        .get("x-forwarded-for")
        .and_then(|value| value.to_str().ok())
        .and_then(|value| value.split(',').next())
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .unwrap_or("unknown")
        .to_string();
    let now = Instant::now();
    let mut limiter = state
        .rate_limiter
        .lock()
        .map_err(|_| AppError::internal("rate limiter unavailable"))?;
    let hits = limiter.entry(client).or_default();
    hits.retain(|hit| now.duration_since(*hit) < RATE_LIMIT_WINDOW);
    if hits.len() >= RATE_LIMIT_REQUESTS {
        return Err(AppError::new(
            StatusCode::TOO_MANY_REQUESTS,
            "too many generation requests; please wait and try again",
        ));
    }
    hits.push(now);
    Ok(())
}

async fn prefill(State(state): State<AppState>) -> Result<Json<serde_json::Value>, AppError> {
    let Some(url) = state.sheet_csv_url else {
        return Err(AppError::new(
            StatusCode::NOT_FOUND,
            "sheet prefill is not configured",
        ));
    };
    let text = reqwest::Client::builder()
        .timeout(Duration::from_secs(10))
        .build()
        .map_err(|err| AppError::bad_gateway(format!("could not create sheet client: {err}")))?
        .get(url)
        .send()
        .await
        .map_err(|err| AppError::bad_gateway(format!("could not fetch sheet: {err}")))?
        .text()
        .await
        .map_err(|err| AppError::bad_gateway(format!("could not read sheet: {err}")))?;
    let mut reader = csv::Reader::from_reader(text.as_bytes());
    let headers = reader
        .headers()
        .map_err(|err| AppError::bad_gateway(format!("could not parse sheet headers: {err}")))?
        .clone();
    let mut records = Vec::new();
    for record in reader.records() {
        let record =
            record.map_err(|err| AppError::bad_gateway(format!("could not parse sheet: {err}")))?;
        let mut row = serde_json::Map::new();
        for (header, value) in headers.iter().zip(record.iter()) {
            row.insert(header.to_string(), json!(value));
        }
        records.push(json!(row));
    }
    Ok(Json(json!({ "records": records })))
}

fn validate_order(order: &ServiceOrder) -> Result<(), AppError> {
    if order.components.len() > MAX_COMPONENTS {
        return Err(AppError::new(
            StatusCode::BAD_REQUEST,
            format!("service order is limited to {MAX_COMPONENTS} components"),
        ));
    }

    for component in &order.components {
        match component {
            Component::Psalm { reference } | Component::Scripture { reference, .. } => {
                validate_field(reference, "reference")?;
                if matches!(component, Component::Psalm { .. }) {
                    Psalm::find(reference)
                        .map_err(|err| AppError::new(StatusCode::BAD_REQUEST, err.to_string()))?;
                }
            }
            Component::Hymn { url } => {
                validate_field(url, "hymn URL")?;
                validate_hymnary_url(url)
                    .map_err(|err| AppError::new(StatusCode::BAD_REQUEST, err.to_string()))?;
            }
            Component::Catechism { question } => {
                Catechism::find(*question)
                    .map_err(|err| AppError::new(StatusCode::BAD_REQUEST, err.to_string()))?;
            }
            Component::Fixed { key, title } => {
                validate_field(key, "component key")?;
                FixedComponent::find(key)
                    .map_err(|err| AppError::new(StatusCode::BAD_REQUEST, err.to_string()))?;
                if let Some(title) = title {
                    validate_field(title, "title")?;
                }
            }
            Component::Sermon { title, text } => {
                validate_field(title, "title")?;
                validate_field(text, "sermon text")?;
            }
        }
    }
    Ok(())
}

fn validate_field(value: &str, name: &str) -> Result<(), AppError> {
    if value.trim().is_empty() {
        return Err(AppError::new(
            StatusCode::BAD_REQUEST,
            format!("{name} is required"),
        ));
    }
    if value.len() > MAX_FIELD_LEN {
        return Err(AppError::new(
            StatusCode::BAD_REQUEST,
            format!("{name} is too long"),
        ));
    }
    Ok(())
}

#[derive(Debug)]
struct AppError {
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

    fn internal(message: impl Into<String>) -> Self {
        Self::new(StatusCode::INTERNAL_SERVER_ERROR, message)
    }

    fn bad_gateway(message: impl Into<String>) -> Self {
        Self::new(StatusCode::BAD_GATEWAY, message)
    }
}

impl IntoResponse for AppError {
    fn into_response(self) -> Response {
        (self.status, Json(json!({ "error": self.message }))).into_response()
    }
}

const INDEX_HTML: &str = r#"<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Church PowerPoint Maker</title>
  <style>
    body{font-family:system-ui,-apple-system,"Segoe UI",sans-serif;max-width:920px;margin:0 auto;padding:2rem;background:#f7f1e8;color:#241f1a}
    header{margin-bottom:2rem} h1{font-family:Georgia,serif;font-size:clamp(2rem,6vw,4rem);line-height:.95;margin:0;color:#7d1f1f}
    form{display:grid;gap:1rem}.card{background:#fffaf3;border:1px solid #dfcfbb;border-radius:18px;padding:1rem;box-shadow:0 8px 28px #4d33221a}
    .row{display:grid;grid-template-columns:150px 1fr 1fr auto;gap:.5rem;margin:.5rem 0}label{font-weight:700}input,select,button{font:inherit;padding:.7rem;border:1px solid #cdbba5;border-radius:10px;background:white}
    button{background:#7d1f1f;color:white;border:0;cursor:pointer}.secondary{background:#3f3a35}.danger{background:#866}#status{min-height:1.5rem;font-weight:700}
    @media(max-width:760px){.row{grid-template-columns:1fr}.row button{width:100%}}
  </style>
</head>
<body>
  <header><h1>Church PowerPoint Maker</h1><p>Enter the service order, generate a deck, and download the finished .pptx.</p></header>
  <form id="form" class="card">
    <label>Date <input id="date" type="date" required></label>
    <div><label>Components</label><div id="components"></div></div>
    <p><button type="button" class="secondary" id="add">Add component</button> <button type="submit">Generate PowerPoint</button></p>
    <div id="status"></div>
  </form>
  <script>
    const components = document.querySelector('#components');
    const status = document.querySelector('#status');
    const types = ['scripture','psalm','hymn','catechism','fixed','sermon'];
    document.querySelector('#date').valueAsDate = new Date();
    function addRow(type='scripture', value='', title='') {
      const row = document.createElement('div'); row.className = 'row';
      row.innerHTML = `<select>${types.map(t=>`<option ${t===type?'selected':''}>${t}</option>`).join('')}</select><input placeholder="Reference / URL / key / question" value="${value}"><input placeholder="Optional title" value="${title}"><button type="button" class="danger">Remove</button>`;
      row.querySelector('button').onclick = () => row.remove(); components.append(row);
    }
    document.querySelector('#add').onclick = () => addRow(); addRow('fixed','confession','Confession'); addRow('scripture','Genesis 1:1','First Reading');
    document.querySelector('#form').onsubmit = async e => {
      e.preventDefault(); status.textContent = 'Generating...';
      const order = {date: document.querySelector('#date').value, components: [...components.children].map(row => {
        const [sel, input, title] = row.querySelectorAll('select,input'); const type = sel.value; const v = input.value.trim(); const t = title.value.trim();
        if (type === 'psalm') return {type, reference:v}; if (type === 'hymn') return {type, url:v}; if (type === 'catechism') return {type, question:Number(v)};
        if (type === 'fixed') return {type, key:v, title:t||null}; if (type === 'sermon') return {type, title:t||'Sermon', text:v}; return {type, reference:v, title:t||null};
      })};
      try { const res = await fetch('/api/generate',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify(order)}); if(!res.ok){throw new Error((await res.json()).error||res.statusText)}
        const blob = await res.blob(); const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `service-${order.date}.pptx`; a.click(); URL.revokeObjectURL(a.href); status.textContent = 'Downloaded.';
      } catch(err) { status.textContent = err.message; }
    };
  </script>
</body>
</html>"#;
