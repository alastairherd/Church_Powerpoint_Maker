use axum::error_handling::HandleErrorLayer;
use axum::BoxError;
use axum::Json;
use deck_builder::LiveSources;
use http::StatusCode;
use serde_json::json;
use server::store::{MemoryObjectStore, ObjectStore, R2ObjectStore};
use server::{app, AppConfig};
use std::net::SocketAddr;
use std::sync::Arc;
use std::time::Duration;
use tower::limit::ConcurrencyLimitLayer;
use tower::timeout::TimeoutLayer;
use tower::ServiceBuilder;

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let esv_api_key = std::env::var("ESV_API_KEY")
        .map_err(|_| anyhow::anyhow!("ESV_API_KEY must be set before starting the server"))?;
    let config = AppConfig::from_env()?;
    let port = std::env::var("PORT")
        .ok()
        .and_then(|value| value.parse::<u16>().ok())
        .unwrap_or(8080);

    let store: Arc<dyn ObjectStore> = if std::env::var("OBJECT_STORE").as_deref() == Ok("memory") {
        Arc::new(MemoryObjectStore::default())
    } else {
        Arc::new(R2ObjectStore::from_env()?)
    };
    let app = app(Arc::new(LiveSources::new(esv_api_key)?), store, config).layer(
        ServiceBuilder::new()
            .layer(HandleErrorLayer::new(|err: BoxError| async move {
                (
                    StatusCode::REQUEST_TIMEOUT,
                    Json(json!({ "error": format!("request failed or timed out: {err}") })),
                )
            }))
            .layer(ConcurrencyLimitLayer::new(4))
            .layer(TimeoutLayer::new(Duration::from_secs(30))),
    );
    let addr = SocketAddr::from(([0, 0, 0, 0], port));
    let listener = tokio::net::TcpListener::bind(addr).await?;
    axum::serve(listener, app).await?;
    Ok(())
}
