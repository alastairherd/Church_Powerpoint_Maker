use argon2::password_hash::SaltString;
use argon2::{Argon2, PasswordHasher};
use async_trait::async_trait;
use axum::body::{to_bytes, Body};
use deck_builder::{Scripture, Sources};
use http::{Request, StatusCode};
use pptx_template::Presentation;
use server::{app_with_sources, AppConfig};
use std::sync::Arc;
use std::time::Duration;
use tower::ServiceExt;

struct TestSources;

#[async_trait]
impl Sources for TestSources {
    async fn scripture(&self, reference: &str) -> anyhow::Result<Scripture> {
        Ok(Scripture {
            reference: reference.to_string(),
            text: "[1] Test scripture".to_string(),
        })
    }
}

fn test_app() -> axum::Router {
    let salt = SaltString::encode_b64(b"twpc-test-salt").unwrap();
    let hash = Argon2::default()
        .hash_password(b"correct horse", &salt)
        .unwrap()
        .to_string();
    app_with_sources(
        Arc::new(TestSources),
        AppConfig {
            password_hash: hash,
            session_signing_secret: "test-session-signing-secret-at-least-32-bytes".into(),
            secure_cookies: false,
            session_ttl: Duration::from_secs(3600),
        },
    )
}

async fn authenticated() -> (axum::Router, String, String) {
    let app = test_app();
    let response = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/login")
                .header("content-type", "application/x-www-form-urlencoded")
                .body(Body::from("display_name=Test+Staff&password=correct+horse"))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::SEE_OTHER);
    let cookie = response.headers()["set-cookie"]
        .to_str()
        .unwrap()
        .split(';')
        .next()
        .unwrap()
        .to_string();
    let session = app
        .clone()
        .oneshot(
            Request::builder()
                .uri("/api/session")
                .header("cookie", &cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(session.into_body(), usize::MAX).await.unwrap();
    let csrf = serde_json::from_slice::<serde_json::Value>(&body).unwrap()["csrf"]
        .as_str()
        .unwrap()
        .to_string();
    (app, cookie, csrf)
}

#[tokio::test]
async fn health_and_login_are_public_but_builder_is_protected() {
    let app = test_app();
    let health = app
        .clone()
        .oneshot(
            Request::builder()
                .uri("/healthz")
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(health.status(), StatusCode::OK);

    let builder = app
        .oneshot(
            Request::builder()
                .uri("/")
                .header("accept", "application/json")
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(builder.status(), StatusCode::UNAUTHORIZED);
}

#[tokio::test]
async fn mutating_requests_require_csrf() {
    let (app, cookie, _) = authenticated().await;
    let response = app
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/api/services")
                .header("cookie", cookie)
                .header("content-type", "application/json")
                .body(Body::from(
                    r#"{"name":"Morning service","date":"2026-07-12","preset":"am"}"#,
                ))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::FORBIDDEN);
}

#[tokio::test]
async fn authenticated_navigation_renders_distinct_workspaces() {
    let (app, cookie, _) = authenticated().await;
    for (path, marker) in [
        ("/", "Service order"),
        ("/library", "Choose and review stored songs"),
        ("/admin", "Staff settings"),
    ] {
        let response = app
            .clone()
            .oneshot(
                Request::builder()
                    .uri(path)
                    .header("cookie", &cookie)
                    .header("accept", "text/html")
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(response.status(), StatusCode::OK, "{path}");
        let body = to_bytes(response.into_body(), usize::MAX).await.unwrap();
        assert!(String::from_utf8_lossy(&body).contains(marker), "{path}");
    }

    let psalm = app
        .oneshot(
            Request::builder()
                .uri("/api/psalm?reference=Psalm%2023%3A1%E2%80%936")
                .header("cookie", cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(psalm.status(), StatusCode::OK);
    let body = to_bytes(psalm.into_body(), usize::MAX).await.unwrap();
    let psalm: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(psalm["meter"], "11 11 11");
    assert!(!psalm["slides"].as_array().unwrap().is_empty());
}

#[tokio::test]
async fn serves_the_editor_controller_module() {
    let (app, cookie, _) = authenticated().await;
    let response = app
        .oneshot(
            Request::builder()
                .uri("/static/editor-controller.js")
                .header("cookie", cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    let body = to_bytes(response.into_body(), usize::MAX).await.unwrap();
    assert!(String::from_utf8_lossy(&body).contains("createEditorController"));
}

#[tokio::test]
async fn scripture_and_psalm_shapes_match_editor_loaders() {
    let (app, cookie, _) = authenticated().await;
    let scripture = app
        .clone()
        .oneshot(
            Request::builder()
                .uri("/api/scripture?reference=Psalm%2096%3A2")
                .header("cookie", &cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(scripture.status(), StatusCode::OK);
    let body = to_bytes(scripture.into_body(), usize::MAX).await.unwrap();
    let scripture: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(scripture["ok"], true);
    assert_eq!(scripture["reference"], "Psalm 96:2");
    assert!(scripture["text"].is_string());

    let psalm = app
        .oneshot(
            Request::builder()
                .uri("/api/psalm?reference=Psalm%2023%3A1%E2%80%936")
                .header("cookie", cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(psalm.status(), StatusCode::OK);
    let body = to_bytes(psalm.into_body(), usize::MAX).await.unwrap();
    let psalm: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert!(psalm["reference"].is_string());
    assert!(psalm["meter"].is_string());
    assert!(psalm["slides"]
        .as_array()
        .unwrap()
        .iter()
        .all(serde_json::Value::is_string));
}

#[tokio::test]
async fn stale_autosave_returns_a_conflict_error_shape() {
    let (app, cookie, csrf) = authenticated().await;
    let created = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/api/services")
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(
                    r#"{"name":"Stale test","date":"2026-07-19","preset":"am"}"#,
                ))
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(created.into_body(), usize::MAX).await.unwrap();
    let mut service: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let id = service["id"].as_str().unwrap().to_string();

    let locked = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/lock"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(locked.into_body(), usize::MAX).await.unwrap();
    let lease: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let token = lease["token"].as_str().unwrap();
    service["revision"] = serde_json::json!(999);

    let response = app
        .oneshot(
            Request::builder()
                .method("PUT")
                .uri(format!("/api/services/{id}/autosave"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("x-lease-token", token)
                .header("content-type", "application/json")
                .body(Body::from(serde_json::to_vec(&service).unwrap()))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::CONFLICT);
    let body = to_bytes(response.into_body(), usize::MAX).await.unwrap();
    let body: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert!(body["error"].is_string());
}

#[tokio::test]
async fn song_catalogue_selection_resolves_during_generation() {
    let (app, cookie, csrf) = authenticated().await;
    let created_song = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/api/songs")
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(
                    r#"{"title":"Test Hymn","aliases":["A test song"],"variant_label":"Test version","author_owner":"Test Author","rights_status":"public_domain","ccli_song_number":null,"lyric_slides":["First verse","Second verse"],"credits":"Test Author"}"#,
                ))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(created_song.status(), StatusCode::OK);
    let body = to_bytes(created_song.into_body(), usize::MAX)
        .await
        .unwrap();
    let song: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let song_id = song["id"].as_str().unwrap().to_string();

    let search = app
        .clone()
        .oneshot(
            Request::builder()
                .uri("/api/songs?q=test%20song")
                .header("cookie", &cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(search.into_body(), usize::MAX).await.unwrap();
    let matches: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(matches.as_array().unwrap().len(), 1);

    let created_service = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/api/services")
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(
                    r#"{"name":"Song test service","date":"2026-07-19","preset":"am"}"#,
                ))
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(created_service.into_body(), usize::MAX)
        .await
        .unwrap();
    let mut service: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let service_id = service["id"].as_str().unwrap().to_string();

    let locked = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{service_id}/lock"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(locked.into_body(), usize::MAX).await.unwrap();
    let lease: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let token = lease["token"].as_str().unwrap();

    let component = service["components"]
        .as_array_mut()
        .unwrap()
        .iter_mut()
        .find(|component| component["type"] == "song")
        .unwrap();
    component["title"] = serde_json::json!("Test Hymn");
    component["song"] = serde_json::json!({
        "entity_id": song_id,
        "version": 1,
        "slide_count": 2
    });
    component["lyric_slides"] = serde_json::json!([]);

    let updated = app
        .clone()
        .oneshot(
            Request::builder()
                .method("PUT")
                .uri(format!("/api/services/{service_id}"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("x-lease-token", token)
                .header("content-type", "application/json")
                .body(Body::from(serde_json::to_vec(&service).unwrap()))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(updated.status(), StatusCode::OK);

    let generated = app
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{service_id}/generate"))
                .header("cookie", cookie)
                .header("x-csrf-token", csrf)
                .header("x-lease-token", token)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(generated.status(), StatusCode::OK);
    let body = to_bytes(generated.into_body(), usize::MAX).await.unwrap();
    Presentation::open_bytes(&body).unwrap().validate().unwrap();
}

#[tokio::test]
async fn administration_versions_the_ccli_setting() {
    let (app, cookie, csrf) = authenticated().await;
    let updated = app
        .clone()
        .oneshot(
            Request::builder()
                .method("PUT")
                .uri("/api/settings")
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(r#"{"ccli_licence_number":"654321"}"#))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(updated.status(), StatusCode::OK);
    let body = to_bytes(updated.into_body(), usize::MAX).await.unwrap();
    let settings: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(settings["version"], 2);
    assert_eq!(settings["ccli_licence_number"], "654321");

    let current = app
        .oneshot(
            Request::builder()
                .uri("/api/settings")
                .header("cookie", cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(current.into_body(), usize::MAX).await.unwrap();
    let settings: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(settings["version"], 2);
    assert_eq!(settings["ccli_licence_number"], "654321");
}

#[tokio::test]
async fn creates_locks_and_generates_an_immutable_revision() {
    let (app, cookie, csrf) = authenticated().await;
    let created = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/api/services")
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(
                    r#"{"name":"Morning service","date":"2026-07-12","preset":"am"}"#,
                ))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(created.status(), StatusCode::CREATED);
    let body = to_bytes(created.into_body(), usize::MAX).await.unwrap();
    let service: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let id = service["id"].as_str().unwrap();

    let locked = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/lock"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(locked.into_body(), usize::MAX).await.unwrap();
    let lease: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let token = lease["token"].as_str().unwrap();

    let generated = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/generate"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("x-lease-token", token)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(generated.status(), StatusCode::OK);
    let body = to_bytes(generated.into_body(), usize::MAX).await.unwrap();
    assert!(body.starts_with(b"PK"));
    Presentation::open_bytes(&body).unwrap().validate().unwrap();

    let history = app
        .oneshot(
            Request::builder()
                .uri(format!("/api/services/{id}/history"))
                .header("cookie", cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(history.into_body(), usize::MAX).await.unwrap();
    let revisions: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(revisions.as_array().unwrap().len(), 1);
    assert_eq!(revisions[0]["revision"], 1);
}
