use argon2::password_hash::SaltString;
use argon2::{Argon2, PasswordHasher};
use async_trait::async_trait;
use axum::body::{to_bytes, Body};
use deck_builder::{Scripture, Sources};
use http::header::{CONTENT_DISPOSITION, CONTENT_TYPE};
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

fn test_app_with_background(background_deck_preparation: bool) -> axum::Router {
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
            background_deck_preparation,
        },
    )
}

fn test_app() -> axum::Router {
    test_app_with_background(false)
}

async fn authenticated() -> (axum::Router, String, String) {
    let app = test_app();
    let (cookie, csrf) = login(&app, "Test Staff").await;
    (app, cookie, csrf)
}

async fn authenticated_with_background() -> (axum::Router, String, String) {
    let app = test_app_with_background(true);
    let (cookie, csrf) = login(&app, "Test Staff").await;
    (app, cookie, csrf)
}

async fn login(app: &axum::Router, display_name: &str) -> (String, String) {
    let response = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/login")
                .header("content-type", "application/x-www-form-urlencoded")
                .body(Body::from(format!(
                    "display_name={}&password=correct+horse",
                    display_name.replace(' ', "+")
                )))
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
    (cookie, csrf)
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
        ("/generated", "Generated service decks"),
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
async fn teaching_loader_accepts_friendly_selections_for_every_source() {
    let (app, cookie, _) = authenticated().await;
    let cases = [
        (
            "/api/teaching?source=westminster_shorter_catechism&selection=Q.%201",
            "1",
            "What is the chief end of man?",
            "glorify God",
        ),
        (
            "/api/teaching?source=heidelberg1891&selection=Q1",
            "1",
            "What is your only comfort in life and death?",
            "faithful Saviour",
        ),
        (
            "/api/teaching?source=westminster_confession_original_british&selection=1.2",
            "1.2",
            "Chapter 1: Of the Holy Scripture",
            "Word of God written",
        ),
    ];
    for (uri, selection, question, answer_fragment) in cases {
        let teaching = app
            .clone()
            .oneshot(
                Request::builder()
                    .uri(uri)
                    .header("cookie", &cookie)
                    .body(Body::empty())
                    .unwrap(),
            )
            .await
            .unwrap();
        assert_eq!(teaching.status(), StatusCode::OK, "{uri}");
        let body = to_bytes(teaching.into_body(), usize::MAX).await.unwrap();
        let teaching: serde_json::Value = serde_json::from_slice(&body).unwrap();
        assert_eq!(teaching["selection"], selection, "{uri}");
        assert_eq!(teaching["question"], question, "{uri}");
        assert!(
            teaching["answer"]
                .as_str()
                .unwrap()
                .contains(answer_fragment),
            "{uri}"
        );
    }

    let invalid = app
        .oneshot(
            Request::builder()
                .uri("/api/teaching?source=westminster_confession_original_british&selection=nonsense")
                .header("cookie", cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(invalid.status(), StatusCode::BAD_REQUEST);
    let body = to_bytes(invalid.into_body(), usize::MAX).await.unwrap();
    let error: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert!(error["error"]
        .as_str()
        .unwrap()
        .contains("confession chapter"));
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

    service["revision"] = serde_json::json!(999);

    let response = app
        .oneshot(
            Request::builder()
                .method("PUT")
                .uri(format!("/api/services/{id}/autosave"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
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
async fn generates_an_immutable_revision_without_locking() {
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

    let generated = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/generate"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
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
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/services/{id}/history"))
                .header("cookie", &cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(history.into_body(), usize::MAX).await.unwrap();
    let revisions: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(revisions.as_array().unwrap().len(), 1);
    assert_eq!(revisions[0]["revision"], 1);

    let generated = app
        .clone()
        .oneshot(
            Request::builder()
                .uri("/api/generated")
                .header("cookie", &cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(generated.status(), StatusCode::OK);
    let body = to_bytes(generated.into_body(), usize::MAX).await.unwrap();
    let generated: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(generated.as_array().unwrap().len(), 1);
    assert_eq!(generated[0]["service_name"], "Morning service");
    assert_eq!(generated[0]["revision"], 1);
    assert_eq!(
        generated[0]["download_url"],
        format!("/api/services/{id}/revisions/1/download")
    );

    let downloaded = app
        .oneshot(
            Request::builder()
                .uri(format!("/api/services/{id}/revisions/1/download"))
                .header("cookie", cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(downloaded.status(), StatusCode::OK);
    assert_eq!(
        downloaded
            .headers()
            .get(CONTENT_TYPE)
            .unwrap()
            .to_str()
            .unwrap(),
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    assert_eq!(
        downloaded
            .headers()
            .get(CONTENT_DISPOSITION)
            .unwrap()
            .to_str()
            .unwrap(),
        "attachment; filename=\"Morning-service-2026-07-12-r1.pptx\""
    );
    let body = to_bytes(downloaded.into_body(), usize::MAX).await.unwrap();
    assert!(body.starts_with(b"PK"));
    Presentation::open_bytes(&body).unwrap().validate().unwrap();
}

#[tokio::test]
async fn full_deck_background_preparation_is_disabled_by_default() {
    let (app, cookie, csrf) = authenticated().await;
    let id = create_service(&app, &cookie, &csrf, "Foreground generation").await;
    let service = get_json(&app, &cookie, &format!("/api/services/{id}")).await;
    let revision = service["revision"].as_u64().unwrap();

    let response = app
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/prepare"))
                .header("cookie", cookie)
                .header("x-csrf-token", csrf)
                .header("content-type", "application/json")
                .body(Body::from(format!(r#"{{"revision":{revision}}}"#)))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    let body = to_bytes(response.into_body(), usize::MAX).await.unwrap();
    let body: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(body["status"], "disabled");
}

#[tokio::test]
async fn prepares_an_exact_revision_without_creating_history_and_generation_reuses_it() {
    let (app, cookie, csrf) = authenticated_with_background().await;
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
                    r#"{"name":"Prepared service","date":"2026-07-12","preset":"am"}"#,
                ))
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(created.into_body(), usize::MAX).await.unwrap();
    let service: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let id = service["id"].as_str().unwrap();
    let revision = service["revision"].as_u64().unwrap();

    let stale = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/prepare"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(format!(r#"{{"revision":{}}}"#, revision + 1)))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(stale.status(), StatusCode::CONFLICT);

    let prepared = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/prepare"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(format!(r#"{{"revision":{revision}}}"#)))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(prepared.status(), StatusCode::ACCEPTED);

    let mut became_ready = false;
    for _ in 0..400 {
        tokio::time::sleep(Duration::from_millis(50)).await;
        let status = app
            .clone()
            .oneshot(
                Request::builder()
                    .method("POST")
                    .uri(format!("/api/services/{id}/prepare"))
                    .header("cookie", &cookie)
                    .header("x-csrf-token", &csrf)
                    .header("content-type", "application/json")
                    .body(Body::from(format!(r#"{{"revision":{revision}}}"#)))
                    .unwrap(),
            )
            .await
            .unwrap()
            .status();
        if status == StatusCode::OK {
            became_ready = true;
            break;
        }
        assert_eq!(status, StatusCode::ACCEPTED);
    }
    assert!(became_ready, "background deck preparation did not finish");

    let history = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/services/{id}/history"))
                .header("cookie", &cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let history = to_bytes(history.into_body(), usize::MAX).await.unwrap();
    let history: serde_json::Value = serde_json::from_slice(&history).unwrap();
    assert_eq!(history, serde_json::json!([]));

    let generated = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/generate"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(generated.status(), StatusCode::OK);
    assert_eq!(generated.headers()["x-deck-preparation"], "hit");
    let body = to_bytes(generated.into_body(), usize::MAX).await.unwrap();
    Presentation::open_bytes(&body).unwrap().validate().unwrap();
}

/// The live service keeps being edited after a deck is handed out, so the history has to keep
/// its own copy of what each PowerPoint was built from.
#[tokio::test]
async fn generated_history_records_the_service_each_deck_was_built_from() {
    let (app, cookie, csrf) = authenticated().await;
    let id = create_service(&app, &cookie, &csrf, "Morning service").await;
    generate(&app, &cookie, &csrf, &id).await;

    let listing = get_json(&app, &cookie, "/api/generated").await;
    let snapshot_url = listing[0]["snapshot_url"].as_str().expect("snapshot url");
    assert_eq!(
        snapshot_url,
        format!("/api/services/{id}/revisions/1/snapshot")
    );
    assert_eq!(
        listing[0]["restore_url"].as_str().expect("restore url"),
        format!("/api/services/{id}/revisions/1/restore")
    );

    let snapshot = get_json(&app, &cookie, snapshot_url).await;
    assert_eq!(snapshot["id"], id.as_str());
    assert_eq!(snapshot["name"], "Morning service");
    assert_eq!(snapshot["preset"], "am");
    let components = snapshot["components"].as_array().expect("components");
    assert!(
        !components.is_empty(),
        "the snapshot keeps the order of service, not just its name"
    );
    assert!(components
        .iter()
        .all(|component| component["type"].is_string()));

    // Renaming the service afterwards must not rewrite what the history says was generated.
    let mut service = get_json(&app, &cookie, &format!("/api/services/{id}")).await;
    service["name"] = serde_json::json!("Renamed service");
    let renamed = app
        .clone()
        .oneshot(
            Request::builder()
                .method("PUT")
                .uri(format!("/api/services/{id}"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(serde_json::to_vec(&service).unwrap()))
                .unwrap(),
        )
        .await
        .unwrap();
    assert!(
        renamed.status().is_success(),
        "rename failed: {}",
        renamed.status()
    );

    let listing = get_json(&app, &cookie, "/api/generated").await;
    assert_eq!(
        listing[0]["service_name"], "Morning service",
        "history shows the name the deck was generated under"
    );
}

/// Reusing a generated deck must fork the saved settings into a new draft. It must never turn
/// the historical snapshot back into the live service or overwrite later edits to that service.
#[tokio::test]
async fn generated_service_settings_can_be_restored_as_a_new_draft() {
    let (app, cookie, csrf) = authenticated().await;
    let id = create_service(&app, &cookie, &csrf, "Morning service").await;
    generate(&app, &cookie, &csrf, &id).await;

    let snapshot_url = format!("/api/services/{id}/revisions/1/snapshot");
    let snapshot = get_json(&app, &cookie, &snapshot_url).await;

    // Prove restoration uses the immutable generated revision, not the live service that may
    // since have been renamed and rearranged.
    let mut live = get_json(&app, &cookie, &format!("/api/services/{id}")).await;
    live["name"] = serde_json::json!("Renamed live service");
    live["components"].as_array_mut().unwrap().reverse();
    let updated = app
        .clone()
        .oneshot(
            Request::builder()
                .method("PUT")
                .uri(format!("/api/services/{id}"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(serde_json::to_vec(&live).unwrap()))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(updated.status(), StatusCode::OK);

    let restored_response = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/revisions/1/restore"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(restored_response.status(), StatusCode::CREATED);
    let body = to_bytes(restored_response.into_body(), usize::MAX)
        .await
        .unwrap();
    let restored: serde_json::Value = serde_json::from_slice(&body).unwrap();

    let restored_id = restored["id"].as_str().expect("new service id");
    assert_ne!(restored_id, id);
    assert_eq!(restored["name"], "Copy of Morning service");
    assert_eq!(restored["date"], snapshot["date"]);
    assert_eq!(restored["preset"], snapshot["preset"]);
    assert_eq!(restored["period"], snapshot["period"]);
    assert_eq!(restored["pm_style"], snapshot["pm_style"]);
    assert_eq!(restored["lords_supper"], snapshot["lords_supper"]);
    assert_eq!(restored["components"], snapshot["components"]);
    assert_eq!(restored["status"], "draft");
    assert_eq!(restored["revision"], 0);
    assert_eq!(restored["audit"]["created_by"], "Test Staff");

    let source = get_json(&app, &cookie, &format!("/api/services/{id}")).await;
    assert_eq!(source["name"], "Renamed live service");
    assert_ne!(source["components"], restored["components"]);
    assert_eq!(
        get_json(&app, &cookie, &snapshot_url).await["components"],
        snapshot["components"],
        "restoring must not rewrite the generated revision"
    );
    assert_eq!(
        get_json(&app, &cookie, &format!("/api/services/{restored_id}")).await["components"],
        snapshot["components"],
        "the restored draft is persisted and opens like any other service"
    );
}

/// Downloading runs every stored deck past validation so broken ones can be mended. A deck that
/// was fine to begin with must come back exactly as it was stored, not silently rewritten.
/// Repairing a genuinely broken package is covered in `pptx-template`.
#[tokio::test]
async fn downloading_a_sound_deck_returns_it_untouched() {
    let (app, cookie, csrf) = authenticated().await;
    let id = create_service(&app, &cookie, &csrf, "Morning service").await;
    generate(&app, &cookie, &csrf, &id).await;

    let first = download(&app, &cookie, &id).await;
    Presentation::open_bytes(&first)
        .unwrap()
        .validate()
        .unwrap();
    let second = download(&app, &cookie, &id).await;
    assert_eq!(
        first, second,
        "a sound deck must not be rewritten by being downloaded"
    );
}

async fn create_service(app: &axum::Router, cookie: &str, csrf: &str, name: &str) -> String {
    let created = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/api/services")
                .header("cookie", cookie)
                .header("x-csrf-token", csrf)
                .header("content-type", "application/json")
                .body(Body::from(format!(
                    r#"{{"name":"{name}","date":"2026-07-12","preset":"am"}}"#
                )))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(created.status(), StatusCode::CREATED);
    let body = to_bytes(created.into_body(), usize::MAX).await.unwrap();
    let service: serde_json::Value = serde_json::from_slice(&body).unwrap();
    service["id"].as_str().unwrap().to_string()
}

async fn generate(app: &axum::Router, cookie: &str, csrf: &str, id: &str) {
    let generated = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/services/{id}/generate"))
                .header("cookie", cookie)
                .header("x-csrf-token", csrf)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(generated.status(), StatusCode::OK);
}

async fn download(app: &axum::Router, cookie: &str, id: &str) -> Vec<u8> {
    let downloaded = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/services/{id}/revisions/1/download"))
                .header("cookie", cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(downloaded.status(), StatusCode::OK);
    to_bytes(downloaded.into_body(), usize::MAX)
        .await
        .unwrap()
        .to_vec()
}

async fn get_json(app: &axum::Router, cookie: &str, uri: &str) -> serde_json::Value {
    let response = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(uri)
                .header("cookie", cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK, "GET {uri}");
    let body = to_bytes(response.into_body(), usize::MAX).await.unwrap();
    serde_json::from_slice(&body).unwrap()
}

const SONG_POWERPOINT: &[u8] = include_bytes!("../../deck-builder/assets/template.pptx");

#[tokio::test]
async fn adding_a_song_by_uploading_a_powerpoint_stores_its_slides() {
    let (app, cookie, csrf) = authenticated().await;

    let created = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/api/songs")
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(
                    r#"{"title":"And Can it Be","aliases":["And can it be that I should gain"],"variant_label":"","author_owner":"Charles Wesley","rights_status":"public_domain","ccli_song_number":"522221","lyric_slides":[],"credits":""}"#,
                ))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(created.status(), StatusCode::OK);
    let body = to_bytes(created.into_body(), usize::MAX).await.unwrap();
    let song: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let song_id = song["id"].as_str().unwrap().to_string();
    // No PowerPoint yet, so no version has been written.
    assert_eq!(song["current_version"], 0);
    assert_eq!(song["slide_count"], 0);

    let empty_preview = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/songs/{song_id}/preview"))
                .header("cookie", &cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(empty_preview.status(), StatusCode::OK);
    let body = to_bytes(empty_preview.into_body(), usize::MAX)
        .await
        .unwrap();
    let preview: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(preview["slides"].as_array().unwrap().len(), 0);

    let uploaded = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/songs/{song_id}/upload"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("x-source-filename", "And Can it Be.pptx")
                .body(Body::from(SONG_POWERPOINT))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(uploaded.status(), StatusCode::OK);
    let body = to_bytes(uploaded.into_body(), usize::MAX).await.unwrap();
    let song: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(song["current_version"], 1);
    assert_eq!(song["source_filename"], "And Can it Be.pptx");
    let slide_count = song["slide_count"].as_u64().unwrap();
    assert!(slide_count > 0);

    let preview = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/songs/{song_id}/preview"))
                .header("cookie", &cookie)
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(preview.into_body(), usize::MAX).await.unwrap();
    let preview: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let slides = preview["slides"].as_array().unwrap();
    assert_eq!(slides.len() as u64, slide_count);
    assert!(slides
        .iter()
        .any(|slide| !slide.as_str().unwrap_or_default().trim().is_empty()));

    // A second upload becomes the next version rather than overwriting the first.
    let uploaded_again = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/songs/{song_id}/upload"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .body(Body::from(SONG_POWERPOINT))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(uploaded_again.status(), StatusCode::OK);
    let body = to_bytes(uploaded_again.into_body(), usize::MAX)
        .await
        .unwrap();
    let song: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(song["current_version"], 2);
}

#[tokio::test]
async fn uploading_something_that_is_not_a_powerpoint_is_rejected() {
    let (app, cookie, csrf) = authenticated().await;
    let created = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/api/songs")
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .header("content-type", "application/json")
                .body(Body::from(
                    r#"{"title":"Not a song","rights_status":"unknown","lyric_slides":[]}"#,
                ))
                .unwrap(),
        )
        .await
        .unwrap();
    let body = to_bytes(created.into_body(), usize::MAX).await.unwrap();
    let song: serde_json::Value = serde_json::from_slice(&body).unwrap();
    let song_id = song["id"].as_str().unwrap().to_string();

    let rejected = app
        .clone()
        .oneshot(
            Request::builder()
                .method("POST")
                .uri(format!("/api/songs/{song_id}/upload"))
                .header("cookie", &cookie)
                .header("x-csrf-token", &csrf)
                .body(Body::from("this is a plain text file, not a deck"))
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(rejected.status(), StatusCode::BAD_REQUEST);
    let body = to_bytes(rejected.into_body(), usize::MAX).await.unwrap();
    let error: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert!(error["error"]
        .as_str()
        .unwrap()
        .contains("PowerPoint was rejected"));
}
