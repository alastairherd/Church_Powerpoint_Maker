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
