use async_trait::async_trait;
use axum::body::{to_bytes, Body};
use deck_builder::{Component, Hymn, Scripture, ServiceOrder, Sources};
use http::{Request, StatusCode};
use pptx_template::Presentation;
use server::app_with_sources;
use std::sync::Arc;
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

    async fn hymn(&self, _url: &str) -> anyhow::Result<Hymn> {
        Ok(Hymn {
            title: "Test Hymn".to_string(),
            stanzas: vec!["Verse one".to_string()],
            author: "Author".to_string(),
            composer: "Composer".to_string(),
            tune: "Tune".to_string(),
            copyright: "Public Domain".to_string(),
        })
    }
}

#[tokio::test]
async fn healthz_returns_ok() {
    let app = app_with_sources(Arc::new(TestSources), None);
    let response = app
        .oneshot(
            Request::builder()
                .uri("/healthz")
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();

    assert_eq!(response.status(), StatusCode::OK);
}

#[tokio::test]
async fn generate_returns_downloadable_pptx() {
    let app = app_with_sources(Arc::new(TestSources), None);
    let order = ServiceOrder {
        date: chrono::NaiveDate::from_ymd_opt(2026, 7, 12).unwrap(),
        components: vec![Component::Scripture {
            reference: "Genesis 1:1".to_string(),
            title: Some("First Reading".to_string()),
        }],
    };
    let body = serde_json::to_vec(&order).unwrap();
    let response = app
        .oneshot(
            Request::builder()
                .method("POST")
                .uri("/api/generate")
                .header("content-type", "application/json")
                .body(Body::from(body))
                .unwrap(),
        )
        .await
        .unwrap();

    assert_eq!(response.status(), StatusCode::OK);
    assert_eq!(
        response.headers()["content-type"],
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    assert!(response.headers()["content-disposition"]
        .to_str()
        .unwrap()
        .contains("service-2026-07-12.pptx"));

    let body = to_bytes(response.into_body(), usize::MAX).await.unwrap();
    assert!(body.starts_with(b"PK"));
    Presentation::open_bytes(&body).unwrap().validate().unwrap();
}
