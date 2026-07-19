use async_trait::async_trait;
use chrono::NaiveDate;
use deck_builder::{build_deck, Scripture, ServicePreset, ServiceRecord, Sources};

struct DemoSources;

#[async_trait]
impl Sources for DemoSources {
    async fn scripture(&self, reference: &str) -> anyhow::Result<Scripture> {
        Ok(Scripture {
            reference: reference.to_string(),
            text: "[1] The Saviour showed favour to his neighbour.".to_string(),
        })
    }
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let service = ServiceRecord::new(
        "demo",
        "Morning service",
        NaiveDate::from_ymd_opt(2026, 7, 12).expect("valid date"),
        ServicePreset::Am,
        "Demo user",
    );
    let bytes = build_deck(&service, &DemoSources, "522221").await?;
    std::fs::write("out.pptx", bytes)?;
    println!("Wrote out.pptx");
    Ok(())
}
