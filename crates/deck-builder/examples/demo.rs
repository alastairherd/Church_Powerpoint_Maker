use async_trait::async_trait;
use chrono::NaiveDate;
use deck_builder::{build_deck, Component, Hymn, Scripture, ServiceOrder, Sources};

struct DemoSources;

#[async_trait]
impl Sources for DemoSources {
    async fn scripture(&self, reference: &str) -> anyhow::Result<Scripture> {
        Ok(Scripture {
            reference: reference.to_string(),
            text: "[1] The Savior showed favor to his neighbor. [2] His judgment is true."
                .to_string(),
        })
    }

    async fn hymn(&self, _url: &str) -> anyhow::Result<Hymn> {
        Ok(Hymn {
            title: "Amazing Grace".to_string(),
            stanzas: vec![
                "Amazing grace! how sweet the sound\nThat saved a wretch like me".to_string(),
                "Through many dangers, toils and snares\nI have already come".to_string(),
            ],
            author: "John Newton".to_string(),
            composer: "Unknown".to_string(),
            tune: "NEW BRITAIN".to_string(),
            copyright: "Public Domain".to_string(),
        })
    }
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let order = ServiceOrder {
        date: NaiveDate::from_ymd_opt(2026, 7, 12).expect("valid date"),
        components: vec![
            Component::Fixed {
                key: "confession".to_string(),
                title: Some("Confession".to_string()),
            },
            Component::Psalm {
                reference: "Psalm 1:1-3 (a)".to_string(),
            },
            Component::Scripture {
                reference: "Genesis 1:1-2".to_string(),
                title: Some("First Reading".to_string()),
            },
            Component::Catechism { question: 1 },
            Component::Hymn {
                url: "https://hymnary.org/text/amazing_grace_how_sweet_the_sound".to_string(),
            },
        ],
    };

    let bytes = build_deck(&order, &DemoSources).await?;
    std::fs::write("out.pptx", bytes)?;
    println!("Wrote out.pptx");
    Ok(())
}
