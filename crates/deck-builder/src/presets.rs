use crate::domain::{ServiceComponent, ServicePreset, TeachingSource};

impl ServicePreset {
    pub fn components(self) -> Vec<ServiceComponent> {
        let names: Vec<ComponentSeed> = match self {
            Self::Am => am(),
            Self::TraditionalPm => traditional_pm(),
            Self::PraiseAndWorshipPm => praise_and_worship_pm(),
            Self::AmLordsSupper => am_lords_supper(),
            Self::PmLordsSupper => pm_lords_supper(),
        };
        names
            .into_iter()
            .enumerate()
            .map(|(index, seed)| seed.component(format!("component-{}", index + 1)))
            .collect()
    }
}

enum ComponentSeed {
    Welcome,
    Notices,
    Call,
    Prayer(&'static str),
    Song,
    Psalm,
    Reading(&'static str),
    Teaching,
    Liturgy(&'static str, &'static str),
}

impl ComponentSeed {
    fn component(self, id: String) -> ServiceComponent {
        match self {
            Self::Welcome => ServiceComponent::Welcome {
                id,
                heading: "Welcome".into(),
            },
            Self::Notices => ServiceComponent::Notices {
                id,
                heading: "Notices".into(),
                rows: Vec::new(),
            },
            Self::Call => ServiceComponent::CallToWorship {
                id,
                heading: "Call to Worship".into(),
                reference: String::new(),
                text: String::new(),
                external_source_failed: false,
            },
            Self::Prayer(heading) => ServiceComponent::CuePrayer {
                id,
                heading: heading.into(),
                cue: String::new(),
                text: String::new(),
            },
            Self::Song => ServiceComponent::Song {
                id,
                title: "Choose a song".into(),
                song: None,
                lyric_slides: Vec::new(),
                credits: String::new(),
            },
            Self::Psalm => ServiceComponent::Psalm {
                id,
                heading: "Psalm".into(),
                reference: String::new(),
                show_verse_numbers: true,
                tune: None,
                slide_breaks: Vec::new(),
            },
            Self::Reading(heading) => ServiceComponent::Reading {
                id,
                heading: heading.into(),
                reference: String::new(),
                bible_page: None,
            },
            Self::Teaching => ServiceComponent::Teaching {
                id,
                heading: "Teaching".into(),
                source: TeachingSource::WestminsterShorterCatechism,
                selection: String::new(),
                text: String::new(),
            },
            Self::Liturgy(heading, key) => ServiceComponent::LiturgyBlock {
                id,
                heading: heading.into(),
                key: key.into(),
                version: None,
                text: String::new(),
            },
        }
    }
}

fn am() -> Vec<ComponentSeed> {
    use ComponentSeed::*;
    vec![
        Welcome,
        Notices,
        Call,
        Song,
        Prayer("Prayer of Praise and Adoration"),
        Liturgy("Confession", "confession"),
        Liturgy("Assurance of Forgiveness", "assurance"),
        Liturgy("Lord's Prayer", "lords_prayer"),
        Psalm,
        Reading("New Testament Reading"),
        Teaching,
        Song,
        Reading("Reading and Sermon"),
        Liturgy("Apostles' Creed", "apostles_creed"),
        Prayer("Prayers of Intercession"),
        Song,
        Liturgy("The Grace", "grace"),
        Prayer("Join us for refreshments"),
    ]
}

fn traditional_pm() -> Vec<ComponentSeed> {
    use ComponentSeed::*;
    vec![
        Welcome,
        Notices,
        Call,
        Song,
        Prayer("Prayer of Adoration"),
        Liturgy("Confession", "confession"),
        Liturgy("Assurance of Forgiveness", "assurance"),
        Liturgy("Lord's Prayer", "lords_prayer"),
        Psalm,
        Teaching,
        Song,
        Reading("Reading and Sermon"),
        Liturgy("Apostles' Creed", "apostles_creed"),
        Prayer("Prayers of Intercession"),
        Song,
        Liturgy("The Grace", "grace"),
        Prayer("Join us for refreshments"),
    ]
}

fn praise_and_worship_pm() -> Vec<ComponentSeed> {
    use ComponentSeed::*;
    vec![
        Welcome,
        Notices,
        Prayer("Opening Prayer"),
        Psalm,
        Liturgy("Confession", "confession"),
        Liturgy("Assurance of Forgiveness", "assurance"),
        Liturgy("Lord's Prayer", "lords_prayer"),
        Reading("Leader and Reading"),
        Song,
        Prayer("Time of Prayer"),
        Song,
        Song,
        Reading("Leader and Reading"),
        Prayer("Time of Prayer"),
        Song,
        Prayer("Closing Prayer"),
        Liturgy("The Grace", "grace"),
        Prayer("Join us for refreshments"),
    ]
}

fn am_lords_supper() -> Vec<ComponentSeed> {
    use ComponentSeed::*;
    vec![
        Welcome,
        Notices,
        Call,
        Song,
        Prayer("Prayer of Praise and Adoration"),
        Liturgy("Prayer for Purity", "prayer_for_purity"),
        Liturgy("The Ten Commandments", "ten_commandments"),
        Liturgy("Lord's Prayer", "lords_prayer"),
        Reading("New Testament Reading"),
        Psalm,
        Teaching,
        Reading("Reading and Sermon"),
        Song,
        Liturgy("Confession", "confession"),
        Liturgy("Assurance of Forgiveness", "assurance"),
        Liturgy("Comfortable Words", "comfortable_words"),
        Liturgy("Prayer of Humble Access", "humble_access"),
        Liturgy("Prayer of Consecration", "consecration"),
        Prayer("Prayers of Intercession"),
        Song,
        Liturgy("Final Blessing", "final_blessing"),
        Prayer("Join us for refreshments"),
    ]
}

fn pm_lords_supper() -> Vec<ComponentSeed> {
    use ComponentSeed::*;
    vec![
        Welcome,
        Notices,
        Call,
        Song,
        Liturgy("Prayer for Purity", "prayer_for_purity"),
        Liturgy("The Ten Commandments", "ten_commandments"),
        Liturgy("Lord's Prayer", "lords_prayer"),
        Psalm,
        Teaching,
        Song,
        Reading("Reading and Sermon"),
        Song,
        Liturgy("Confession", "confession"),
        Liturgy("Assurance of Forgiveness", "assurance"),
        Liturgy("Comfortable Words", "comfortable_words"),
        Liturgy("Prayer of Humble Access", "humble_access"),
        Liturgy("Prayer of Consecration", "consecration"),
        Prayer("Prayers of Intercession"),
        Song,
        Liturgy("Final Blessing", "final_blessing"),
        Prayer("Join us for refreshments"),
    ]
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn all_presets_have_stable_unique_component_ids() {
        for preset in ServicePreset::all() {
            let components = preset.components();
            let ids: std::collections::HashSet<_> =
                components.iter().map(ServiceComponent::id).collect();
            assert_eq!(ids.len(), components.len(), "{preset:?}");
        }
    }

    #[test]
    fn communion_presets_include_the_full_liturgy() {
        for preset in [ServicePreset::AmLordsSupper, ServicePreset::PmLordsSupper] {
            let headings: Vec<_> = preset
                .components()
                .iter()
                .map(|component| component.heading().to_string())
                .collect();
            for required in [
                "Prayer for Purity",
                "The Ten Commandments",
                "Comfortable Words",
                "Prayer of Humble Access",
                "Prayer of Consecration",
                "Final Blessing",
            ] {
                assert!(headings.iter().any(|heading| heading == required));
            }
        }
    }

    #[test]
    fn praise_and_worship_restores_confession() {
        let components = ServicePreset::PraiseAndWorshipPm.components();
        let headings: Vec<_> = components.iter().map(ServiceComponent::heading).collect();
        assert!(headings.contains(&"Confession"));
        assert!(headings.contains(&"Assurance of Forgiveness"));
        assert!(headings.contains(&"Lord's Prayer"));
    }
}
