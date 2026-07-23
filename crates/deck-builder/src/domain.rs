use chrono::{DateTime, NaiveDate, Utc};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct AuditMetadata {
    pub created_at: DateTime<Utc>,
    pub created_by: String,
    pub updated_at: DateTime<Utc>,
    pub updated_by: String,
}

impl AuditMetadata {
    pub fn new(staff_name: impl Into<String>) -> Self {
        let staff_name = staff_name.into();
        let now = Utc::now();
        Self {
            created_at: now,
            created_by: staff_name.clone(),
            updated_at: now,
            updated_by: staff_name,
        }
    }

    pub fn touch(&mut self, staff_name: impl Into<String>) {
        self.updated_at = Utc::now();
        self.updated_by = staff_name.into();
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum ServicePeriod {
    Am,
    Pm,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum PmStyle {
    Traditional,
    PraiseAndWorship,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum ServicePreset {
    Am,
    TraditionalPm,
    PraiseAndWorshipPm,
    AmLordsSupper,
    PmLordsSupper,
}

impl ServicePreset {
    pub fn is_lords_supper(self) -> bool {
        matches!(self, Self::AmLordsSupper | Self::PmLordsSupper)
    }

    pub fn label(self) -> &'static str {
        match self {
            Self::Am => "Morning service",
            Self::TraditionalPm => "Traditional evening service",
            Self::PraiseAndWorshipPm => "Praise and worship evening",
            Self::AmLordsSupper => "Morning Lord's Supper",
            Self::PmLordsSupper => "Evening Lord's Supper",
        }
    }

    pub fn all() -> [Self; 5] {
        [
            Self::Am,
            Self::TraditionalPm,
            Self::PraiseAndWorshipPm,
            Self::AmLordsSupper,
            Self::PmLordsSupper,
        ]
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum ServiceStatus {
    Draft,
    Completed,
    Archived,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ServiceRecord {
    pub id: String,
    pub name: String,
    pub date: NaiveDate,
    #[serde(default)]
    pub period: Option<ServicePeriod>,
    #[serde(default)]
    pub pm_style: Option<PmStyle>,
    #[serde(default)]
    pub lords_supper: bool,
    pub preset: ServicePreset,
    pub status: ServiceStatus,
    pub revision: u64,
    pub components: Vec<ServiceComponent>,
    pub audit: AuditMetadata,
}

impl ServiceRecord {
    pub fn new(
        id: impl Into<String>,
        name: impl Into<String>,
        date: NaiveDate,
        preset: ServicePreset,
        staff_name: impl Into<String>,
    ) -> Self {
        let staff_name = staff_name.into();
        let (period, pm_style, lords_supper) = match preset {
            ServicePreset::Am => (Some(ServicePeriod::Am), None, false),
            ServicePreset::TraditionalPm => {
                (Some(ServicePeriod::Pm), Some(PmStyle::Traditional), false)
            }
            ServicePreset::PraiseAndWorshipPm => (
                Some(ServicePeriod::Pm),
                Some(PmStyle::PraiseAndWorship),
                false,
            ),
            ServicePreset::AmLordsSupper => (Some(ServicePeriod::Am), None, true),
            ServicePreset::PmLordsSupper => {
                (Some(ServicePeriod::Pm), Some(PmStyle::Traditional), true)
            }
        };
        Self {
            id: id.into(),
            name: name.into(),
            date,
            period,
            pm_style,
            lords_supper,
            preset,
            status: ServiceStatus::Draft,
            revision: 0,
            components: preset.components(),
            audit: AuditMetadata::new(staff_name),
        }
    }

    pub fn mark_edited(&mut self, staff_name: impl Into<String>) {
        if self.status == ServiceStatus::Completed {
            self.status = ServiceStatus::Draft;
        }
        self.revision = self.revision.saturating_add(1);
        self.audit.touch(staff_name);
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct NoticeRow {
    pub when: String,
    pub title: String,
    pub details: String,
    #[serde(default)]
    pub emphasis: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct VersionPin {
    pub entity_id: String,
    pub version: u64,
    #[serde(default)]
    pub slide_count: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ImagePin {
    pub object_key: String,
    pub version: u64,
    pub content_type: String,
    pub alt_text: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum ServiceComponent {
    Welcome {
        id: String,
        heading: String,
    },
    Notices {
        id: String,
        heading: String,
        rows: Vec<NoticeRow>,
    },
    CallToWorship {
        id: String,
        heading: String,
        reference: String,
        text: String,
        #[serde(default)]
        external_source_failed: bool,
    },
    CuePrayer {
        id: String,
        heading: String,
        #[serde(default)]
        cue: String,
        #[serde(default)]
        text: String,
    },
    Song {
        id: String,
        title: String,
        #[serde(default)]
        song: Option<VersionPin>,
        #[serde(default)]
        lyric_slides: Vec<String>,
        #[serde(default)]
        credits: String,
    },
    Psalm {
        id: String,
        heading: String,
        reference: String,
        #[serde(default)]
        tune: Option<VersionPin>,
        #[serde(default)]
        slide_breaks: Vec<String>,
    },
    Reading {
        id: String,
        heading: String,
        reference: String,
        #[serde(default)]
        bible_page: Option<u16>,
    },
    Teaching {
        id: String,
        heading: String,
        source: TeachingSource,
        selection: String,
        #[serde(default)]
        text: String,
    },
    LiturgyBlock {
        id: String,
        heading: String,
        key: String,
        #[serde(default)]
        version: Option<u64>,
        #[serde(default)]
        text: String,
    },
    CustomTextImage {
        id: String,
        heading: String,
        slides: Vec<String>,
        #[serde(default)]
        image: Option<ImagePin>,
    },
}

impl ServiceComponent {
    pub fn id(&self) -> &str {
        match self {
            Self::Welcome { id, .. }
            | Self::Notices { id, .. }
            | Self::CallToWorship { id, .. }
            | Self::CuePrayer { id, .. }
            | Self::Song { id, .. }
            | Self::Psalm { id, .. }
            | Self::Reading { id, .. }
            | Self::Teaching { id, .. }
            | Self::LiturgyBlock { id, .. }
            | Self::CustomTextImage { id, .. } => id,
        }
    }

    pub fn heading(&self) -> &str {
        match self {
            Self::Welcome { heading, .. }
            | Self::Notices { heading, .. }
            | Self::CallToWorship { heading, .. }
            | Self::CuePrayer { heading, .. }
            | Self::Psalm { heading, .. }
            | Self::Reading { heading, .. }
            | Self::Teaching { heading, .. }
            | Self::LiturgyBlock { heading, .. }
            | Self::CustomTextImage { heading, .. } => heading,
            Self::Song { title, .. } => title,
        }
    }

    pub fn estimated_slides(&self) -> usize {
        match self {
            Self::Notices { rows, .. } => rows.len().max(1).div_ceil(5),
            Self::Song {
                song, lyric_slides, ..
            } => song
                .as_ref()
                .map(|pin| pin.slide_count)
                .unwrap_or(lyric_slides.len())
                .max(1),
            Self::Psalm { slide_breaks, .. } => slide_breaks.len().max(1),
            Self::LiturgyBlock { text, .. } => text.split("\n\n").count().max(1),
            Self::CustomTextImage { slides, .. } => slides.len().max(1),
            _ => 1,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum TeachingSource {
    WestminsterShorterCatechism,
    Heidelberg1891,
    WestminsterConfessionOriginalBritish,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum RightsStatus {
    PublicDomain,
    CcliCovered,
    DirectPermission,
    Unknown,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct SongRecord {
    pub id: String,
    pub title: String,
    #[serde(default)]
    pub aliases: Vec<String>,
    #[serde(default)]
    pub variant_label: String,
    #[serde(default)]
    pub author_owner: String,
    pub rights_status: RightsStatus,
    #[serde(default)]
    pub ccli_song_number: Option<String>,
    #[serde(default)]
    pub source_filename: Option<String>,
    pub slide_count: usize,
    pub current_version: u64,
    #[serde(default)]
    pub archived: bool,
    #[serde(default)]
    pub review_warnings: Vec<String>,
    pub audit: AuditMetadata,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct SongVersion {
    pub song_id: String,
    pub version: u64,
    pub object_key: String,
    pub sha256: String,
    pub slide_count: usize,
    pub extracted_text: Vec<String>,
    pub created_at: DateTime<Utc>,
    pub created_by: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct TuneRecord {
    pub id: String,
    pub name: String,
    pub composer: String,
    pub meter: String,
    pub notes: String,
    pub version: u64,
    pub archived: bool,
    pub audit: AuditMetadata,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct LiturgyVersion {
    pub key: String,
    pub version: u64,
    pub heading: String,
    pub text: String,
    pub created_at: DateTime<Utc>,
    pub created_by: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct GeneratedDeckVersion {
    pub service_id: String,
    pub revision: u64,
    pub object_key: String,
    pub generated_at: DateTime<Utc>,
    pub generated_by: String,
    pub expires_at: DateTime<Utc>,
    pub source_revision: u64,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct GlobalSettingsVersion {
    pub version: u64,
    pub ccli_licence_number: String,
    pub created_at: DateTime<Utc>,
    pub created_by: String,
}

impl Default for GlobalSettingsVersion {
    fn default() -> Self {
        Self {
            version: 1,
            ccli_licence_number: "522221".to_string(),
            created_at: Utc::now(),
            created_by: "system".to_string(),
        }
    }
}
