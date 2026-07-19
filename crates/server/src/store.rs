use async_trait::async_trait;
use aws_credential_types::Credentials;
use aws_sdk_s3::config::Region;
use aws_sdk_s3::primitives::ByteStream;
use std::collections::BTreeMap;
use std::sync::{Arc, RwLock};
use thiserror::Error;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoredObject {
    pub bytes: Vec<u8>,
    pub etag: String,
    pub content_type: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PutCondition {
    Any,
    IfMatch(String),
    IfNoneMatch,
}

#[derive(Debug, Error)]
pub enum StoreError {
    #[error("object not found: {0}")]
    NotFound(String),
    #[error("object has changed since it was read")]
    PreconditionFailed,
    #[error("object store is unavailable: {0}")]
    Unavailable(String),
}

#[async_trait]
pub trait ObjectStore: Send + Sync {
    async fn get(&self, key: &str) -> Result<StoredObject, StoreError>;
    async fn put(
        &self,
        key: &str,
        bytes: Vec<u8>,
        content_type: &str,
        condition: PutCondition,
    ) -> Result<StoredObject, StoreError>;
    async fn delete(&self, key: &str) -> Result<(), StoreError>;
    async fn list(&self, prefix: &str) -> Result<Vec<String>, StoreError>;
}

#[derive(Clone)]
pub struct R2ObjectStore {
    client: aws_sdk_s3::Client,
    bucket: String,
}

impl R2ObjectStore {
    pub fn new(
        account_id: impl Into<String>,
        bucket: impl Into<String>,
        access_key_id: impl Into<String>,
        secret_access_key: impl Into<String>,
    ) -> Result<Self, StoreError> {
        let account_id = account_id.into();
        let bucket = bucket.into();
        let access_key_id = access_key_id.into();
        let secret_access_key = secret_access_key.into();
        if account_id.trim().is_empty()
            || bucket.trim().is_empty()
            || access_key_id.trim().is_empty()
            || secret_access_key.trim().is_empty()
        {
            return Err(StoreError::Unavailable(
                "R2 account, bucket and credentials are required".into(),
            ));
        }
        let credentials = Credentials::new(
            access_key_id,
            secret_access_key,
            None,
            None,
            "twpc-r2-configuration",
        );
        let config = aws_sdk_s3::Config::builder()
            .behavior_version_latest()
            .endpoint_url(format!("https://{account_id}.r2.cloudflarestorage.com"))
            .region(Region::new("auto"))
            .credentials_provider(credentials)
            .force_path_style(true)
            .build();
        Ok(Self {
            client: aws_sdk_s3::Client::from_conf(config),
            bucket,
        })
    }

    pub fn from_env() -> Result<Self, StoreError> {
        let env = |name: &str| {
            std::env::var(name).map_err(|_| StoreError::Unavailable(format!("{name} is required")))
        };
        Self::new(
            env("R2_ACCOUNT_ID")?,
            env("R2_BUCKET")?,
            env("R2_ACCESS_KEY_ID")?,
            env("R2_SECRET_ACCESS_KEY")?,
        )
    }
}

#[async_trait]
impl ObjectStore for R2ObjectStore {
    async fn get(&self, key: &str) -> Result<StoredObject, StoreError> {
        let response = self
            .client
            .get_object()
            .bucket(&self.bucket)
            .key(key)
            .send()
            .await
            .map_err(|error| classify_s3_error(key, &error.to_string()))?;
        let etag = response.e_tag().unwrap_or_default().to_string();
        let content_type = response
            .content_type()
            .unwrap_or("application/octet-stream")
            .to_string();
        let bytes = response
            .body
            .collect()
            .await
            .map_err(|error| StoreError::Unavailable(error.to_string()))?
            .into_bytes()
            .to_vec();
        Ok(StoredObject {
            bytes,
            etag,
            content_type,
        })
    }

    async fn put(
        &self,
        key: &str,
        bytes: Vec<u8>,
        content_type: &str,
        condition: PutCondition,
    ) -> Result<StoredObject, StoreError> {
        let mut request = self
            .client
            .put_object()
            .bucket(&self.bucket)
            .key(key)
            .content_type(content_type)
            .body(ByteStream::from(bytes.clone()));
        request = match condition {
            PutCondition::Any => request,
            PutCondition::IfMatch(etag) => request.if_match(etag),
            PutCondition::IfNoneMatch => request.if_none_match("*"),
        };
        let response = request
            .send()
            .await
            .map_err(|error| classify_s3_error(key, &error.to_string()))?;
        Ok(StoredObject {
            bytes,
            etag: response.e_tag().unwrap_or_default().to_string(),
            content_type: content_type.to_string(),
        })
    }

    async fn delete(&self, key: &str) -> Result<(), StoreError> {
        self.client
            .delete_object()
            .bucket(&self.bucket)
            .key(key)
            .send()
            .await
            .map_err(|error| classify_s3_error(key, &error.to_string()))?;
        Ok(())
    }

    async fn list(&self, prefix: &str) -> Result<Vec<String>, StoreError> {
        let mut keys = Vec::new();
        let mut continuation: Option<String> = None;
        loop {
            let response = self
                .client
                .list_objects_v2()
                .bucket(&self.bucket)
                .prefix(prefix)
                .set_continuation_token(continuation)
                .send()
                .await
                .map_err(|error| classify_s3_error(prefix, &error.to_string()))?;
            keys.extend(
                response
                    .contents()
                    .iter()
                    .filter_map(|object| object.key().map(str::to_string)),
            );
            if response.is_truncated() != Some(true) {
                break;
            }
            continuation = response.next_continuation_token().map(str::to_string);
        }
        Ok(keys)
    }
}

fn classify_s3_error(key: &str, message: &str) -> StoreError {
    if message.contains("NoSuchKey") || message.contains("NotFound") || message.contains("404") {
        StoreError::NotFound(key.to_string())
    } else if message.contains("PreconditionFailed") || message.contains("412") {
        StoreError::PreconditionFailed
    } else {
        StoreError::Unavailable(message.to_string())
    }
}

#[derive(Clone, Default)]
pub struct MemoryObjectStore {
    inner: Arc<RwLock<MemoryState>>,
}

#[derive(Default)]
struct MemoryState {
    objects: BTreeMap<String, StoredObject>,
    next_etag: u64,
}

#[async_trait]
impl ObjectStore for MemoryObjectStore {
    async fn get(&self, key: &str) -> Result<StoredObject, StoreError> {
        self.inner
            .read()
            .map_err(|_| StoreError::Unavailable("read lock poisoned".into()))?
            .objects
            .get(key)
            .cloned()
            .ok_or_else(|| StoreError::NotFound(key.to_string()))
    }

    async fn put(
        &self,
        key: &str,
        bytes: Vec<u8>,
        content_type: &str,
        condition: PutCondition,
    ) -> Result<StoredObject, StoreError> {
        let mut state = self
            .inner
            .write()
            .map_err(|_| StoreError::Unavailable("write lock poisoned".into()))?;
        let existing = state.objects.get(key);
        let allowed = match condition {
            PutCondition::Any => true,
            PutCondition::IfNoneMatch => existing.is_none(),
            PutCondition::IfMatch(ref wanted) => {
                existing.is_some_and(|object| object.etag == *wanted)
            }
        };
        if !allowed {
            return Err(StoreError::PreconditionFailed);
        }
        state.next_etag = state.next_etag.saturating_add(1);
        let object = StoredObject {
            bytes,
            etag: format!("\"{:016x}\"", state.next_etag),
            content_type: content_type.to_string(),
        };
        state.objects.insert(key.to_string(), object.clone());
        Ok(object)
    }

    async fn delete(&self, key: &str) -> Result<(), StoreError> {
        self.inner
            .write()
            .map_err(|_| StoreError::Unavailable("write lock poisoned".into()))?
            .objects
            .remove(key);
        Ok(())
    }

    async fn list(&self, prefix: &str) -> Result<Vec<String>, StoreError> {
        Ok(self
            .inner
            .read()
            .map_err(|_| StoreError::Unavailable("read lock poisoned".into()))?
            .objects
            .keys()
            .filter(|key| key.starts_with(prefix))
            .cloned()
            .collect())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[tokio::test]
    async fn conditional_writes_prevent_lost_updates() {
        let store = MemoryObjectStore::default();
        let first = store
            .put(
                "entities/services/one.json",
                b"one".to_vec(),
                "application/json",
                PutCondition::IfNoneMatch,
            )
            .await
            .unwrap();
        assert!(matches!(
            store
                .put(
                    "entities/services/one.json",
                    b"stale".to_vec(),
                    "application/json",
                    PutCondition::IfMatch("\"stale\"".into()),
                )
                .await,
            Err(StoreError::PreconditionFailed)
        ));
        let updated = store
            .put(
                "entities/services/one.json",
                b"two".to_vec(),
                "application/json",
                PutCondition::IfMatch(first.etag),
            )
            .await
            .unwrap();
        assert_eq!(updated.bytes, b"two");
    }
}
