# TWPC Service Builder

The supported application is an authenticated Rust/Axum website for preparing and generating TWPC service PowerPoints. Staff choose a service preset, arrange and edit its components, review warnings, and generate immutable `.pptx` revisions. The old Python generator is retained for reference only in [`legacy/python/`](legacy/python/).

## What is implemented

- AM, traditional PM, praise-and-worship PM, AM Lord's Supper, and PM Lord's Supper presets based on the supplied services
- Ordered, editable components for notices, call to worship, prayer cues, songs, psalms, readings, teaching, liturgy, and custom slides
- Staff sign-in using an Argon2 password hash, signed HTTP-only sessions, CSRF protection, login throttling, and audit display names
- Five-minute editing leases, autosave conflict detection, archive/restore, generated revision history, and 730-day deck-expiry metadata
- Server-side ESV lookup with manual-entry fallback
- Conditional object writes through an injected object-store trait, with Cloudflare R2 in production and an in-memory test implementation
- Server-rendered Askama pages, static CSS, and browser ES modules with no Node toolchain
- Embedded Sing Psalms, WSC, liturgy, and canonical 4:3 TWPC PowerPoint assets

## Configuration

Production requires:

```text
ESV_API_KEY
STAFF_PASSWORD_HASH
SESSION_SIGNING_SECRET
R2_ACCOUNT_ID
R2_BUCKET
R2_ACCESS_KEY_ID
R2_SECRET_ACCESS_KEY
```

`STAFF_PASSWORD_HASH` must be an Argon2 PHC string. `SESSION_SIGNING_SECRET` must contain at least 32 characters. Configure the R2 bucket as private and apply a 730-day lifecycle rule to `generated/services/`; entity definitions and immutable source versions must not use that expiry rule.

For disposable local development, set `OBJECT_STORE=memory`. Production defaults to R2 and fails closed when any R2 setting is missing.

## Development

The project pins Rust 1.97.1 in `rust-toolchain.toml`, CI, and Docker.

```bash
cargo fmt --check
cargo clippy --workspace --all-targets -- -D warnings
cargo test --workspace
cargo run -p server
```

The Docker image is the deployment source of truth:

```bash
docker build -t twpc-service-builder .
docker run --rm -p 8080:8080 --env-file .env twpc-service-builder
```

The supplied 86-deck song archive has a dry-run capable, idempotent importer. It validates the exact expected total of 365 ordered slides before writing immutable R2 objects:

```bash
cargo run -p server --bin import-song-library -- "Attachments-sample services.zip" --dry-run
cargo run -p server --bin import-song-library -- "Attachments-sample services.zip"
```

Only `/login`, `/healthz`, and the immutable static assets used by the login page are public. Service, preview, history, reporting, and generation APIs require a staff session.

## Repository layout

```text
crates/pptx-template/   OOXML package and slide operations
crates/deck-builder/    domain model, presets, embedded data, deck generation
crates/server/          Axum routes, authentication, storage, templates and UI
legacy/python/          unsupported historical generator and its assets
```
