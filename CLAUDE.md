# CLAUDE.md — project guide for LLM assistants

TWPC Service Builder: a web app that lets church staff assemble an order of service
(songs, psalms, readings, notices, liturgy, teaching) and generate a PowerPoint deck
from a TWPC-branded template. Active development happens on the Rust web app
(`release/rust-web-app` branch); `legacy/python/` is an earlier prototype kept for
reference only.

## Workspace layout

- `crates/deck-builder` — domain model (`domain.rs`: `ServiceRecord`, `ServiceComponent`,
  presets) and PPTX generation (`build_deck` in `lib.rs`, a big match over component
  types that clones seed slides from `assets/template.pptx`). Text utilities live in
  `textproc.rs` (British spellings, scripture/psalm normalisation). External data
  sources implement the `Sources` trait (`lib.rs`); the live ESV API client is
  `sources/esv.rs` (needs `ESV_API_KEY`).
- `crates/pptx-template` — low-level OpenXML editing: open a .pptx, clone slides, set
  shape text/`Run`s (font, size, colour, bold/italic/underline/superscript), positions.
  No external PPTX library; XML is manipulated directly.
- `crates/server` — axum web server. Sessions + CSRF, JSON API (`/api/services`,
  `/api/scripture`, `/api/psalm`, `/api/teaching`, `/api/songs`, generation), askama
  HTML templates in `templates/`, and the frontend in `static/`.

## Frontend

Vanilla ES modules, no framework, no build step:

- `crates/server/static/app.js` — all DOM construction and rendering for the builder
  page (`createEditorApp`). Three-column layout: order list, editor panel, review panel.
- `crates/server/static/editor-controller.js` — state, debounced autosave
  (PUT `/api/services/:id/autosave`), and the psalm/ESV/teaching loaders
  (`createEditorController`). Render callbacks are injected from app.js.
- `crates/server/static/app.css` — single hand-written stylesheet, CSS custom
  properties, grid layout.

**Important:** static assets and all JSON data are embedded into the server binary
with `include_str!`/`include_bytes!` (`crates/server/src/lib.rs`,
`crates/deck-builder/src/lib.rs`). Editing a static or asset file requires a
`cargo build` for the running server to serve the change.

## Embedded data (`crates/deck-builder/assets/`)

- `template.pptx` — the seed deck; slides are cloned by index (`SEED_*` constants).
- `psalms.json` — Sing Psalms texts with verse-numbered stanzas.
- `wsc.json`, `heidelberg.json` — catechism Q&As (`{Number, Question, Answer}`).
- `wcf.json` — Westminster Confession chapters/sections. All three teaching sources
  resolve through `Teaching::find(source, selection)`; selections accept `Q. 1` style
  for catechisms and `21` / `21.8` for the confession.
- `components.json` — fixed liturgy wording (confession, assurance, Lord's Supper).

## Build, test, lint

- `cargo test --workspace` — Rust tests (integration tests in `crates/*/tests/`).
- `cargo fmt --check` and `cargo clippy --workspace --all-targets -- -D warnings` — both
  enforced by CI (`.github/workflows/ci.yml`).
- `npm test` — frontend vitest suite in `tests/` (jsdom; `tests/helpers/editor-fixture.js`
  builds the DOM). Not run by CI, run it locally.
- On this Raspberry Pi there is no host Rust toolchain: build/test inside Docker with
  `rust:1.97.1`, workspace bind-mounted read-only, using the warm named volumes
  `church-powerpoint-target` and `church-powerpoint-cargo-registry`
  (`CARGO_TARGET_DIR=/target`, `CARGO_HOME=/cargo-home`, `--cpus=2`, run as root).
  Avoid parallel/cold rebuilds — they have destabilised the Pi.

## PPTX correctness pitfalls

PowerPoint enforces invariants that schema validators (including the Open XML SDK)
do not. Known ones: `p:sldMasterId`/`p:sldLayoutId` ids must be globally unique, all
reachable masters must be registered, and each master must own its theme part. Real
PowerPoint is the only reliable oracle. Full write-up:
`docs/powerpoint-repair-postmortem.md`. Diagnose package structure with small Python
`zipfile`/ElementTree scripts rather than rebuilding repeatedly.

## Conventions

- Commit messages: imperative sentence, no prefixes; wrap body at ~72 chars.
- Frontend tests assert DOM stability (nodes not rebuilt while typing) — preserve the
  targeted-render scopes (`field`/`heading`/`summary`/`structural`) when editing
  `app.js`/`editor-controller.js`.
- All church-facing text uses British English; `textproc::british_spellings` converts
  fetched ESV text.
