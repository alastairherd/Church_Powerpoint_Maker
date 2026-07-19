# Church PowerPoint Maker — Rust Web App: Implementation Plan

**Goal:** Rebuild [alastairherd/Church_Powerpoint_Maker](https://github.com/alastairherd/Church_Powerpoint_Maker) as a Rust web application. A visitor opens the site, fills in (or confirms) the details of the week's service, clicks Generate, and downloads a ready-to-use `.pptx`.

This document is written to be executed by a coding agent. Work through the phases in order; each phase ends with a verifiable checkpoint.

---

## 1. What the original Python project does

Understanding the original is required before porting. `full_code.py` does the following:

1. **Reads the weekly service plan** from a published Google Sheet (fetched as CSV) — date, psalm/hymn selections, scripture readings, catechism question number, sermon title, etc.
2. **Fetches scripture text** from the ESV API (`api.esv.org`) using an API key.
3. **Scrapes hymn lyrics** from hymnary.org given a hymn URL.
4. **Loads local data files:**
   - `psalms.json` — full lyrics of the Sing Psalms psalter, keyed by psalm reference
   - `wsc.json` — Westminster Shorter Catechism questions and answers
   - `components.json` — fixed liturgical elements and service order
5. **Processes text:** converts American to British spellings; renders verse numbers as superscript runs; splits long passages across multiple slides.
6. **Fills a template** (`template.pptx`) using python-pptx: clones placeholder slides for each service component and writes title/body/copyright text into placeholders.
7. **Saves** the finished deck.

**⚠️ Security note:** the original `full_code.py` contains a hardcoded ESV API key committed to the public repo. That key must be treated as compromised — revoke it at api.esv.org and issue a new one before this project goes live. The new key lives only in the host's secret/environment-variable store, never in the repository.

---

## 2. Architecture decision

### Constraints that shape the design

- The **ESV API key must remain secret**, so scripture fetching must happen server-side (or edge-side), never in the browser.
- **hymnary.org scraping cannot run in the browser** (CORS), so it also needs a server or edge function.
- Traffic is tiny and bursty: realistically a handful of generations per week. Cold starts are acceptable.
- A `.pptx` is a ZIP of XML files. Generating one in memory is cheap (< a few MB, < 1s), so no disk or database is required. The app can be fully stateless.

### Option A — Axum server (recommended)

A single Rust binary using [Axum](https://github.com/tokio-rs/axum):

- Serves the static frontend (one HTML page + a little JS/CSS) via `tower-http`'s `ServeDir` or embedded with `include_str!`.
- Exposes `POST /api/generate` which returns the `.pptx` bytes with `Content-Disposition: attachment`.
- Uses `reqwest` for the ESV API and hymnary.org; `scraper` crate for HTML parsing.
- Full standard library and crate ecosystem available; easiest to develop and test locally (`cargo run`).
- Deploys as a Docker container or native build on Render/Northflank.

### Option B — Cloudflare Workers (Rust → WASM via `workers-rs`)

- The whole app compiles to WASM and runs at the edge. Free tier is generous (100k requests/day) with no cold-start penalty worth noticing.
- Secrets (ESV key) are handled by Wrangler's secret store.
- **Trade-offs:** the WASM target restricts crates — no `tokio`/`reqwest` (use the Workers `fetch` API instead), no threads, and some zip/xml crates need `wasm` feature flags or alternatives. Local dev runs through `wrangler dev` rather than plain `cargo run`. Debugging is harder.

### Option C — Vercel

- Rust is supported only through a community serverless runtime; the platform is optimized for JS/Next.js. Workable but the most awkward fit of the three. Not recommended unless there's a strong reason.

### Decision

**Build Option A (Axum).** It is the simplest to develop, test, and debug, and the pptx engine (Phase 2) stays a plain Rust library that could later be recompiled for Workers if desired. Structure the workspace so the core logic is host-agnostic (see §4).

---

## 3. Hosting comparison (state of play, mid-2026)

| Platform | Free tier | Cold starts | Rust story | Custom domain on free | Notes |
|---|---|---|---|---|---|
| **Render** | Free web services, no credit card | Yes — idle services sleep, ~1 min wake | Auto-detected native builds or Docker | Yes | Recommended starting point. Wake delay is fine for weekly use. |
| **Northflank** | 2 services, 1 vCPU, 1 GB RAM | No forced sleep | Docker/buildpacks | Yes | Strongest genuinely free tier; good upgrade path. |
| **Cloudflare Workers** | 100k req/day | Effectively none | First-class via `workers-rs`, but WASM constraints | `workers.dev` subdomain; custom domains supported | Best if we ever rewrite for Option B. |
| **Vercel** | Hobby tier, serverless functions | Minor | Community runtime only | Yes | Frontend-focused; awkward for Rust. |
| **Fly.io** | None for new users (trial only, card required) | n/a | Excellent (Docker) | Yes | ~$2–5/mo minimum for always-on. |
| **Railway** | Trial/usage credits, not perpetual free | n/a | Good (Docker/Nixpacks) | Yes | Nice DX, but not free long-term. |

**Plan: deploy to Render free tier first** (`churchdeck.onrender.com` or similar). Include a `Dockerfile` so the app is portable to Northflank or a VPS with zero code changes. No custom domain needed initially — all these hosts provide a free HTTPS subdomain.

---

## 4. Project structure

Cargo workspace with three crates so the core stays reusable and testable:

```
church-deck/
├── Cargo.toml                 # workspace
├── Dockerfile
├── render.yaml                # Render blueprint (optional)
├── crates/
│   ├── pptx-template/         # the mini python-pptx port (library)
│   │   ├── src/
│   │   │   ├── lib.rs
│   │   │   ├── package.rs     # zip open/save, [Content_Types], rels
│   │   │   ├── presentation.rs
│   │   │   ├── slide.rs       # clone, reorder, delete slides
│   │   │   ├── placeholder.rs # find by idx/type, set text
│   │   │   ├── text.rs        # paragraphs, runs, superscript, formatting
│   │   │   └── xmlutil.rs     # quick-xml helpers
│   │   └── tests/
│   ├── deck-builder/          # domain logic (library)
│   │   ├── src/
│   │   │   ├── lib.rs
│   │   │   ├── model.rs       # ServiceOrder, Component enums
│   │   │   ├── sources/
│   │   │   │   ├── esv.rs     # ESV API client
│   │   │   │   ├── hymnary.rs # scraper
│   │   │   │   ├── psalms.rs  # embedded psalms.json
│   │   │   │   ├── wsc.rs     # embedded wsc.json
│   │   │   │   └── sheet.rs   # Google Sheet CSV fetch/parse
│   │   │   ├── textproc.rs    # UK spellings, verse superscripts, slide splitting
│   │   │   └── build.rs       # ServiceOrder -> pptx bytes
│   │   └── assets/            # template.pptx, psalms.json, wsc.json, components.json
│   └── server/                # Axum binary
│       └── src/main.rs
```

Key dependencies: `zip`, `quick-xml`, `serde`/`serde_json`, `axum`, `tokio`, `reqwest`, `scraper`, `regex`, `csv`, `thiserror`, `anyhow`.

Data files and `template.pptx` are embedded into the binary with `include_bytes!`/`include_str!` so deployment is a single artifact.

---

## 5. Phase 1 — `pptx-template` crate (the mini python-pptx port)

**Scope discipline:** implement only what this project needs. Do not attempt general OOXML coverage. No charts, images, tables, or theme editing.

### 5.1 How a .pptx works (for the agent)

A `.pptx` is a ZIP containing, among other things:

- `ppt/presentation.xml` — lists slide IDs in order (`<p:sldIdLst>`)
- `ppt/_rels/presentation.xml.rels` — maps relationship IDs to slide part filenames
- `ppt/slides/slideN.xml` — each slide's shape tree; placeholders are `<p:sp>` elements whose `<p:nvSpPr>/<p:nvPr>/<p:ph>` carries a `type` and/or `idx` attribute
- `[Content_Types].xml` — must declare every part, including each slide
- Text lives in `<p:txBody>` → `<a:p>` (paragraph) → `<a:r>` (run) → `<a:t>` (text). Superscript is a run property: `<a:rPr baseline="30000"/>`.

### 5.2 Public API (target)

```rust
let mut pres = Presentation::open_bytes(TEMPLATE_BYTES)?;

// Template slides are addressed by index; clone one as a new slide at the end
let idx = pres.clone_slide(SONG_TEMPLATE_IDX)?;
let slide = pres.slide_mut(idx)?;
slide.placeholder_by_idx(0)?.set_text("Psalm 23");
slide.placeholder_by_idx(1)?.set_rich_text(&runs); // Vec<Run> with per-run formatting
slide.placeholder_by_idx(2)?.set_text("Sing Psalms © 2003");

pres.delete_slide(SONG_TEMPLATE_IDX)?;             // remove unused template slides
pres.reorder(&final_order)?;
let bytes: Vec<u8> = pres.save_bytes()?;
```

`Run` should support at minimum: text, `superscript: bool`, `bold`, `italic`. `set_rich_text` clears existing paragraphs in the placeholder and writes new `<a:p>`/`<a:r>` structures, copying the placeholder's existing default run properties where present so template fonts/sizes are preserved.

### 5.3 Cloning a slide correctly (the fiddly part)

Cloning slide N as a new slide requires all of:

1. Copy `ppt/slides/slideN.xml` to a new part `ppt/slides/slideM.xml` (M = next free number).
2. Copy `ppt/slides/_rels/slideN.xml.rels` to `ppt/slides/_rels/slideM.xml.rels` (preserves the layout relationship).
3. Add an `<Override>` for the new part in `[Content_Types].xml`.
4. Add a new `<Relationship>` in `ppt/_rels/presentation.xml.rels` with a fresh `rId`.
5. Append a `<p:sldId>` (fresh numeric id ≥ 256, unique) referencing that `rId` in `<p:sldIdLst>`.

Deletion is the inverse. Get this right once, test it hard (see 5.4), and everything else is easy.

### 5.4 Tests / checkpoint

- Round-trip test: open template → save unchanged → output opens in LibreOffice/PowerPoint without repair prompts. Automate a structural check: unzip output, assert content-types entries match slide parts, assert every `r:embed`/`r:id` referenced in slide XML exists in its rels file.
- Clone test: clone a slide 10×, set distinct text in each, verify order and text by re-parsing the output.
- Superscript test: verify `baseline` attribute present on the expected runs.
- **Manual checkpoint:** open a generated file in real PowerPoint (or PowerPoint online) and confirm no repair dialog. This is the acceptance bar for Phase 1.

---

## 6. Phase 2 — `deck-builder` crate

### 6.1 Domain model

```rust
enum Component {
    Psalm { reference: String },          // from embedded psalms.json
    Hymn { url: String },                 // scraped from hymnary.org
    Scripture { reference: String },      // ESV API
    Catechism { question: u16 },          // embedded wsc.json
    Fixed { key: String },                // from components.json (call to worship, benediction, …)
    Sermon { title: String, text: String },
}

struct ServiceOrder {
    date: NaiveDate,
    components: Vec<Component>,
}
```

Port the service-order logic from the Python code and `components.json` — the mapping from sheet columns to components, and which template slide each component type uses.

### 6.2 Data sources

- **ESV API** (`GET https://api.esv.org/v3/passage/text/`): request with the params the Python code uses (no footnotes/headings, keep verse numbers). API key from env var `ESV_API_KEY`. Handle 401/rate-limit errors with a clear message surfaced to the user.
- **hymnary.org scraper:** port the Python scraping logic using `scraper` with the same CSS selectors; verify selectors against the live site during development since they may have changed. Set a descriptive User-Agent. **Copyright note:** only public-domain hymn texts should be included in decks; carry over whatever the original project does for copyright lines, and display the hymn's copyright/attribution on the slide as the original does.
- **Psalms / WSC / fixed components:** embed the JSON files from the original repo; deserialize with serde at startup (`once_cell`/`LazyLock`).
- **Google Sheet (optional feature):** the original pulls the week's plan from a published-CSV Google Sheet. Support this as a convenience: a `GET /api/prefill` endpoint fetches and parses the CSV (via the `csv` crate) and returns JSON the frontend uses to pre-populate the form. The sheet URL comes from an env var `SHEET_CSV_URL`. If unset, the feature is hidden. Manual form entry must always work — do not make the sheet a hard dependency.

### 6.3 Text processing (`textproc.rs`)

Port from the Python:

- **American → British spelling** replacements (port the exact replacement table from the Python source).
- **Verse numbers → superscript runs:** parse ESV output; verse markers like `[23]` become superscript runs, body text becomes normal runs.
- **Slide splitting:** long passages/lyrics split across multiple cloned slides. Port the original's line/character thresholds; make them constants with the Python values as defaults.
- Stanza handling for psalms/hymns: one stanza (or stanza pair, matching the original) per slide.

### 6.4 Builder

`build_deck(order: &ServiceOrder, sources: &impl Sources) -> Result<Vec<u8>>` — pure function from a resolved service order to pptx bytes, so it is unit-testable with mocked sources. `Sources` is a trait so ESV/hymnary can be faked in tests.

**Checkpoint:** a CLI harness (`cargo run -p deck-builder --example demo`) that builds a full sample service offline (mocked ESV/hymn responses) and writes `out.pptx`; verify manually in PowerPoint.

---

## 7. Phase 3 — `server` crate (Axum)

### Endpoints

- `GET /` — serves the form page (embedded static HTML/CSS/JS; no frontend framework).
- `GET /api/prefill` — optional Google Sheet prefill (see 6.2).
- `POST /api/generate` — JSON body describing the service order; responds with the pptx bytes, `Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation` and `Content-Disposition: attachment; filename="service-YYYY-MM-DD.pptx"`.
- `GET /healthz` — 200 OK (Render health checks).

### Frontend (single page)

- Date picker; dynamic list of components (add/remove/reorder rows; each row = type dropdown + reference field).
- "Load this week's plan" button if prefill is configured.
- Generate button → `fetch` POST → trigger browser download from the blob.
- Show progress state and readable errors (e.g., "ESV API rejected the request", "Couldn't read hymn page").
- Keep it dependency-free vanilla JS; total frontend under ~300 lines.

### Hardening

- Request timeout (e.g., 30 s) and a concurrency limit (`tower` layers) — generation does outbound HTTP calls.
- Basic rate limit per IP (e.g., `tower_governor`) to protect the ESV key quota.
- Validate inputs: cap component count (e.g., ≤ 30), cap reference string lengths, allowlist hymn URLs to `hymnary.org` hosts only (prevents SSRF via the scraper).
- Never log the API key; read it once at startup and fail fast with a clear message if missing.

---

## 8. Phase 4 — Deployment (Render)

1. `Dockerfile`: multi-stage build — `rust:1-slim` builder → `debian:stable-slim` (or distroless) runtime; final image just the binary. Bind to `0.0.0.0:$PORT` (Render injects `PORT`).
2. Create a Render **Web Service** from the GitHub repo, Docker runtime, free plan.
3. Set env vars in Render dashboard: `ESV_API_KEY` (the **new**, regenerated key), optionally `SHEET_CSV_URL`.
4. Health check path `/healthz`.
5. Verify end-to-end on `*.onrender.com`, including the ~1 min cold-start wake behaving acceptably (frontend should show a loading state on first request).
6. Optional later: attach a custom domain (Render free tier supports this) — e.g., a free `is-a.dev` subdomain or a cheap .uk/.xyz.

---

## 9. Testing summary

- **Unit:** text processing (spellings, superscript parsing, slide splitting) with fixtures copied from real ESV responses; JSON deserialization of all three data files; XML placeholder finding.
- **Integration:** full deck build with mocked sources; structural validation of output zip (content types ↔ parts ↔ rels all consistent).
- **Manual acceptance:** open generated decks in desktop PowerPoint and LibreOffice — no repair prompt, correct fonts/formatting inherited from template, superscripts render, slide order correct.
- **CI:** GitHub Actions — `cargo fmt --check`, `clippy -D warnings`, `cargo test` on every push.

---

## 10. Milestones (suggested order of execution)

1. Workspace scaffold + CI + Dockerfile skeleton.
2. `pptx-template`: open/save round-trip passing PowerPoint's repair check.
3. `pptx-template`: clone/delete/reorder + placeholder text + rich runs (superscript). Phase 1 checkpoint.
4. `deck-builder`: embedded data sources (psalms, WSC, components) + text processing with unit tests.
5. `deck-builder`: ESV client + hymnary scraper behind the `Sources` trait; CLI demo produces a full deck. Phase 2 checkpoint.
6. `server`: endpoints + frontend form; local end-to-end works.
7. Hardening (limits, validation, SSRF allowlist) + Sheet prefill.
8. Deploy to Render; end-to-end verification. 🎉

---

## 11. Open questions to resolve before/while building

1. Does the existing `template.pptx` from the repo remain the design source of truth? (Recommended: yes — copy it into `crates/deck-builder/assets/`.) Document its placeholder indices per layout once, in a comment.
2. Should the site expose the Google Sheet prefill publicly, or is manual entry enough for v1?
3. Any authentication needed (is this public, or should a simple shared passcode gate `POST /api/generate` to protect the ESV quota)? A single `SITE_PASSCODE` env var check is a cheap option.
4. Confirm hymnary.org's current HTML structure and terms; adjust selectors and keep attribution/copyright lines on slides.
