# Service Content Parity Investigation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Investigate and correct the remaining content-parity gaps between the original Python deck generator and the new Rust web app.

**Architecture:** Keep the `pptx-template` crate focused on OOXML mechanics and put service-content decisions in `deck-builder`. Use failing tests that compare generated deck text against the original Python behavior before changing implementation.

**Tech Stack:** Rust workspace, `deck-builder`, `pptx-template`, Axum server, mocked `Sources` implementations, embedded `template.pptx`, `psalms.json`, `wsc.json`, and `components.json`.

## Global Constraints

- Do not use the committed historical ESV key for production; it remains compromised.
- Preserve server-side ESV and Hymnary fetching so secrets and scraping do not run in the browser.
- Keep generated `.pptx` files structurally valid and parse every generated slide XML part in tests.
- Match the original Python output where the original behavior is intentional.
- Use TDD for each behavior change: write the failing test, verify it fails, then implement.

---

### Task 1: Reading Slide Content Parity

**Files:**
- Modify: `crates/deck-builder/src/lib.rs`
- Test: `crates/deck-builder/tests/build_deck.rs`

**Interfaces:**
- Consumes: `Component::Scripture { reference, title }`, `Sources::scripture(reference) -> Scripture`
- Produces: reading slides whose visible text matches the intended design for references, body text, and page marker handling

- [ ] **Step 1: Confirm intended reading behavior from the original project**

Read `slide_making.py` lines 156-160 and `full_code.py` lines 587-591. Record whether reading slides should show only `"<reference>\n\npg. X"`, full ESV scripture text, or both.

- [ ] **Step 2: Write the failing test**

Add a test in `crates/deck-builder/tests/build_deck.rs` that builds a service order containing:

```rust
Component::Scripture {
    reference: "Genesis 1:1".to_string(),
    title: Some("First Reading".to_string()),
}
```

Assert the generated deck text contains the expected reference and page marker according to Step 1.

- [ ] **Step 3: Run the focused test and verify it fails**

Run:

```bash
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test -p deck-builder reading_slide
```

Expected: FAIL because the current implementation places fetched scripture body text on the reading slide and does not preserve the original reading-reference/page-marker behavior.

- [ ] **Step 4: Implement the minimal reading-slide change**

Update the `Component::Scripture` branch in `build_deck` to match the confirmed behavior from Step 1. If full scripture text is still desired elsewhere, introduce a separate component variant in a later plan rather than overloading reading slides.

- [ ] **Step 5: Verify**

Run:

```bash
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test -p deck-builder reading_slide
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test --workspace
```

Expected: all tests pass.

---

### Task 2: Hymnary Metadata And Tune Selection

**Files:**
- Modify: `crates/deck-builder/src/sources/hymnary.rs`
- Modify: `crates/deck-builder/src/lib.rs`
- Modify: `crates/server/src/lib.rs`
- Test: `crates/deck-builder/tests/build_deck.rs`
- Test: `crates/server/tests/endpoints.rs`

**Interfaces:**
- Consumes: hymn text URL and tune URL or tune identifier entered through the web form
- Produces: `Hymn { title, stanzas, author, composer, tune, copyright }` populated like the Python scraper output

- [ ] **Step 1: Confirm original hymn inputs**

Read `functions.py` lines 103-200 and `slide_making.py` lines 123-132. Confirm that the original Python flow passed both `url1` for hymn lyrics and `url2` for tune details.

- [ ] **Step 2: Write a failing Hymnary parser fixture test**

Create an inline HTML fixture in `crates/deck-builder/tests/build_deck.rs` with the same relevant labels used by Hymnary:

```html
<span class="hy_infoLabel">Title:</span><span>Amazing Grace</span>
<span class="hy_infoLabel">Author:</span><span>John Newton</span>
<span class="hy_infoLabel">Copyright:</span><span>Public Domain</span>
<div id="at_fulltext" class="authority_section"><div><div class="authority_columns"><p>1 Amazing grace</p></div></div></div>
```

Assert `parse_hymnary_page` returns `title == "Amazing Grace"` instead of `"Unable to find"`.

- [ ] **Step 3: Write a failing tune metadata test**

Add a tune fixture with:

```html
<span class="hy_infoLabel">Title:</span><span>NEW BRITAIN</span>
<span class="hy_infoLabel">Composer:</span><span>Unknown</span>
<span class="hy_infoLabel">Meter:</span><span>C.M.</span>
```

Assert the scraper extracts tune title, composer, and meter.

- [ ] **Step 4: Run focused tests and verify they fail**

Run:

```bash
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test -p deck-builder hymnary
```

Expected: FAIL because the current parser does not reliably find sibling metadata and there is no tune URL input path.

- [ ] **Step 5: Implement minimal scraper fixes**

Fix label-sibling traversal in `parse_hymnary_page`. Add a tune parsing function mirroring the original Python `tune_details(url2)` behavior.

- [ ] **Step 6: Add a tune field to the request model and frontend**

Extend `Component::Hymn` so it can carry both the lyrics URL and tune URL. Update the server's vanilla JS form so hymn rows can provide both values without breaking other component types.

- [ ] **Step 7: Verify**

Run:

```bash
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test -p deck-builder hymnary
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test -p server
```

Expected: all focused tests pass.

---

### Task 3: Psalm And Hymn Page Marker Parity

**Files:**
- Modify: `crates/deck-builder/src/lib.rs`
- Test: `crates/deck-builder/tests/build_deck.rs`

**Interfaces:**
- Consumes: generated hymn and psalm slide text
- Produces: hymn and psalm slides that preserve the original `pg. X` marker where the original design expects it

- [ ] **Step 1: Confirm original page-marker behavior**

Search the Python source for `pg. X`, page number handling, and any placeholder convention in `template.pptx`. Confirm whether `pg. X` belongs on hymn slides, psalm slides, reading slides, or all of them.

- [ ] **Step 2: Write failing deck text tests**

Add tests in `crates/deck-builder/tests/build_deck.rs` that build one psalm and one hymn using `MockSources`. Assert the generated slide text includes `pg. X` exactly where Step 1 says it should appear.

- [ ] **Step 3: Run focused tests and verify failure**

Run:

```bash
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test -p deck-builder page_marker
```

Expected: FAIL because the current Rust builder does not add page markers for hymns or psalms.

- [ ] **Step 4: Implement minimal page-marker placement**

Update the hymn and/or psalm slide writer logic in `build_deck` to add `pg. X` to the same placeholder/body text location used by the original design.

- [ ] **Step 5: Verify**

Run:

```bash
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test -p deck-builder page_marker
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test --workspace
```

Expected: all tests pass.

---

### Task 4: Tune Formatting In Generated Slides

**Files:**
- Modify: `crates/deck-builder/src/lib.rs`
- Test: `crates/deck-builder/tests/build_deck.rs`

**Interfaces:**
- Consumes: `Hymn::tune`, `Hymn::composer`, `Psalm::meter`, and parsed tune metadata
- Produces: clean, human-readable copyright/tune lines matching the original Python deck format

- [ ] **Step 1: Capture current malformed formatting**

Use the running web app or `cargo run -p deck-builder --example demo` to generate a deck. Extract slide text and record the exact tune formatting that looks wrong.

- [ ] **Step 2: Write a failing formatting test**

Add a test in `crates/deck-builder/tests/build_deck.rs` that builds a hymn with:

```rust
Hymn {
    title: "Amazing Grace".to_string(),
    stanzas: vec!["Amazing grace".to_string()],
    author: "John Newton".to_string(),
    composer: "Unknown".to_string(),
    tune: "NEW BRITAIN".to_string(),
    copyright: "Public Domain".to_string(),
}
```

Assert the generated copyright placeholder includes clean lines for `Words:`, `Composer:`, `Tune:`, copyright, and CCLI.

- [ ] **Step 3: Run focused test and verify failure**

Run:

```bash
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test -p deck-builder tune_formatting
```

Expected: FAIL until tune/copyright formatting is normalized.

- [ ] **Step 4: Implement minimal formatting normalization**

Update only the string formatting in the hymn and psalm branches of `build_deck`. Do not change scraper behavior in this task.

- [ ] **Step 5: Verify**

Run:

```bash
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test -p deck-builder tune_formatting
docker run --rm -v "$PWD":/app -w /app rust:1-slim cargo test --workspace
```

Expected: all tests pass.

---

## Self-Review

- Spec coverage: Covers reading reference behavior, hymn metadata, tune input, page markers, and tune formatting.
- Placeholder scan: No TBD/TODO placeholders; each task has concrete files, tests, and commands.
- Type consistency: Uses current `Component`, `Hymn`, `Psalm`, `Sources`, and `build_deck` names from the Rust workspace.
