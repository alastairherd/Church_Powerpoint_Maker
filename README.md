# TWPC Service Builder

A web app for assembling a church order of service and generating the PowerPoint
deck for the Sunday screens.

Staff pick a preset (morning, traditional evening, praise and worship, or either
Lord's Supper order), then fill in the week's specifics: hymns from the song
library, a psalm, scripture readings, notices, a catechism or confession reading,
and the sermon details. Liturgy that does not change week to week — the
confession, the assurance of forgiveness, the Lord's Prayer, the creed — is filled
in automatically. Pressing Generate produces a `.pptx` built from the
TWPC-branded template, ready to open and project.

## Editor and generation behaviour

The editor autosaves without rebuilding the whole page on every keystroke. Counts
and validation are batched, component fields keep their focus while staff type,
and drag-and-drop ordering avoids unnecessary DOM work. Psalm, scripture and
teaching lookups are read-only until the corresponding Load button is pressed;
the app may prefetch them when a staff member focuses or points at that action so
the explicit load feels faster. Shared prefetched requests still retain the
normal ten-second loader timeout.

The song picker loads the complete matching set into a bounded, scrollable list.
There is no hidden fourteen-song result limit.

Psalm body text is authored at 32pt. Automatic Psalm pagination greedily keeps
whole source stanzas together according to the available rendered height. Two
typical stanzas are the design target, but three or more short stanzas may share
a slide when they fit. A stanza rolls intact to the next slide when adding it
would overflow; an individually oversized stanza is kept whole so no text is
silently lost. Staff-edited slide breaks remain authoritative.

### Background deck preparation

Full-deck background preparation is experimental and disabled by default. It was
found to compete with foreground Generate on CPU-constrained hosting, because
PPTX assembly and ZIP compression are CPU-heavy. Safe read-only source
prefetching remains enabled, but production uses the original single foreground
deck build unless `BACKGROUND_DECK_PREPARATION=true` is explicitly configured.

When explicitly enabled, an autosave that remains unchanged for 3.5 seconds may
prepare the exact saved revision in the background:

- preparation runs on one dedicated worker and uses a bounded queue;
- completed decks are cached in memory for up to 20 minutes, with a maximum of
  four decks or 64 MiB;
- cache identity includes the service ID and revision, global settings version,
  and embedded-template generation;
- a later edit or settings change invalidates the corresponding prepared work;
- preparation never marks a service complete, publishes a deck, or writes to
  generated history;
- Generate uses an exact result only when it is already complete. It never waits
  behind the background queue; otherwise it immediately uses the normal
  foreground build.

Successful generation still performs all permanent writes in one place: it
stores the PPTX, records an immutable service snapshot in generated history, and
marks the live service complete. The response header `x-deck-preparation` is
`hit` when prepared bytes were reused and `miss` when a foreground build was
needed.

From the Generated PowerPoints page, **Use as starting point** restores the
immutable settings saved with that generated revision into a new draft. It does
not modify the original service or its generated history.

## Running it

The server needs a few things from the environment:

| Variable | Purpose |
| --- | --- |
| `ESV_API_KEY` | Fetching scripture text |
| `STAFF_PASSWORD_HASH` | The shared staff password, hashed — generate one with `cargo run --bin hash-password` |
| `SESSION_SIGNING_SECRET` | Signing session cookies; at least 32 bytes |
| `R2_ACCOUNT_ID`, `R2_BUCKET`, `R2_ACCESS_KEY_ID`, `R2_SECRET_ACCESS_KEY` | Cloudflare R2, where services and the song library are stored |
| `PORT` | Defaults to 8080 |
| `OBJECT_STORE` | Set to `memory` to run without R2; nothing is persisted |
| `COOKIE_SECURE` | Set to `false` when developing over plain HTTP |
| `BACKGROUND_DECK_PREPARATION` | Experimental full-deck preparation; defaults to `false` because it can compete for CPU on constrained hosts |

```sh
cargo run -p server
```

For a quick local look, `OBJECT_STORE=memory` avoids needing R2 credentials at
all. Deployment is by Docker; `render.yaml` describes the hosted setup.

## Building and testing

```sh
cargo test --workspace        # Rust: unit tests and HTTP endpoint tests
npm test                      # frontend: vitest under jsdom
cargo fmt --check
cargo clippy --workspace --all-targets -- -D warnings
bash scripts/validate-openxml.sh  # representative deck + Microsoft Open XML SDK
```

Formatting and Clippy are enforced by CI. `npm test` is not, so run it locally
when you touch anything in `crates/server/static/`. Run the Open XML validator
whenever presentation generation or package structure changes.

**Static assets and JSON data are compiled into the binary** with `include_str!`
and `include_bytes!`. Editing a file under `crates/server/static/`,
`crates/server/templates/` or `crates/deck-builder/assets/` has no effect on a
running server until you rebuild.

## Layout

- **`crates/deck-builder`** — the service model (presets, components) and deck
  generation. `build_deck` walks the order of service and clones seed slides from
  the branded template. Embedded data lives in `assets/`: the template itself,
  the psalter, the catechisms and confession, and the fixed liturgy wording.
- **`crates/pptx-template`** — OpenXML editing done by hand: open a package, clone
  slides, set shape text and run formatting, import slides from another deck.
  There is no PowerPoint library underneath it.
- **`crates/server`** — the axum web server. Sessions and CSRF, the JSON API, the
  askama templates, and the R2 object store. Also two small binaries:
  `hash-password` and `import-song-library`.

The frontend is vanilla ES modules with no framework and no build step.

Songs are not kept in this repository. Each one is a PowerPoint held in object
storage and imported into the generated deck slide for slide, so the lyrics
appear exactly as they were laid out. Staff add them through the song library
page.

## A warning about PowerPoint

PowerPoint enforces rules that schema validators, including the Open XML SDK,
happily ignore — slide master and layout IDs must be globally unique, every
reachable master must be registered, and each master must own its theme part.
Real PowerPoint is the only trustworthy check that a generated deck opens. If you
are changing anything in `pptx-template`, read
[`docs/powerpoint-repair-postmortem.md`](docs/powerpoint-repair-postmortem.md)
first. For presentation-generation changes, validate a representative deck with
`scripts/validate-openxml.sh` and, where PowerPoint is available, open it there
and inspect representative dense slides as the final smoke test. Running Office
inside Docker is not the expected workflow; the Open XML SDK container provides
schema validation, while a licensed desktop PowerPoint installation remains the
application-level check.

## Conventions

All congregation-facing text uses British English; scripture fetched from the ESV
API is converted by `textproc::british_spellings`. Commit messages are an
imperative sentence with no prefix, body wrapped at about 72 characters.

`CLAUDE.md` is the working guide for LLM assistants; `PRODUCT.md` and `DESIGN.md`
cover the product intent and interface design.
