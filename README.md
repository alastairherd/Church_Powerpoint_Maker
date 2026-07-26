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
```

The last two are enforced by CI; `npm test` is not, so run it locally when you
touch anything in `crates/server/static/`.

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
first.

## Conventions

All congregation-facing text uses British English; scripture fetched from the ESV
API is converted by `textproc::british_spellings`. Commit messages are an
imperative sentence with no prefix, body wrapped at about 72 characters.

`CLAUDE.md` is the working guide for LLM assistants; `PRODUCT.md` and `DESIGN.md`
cover the product intent and interface design.
