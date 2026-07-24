# Postmortem: generated PPTX files PowerPoint refused to open (July 2026)

Audience: future LLM agents (and humans) working on `crates/pptx-template`.
Read this before touching slide-master import, registration, or validation code.

## Symptom

Decks generated with an imported song presentation (`Presentation::import_slides`)
failed in Microsoft PowerPoint. Progression across downloaded revisions of the
same service deck:

| Revision | State of the package | PowerPoint behaviour |
|----------|----------------------|----------------------|
| r13 | `slideMaster2.xml` present but missing from `p:sldMasterIdLst` | opened only after "repair" |
| r14 | master registered, but its `p:sldMasterId/@id` reused an existing `p:sldLayoutId/@id`, and all 11 of its layout IDs duplicated master1's | would not open |
| r15 | master ID collision fixed; the 11 duplicate layout IDs remained | would not open |
| r16 | all IDs unique; both masters' `.rels` pointed at the same `ppt/theme/theme1.xml` | would not open |
| r16 + private theme copy | — | opened cleanly |

## Root causes (three stacked defects)

1. **Unregistered reachable master.** Import copied master/layout parts but never
   added the master to `p:sldMasterIdLst` / `presentation.xml.rels`.
   Fixed in `4203460` (reachable-master traversal + registration).
2. **Duplicate master/layout IDs.** `p:sldMasterId/@id` and `p:sldLayoutId/@id`
   share one ID space and must be globally unique across *all* registered
   masters. The imported master kept its original layout IDs, which duplicated
   the destination master's. Fixed in `489bd08`
   (`normalize_slide_master_layout_ids` renumbers on registration; `r:id`
   attributes are left untouched because the relationship, not the numeric ID,
   identifies the layout part).
3. **Two masters sharing one theme part.** The song deck's theme was
   byte-identical to the template's, so import deduplication collapsed them
   onto `theme1.xml`. PowerPoint requires a one-to-one master↔theme pairing.
   Fixed in `0af924d` (theme parts excluded from import dedup; validation
   rejects registered masters sharing a theme).

Each fix exposed the next: PowerPoint stops at the first violated invariant,
so "still broken after the fix" did **not** mean the fix was wrong.

## PowerPoint invariants the Open XML SDK does NOT check

`scripts/validate-openxml.sh` (Microsoft's DocumentFormat.OpenXml validator)
passed every one of the rejected files above. It validates schema and package
structure, not PowerPoint's application-level rules. Known rules it misses:

- every master reachable from a slide must be listed in `p:sldMasterIdLst`;
- `p:sldMasterId/@id` and `p:sldLayoutId/@id` must be unique across the whole
  presentation (all values ≥ 2147483648);
- each slide master must reference its **own** theme part — never share one;
- each layout listed in a master's `p:sldLayoutIdLst` should have `.rels`
  pointing back to that same master, and appear in exactly one master's list.

The only true oracle is opening the file in real PowerPoint.

## Investigation method that worked

- **Never rebuild to diagnose.** Inspect the failing `.pptx` directly with a
  short Python script (`zipfile` + `ElementTree`/regex): dump
  `p:sldMasterIdLst`, each master's `p:sldLayoutIdLst`, resolve every `r:id`
  through the `.rels`, check ID uniqueness, backlinks, content-type overrides,
  and XML well-formedness. Compare against the known-good
  `crates/deck-builder/assets/template.pptx`.
- **Differential analysis** across the failing revisions (r13→r16) shows which
  invariant each code change fixed and which one still fails.
- **Binary-patch the failing file** to test a hypothesis before writing Rust:
  e.g. r16 with a duplicated `theme4.xml` for master2 (a ~20-line Python zip
  rewrite) opened cleanly in PowerPoint, proving the theme diagnosis. This is
  minutes instead of a full build cycle.
- **One Docker verification pass at the end**, not during diagnosis. This
  machine (Raspberry Pi 5) has no host Rust toolchain; repeated cold Docker
  builds have destabilised it. Use the warm root-owned volumes
  `church-powerpoint-target` and `church-powerpoint-cargo-registry`, run the
  container as root with the workspace bind-mounted read-only, `--cpus=2`:
  `cargo fmt --check && cargo clippy --workspace --all-targets -- -D warnings
  && cargo test --workspace` in `rust:1.97.1`.

## Where the guardrails live now

- `Presentation::validate()` (called on every save path and by
  `validate_song_source`) rejects: unregistered reachable masters, duplicate
  master/layout IDs, and masters sharing a theme part.
- Regression tests in `crates/pptx-template/tests/presentation.rs` reconstruct
  the exact r15 and r16 package shapes and assert validation rejects them
  (`validation_rejects_r15_layout_ids_overlapping_across_registered_masters`,
  `validation_rejects_registered_masters_sharing_a_theme_part`).

## Known residual risks (unfixed as of 0af924d)

- Layout import dedup compares a layout's master by *content similarity* and
  relationship *shapes* only. A source master that differs from the
  destination master only in relationship targets (e.g. a different background
  image) could dedup a layout onto the destination master's layout while still
  creating a new master — recreating a cross-master backlink violation.
  Validation does not yet check layout↔master backlinks.
- `attr(tag, "id")` uses `\bid="` which also matches `r:id="`; safe only while
  generated XML always puts `id` before `r:id`.
- `add_slide_from_layout` copies a layout's whole inner XML; if used with a
  layout containing `<p:hf>` (slideLayout7 has one), the element is invalid
  inside `p:sld`. Current callers only use layouts 12/13.
- CI (`.github/workflows/ci.yml`) does not run `scripts/validate-openxml.sh`,
  and no automated check opens the deck in real PowerPoint.
