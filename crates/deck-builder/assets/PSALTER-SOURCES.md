Psalter texts come from `alastairherd/psalms_website` at revision
`0c2ff87293fda215b75f6599e7bab983d7470df4`.

- `psalms.json`: the Sing Psalms entries (versions a/b/c) in `static/psalms.json`.
  The existing embedded file matches these entries exactly.
- `scottish-psalms.json`: `static/1650-psalms.json`, with version `1650`
  normalised to `a` and `1650b` to `b`. Each Body key is prefixed to its
  stanza text for the existing optional superscript-number rendering.
  Text and stanza boundaries otherwise match the source. The source's keys
  identify stanzas, which do not always correspond one-to-one to Bible verses.

Both datasets are embedded; loading and generating psalms needs no network access.
The editor's `psalmVariants` catalogue is derived from the non-a versions in
these two files and should be updated alongside the datasets.
