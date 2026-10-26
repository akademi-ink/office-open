# office-open (Akademi)

Akademi's OOXML engine, focused on Word (.docx). It is a fork of
[DemoMacro/office-open](https://github.com/DemoMacro/office-open) at v0.14.5 and no longer
tracks upstream. The upstream main as of 2026-10-05 is kept on branch
`archive/upstream-2026-10-05`.

| Package             | Role                                           |
| ------------------- | ---------------------------------------------- |
| `@office-open/xml`  | XML parsing and serialization                  |
| `@office-open/core` | OPC packaging, drawings, charts, themes, units |
| `@office-open/docx` | Word generation, parsing and patching          |
| `@office-open/pptx` | PowerPoint export (kept for Akademi decks)     |

The packages are consumed from source by
[akademi-editor](https://github.com/akademi-ink/akademi-editor) as a git subtree in
`office-open/`. Fix bugs here or in that subtree, then sync with `git subtree push/pull`.
Every change is measured against Word with Akademi's fidelity harness before it lands.

MIT licensed; see [LICENSE](./LICENSE).
