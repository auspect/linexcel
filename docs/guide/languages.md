# Languages

`language=` drives both the AI system prompt and the viewer interface. Nine are
available:

| Code | Language | Code | Language | Code | Language |
| --- | --- | --- | --- | --- | --- |
| `en` | English (default) | `it` | Italiano | `nl` | Nederlands |
| `fr` | Français | `pt` | Português | `ja` | 日本語 |
| `es` | Español | `de` | Deutsch | `zh` | 简体中文 |

```python
docs = result.document(base_url=..., model=..., language="de")
result.save_html("out.html", docs=docs, language="de")
```

The two are independent: a German report can carry English cards if you ask for
one language when documenting and another when exporting.

## A closed allowlist, not free text

Any other value raises `ValueError`. This is a constraint rather than an
oversight: `language` selects a stored system prompt *and* is interpolated into
the generated JavaScript, so an arbitrary string would be both a
prompt-injection and an interpolation vector.

Reports embed only the requested language plus the English fallback, so the set
can grow without every exported file growing with it.

## Adding one

Add a directory under `src/linexcel/assets/prompts/`, named with the language
code, holding three Markdown files:

```
src/linexcel/assets/prompts/sv/
├── node.md        # the card written for one calculation node
├── workbook.md    # the workbook overview
└── vision.md      # describing a sheet screenshot
```

Then extend `linexcel.i18n.UI_STRINGS` with the interface strings. The test
suite asserts the four stay in sync, so a partial addition fails the build
rather than surfacing as raw interface keys in a report.

The prompts are prose, one file per language, precisely so that translating
them needs no Python: copy the `en/` directory, translate the three files,
add the interface strings.

!!! note "Translation provenance"

    English and French were written by hand. The other seven languages — the
    interface strings *and* the AI system prompts — were produced with AI
    assistance and have not been reviewed by native speakers.

    This matters more for the prompts than for the interface: their wording
    steers how the model writes each card, so an awkward phrasing degrades
    output quality rather than just looking odd. Corrections are welcome —
    interface strings live in `linexcel.i18n`, prompts under
    `src/linexcel/assets/prompts/<language>/`.
