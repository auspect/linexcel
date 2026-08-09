# Workbook context & screenshots

## Workbook context

`result.workbook_context` extracts bounded first rows and columns for every
sheet, without assuming a header row. It also exposes comments, merged cells,
frozen panes, hidden columns, and sheet visibility using `openpyxl`; Excel is
not launched.

```python
ctx = result.workbook_context
for sheet in ctx["sheets"]:
    print(sheet["name"], sheet["dimensions"], sheet["preview_range"])
    print("  frozen:", sheet["freeze_panes"], "hidden:", sheet["hidden_columns"])
    for comment in sheet["comments"]:
        print(f"  {comment['cell']} ({comment['author']}): {comment['text']}")
```

The mapping is keyed `filename`, `sheets`, `stats` and `warnings`; per-sheet
context lives under `sheets`.

This context has two consumers. The HTML report renders it in the **Sheets**
tab, and [`document_workbook()`](ai.md#workbook-overview) sends it to the model
so the AI overview describes the file a reader opens — its titles, labels,
comments and hidden columns — rather than only the graph its formulas make.
Pass `include_context=False` to keep cell contents off the wire.

## Screenshots

Generate and embed high-resolution sheet screenshots using LibreOffice Calc:

```python
screenshots = result.save_screenshots("screenshots/")
result.save_html("out.html", screenshots=screenshots)
```

Each sheet is rendered whole, onto a single image, and the result is keyed by
sheet name — which is what lets the **Sheets** tab show every sheet's image
beside its comments, frozen panes and first cells:

```python
result.save_screenshots("screenshots/")
# {'Ventes': [PosixPath('screenshots/demo-Ventes.png')],
#  'Synthese': [PosixPath('screenshots/demo-Synthese.png')], ...}
```

Pass `per_sheet=False` for the flat `list[Path]` of print pages instead, laid
out by the workbook's own page setup; the report then shows them in a separate
**Visual preview** tab, since no page can be tied to a sheet. That flat list is
also what you get back when the renderer does not produce exactly one page per
sheet — an older LibreOffice ignores the single-page option — because a
screenshot filed under the wrong sheet is worse than one filed under none.

Requires LibreOffice and Poppler's `pdftoppm`. Both are located on `PATH` or in
their standard install directory, so the Windows and macOS installers work
without any `PATH` setup:

=== "Debian / Ubuntu"

    ```bash
    sudo apt install libreoffice-calc poppler-utils
    ```

=== "Windows"

    ```powershell
    winget install TheDocumentFoundation.LibreOffice
    winget install oschwartz10612.Poppler
    ```

=== "macOS"

    ```bash
    brew install --cask libreoffice
    brew install poppler
    ```

LibreOffice runs headless with a throwaway user profile, so rendering works even
while LibreOffice is open on the desktop, and never touches your own settings.

!!! note "Screenshots are for readers, not for the model"

    The PNGs are embedded in the report for a human to look at. The AI overview
    never receives an image: it is given the same facts in text form, read
    deterministically from the file by `openpyxl`. So the overview describes the
    sheets as they look even when no renderer is installed at all, and no vision
    model is needed.
