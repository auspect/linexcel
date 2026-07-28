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

## Screenshots

Generate and embed high-resolution sheet screenshots using LibreOffice Calc:

```python
screenshots = result.save_screenshots("screenshots/")
result.save_html("out.html", screenshots=screenshots)
```

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
