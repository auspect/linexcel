# Workbook context & screenshots

## Workbook context

`result.workbook_context` extracts bounded first rows and columns for every
sheet, without assuming a header row. It also exposes comments, merged cells,
frozen panes, hidden columns, and sheet visibility using `openpyxl`; Excel is
not launched.

```python
ctx = result.workbook_context
for sheet_name, sheet_ctx in ctx.items():
    print(f"{sheet_name}: {sheet_ctx}")
```

## Screenshots

Generate and embed high-resolution sheet screenshots using LibreOffice Calc:

```python
screenshots = result.save_screenshots("screenshots/")
result.save_html("out.html", screenshots=screenshots)
```

Requires LibreOffice and Poppler's `pdftoppm`:
```bash
sudo apt install libreoffice-calc poppler-utils
```
