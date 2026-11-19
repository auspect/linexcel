# Visual verification and independent AI review policy

## Rule: always validate visually with an independent AI reviewer

Whenever modifications touch analysis, evaluation, documentation, screenshots, or the HTML/CSS/JS viewer:

1. **Mandatory visual inspection across screen sizes:**
   - Never rely on unit tests, green exit codes, or headless pass statuses alone for UI/viewer changes.
   - Always launch Playwright/Chromium/Edge browser audits capturing both **desktop (1440×900)** and **mobile (390×844)** viewports.
   - Every available dashboard tab (**Graph, Overview, Sheets, Visual preview**) and every sheet in the sidebar must be traversed and captured.

2. **Independent AI reviewer verification:**
   - Inspect rendered components for layout anomalies:
     - Check element bounding boxes (e.g. SVG icons, warning notices, cards, grid tables) to ensure no unbounded stretching or viewport overflow occurs.
     - Verify warning scoping: notices specific to one sheet or entity must never leak onto unrelated sheets.
     - Verify scroll position handling: navigation between items must reset viewports appropriately.
   - Cross-examine AI documentation and image descriptions against actual spreadsheet screenshots and source formulas. Identify and document factual hallucinations or misleading claims.
   - Record findings, discrepancies, and remaining limits honestly in a dedicated independent review report artifact before finalizing any PR.
