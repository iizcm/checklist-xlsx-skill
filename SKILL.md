---
name: checklist-xlsx
description: Generate formatted Excel checklists with color-coded sections, priority levels, input cells, auto-filters, and frozen headers using openpyxl.
version: 1.0.0
tags: [excel, checklist, xlsx, openpyxl, productivity]
---

# Checklist XLSX Generator

Generate professional, fillable checklists as .xlsx files from structured data. 

## When to Use

User asks for a checklist, packing list, preparation list, form, or any tabular "tick-off" document in Excel format. Also triggers when user says "buatin checklist", "checklist lahiran", "packing list", etc.

## Pattern

Use `openpyxl` to generate the file directly from Python code. Then serve via HTTP download link if user has file-delivery preference (`deliver-via-download-link` skill).

## Required Styling (Always Apply)

1. **Title row** — bold white text on blue header fill (`2F75B5`)
2. **Subtitle row** — italic gray font, shows creation date + instruction
3. **Column headers** — dark gray (`333333`) background, white text
4. **Section dividers** — colored header row per category (red/blue/green/orange/purple fills)
5. **Input columns** — yellow fill (`FFF2CC`) for user-editable fields
6. **Priority column** — CRITICAL (red), HIGH (orange), NORMAL (gray)
7. **Legend row** — at bottom explaining fill colors and priority levels
8. **Font** — Arial throughout; bold for headers, italic+gray for notes
9. **Row height** — ~22px minimum
10. **Freeze panes** — `ws.freeze_panes = 'A4'`
11. **Auto-filter** — `ws.auto_filter.ref = 'A3:F<N>'`
12. **Text wrap** — enable on description/note columns

## Standard Column Layout

| Col | Header | Width | Purpose |
|-----|--------|-------|---------|
| A | No | 5 | Item number |
| B | Item | 42 | Task/item name |
| C | Siap? (v/x) | 16 | Checkbox input (yellow) |
| D | Keterangan | 32 | Pre-filled note/description |
| E | Catatan | 32 | User notes (yellow) |
| F | Priority | 16 | CRITICAL / HIGH / NORMAL |

## Data Structure (Python)

```python
sections = [
    ("SECTION TITLE", PatternFill(start_color='COLOR', end_color='COLOR', fill_type='solid'), [
        ("Item Name", "Note description", "PRIORITY"),
        ...
    ]),
]
```

Colors per section: `C00000` (red), `5B9BD5` (blue), `70AD47` (green), `ED7D31` (orange), `7030A0` (purple)

## Delivery

After generating:
1. Test with `curl -s -o /dev/null -w "%{http_code}" http://127.0.0.1:8000/FILE`
2. Return download link: `http://134.199.170.183:8000/filename.xlsx`
3. If no download server running, start one first.

## Pitfalls

- Don't use LibreOffice recalc.py unless file has formulas (checklists don't need it)
- Don't hardcode numeric values in formula cells — this is a fillable form, not a calculation sheet
- Always test the HTTP server returns 200 before giving the link
- Write script to `/tmp/` or task-specific temp dir, execute, then clean up
