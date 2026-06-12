# UDS Weekly Operations Dashboard

A fully automated Python pipeline that turns the weekly UDS logistics workbook (`UDS_PPT_Data_Wk_*.xlsx`) into two ready-to-use dashboards — a single-file interactive web dashboard and a native Excel workbook.

---

## What it produces

| Output | Description |
|---|---|
| `data.js` | Parsed JSON (all 12 sheets, 48 KPIs, 33 charts) injected as `window.DASHBOARD` |
| `index.html` | Dark, single-file interactive dashboard — Chart.js 4.4.1 via CDN, no build step |
| `UDS_Dashboard.xlsx` | Native Excel workbook — 13 tabs, 33 openpyxl charts, formatted KPI cards & tables |

### Dashboard sections (12 sheets)

| # | Sheet | Key metrics |
|---|---|---|
| 1 | CCI Slide | Complaint Contact Index — All / Non-Meesho / Meesho |
| 2 | NEW EI% | Escalation Index per 1 000 shipments |
| 3 | Source Wise Del & Pik | UDS, Delivered, RTO by complaint source |
| 4 | Client Wise Del & Pik | Del% and RTO% by client type |
| 5 | POD Wise CCI_Del, Pik & Fake % | Delivered%, UDS%, Fake% by pincode-OD |
| 6 | Issue Type Wise_Fwd & Rev | AWB count ranking by issue type |
| 7 | ZONE WISE | Del%, RTO%, Fake% by logistics zone |
| 8 | Zone Wise - C2 | C2 Escalation Index by zone |
| 9 | RTO & Cancelled Reasons | Top RTO and cancellation drivers |
| 10 | Action User | Top users driving cancellations & RTOs |
| 11 | POST UDS FAKE | Post-UDS fake delivery rates by POD |
| 12 | Pre & Post_Data | Pre-UDS vs Post-UDS delivery improvement |

---

## Quick start

### Requirements

```bash
pip install openpyxl
```

Python 3.10+ required. Internet is only needed when *viewing* `index.html` (Chart.js CDN).

### Run

```bash
# Drop your weekly workbook in this folder, then:
python run_all.py

# Optionally open the dashboard in your browser immediately:
python run_all.py --open
```

That's it. The newest `UDS_PPT_Data_Wk_*.xlsx` in the folder is auto-selected.

---

## File structure

```
.
├── build_data.py        # Step 1 — parse workbook → data.js
├── build_excel.py       # Step 2 — data.js → UDS_Dashboard.xlsx
├── index.html           # Static dashboard (loads data.js + Chart.js CDN)
├── run_all.py           # Orchestrator: runs both steps, prints summary
├── data.js              # Generated — window.DASHBOARD JSON
└── UDS_Dashboard.xlsx   # Generated — native Excel dashboard
```

---

## How it works

### `build_data.py`

- **Auto-detects** the newest `UDS_PPT_Data_Wk_*.xlsx` in the folder or its parent (ignores files named `Dashboard` or starting with `~$`).
- Reads with `openpyxl` in `read_only + data_only` mode (cached values, no formula evaluation).
- Two generic readers handle every sheet layout:
  - `read_block` — two-row merged header (week label row + metric sub-row), de-duplicates repeated sub-column names (`%` → `%`, `%#2`).
  - `read_simple` — single header row for ranking/reason tables.
- **Front-anchored windowing** — a `WINDOW` dict limits each sheet to its first N week-groups, keeping the curated left side and ignoring the scratch region on the right.
- Emits `window.DASHBOARD = {...}` as `data.js` (UTF-8, NaN/inf → null).

### `index.html`

- Fixed dark sidebar with one nav item per sheet; main pane with sticky topbar.
- Per section: blue-accent narrative callout, 4-up KPI grid with coloured ▲/▼ delta, 2-up responsive chart grid, formatted tables with bold Grand Total rows.
- Charts render **lazily** on first section visit — no wasted work on load.
- Formatters: `pct2` → `x.xx%`, `pct1` → `x.x%`, `num2` → `x.xx`, `int` → Indian grouping, `pp` → `+x.xx pp`.
- Horizontal bar charts use a dedicated category-axis config (no `callback: undefined` bug).
- Print stylesheet expands all sections.

### `build_excel.py`

- Loads `data.js` by stripping the JS wrapper and parsing JSON — no Node.js required.
- **Dashboard** tab: title, subtitle, clickable index linking each sheet tab (`cell.hyperlink`).
- Per sheet tab: title band, wrapped narrative in light-blue box with thick accent border, 4 KPI "cards" (3-row × 3-col merged cells: small grey label, 18pt bold value, coloured delta line), native openpyxl charts reading from hidden data blocks in columns AB+, formatted banded tables with navy header and bold Grand Total.
- Number formats: `pct2 → 0.00%`, `int → #,##0`, `pp → +0.00" pp";[Red]-0.00" pp";0" pp"` (× 100 before writing).
- Zero formula errors — `num()` guard ensures no `#REF!` / `#DIV/0!` strings reach cells.

---

## Weekly automation

Drop the new `UDS_PPT_Data_Wk_<n>.xlsx` in the folder and run:

```bash
python run_all.py
```

No code changes needed. The newest file is auto-selected; the front-anchored window automatically tracks the latest week. If a curated block genuinely widens (new permanent columns added), bump that sheet's count in the `WINDOW` dict at the top of `build_data.py` — that's the only maintenance knob.

---

## Acceptance checks

| Check | Target | Status |
|---|---|---|
| Sheets in data.js | 12 | Pass |
| KPIs per sheet | 4 | Pass |
| Charts total | 33 | Pass |
| All-null chart series | 0 | Pass |
| Tabs in xlsx | 13 | Pass |
| Formula-error cells in xlsx | 0 | Pass |
