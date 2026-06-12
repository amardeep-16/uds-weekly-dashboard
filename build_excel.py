#!/usr/bin/env python3
"""build_excel.py — data.js → UDS_Dashboard.xlsx (openpyxl)"""

import json, os, re, math

# Load data.js via execjs-free approach (exec the JS assignment in Python)
def load_data_js(path):
    with open(path, "r", encoding="utf-8") as f:
        js = f.read()
    # Strip "window.DASHBOARD = " prefix and trailing ";"
    json_str = re.sub(r"^window\.DASHBOARD\s*=\s*", "", js.strip())
    json_str = json_str.rstrip(";").strip()
    return json.loads(json_str)


import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, DoughnutChart, Reference

# ─── Palette ──────────────────────────────────────────────────────────────
ACCENT_COLORS = [
    "3D7EFF","6C63FF","2DD4BF","F87171","FBBF24",
    "A78BFA","34D399","FB923C","60A5FA","E879F9",
]
NAVY   = "0F1B2E"
WHITE  = "FFFFFF"
CARD   = "1E2435"
MUTED  = "8892B0"
ACCENT = "3D7EFF"
GOOD   = "2DD4BF"
BAD    = "F87171"
LIGHT_BLUE = "E8F0FE"
BORDER_COL = "2A3050"

TAB_COLORS = [
    "3D7EFF","6C63FF","2DD4BF","F87171","FBBF24",
    "A78BFA","34D399","FB923C","60A5FA","E879F9",
    "4ADE80","F472B6",
]

# ─── Style helpers ─────────────────────────────────────────────────────────
def _font(bold=False, size=11, color=WHITE, italic=False):
    return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)

def _fill(color):
    return PatternFill("solid", fgColor=color)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border(style="thin", color=BORDER_COL):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _thick_left(color=ACCENT):
    return Border(left=Side(style="thick", color=color),
                  right=Side(style="thin", color=BORDER_COL),
                  top=Side(style="thin", color=BORDER_COL),
                  bottom=Side(style="thin", color=BORDER_COL))

def set_cell(ws, row, col, value, bold=False, size=11, color=WHITE,
             fill=None, align_h="left", align_v="center", wrap=False,
             border=None, nf=None, italic=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _font(bold, size, color, italic)
    c.alignment = _align(align_h, align_v, wrap)
    if fill:
        c.fill = _fill(fill)
    if border:
        c.border = border
    if nf:
        c.number_format = nf
    return c

def merge_range(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

# ─── Number formats ────────────────────────────────────────────────────────
NF = {
    "pct2": "0.00%",
    "pct1": "0.0%",
    "num2": "0.00",
    "int":  "#,##0",
    "pp":   '+0.00" pp";[Red]-0.00" pp";0" pp"',
    "t":    "@",
}

def nf_for(f):
    return NF.get(f, "General")

def write_val(ws, row, col, raw, fmt_key, *, bold=False, fill_color=None,
              font_color=WHITE, border=None):
    """Write a value with correct number format. Never writes #... error strings."""
    val = raw
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        val = None
    # Scale pp: stored as fraction, Excel format multiplies by 100 via custom fmt
    if fmt_key == "pp" and val is not None:
        val = val * 100  # store as plain number; format adds " pp"
    c = ws.cell(row=row, column=col, value=val)
    c.number_format = nf_for(fmt_key)
    c.font = _font(bold=bold, color=font_color)
    c.alignment = _align("right", "center")
    if fill_color:
        c.fill = _fill(fill_color)
    if border:
        c.border = border
    return c

def sanitize_name(name, max_len=31):
    clean = re.sub(r'[\[\]:*?/\\]', '', name)
    return clean[:max_len]

# ─── KPI card (3×3 merged block) ──────────────────────────────────────────
def kpi_card(ws, row, col, kpi):
    """Draw a 3-row × 3-col KPI card starting at (row, col)."""
    # Row 0: small caps label (merge across 3 cols)
    merge_range(ws, row, col, row, col+2)
    lbl = ws.cell(row=row, column=col, value=kpi.get("label",""))
    lbl.font = Font(name="Calibri", size=9, color=MUTED, bold=False)
    lbl.alignment = _align("center", "center")
    lbl.fill = _fill("111827")

    # Row 1: big value (merge across 3 cols)
    merge_range(ws, row+1, col, row+1, col+2)
    raw = kpi.get("value")
    fmt_k = kpi.get("fmt","num2")
    if fmt_k == "pp" and raw is not None:
        raw = raw * 100
    val_c = ws.cell(row=row+1, column=col, value=raw)
    val_c.number_format = NF.get(fmt_k, "General")
    val_c.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    val_c.alignment = _align("center", "center")
    val_c.fill = _fill(CARD)

    # Row 2: delta / text line (merge across 3 cols)
    merge_range(ws, row+2, col, row+2, col+2)
    d = kpi.get("delta")
    df = kpi.get("deltafmt","pp")
    d_text = ""
    if d is not None:
        sign = "+" if d >= 0 else ""
        if df == "pp":
            d_text = f"{sign}{d*100:.2f} pp"
        elif df == "pct2":
            d_text = f"{sign}{d*100:.2f}%"
        elif df in ("num2","num"):
            d_text = f"{sign}{d:.2f}"
        else:
            d_text = f"{sign}{d:.2f}"
    bottom_text = d_text or kpi.get("text","") or ""
    d_c = ws.cell(row=row+2, column=col, value=bottom_text)
    goodlow = kpi.get("goodlow", True)
    good = (d is not None and ((goodlow and d < 0) or (not goodlow and d > 0)))
    d_c.font = Font(name="Calibri", size=10, color=GOOD if good else (BAD if d else MUTED))
    d_c.alignment = _align("center", "center")
    d_c.fill = _fill(CARD)

    # Outer border on each row's merged span top-left cell
    for r in (row, row+1, row+2):
        ws.cell(row=r, column=col).border = _border(color=BORDER_COL)

# ─── Native chart builder ─────────────────────────────────────────────────
def build_native_chart(ws, chart_spec, data_start_row, data_start_col,
                        chart_row, chart_col, chart_w=10, chart_h=15):
    """
    Write chart data to hidden columns (AB+), create an openpyxl chart.
    Returns the chart object.
    """
    fmt_k = chart_spec.get("fmt", "num2")
    ctype = chart_spec.get("type", "bar")
    is_h  = chart_spec.get("horizontal", False)
    is_don= ctype == "doughnut"
    is_line = ctype == "line"

    labels   = chart_spec.get("labels", [])
    datasets = chart_spec.get("datasets", [])

    # Write hidden data block starting at data_start_col
    dc = data_start_col
    # Row 1: header
    ws.cell(row=data_start_row, column=dc, value="Category")
    for j, ds in enumerate(datasets):
        ws.cell(row=data_start_row, column=dc+1+j, value=ds.get("label",""))

    # Rows 2+: data
    for i, lbl in enumerate(labels):
        ws.cell(row=data_start_row+1+i, column=dc, value=lbl)
        for j, ds in enumerate(datasets):
            data_list = ds.get("data", [])
            val = data_list[i] if i < len(data_list) else None
            if val is not None and isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
                val = None
            if fmt_k == "pp" and val is not None:
                val = val * 100
            cell = ws.cell(row=data_start_row+1+i, column=dc+1+j, value=val)
            cell.number_format = NF.get(fmt_k, "General")

    n_rows = len(labels)
    n_ds   = len(datasets)

    # Create chart
    if is_don:
        ch = DoughnutChart()
    elif is_line:
        ch = LineChart()
        ch.grouping = "standard"
    else:
        ch = BarChart()
        if is_h:
            ch.barDir = "bar"
        else:
            ch.barDir = "col"
        ch.grouping = "clustered"

    ch.title = chart_spec.get("title","")
    ch.style = 10

    y_nf = NF.get(fmt_k, "General")

    for j in range(n_ds):
        data_ref = Reference(ws,
                             min_col=dc+1+j, max_col=dc+1+j,
                             min_row=data_start_row, max_row=data_start_row+n_rows)
        ch.add_data(data_ref, titles_from_data=True)

        # Apply series color
        if j < len(ch.series):
            ser = ch.series[j]
            col_hex = ACCENT_COLORS[j % len(ACCENT_COLORS)]
            from openpyxl.chart.data_source import NumDataSource, NumRef
            try:
                from openpyxl.drawing.fill import PatternFillProperties
                from openpyxl.chart.series import SeriesLabel
                ser.graphicalProperties.solidFill = col_hex
                if not is_don:
                    ser.graphicalProperties.line.solidFill = col_hex
            except Exception:
                pass

    # Category axis reference
    if n_rows > 0:
        cat_ref = Reference(ws,
                            min_col=dc, max_col=dc,
                            min_row=data_start_row+1, max_row=data_start_row+n_rows)
        try:
            ch.set_categories(cat_ref)
        except Exception:
            pass

    # Axis formatting
    try:
        if not is_don:
            if hasattr(ch, 'y_axis'):
                ch.y_axis.numFmt = y_nf
                ch.y_axis.delete = False
            if hasattr(ch, 'x_axis'):
                if is_h:
                    ch.x_axis.numFmt = y_nf
                else:
                    ch.x_axis.delete = False
    except Exception:
        pass

    # Legend: hide for single-series non-doughnut
    if n_ds <= 1 and not is_don:
        ch.legend = None

    ch.width  = chart_w * 1.5
    ch.height = chart_h * 0.5

    ws.add_chart(ch, f"{get_column_letter(chart_col)}{chart_row}")
    return ch

# ─── Sheet builder ────────────────────────────────────────────────────────
def build_sheet(wb, sheet_data, tab_idx):
    name = sanitize_name(sheet_data.get("name", f"Sheet{tab_idx}"))
    ws = wb.create_sheet(title=name)
    ws.sheet_view.showGridLines = False
    ws.tab_color = TAB_COLORS[tab_idx % len(TAB_COLORS)]

    kpis   = sheet_data.get("kpis",   [])
    charts = sheet_data.get("charts", [])
    tables = sheet_data.get("tables", [])

    # ── Title band ─────────────────────────────────
    ws.row_dimensions[1].height = 30
    ws.column_dimensions["A"].width = 20
    merge_range(ws, 1, 1, 1, 14)
    c = ws.cell(row=1, column=1, value=sheet_data.get("name",""))
    c.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
    c.fill = _fill(NAVY)
    c.alignment = _align("left","center")

    # subtitle
    merge_range(ws, 2, 1, 2, 14)
    c2 = ws.cell(row=2, column=1, value=sheet_data.get("subtitle",""))
    c2.font = Font(name="Calibri", size=10, color=MUTED, italic=True)
    c2.fill = _fill("0d1424")
    c2.alignment = _align("left","center")

    # ── Narrative ──────────────────────────────────
    ws.row_dimensions[3].height = 14
    ws.row_dimensions[4].height = 40
    merge_range(ws, 4, 1, 4, 14)
    narr = ws.cell(row=4, column=1, value=sheet_data.get("narrative",""))
    narr.font = Font(name="Calibri", size=10, color="C8D6F0", italic=True)
    narr.fill = _fill("1a2540")
    narr.alignment = _align("left","center", wrap=True)
    narr.border = _thick_left(ACCENT)

    # ── KPI cards (4-up) ──────────────────────────
    row_kpi = 5
    for i, kpi in enumerate(kpis[:4]):
        kpi_card(ws, row_kpi, 1 + i*3, kpi)

    ws.row_dimensions[row_kpi].height   = 16
    ws.row_dimensions[row_kpi+1].height = 30
    ws.row_dimensions[row_kpi+2].height = 18
    for col_w in range(1, 13):
        ws.column_dimensions[get_column_letter(col_w)].width = 14

    # Freeze below KPIs
    ws.freeze_panes = ws.cell(row=row_kpi+3, column=1)

    # ── Charts (2-up grid) ────────────────────────
    # Hidden data starts at col 28 (AB)
    hidden_col_start = 28
    hc = hidden_col_start

    chart_row_start = row_kpi + 4
    crow = chart_row_start

    for ci, chart_spec in enumerate(charts):
        col_pos = 1 if ci % 2 == 0 else 8
        if ci % 2 == 0 and ci > 0:
            crow += 16

        build_native_chart(ws, chart_spec,
                           data_start_row=1, data_start_col=hc,
                           chart_row=crow, chart_col=col_pos,
                           chart_w=9, chart_h=14)
        hc += len(chart_spec.get("datasets",[])) + 3 + len(chart_spec.get("labels",[]))

    # ── Tables ────────────────────────────────────
    if charts:
        tbl_row = crow + 17
    else:
        tbl_row = row_kpi + 5

    for tbl in tables:
        headers = tbl.get("headers",[])
        rows    = tbl.get("rows",[])
        fmts    = tbl.get("fmt",[])
        title   = tbl.get("title","")

        # Table title
        merge_range(ws, tbl_row, 1, tbl_row, max(len(headers), 1))
        tc = ws.cell(row=tbl_row, column=1, value=title)
        tc.font  = Font(name="Calibri", size=10, bold=True, color=WHITE)
        tc.fill  = _fill(NAVY)
        tc.alignment = _align("left","center")
        tbl_row += 1

        # Header row
        for ci, h in enumerate(headers):
            c_h = ws.cell(row=tbl_row, column=ci+1, value=h)
            c_h.font  = Font(name="Calibri", size=9, bold=True, color=WHITE)
            c_h.fill  = _fill(NAVY)
            c_h.alignment = _align("center","center")
            c_h.border = _border(color="1a2540")
        tbl_row += 1

        # Data rows
        for ri, row_data in enumerate(rows):
            is_gt = (isinstance(row_data[0], str) and
                     "grand total" in row_data[0].lower()) if row_data else False
            banded_fill = "141c2e" if ri % 2 == 0 else "111827"
            row_fill = "162040" if is_gt else banded_fill

            for ci, val in enumerate(row_data):
                f = fmts[ci] if ci < len(fmts) else "t"
                c_d = ws.cell(row=tbl_row, column=ci+1)
                if isinstance(val, str) or f == "t":
                    c_d.value = val
                    c_d.number_format = "@"
                    c_d.alignment = _align("left","center")
                else:
                    if f == "pp" and val is not None:
                        val = val * 100
                    c_d.value = val
                    c_d.number_format = NF.get(f,"General")
                    c_d.alignment = _align("right","center")
                c_d.font  = Font(name="Calibri", size=9, bold=is_gt,
                                  color=ACCENT if is_gt else WHITE)
                c_d.fill  = _fill(row_fill)
                c_d.border = _border(style="hair")
            tbl_row += 1

        tbl_row += 2  # gap between tables

    return ws


# ─── Dashboard index tab ──────────────────────────────────────────────────
def build_index(wb, sheets_data, meta):
    ws = wb.create_sheet(title="Dashboard", index=0)
    ws.sheet_view.showGridLines = False
    ws.tab_color = "0F1B2E"

    # Title
    merge_range(ws, 1, 1, 1, 10)
    c = ws.cell(row=1, column=1, value=meta.get("title","UDS Dashboard"))
    c.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    c.fill = _fill(NAVY)
    c.alignment = _align("center","center")

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 20
    merge_range(ws, 2, 1, 2, 10)
    sub = ws.cell(row=2, column=1, value=f"Source: {meta.get('source','')} | Generated: {meta.get('generated','')}")
    sub.font = Font(name="Calibri", size=10, color=MUTED)
    sub.fill = _fill("0d1424")
    sub.alignment = _align("center","center")

    # Index
    ws.row_dimensions[3].height = 10
    hdr = ws.cell(row=4, column=1, value="Sheet")
    hdr.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    hdr.fill = _fill(NAVY)
    hdr2 = ws.cell(row=4, column=2, value="Subtitle")
    hdr2.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    hdr2.fill = _fill(NAVY)
    hdr3 = ws.cell(row=4, column=3, value="Latest Week")
    hdr3.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    hdr3.fill = _fill(NAVY)

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 16

    for i, s in enumerate(sheets_data):
        row = 5 + i
        name = sanitize_name(s.get("name",""))
        c_link = ws.cell(row=row, column=1, value=name)
        c_link.hyperlink = f"#'{name}'!A1"
        c_link.font = Font(name="Calibri", size=10, color=ACCENT, underline="single")
        c_link.fill = _fill("141c2e" if i%2==0 else "111827")
        c_sub = ws.cell(row=row, column=2, value=s.get("subtitle",""))
        c_sub.font = Font(name="Calibri", size=10, color=MUTED)
        c_sub.fill = _fill("141c2e" if i%2==0 else "111827")
        c_lw = ws.cell(row=row, column=3, value=s.get("latest_week",""))
        c_lw.font = Font(name="Calibri", size=10, color=WHITE)
        c_lw.fill = _fill("141c2e" if i%2==0 else "111827")

    return ws


# ─── Main ──────────────────────────────────────────────────────────────────
def main():
    here = os.path.dirname(os.path.abspath(__file__))
    data_js = os.path.join(here, "data.js")
    out_path = os.path.join(here, "UDS_Dashboard.xlsx")

    if not os.path.exists(data_js):
        raise FileNotFoundError(f"data.js not found at {data_js}. Run build_data.py first.")

    print(f"Loading {data_js} …")
    D = load_data_js(data_js)

    wb = openpyxl.Workbook()
    # Remove default sheet
    for s in wb.sheetnames:
        del wb[s]

    meta = D.get("meta", {})
    sheets = D.get("sheets", [])

    print(f"Building Dashboard index …")
    build_index(wb, sheets, meta)

    for i, s in enumerate(sheets):
        print(f"  Building sheet: {s.get('name','')}")
        build_sheet(wb, s, i)

    wb.save(out_path)
    print(f"\nOK UDS_Dashboard.xlsx written -> {out_path}")
    print(f"  {len(wb.sheetnames)} tabs")


if __name__ == "__main__":
    main()
