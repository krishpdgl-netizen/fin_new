"""
main.py
-------
The ENTIRE backend in one file. No other Python files needed.

What it does:
  - Login with a shared access code
  - Save/load quarterly financial entries
  - Read/write the data as an Excel file stored in a GitHub repo
    (via the GitHub Contents API) - so there's no database to set up
  - Build the Reports table (Q1-Q4, H1, 9M, Annual + QoQ%/YoY% growth)
  - Serve a dashboard summary

Environment variables needed (set these in Vercel Project Settings):
  GITHUB_TOKEN   - GitHub Personal Access Token (Contents: Read & Write)
  GITHUB_REPO    - "your-username/your-repo"
  GITHUB_BRANCH  - usually "main"
  EXCEL_PATH     - where the .xlsx lives in the repo, e.g. "data/finance_data.xlsx"
  ACCESS_CODE    - the shared password used to log in
"""

import os
import io
import base64
from datetime import datetime, timezone
from collections import defaultdict
from typing import Dict

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, Series

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel


# ============================================================================
# SETTINGS
# ============================================================================
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO = os.environ.get("GITHUB_REPO", "")
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH", "main")
EXCEL_PATH = os.environ.get("EXCEL_PATH", "data/finance_data.xlsx")
ACCESS_CODE = os.environ.get("ACCESS_CODE", "finance2026")

GITHUB_ = "https://api.github.com"
GH_HEADERS = {
    "Authorization": f"Bearer {GITHUB_TOKEN}",
    "Accept": "application/vnd.github+json",
}

SHEET_NAME = "Entries"
HEADER_ROW = ["fiscal_year", "quarter", "category", "particular", "amount", "entered_by", "updated_at"]

QUARTER_ORDER = ["Q1", "Q2", "Q3", "Q4"]

STANDARD_PARTICULARS = {
    "Revenue": ["Revenue from Operations", "Other Income"],
    "Expenses": [
        "Cost of Materials / COGS",
        "Employee Benefit Expense",
        "Finance Costs",
        "Depreciation & Amortization",
        "Other Expenses",
    ],
    "Other": [],
}

COMPUTED_ROWS = [
    "Total Revenue",
    "Total Expenses",
    "EBITDA",
    "Profit Before Tax (PBT)",
    "Tax Expense",
    "Net Profit",
    "Net Profit Margin %",
]


# ============================================================================
# GITHUB READ / WRITE (the "database")
# ============================================================================
def _contents_url(path):
    return f"{GITHUB_}/repos/{GITHUB_REPO}/contents/{path}"


def _get_file_meta():
    """Returns (sha, raw_bytes). raw_bytes is None if the file doesn't exist yet."""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        raise RuntimeError("GITHUB_TOKEN / GITHUB_REPO are not set on the server.")
    resp = requests.get(_contents_url(EXCEL_PATH), headers=GH_HEADERS, params={"ref": GITHUB_BRANCH}, timeout=30)
    if resp.status_code == 404:
        return None, None
    resp.raise_for_status()
    data = resp.json()
    return data["sha"], base64.b64decode(data["content"])


def _blank_workbook_bytes():
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADER_ROW)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def load_rows():
    """Every saved line item, as a list of dicts."""
    _, raw = _get_file_meta()
    if raw is None:
        return []
    wb = load_workbook(io.BytesIO(raw), data_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        return []
    header = all_rows[0]
    return [dict(zip(header, r)) for r in all_rows[1:] if r[0] is not None]


def save_rows(rows, commit_message):
    """Overwrite the workbook in GitHub with the full row list."""
    sha, _ = _get_file_meta()
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADER_ROW)
    for r in rows:
        ws.append([r.get(col) for col in HEADER_ROW])
    buf = io.BytesIO()
    wb.save(buf)

    payload = {
        "message": commit_message,
        "content": base64.b64encode(buf.getvalue()).decode(),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    resp = requests.put(_contents_url(EXCEL_PATH), headers=GH_HEADERS, json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_raw_file_bytes():
    _, raw = _get_file_meta()
    return raw if raw is not None else _blank_workbook_bytes()


def get_download_url():
    return f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/{EXCEL_PATH}"


# ============================================================================
# FINANCE CALCULATIONS
# ============================================================================
def _qkey(fy, q):
    return (int(fy), q)


def group_by_period(rows):
    grouped = defaultdict(dict)
    for r in rows:
        fy, q, particular, amount = r["fiscal_year"], r["quarter"], r["particular"], r["amount"]
        if fy is None or q is None or particular is None:
            continue
        grouped[_qkey(fy, q)][particular] = float(amount or 0)
    return grouped


def compute_subtotals(values, categories):
    rev_items = [p for p in categories.get("Revenue", []) if p in values]
    exp_items = [p for p in categories.get("Expenses", []) if p in values]

    total_revenue = sum(values.get(p, 0) for p in rev_items)
    depreciation = values.get("Depreciation & Amortization", 0)
    finance_costs = values.get("Finance Costs", 0)
    tax = values.get("Tax Expense", 0)
    total_expenses = sum(values.get(p, 0) for p in exp_items)

    ebitda = total_revenue - total_expenses + depreciation + finance_costs
    pbt = total_revenue - total_expenses
    net_profit = pbt - tax
    margin = (net_profit / total_revenue * 100) if total_revenue else 0

    return {
        "Total Revenue": total_revenue,
        "Total Expenses": total_expenses,
        "EBITDA": ebitda,
        "Profit Before Tax (PBT)": pbt,
        "Tax Expense": tax,
        "Net Profit": net_profit,
        "Net Profit Margin %": margin,
    }


def _sum_periods(period_dicts):
    out = defaultdict(float)
    for d in period_dicts:
        for k, v in d.items():
            out[k] += v or 0
    return dict(out)


def build_report(rows, fiscal_year, categories=None):
    categories = categories or STANDARD_PARTICULARS
    grouped = group_by_period(rows)

    quarter_raw = {q: grouped.get(_qkey(fiscal_year, q), {}) for q in QUARTER_ORDER}
    prev_quarter_raw = {q: grouped.get(_qkey(fiscal_year - 1, q), {}) for q in QUARTER_ORDER}
    available = [q for q in QUARTER_ORDER if quarter_raw[q]]

    columns = list(available)
    cumulative = {}
    if all(q in available for q in ["Q1", "Q2"]):
        cumulative["H1 (6M)"] = _sum_periods([quarter_raw["Q1"], quarter_raw["Q2"]])
        columns.append("H1 (6M)")
    if all(q in available for q in ["Q1", "Q2", "Q3"]):
        cumulative["9M"] = _sum_periods([quarter_raw["Q1"], quarter_raw["Q2"], quarter_raw["Q3"]])
        columns.append("9M")
    if all(q in available for q in QUARTER_ORDER):
        cumulative["FY (Annual)"] = _sum_periods([quarter_raw[q] for q in QUARTER_ORDER])
        columns.append("FY (Annual)")

    all_period_values = {**{q: quarter_raw[q] for q in available}, **cumulative}

    particulars = []
    for cat in ["Revenue", "Expenses", "Other"]:
        for p in categories.get(cat, []):
            particulars.append((cat, p))
    seen = {p for _, p in particulars}
    for q in available:
        for p in quarter_raw[q]:
            if p not in seen:
                particulars.append(("Other", p))
                seen.add(p)

    line_items = []
    for cat, particular in particulars:
        if particular in COMPUTED_ROWS:
            continue
        values = {col: all_period_values.get(col, {}).get(particular, 0) for col in columns}
        line_items.append({"particular": particular, "category": cat, "values": values})

    computed = {col: compute_subtotals(all_period_values[col], categories) for col in columns}
    for name in COMPUTED_ROWS:
        line_items.append({
            "particular": name,
            "category": "Computed",
            "values": {col: computed[col][name] for col in columns},
        })

    growth = defaultdict(dict)
    quarter_full = {q: {**quarter_raw[q], **computed.get(q, {})} for q in available}
    prev_full = {}
    for q in QUARTER_ORDER:
        base = prev_quarter_raw[q]
        prev_full[q] = {**base, **compute_subtotals(base, categories)} if base else {}

    for li in line_items:
        name = li["particular"]
        for i, q in enumerate(available):
            cur = quarter_full[q].get(name, 0)
            qoq = None
            if i > 0:
                prev_val = quarter_full[available[i - 1]].get(name, 0)
                if prev_val:
                    qoq = (cur - prev_val) / abs(prev_val) * 100
            yoy = None
            prev_year_val = prev_full.get(q, {}).get(name)
            if prev_year_val:
                yoy = (cur - prev_year_val) / abs(prev_year_val) * 100
            growth[name][q] = {"qoq": qoq, "yoy": yoy}

    return {"fiscal_year": fiscal_year, "columns": columns, "line_items": line_items, "growth": growth}


def list_years(rows):
    return sorted({int(r["fiscal_year"]) for r in rows if r.get("fiscal_year") is not None}, reverse=True)


def latest_summary(rows):
    grouped = group_by_period(rows)
    if not grouped:
        return None
    fy, q = max(grouped.keys(), key=lambda k: (k[0], QUARTER_ORDER.index(k[1])))
    report = build_report(rows, fy)
    rev = next((li for li in report["line_items"] if li["particular"] == "Total Revenue"), None)
    npf = next((li for li in report["line_items"] if li["particular"] == "Net Profit"), None)
    margin = next((li for li in report["line_items"] if li["particular"] == "Net Profit Margin %"), None)
    g = report["growth"]
    return {
        "fiscal_year": fy,
        "quarter": q,
        "total_revenue": rev["values"].get(q, 0) if rev else 0,
        "net_profit": npf["values"].get(q, 0) if npf else 0,
        "net_margin": margin["values"].get(q, 0) if margin else 0,
        "qoq_revenue": g.get("Total Revenue", {}).get(q, {}).get("qoq"),
        "yoy_revenue": g.get("Total Revenue", {}).get(q, {}).get("yoy"),
        "qoq_net_profit": g.get("Net Profit", {}).get(q, {}).get("qoq"),
        "yoy_net_profit": g.get("Net Profit", {}).get(q, {}).get("yoy"),
    }

_HEADER_FILL = PatternFill("solid", fgColor="111827")
_HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
_COMPUTED_FILL = PatternFill("solid", fgColor="F3F4F6")
_COMPUTED_FONT = Font(bold=True)
_HIGHLIGHT_FONT = Font(bold=True, color="2563EB")
_UP_FILL = PatternFill("solid", fgColor="DCFCE7")
_UP_FONT = Font(color="15803D", bold=True)
_DOWN_FILL = PatternFill("solid", fgColor="FEE2E2")
_DOWN_FONT = Font(color="B91C1C", bold=True)
_FLAT_FILL = PatternFill("solid", fgColor="F3F4F6")
_FLAT_FONT = Font(color="6B7280")
_THIN_BORDER = Border(bottom=Side(style="thin", color="E5E7EB"))
_TITLE_FONT = Font(bold=True, size=14, color="111827")

_PERCENT_ROWS = {"Net Profit Margin %"}
_HIGHLIGHT_ROWS = {"Total Revenue", "Net Profit"}


def _autosize_columns(ws, num_cols, width=16, first_col_width=30):
    ws.column_dimensions["A"].width = first_col_width
    for i in range(2, num_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = width


_SECTION_AMT_FILL = PatternFill("solid", fgColor="2563EB")
_SECTION_QOQ_FILL = PatternFill("solid", fgColor="15803D")
_SECTION_YOY_FILL = PatternFill("solid", fgColor="B45309")
_SECTION_FONT = Font(bold=True, color="FFFFFF", size=10)


def _write_year_sheet(ws, report, fiscal_year):
    """
    One sheet per fiscal year with three side-by-side blocks:
    Particular | <Amounts: Q1..Q4, H1, 9M, Annual> | <QoQ%: Q1..Q4> | <YoY%: Q1..Q4>
    plus a bar chart of Total Revenue / EBITDA / Net Profit by quarter.
    """
    columns = report["columns"]                              # e.g. Q1..Q4, H1 (6M), 9M, FY (Annual)
    quarter_cols = [c for c in columns if c in QUARTER_ORDER]  # growth only applies to quarters
    n_amt = len(columns)
    n_growth = len(quarter_cols)

    amt_start = 2
    qoq_start = amt_start + n_amt + 1   # 1 blank column as a visual gap
    yoy_start = qoq_start + n_growth + 1 if n_growth else qoq_start
    total_cols = (yoy_start + n_growth - 1) if n_growth else (amt_start + n_amt - 1)

    # ---- Title ----
    ws["A1"] = f"FY {fiscal_year} — Financial Report"
    ws["A1"].font = _TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)

    # ---- Section headers (row 2) ----
    def _section(col_start, col_end, label, fill):
        if col_end < col_start:
            return
        ws.merge_cells(start_row=2, start_column=col_start, end_row=2, end_column=col_end)
        cell = ws.cell(row=2, column=col_start, value=label)
        cell.font = _SECTION_FONT
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")

    _section(amt_start, amt_start + n_amt - 1, "AMOUNTS", _SECTION_AMT_FILL)
    if n_growth:
        _section(qoq_start, qoq_start + n_growth - 1, "QoQ % GROWTH", _SECTION_QOQ_FILL)
        _section(yoy_start, yoy_start + n_growth - 1, "YoY % GROWTH", _SECTION_YOY_FILL)

    # ---- Column sub-headers (row 3) ----
    header_row = 3
    name_header = ws.cell(row=header_row, column=1, value="Particular")
    name_header.fill = _HEADER_FILL
    name_header.font = _HEADER_FONT

    for j, col in enumerate(columns):
        c = ws.cell(row=header_row, column=amt_start + j, value=col)
        c.fill = _HEADER_FILL
        c.font = _HEADER_FONT
        c.alignment = Alignment(horizontal="right")

    for j, col in enumerate(quarter_cols):
        for start_col in (qoq_start, yoy_start):
            c = ws.cell(row=header_row, column=start_col + j, value=col)
            c.fill = _HEADER_FILL
            c.font = _HEADER_FONT
            c.alignment = Alignment(horizontal="right")

    # ---- Data rows ----
    row_index = {}
    r = header_row + 1
    for li in report["line_items"]:
        name = li["particular"]
        is_computed = li["category"] == "Computed"
        is_highlight = name in _HIGHLIGHT_ROWS
        is_percent = name in _PERCENT_ROWS

        name_cell = ws.cell(row=r, column=1, value=name)
        if is_computed:
            name_cell.fill = _COMPUTED_FILL
            name_cell.font = _HIGHLIGHT_FONT if is_highlight else _COMPUTED_FONT
        name_cell.border = _THIN_BORDER

        # amounts block
        for j, col in enumerate(columns):
            val = li["values"].get(col, 0)
            c = ws.cell(row=r, column=amt_start + j, value=val)
            c.alignment = Alignment(horizontal="right")
            c.number_format = '0.0"%"' if is_percent else "#,##0"
            c.border = _THIN_BORDER
            if is_computed:
                c.fill = _COMPUTED_FILL
                c.font = _HIGHLIGHT_FONT if is_highlight else _COMPUTED_FONT

        # QoQ / YoY blocks
        for j, col in enumerate(quarter_cols):
            g = report["growth"].get(name, {}).get(col, {})
            for start_col, mode in ((qoq_start, "qoq"), (yoy_start, "yoy")):
                val = g.get(mode)
                c = ws.cell(row=r, column=start_col + j)
                c.alignment = Alignment(horizontal="right")
                c.border = _THIN_BORDER
                if val is None:
                    c.value = "—"
                    c.fill = _FLAT_FILL
                    c.font = _FLAT_FONT
                else:
                    c.value = val
                    c.number_format = '+0.0"%";-0.0"%"'
                    if val > 0:
                        c.fill, c.font = _UP_FILL, _UP_FONT
                    elif val < 0:
                        c.fill, c.font = _DOWN_FILL, _DOWN_FONT
                    else:
                        c.fill, c.font = _FLAT_FILL, _FLAT_FONT

        row_index[name] = r
        r += 1

    last_data_row = r - 1
    _autosize_columns(ws, total_cols)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=amt_start).coordinate

    # ---- Bar chart: Total Revenue / EBITDA / Net Profit across quarters ----
    if quarter_cols:
        chart_rows = [n for n in ("Total Revenue", "EBITDA", "Net Profit") if n in row_index]
        if chart_rows:
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "clustered"
            chart.title = f"FY{fiscal_year} Quarterly Trend"
            chart.y_axis.title = "Amount"
            chart.x_axis.title = "Quarter"
            chart.style = 10
            chart.width = 24
            chart.height = 10

            cat_col_start = amt_start
            cat_col_end = amt_start + len(quarter_cols) - 1
            cats = Reference(ws, min_col=cat_col_start, max_col=cat_col_end,
                              min_row=header_row, max_row=header_row)

            for name in chart_rows:
                row_num = row_index[name]
                data_ref = Reference(ws, min_col=cat_col_start, max_col=cat_col_end,
                                      min_row=row_num, max_row=row_num)
                chart.series.append(Series(data_ref, title=name))

            chart.set_categories(cats)
            ws.add_chart(chart, f"A{last_data_row + 3}")


def build_styled_report_workbook(rows):
    years = list_years(rows)
    wb = Workbook()
    wb.remove(wb.active)

    if not years:
        ws = wb.create_sheet("No Data")
        ws["A1"] = "No entries yet — add a quarter from the Entry page first."
        ws["A1"].font = _TITLE_FONT
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    for fy in years:
        report = build_report(rows, fy)
        if not report["columns"]:
            continue
        ws = wb.create_sheet(f"FY{fy}"[:31])
        _write_year_sheet(ws, report, fy)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
# ============================================================================
# API
# ============================================================================
app = FastAPI(title="Finance Management System")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


class LoginIn(BaseModel):
    name: str
    access_code: str


class EntryIn(BaseModel):
    fiscal_year: int
    quarter: str
    entered_by: str
    particulars: Dict[str, Dict[str, float]]


@app.post("/api/login")
def login(data: LoginIn):
    if data.access_code != ACCESS_CODE:
        return {"success": False, "message": "Incorrect access code."}
    if not data.name.strip():
        return {"success": False, "message": "Please enter your name."}
    return {"success": True, "name": data.name.strip()}


@app.get("/api/particulars")
def get_particulars():
    return STANDARD_PARTICULARS


@app.post("/api/entry")
def save_entry(data: EntryIn):
    if data.quarter not in QUARTER_ORDER:
        raise HTTPException(400, "quarter must be one of Q1, Q2, Q3, Q4")
    try:
        rows = load_rows()
    except Exception as e:
        raise HTTPException(500, f"Could not read the workbook from GitHub: {e}")

    rows = [r for r in rows if not (int(r["fiscal_year"]) == data.fiscal_year and r["quarter"] == data.quarter)]

    now = datetime.now(timezone.utc).isoformat()
    for category, items in data.particulars.items():
        for particular, amount in items.items():
            if not particular:
                continue
            rows.append({
                "fiscal_year": data.fiscal_year,
                "quarter": data.quarter,
                "category": category,
                "particular": particular,
                "amount": float(amount or 0),
                "entered_by": data.entered_by,
                "updated_at": now,
            })

    try:
        save_rows(rows, f"Update {data.quarter} FY{data.fiscal_year} financials ({data.entered_by})")
    except Exception as e:
        raise HTTPException(500, f"Could not save the workbook to GitHub: {e}")

    return {"success": True}


@app.get("/api/entry/{fiscal_year}/{quarter}")
def get_entry(fiscal_year: int, quarter: str):
    rows = load_rows()
    filtered = [r for r in rows if int(r["fiscal_year"]) == fiscal_year and r["quarter"] == quarter]
    out = {"Revenue": {}, "Expenses": {}, "Other": {}}
    for r in filtered:
        out.setdefault(r["category"], {})[r["particular"]] = r["amount"]
    return out


@app.get("/api/years")
def get_years():
    return {"years": list_years(load_rows())}


@app.get("/api/report/{fiscal_year}")
def get_report(fiscal_year: int):
    return build_report(load_rows(), fiscal_year)


@app.get("/api/dashboard-summary")
def dashboard_summary():
    rows = load_rows()
    return {"summary": latest_summary(rows), "years": list_years(rows), "entry_count": len(rows)}


@app.get("/api/download")
def download_excel():
    try:
        rows = load_rows()
        styled_bytes = build_styled_report_workbook(rows)
    except Exception as e:
        raise HTTPException(500, f"Could not build the report workbook: {e}")
    return StreamingResponse(
        io.BytesIO(styled_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=finance_report.xlsx"},
    )


@app.get("/api/download-url")
def download_url_route():
    return {"url": get_download_url()}
