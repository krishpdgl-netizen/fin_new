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

GITHUB_API = "https://api.github.com"
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
    return f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{path}"


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
        raw = get_raw_file_bytes()
    except Exception as e:
        raise HTTPException(500, f"Could not fetch the workbook from GitHub: {e}")
    return StreamingResponse(
        io.BytesIO(raw),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=finance_data.xlsx"},
    )


@app.get("/api/download-url")
def download_url_route():
    return {"url": get_download_url()}
