import os
import io
from datetime import datetime, timezone

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Dict

import github_helper
import finance_logic

# Vercel routes every request under /api/* to this file (see vercel.json),
# so all routes below are declared with the /api prefix.
app = FastAPI(title="Finance Management System")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

ACCESS_CODE = os.environ.get("ACCESS_CODE", "finance2026")


# ---------------------------------------------------------------- MODELS ----
class LoginIn(BaseModel):
    name: str
    access_code: str


class EntryIn(BaseModel):
    fiscal_year: int
    quarter: str  # Q1 / Q2 / Q3 / Q4
    entered_by: str
    particulars: Dict[str, Dict[str, float]]


# ------------------------------------------------------------------ AUTH ----
@app.post("/api/login")
def login(data: LoginIn):
    if data.access_code != ACCESS_CODE:
        return {"success": False, "message": "Incorrect access code."}
    if not data.name.strip():
        return {"success": False, "message": "Please enter your name."}
    return {"success": True, "name": data.name.strip()}


# ------------------------------------------------------------ PARTICULARS ----
@app.get("/api/particulars")
def get_particulars():
    return finance_logic.STANDARD_PARTICULARS


# ------------------------------------------------------------------ ENTRY ----
@app.post("/api/entry")
def save_entry(data: EntryIn):
    if data.quarter not in finance_logic.QUARTER_ORDER:
        raise HTTPException(400, "quarter must be one of Q1, Q2, Q3, Q4")

    try:
        rows = github_helper.load_rows()
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
        github_helper.save_rows(
            rows,
            commit_message=f"Update {data.quarter} FY{data.fiscal_year} financials ({data.entered_by})",
        )
    except Exception as e:
        raise HTTPException(500, f"Could not save the workbook to GitHub: {e}")

    return {"success": True}


@app.get("/api/entry/{fiscal_year}/{quarter}")
def get_entry(fiscal_year: int, quarter: str):
    rows = github_helper.load_rows()
    filtered = [r for r in rows if int(r["fiscal_year"]) == fiscal_year and r["quarter"] == quarter]
    out = {"Revenue": {}, "Expenses": {}, "Other": {}}
    for r in filtered:
        out.setdefault(r["category"], {})[r["particular"]] = r["amount"]
    return out


# ----------------------------------------------------------------- REPORT ----
@app.get("/api/years")
def get_years():
    rows = github_helper.load_rows()
    return {"years": finance_logic.list_available_fiscal_years(rows)}


@app.get("/api/report/{fiscal_year}")
def get_report(fiscal_year: int):
    rows = github_helper.load_rows()
    return finance_logic.build_report(rows, fiscal_year)


@app.get("/api/dashboard-summary")
def dashboard_summary():
    rows = github_helper.load_rows()
    summary = finance_logic.latest_quarter_summary(rows)
    years = finance_logic.list_available_fiscal_years(rows)
    return {"summary": summary, "years": years, "entry_count": len(rows)}


# ---------------------------------------------------------------- DOWNLOAD ----
@app.get("/api/download")
def download_excel():
    try:
        raw = github_helper.get_raw_file_bytes()
    except Exception as e:
        raise HTTPException(500, f"Could not fetch the workbook from GitHub: {e}")
    return StreamingResponse(
        io.BytesIO(raw),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=finance_data.xlsx"},
    )


@app.get("/api/download-url")
def download_url():
    return {"url": github_helper.download_url()}
