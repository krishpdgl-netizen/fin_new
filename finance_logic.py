"""
finance_logic.py
-----------------
Turns the flat list of (fiscal_year, quarter, category, particular, amount) rows
loaded from the Excel workbook into the pivoted report the frontend shows:
particulars down the rows, periods (Q1, Q2, Q3, Q4, H1, 9M, FY) across the columns,
plus QoQ% and YoY% growth.
"""
from collections import defaultdict

QUARTER_ORDER = ["Q1", "Q2", "Q3", "Q4"]

STANDARD_PARTICULARS = {
    "Revenue": [
        "Revenue from Operations",
        "Other Income",
    ],
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


def _quarter_key(fy, q):
    return (int(fy), q)


def group_by_period(rows):
    """{ (fiscal_year:int, quarter:str): { particular: amount } }"""
    grouped = defaultdict(dict)
    for r in rows:
        fy, q, particular, amount = r["fiscal_year"], r["quarter"], r["particular"], r["amount"]
        if fy is None or q is None or particular is None:
            continue
        grouped[_quarter_key(fy, q)][particular] = float(amount or 0)
    return grouped


def compute_subtotals(period_values: dict, particular_categories: dict):
    """Given {particular: amount} for one period, compute the derived rows."""
    revenue_items = [p for p in particular_categories.get("Revenue", []) if p in period_values]
    expense_items = [p for p in particular_categories.get("Expenses", []) if p in period_values]

    total_revenue = sum(period_values.get(p, 0) for p in revenue_items)
    depreciation = period_values.get("Depreciation & Amortization", 0)
    finance_costs = period_values.get("Finance Costs", 0)
    tax = period_values.get("Tax Expense", 0)

    # Total expenses excludes tax (tax is applied after PBT)
    total_expenses = sum(period_values.get(p, 0) for p in expense_items)

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


def _sum_periods(periods_values: list):
    """Sum a list of {particular: amount} dicts into one (for H1 / 9M / FY roll-ups)."""
    out = defaultdict(float)
    for pv in periods_values:
        for k, v in pv.items():
            out[k] += v or 0
    return dict(out)


def build_report(rows, fiscal_year: int, particular_categories: dict = None):
    """
    Returns:
    {
      "columns": ["Q1","Q2","Q3","Q4","H1 (6M)","9M","FY (Annual)"],
      "line_items": [ {"particular": ..., "category": ..., "values": {col: amount}} , ... ],
      "growth": { "Total Revenue": {"Q2": {"qoq": .., "yoy": ..}, ...}, ... }
    }
    """
    particular_categories = particular_categories or STANDARD_PARTICULARS
    grouped = group_by_period(rows)

    # raw values per quarter, this fiscal year
    quarter_raw = {q: grouped.get(_quarter_key(fiscal_year, q), {}) for q in QUARTER_ORDER}
    # raw values per quarter, previous fiscal year (for YoY)
    prev_quarter_raw = {q: grouped.get(_quarter_key(fiscal_year - 1, q), {}) for q in QUARTER_ORDER}

    available_quarters = [q for q in QUARTER_ORDER if quarter_raw[q]]

    columns = list(available_quarters)
    cumulative_raw = {}
    if all(q in available_quarters for q in ["Q1", "Q2"]):
        cumulative_raw["H1 (6M)"] = _sum_periods([quarter_raw["Q1"], quarter_raw["Q2"]])
        columns.append("H1 (6M)")
    if all(q in available_quarters for q in ["Q1", "Q2", "Q3"]):
        cumulative_raw["9M"] = _sum_periods([quarter_raw["Q1"], quarter_raw["Q2"], quarter_raw["Q3"]])
        columns.append("9M")
    if all(q in available_quarters for q in QUARTER_ORDER):
        cumulative_raw["FY (Annual)"] = _sum_periods([quarter_raw[q] for q in QUARTER_ORDER])
        columns.append("FY (Annual)")

    all_period_values = {**{q: quarter_raw[q] for q in available_quarters}, **cumulative_raw}

    # all particular names that actually have data, in category order
    all_particulars = []
    for cat in ["Revenue", "Expenses", "Other"]:
        for p in particular_categories.get(cat, []):
            all_particulars.append((cat, p))
    # include any custom particulars entered that aren't in the standard list
    seen = {p for _, p in all_particulars}
    for q in available_quarters:
        for p in quarter_raw[q]:
            if p not in seen:
                all_particulars.append(("Other", p))
                seen.add(p)

    line_items = []
    for cat, particular in all_particulars:
        if particular in COMPUTED_ROWS:
            # avoid showing e.g. "Tax Expense" twice - once as raw input, once as computed subtotal
            continue
        values = {col: all_period_values.get(col, {}).get(particular, 0) for col in columns}
        line_items.append({"particular": particular, "category": cat, "values": values})

    # computed subtotal rows, for every column
    computed = {}
    for col in columns:
        computed[col] = compute_subtotals(all_period_values[col], particular_categories)
    for name in COMPUTED_ROWS:
        line_items.append({
            "particular": name,
            "category": "Computed",
            "values": {col: computed[col][name] for col in columns},
        })

    # growth: QoQ (vs previous quarter, this FY) and YoY (vs same quarter, prior FY)
    growth = defaultdict(dict)
    quarter_full_values = {q: {**quarter_raw[q], **computed.get(q, {})} for q in available_quarters}
    prev_quarter_full_values = {}
    for q in QUARTER_ORDER:
        base = prev_quarter_raw[q]
        prev_quarter_full_values[q] = {**base, **compute_subtotals(base, particular_categories)} if base else {}

    row_names = [li["particular"] for li in line_items]
    for name in row_names:
        for i, q in enumerate(available_quarters):
            cur = quarter_full_values[q].get(name, 0)
            qoq = None
            if i > 0:
                prev_q = available_quarters[i - 1]
                prev_val = quarter_full_values[prev_q].get(name, 0)
                if prev_val:
                    qoq = (cur - prev_val) / abs(prev_val) * 100
            yoy = None
            prev_year_val = prev_quarter_full_values.get(q, {}).get(name)
            if prev_year_val:
                yoy = (cur - prev_year_val) / abs(prev_year_val) * 100
            growth[name][q] = {"qoq": qoq, "yoy": yoy}

    return {
        "fiscal_year": fiscal_year,
        "columns": columns,
        "line_items": line_items,
        "growth": growth,
    }


def list_available_fiscal_years(rows):
    years = sorted({int(r["fiscal_year"]) for r in rows if r.get("fiscal_year") is not None}, reverse=True)
    return years


def latest_quarter_summary(rows):
    """For the dashboard cards: latest quarter's Total Revenue / Net Profit / QoQ / YoY."""
    grouped = group_by_period(rows)
    if not grouped:
        return None
    latest_key = max(grouped.keys(), key=lambda k: (k[0], QUARTER_ORDER.index(k[1])))
    fy, q = latest_key
    report = build_report(rows, fy)
    idx = report["columns"].index(q) if q in report["columns"] else None
    rev_row = next((li for li in report["line_items"] if li["particular"] == "Total Revenue"), None)
    np_row = next((li for li in report["line_items"] if li["particular"] == "Net Profit"), None)
    margin_row = next((li for li in report["line_items"] if li["particular"] == "Net Profit Margin %"), None)
    return {
        "fiscal_year": fy,
        "quarter": q,
        "total_revenue": rev_row["values"].get(q, 0) if rev_row else 0,
        "net_profit": np_row["values"].get(q, 0) if np_row else 0,
        "net_margin": margin_row["values"].get(q, 0) if margin_row else 0,
        "qoq_revenue": report["growth"].get("Total Revenue", {}).get(q, {}).get("qoq"),
        "yoy_revenue": report["growth"].get("Total Revenue", {}).get(q, {}).get("yoy"),
        "qoq_net_profit": report["growth"].get("Net Profit", {}).get(q, {}).get("qoq"),
        "yoy_net_profit": report["growth"].get("Net Profit", {}).get(q, {}).get("yoy"),
    }
