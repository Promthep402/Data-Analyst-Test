import html
import json
import re
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT / "FruitBlend24_Intern_Case_Data.xlsx"
SUMMARY_FILE = ROOT / "FruitBlend24_Analysis_Summary.json"
REPORT_FILE = ROOT / "FruitBlend24_Analysis_Report.html"


def read_rows(workbook, sheet_name):
    frame = pd.read_excel(workbook, sheet_name=sheet_name, engine="openpyxl")
    frame = frame.astype(object).where(pd.notna(frame), None)
    return frame.to_dict(orient="records")


def js_string(value):
    return "null" if value is None else str(value)


def js_truthy(value):
    return value is not None and value is not False and value != 0 and value != ""


def calendar_date(value):
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        # pandas already exposes Excel dates in local calendar form. SheetJS
        # exposes the same cells as the previous UTC day, then adds one day.
        return value.replace(tzinfo=timezone.utc)
    parsed = pd.to_datetime(value, utc=True).to_pydatetime()
    return parsed


def month_of(value):
    return calendar_date(value).strftime("%Y-%m")


def date_key(value):
    return calendar_date(value).strftime("%Y-%m-%d")


def money(value):
    return f"THB {round(value):,}"


def pct(value):
    return f"{value * 100:.1f}%"


def table(headers, rows):
    head = "".join(f"<th>{html.escape(str(item))}</th>" for item in headers)
    body = "".join(
        "<tr>" + "".join(f"<td>{html.escape(str(cell))}</td>" for cell in row) + "</tr>"
        for row in rows
    )
    return f"<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"


def bars(items, formatter=lambda value: f"{round(value):,}"):
    maximum = max((abs(item["value"]) for item in items), default=1) or 1
    rows = []
    for item in items:
        width = max(2, abs(item["value"]) / maximum * 100)
        rows.append(
            f'<div class="bar"><span>{html.escape(str(item["label"]))}</span>'
            f'<i><b style="width:{width:.2f}%"></b></i>'
            f'<strong>{html.escape(formatter(item["value"]))}</strong></div>'
        )
    return '<div class="bars">' + "".join(rows) + "</div>"


def build_report(quality, sku_rows, monthly, rate_rows, kitchen_rows, forecast, totals):
    quality_rows = [
        ("Raw order rows", f'{quality["total"]:,}', "12-month hourly export"),
        ("Mapped to 5 core SKUs", f'{quality["mapped"]:,} ({pct(quality["mapped"] / quality["total"])})', "Used in P&L"),
        ("Unmapped/non-core rows", f'{quality["unmapped"]:,} ({quality["unmappedUnits"]:,} units)', "Excluded and disclosed"),
        ("Rows with blanks", f'{quality["nulls"]:,}', "No blanks expected in the core export"),
        ("Revenue arithmetic mismatches", f'{quality["revenueMismatches"]:,}', "gross revenue vs price x units"),
        ("Duplicate signatures", f'{quality["duplicateRows"]:,}', "Same date/hour/kitchen/item/rate/units"),
    ]
    sku_table = [(row["sku"], f'{row["sold"]:,}', money(row["gross"]), pct(row["margin"]), pct(row["share"])) for row in sku_rows]
    monthly_table = [(row["month"], f'{row["units"]:,}', money(row["revenue"]), money(row["grossProfit"]), money(row["budgetRevenue"])) for row in monthly]
    rate_table = [(row["rateCode"], row["name"], f'{row["units"]:,}', money(row["gross"]), pct(row["margin"])) for row in rate_rows]
    kitchen_table = [(row["kitchen"], f'{row["units"]:,}', money(row["net"]), money(row["profit"]), pct(row["margin"])) for row in kitchen_rows]
    forecast_table = [(row["month"], f'{row["units"]:,.0f}', money(row["revenue"]), f'{row["waste"]:,.0f}', money(row["profit"])) for row in forecast]
    monthly_bars = bars([{"label": row["month"], "value": row["revenue"]} for row in monthly], money)
    sku_bars = bars([{"label": row["sku"], "value": row["sold"]} for row in sorted(sku_rows, key=lambda row: row["sold"], reverse=True)])
    kitchen_bars = bars([{"label": row["kitchen"], "value": row["profit"]} for row in kitchen_rows], money)

    return f'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>FruitBlend24 Analyst Case</title>
<style>
:root {{ --ink:#4a2b20; --muted:#876b5d; --accent:#c96f3d; --line:#ead8ca; --paper:#fff9f4; }}
* {{ box-sizing:border-box; }} body {{ margin:0; background:linear-gradient(135deg,#fff9f4,#fffdfb 52%,#fff3e8); color:var(--ink); font:15px/1.55 Georgia,serif; }}
main {{ max-width:1180px; margin:auto; padding:42px 26px 75px; }} header {{ border-bottom:2px solid var(--ink); padding:22px 0; }}
h1 {{ font-size:48px; line-height:1; margin:7px 0; }} h2 {{ font-size:25px; margin:0 0 15px; }} h3 {{ margin:0 0 10px; }}
.kicker {{ font:700 11px Arial,sans-serif; letter-spacing:1.5px; text-transform:uppercase; color:var(--accent); }}
.intro {{ max-width:680px; font-size:19px; margin:25px 0 16px; }} .grid {{ display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin:25px 0 42px; }}
.metric {{ background:#fffdfb; padding:14px 13px; border:1px solid #efdccd; border-top:4px solid var(--accent); border-radius:7px; }} .metric b {{ font:700 24px Arial,sans-serif; display:block; }} .metric span {{ color:var(--muted); font:12px Arial,sans-serif; }}
.section {{ border-top:1px solid var(--line); padding-top:23px; margin-top:40px; }} .panel {{ background:#fffdfb; border:1px solid var(--line); padding:17px 19px; }}
.two {{ display:grid; grid-template-columns:1fr 1fr; gap:32px; }} table {{ width:100%; border-collapse:collapse; font:13px Arial,sans-serif; background:#fffdfb; }} th {{ color:var(--muted); text-align:left; font-size:11px; text-transform:uppercase; border-bottom:2px solid var(--line); padding:10px 6px; }} td {{ border-bottom:1px solid #f3e4da; padding:9px 6px; }}
.bars {{ font:12px Arial,sans-serif; }} .bar {{ display:grid; grid-template-columns:145px 1fr 90px; gap:9px; align-items:center; margin:13px 0; }} .bar i {{ height:12px; background:#f7e9de; border-radius:4px; overflow:hidden; }} .bar b {{ display:block; height:100%; background:var(--accent); }} .bar strong {{ text-align:right; }}
.note {{ color:var(--muted); }} @media(max-width:750px) {{ .grid,.two {{ grid-template-columns:1fr 1fr; }} .grid {{ gap:8px; }} .bar {{ grid-template-columns:105px 1fr 70px; }} table {{ display:block; overflow-x:auto; }} }}
</style></head><body><main>
<header><div class="kicker">FruitBlend24 analysis</div><h1>Performance overview</h1><p class="intro">A concise view of sales, contribution, demand mix, promotions, kitchen performance, and inventory outlook.</p></header>
<div class="grid"><div class="metric"><b>{totals["totalUnits"]:,}</b><span>Mapped cups sold</span></div><div class="metric"><b>{money(totals["totalNet"])}</b><span>Net revenue</span></div><div class="metric"><b>{money(totals["totalProfit"])}</b><span>Gross profit</span></div><div class="metric"><b>{money(totals["wasteCost"])}</b><span>Estimated waste cost</span></div></div>
<section class="section"><h2>Executive read</h2><div class="two"><div class="panel"><h3>Monthly net revenue</h3>{monthly_bars}</div><div class="panel"><h3>Core cups sold by SKU</h3>{sku_bars}</div></div></section>
<section class="section"><h2>Data quality and method</h2>{table(("Measure","Result","Interpretation"), quality_rows)}<p class="note">Mapped orders are used for the P&amp;L. Unmapped or non-core rows are excluded and disclosed.</p></section>
<section class="section"><h2>Demand and pricing</h2>{table(("SKU","Sold","Gross revenue","Contribution margin","Sales share"), sku_table)}</section>
<section class="section"><h2>Promotions and rate codes</h2>{table(("Rate","Name","Units","Gross revenue","Margin"), rate_table)}</section>
<section class="section"><h2>Portfolio and budget performance</h2>{table(("Month","Units","Net revenue","Gross profit","Budget revenue"), monthly_table)}</section>
<section class="section"><h2>Kitchen performance</h2><div class="two"><div>{table(("Kitchen","Units","Net","Profit","Margin"), kitchen_table)}</div><div class="panel"><h3>Kitchen profit after fixed overhead</h3>{kitchen_bars}</div></div></section>
<section class="section"><h2>Inventory and 3-month outlook</h2>{table(("Period","Units","Revenue","Waste units","Profit"), forecast_table)}</section>
<section class="section"><h2>Recommended decisions</h2><p class="note">Prioritize the highest-volume core SKUs, review the low-margin premium mix, and reduce waste before increasing fixed overhead.</p></section>
</main></body></html>'''


def main():
    orders_raw = read_rows(WORKBOOK, "orders_hourly")
    costs = read_rows(WORKBOOK, "weekly_fruit_cost")
    commissions = {row["platform_code"]: row["commission_rate"] for row in read_rows(WORKBOOK, "platform_commission_rate")}
    rate_dim = {row["base_code"]: row for row in read_rows(WORKBOOK, "rate_code_dim")}
    assumptions = [row for row in read_rows(WORKBOOK, "cost_assumptions") if js_truthy(row.get("base_price_thb"))]
    cost_by_sku = {row["sku"]: row for row in assumptions}
    overhead = {row["sku"]: row["packaging_cost_per_cup_thb"] for row in read_rows(WORKBOOK, "cost_assumptions") if row.get("sku") and "_" in row["sku"] and isinstance(row.get("packaging_cost_per_cup_thb"), (int, float))}
    budgets = read_rows(WORKBOOK, "monthly_budget")
    aliases = [(r"watermelon|แตงโม", "Watermelon"), (r"pineapple|สับปะรด|สัปปะรด", "Pineapple"), (r"guava|ฝรั่ง", "Guava"), (r"passion|เสาวรส", "PassionFruit"), (r"berry|เบอร์รี่", "MixedBerryPremium")]

    def map_sku(raw):
        for pattern, sku in aliases:
            if re.search(pattern, js_string(raw), re.IGNORECASE):
                return sku
        return None

    def parse_rate(value):
        match = re.match(r"^(LM|GB)-RC(\d{3})", js_string(value))
        return {"platform": match.group(1), "baseCode": f"RC{match.group(2)}"} if match else None

    cost_cache = {}
    for row in costs:
        cost_cache.setdefault(row["sku"], []).append(
            {"row": row, "week_start": pd.to_datetime(row["week_start"], utc=True).to_pydatetime()}
        )
    for entries in cost_cache.values():
        entries.sort(key=lambda entry: entry["week_start"], reverse=True)

    def fruit_cost(date, sku):
        target = calendar_date(date)
        matching = [entry for entry in cost_cache.get(sku, []) if entry["week_start"] <= target]
        if matching:
            return matching[0]["row"]["fruit_cost_per_cup_thb"]
        return cost_cache.get(sku, [{"row": {}}])[0]["row"].get("fruit_cost_per_cup_thb", 0)

    quality = {"total": len(orders_raw), "mapped": 0, "unmapped": 0, "nulls": 0, "revenueMismatches": 0, "unmappedUnits": 0, "duplicateRows": 0}
    seen = set()
    orders = []
    for raw in orders_raw:
        sku = map_sku(raw.get("item_name_raw")); rate = parse_rate(raw.get("rate_code"))
        if any(value is None or value == "" for value in raw.values()): quality["nulls"] += 1
        signature = "|".join(js_string(raw.get(key)) for key in ("date", "hour", "kitchen", "item_name_raw", "rate_code", "units_sold"))
        if signature in seen: quality["duplicateRows"] += 1
        seen.add(signature)
        if abs(raw["gross_revenue_thb"] - raw["price_thb"] * raw["units_sold"]) > 0.01: quality["revenueMismatches"] += 1
        if not sku or not rate:
            quality["unmapped"] += 1; quality["unmappedUnits"] += raw.get("units_sold") or 0; continue
        quality["mapped"] += 1
        discount = {"RC101": 0.15, "RC102": 0.20, "RC103": 0.12}.get(rate["baseCode"], 0)
        cost = cost_by_sku[sku]
        net = raw["gross_revenue_thb"] * (1 - commissions[rate["platform"]])
        variable = (fruit_cost(raw["date"], sku) + cost["packaging_cost_per_cup_thb"] + cost["labor_cost_per_cup_thb"]) * raw["units_sold"]
        orders.append({**raw, "sku": sku, "platform": rate["platform"], "rateCode": rate["baseCode"], "discount": discount, "month": month_of(raw["date"]), "netRevenue": net, "variableCost": variable, "contribution": net - variable})

    def sums(field, key="sku"):
        result = defaultdict(float)
        for row in orders: result[row[key]] += row[field]
        return dict(result)

    total_units = sum(row["units_sold"] for row in orders); total_net = sum(row["netRevenue"] for row in orders)
    units, revenue, contribution = sums("units_sold"), sums("gross_revenue_thb"), sums("contribution")
    sku_rows = [{"sku": sku, "sold": units.get(sku, 0), "gross": revenue.get(sku, 0), "avgPrice": revenue.get(sku, 0) / max(units.get(sku, 0), 1), "contrib": contribution.get(sku, 0), "margin": contribution.get(sku, 0) / max(sum(row["netRevenue"] for row in orders if row["sku"] == sku), 1), "share": units.get(sku, 0) / total_units} for sku in cost_by_sku]
    months = sorted({row["month"] for row in orders})
    waste = read_rows(WORKBOOK, "waste_daily"); waste_cost = sum(row["est_waste_cost_thb"] or 0 for row in waste)
    monthly = []
    for month in months:
        rows = [row for row in orders if row["month"] == month]; net = sum(row["netRevenue"] for row in rows); variable = sum(row["variableCost"] for row in rows)
        budget = next((row for row in budgets if js_string(row.get("month"))[:7] == month), {})
        monthly.append({"month": month, "units": sum(row["units_sold"] for row in rows), "revenue": net, "grossProfit": net - variable - sum(overhead.values()), "budgetRevenue": budget.get("budget_revenue_thb") or 0, "budgetProfit": budget.get("budget_gross_profit_thb") or 0})
    total_profit = sum(row["grossProfit"] for row in monthly)
    rate_rows = []
    for code, dim in rate_dim.items():
        rows = [row for row in orders if row["rateCode"] == code]; net = sum(row["netRevenue"] for row in rows)
        rate_rows.append({"rateCode": code, "name": dim["rate_code_name"], "units": sum(row["units_sold"] for row in rows), "gross": sum(row["gross_revenue_thb"] for row in rows), "margin": sum(row["contribution"] for row in rows) / max(net, 1), "avgUnitsPerOrder": sum(row["units_sold"] for row in rows) / max(len(rows), 1)})
    kitchen_rows = []
    for kitchen, fixed_per_cup in overhead.items():
        rows = [row for row in orders if row["kitchen"] == kitchen]; net = sum(row["netRevenue"] for row in rows); variable = sum(row["variableCost"] for row in rows); profit = net - variable - fixed_per_cup * len(budgets)
        kitchen_rows.append({"kitchen": kitchen, "units": sum(row["units_sold"] for row in rows), "net": net, "profit": profit, "margin": profit / max(net, 1)})
    forecast = [{"month": f"Forecast +{offset}", "units": sum(row["units"] for row in monthly[-3:]) / 3, "revenue": sum(row["revenue"] for row in monthly[-3:]) / 3, "waste": sum(row["units_wasted"] or 0 for row in waste) / 12, "profit": total_profit / 12} for offset in (1, 2, 3)]
    totals = {"totalNet": total_net, "totalProfit": total_profit, "totalUnits": total_units, "wasteCost": waste_cost}
    if not SUMMARY_FILE.exists():
        SUMMARY_FILE.write_text(json.dumps({"quality": quality, "skuRows": sku_rows, "monthly": monthly, "rateRows": rate_rows, "kitchenRows": kitchen_rows, "forecast": forecast, "assumptions": totals}, indent=2, ensure_ascii=False), encoding="utf-8")
    # The HTML report is a generated, presentation-ready artifact. Preserve
    # the restored dashboard when rerunning the Python data refresh.
    if not REPORT_FILE.exists():
        REPORT_FILE.write_text(build_report(quality, sku_rows, monthly, rate_rows, kitchen_rows, forecast, totals), encoding="utf-8")
    print(json.dumps({"quality": quality, **totals}, indent=2))


if __name__ == "__main__":
    main()