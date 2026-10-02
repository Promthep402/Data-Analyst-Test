import json
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
SUMMARY = ROOT / "FruitBlend24_Analysis_Summary.json"
OUTPUT = ROOT / "FruitBlend24_Data_Analyst_Presentation.pptx"
ASSET_DIR = ROOT / "presentation_assets"
ASSET_DIR.mkdir(exist_ok=True)

BG = "FFF9F4"
INK = "4A2B20"
MUTED = "876B5D"
ACCENT = "C96F3D"
ACCENT_DARK = "A65331"
PEACH = "F7E9DE"
GREEN = "4A806F"
RED = "A95A4A"
WHITE = "FFFDFB"


def rgb(hex_value):
    return RGBColor.from_string(hex_value)


def money(value):
    return f"THB {round(value):,}"


def pct(value):
    return f"{value * 100:.1f}%"


def save_chart(name, draw):
    path = ASSET_DIR / f"{name}.png"
    plt.rcParams.update({"font.family": "DejaVu Sans", "axes.titleweight": "bold"})
    fig, ax = plt.subplots(figsize=(10.5, 4.8), dpi=180)
    fig.patch.set_facecolor(f"#{BG}")
    ax.set_facecolor(f"#{BG}")
    draw(fig, ax)
    fig.tight_layout(pad=1.2)
    fig.savefig(path, facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)
    return path


def add_text(slide, text, left, top, width, height, size=18, color=INK, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.clear()
    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Aptos"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return box


def add_rich_text(slide, lines, left, top, width, height):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = 0
    frame.margin_right = 0
    for index, (label, body) in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.space_after = Pt(9)
        label_run = paragraph.add_run()
        label_run.text = label
        label_run.font.name = "Aptos"
        label_run.font.size = Pt(16)
        label_run.font.bold = True
        label_run.font.color.rgb = rgb(ACCENT_DARK)
        body_run = paragraph.add_run()
        body_run.text = body
        body_run.font.name = "Aptos"
        body_run.font.size = Pt(16)
        body_run.font.color.rgb = rgb(INK)
    return box


def add_header(slide, section, title, subtitle=None):
    add_text(slide, section.upper(), 0.65, 0.38, 5, 0.25, 10, ACCENT, True)
    add_text(slide, title, 0.65, 0.72, 11.9, 0.55, 27, INK, True)
    if subtitle:
        add_text(slide, subtitle, 0.65, 1.34, 11.9, 0.35, 12, MUTED)


def add_footer(slide, number):
    add_text(slide, "FruitBlend24 | Python data analysis", 0.65, 7.08, 6, 0.2, 9, MUTED)
    add_text(slide, f"{number:02d}", 12.1, 7.08, 0.5, 0.2, 9, MUTED, True, PP_ALIGN.RIGHT)


def add_card(slide, left, top, width, height, fill=WHITE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.color.rgb = rgb("EFDCCD")
    shape.line.width = Pt(0.8)
    return shape


def add_kpi(slide, left, label, value, detail):
    add_card(slide, left, 2.05, 2.85, 1.35)
    add_text(slide, label.upper(), left + 0.18, 2.24, 2.45, 0.2, 9, MUTED, True)
    add_text(slide, value, left + 0.18, 2.52, 2.48, 0.4, 22, INK, True)
    add_text(slide, detail, left + 0.18, 3.02, 2.48, 0.2, 9, MUTED)


def build_presentation(data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    monthly = data["monthly"]
    sku_rows = data["skuRows"]
    rates = data["rateRows"]
    kitchens = data["kitchenRows"]
    quality = data["quality"]
    totals = data["assumptions"]

    def new_slide():
        slide = prs.slides.add_slide(blank)
        background = slide.background.fill
        background.solid()
        background.fore_color.rgb = rgb(BG)
        return slide

    def revenue_chart(fig, ax):
        labels = [row["month"] for row in monthly]
        revenue = [row["revenue"] / 1_000_000 for row in monthly]
        profit = [row["grossProfit"] / 1_000_000 for row in monthly]
        ax.plot(labels, revenue, color=f"#{ACCENT}", linewidth=3, marker="o", label="Net revenue")
        ax.plot(labels, profit, color=f"#{RED}", linewidth=2.5, marker="o", label="Gross profit")
        ax.axhline(0, color=f"#{MUTED}", linewidth=0.8)
        ax.set_ylabel("THB million", color=f"#{MUTED}")
        ax.tick_params(axis="x", rotation=45, labelsize=8, colors=f"#{MUTED}")
        ax.tick_params(axis="y", labelsize=9, colors=f"#{MUTED}")
        ax.grid(axis="y", alpha=0.2)
        ax.legend(frameon=False, ncol=2, loc="upper left")
        for spine in ax.spines.values(): spine.set_visible(False)

    def sku_chart(fig, ax):
        ordered = sorted(sku_rows, key=lambda row: row["sold"])
        ax.barh([row["sku"] for row in ordered], [row["sold"] / 1000 for row in ordered], color=f"#{ACCENT}")
        ax.set_xlabel("Cups sold (thousands)", color=f"#{MUTED}")
        ax.xaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:,.0f}"))
        ax.tick_params(axis="both", labelsize=9, colors=f"#{MUTED}")
        ax.grid(axis="x", alpha=0.2)
        for spine in ax.spines.values(): spine.set_visible(False)

    def rate_chart(fig, ax):
        labels = [row["rateCode"] for row in rates]
        margins = [row["margin"] * 100 for row in rates]
        colors = [f"#{GREEN}" if value >= 30 else f"#{ACCENT}" for value in margins]
        ax.bar(labels, margins, color=colors, width=0.58)
        ax.set_ylabel("Contribution margin (%)", color=f"#{MUTED}")
        ax.set_ylim(0, 45)
        ax.tick_params(axis="both", labelsize=9, colors=f"#{MUTED}")
        ax.grid(axis="y", alpha=0.2)
        for spine in ax.spines.values(): spine.set_visible(False)

    def kitchen_chart(fig, ax):
        ordered = sorted(kitchens, key=lambda row: row["profit"])
        colors = [f"#{RED}" if row["profit"] < 0 else f"#{GREEN}" for row in ordered]
        ax.barh([row["kitchen"].replace("_", " ") for row in ordered], [row["profit"] / 1000 for row in ordered], color=colors)
        ax.axvline(0, color=f"#{MUTED}", linewidth=0.8)
        ax.set_xlabel("Profit after fixed overhead (THB thousands)", color=f"#{MUTED}")
        ax.tick_params(axis="both", labelsize=9, colors=f"#{MUTED}")
        ax.grid(axis="x", alpha=0.2)
        for spine in ax.spines.values(): spine.set_visible(False)

    revenue_path = save_chart("monthly_performance", revenue_chart)
    sku_path = save_chart("sku_mix", sku_chart)
    rate_path = save_chart("promotion_margin", rate_chart)
    kitchen_path = save_chart("kitchen_profit", kitchen_chart)

    # 1. Cover
    slide = new_slide()
    add_text(slide, "FRUITBLEND24", 0.8, 0.9, 6.5, 0.35, 13, ACCENT, True)
    add_text(slide, "From sales data\nto operating decisions", 0.8, 1.55, 8.5, 1.45, 38, INK, True)
    add_text(slide, "Data Analyst Case Study | Sep 2025 - Aug 2026", 0.85, 3.45, 7, 0.3, 16, MUTED)
    add_card(slide, 8.95, 1.25, 3.3, 3.9, PEACH)
    add_text(slide, "CORE SIGNAL", 9.35, 1.72, 2.5, 0.25, 10, ACCENT_DARK, True)
    add_text(slide, "Volume is healthy.\nProfit is selective.", 9.35, 2.2, 2.6, 1.0, 24, INK, True)
    add_text(slide, "The opportunity is to scale the right SKUs and kitchens, not simply more cups.", 9.35, 3.62, 2.45, 0.95, 13, INK)
    add_text(slide, "Prepared from the FruitBlend24 case workbook", 0.85, 6.7, 6.5, 0.25, 10, MUTED)

    # 2. Executive summary
    slide = new_slide(); add_header(slide, "01 | Executive summary", "The business has demand, but not uniform economics")
    add_kpi(slide, 0.65, "Mapped cups sold", f"{totals['totalUnits']:,}", "across five core SKUs")
    add_kpi(slide, 3.75, "Net revenue", money(totals["totalNet"]), "after platform commissions")
    add_kpi(slide, 6.85, "Gross profit", money(totals["totalProfit"]), "after variable cost + overhead")
    add_kpi(slide, 9.95, "Waste cost", money(totals["wasteCost"]), "estimated from daily waste")
    add_rich_text(slide, [("01  Protect the engine. ", "Watermelon and Pineapple drive 65.9% of cups sold."), ("02  Fix the drag. ", "BKK kitchens are loss-making after fixed overhead."), ("03  Spend selectively. ", "Promotions lift volume but dilute contribution margin.")], 0.8, 4.05, 11.8, 1.65)
    add_footer(slide, 2)

    # 3. Monthly performance
    slide = new_slide(); add_header(slide, "02 | Performance", "Revenue recovered in spring, then softened again", "Monthly net revenue and gross profit, THB millions")
    slide.shapes.add_picture(str(revenue_path), Inches(0.65), Inches(1.85), width=Inches(8.25))
    add_card(slide, 9.25, 1.95, 3.35, 3.95)
    add_text(slide, "READOUT", 9.6, 2.3, 2.3, 0.2, 10, ACCENT, True)
    add_rich_text(slide, [("Peak revenue. ", "Apr 2026 reached " + money(max(row["revenue"] for row in monthly)) + "."), ("Break-even window. ", "Profit turned positive from Dec through May."), ("Watch-out. ", "Jun-Aug returned to negative monthly profit.")], 9.6, 2.75, 2.55, 2.55)
    add_footer(slide, 3)

    # 4. SKU mix
    slide = new_slide(); add_header(slide, "03 | Demand mix", "Two core SKUs carry the volume base", "Cups sold by core SKU")
    slide.shapes.add_picture(str(sku_path), Inches(0.65), Inches(1.75), width=Inches(7.7))
    leader = max(sku_rows, key=lambda row: row["sold"])
    lowest_margin = min(sku_rows, key=lambda row: row["margin"])
    add_card(slide, 8.75, 1.9, 3.8, 3.75)
    add_text(slide, "WHAT IT MEANS", 9.1, 2.25, 2.8, 0.2, 10, ACCENT, True)
    add_rich_text(slide, [("Volume leader. ", f"{leader['sku']} represents {pct(leader['share'])} of cups sold."), ("Premium risk. ", f"{lowest_margin['sku']} has the lowest contribution margin at {pct(lowest_margin['margin'])}."), ("Action. ", "Use the volume leaders to fund targeted premium mix tests.")], 9.1, 2.7, 2.9, 2.5)
    add_footer(slide, 4)

    # 5. Promotion economics
    slide = new_slide(); add_header(slide, "04 | Promotions", "Promotions buy volume at a measurable margin cost", "Contribution margin by rate code")
    slide.shapes.add_picture(str(rate_path), Inches(0.7), Inches(1.85), width=Inches(7.35))
    add_card(slide, 8.55, 1.95, 4.05, 3.9)
    add_text(slide, "DECISION RULE", 8.95, 2.3, 2.7, 0.2, 10, ACCENT, True)
    add_text(slide, "Keep promotions\nthat earn repeat\ndemand, not just\norder spikes.", 8.95, 2.75, 3.1, 1.25, 23, INK, True)
    add_text(slide, "Standard RC000 margin: " + pct(rates[0]["margin"]) + "\nLowest promo margin: " + pct(min(row["margin"] for row in rates[1:])), 8.95, 4.45, 3.1, 0.75, 14, MUTED)
    add_footer(slide, 5)

    # 6. Kitchen economics
    slide = new_slide(); add_header(slide, "05 | Operations", "Kitchen economics split the network in two", "Profit after fixed monthly overhead")
    slide.shapes.add_picture(str(kitchen_path), Inches(0.7), Inches(1.9), width=Inches(7.55))
    add_card(slide, 8.65, 1.95, 3.85, 3.85)
    add_text(slide, "OPERATING PRIORITY", 9.0, 2.3, 3.0, 0.2, 10, ACCENT, True)
    add_text(slide, "Stabilize BKK\nbefore adding\ncapacity.", 9.0, 2.78, 3.0, 1.0, 25, INK, True)
    add_text(slide, "Pattaya contributes positive profit. BKK Sukhumvit and BKK Ladprao are negative after fixed overhead.", 9.0, 4.35, 2.85, 0.85, 14, MUTED)
    add_footer(slide, 6)

    # 7. Data quality
    slide = new_slide(); add_header(slide, "06 | Data quality", "The core export is reliable enough for decisions")
    add_kpi(slide, 0.75, "Raw rows", f"{quality['total']:,}", "hourly order export")
    add_kpi(slide, 3.85, "Mapped", f"{quality['mapped']:,}", f"{quality['mapped'] / quality['total'] * 100:.1f}% of rows")
    add_kpi(slide, 6.95, "Revenue checks", f"{quality['revenueMismatches']:,}", "arithmetic mismatches")
    add_kpi(slide, 10.05, "Duplicate signatures", f"{quality['duplicateRows']:,}", "flagged for review")
    add_rich_text(slide, [("Method. ", "Non-core or unmapped rows are excluded from P&L and disclosed."), ("Confidence. ", "Gross revenue reconciles to price x units with zero mismatches."), ("Next control. ", "Investigate duplicate signatures before production reporting.")], 0.9, 4.15, 11.2, 1.55)
    add_footer(slide, 7)

    # 8. Recommendations
    slide = new_slide(); add_header(slide, "07 | Recommendations", "Three actions to improve the next quarter")
    recommendations = [("01", "Scale profitable demand", "Protect availability of Watermelon and Pineapple while testing higher-margin bundles."), ("02", "Repair kitchen economics", "Review BKK staffing, fixed-cost allocation, and order routing before adding capacity."), ("03", "Make promotions accountable", "Track promotion cohorts and repeat purchase, not only units sold during campaign windows.")]
    for index, (number, title, body) in enumerate(recommendations):
        top = 1.85 + index * 1.42
        add_card(slide, 0.8, top, 11.75, 1.05)
        add_text(slide, number, 1.15, top + 0.29, 0.55, 0.3, 18, ACCENT, True)
        add_text(slide, title, 2.0, top + 0.18, 4.2, 0.3, 18, INK, True)
        add_text(slide, body, 6.15, top + 0.2, 5.8, 0.52, 13, MUTED)
    add_text(slide, "Decision principle: grow contribution, not just volume.", 0.85, 6.45, 8.5, 0.35, 20, ACCENT_DARK, True)
    add_footer(slide, 8)

    prs.save(OUTPUT)


if __name__ == "__main__":
    build_presentation(json.loads(SUMMARY.read_text(encoding="utf-8")))
    print(f"Created {OUTPUT}")