"""Workbooks used by ``validate_manual.py``.

Two fixtures, with different jobs:

``build_sales_workbook`` is a realistic B2B commercial and margin review
workbook (``Summary``, ``Sales``, ``Catalog``, ``Assumptions``). It is what the
README screenshots are captured from, so it combines corporate spreadsheet
conventions — an executive KPI scorecard, regional plan vs. actuals, a 40-order
B2B ledger, a product pricing catalog, defined policy parameters, charts,
comments, frozen panes, merged banners and hidden cost/baseline columns — while
keeping its lineage graph readable at a glance through stretched-formula
grouping.

``build_stress_workbook`` is the opposite. It is an English workbook built to
break things: every Excel error value, references the analyser cannot resolve,
sheet names that need quoting, a circular pair, formulas past the dossier
truncation limit, and cell text chosen to escape the generated HTML if the
escaping is wrong. Nothing in it is meant to look like a real business file —
each sheet is a category of failure, named after what it attacks.

Not covered here: VBA. ``openpyxl`` can preserve an existing ``vbaProject.bin``
but cannot author one, so exercising ``linexcel.vba`` needs a real ``.xlsm``
recorded by Excel. ``tests/test_lineage.py`` covers that path with injected
modules instead.
"""

from __future__ import annotations

import datetime
import io

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ──────────────────────────────────────────────
# 1. B2B commercial & margin review — showcase fixture
# ──────────────────────────────────────────────

CATALOG_ITEMS: tuple[tuple[str, str, str, int, int, str], ...] = (
    (
        "NS-CORE-100",
        "Northstar Core Platform",
        "Cloud Core",
        15_000,
        3_300,
        "Annual SaaS",
    ),
    (
        "NS-CORE-500",
        "Northstar Enterprise Suite",
        "Cloud Core",
        32_000,
        6_400,
        "Annual SaaS",
    ),
    (
        "NS-GOV-200",
        "Lineage & Governance Pro",
        "Data Governance",
        20_000,
        4_400,
        "Annual SaaS",
    ),
    ("NS-SEC-300", "Zero-Trust Audit Shield", "Security", 18_000, 4_320, "Annual SaaS"),
    (
        "NS-AI-600",
        "Automated Anomaly Copilot",
        "AI Analytics",
        24_000,
        6_240,
        "Annual SaaS",
    ),
    (
        "NS-EDGE-410",
        "Edge Telemetry Appliance",
        "Hybrid Edge",
        12_000,
        6_240,
        "3-Yr Appliance",
    ),
    (
        "NS-EDGE-820",
        "High-Throughput Relay Node",
        "Hybrid Edge",
        26_000,
        14_300,
        "3-Yr Appliance",
    ),
    (
        "NS-SVC-900",
        "Enterprise Onboarding Pack",
        "Advisory",
        10_000,
        5_800,
        "Fixed Scope",
    ),
)

SALES_ORDERS: tuple[tuple[str, datetime.date, str, str, str, int, float], ...] = (
    (
        "SO-26301",
        datetime.date(2026, 7, 2),
        "Meridian Health Systems",
        "North America",
        "NS-CORE-500",
        2,
        0.05,
    ),
    (
        "SO-26302",
        datetime.date(2026, 7, 6),
        "Helvetia Private Bank",
        "EMEA",
        "NS-GOV-200",
        2,
        0.05,
    ),
    (
        "SO-26303",
        datetime.date(2026, 7, 8),
        "Sapporo Precision Robotics",
        "APAC",
        "NS-EDGE-820",
        2,
        0.08,
    ),
    (
        "SO-26304",
        datetime.date(2026, 7, 10),
        "Vanguard Freight Logistics",
        "North America",
        "NS-CORE-100",
        3,
        0.08,
    ),
    (
        "SO-26305",
        datetime.date(2026, 7, 13),
        "Andes Digital Payments",
        "LATAM",
        "NS-CORE-100",
        2,
        0.10,
    ),
    (
        "SO-26306",
        datetime.date(2026, 7, 15),
        "Nordic Grid Transmission",
        "EMEA",
        "NS-SEC-300",
        2,
        0.05,
    ),
    (
        "SO-26307",
        datetime.date(2026, 7, 17),
        "Kestrel Aerospace Group",
        "North America",
        "NS-AI-600",
        2,
        0.05,
    ),
    (
        "SO-26308",
        datetime.date(2026, 7, 21),
        "Pacific Harbour Ports",
        "APAC",
        "NS-CORE-100",
        2,
        0.05,
    ),
    (
        "SO-26309",
        datetime.date(2026, 7, 23),
        "Rheinland Specialty Chem",
        "EMEA",
        "NS-CORE-500",
        1,
        0.00,
    ),
    (
        "SO-26310",
        datetime.date(2026, 7, 27),
        "Clearwater BioLabs",
        "North America",
        "NS-GOV-200",
        2,
        0.10,
    ),
    (
        "SO-26311",
        datetime.date(2026, 7, 29),
        "Cordillera Mining Corp",
        "LATAM",
        "NS-EDGE-410",
        2,
        0.05,
    ),
    (
        "SO-26312",
        datetime.date(2026, 7, 31),
        "Caledonia Offshore Energy",
        "EMEA",
        "NS-EDGE-820",
        1,
        0.05,
    ),
    (
        "SO-26313",
        datetime.date(2026, 8, 3),
        "Redwood Cloud Retail",
        "North America",
        "NS-CORE-500",
        2,
        0.10,
    ),
    (
        "SO-26314",
        datetime.date(2026, 8, 5),
        "Tasman Telecom Holdings",
        "APAC",
        "NS-SEC-300",
        2,
        0.08,
    ),
    (
        "SO-26315",
        datetime.date(2026, 8, 7),
        "Aquitaine Rail Freight",
        "EMEA",
        "NS-CORE-100",
        2,
        0.05,
    ),
    (
        "SO-26316",
        datetime.date(2026, 8, 11),
        "Summit National Bancorp",
        "North America",
        "NS-SEC-300",
        3,
        0.10,
    ),
    (
        "SO-26317",
        datetime.date(2026, 8, 13),
        "Paulista Agritech S.A.",
        "LATAM",
        "NS-GOV-200",
        1,
        0.05,
    ),
    (
        "SO-26318",
        datetime.date(2026, 8, 14),
        "Hanover Industrial IoT",
        "EMEA",
        "NS-EDGE-410",
        3,
        0.10,
    ),
    (
        "SO-26319",
        datetime.date(2026, 8, 18),
        "Singapura Clearing House",
        "APAC",
        "NS-CORE-500",
        1,
        0.00,
    ),
    (
        "SO-26320",
        datetime.date(2026, 8, 20),
        "Ironbridge Semiconductor",
        "North America",
        "NS-AI-600",
        2,
        0.08,
    ),
    (
        "SO-26321",
        datetime.date(2026, 8, 24),
        "Lombardia MedTech SpA",
        "EMEA",
        "NS-AI-600",
        2,
        0.10,
    ),
    (
        "SO-26322",
        datetime.date(2026, 8, 26),
        "Boreal Timber & Pulp",
        "North America",
        "NS-EDGE-410",
        2,
        0.05,
    ),
    (
        "SO-26323",
        datetime.date(2026, 8, 28),
        "Monterrey Auto Components",
        "LATAM",
        "NS-SEC-300",
        1,
        0.05,
    ),
    (
        "SO-26324",
        datetime.date(2026, 8, 31),
        "Kansai Smart Mobility",
        "APAC",
        "NS-AI-600",
        2,
        0.10,
    ),
    (
        "SO-26325",
        datetime.date(2026, 9, 2),
        "Chesapeake Utilities",
        "North America",
        "NS-GOV-200",
        2,
        0.05,
    ),
    (
        "SO-26326",
        datetime.date(2026, 9, 4),
        "Benelux Maritime Group",
        "EMEA",
        "NS-CORE-500",
        2,
        0.12,
    ),
    (
        "SO-26327",
        datetime.date(2026, 9, 8),
        "Heartland Mutual Insurance",
        "North America",
        "NS-SVC-900",
        2,
        0.00,
    ),
    (
        "SO-26328",
        datetime.date(2026, 9, 9),
        "Seoul FinCloud Partners",
        "APAC",
        "NS-GOV-200",
        2,
        0.08,
    ),
    (
        "SO-26329",
        datetime.date(2026, 9, 11),
        "Iberia Renewables S.L.",
        "EMEA",
        "NS-GOV-200",
        2,
        0.08,
    ),
    (
        "SO-26330",
        datetime.date(2026, 9, 14),
        "Cascadia Autonomous Lab",
        "North America",
        "NS-EDGE-820",
        2,
        0.10,
    ),
    (
        "SO-26331",
        datetime.date(2026, 9, 16),
        "Rio Plata Logistics",
        "LATAM",
        "NS-CORE-100",
        1,
        0.00,
    ),
    (
        "SO-26332",
        datetime.date(2026, 9, 17),
        "Danube Pharma Distribution",
        "EMEA",
        "NS-SEC-300",
        2,
        0.10,
    ),
    (
        "SO-26333",
        datetime.date(2026, 9, 21),
        "Granite State Defense",
        "North America",
        "NS-CORE-500",
        1,
        0.00,
    ),
    (
        "SO-26334",
        datetime.date(2026, 9, 22),
        "Auckland Grid Analytics",
        "APAC",
        "NS-SVC-900",
        2,
        0.05,
    ),
    (
        "SO-26335",
        datetime.date(2026, 9, 23),
        "Solvay Precision Optics",
        "EMEA",
        "NS-SVC-900",
        2,
        0.05,
    ),
    (
        "SO-26336",
        datetime.date(2026, 9, 24),
        "Prairie Grain Exchange",
        "North America",
        "NS-CORE-100",
        2,
        0.05,
    ),
    (
        "SO-26337",
        datetime.date(2026, 9, 25),
        "Bogota Cloud Bancaria",
        "LATAM",
        "NS-SVC-900",
        1,
        0.00,
    ),
    (
        "SO-26338",
        datetime.date(2026, 9, 28),
        "Osaka Industrial Systems",
        "APAC",
        "NS-CORE-100",
        2,
        0.10,
    ),
    (
        "SO-26339",
        datetime.date(2026, 9, 29),
        "Skagerak Marine Services",
        "EMEA",
        "NS-CORE-100",
        2,
        0.08,
    ),
    (
        "SO-26340",
        datetime.date(2026, 9, 29),
        "Apex Midstream Partners",
        "North America",
        "NS-SEC-300",
        2,
        0.05,
    ),
)

REGIONAL_PLANS: tuple[tuple[str, str, int, str, int], ...] = (
    ("North America", "M. Alvarez", 520_000, "Enterprise Cloud & Security", 475_000),
    ("EMEA", "C. Lindqvist", 385_000, "Core Platform & Data Governance", 352_000),
    ("APAC", "K. Takahashi", 250_000, "Hybrid Edge & AI Telemetry", 218_000),
    ("LATAM", "R. Mendes", 115_000, "Mid-Market Platform Expansion", 98_000),
)


def build_sales_workbook() -> bytes:
    """Generate a realistic 4-sheet B2B commercial and margin review workbook."""
    wb = Workbook()

    # Corporate typography & palette
    font_title = Font(name="Segoe UI", size=14, bold=True, color="0F2942")
    font_subtitle = Font(name="Segoe UI", size=9.5, italic=True, color="52657A")
    font_section = Font(name="Segoe UI", size=10.5, bold=True, color="0F2942")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Segoe UI", size=10, color="1E293B")
    font_muted = Font(name="Segoe UI", size=9.5, color="475569")
    font_bold = Font(name="Segoe UI", size=10, bold=True, color="0F2942")
    font_input = Font(name="Segoe UI", size=10, bold=True, color="1B4F8C")

    fill_header = PatternFill(
        start_color="163A5F", end_color="163A5F", fill_type="solid"
    )
    fill_zebra = PatternFill(
        start_color="F7F9FB", end_color="F7F9FB", fill_type="solid"
    )
    fill_total = PatternFill(
        start_color="EEF3F8", end_color="EEF3F8", fill_type="solid"
    )
    fill_highlight = PatternFill(
        start_color="EBF2F9", end_color="EBF2F9", fill_type="solid"
    )
    fill_input = PatternFill(
        start_color="F2F7FD", end_color="F2F7FD", fill_type="solid"
    )

    border_thin = Border(
        left=Side(style="thin", color="D8E1EB"),
        right=Side(style="thin", color="D8E1EB"),
        top=Side(style="thin", color="D8E1EB"),
        bottom=Side(style="thin", color="D8E1EB"),
    )
    border_header = Border(
        left=Side(style="thin", color="2B547E"),
        right=Side(style="thin", color="2B547E"),
        top=Side(style="medium", color="0F2942"),
        bottom=Side(style="medium", color="0F2942"),
    )
    border_total = Border(
        left=Side(style="thin", color="D8E1EB"),
        right=Side(style="thin", color="D8E1EB"),
        top=Side(style="thin", color="163A5F"),
        bottom=Side(style="double", color="163A5F"),
    )
    border_banner = Border(bottom=Side(style="medium", color="163A5F"))

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_note = Alignment(horizontal="left", vertical="center", indent=1)
    align_right = Alignment(horizontal="right", vertical="center")

    fmt_currency = "$#,##0.00;-$#,##0.00"
    fmt_percent = "0.0%"
    fmt_signed_pct = "+0.0%;-0.0%;0.0%"
    fmt_int = "#,##0"
    fmt_date = "yyyy-mm-dd"

    def style_banner(ws, end_col: int, title: str, subtitle: str) -> None:
        end_letter = get_column_letter(end_col)
        ws.row_dimensions[1].height = 10
        ws.row_dimensions[2].height = 25
        ws.row_dimensions[3].height = 18
        ws.row_dimensions[4].height = 10
        ws.column_dimensions["A"].width = 3
        ws.merge_cells(f"B2:{end_letter}2")
        ws.merge_cells(f"B3:{end_letter}3")
        ws["B2"] = title
        ws["B2"].font = font_title
        ws["B2"].alignment = align_left
        ws["B3"] = subtitle
        ws["B3"].font = font_subtitle
        ws["B3"].alignment = align_left
        for col in range(2, end_col + 1):
            ws.cell(row=3, column=col).border = border_banner

    def style_headers(ws, row: int, specs: list[tuple[str, Alignment]]) -> None:
        ws.row_dimensions[row].height = 23
        for idx, (label, alignment) in enumerate(specs, start=2):
            cell = ws.cell(row=row, column=idx, value=label)
            cell.font = font_header
            cell.fill = fill_header
            cell.border = border_header
            cell.alignment = alignment

    # ── 1. SHEET: Summary (Executive commercial & margin overview) ──
    summary = wb.active
    summary.title = "Summary"
    summary.sheet_properties.tabColor = "163A5F"
    style_banner(
        summary,
        6,
        "Northstar Cloud — Q3 FY26 Commercial & Margin Review",
        "Executive Summary · Closed-Won Actuals vs. Q3 Plan · Currency: USD",
    )

    style_headers(
        summary,
        5,
        [
            ("Executive KPI", align_left),
            ("Q3 Actual", align_right),
            ("Governance / Context", align_note),
        ],
    )
    summary["C5"].comment = Comment(
        "Net revenue reflects Q3 closed-won contracts net of discounts; "
        "partner rebates and sales commissions are accrued in C10.",
        "FP&A — Corporate Finance",
    )

    kpis = [
        (
            "Booked Net Revenue",
            "=SUM(Sales!$J$6:$J$45)",
            fmt_currency,
            "40 closed-won B2B contracts",
            False,
        ),
        (
            "Variance vs. Revenue Plan",
            "=ROUND((C6-Assumptions!$D$16)/Assumptions!$D$16, 4)",
            fmt_signed_pct,
            "vs. $1,270,000 approved Q3 plan",
            False,
        ),
        (
            "Gross Profit",
            "=SUM(Sales!$J$6:$J$45)-SUM(Sales!$H$6:$H$45)",
            fmt_currency,
            "Net of direct cloud & OEM COGS",
            False,
        ),
        (
            "Blended Gross Margin %",
            "=ROUND(C8/C6, 4)",
            fmt_percent,
            "Above 62.0% Floor (POL-GM-62)",
            False,
        ),
        (
            "Net Commercial Contribution",
            "=ROUND(C8-C6*RebateRate-C8*Assumptions!$D$8, 2)",
            fmt_currency,
            "After 3.5% rebate & 6.5% commission",
            True,
        ),
    ]
    for offset, (metric, formula, num_fmt, note, emphasized) in enumerate(kpis):
        r = 6 + offset
        summary.row_dimensions[r].height = 21
        row_fill = (
            fill_highlight if emphasized else (fill_zebra if r % 2 == 1 else None)
        )
        row_border = border_total if emphasized else border_thin

        c_lbl = summary.cell(row=r, column=2, value=metric)
        c_lbl.font = font_bold
        c_lbl.alignment = align_left
        c_lbl.border = row_border

        c_val = summary.cell(row=r, column=3, value=formula)
        c_val.font = font_bold if emphasized else font_data
        c_val.alignment = align_right
        c_val.number_format = num_fmt
        c_val.border = row_border

        c_note = summary.cell(row=r, column=4, value=note)
        c_note.font = font_bold if str(note).startswith("=") else font_muted
        c_note.alignment = align_note
        c_note.border = row_border

        if row_fill:
            c_lbl.fill = row_fill
            c_val.fill = row_fill
            c_note.fill = row_fill

    summary.row_dimensions[11].height = 10
    summary.row_dimensions[12].height = 10
    summary.row_dimensions[13].height = 20
    summary["B13"] = "Regional Commercial Revenue & Plan Breakdown"
    summary["B13"].font = font_section
    summary["B13"].alignment = align_left

    style_headers(
        summary,
        14,
        [
            ("Region", align_left),
            ("Actual Net Revenue", align_right),
            ("Q3 Revenue Plan", align_right),
            ("Regional VP", align_note),
            ("Primary Growth Segment", align_left),
            ("FY25 Q3 Baseline", align_right),
        ],
    )

    for idx, (region, vp_name, plan_amt, segment, fy25_baseline) in enumerate(
        REGIONAL_PLANS
    ):
        r = 15 + idx
        summary.row_dimensions[r].height = 20
        row_fill = fill_zebra if idx % 2 == 1 else None

        values_and_formats = [
            (region, align_left, None, font_bold),
            (
                f"=SUMIFS(Sales!$J$6:$J$45, Sales!$E$6:$E$45, B{r})",
                align_right,
                fmt_currency,
                font_data,
            ),
            (plan_amt, align_right, fmt_currency, font_data),
            (vp_name, align_note, None, font_data),
            (segment, align_left, None, font_muted),
            (fy25_baseline, align_right, fmt_currency, font_muted),
        ]
        for col_offset, (val, align, nfmt, fnt) in enumerate(
            values_and_formats, start=2
        ):
            cell = summary.cell(row=r, column=col_offset, value=val)
            cell.font = fnt
            cell.alignment = align
            cell.border = border_thin
            if nfmt:
                cell.number_format = nfmt
            if row_fill:
                cell.fill = row_fill

    for letter, width in (
        ("B", 30),
        ("C", 20),
        ("D", 36),
        ("E", 19),
        ("F", 34),
        ("G", 18),
    ):
        summary.column_dimensions[letter].width = width
    summary.column_dimensions["G"].hidden = True
    summary.freeze_panes = "B6"

    rev_chart = BarChart()
    rev_chart.type = "col"
    rev_chart.grouping = "clustered"
    rev_chart.overlap = -15
    rev_chart.style = 10
    rev_chart.title = "Regional Net Revenue: Q3 Actual vs. Plan"
    rev_chart.y_axis.title = "USD ($)"  # ty: ignore[invalid-assignment]
    rev_chart.x_axis.title = "Region"  # ty: ignore[invalid-assignment]
    rev_data = Reference(summary, min_col=3, max_col=4, min_row=14, max_row=18)
    rev_cats = Reference(summary, min_col=2, min_row=15, max_row=18)
    rev_chart.add_data(rev_data, titles_from_data=True)
    rev_chart.set_categories(rev_cats)
    rev_chart.width = 13.6
    rev_chart.height = 5.7
    summary.add_chart(rev_chart, "E5")
    _print_wide(summary)

    # ── 2. SHEET: Sales (40 closed-won B2B enterprise orders) ──
    sales = wb.create_sheet("Sales")
    sales.sheet_properties.tabColor = "1F6B5C"
    style_banner(
        sales,
        10,
        "Q3 FY26 Closed-Won Commercial Order Ledger",
        "40 Verified Enterprise & Mid-Market Contracts · Jul–Sep 2026 · Currency: USD",
    )

    style_headers(
        sales,
        5,
        [
            ("Order ID", align_center),
            ("Close Date", align_center),
            ("Customer Account", align_left),
            ("Region", align_left),
            ("SKU", align_center),
            ("Qty", align_right),
            ("Direct Cost", align_right),
            ("Discount %", align_right),
            ("Net Revenue", align_right),
        ],
    )
    sales["I5"].comment = Comment(
        "Negotiated volume and multi-year prepay discounts reconciled against signed order forms.",
        "Sales Operations",
    )

    unit_costs = {sku: cost for sku, _n, _f, _p, cost, _bm in CATALOG_ITEMS}
    for idx, (order_id, close_dt, account, region, sku, qty, disc) in enumerate(
        SALES_ORDERS
    ):
        r = 6 + idx
        sales.row_dimensions[r].height = 19
        row_fill = fill_zebra if idx % 2 == 1 else None
        row_specs = [
            (order_id, align_center, None),
            (close_dt, align_center, fmt_date),
            (account, align_left, None),
            (region, align_left, None),
            (sku, align_center, None),
            (qty, align_right, fmt_int),
            (qty * unit_costs[sku], align_right, fmt_currency),
            (disc, align_right, fmt_percent),
            (
                f"=ROUND(G{r}*INDEX(Catalog!$E$6:$E$13, MATCH(F{r}, Catalog!$B$6:$B$13, 0))*(1-I{r}), 2)",
                align_right,
                fmt_currency,
            ),
        ]
        for col_offset, (val, align, nfmt) in enumerate(row_specs, start=2):
            cell = sales.cell(row=r, column=col_offset, value=val)
            cell.font = font_data
            cell.alignment = align
            cell.border = border_thin
            if nfmt:
                cell.number_format = nfmt
            if row_fill:
                cell.fill = row_fill

    for letter, width in (
        ("B", 13),
        ("C", 13),
        ("D", 29),
        ("E", 17),
        ("F", 15),
        ("G", 8),
        ("H", 16),
        ("I", 13),
        ("J", 18),
    ):
        sales.column_dimensions[letter].width = width
    sales.column_dimensions["H"].hidden = True
    sales.freeze_panes = "B6"
    _print_wide(sales)

    # ── 3. SHEET: Catalog (Enterprise product pricing & unit costs) ──
    cat = wb.create_sheet("Catalog")
    cat.sheet_properties.tabColor = "3B6998"
    style_banner(
        cat,
        7,
        "Enterprise Product Catalog & Unit Cost Schedule",
        "Annual Subscription & Appliance SKUs · Standard List Pricing & Direct COGS (USD)",
    )

    style_headers(
        cat,
        5,
        [
            ("SKU", align_center),
            ("Product Name", align_left),
            ("Product Family", align_left),
            ("Annual List Price", align_right),
            ("Direct Unit Cost", align_right),
            ("Billing Model", align_note),
        ],
    )
    cat["F5"].comment = Comment(
        "Direct unit cost includes cloud compute, third-party OEM royalties, and L3 support allocation.",
        "Cost Accounting",
    )

    for idx, (sku, name, family, list_price, unit_cost, billing_model) in enumerate(
        CATALOG_ITEMS
    ):
        r = 6 + idx
        cat.row_dimensions[r].height = 20
        row_fill = fill_zebra if idx % 2 == 1 else None
        row_specs = [
            (sku, align_center, None, font_bold),
            (name, align_left, None, font_data),
            (family, align_left, None, font_muted),
            (list_price, align_right, fmt_currency, font_data),
            (unit_cost, align_right, fmt_currency, font_data),
            (billing_model, align_note, None, font_muted),
        ]
        for col_offset, (val, align, nfmt, fnt) in enumerate(row_specs, start=2):
            cell = cat.cell(row=r, column=col_offset, value=val)
            cell.font = fnt
            cell.alignment = align
            cell.border = border_thin
            if nfmt:
                cell.number_format = nfmt
            if row_fill:
                cell.fill = row_fill

    for letter, width in (
        ("B", 16),
        ("C", 29),
        ("D", 19),
        ("E", 19),
        ("F", 18),
        ("G", 20),
        ("H", 3),
    ):
        cat.column_dimensions[letter].width = width
    cat.freeze_panes = "B6"

    cat_chart = BarChart()
    cat_chart.type = "bar"
    cat_chart.grouping = "clustered"
    cat_chart.overlap = -10
    cat_chart.style = 10
    cat_chart.title = "Annual List Price vs. Direct Cost by SKU"
    cat_chart.y_axis.title = "USD ($)"  # ty: ignore[invalid-assignment]
    cat_chart.x_axis.title = "SKU"  # ty: ignore[invalid-assignment]
    cat_data = Reference(cat, min_col=5, max_col=6, min_row=5, max_row=13)
    cat_cats = Reference(cat, min_col=2, min_row=6, max_row=13)
    cat_chart.add_data(cat_data, titles_from_data=True)
    cat_chart.set_categories(cat_cats)
    cat_chart.width = 13.5
    cat_chart.height = 7.0
    cat.add_chart(cat_chart, "I5")
    _print_wide(cat)

    # ── 4. SHEET: Assumptions (Commercial policy & Q3 regional plan) ──
    params = wb.create_sheet("Assumptions")
    params.sheet_properties.tabColor = "6C7A89"
    style_banner(
        params,
        5,
        "Northstar Cloud — Q3 FY26 Commercial Plan & Margin Policy",
        "Approved FP&A Baseline · Effective Jul 1 – Sep 30, 2026 · Currency: USD",
    )

    style_headers(
        params,
        5,
        [
            ("Policy Parameter", align_left),
            ("Policy Code", align_center),
            ("Approved Rate", align_right),
            ("Governance Scope", align_note),
        ],
    )

    policies = [
        (
            "Blended Gross Margin Floor",
            "POL-GM-62",
            0.62,
            "Board minimum quarterly blended gross margin",
        ),
        (
            "Partner Channel Rebate Rate",
            "POL-REB-35",
            0.035,
            "Accrued on booked net revenue across regions",
        ),
        (
            "Sales Commission Rate",
            "POL-COM-65",
            0.065,
            "Variable compensation accrued on gross profit",
        ),
        (
            "Standard Discount Cap",
            "POL-DSC-15",
            0.15,
            "Discounts above 15.0% require deal-desk escalation",
        ),
    ]
    for idx, (label, code, rate, scope_note) in enumerate(policies):
        r = 6 + idx
        params.row_dimensions[r].height = 20
        c_lbl = params.cell(row=r, column=2, value=label)
        c_lbl.font = font_bold
        c_lbl.alignment = align_left
        c_lbl.border = border_thin

        c_code = params.cell(row=r, column=3, value=code)
        c_code.font = font_muted
        c_code.alignment = align_center
        c_code.border = border_thin

        c_val = params.cell(row=r, column=4, value=rate)
        c_val.font = font_input
        c_val.fill = fill_input
        c_val.alignment = align_right
        c_val.number_format = fmt_percent
        c_val.border = border_thin

        c_note = params.cell(row=r, column=5, value=scope_note)
        c_note.font = font_muted
        c_note.alignment = align_note
        c_note.border = border_thin

    params["D6"].comment = Comment(
        "Approved at Q3 Operating Committee (ref. FIN-2026-Q3). Changes require VP Finance sign-off.",
        "Corporate FP&A",
    )

    params.row_dimensions[10].height = 12
    style_headers(
        params,
        11,
        [
            ("Region", align_left),
            ("Regional VP", align_left),
            ("Q3 Revenue Plan", align_right),
            ("Primary Growth Segment", align_note),
        ],
    )

    for idx, (region, vp_name, plan_amt, segment, _fy25) in enumerate(REGIONAL_PLANS):
        r = 12 + idx
        params.row_dimensions[r].height = 20
        c_reg = params.cell(row=r, column=2, value=region)
        c_reg.font = font_bold
        c_reg.alignment = align_left
        c_reg.border = border_thin

        c_vp = params.cell(row=r, column=3, value=vp_name)
        c_vp.font = font_data
        c_vp.alignment = align_left
        c_vp.border = border_thin

        c_plan = params.cell(row=r, column=4, value=plan_amt)
        c_plan.font = font_input
        c_plan.fill = fill_input
        c_plan.alignment = align_right
        c_plan.number_format = fmt_currency
        c_plan.border = border_thin

        c_seg = params.cell(row=r, column=5, value=segment)
        c_seg.font = font_muted
        c_seg.alignment = align_note
        c_seg.border = border_thin

    params.row_dimensions[16].height = 22
    total_cells = [
        ("Total Q3 Revenue Plan", align_left, None, font_bold),
        ("Global Commercial", align_left, None, font_muted),
        ("=SUM(D12:D15)", align_right, fmt_currency, font_bold),
        ("Approved Jun 25, 2026", align_note, None, font_muted),
    ]
    for col_offset, (val, align, nfmt, fnt) in enumerate(total_cells, start=2):
        cell = params.cell(row=16, column=col_offset, value=val)
        cell.font = fnt
        cell.fill = fill_total
        cell.alignment = align
        cell.border = border_total
        if nfmt:
            cell.number_format = nfmt

    for letter, width in (("B", 31), ("C", 20), ("D", 20), ("E", 46)):
        params.column_dimensions[letter].width = width
    params.freeze_panes = "B6"
    _print_wide(params)

    wb.defined_names.add(DefinedName("RebateRate", attr_text="Assumptions!$D$7"))

    return _to_bytes(wb)


# ──────────────────────────────────────────────
# 2. Stress workbook — the hostile one
# ──────────────────────────────────────────────

#: Rows in the stretched table. Large enough that grouping must collapse it to a
#: single node, small enough that LibreOffice still renders it quickly.
STRESS_ROWS = 250

#: Strings that must survive extraction, JSON embedding and HTML rendering
#: unchanged. Each one has broken something in some report generator: the first
#: two close the script tag the graph is embedded in, the next two collide with
#: the viewer's own template placeholders, ``$&`` is a regex replacement
#: reference, and U+2028 is a newline to a JavaScript string literal but not to
#: JSON.
HOSTILE_STRINGS = (
    "</script><script>alert(1)</script>",
    '"><img src=x onerror=alert(1)>',
    "__TITLE__",
    "__GRAPH_JSON__",
    "value is $& and $1",
    "line\u2028separator",
    "quote\" apostrophe' backslash\\ tab\t",
    "=cmd|'/c calc'!A1",  # CSV-injection shape, harmless as a stored string
)


def build_stress_workbook() -> bytes:
    """Generate an English workbook that attacks the analysis pipeline.

    Every sheet is one category of problem, so a report that renders all of them
    has exercised: error propagation, unresolvable references, quoting, deep and
    oversized formulas, hidden and empty sheets, and HTML escaping.
    """
    wb = Workbook()
    _stress_data(wb)
    _stress_lookups(wb)
    _stress_errors(wb)
    _stress_volatile(wb)
    _stress_cross_refs(wb)
    _stress_quoting(wb)
    _stress_hidden(wb)
    _stress_hostile(wb)
    wb.create_sheet("Empty Sheet")  # no cells at all: every bound is degenerate
    _stress_names(wb)
    return _to_bytes(wb)


def _stress_data(wb: Workbook) -> None:
    """A wide table of mixed types, and the stretched formulas over it."""
    ws = wb.active
    ws.title = "Sales Data"

    headers = ["Region", "Rep", "Booked", "Units", "Price", "Revenue", "Flag"]
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=text)
        cell.font = Font(bold=True)

    epoch = datetime.date(2026, 1, 1)
    for r in range(2, STRESS_ROWS + 2):
        ws.cell(row=r, column=1, value=["North", "South", "East", "West"][r % 4])
        ws.cell(row=r, column=2, value=f"Rep {r - 1:03d}")
        ws.cell(row=r, column=3, value=epoch + datetime.timedelta(days=r)).number_format
        ws.cell(row=r, column=3).number_format = "yyyy-mm-dd"
        # Every 10th row is blank and every 17th holds a number stored as text:
        # both make an aggregate disagree with a naive row count.
        if r % 10 == 0:
            pass
        elif r % 17 == 0:
            ws.cell(row=r, column=4, value=str(r % 9 + 1))
        else:
            ws.cell(row=r, column=4, value=r % 9 + 1)
        ws.cell(row=r, column=5, value=round(4.25 + (r % 23) * 1.5, 2))
        ws.cell(row=r, column=6, value=f"=IFERROR(D{r}*E{r}, 0)")
        ws.cell(row=r, column=7, value=f'=IF(F{r}>60, "HIGH", "LOW")')

    # Dates before 1900-03-01, where Excel's deliberate leap-year bug puts the
    # serial one day off every other implementation.
    ws["I1"] = "Lotus-bug window"
    for offset, row in enumerate(range(2, 8)):
        ws.cell(row=row, column=9, value=datetime.date(1900, 1, 1 + offset))
        ws.cell(row=row, column=9).number_format = "yyyy-mm-dd"

    ws["K1"] = "Whole-column aggregate"
    ws["K2"] = "=SUM(F:F)"
    ws["K3"] = "=COUNT(D:D)"
    ws["K4"] = "=COUNTA(D:D)"  # disagrees with K3: the text-numbers count here

    # A real table, so the structured reference on 'Cross Refs' has a referent.
    table = Table(displayName="SalesTable", ref=f"A1:G{STRESS_ROWS + 1}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(table)

    ws.freeze_panes = "B2"
    ws.row_dimensions[3].hidden = True
    ws.column_dimensions["E"].hidden = True
    ws.merge_cells("M1:P1")
    ws["M1"] = "Merged banner spanning four columns"
    ws["A1"].comment = Comment(
        "Region codes come from the CRM export, not from finance. "
        "Réf. ticket #4821 — вопрос про кодировку.",
        "Data Steward",
    )
    ws.add_data_validation(_units_validation(ws))
    ws.conditional_formatting.add(
        f"F2:F{STRESS_ROWS + 1}",
        CellIsRule(operator="greaterThan", formula=["60"], font=Font(bold=True)),
    )
    _print_wide(ws)


def _units_validation(ws) -> DataValidation:
    validation = DataValidation(
        type="whole", operator="between", formula1=1, formula2=9
    )
    validation.add(f"D2:D{STRESS_ROWS + 1}")
    return validation


def _stress_lookups(wb: Workbook) -> None:
    """Lookup and aggregate functions, plus formulas built to be too big."""
    ws = wb.create_sheet("Lookups")
    ws["A1"] = "Function"
    ws["B1"] = "Result"
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)

    cases = [
        ("VLOOKUP exact", "=VLOOKUP(\"Rep 007\", 'Sales Data'!B:F, 5, FALSE)"),
        ("VLOOKUP missing", "=VLOOKUP(\"Rep 999\", 'Sales Data'!B:F, 5, FALSE)"),
        (
            "INDEX/MATCH",
            "=INDEX('Sales Data'!F:F, MATCH(\"Rep 012\", 'Sales Data'!B:B, 0))",
        ),
        (
            "XLOOKUP",
            "=XLOOKUP(\"Rep 020\", 'Sales Data'!B:B, 'Sales Data'!F:F, \"none\")",
        ),
        ("SUMIFS", "=SUMIFS('Sales Data'!F:F, 'Sales Data'!A:A, \"North\")"),
        (
            "COUNTIFS",
            "=COUNTIFS('Sales Data'!A:A, \"East\", 'Sales Data'!G:G, \"HIGH\")",
        ),
        ("SUMPRODUCT", "=SUMPRODUCT('Sales Data'!D2:D100, 'Sales Data'!E2:E100)"),
        ("TEXTJOIN", "=TEXTJOIN(\", \", TRUE, 'Sales Data'!A2:A6)"),
        (
            "LET",
            "=LET(total, SUM('Sales Data'!F:F), n, COUNT('Sales Data'!F:F), total/n)",
        ),
        ("Array constant", "=SUM({1,2,3;4,5,6})"),
        ("Intersection", "=SUM('Sales Data'!D2:D10 'Sales Data'!D5:D20)"),
    ]
    for row, (label, formula) in enumerate(cases, start=2):
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=formula)

    ws["A20"] = "Nested IF, 30 deep"
    ws["B20"] = _nested_if(30)
    ws["A21"] = "Formula past the dossier limit"
    ws["B21"] = _oversized_sum(700)
    ws["A22"] = "Deeply parenthesised"
    ws["B22"] = "=" + "(" * 40 + "1+2" + ")" * 40
    ws.column_dimensions["A"].width = 32
    _print_wide(ws)


def _nested_if(depth: int) -> str:
    """``=IF(A1>1, IF(A1>2, … , 0), 0)`` — a chain the decomposer must walk."""
    formula = "0"
    for level in range(depth, 0, -1):
        formula = f"IF('Sales Data'!K2>{level}, {formula}, {level})"
    return "=" + formula


def _oversized_sum(terms: int) -> str:
    """A formula whose dossier exceeds ``aidoc.MAX_DOSSIER_CHARS``.

    With 700 terms this deliberately exceeds Excel's 8,192-character limit.
    It tests hostile OOXML and bounded AI dossiers, not Excel compatibility.
    """
    return "=" + "+".join(
        f"ROUND('Sales Data'!E{i % 200 + 2}, 2)" for i in range(terms)
    )


def _stress_errors(wb: Workbook) -> None:
    """One cell per Excel error value, and the guards that swallow them."""
    ws = wb.create_sheet("Errors")
    ws["A1"] = "Error"
    ws["B1"] = "Raw"
    ws["C1"] = "Guarded"
    for cell in ("A1", "B1", "C1"):
        ws[cell].font = Font(bold=True)

    cases = [
        ("#DIV/0!", "=1/0"),
        ("#VALUE!", '="text"+1'),
        ("#N/A", "=NA()"),
        ("#NAME?", "=NoSuchFunction(1)"),
        ("#NUM!", "=SQRT(-1)"),
        ("#NULL!", "=SUM(A1:A2 B5:B6)"),
        ("#REF!", "=#REF!"),
    ]
    for row, (label, formula) in enumerate(cases, start=2):
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=formula)
        ws.cell(row=row, column=3, value=f'=IFERROR({formula[1:]}, "handled")')

    ws["A12"] = "Aggregate over errors"
    ws["B12"] = "=SUM(B2:B8)"  # itself an error: errors propagate through SUM
    ws["B13"] = "=AGGREGATE(9, 6, B2:B8)"  # 6 = ignore errors
    ws.column_dimensions["A"].width = 24
    _print_wide(ws)


def _stress_volatile(wb: Workbook) -> None:
    """Functions whose value depends on when — or on a string built at runtime."""
    ws = wb.create_sheet("Volatile")
    rows = [
        ("TODAY", "=TODAY()"),
        ("NOW", "=NOW()"),
        ("RAND", "=RAND()"),
        ("RANDBETWEEN", "=RANDBETWEEN(1, 100)"),
        # INDIRECT and OFFSET build their target at evaluation time, so no static
        # analysis can draw the edge — these must land as unresolved, not wrong.
        ("INDIRECT literal", "=INDIRECT(\"'Sales Data'!F2\")"),
        ("INDIRECT computed", "=INDIRECT(\"'Sales Data'!F\" & 2 + 3)"),
        ("OFFSET", "=OFFSET('Sales Data'!A1, 5, 5)"),
        ("INDIRECT of a name", '=INDIRECT("TaxRate")'),
    ]
    ws["A1"] = "Function"
    ws["B1"] = "Value"
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    for row, (label, formula) in enumerate(rows, start=2):
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=formula)
    ws.column_dimensions["A"].width = 22
    _print_wide(ws)


def _stress_cross_refs(wb: Workbook) -> None:
    """References that leave the sheet, the workbook, or logic itself."""
    ws = wb.create_sheet("Cross Refs")
    rows = [
        ("3-D reference", "=SUM('Sales Data:Lookups'!A1)"),
        ("External workbook", "='[Budget FY26.xlsx]Annual'!$B$4"),
        ("External, unopened path", "='C:\\Reports\\[Old.xlsx]Sheet1'!A1"),
        ("Structured reference", "=SUM(SalesTable[Revenue])"),
        ("Structured, this row", "=SalesTable[[#This Row],[Units]]"),
        ("Quoted sheet name", "='O''Brien''s Café'!B2"),
        ("Defined name", "=TaxRate"),
        ("Sheet-scoped name", "=LocalLimit"),
        ("Name holding a constant", "=CompanyName"),
        ("Name holding a formula", "=RevenueColumn"),
    ]
    ws["A1"] = "Reference kind"
    ws["B1"] = "Formula"
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    for row, (label, formula) in enumerate(rows, start=2):
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=formula)

    # A circular pair. Excel reports it rather than evaluating; the analyser must
    # come back with something instead of recursing until it dies.
    ws["A15"] = "Circular pair"
    ws["B15"] = "=B16+1"
    ws["B16"] = "=B15+1"
    ws["A17"] = "Self reference"
    ws["B17"] = "=B17"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 40
    _print_wide(ws)


def _stress_quoting(wb: Workbook) -> None:
    """A sheet whose own name has to be quoted everywhere it is mentioned."""
    ws = wb.create_sheet("O'Brien's Café")
    ws["A1"] = "Sheet names carry apostrophes, spaces and accents"
    ws["A1"].font = Font(bold=True)
    ws["B2"] = 42
    ws["B3"] = "=B2*2"
    ws["B4"] = "=SUM('Sales Data'!F2:F10) + B3"
    ws["A6"] = "Referenced from 'Cross Refs' as 'O''Brien''s Café'!B2"
    ws.column_dimensions["A"].width = 50
    _print_wide(ws)


def _stress_hidden(wb: Workbook) -> None:
    """Sheets the reader never sees, which still feed the visible ones."""
    config = wb.create_sheet("Hidden Config")
    config["A1"] = "Setting"
    config["B1"] = "Value"
    config["A2"] = "Tax rate"
    config["B2"] = 0.196
    config["A3"] = "Local limit"
    config["B3"] = 250_000
    config.sheet_state = "hidden"

    archive = wb.create_sheet("Very Hidden Archive")
    archive["A1"] = "Only reachable through the VBA project editor"
    archive["A2"] = "=SUM('Sales Data'!F:F)"
    archive.sheet_state = "veryHidden"


def _stress_hostile(wb: Workbook) -> None:
    """Cell text, comments and a defined name aimed at the HTML escaping.

    None of these should ever execute. If any does, the failure is visible as a
    dialog or a broken report rather than as a silent wrong answer.
    """
    ws = wb.create_sheet("Hostile Text")
    ws["A1"] = "Payload"
    ws["B1"] = "Stored"
    ws["C1"] = "Concatenated into a formula"
    for cell in ("A1", "B1", "C1"):
        ws[cell].font = Font(bold=True)

    for row, payload in enumerate(HOSTILE_STRINGS, start=2):
        ws.cell(row=row, column=1, value=f"case {row - 1}")
        ws.cell(row=row, column=2, value=payload)
        ws.cell(row=row, column=3, value=f'=B{row} & " (copied)"')

    ws["A12"] = "Very long single value"
    ws["B12"] = "long " * 2_000
    ws["A13"] = "Right-to-left text"
    ws["B13"] = "מכירות רבעוניות"
    ws["A14"] = "Emoji and combining marks"
    ws["B14"] = "revenue 📈 café e\u0301"

    ws["B2"].comment = Comment(
        "A comment that also closes the tag: </script><b>bold</b>",
        "</script>",
    )
    ws.column_dimensions["B"].width = 46
    _print_wide(ws)


def _stress_names(wb: Workbook) -> None:
    """Defined names of every shape the graph has to resolve."""
    wb.defined_names.add(DefinedName("TaxRate", attr_text="'Hidden Config'!$B$2"))
    wb.defined_names.add(DefinedName("CompanyName", attr_text='"Contoso Ltd"'))
    wb.defined_names.add(
        DefinedName("RevenueColumn", attr_text="'Sales Data'!$F$2:$F$251")
    )
    wb.defined_names.add(DefinedName("Café_Total", attr_text="'O''Brien''s Café'!$B$4"))
    # A name pointing at a range that no longer exists.
    wb.defined_names.add(DefinedName("BrokenName", attr_text="#REF!$A$1"))
    # Sheet-scoped: same short name could exist on another sheet with another
    # target, so resolution has to consider scope and not just the name.
    cross_refs = wb["Cross Refs"]
    cross_refs.defined_names.add(
        DefinedName("LocalLimit", attr_text="'Hidden Config'!$B$3")
    )


# ──────────────────────────────────────────────
# Shared helpers
# ──────────────────────────────────────────────


def _print_wide(ws) -> None:
    """Landscape A3, fitted to width — keeps the rendered PNGs legible."""
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _to_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
