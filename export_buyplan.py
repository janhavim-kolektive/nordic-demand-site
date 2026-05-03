"""
export_buyplan.py
=================
Export filtered buy plan to Excel for vendor orders.
Run in Cloud Shell from ~/nordic_site/

Usage examples:
  python3 export_buyplan.py
      --brands "CAT,Lambretta"
      --months "2026-11,2026-12"
      --styles "COLORADO,LAM0025"
      --min_qty 1

  python3 export_buyplan.py --brands "Original Penguin" --months "2026-11"
  python3 export_buyplan.py --brands "CAT,Rockport Comfort" --months "2026-11,2026-12"
"""

import argparse
import sys
import os
from datetime import datetime

os.system("pip install openpyxl google-cloud-bigquery db-dtypes pyarrow --break-system-packages -q")

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from google.cloud import bigquery

PROJECT = "inventory-planning-model"
DATASET = "Model"

# ── Args ──────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(description="Export filtered buy plan")
parser.add_argument("--brands",  default="", help="Comma-separated brands e.g. 'CAT,Lambretta'")
parser.add_argument("--months",  default="", help="Comma-separated months e.g. '2026-11,2026-12'")
parser.add_argument("--styles",  default="", help="Comma-separated styles e.g. 'COLORADO,LAM0025'")
parser.add_argument("--min_qty", default=1, type=int, help="Minimum buy_qty to include")
parser.add_argument("--level",   default="sku", choices=["sku","style"],
                    help="sku = colour+size detail, style = aggregated")
args = parser.parse_args()

# ── BigQuery client ───────────────────────────────────────────────
client = bigquery.Client(project=PROJECT)

def make_where():
    parts = [f"buy_qty >= {args.min_qty}"]
    if args.brands:
        bl = [f"'{b.strip()}'" for b in args.brands.split(",") if b.strip()]
        parts.append(f"brand IN ({','.join(bl)})")
    if args.months:
        # year_month stored as YYYY-MM-01 or YYYY-MM — handle both
        ml = []
        for m in args.months.split(","):
            m = m.strip()
            if m:
                ml.append(f"'{m}-01'")
                ml.append(f"'{m}'")
        parts.append(f"year_month IN ({','.join(ml)})")
    if args.styles:
        sl = [f"'{s.strip()}'" for s in args.styles.split(",") if s.strip()]
        parts.append(f"style IN ({','.join(sl)})")
    return "WHERE " + " AND ".join(parts)

where = make_where()
print(f"\nFilter: {where}")

# ── Query ─────────────────────────────────────────────────────────
if args.level == "style":
    sql = f"""
        SELECT
            customer, brand, style,
            SUBSTR(year_month, 1, 7)          AS month,
            ROUND(SUM(forecast_p50), 0)       AS forecast_p50,
            ROUND(SUM(forecast_p90), 0)       AS forecast_p90,
            ROUND(SUM(buy_qty), 0)            AS buy_qty,
            MAX(moq)                          AS moq,
            MAX(CAST(new_style_flag AS INT64))    AS new_style,
            MAX(CAST(below_moq_flag AS INT64))    AS below_moq,
            MAX(CAST(zero_forecast_flag AS INT64)) AS zero_forecast
        FROM `{PROJECT}`.`{DATASET}`.`buy_plan`
        {where}
        GROUP BY customer, brand, style, SUBSTR(year_month,1,7)
        ORDER BY brand, style, customer, month
    """
else:
    sql = f"""
        SELECT
            customer, brand, style, colour, size,
            SUBSTR(year_month, 1, 7)          AS month,
            ROUND(forecast_p50, 0)            AS forecast_p50,
            ROUND(forecast_p90, 0)            AS forecast_p90,
            ROUND(buy_qty, 0)                 AS buy_qty,
            moq,
            safety_stock,
            CAST(new_style_flag AS INT64)     AS new_style,
            CAST(below_moq_flag AS INT64)     AS below_moq,
            CAST(zero_forecast_flag AS INT64) AS zero_forecast
        FROM `{PROJECT}`.`{DATASET}`.`buy_plan`
        {where}
        ORDER BY brand, style, colour, size, customer, month
    """

print("Querying BigQuery…")
try:
    df = client.query(sql).to_dataframe()
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)

if df.empty:
    print("No rows returned — check your filters. Brands and styles are case-sensitive.")
    sys.exit(0)

print(f"Rows returned: {len(df):,}")

# ── Build Excel ───────────────────────────────────────────────────
wb = Workbook()

# ── Sheet 1: Buy Plan ─────────────────────────────────────────────
ws = wb.active
ws.title = "Buy Plan"

# Colours
NAVY   = "0A1628"
TEAL   = "0D7377"
GOLD   = "D4A017"
LIGHT  = "EEF4FB"
WHITE  = "FFFFFF"
RED_BG = "FCEBEB"
AMB_BG = "FAEEDA"
GRN_BG = "EAF3DE"

navy_fill  = PatternFill("solid", fgColor=NAVY)
teal_fill  = PatternFill("solid", fgColor=TEAL)
light_fill = PatternFill("solid", fgColor=LIGHT)
red_fill   = PatternFill("solid", fgColor="FFC7CE")
amb_fill   = PatternFill("solid", fgColor="FFEB9C")
grn_fill   = PatternFill("solid", fgColor="C6EFCE")
white_fill = PatternFill("solid", fgColor=WHITE)

thin = Border(
    left=Side(style='thin', color="C5D5E8"),
    right=Side(style='thin', color="C5D5E8"),
    top=Side(style='thin', color="C5D5E8"),
    bottom=Side(style='thin', color="C5D5E8")
)

# ── Header block ─────────────────────────────────────────────────
ws.merge_cells("A1:J1")
ws["A1"] = "NORDIC KOLEKTIVE VENTURES · DEMAND INTELLIGENCE"
ws["A1"].font = Font(name="Arial", size=14, bold=True, color=WHITE)
ws["A1"].fill = navy_fill
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 30

ws.merge_cells("A2:J2")
brands_str  = args.brands  if args.brands  else "All brands"
months_str  = args.months  if args.months  else "All months"
styles_str  = args.styles  if args.styles  else "All styles"
ws["A2"] = (f"Buy Plan Export  |  Brands: {brands_str}  |  "
            f"Months: {months_str}  |  Styles: {styles_str}  |  "
            f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}")
ws["A2"].font = Font(name="Arial", size=10, color=WHITE)
ws["A2"].fill = teal_fill
ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 20

ws.append([])  # blank row 3

# ── Column headers ────────────────────────────────────────────────
if args.level == "style":
    headers = ["Customer","Brand","Style","Month",
               "Forecast P50","Forecast P90","Buy Qty","MOQ",
               "New Style","Below MOQ","Zero Fcst"]
else:
    headers = ["Customer","Brand","Style","Colour","Size","Month",
               "Forecast P50","Forecast P90","Buy Qty","MOQ","Safety Stock",
               "New Style","Below MOQ","Zero Fcst"]

ws.append(headers)
hrow = ws.max_row
for i, h in enumerate(headers, 1):
    cell = ws.cell(row=hrow, column=i)
    cell.value = h
    cell.fill = navy_fill
    cell.font = Font(name="Arial", size=10, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin
ws.row_dimensions[hrow].height = 22

# ── Data rows ─────────────────────────────────────────────────────
for _, row in df.iterrows():
    if args.level == "style":
        values = [
            row.get("customer",""), row.get("brand",""), row.get("style",""),
            row.get("month",""),
            int(row.get("forecast_p50") or 0), int(row.get("forecast_p90") or 0),
            int(row.get("buy_qty") or 0), int(row.get("moq") or 0),
            "Yes" if row.get("new_style") else "",
            "Yes" if row.get("below_moq") else "",
            "Yes" if row.get("zero_forecast") else "",
        ]
    else:
        values = [
            row.get("customer",""), row.get("brand",""), row.get("style",""),
            row.get("colour",""), row.get("size",""), row.get("month",""),
            int(row.get("forecast_p50") or 0), int(row.get("forecast_p90") or 0),
            int(row.get("buy_qty") or 0), int(row.get("moq") or 0),
            int(row.get("safety_stock") or 0),
            "Yes" if row.get("new_style") else "",
            "Yes" if row.get("below_moq") else "",
            "Yes" if row.get("zero_forecast") else "",
        ]
    ws.append(values)
    r = ws.max_row

    # Alternate row fill
    fill = light_fill if r % 2 == 0 else white_fill
    for c in range(1, len(headers)+1):
        cell = ws.cell(row=r, column=c)
        cell.fill = fill
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(vertical="center")
        cell.border = thin

    # Flag colouring — last 3 cols
    n = len(headers)
    if str(values[-3]) == "Yes":  # new style
        ws.cell(row=r, column=n-2).fill = PatternFill("solid", fgColor="E6F1FB")
    if str(values[-2]) == "Yes":  # below moq
        ws.cell(row=r, column=n-1).fill = amb_fill
    if str(values[-1]) == "Yes":  # zero forecast
        ws.cell(row=r, column=n).fill = red_fill

# ── Column widths ─────────────────────────────────────────────────
col_widths = {
    "sku":   [28,18,14,14,6,10,14,14,10,8,12,11,11,11],
    "style": [28,18,14,10,14,14,10,8,11,11,11]
}
for i, w in enumerate(col_widths[args.level], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.freeze_panes = f"A{hrow+1}"

# ── Sheet 2: Summary ──────────────────────────────────────────────
ws2 = wb.create_sheet("Brand Summary")
ws2.merge_cells("A1:F1")
ws2["A1"] = "BRAND SUMMARY"
ws2["A1"].font = Font(name="Arial", size=13, bold=True, color=WHITE)
ws2["A1"].fill = navy_fill
ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[1].height = 28

summary = (df.groupby("brand").agg(
    total_buy=("buy_qty","sum"),
    total_p50=("forecast_p50","sum"),
    rows=("buy_qty","count"),
    months=("month","nunique")
).reset_index().sort_values("total_buy", ascending=False))

summary["ratio"] = (summary["total_buy"] / summary["total_p50"].replace(0,1)).round(2)

sh = ["Brand","Total Buy Qty","Total P50","Buy/P50 Ratio","Months","Lines"]
ws2.append(sh)
hr2 = ws2.max_row
for i, h in enumerate(sh, 1):
    c = ws2.cell(row=hr2, column=i)
    c.value = h
    c.fill = teal_fill
    c.font = Font(name="Arial", size=10, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center")
    c.border = thin

for _, row in summary.iterrows():
    vals = [row["brand"], int(row["total_buy"]), int(row["total_p50"]),
            float(row["ratio"]), int(row["months"]), int(row["rows"])]
    ws2.append(vals)
    r = ws2.max_row
    for ci, v in enumerate(vals, 1):
        cell = ws2.cell(row=r, column=ci)
        cell.fill = light_fill if r%2==0 else white_fill
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(vertical="center")
        cell.border = thin
    # Ratio colouring
    ratio = float(row["ratio"])
    rc = ws2.cell(row=r, column=4)
    if 1.0 <= ratio <= 2.5: rc.fill = grn_fill
    elif ratio > 2.5: rc.fill = amb_fill
    else: rc.fill = red_fill

for i, w in enumerate([28,15,13,14,10,8], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

ws2.freeze_panes = "A3"

# ── Sheet 3: Flag Guide ───────────────────────────────────────────
ws3 = wb.create_sheet("Flag Guide")
ws3.merge_cells("A1:C1")
ws3["A1"] = "FLAG REFERENCE GUIDE"
ws3["A1"].font = Font(name="Arial", size=12, bold=True, color=WHITE)
ws3["A1"].fill = navy_fill
ws3["A1"].alignment = Alignment(horizontal="center")
ws3.row_dimensions[1].height = 25
ws3.append(["Flag","Meaning","Buyer Action"])
for i in range(1,4):
    c = ws3.cell(row=2, column=i)
    c.fill = teal_fill
    c.font = Font(name="Arial", size=10, bold=True, color=WHITE)
    c.border = thin

flags = [
    ("New Style = Yes","Style has fewer than 3 months of sales history","Treat forecast as indicative — wide uncertainty"),
    ("Below MOQ = Yes","Forecast is below the vendor minimum order quantity","Decide: order MOQ anyway or skip this line"),
    ("Zero Fcst = Yes","Model predicts zero demand for this period","Skip unless you have specific pipeline knowledge"),
    ("Ratio 1.0–2.5×","Buy quantity is 1–2.5× the P50 forecast","Healthy range — trust the model"),
    ("Ratio > 2.5×","Over-ordering vs forecast","Review before committing to vendor"),
]
for fi, (f, m, a) in enumerate(flags):
    ws3.append([f, m, a])
    r = ws3.max_row
    fills = [amb_fill, red_fill, red_fill, grn_fill, amb_fill]
    for ci in range(1,4):
        c = ws3.cell(row=r, column=ci)
        c.fill = fills[fi]
        c.font = Font(name="Arial", size=10)
        c.border = thin

for i, w in enumerate([22,45,42], 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# ── Save ──────────────────────────────────────────────────────────
stamp = datetime.now().strftime("%Y%m%d_%H%M")
brands_safe = args.brands.replace(",","_").replace(" ","").replace("'","")[:30] if args.brands else "ALL"
months_safe = args.months.replace(",","_").replace(" ","") if args.months else "ALL"
fname = f"BuyPlan_{brands_safe}_{months_safe}_{stamp}.xlsx"

wb.save(fname)
print(f"\n✓ Saved: {fname}")
print(f"  Rows: {len(df):,}")
print(f"  Brands: {df['brand'].nunique()}")
print(f"  Months: {df['month'].nunique()}")
print(f"  Total buy qty: {int(df['buy_qty'].sum()):,}")
print(f"\nDownload with:")
print(f"  cloudshell download ~/{fname}")
