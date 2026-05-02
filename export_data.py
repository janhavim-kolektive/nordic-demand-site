"""
export_data.py — Export BigQuery data to JSON for Netlify site
Run after every pipeline: python3 export_data.py
"""

import json, os
from datetime import datetime
from google.cloud import bigquery

PROJECT = "inventory-planning-model"
DATASET = "Model"
T = lambda name: f"`{PROJECT}`.`{DATASET}`.`{name}`"

client = bigquery.Client(project=PROJECT)

def bq(sql):
    return client.query(sql).to_dataframe()

os.makedirs("public/data", exist_ok=True)
print("Exporting data from BigQuery...")

# ── 1. FORECAST ────────────────────────────────────────────────────
print("  [1/5] Forecast data...")
brand_monthly = bq(f"""
    SELECT brand, year_month,
           ROUND(SUM(p10), 0)          AS p10,
           ROUND(SUM(p50_ensemble), 0) AS p50,
           ROUND(SUM(p90), 0)          AS p90
    FROM {T('forecasts_ensemble_reconciled')}
    GROUP BY brand, year_month
    ORDER BY brand, year_month
""")
brand_list = bq(f"""
    SELECT DISTINCT brand
    FROM {T('forecasts_ensemble_reconciled')}
    ORDER BY brand
""")
with open("public/data/forecast.json", "w") as f:
    json.dump({
        "brands": brand_list["brand"].tolist(),
        "monthly": brand_monthly.to_dict(orient="records")
    }, f)

# ── 2. BUY PLAN ────────────────────────────────────────────────────
print("  [2/5] Buy plan data...")
brand_summary = bq(f"""
    SELECT brand,
           ROUND(SUM(buy_qty)/12.0, 0)      AS buy_per_month,
           ROUND(SUM(forecast_p50)/12.0, 0) AS p50_per_month,
           ROUND(SUM(safety_stock)/12.0, 0) AS safety_per_month,
           ROUND(SAFE_DIVIDE(SUM(buy_qty), NULLIF(SUM(forecast_p50),0)), 2) AS ratio,
           COUNTIF(buy_qty > 0)             AS active_lines,
           COUNTIF(below_moq_flag = 1)      AS below_moq_count
    FROM {T('buy_plan')}
    GROUP BY brand
    ORDER BY buy_per_month DESC
    LIMIT 20
""")
bp_kpis = bq(f"""
    SELECT
        SUM(buy_qty)              AS total_buy,
        SUM(forecast_p50)         AS total_p50,
        SUM(safety_stock)         AS total_safety,
        COUNTIF(buy_qty > 0)      AS actionable_lines,
        COUNTIF(below_moq_flag=1) AS below_moq_lines,
        COUNT(*)                  AS total_lines
    FROM {T('buy_plan')}
""")
with open("public/data/buyplan.json", "w") as f:
    json.dump({
        "kpis": bp_kpis.to_dict(orient="records")[0],
        "brand_summary": brand_summary.to_dict(orient="records")
    }, f)

# ── 3. RECOMMENDATIONS ─────────────────────────────────────────────
print("  [3/5] Recommendations data...")

reorder = bq(f"""
    SELECT customer, brand,
           last_order_month,
           months_overdue,
           avg_order_qty,
           priority
    FROM {T('recommendations_reorder')}
    WHERE priority IN ('URGENT','DUE')
    ORDER BY months_overdue DESC
    LIMIT 50
""")

crosssell = bq(f"""
    SELECT customer,
           recommended_brand AS brand,
           ROUND(confidence_score, 1)  AS confidence,
           similar_customers_buying    AS similar_customers,
           priority
    FROM {T('recommendations_crosssell')}
    WHERE priority = 'HIGH'
    ORDER BY confidence_score DESC
    LIMIT 50
""")

newstyle = bq(f"""
    SELECT customer, brand,
           recommended_style           AS style,
           ROUND(recommendation_score, 1) AS score,
           ROUND(colour_compatibility * 100, 0) AS colour_match_pct,
           priority
    FROM {T('recommendations_newstyle')}
    WHERE priority = 'HIGH'
      AND recommendation_score >= 8
    ORDER BY recommendation_score DESC
    LIMIT 100
""")

rec_counts = bq(f"""
    SELECT
        COUNTIF(rec_type='REORDER' AND priority='URGENT')  AS reorder_urgent,
        COUNTIF(rec_type='REORDER')                        AS reorder_total,
        COUNTIF(rec_type='CROSS_SELL' AND priority='HIGH') AS xsell_high,
        COUNTIF(rec_type='NEW_STYLE'  AND priority='HIGH') AS newstyle_high
    FROM {T('recommendations_summary')}
""")

with open("public/data/recommendations.json", "w") as f:
    json.dump({
        "counts":    rec_counts.to_dict(orient="records")[0],
        "reorder":   reorder.fillna("").to_dict(orient="records"),
        "crosssell": crosssell.to_dict(orient="records"),
        "newstyle":  newstyle.to_dict(orient="records")
    }, f)

# ── 4. ACCURACY ─────────────────────────────────────────────────────
print("  [4/5] Accuracy data...")
validation = bq(f"""
    SELECT brand,
           ROUND(ens_monthly, 0)         AS forecast,
           ROUND(actual_monthly_avg, 0)  AS actual,
           ROUND(ratio, 2)               AS ratio,
           ROUND(ABS(ens_monthly - actual_monthly_avg)
                 / NULLIF(actual_monthly_avg,0) * 100, 1) AS pct_error,
           status
    FROM {T('ensemble_brand_validation')}
    WHERE actual_monthly_avg > 0
    ORDER BY actual_monthly_avg DESC
""")
try:
    walkforward = bq(f"""
        SELECT window_label, naive_mae, stat_mae
        FROM {T('mae_walkforward')}
        ORDER BY window_label
    """)
    wf_data = walkforward.to_dict(orient="records")
except Exception:
    wf_data = []
try:
    features = bq(f"""
        SELECT feature, importance,
               ROUND(importance/SUM(importance) OVER()*100,1) AS pct
        FROM {T('lgbm_feature_importance')}
        ORDER BY importance DESC
        LIMIT 15
    """)
    feat_data = features.to_dict(orient="records")
except Exception:
    feat_data = []

with open("public/data/accuracy.json", "w") as f:
    json.dump({
        "validation":  validation.to_dict(orient="records"),
        "walkforward": wf_data,
        "features":    feat_data
    }, f)

# ── 5. META ─────────────────────────────────────────────────────────
print("  [5/5] Meta...")
with open("public/data/meta.json", "w") as f:
    json.dump({
        "last_updated": datetime.now().strftime("%d %B %Y, %H:%M"),
        "pipeline":     "May 2026",
        "project":      PROJECT,
        "dataset":      DATASET
    }, f)

print()
print("Done! All 5 files written to public/data/")
print("  forecast.json · buyplan.json · recommendations.json · accuracy.json · meta.json")
print()
print("Next: git add . && git commit -m 'update' && git push")
