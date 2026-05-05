from pathlib import Path

path = Path.home() / "nordic_site" / "export_data.py"
text = path.read_text()

text = text.replace(
"""    HAVING SUM(buy_qty) > 0
    ORDER BY buy_per_month DESC""",
"""    HAVING SUM(COALESCE(forecast_p50,0)) > 0 OR SUM(COALESCE(buy_qty,0)) > 0
    ORDER BY buy_per_month DESC"""
)

text = text.replace(
"""ROW_NUMBER() OVER (PARTITION BY brand ORDER BY buy_qty DESC) AS rn
        FROM {T('buy_plan')}
        WHERE buy_qty > 0""",
"""ROW_NUMBER() OVER (
                   PARTITION BY brand
                   ORDER BY COALESCE(buy_qty,0) DESC, COALESCE(forecast_p50,0) DESC
               ) AS rn
        FROM {T('buy_plan')}
        WHERE COALESCE(forecast_p50,0) > 0
           OR COALESCE(forecast_p90,0) > 0
           OR COALESCE(buy_qty,0) > 0"""
)

text = text.replace(
"""ROUND(buy_qty,0)      AS buy_qty,""",
"""ROUND(COALESCE(buy_qty,0),0)      AS buy_qty,"""
)

text = text.replace(
"""ROUND(forecast_p50,0) AS p50,
               ROUND(forecast_p90,0) AS p90,""",
"""ROUND(COALESCE(forecast_p50,0),0) AS p50,
               ROUND(COALESCE(forecast_p90,0),0) AS p90,"""
)

path.write_text(text)
print("Strong Buy Plan export fix applied")
