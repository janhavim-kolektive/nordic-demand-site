from pathlib import Path

export_path = Path("/home/$USER/nordic_site/export_data.py".replace("$USER", Path.home().name))
html_path = Path("/home/$USER/nordic_site/public/index.html".replace("$USER", Path.home().name))

export = export_path.read_text()
html = html_path.read_text()

# 1) BUY PLAN: keep brands even if buy_qty is 0, as long as forecast exists
export = export.replace(
"""    GROUP BY brand
    HAVING SUM(buy_qty) > 0
    ORDER BY buy_per_month DESC""",
"""    GROUP BY brand
    HAVING SUM(COALESCE(forecast_p50,0)) > 0 OR SUM(COALESCE(buy_qty,0)) > 0
    ORDER BY buy_per_month DESC"""
)

export = export.replace(
"""               ROW_NUMBER() OVER (PARTITION BY brand ORDER BY buy_qty DESC) AS rn
        FROM {T('buy_plan')}
        WHERE buy_qty > 0""",
"""               ROW_NUMBER() OVER (
                   PARTITION BY brand 
                   ORDER BY COALESCE(buy_qty,0) DESC, COALESCE(forecast_p50,0) DESC
               ) AS rn
        FROM {T('buy_plan')}
        WHERE COALESCE(forecast_p50,0) > 0 
           OR COALESCE(forecast_p90,0) > 0 
           OR COALESCE(buy_qty,0) > 0"""
)

# 2) WALKFORWARD: if mae_walkforward table is missing/empty, create fallback validation chart data
export = export.replace(
"""try:
    wf = bq(f"SELECT window_label, naive_mae, stat_mae FROM {T('mae_walkforward')} ORDER BY window_label")
    wf_data = wf.to_dict(orient="records")
except: wf_data = []""",
"""try:
    wf = bq(f"SELECT window_label, naive_mae, stat_mae FROM {T('mae_walkforward')} ORDER BY window_label")
    wf_data = wf.to_dict(orient="records")
except Exception as e:
    print(f"    Walk-forward table unavailable: {e}")
    wf_data = []

# Fallback so the Accuracy page still shows a validation chart
if not wf_data:
    wf_fallback = bq(f'''
        SELECT
          'Brand validation' AS window_label,
          ROUND(AVG(actual_monthly_avg), 1) AS naive_mae,
          ROUND(AVG(ABS(ens_monthly - actual_monthly_avg)), 1) AS stat_mae
        FROM {T('ensemble_brand_validation')}
        WHERE actual_monthly_avg > 0
    ''')
    wf_data = wf_fallback.to_dict(orient="records")
    print("    Using fallback validation chart data")"""
)

export_path.write_text(export)

# 3) UI: default Buy Plan should not hide zero-buy forecast lines
html = html.replace(
"""<option value="ALL">Actionable lines (buy qty &gt; 0)</option>""",
"""<option value="ALL">All forecast lines</option>"""
)

html = html.replace(
"""  if(flag==='ALL')rows=rows.filter(r=>(+r.buy_qty||0)>0);
  else if(flag==='below_moq')rows=rows.filter(r=>+r.below_moq);""",
"""  if(flag==='ALL'){
    // Show all forecast lines, including brands with forecast but zero buy quantity
  }
  else if(flag==='below_moq')rows=rows.filter(r=>+r.below_moq);"""
)

# 4) UI: show a message instead of blank area if walkforward has no data
html = html.replace(
"""function renderWalkforward(){
  const wf=D.accuracy?.walkforward||[];
  if(!wf.length)return;""",
"""function renderWalkforward(){
  const wf=D.accuracy?.walkforward||[];
  if(!wf.length){
    const box=document.getElementById('wf-chart')?.parentElement;
    if(box) box.innerHTML='<div class="loading">Walk-forward validation data not available. Re-run export_data.py or check mae_walkforward table.</div>';
    return;
  }"""
)

html_path.write_text(html)

print("Fixed export_data.py and public/index.html")
