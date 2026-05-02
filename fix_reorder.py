import re

with open('/home/janhavi_m/nordic_site/export_data.py', 'r') as f:
    content = f.read()

old = '''    reorder = bq(f"""
    SELECT customer, brand, last_order_month, overdue_months,
           avg_order_qty, priority
    FROM {T('recommendations_reorder')}
    WHERE priority IN (\'URGENT\',\'DUE\')
    ORDER BY overdue_months DESC
    LIMIT 50
""")'''

new = '''    reorder = bq(f"""
    SELECT customer, brand, last_order_month,
           months_overdue, avg_order_qty, priority
    FROM {T('recommendations_reorder')}
    WHERE priority IN (\'URGENT\',\'DUE\')
    ORDER BY months_overdue DESC
    LIMIT 50
""")'''

content = content.replace(old, new)
with open('/home/janhavi_m/nordic_site/export_data.py', 'w') as f:
    f.write(content)
print("Fixed")
