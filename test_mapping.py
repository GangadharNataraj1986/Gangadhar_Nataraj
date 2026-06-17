import sys
sys.path.insert(0, r'c:\Users\e187725\OneDrive - Applied Materials\Desktop\Orphan Anlsysis Pyt')

from Problem_Solution_Agent_PSS.inventory_demand_cost_query import _clean_parts
import pyodbc

parts = ['0055-03777']
cleaned_parts = _clean_parts(parts)
print(f'Cleaned parts: {cleaned_parts}')

# Build part filter for SQL query
part_filter = ", ".join(f"'{p}'" for p in cleaned_parts)
print(f'Part filter: {part_filter}')

sql = f"""
SELECT COUNT(*) as cnt
FROM prd.pd_mm.factknxssplyconsmp
WHERE UPPER(REGEXP_REPLACE(CAST(dpart AS STRING), '[^A-Za-z0-9]', '')) IN ({part_filter})
"""

print(f'SQL: {sql}')

conn = pyodbc.connect('DSN=Spark-PRD', autocommit=True)
try:
    cursor = conn.cursor()
    cursor.execute(sql)
    rows = cursor.fetchall()
    print(f'Count result: {rows[0][0]}')
finally:
    conn.close()

