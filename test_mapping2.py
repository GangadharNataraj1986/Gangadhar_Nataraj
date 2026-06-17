import sys
sys.path.insert(0, r'c:\Users\e187725\OneDrive - Applied Materials\Desktop\Orphan Anlsysis Pyt')

# Force reload of the module
if 'Problem_Solution_Agent_PSS.inventory_demand_cost_query' in sys.modules:
    del sys.modules['Problem_Solution_Agent_PSS.inventory_demand_cost_query']

from Problem_Solution_Agent_PSS.inventory_demand_cost_query import fetch_inventory_demand_mapping

rows = fetch_inventory_demand_mapping(['0055-03777'])
print(f'Rows returned: {len(rows)}')
if rows:
    print(f'First row system_number: {rows[0].get("system_number")}')
    print(f'First row consumption_qty: {rows[0].get("consumption_qty")}')
else:
    print('No rows returned')
