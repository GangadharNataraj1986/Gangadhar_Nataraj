import sys
sys.path.append(r'c:\Users\e187725\OneDrive - Applied Materials\Desktop\Orphan Anlsysis Pyt\Problem_Solution_Agent_PSS')

from inventory_demand_cost_query import fetch_open_purchase_order_details

print("Testing PO query with part 0195-06226...")
rows = fetch_open_purchase_order_details(['0195-06226'])
print(f'ROW_COUNT: {len(rows)}')

if rows:
    print('Sample row keys:', list(rows[0].keys()))
    print('\nSample row (first):')
    for key, val in rows[0].items():
        print(f'  {key}: {val}')
else:
    print("No rows returned!")
