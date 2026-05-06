"""
Databricks SQL Query Test via ODBC
Executes a sample SQL query against part_dim table and displays results
"""

import pyodbc

# Configuration
DSN = "Spark-PRD"

def execute_sql_query():
    """Connect to Databricks via ODBC and execute SQL query"""
    
    # SQL Query
    sql_query = """
    SELECT
      material        AS part_number,
      material_desc   AS part_description,
      part_category,
      material_class,
      network,
      cost,
      part_status_desc
    FROM prd.ud_apv.part_dim
    WHERE material IS NOT NULL
    ORDER BY material
    LIMIT 20
    """
    
    try:
        # Connect to Databricks using DSN with autocommit
        print(f"Connecting to Databricks via ODBC DSN '{DSN}'...")
        conn = pyodbc.connect(f"DSN={DSN}", autocommit=True)
        cursor = conn.cursor()
        
        # Execute SQL query
        print("\nExecuting SQL query...")
        print("-" * 80)
        cursor.execute(sql_query)
        rows = cursor.fetchall()
        
        # Display results
        print(f"\nQuery executed successfully!")
        print(f"Total records retrieved: {len(rows)}\n")
        
        if rows:
            # Print column headers
            columns = [desc[0] for desc in cursor.description]
            print("Results:")
            print("-" * 80)
            print(f"{columns[0]:<15} | {columns[1]:<30} | {columns[2]:<20} | {columns[3]:<20} | {columns[4]:<15} | {columns[5]:<15} | {columns[6]:<20}")
            print("-" * 80)
            
            # Print data rows
            for i, row in enumerate(rows, 1):
                print(f"{str(row[0]):<15} | {str(row[1]):<30} | {str(row[2]):<20} | {str(row[3]):<20} | {str(row[4]):<15} | {str(row[5]):<15} | {str(row[6]):<20}")
        else:
            print("No records found")
        
        cursor.close()
        conn.close()
        print("\nConnection closed successfully")
        
    except pyodbc.Error as e:
        print(f"ODBC Error: {e}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    execute_sql_query()
