"""
Databricks ODBC Connection Script
Connects to Databricks using pyodbc and retrieves accessible schemas
Uses pre-configured DSN
"""

import pyodbc

# Configuration - use pre-configured DSN
DSN = "Spark-PRD"  # Pre-configured ODBC DSN

def connect_odbc_and_fetch_schemas():
    """Connect to Databricks via pre-configured ODBC DSN and fetch schema list"""
    
    # Use pre-configured DSN (should have all settings stored)
    try:
        # Connect to Databricks using DSN without autocommit
        print(f"Connecting to Databricks via ODBC DSN '{DSN}'...")
        conn = pyodbc.connect(f"DSN={DSN}", autocommit=True)
        cursor = conn.cursor()
        
        # Execute SHOW SCHEMAS query
        print("Fetching schemas...\n")
        cursor.execute("SHOW SCHEMAS")
        rows = cursor.fetchall()
        
        # Display results
        print(f"Connection established successfully!")
        print(f"\nTotal Schemas: {len(rows)}\n")
        print("Schemas accessible:")
        print("-" * 50)
        
        for i, row in enumerate(rows, 1):
            schema_name = row[0]
            print(f"{i:3d}. {schema_name}")
        
        cursor.close()
        conn.close()
        
    except pyodbc.Error as e:
        print(f"ODBC Connection Error: {e}")
        print(f"\nNote: DSN '{DSN}' may need proper credentials configured in Windows ODBC Data Sources")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    connect_odbc_and_fetch_schemas()
