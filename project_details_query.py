"""
Query Project details from Databricks using a Project ID.
Supports single or multiple Project IDs entered as comma/space/newline separated text.

Examples:
    python project_details_query.py --project "12550520"
    python project_details_query.py --project "12550520,12550521 12550522"
"""

from __future__ import annotations

import argparse
import re
from typing import Any, Dict, List

import pyodbc

DSN = "Spark-PRD"

# Project statuses that disqualify a record from summarisation.
_INACTIVE_STATUSES = {"Closed", "Cancelled", "Canceled", "On Hold"}

SQL_TEMPLATE = """
SELECT
  pf.project_id        AS core_project_number,
  pf.project_name,
  pf.project_scope_desc AS defined_scope,
  pf.deliverable        AS deliverables,
  st.ec_status          AS status
FROM
  prd.rd_core.tbl_projectx_projectform pf
  LEFT JOIN prd.rd_core.tbl_projectx_ec_status st
    ON pf.project_status_id = st.ec_status_id
WHERE
  pf.project_id IN ({placeholders})
ORDER BY
  pf.project_id
"""


def parse_project_ids(project_field_value: str) -> List[int]:
    """Extract numeric Project IDs from free-form input and keep first-seen order."""
    if not project_field_value or not project_field_value.strip():
        raise ValueError("Project field is empty. Enter one or more Project IDs.")

    raw_ids = re.findall(r"\d+", project_field_value)
    if not raw_ids:
        raise ValueError("No numeric Project IDs found in input.")

    # Deduplicate while preserving order.
    unique_ids = list(dict.fromkeys(int(x) for x in raw_ids))
    return unique_ids


def query_project_details(project_field_value: str) -> List[Dict[str, Any]]:
    """Query Databricks and return a list of project detail dicts."""
    project_ids = parse_project_ids(project_field_value)
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in project_ids)
    sql = SQL_TEMPLATE.format(placeholders=placeholders)

    print(f"Connecting to Databricks via ODBC DSN '{DSN}'...")
    conn = pyodbc.connect(f"DSN={DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        try:
            print(f"Running query for Project IDs: {', '.join(map(str, project_ids))}")
            cursor.execute(sql)
            rows = cursor.fetchall()
        finally:
            cursor.close()
    finally:
        conn.close()

    results: List[Dict[str, Any]] = []
    found_ids: set = set()

    for row in rows:
        project_id  = int(row[0]) if row[0] is not None else None
        name        = str(row[1]).strip() if row[1] is not None else ""
        scope       = str(row[2]).strip() if row[2] is not None else ""
        deliverable = str(row[3]).strip() if row[3] is not None else ""
        status      = str(row[4]).strip() if row[4] is not None else ""

        if project_id is not None:
            found_ids.add(project_id)

        results.append({
            "core_project_number": project_id,
            "project_name":        name,
            "defined_scope":       scope,
            "deliverables":        deliverable,
            "status":              status,
        })

    not_found = [pid for pid in project_ids if pid not in found_ids]
    if not_found:
        print(f"WARNING: No records found for Project ID(s): {', '.join(map(str, not_found))}")

    return results


def fetch_project_records(project_field_value: str) -> Dict[str, Any]:
    """Fetch Project records and split into valid / skipped / not_found.

    Returns a dict with three keys:
        "valid"     - list of dicts: project_id, project_name, defined_scope,
                      deliverables, status
        "skipped"   - list of dicts: project_id, status
                      (filtered out by inactive status)
        "not_found" - list of int (IDs present in input but absent in Databricks)
    """
    project_ids = parse_project_ids(project_field_value)
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in project_ids)
    sql = SQL_TEMPLATE.format(placeholders=placeholders)

    conn = pyodbc.connect(f"DSN={DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        try:
            cursor.execute(sql)
            rows = cursor.fetchall()
        finally:
            cursor.close()
    finally:
        conn.close()

    found_ids: set = set()
    valid: List[Dict[str, Any]] = []
    skipped: List[Dict[str, Any]] = []

    for row in rows:
        project_id = int(row[0]) if row[0] is not None else None
        name = str(row[1]).strip() if row[1] is not None else ""
        scope = str(row[2]).strip() if row[2] is not None else ""
        deliverable = str(row[3]).strip() if row[3] is not None else ""
        status = str(row[4]).strip() if row[4] is not None else ""

        if project_id is None:
            continue

        found_ids.add(project_id)

        if status in _INACTIVE_STATUSES:
            skipped.append({"project_id": project_id, "status": status})
        else:
            valid.append({
                "project_id": project_id,
                "project_name": name,
                "defined_scope": scope,
                "deliverables": deliverable,
                "status": status,
            })

    not_found = [pid for pid in project_ids if pid not in found_ids]
    return {"valid": valid, "skipped": skipped, "not_found": not_found}


def print_results(results: List[Dict[str, Any]]) -> None:
    if not results:
        print("No matching project records found.")
        return

    print(f"\nTotal project records retrieved: {len(results)}")
    print("-" * 140)
    print(f"{'Project ID':<15} | {'Project Name':<35} | {'Status':<20} | {'Scope':<30} | {'Deliverables':<30}")
    print("-" * 140)

    for rec in results:
        pid    = str(rec["core_project_number"]) if rec["core_project_number"] is not None else ""
        name   = rec["project_name"]
        status = rec["status"]
        scope  = rec["defined_scope"]
        deliv  = rec["deliverables"]

        name_s   = (name[:32]   + "...") if len(name)   > 35 else name
        scope_s  = (scope[:27]  + "...") if len(scope)  > 30 else scope
        deliv_s  = (deliv[:27]  + "...") if len(deliv)  > 30 else deliv

        print(f"{pid:<15} | {name_s:<35} | {status:<20} | {scope_s:<30} | {deliv_s:<30}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Query Project details from Databricks.")
    parser.add_argument(
        "--project",
        dest="project_field_value",
        help="Project ID(s) — comma/space/newline separated.",
    )
    args = parser.parse_args()

    project_field_value = args.project_field_value
    if not project_field_value:
        project_field_value = input("Enter Project ID(s): ").strip()

    try:
        results = query_project_details(project_field_value)
        print_results(results)
    except ValueError as exc:
        print(f"Input Error: {exc}")
    except pyodbc.Error as exc:
        print(f"ODBC Error: {exc}")
    except Exception as exc:
        print(f"Unexpected Error: {exc}")


if __name__ == "__main__":
    main()
