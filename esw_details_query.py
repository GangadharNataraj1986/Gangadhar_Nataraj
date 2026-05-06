"""
Query ESW details from Databricks using an ESW number field value.
Supports single or multiple ESW numbers entered as comma/space/newline separated text.

Examples:
    python esw_details_query.py --esw "20060599"
    python esw_details_query.py --esw "20060599,20060600 20060601"
"""

from __future__ import annotations

import argparse
import re
from typing import Any, Dict, List

import pyodbc

DSN = "Spark-PRD"

# ESW statuses that disqualify a record from summarisation.
_INACTIVE_STATUSES = {"Closed", "Rejected", "Expired", "Cancelled", "Canceled"}

SQL_TEMPLATE = """
SELECT
    ec_number AS esw_number,
    title,
    problem,
    solution,
    status
FROM prd.rd_core.view_projectx_ec_form_data
WHERE
    ec_number IN ({placeholders})
    AND status NOT IN ('Closed', 'Rejected', 'Expired', 'Cancelled', 'Canceled')
ORDER BY ec_number
"""


def parse_esw_numbers(esw_field_value: str) -> List[int]:
    """Extract numeric ESW numbers from free-form input and keep first-seen order."""
    if not esw_field_value or not esw_field_value.strip():
        raise ValueError("ESW field is empty. Enter one or more ESW numbers.")

    raw_ids = re.findall(r"\d+", esw_field_value)
    if not raw_ids:
        raise ValueError("No numeric ESW numbers found in input.")

    # Deduplicate while preserving order.
    unique_ids = list(dict.fromkeys(int(x) for x in raw_ids))
    return unique_ids


def query_esw_details(esw_field_value: str) -> List[Dict[str, Any]]:
    """Query Databricks and return a list of ESW detail dicts."""
    esw_numbers = parse_esw_numbers(esw_field_value)
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in esw_numbers)
    sql = SQL_TEMPLATE.format(placeholders=placeholders)

    print(f"Connecting to Databricks via ODBC DSN '{DSN}'...")
    conn = pyodbc.connect(f"DSN={DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        try:
            print(f"Running query for ESW number(s): {', '.join(map(str, esw_numbers))}")
            cursor.execute(sql)
            rows = cursor.fetchall()
        finally:
            cursor.close()
    finally:
        conn.close()

    results: List[Dict[str, Any]] = []
    found_ids: set = set()

    for row in rows:
        esw_number = int(row[0]) if row[0] is not None else None
        title = str(row[1]).strip() if row[1] is not None else ""
        problem = str(row[2]).strip() if row[2] is not None else ""
        solution = str(row[3]).strip() if row[3] is not None else ""
        status = str(row[4]).strip() if row[4] is not None else ""

        if esw_number is not None:
            found_ids.add(esw_number)

        results.append({
            "esw_number": esw_number,
            "title": title,
            "problem": problem,
            "solution": solution,
            "status": status,
        })

    not_found = [eid for eid in esw_numbers if eid not in found_ids]
    if not_found:
        print(f"WARNING: No records found for ESW number(s): {', '.join(map(str, not_found))}")

    return results


def fetch_esw_records(esw_field_value: str) -> Dict[str, Any]:
    """Fetch ESW records and split into valid / skipped / not_found.

    Returns a dict with three keys:
        "valid"     - list of dicts: esw_number, title, problem, solution, status
        "skipped"   - list of dicts: esw_number, status
                      (filtered out by inactive status)
        "not_found" - list of int (IDs present in input but absent in Databricks)
    """
    esw_numbers = parse_esw_numbers(esw_field_value)
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in esw_numbers)

    sql = f"""
    SELECT
        ec_number AS esw_number,
        title,
        problem,
        solution,
        status
    FROM prd.rd_core.view_projectx_ec_form_data
    WHERE ec_number IN ({placeholders})
    ORDER BY ec_number
    """

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
        esw_number = int(row[0]) if row[0] is not None else None
        title = str(row[1]).strip() if row[1] is not None else ""
        problem = str(row[2]).strip() if row[2] is not None else ""
        solution = str(row[3]).strip() if row[3] is not None else ""
        status = str(row[4]).strip() if row[4] is not None else ""

        if esw_number is None:
            continue

        found_ids.add(esw_number)

        if status in _INACTIVE_STATUSES:
            skipped.append({"esw_number": esw_number, "status": status})
        else:
            valid.append({
                "esw_number": esw_number,
                "title": title,
                "problem": problem,
                "solution": solution,
                "status": status,
            })

    not_found = [eid for eid in esw_numbers if eid not in found_ids]
    return {"valid": valid, "skipped": skipped, "not_found": not_found}


def print_results(results: List[Dict[str, Any]]) -> None:
    if not results:
        print("No matching ESW records found.")
        return

    print(f"\nTotal ESW records retrieved: {len(results)}")
    print("-" * 170)
    print(f"{'ESW Number':<12} | {'Status':<20} | {'Title':<35} | {'Problem':<45} | {'Solution':<45}")
    print("-" * 170)

    for rec in results:
        esw = str(rec["esw_number"]) if rec["esw_number"] is not None else ""
        status = rec["status"]
        title = rec["title"]
        problem = rec["problem"]
        solution = rec["solution"]

        title_s = (title[:32] + "...") if len(title) > 35 else title
        problem_s = (problem[:42] + "...") if len(problem) > 45 else problem
        solution_s = (solution[:42] + "...") if len(solution) > 45 else solution

        print(f"{esw:<12} | {status:<20} | {title_s:<35} | {problem_s:<45} | {solution_s:<45}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Query ESW details from Databricks.")
    parser.add_argument(
        "--esw",
        dest="esw_field_value",
        help="ESW number(s) - comma/space/newline separated.",
    )
    args = parser.parse_args()

    esw_field_value = args.esw_field_value
    if not esw_field_value:
        esw_field_value = input("Enter ESW number(s): ").strip()

    try:
        results = query_esw_details(esw_field_value)
        print_results(results)
    except ValueError as exc:
        print(f"Input Error: {exc}")
    except pyodbc.Error as exc:
        print(f"ODBC Error: {exc}")
    except Exception as exc:
        print(f"Unexpected Error: {exc}")


if __name__ == "__main__":
    main()
