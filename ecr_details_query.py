"""
Query ECR details from Databricks using an ECR number field value.
Supports single or multiple ECR numbers entered as comma/space/newline separated text.

Examples:
    python ecr_details_query.py --ecr "3173415"
    python ecr_details_query.py --ecr "3173415,3173416 3173417"
"""

from __future__ import annotations

import argparse
import re
from typing import Any, Dict, List

import pyodbc

DSN = "Spark-PRD"

# ECR statuses that disqualify a record from summarisation.
_INACTIVE_STATUSES = {"Closed", "Rejected", "Expired", "Cancelled", "Canceled"}

SQL_TEMPLATE = """
SELECT
    ec_number,
    problem,
    solution,
    status
FROM prd.rd_core.view_projectx_ec_form_data
WHERE
    ec_number IN ({placeholders})
ORDER BY ec_number
"""


def parse_ecr_numbers(ecr_field_value: str) -> List[int]:
    """Extract numeric ECR numbers from free-form input and keep first-seen order."""
    if not ecr_field_value or not ecr_field_value.strip():
        raise ValueError("ECR field is empty. Enter one or more ECR numbers.")

    raw_ids = re.findall(r"\d+", ecr_field_value)
    if not raw_ids:
        raise ValueError("No numeric ECR numbers found in input.")

    # Deduplicate while preserving order.
    unique_ids = list(dict.fromkeys(int(x) for x in raw_ids))
    return unique_ids


def query_ecr_details(ecr_field_value: str) -> List[Dict[str, Any]]:
    """Query Databricks and return a list of ECR detail dicts."""
    ecr_numbers = parse_ecr_numbers(ecr_field_value)
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in ecr_numbers)
    sql = SQL_TEMPLATE.format(placeholders=placeholders)

    print(f"Connecting to Databricks via ODBC DSN '{DSN}'...")
    conn = pyodbc.connect(f"DSN={DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        try:
            print(f"Running query for ECR number(s): {', '.join(map(str, ecr_numbers))}")
            cursor.execute(sql)
            rows = cursor.fetchall()
        finally:
            cursor.close()
    finally:
        conn.close()

    results: List[Dict[str, Any]] = []
    found_ids: set = set()

    for row in rows:
        ecr_number = int(row[0]) if row[0] is not None else None
        problem    = str(row[1]).strip() if row[1] is not None else ""
        solution   = str(row[2]).strip() if row[2] is not None else ""
        status     = str(row[3]).strip() if row[3] is not None else ""

        if ecr_number is not None:
            found_ids.add(ecr_number)

        results.append({
            "ecr_number": ecr_number,
            "problem":    problem,
            "solution":   solution,
            "status":     status,
        })

    not_found = [eid for eid in ecr_numbers if eid not in found_ids]
    if not_found:
        print(f"WARNING: No records found for ECR number(s): {', '.join(map(str, not_found))}")

    return results


def fetch_ecr_records(ecr_field_value: str) -> Dict[str, Any]:
    """Fetch ECR records and split into valid / skipped / not_found.

    Returns a dict with three keys:
        "valid"     - list of dicts: ecr_number, problem, solution, status
        "skipped"   - list of dicts: ecr_number, status
                      (filtered out by inactive status)
        "not_found" - list of int (IDs present in input but absent in Databricks)
    """
    ecr_numbers = parse_ecr_numbers(ecr_field_value)
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in ecr_numbers)

    sql = f"""
    SELECT
        ec_number,
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
        ecr_number = int(row[0]) if row[0] is not None else None
        problem    = str(row[1]).strip() if row[1] is not None else ""
        solution   = str(row[2]).strip() if row[2] is not None else ""
        status     = str(row[3]).strip() if row[3] is not None else ""

        if ecr_number is None:
            continue

        found_ids.add(ecr_number)

        if status in _INACTIVE_STATUSES:
            skipped.append({"ecr_number": ecr_number, "status": status})
        else:
            valid.append({
                "ecr_number": ecr_number,
                "problem":    problem,
                "solution":   solution,
                "status":     status,
            })

    not_found = [eid for eid in ecr_numbers if eid not in found_ids]
    return {"valid": valid, "skipped": skipped, "not_found": not_found}


def print_results(results: List[Dict[str, Any]]) -> None:
    """Pretty-print ECR query results to stdout."""
    if not results:
        print("No results returned.")
        return
    for rec in results:
        print("-" * 60)
        print(f"ECR Number : {rec.get('ecr_number')}")
        print(f"Status     : {rec.get('status')}")
        print(f"Problem    : {rec.get('problem', '')[:300]}")
        print(f"Solution   : {rec.get('solution', '')[:300]}")
    print("-" * 60)


def main() -> None:
    parser = argparse.ArgumentParser(description="Query ECR details from Databricks.")
    parser.add_argument(
        "--ecr",
        required=True,
        help="ECR number(s) to query. Comma/space/newline separated.",
    )
    args = parser.parse_args()

    result = fetch_ecr_records(args.ecr)

    valid     = result.get("valid", [])
    skipped   = result.get("skipped", [])
    not_found = result.get("not_found", [])

    if valid:
        print(f"\n=== Valid ECR Records ({len(valid)}) ===")
        print_results(valid)

    if skipped:
        print(f"\n=== Skipped (inactive status) ({len(skipped)}) ===")
        for s in skipped:
            print(f"  ECR {s['ecr_number']} – {s['status']}")

    if not_found:
        print(f"\n=== Not Found ({len(not_found)}) ===")
        for n in not_found:
            print(f"  ECR {n}")


if __name__ == "__main__":
    main()
