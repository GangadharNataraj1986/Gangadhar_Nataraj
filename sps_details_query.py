"""
Query SPS details from Databricks using an SPS ID field value.
Supports single or multiple SPS IDs entered as comma/space/newline separated text.

Examples:
    python sps_details_query.py --sps "225924"
    python sps_details_query.py --sps "225924,225925 225926"
"""

from __future__ import annotations

import argparse
import re
from typing import Any, Dict, List

import pyodbc

DSN = "Spark-PRD"

# SPS statuses that disqualify a record from summarisation.
_INACTIVE_STATUSES = {"Closed", "Rejected", "Cancelled", "Canceled"}

SQL_TEMPLATE = """
SELECT
    p.sps_id,
    p.part_number,
    p.part_description,
    p.status,
    pr.problem_description AS problem,
    pr.problem_soultion AS proposed_solution,
    s.solution
FROM
    prd.rd_sharepoint.sps_parts p
LEFT JOIN prd.rd_sharepoint.sps_problem pr
    ON p.sps_id = pr.sps_id
LEFT JOIN prd.rd_sharepoint.sps_solution s
    ON p.sps_id = s.sps_id
WHERE
    p.sps_id IN ({placeholders})
        AND p.status NOT IN ('Closed', 'Rejected', 'Cancelled', 'Canceled')
ORDER BY
    p.sps_id
"""


def parse_sps_ids(sps_field_value: str) -> List[int]:
    """Extract numeric SPS IDs from free-form input and keep first-seen order."""
    if not sps_field_value or not sps_field_value.strip():
        raise ValueError("SPS field is empty. Enter one or more SPS IDs.")

    raw_ids = re.findall(r"\d+", sps_field_value)
    if not raw_ids:
        raise ValueError("No numeric SPS IDs found in input.")

    # Deduplicate while preserving order.
    unique_ids = list(dict.fromkeys(int(x) for x in raw_ids))
    return unique_ids


def query_sps_details(sps_field_value: str) -> List[Dict[str, Any]]:
    """Query Databricks and return a list of SPS detail dicts."""
    sps_ids = parse_sps_ids(sps_field_value)
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in sps_ids)
    sql = SQL_TEMPLATE.format(placeholders=placeholders)

    print(f"Connecting to Databricks via ODBC DSN '{DSN}'...")
    conn = pyodbc.connect(f"DSN={DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        try:
            print(f"Running query for SPS IDs: {', '.join(map(str, sps_ids))}")
            cursor.execute(sql)
            rows = cursor.fetchall()
        finally:
            cursor.close()
    finally:
        conn.close()

    results: List[Dict[str, Any]] = []
    found_ids: set = set()

    for row in rows:
        sps_id       = int(row[0]) if row[0] is not None else None
        part_number  = str(row[1]).strip() if row[1] is not None else ""
        part_desc    = str(row[2]).strip() if row[2] is not None else ""
        status       = str(row[3]).strip() if row[3] is not None else ""
        problem      = str(row[4]).strip() if row[4] is not None else ""
        proposed_sol = str(row[5]).strip() if row[5] is not None else ""
        solution     = str(row[6]).strip() if row[6] is not None else ""

        if sps_id is not None:
            found_ids.add(sps_id)

        results.append({
            "sps_id":           sps_id,
            "part_number":      part_number,
            "part_description": part_desc,
            "status":           status,
            "problem":          problem,
            "proposed_solution": proposed_sol,
            "solution":         solution,
        })

    not_found = [sid for sid in sps_ids if sid not in found_ids]
    if not_found:
        print(f"WARNING: No records found for SPS ID(s): {', '.join(map(str, not_found))}")

    return results


def fetch_sps_records(sps_field_value: str) -> Dict[str, Any]:
    """Fetch SPS records and split into valid / skipped / not_found.

    Returns a dict with three keys:
        "valid"     - list of dicts: sps_id, part_number, part_description,
                      status, problem, proposed_solution, solution
        "skipped"   - list of dicts: sps_id, status
                      (filtered out by inactive status)
        "not_found" - list of int (IDs present in input but absent in Databricks)
    """
    sps_ids = parse_sps_ids(sps_field_value)
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in sps_ids)
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
        sps_id       = int(row[0]) if row[0] is not None else None
        part_number  = str(row[1]).strip() if row[1] is not None else ""
        part_desc    = str(row[2]).strip() if row[2] is not None else ""
        status       = str(row[3]).strip() if row[3] is not None else ""
        problem      = str(row[4]).strip() if row[4] is not None else ""
        proposed_sol = str(row[5]).strip() if row[5] is not None else ""
        solution     = str(row[6]).strip() if row[6] is not None else ""

        if sps_id is None:
            continue

        found_ids.add(sps_id)

        if status in _INACTIVE_STATUSES:
            skipped.append({"sps_id": sps_id, "status": status})
        else:
            valid.append({
                "sps_id":            sps_id,
                "part_number":       part_number,
                "part_description":  part_desc,
                "status":            status,
                "problem":           problem,
                "proposed_solution": proposed_sol,
                "solution":          solution,
            })

    not_found = [sid for sid in sps_ids if sid not in found_ids]
    return {"valid": valid, "skipped": skipped, "not_found": not_found}


def print_results(results: List[Dict[str, Any]]) -> None:
    if not results:
        print("No matching SPS records found.")
        return

    print(f"\nTotal SPS records retrieved: {len(results)}")
    print("-" * 160)
    print(f"{'SPS ID':<12} | {'Part Number':<20} | {'Status':<20} | {'Problem':<30} | {'Proposed Solution':<30} | {'Solution':<30}")
    print("-" * 160)

    for rec in results:
        sid     = str(rec["sps_id"]) if rec["sps_id"] is not None else ""
        part    = rec["part_number"]
        status  = rec["status"]
        problem = rec["problem"]
        prop_s  = rec["proposed_solution"]
        sol     = rec["solution"]

        part_s    = (part[:17]    + "...") if len(part)    > 20 else part
        problem_s = (problem[:27] + "...") if len(problem) > 30 else problem
        prop_ss   = (prop_s[:27]  + "...") if len(prop_s)  > 30 else prop_s
        sol_s     = (sol[:27]     + "...") if len(sol)     > 30 else sol

        print(f"{sid:<12} | {part_s:<20} | {status:<20} | {problem_s:<30} | {prop_ss:<30} | {sol_s:<30}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Query SPS details from Databricks.")
    parser.add_argument(
        "--sps",
        dest="sps_field_value",
        help="SPS ID(s) — comma/space/newline separated.",
    )
    args = parser.parse_args()

    sps_field_value = args.sps_field_value
    if not sps_field_value:
        sps_field_value = input("Enter SPS ID(s): ").strip()

    try:
        results = query_sps_details(sps_field_value)
        print_results(results)
    except ValueError as exc:
        print(f"Input Error: {exc}")
    except pyodbc.Error as exc:
        print(f"ODBC Error: {exc}")
    except Exception as exc:
        print(f"Unexpected Error: {exc}")


if __name__ == "__main__":
    main()
