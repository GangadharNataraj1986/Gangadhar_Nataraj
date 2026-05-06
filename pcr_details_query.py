"""
Query PCR details from Databricks using a PCR field value.
Supports single or multiple PCR IDs entered as comma/space/newline separated text.

Examples:
    python pcr_details_query.py --pcr "10032740"
    python pcr_details_query.py --pcr "10032740,10032741 10032742"
"""

from __future__ import annotations

import argparse
import re
from typing import Any, Dict, List

import pyodbc

DSN = "Spark-PRD"

# PCR statuses that disqualify a record from summarisation.
_INACTIVE_STATUSES = {"Closed", "Cancelled", "On Hold"}


def parse_pcr_ids(pcr_field_value: str) -> List[int]:
    """Extract numeric PCR IDs from free-form input and keep first-seen order."""
    if not pcr_field_value or not pcr_field_value.strip():
        raise ValueError("PCR field is empty. Enter one or more PCR IDs.")

    raw_ids = re.findall(r"\d+", pcr_field_value)
    if not raw_ids:
        raise ValueError("No numeric PCR IDs found in input.")

    # Deduplicate while preserving order.
    unique_ids = list(dict.fromkeys(int(x) for x in raw_ids))
    return unique_ids


def build_query(pcr_ids: List[int]) -> str:
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in pcr_ids)
    return f"""
    SELECT
        pcr.pcr_id,
        ec.ec_problem AS problem_statement,
        ec.ec_solution AS solution_statement,
        st.ec_status AS pcr_status
    FROM prd.rd_core.tbl_projectx_ec_pcr pcr
    JOIN prd.rd_core.tbl_projectx_ec ec
        ON pcr.pcr_id = ec.ec_number
    LEFT JOIN prd.rd_core.tbl_projectx_ec_status st
        ON ec.ec_status_id = st.ec_status_id
    WHERE pcr.pcr_id IN ({placeholders})
      AND (st.ec_status IS NULL OR st.ec_status NOT IN ('Closed', 'Canceled', 'On Hold'))
    ORDER BY pcr.pcr_id
    """


def query_pcr_details(pcr_field_value: str) -> List[pyodbc.Row]:
    pcr_ids = parse_pcr_ids(pcr_field_value)
    sql_query = build_query(pcr_ids)

    print(f"Connecting to Databricks via ODBC DSN '{DSN}'...")
    conn = pyodbc.connect(f"DSN={DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        try:
            print(f"Running query for PCR IDs: {', '.join(map(str, pcr_ids))}")
            cursor.execute(sql_query)
            rows = cursor.fetchall()
            return rows
        finally:
            cursor.close()
    finally:
        conn.close()


def fetch_pcr_records(pcr_field_value: str) -> Dict[str, Any]:
    """Fetch PCR records from Databricks and split into valid / skipped / not_found.

    Returns a dict with three keys:
        "valid"     – list of dicts: pcr_id, problem, solution, status, psnnumber
        "skipped"   – list of dicts: pcr_id, status  (filtered out by inactive status)
        "not_found" – list of int   (IDs present in input but absent in Databricks)
    """
    pcr_ids = parse_pcr_ids(pcr_field_value)
    # Databricks Hive SQL does not support positional ? parameters; inline validated ints.
    placeholders = ", ".join(str(n) for n in pcr_ids)

    sql = f"""
    SELECT
        pcr.pcr_id,
        ec.ec_problem          AS problem_statement,
        ec.ec_solution         AS solution_statement,
        st.ec_status           AS pcr_status,
        ec.psnnumber
    FROM prd.rd_core.tbl_projectx_ec_pcr pcr
    JOIN prd.rd_core.tbl_projectx_ec ec
        ON pcr.pcr_id = ec.ec_number
    LEFT JOIN prd.rd_core.tbl_projectx_ec_status st
        ON ec.ec_status_id = st.ec_status_id
    WHERE pcr.pcr_id IN ({placeholders})
    ORDER BY pcr.pcr_id
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
        pcr_id  = int(row[0])
        problem = str(row[1]).strip() if row[1] is not None else ""
        solution = str(row[2]).strip() if row[2] is not None else ""
        status  = str(row[3]).strip() if row[3] is not None else ""
        psn     = str(row[4]).strip() if row[4] is not None else ""
        found_ids.add(pcr_id)

        if status in _INACTIVE_STATUSES:
            skipped.append({"pcr_id": pcr_id, "status": status})
        else:
            valid.append({
                "pcr_id":    pcr_id,
                "problem":   problem,
                "solution":  solution,
                "status":    status,
                "psnnumber": psn if psn else None,
            })

    not_found = [pid for pid in pcr_ids if pid not in found_ids]

    return {"valid": valid, "skipped": skipped, "not_found": not_found}


def print_results(rows: List[pyodbc.Row]) -> None:
    if not rows:
        print("No matching PCR records found (or all are filtered by status).")
        return

    print(f"\nTotal PCR records retrieved: {len(rows)}")
    print("-" * 120)
    print(f"{'PCR ID':<12} | {'PCR Status':<25} | {'Problem':<35} | {'Solution':<35}")
    print("-" * 120)

    for row in rows:
        pcr_id = str(row[0])
        problem = str(row[1]) if row[1] is not None else ""
        solution = str(row[2]) if row[2] is not None else ""
        status = str(row[3]) if row[3] is not None else ""

        # Keep terminal output readable in one row per record.
        problem_short = (problem[:32] + "...") if len(problem) > 35 else problem
        solution_short = (solution[:32] + "...") if len(solution) > 35 else solution

        print(f"{pcr_id:<12} | {status:<25} | {problem_short:<35} | {solution_short:<35}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Query PCR Problem/Solution/Status from Databricks.")
    parser.add_argument(
        "--pcr",
        dest="pcr_field_value",
        help="PCR field input with one or multiple IDs (comma/space/newline separated).",
    )
    args = parser.parse_args()

    pcr_field_value = args.pcr_field_value
    if not pcr_field_value:
        pcr_field_value = input("Enter PCR ID(s): ").strip()

    try:
        rows = query_pcr_details(pcr_field_value)
        print_results(rows)
    except ValueError as exc:
        print(f"Input Error: {exc}")
    except pyodbc.Error as exc:
        print(f"ODBC Error: {exc}")
    except Exception as exc:
        print(f"Unexpected Error: {exc}")


if __name__ == "__main__":
    main()
