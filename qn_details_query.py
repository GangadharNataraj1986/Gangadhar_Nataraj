"""
Query QN (Quality Note) details from Databricks using a QN number field value.
Supports single or multiple QN numbers entered as comma/space/newline separated text.

Examples:
    python qn_details_query.py --qn "42630248"
    python qn_details_query.py --qn "42630248,42630249 42630250"
"""

from __future__ import annotations

import argparse
import re
from typing import Any, Dict, List

import pyodbc

DSN = "Spark-PRD"

# QN statuses that disqualify a record from summarisation.
_INACTIVE_STATUSES = {"Closed", "Cancelled", "Canceled"}

SQL_TEMPLATE = """
SELECT
    qr_id,
    sap_notification_number AS qn_number,
    title,
    problem_statement,
    immediate_fix AS solution,
    immediate_corrective_action,
    comments,
    state AS qn_status,
    date_created,
    date_closed,
    fault,
    fault_level_2,
    assembly,
    part_num,
    part_desc,
    owner,
    originator,
    nc_manufacturing_site,
    dmr_disposition,
    root_cause
FROM
    prd.pd_tw.manufacturing_nc_summary
WHERE
    sap_notification_number IN ({placeholders})
ORDER BY
    sap_notification_number
"""


def parse_qn_numbers(qn_field_value: str) -> List[str]:
    """Extract numeric QN numbers from free-form input and keep first-seen order."""
    if not qn_field_value or not qn_field_value.strip():
        raise ValueError("QN field is empty. Enter one or more QN numbers.")

    raw_ids = re.findall(r"\d+", qn_field_value)
    if not raw_ids:
        raise ValueError("No numeric QN numbers found in input.")

    # Deduplicate while preserving order (kept as strings to match DB type).
    unique_ids = list(dict.fromkeys(raw_ids))
    return unique_ids


def query_qn_details(qn_field_value: str) -> List[Dict[str, Any]]:
    """Query Databricks and return a list of QN detail dicts."""
    qn_numbers = parse_qn_numbers(qn_field_value)
    # Databricks Hive SQL does not support positional ? parameters in IN() clauses.
    # qn_numbers are validated as digit-only strings via re.findall so inlining is safe.
    placeholders = ", ".join(f"'{n}'" for n in qn_numbers)
    sql = SQL_TEMPLATE.format(placeholders=placeholders)

    print(f"Connecting to Databricks via ODBC DSN '{DSN}'...")
    conn = pyodbc.connect(f"DSN={DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        try:
            print(f"Running query for QN number(s): {', '.join(qn_numbers)}")
            cursor.execute(sql)
            rows = cursor.fetchall()
        finally:
            cursor.close()
    finally:
        conn.close()

    results: List[Dict[str, Any]] = []
    found_ids: set = set()

    for row in rows:
        qr_id                      = int(row[0]) if row[0] is not None else None
        qn_number                  = str(row[1]).strip() if row[1] is not None else ""
        title                      = str(row[2]).strip() if row[2] is not None else ""
        problem_statement          = str(row[3]).strip() if row[3] is not None else ""
        solution                   = str(row[4]).strip() if row[4] is not None else ""
        immediate_corrective_action = str(row[5]).strip() if row[5] is not None else ""
        comments                   = str(row[6]).strip() if row[6] is not None else ""
        qn_status                  = str(row[7]).strip() if row[7] is not None else ""
        date_created               = row[8]
        date_closed                = row[9]
        fault                      = str(row[10]).strip() if row[10] is not None else ""
        fault_level_2              = str(row[11]).strip() if row[11] is not None else ""
        assembly                   = str(row[12]).strip() if row[12] is not None else ""
        part_num                   = str(row[13]).strip() if row[13] is not None else ""
        part_desc                  = str(row[14]).strip() if row[14] is not None else ""
        owner                      = str(row[15]).strip() if row[15] is not None else ""
        originator                 = str(row[16]).strip() if row[16] is not None else ""
        nc_manufacturing_site      = str(row[17]).strip() if row[17] is not None else ""
        dmr_disposition            = str(row[18]).strip() if row[18] is not None else ""
        root_cause                 = str(row[19]).strip() if row[19] is not None else ""

        if qn_number:
            found_ids.add(qn_number)

        results.append({
            "qr_id":                       qr_id,
            "qn_number":                   qn_number,
            "title":                       title,
            "problem_statement":           problem_statement,
            "solution":                    solution,
            "immediate_corrective_action": immediate_corrective_action,
            "comments":                    comments,
            "qn_status":                   qn_status,
            "date_created":                date_created,
            "date_closed":                 date_closed,
            "fault":                       fault,
            "fault_level_2":               fault_level_2,
            "assembly":                    assembly,
            "part_num":                    part_num,
            "part_desc":                   part_desc,
            "owner":                       owner,
            "originator":                  originator,
            "nc_manufacturing_site":       nc_manufacturing_site,
            "dmr_disposition":             dmr_disposition,
            "root_cause":                  root_cause,
        })

    not_found = [qn for qn in qn_numbers if qn not in found_ids]
    if not_found:
        print(f"WARNING: No records found for QN number(s): {', '.join(not_found)}")

    return results


def fetch_qn_records(qn_field_value: str) -> Dict[str, Any]:
    """Fetch QN records and split into valid / skipped / not_found.

    Returns a dict with three keys:
        "valid"     - list of dicts with all QN fields
        "skipped"   - list of dicts: qn_number, qn_status
                      (filtered out by inactive status)
        "not_found" - list of str (QN numbers present in input but absent in Databricks)
    """
    qn_numbers = parse_qn_numbers(qn_field_value)
    # Databricks Hive SQL does not support positional ? parameters in IN() clauses.
    # qn_numbers are validated as digit-only strings via re.findall so inlining is safe.
    placeholders = ", ".join(f"'{n}'" for n in qn_numbers)

    sql = f"""
    SELECT
        qr_id,
        sap_notification_number AS qn_number,
        title,
        problem_statement,
        immediate_fix AS solution,
        immediate_corrective_action,
        comments,
        state AS qn_status,
        date_created,
        date_closed,
        fault,
        fault_level_2,
        assembly,
        part_num,
        part_desc,
        owner,
        originator,
        nc_manufacturing_site,
        dmr_disposition,
        root_cause
    FROM prd.pd_tw.manufacturing_nc_summary
    WHERE sap_notification_number IN ({placeholders})
    ORDER BY sap_notification_number
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
        qr_id                      = int(row[0]) if row[0] is not None else None
        qn_number                  = str(row[1]).strip() if row[1] is not None else ""
        title                      = str(row[2]).strip() if row[2] is not None else ""
        problem_statement          = str(row[3]).strip() if row[3] is not None else ""
        solution                   = str(row[4]).strip() if row[4] is not None else ""
        immediate_corrective_action = str(row[5]).strip() if row[5] is not None else ""
        comments                   = str(row[6]).strip() if row[6] is not None else ""
        qn_status                  = str(row[7]).strip() if row[7] is not None else ""
        date_created               = row[8]
        date_closed                = row[9]
        fault                      = str(row[10]).strip() if row[10] is not None else ""
        fault_level_2              = str(row[11]).strip() if row[11] is not None else ""
        assembly                   = str(row[12]).strip() if row[12] is not None else ""
        part_num                   = str(row[13]).strip() if row[13] is not None else ""
        part_desc                  = str(row[14]).strip() if row[14] is not None else ""
        owner                      = str(row[15]).strip() if row[15] is not None else ""
        originator                 = str(row[16]).strip() if row[16] is not None else ""
        nc_manufacturing_site      = str(row[17]).strip() if row[17] is not None else ""
        dmr_disposition            = str(row[18]).strip() if row[18] is not None else ""
        root_cause                 = str(row[19]).strip() if row[19] is not None else ""

        if not qn_number:
            continue

        found_ids.add(qn_number)

        if qn_status in _INACTIVE_STATUSES:
            skipped.append({"qn_number": qn_number, "qn_status": qn_status})
        else:
            valid.append({
                "qr_id":                       qr_id,
                "qn_number":                   qn_number,
                "title":                       title,
                "problem_statement":           problem_statement,
                "solution":                    solution,
                "immediate_corrective_action": immediate_corrective_action,
                "comments":                    comments,
                "qn_status":                   qn_status,
                "date_created":                date_created,
                "date_closed":                 date_closed,
                "fault":                       fault,
                "fault_level_2":               fault_level_2,
                "assembly":                    assembly,
                "part_num":                    part_num,
                "part_desc":                   part_desc,
                "owner":                       owner,
                "originator":                  originator,
                "nc_manufacturing_site":       nc_manufacturing_site,
                "dmr_disposition":             dmr_disposition,
                "root_cause":                  root_cause,
            })

    not_found = [qn for qn in qn_numbers if qn not in found_ids]
    return {"valid": valid, "skipped": skipped, "not_found": not_found}


def print_results(results: List[Dict[str, Any]]) -> None:
    if not results:
        print("No matching QN records found (or all are filtered by status).")
        return

    print(f"\nTotal QN records retrieved: {len(results)}")
    print("-" * 160)
    print(
        f"{'QN Number':<14} | {'Title':<30} | {'Status':<20} | "
        f"{'Problem':<30} | {'Solution':<30} | {'Root Cause':<20}"
    )
    print("-" * 160)

    for rec in results:
        qn     = rec["qn_number"]
        title  = rec["title"]
        status = rec["qn_status"]
        prob   = rec["problem_statement"]
        sol    = rec["solution"]
        rc     = rec["root_cause"]

        title_s  = (title[:27]  + "...") if len(title)  > 30 else title
        prob_s   = (prob[:27]   + "...") if len(prob)   > 30 else prob
        sol_s    = (sol[:27]    + "...") if len(sol)    > 30 else sol
        rc_s     = (rc[:17]     + "...") if len(rc)     > 20 else rc

        print(
            f"{qn:<14} | {title_s:<30} | {status:<20} | "
            f"{prob_s:<30} | {sol_s:<30} | {rc_s:<20}"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Query QN (Quality Note) details from Databricks.")
    parser.add_argument(
        "--qn",
        required=True,
        help="QN number(s) to look up. Comma, space, or newline separated.",
    )
    args = parser.parse_args()

    results = query_qn_details(args.qn)
    print_results(results)


if __name__ == "__main__":
    main()
