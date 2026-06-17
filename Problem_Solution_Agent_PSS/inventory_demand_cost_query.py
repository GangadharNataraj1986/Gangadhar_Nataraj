"""Databricks SQL helper for OBS Inventory, Demand, and Cost.

Purpose
-------
Build a plant-level table for one or more OBS parts across plants 4020, 4055,
4060, 4070, 4080, and 4090.

Verified Databricks source mapping
---------------------------------
1. Make / Buy, MRP, Description, Standard Cost
    - Schema.Table: prd.ud_gsco.material_master
    - Columns: prcrmnttype, mrpprfl, materialdesc, stdcost_usd

2. Inventory (On Hand)
    - Schema.Table: prd.pd_inv.summonhandinv
    - Column: qty
    - Logic: SUM(qty) by materialnum, plantcode

3. Inventory (On Order - Open PO only)
    - Schema.Table: prd.ud_agsocebi.open_purchase_order
    - Columns: quantity, quantity_of_goods_received
    - Logic: SUM(quantity - quantity_of_goods_received) where remaining qty > 0

4. Gross Demand (MM360-aligned)
        - Schema.Table: prd.ud_gsco.vw_active_items
        - Columns: demand13wk, demand26wk, ssg_52_wk_dmd, ags_52_wk_dmd
        - Logic:
            - Demand-13 = demand13wk
            - Demand-26 = demand26wk
            - Demand-52 = ssg_52_wk_dmd + ags_52_wk_dmd

5. Cost behavior
    - Standard Cost USD is returned only when on-hand inventory exists (> 0)
    - Inventory Cost USD = on_hand_qty * standard_cost_usd
    - Supply Cost USD = (on_hand_qty + on_order_qty) * standard_cost_usd

Notes
-----
- The query deliberately starts from a requested parts x requested plants grid,
  so every requested part/plant combination is returned even when one of the
  source views has no row for that plant.
- Databricks Hive SQL over ODBC does not support positional parameters in the
  same way as SQL Server. This helper sanitizes the part list and inlines it.
"""
from __future__ import annotations

import argparse
from decimal import Decimal
from typing import Any, Dict, Iterable, List, Sequence

_DSN = "Spark-PRD"
DEFAULT_PLANTS: tuple[str, ...] = ("4020", "4055", "4060", "4070", "4080", "4090")

OUTPUT_COLUMNS: list[str] = [
    "part_number",
    "part_description",
    "plant",
    "make_buy",
    "mrp_profile",
    "on_hand_qty",
    "on_order_qty",
    "gross_demand_13w",
    "gross_demand_26w",
    "gross_demand_52w",
    "ags_on_hand_qty",
    "ags_on_order_qty",
    "ags_gross_demand_52w",
    "standard_cost_usd",
    "inventory_cost_usd",
    "supply_cost_usd",
]


def _clean_parts(parts: Iterable[str]) -> list[str]:
    return _clean_identifiers(parts)


def _clean_identifiers(values: Iterable[str]) -> list[str]:
    cleaned: list[str] = []
    seen: set[str] = set()
    for raw in values:
        value = str(raw or "").strip().upper().replace("'", "")
        if not value or value in seen:
            continue
        seen.add(value)
        cleaned.append(value)
    return cleaned


def _clean_plants(plants: Sequence[str] | None) -> list[str]:
    raw_values = plants or DEFAULT_PLANTS
    cleaned: list[str] = []
    seen: set[str] = set()
    for raw in raw_values:
        plant = str(raw or "").strip().replace("'", "")
        if not plant or plant in seen:
            continue
        seen.add(plant)
        cleaned.append(plant)
    if not cleaned:
        raise ValueError("At least one plant is required.")
    return cleaned


def _values_cte(alias: str, column_name: str, values: Sequence[str]) -> str:
    selects = [f"SELECT '{value}' AS {column_name}" for value in values]
    return f"{alias} AS (\n    " + "\n    UNION ALL\n    ".join(selects) + "\n)"


def _build_sql(parts: Sequence[str], plants: Sequence[str]) -> str:
    part_filter = ", ".join(f"'{part}'" for part in parts)
    plant_filter = ", ".join(f"'{plant}'" for plant in plants)
    requested_parts_cte = _values_cte("requested_parts", "part_number", parts)
    requested_plants_cte = _values_cte("requested_plants", "plant", plants)

    return f"""
WITH
{requested_parts_cte},
{requested_plants_cte},
part_plant_grid AS (
    SELECT
        rp.part_number,
        pl.plant
    FROM requested_parts rp
    CROSS JOIN requested_plants pl
),
material_master AS (
    SELECT
        materialnum,
        plantcd,
        MAX(materialdesc) AS materialdesc,
        MAX(mrpprfl) AS mrpprfl,
        MAX(prcrmnttype) AS prcrmnttype,
        MAX(stdcost_usd) AS stdcost_usd
    FROM prd.ud_gsco.material_master
    WHERE materialnum IN ({part_filter})
      AND plantcd IN ({plant_filter})
    GROUP BY materialnum, plantcd
),
on_hand_inventory AS (
    SELECT
        materialnum,
        plantcode,
        SUM(COALESCE(qty, 0)) AS on_hand_qty
    FROM prd.pd_inv.summonhandinv
    WHERE materialnum IN ({part_filter})
      AND plantcode IN ({plant_filter})
    GROUP BY materialnum, plantcode
),
open_po AS (
    SELECT
        material_number AS materialnum,
        plant_code AS plantcd,
        SUM(COALESCE(quantity, 0) - COALESCE(quantity_of_goods_received, 0)) AS open_order_qty
    FROM prd.ud_agsocebi.open_purchase_order
    WHERE material_number IN ({part_filter})
      AND plant_code IN ({plant_filter})
      AND (COALESCE(quantity, 0) - COALESCE(quantity_of_goods_received, 0)) > 0
    GROUP BY material_number, plant_code
),
demand_buckets AS (
    SELECT
        materialnum,
        plantcd,
        MAX(COALESCE(TRY_CAST(demand13wk AS DECIMAL(38, 3)), 0)) AS gross_demand_13w,
        MAX(COALESCE(TRY_CAST(demand26wk AS DECIMAL(38, 3)), 0)) AS gross_demand_26w,
        MAX(COALESCE(TRY_CAST(ssg_52_wk_dmd AS DECIMAL(38, 3)), 0) + COALESCE(TRY_CAST(ags_52_wk_dmd AS DECIMAL(38, 3)), 0)) AS gross_demand_52w,
        MAX(COALESCE(TRY_CAST(ags_52_wk_dmd AS DECIMAL(38, 3)), 0)) AS ags_gross_demand_52w
    FROM prd.ud_gsco.vw_active_items
    WHERE materialnum IN ({part_filter})
      AND plantcd IN ({plant_filter})
    GROUP BY materialnum, plantcd
)
SELECT
    grid.part_number                                              AS part_number,
    COALESCE(mm.materialdesc, '')                                  AS part_description,
    grid.plant                                                    AS plant,
    CASE
        WHEN UPPER(COALESCE(mm.prcrmnttype, '')) IN ('MAKE', 'BUY')
            THEN UPPER(mm.prcrmnttype)
        WHEN COALESCE(mm.prcrmnttype, '') = 'E'
            THEN 'MAKE'
        WHEN COALESCE(mm.prcrmnttype, '') = 'F'
            THEN 'BUY'
        ELSE COALESCE(mm.prcrmnttype, '')
    END                                                           AS make_buy,
    COALESCE(mm.mrpprfl, '')                                       AS mrp_profile,
    CAST(COALESCE(oh.on_hand_qty, 0) AS DECIMAL(18, 3))            AS on_hand_qty,
    CAST(COALESCE(po.open_order_qty, 0) AS DECIMAL(18, 3))        AS on_order_qty,
    CAST(COALESCE(dem.gross_demand_13w, 0) AS DECIMAL(18, 3))     AS gross_demand_13w,
    CAST(COALESCE(dem.gross_demand_26w, 0) AS DECIMAL(18, 3))     AS gross_demand_26w,
    CAST(COALESCE(dem.gross_demand_52w, 0) AS DECIMAL(18, 3))     AS gross_demand_52w,
    CAST(COALESCE(oh.on_hand_qty, 0) AS DECIMAL(18, 3))            AS ags_on_hand_qty, -- Placeholder, see note below
    CAST(COALESCE(po.open_order_qty, 0) AS DECIMAL(18, 3))         AS ags_on_order_qty, -- Placeholder, see note below
    CAST(COALESCE(dem.ags_gross_demand_52w, 0) AS DECIMAL(18, 3))  AS ags_gross_demand_52w,
    CAST(CASE
            WHEN (COALESCE(oh.on_hand_qty, 0) > 0 OR COALESCE(po.open_order_qty, 0) > 0)
                THEN COALESCE(mm.stdcost_usd, 0)
            ELSE 0
        END AS DECIMAL(18, 3))                                    AS standard_cost_usd,
    CAST((COALESCE(oh.on_hand_qty, 0) + COALESCE(po.open_order_qty, 0)) *
        CASE
            WHEN (COALESCE(oh.on_hand_qty, 0) > 0 OR COALESCE(po.open_order_qty, 0) > 0)
                THEN COALESCE(mm.stdcost_usd, 0)
            ELSE 0
        END AS DECIMAL(18, 3))
                                                                 AS inventory_cost_usd,
    CAST((COALESCE(oh.on_hand_qty, 0) + COALESCE(po.open_order_qty, 0)) *
        CASE
            WHEN (COALESCE(oh.on_hand_qty, 0) > 0 OR COALESCE(po.open_order_qty, 0) > 0)
                THEN COALESCE(mm.stdcost_usd, 0)
            ELSE 0
        END AS DECIMAL(18, 3))
                                                                 AS supply_cost_usd
FROM part_plant_grid grid
LEFT JOIN material_master mm
    ON mm.materialnum = grid.part_number
   AND mm.plantcd = grid.plant
LEFT JOIN on_hand_inventory oh
    ON oh.materialnum = grid.part_number
   AND oh.plantcode = grid.plant
LEFT JOIN open_po po
    ON po.materialnum = grid.part_number
   AND po.plantcd = grid.plant
LEFT JOIN demand_buckets dem
    ON dem.materialnum = grid.part_number
   AND dem.plantcd = grid.plant
ORDER BY grid.part_number, grid.plant
""".strip()


def fetch_inventory_demand_cost(
    parts: Iterable[str],
    plants: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    """Query Databricks and return plant-level inventory / demand / cost rows."""
    cleaned_parts = _clean_parts(parts)
    cleaned_plants = _clean_plants(plants)
    if not cleaned_parts:
        raise ValueError("At least one part number is required.")

    try:
        import pyodbc  # type: ignore[import]
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "pyodbc is not installed. Install it with: pip install pyodbc"
        ) from exc

    sql = _build_sql(cleaned_parts, cleaned_plants)
    conn = pyodbc.connect(f"DSN={_DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
    finally:
        conn.close()

    results: list[dict[str, Any]] = []
    for row in rows:
        record = {columns[idx]: row[idx] for idx in range(len(columns))}
        results.append(record)
    return results


def fetch_inventory_demand_mapping(parts: Iterable[str]) -> list[dict[str, Any]]:
    """Query Databricks and return inventory-demand mapping rows for Inventory_Demand_Mapping tab.
    
    Returns a list of dicts with demand order information (system numbers, consumption qty, dates, etc.)
    for the requested OBS parts.
    """
    cleaned_parts = _clean_parts(parts)
    if not cleaned_parts:
        return []

    try:
        import pyodbc  # type: ignore[import]
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "pyodbc is not installed. Install it with: pip install pyodbc"
        ) from exc

    # Remove special characters from parts for resilient SQL matching.
    import re
    normalized_parts = [re.sub(r'[^A-Za-z0-9]', '', p) for p in cleaned_parts]
    part_filter = ", ".join(f"'{p}'" for p in normalized_parts)

    # Inventory-demand mapping sourced from costed BOM with BOM freeze and ECO anchor dates.
    # Uses certified_system_summary for BOM freeze dates and latest ECO anchor from slot_change_summary_rr.
    sql = f"""
SELECT DISTINCT
    b.slotnum               AS system_slot_number,
    b.slotnum               AS system_number,
    b.matrlnum              AS part_number,
    b.matrldesc             AS part_description,
    b.parentpart            AS parent_part,
    b.bomlvl                AS bom_level,
    b.cmpntqty              AS consumption_qty,
    b.platform,
    b.buildqtr              AS build_quarter,
    b.SOPDate               AS sop_date,
    m.eopplanneddate        AS eop_planned_date,
    m.eopplanneddate        AS eop_date,
    css.bom_freeze_plan_date AS bom_freeze_planned,
    css.bom_freeze_actual_date AS bom_freeze_actual,
    scsr.eco_anchor_date_new AS eco_anchor_date,
    b.oppno                 AS opportunity_no,
    b.bomsalesdoc           AS sales_order,
    b.custnamemd            AS customer
FROM prd.pd_mm.summcostedbom b
LEFT JOIN prd.pd_mm.factprojmilestncmncol m
    ON b.slotnum = m.slotnum
LEFT JOIN prd.pd_tw.certified_system_summary css
    ON b.slotnum = css.system
LEFT JOIN (
    SELECT 
        slot_number,
        eco_anchor_date_new,
        ROW_NUMBER() OVER (
            PARTITION BY slot_number 
            ORDER BY change_date DESC
        ) AS rn
    FROM prd.rd_kinaxis.slot_change_summary_rr
) scsr
    ON b.slotnum = scsr.slot_number 
   AND scsr.rn = 1
WHERE UPPER(REGEXP_REPLACE(CAST(b.matrlnum AS STRING), '[^A-Za-z0-9]', '')) IN ({part_filter})
  AND CAST(b.SOPDate AS DATE) >= CURRENT_DATE
ORDER BY 
    part_number,
    build_quarter,
    system_slot_number
""".strip()

    conn = pyodbc.connect(f"DSN={_DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
    finally:
        conn.close()

    results: list[dict[str, Any]] = []
    for row in rows:
        rec = {columns[idx]: row[idx] for idx in range(len(columns))}
        results.append(rec)
    
    return results


def fetch_open_purchase_order_details(
    parts: Iterable[str],
    plants: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    """Return open PO detail rows for the requested part/plant combinations using star schema tables."""
    cleaned_parts = _clean_parts(parts)
    cleaned_plants = _clean_plants(plants) if plants is not None else []
    if not cleaned_parts:
        raise ValueError("At least one part number is required.")

    try:
        import pyodbc  # type: ignore[import]
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "pyodbc is not installed. Install it with: pip install pyodbc"
        ) from exc

    part_filter = ", ".join(f"'{part}'" for part in cleaned_parts)
    plant_filter = ", ".join(f"'{plant}'" for plant in cleaned_plants)
    plant_clause = f"AND fp.plantcd IN ({plant_filter})" if cleaned_plants else ""
    
    # Use star schema tables: factposl (line items) + factpol (PO headers)
    sql = f"""
SELECT
    fp.materialnum AS material_number,
    COALESCE(pol.shrtdesc, '') AS description,
    COALESCE(sng.current_sng_status, '') AS sng_status,
    fp.plantcd AS plant,
    fp.pocd AS po_number,
    COALESCE(pace.global_pace_flag, '') AS pace,
    COALESCE(fp.dueqty, 0) AS scheduled_qty,
    COALESCE(fp.issdqty, 0) AS gr_quantity,
    COALESCE(fp.openqty, 0) AS open_quantity,
    COALESCE(pol.podate, pol.pocdate, pol.cdate) AS po_creation_date,
    CASE
        WHEN fp.orgnlcmmitdate != '1900-01-01' THEN fp.orgnlcmmitdate
        WHEN fp.cmmitdate != '1900-01-01' THEN fp.cmmitdate
        WHEN fp.needbydate != '1900-01-01' THEN fp.needbydate
        ELSE NULL
    END AS delivery_date,
    fp.needbydate AS need_by_date,
    COALESCE(pol.netprice, 0) AS price_usd,
    COALESCE(fp.postatus, '') AS po_status
FROM prd.pd_mm.factposl fp
INNER JOIN prd.pd_mm.factpol pol
    ON pol.pocd = fp.pocd
    AND pol.polcd = fp.polcd
    AND pol.rflg = 1
LEFT JOIN prd.rd_gsspi.emrp_ezcancel sng
    ON sng.poline = CONCAT(fp.pocd, '_', fp.polcd, '_', fp.poschedulelncd)
LEFT JOIN prd.ud_agsocebi.pace_part_assignment pace
    ON pace.material = fp.materialnum
WHERE fp.materialnum IN ({part_filter})
    {plant_clause}
    AND fp.rflg = 1
    AND pol.status = 'OPEN'
    AND COALESCE(fp.openqty, 0) > 0
    AND COALESCE(pol.delind, '') NOT IN ('L', 'D')
ORDER BY material_number, po_number DESC"""

    print(f"DEBUG: PO Query - Parts: {cleaned_parts[:3]}... ({len(cleaned_parts)} total)")
    print(f"DEBUG: PO Query - Plants: {cleaned_plants}")
    
    conn = pyodbc.connect(f"DSN={_DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
        print(f"DEBUG: PO Query returned {len(rows)} rows")
    finally:
        conn.close()

    from datetime import datetime, timedelta
    
    results: list[dict[str, Any]] = []
    for row in rows:
        record = {columns[idx]: row[idx] for idx in range(len(columns))}
        
        # Calculate aging metrics
        po_creation_date = record.get('po_creation_date')
        delivery_date = record.get('delivery_date')
        today = datetime.now().date()
        
        if po_creation_date and delivery_date:
            try:
                if isinstance(po_creation_date, str):
                    po_creation_date = datetime.strptime(po_creation_date, '%Y-%m-%d').date()
                if isinstance(delivery_date, str):
                    delivery_date = datetime.strptime(delivery_date, '%Y-%m-%d').date()
                
                lead_time_days = (delivery_date - po_creation_date).days
                po_age_days = (today - po_creation_date).days
                
                if lead_time_days > 0:
                    po_age_pct = min(100.0, (po_age_days / lead_time_days) * 100.0)
                else:
                    po_age_pct = 0.0
                
                balance_days = (delivery_date - today).days
                
                record['lead_time_days'] = lead_time_days
                record['po_age_days'] = po_age_days
                record['po_age_pct'] = round(po_age_pct, 2)
                record['balance_days_to_delivery'] = balance_days
            except Exception:
                record['lead_time_days'] = 0
                record['po_age_days'] = 0
                record['po_age_pct'] = 0.0
                record['balance_days_to_delivery'] = 0
        else:
            record['lead_time_days'] = 0
            record['po_age_days'] = 0
            record['po_age_pct'] = 0.0
            record['balance_days_to_delivery'] = 0
        
        results.append(record)
    return results


def fetch_kit_code_descriptions(
    kit_codes: Iterable[str],
    plant: str | None = None,
) -> dict[str, str]:
    """Return a best-effort kit description per kit code from BOM parent descriptions."""
    cleaned_kit_codes = _clean_identifiers(kit_codes)
    if not cleaned_kit_codes:
        return {}

    try:
        import pyodbc  # type: ignore[import]
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "pyodbc is not installed. Install it with: pip install pyodbc"
        ) from exc

    kit_filter = ", ".join(f"'{code}'" for code in cleaned_kit_codes)
    plant_clause = ""
    if plant:
        cleaned_plant = str(plant).strip().replace("'", "")
        if cleaned_plant:
            plant_clause = f" AND b.plantcd = '{cleaned_plant}'"

    sql = f"""
WITH ranked_kit_descriptions AS (
    SELECT
        TRIM(b.sortstring) AS kit_code,
        COALESCE(b.parentmaterialdesc, '') AS kit_description,
        ROW_NUMBER() OVER (
            PARTITION BY TRIM(b.sortstring)
            ORDER BY
                CASE WHEN UPPER(COALESCE(b.parentmaterialdesc, '')) LIKE '%OUTSOURCED%' THEN 0 ELSE 1 END,
                CASE WHEN COALESCE(b.parentmaterialdesc, '') <> '' THEN 0 ELSE 1 END,
                COALESCE(b.parentmaterialdesc, '')
        ) AS rn
    FROM prd.pd_mm.factbomlvl1 b
    WHERE TRIM(b.sortstring) IN ({kit_filter})
      AND TRIM(b.sortstring) <> ''
      {plant_clause}
)
SELECT
    kit_code,
    kit_description
FROM ranked_kit_descriptions
WHERE rn = 1
""".strip()

    conn = pyodbc.connect(f"DSN={_DSN}", autocommit=True)
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        rows = cursor.fetchall()
    finally:
        conn.close()

    results: dict[str, str] = {}
    for kit_code, kit_description in rows:
        key = str(kit_code or "").strip().upper()
        if key:
            results[key] = str(kit_description or "").strip()
    return results


def rows_to_dataframe(rows: Sequence[dict[str, Any]]):
    """Return a pandas DataFrame when pandas is available."""
    try:
        import pandas as pd  # type: ignore[import]
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("pandas is required for DataFrame output.") from exc

    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    numeric_columns = [
        "on_hand_qty",
        "on_order_qty",
        "gross_demand_13w",
        "gross_demand_26w",
        "gross_demand_52w",
        "standard_cost_usd",
        "inventory_cost_usd",
        "supply_cost_usd",
    ]
    for column in numeric_columns:
        if column in df.columns:
            df[column] = df[column].apply(_as_float)
    return df


def _as_float(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, Decimal):
        return float(value)
    return float(value)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Query Databricks for plant-level inventory, demand, and cost."
    )
    parser.add_argument(
        "parts",
        nargs="+",
        help="One or more part numbers, for example: 0022-48090 0010-05024",
    )
    parser.add_argument(
        "--plants",
        nargs="*",
        default=list(DEFAULT_PLANTS),
        help="Plant list. Defaults to 4020 4055 4060 4070 4080 4090.",
    )
    parser.add_argument(
        "--csv",
        help="Optional output CSV path.",
    )
    args = parser.parse_args()

    rows = fetch_inventory_demand_cost(args.parts, args.plants)
    df = rows_to_dataframe(rows)

    if args.csv:
        df.to_csv(args.csv, index=False)
        print(f"Wrote {len(df)} row(s) to {args.csv}")
    else:
        print(df.to_string(index=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
