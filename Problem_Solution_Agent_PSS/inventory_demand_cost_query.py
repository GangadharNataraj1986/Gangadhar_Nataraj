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

4. Gross Demand (latest snapshot, excluding PlannedPO)
    - Schema.Table: prd.pd_mm.factknxssplyconsmp
    - Columns: snapshotdate, spart, partsite, stype, dduedate, dqty
    - Logic:
      - latest snapshot per (spart, partsite)
      - exclude stype = 'PlannedPO'
      - bucket and sum TRY_CAST(dqty) into 13/26/52-week windows

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
    "standard_cost_usd",
    "inventory_cost_usd",
    "supply_cost_usd",
]


def _clean_parts(parts: Iterable[str]) -> list[str]:
    cleaned: list[str] = []
    seen: set[str] = set()
    for raw in parts:
        part = str(raw or "").strip().upper().replace("'", "")
        if not part or part in seen:
            continue
        seen.add(part)
        cleaned.append(part)
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
latest_snapshot AS (
    SELECT
        spart AS materialnum,
        partsite AS plantcd,
        MAX(snapshotdate) AS snapshotdate
    FROM prd.pd_mm.factknxssplyconsmp
    WHERE spart IN ({part_filter})
      AND partsite IN ({plant_filter})
      AND LOWER(COALESCE(stype, '')) IN ('quotation', 'gforecast')
    GROUP BY spart, partsite
),
demand_buckets AS (
    SELECT
        f.spart AS materialnum,
        f.partsite AS plantcd,
        SUM(CASE
                WHEN DATEDIFF(TRY_CAST(f.dduedate AS DATE), CURRENT_DATE()) BETWEEN 0 AND 91
                    THEN COALESCE(TRY_CAST(f.dqty AS DECIMAL(38, 3)), 0)
                ELSE 0
            END) AS gross_demand_13w,
        SUM(CASE
                WHEN DATEDIFF(TRY_CAST(f.dduedate AS DATE), CURRENT_DATE()) BETWEEN 0 AND 182
                    THEN COALESCE(TRY_CAST(f.dqty AS DECIMAL(38, 3)), 0)
                ELSE 0
            END) AS gross_demand_26w,
        SUM(CASE
                WHEN DATEDIFF(TRY_CAST(f.dduedate AS DATE), CURRENT_DATE()) BETWEEN 0 AND 364
                    THEN COALESCE(TRY_CAST(f.dqty AS DECIMAL(38, 3)), 0)
                ELSE 0
            END) AS gross_demand_52w
    FROM prd.pd_mm.factknxssplyconsmp f
    INNER JOIN latest_snapshot ls
        ON f.spart = ls.materialnum
       AND f.partsite = ls.plantcd
       AND f.snapshotdate = ls.snapshotdate
    WHERE LOWER(COALESCE(f.stype, '')) IN ('quotation', 'gforecast')
    GROUP BY f.spart, f.partsite
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
    CAST(CASE
            WHEN COALESCE(oh.on_hand_qty, 0) > 0 THEN COALESCE(mm.stdcost_usd, 0)
            ELSE 0
        END AS DECIMAL(18, 3))                                    AS standard_cost_usd,
    CAST(COALESCE(oh.on_hand_qty, 0) *
        CASE
            WHEN COALESCE(oh.on_hand_qty, 0) > 0 THEN COALESCE(mm.stdcost_usd, 0)
            ELSE 0
        END AS DECIMAL(18, 3))
                                                                 AS inventory_cost_usd,
    CAST((COALESCE(oh.on_hand_qty, 0) + COALESCE(po.open_order_qty, 0)) *
        CASE
            WHEN COALESCE(oh.on_hand_qty, 0) > 0 THEN COALESCE(mm.stdcost_usd, 0)
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
