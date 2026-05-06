"""Databricks SQL helper for OBS Inventory, Demand, and Cost.

Purpose
-------
Build a plant-level table for one or more OBS parts across plants 4020, 4055,
4060, 4070, 4080, and 4090.

Verified Databricks source mapping
---------------------------------
1. Make / Buy
   - Schema.Table: prd.ud_gsco.vw_active_items
   - Column: prcrmnttype
   - Fallback: prd.ud_gsco.material_master.prcrmnttype

2. MRP Profile
   - Schema.Table: prd.ud_gsco.vw_active_items
   - Column: mrpprfl
   - Fallback: prd.ud_gsco.material_master.mrpprfl

3. Inventory
   - On Hand (OH)
     - Schema.Table: prd.ud_gsco.vw_active_items
     - Column: nettableoh
   - On Order (OO)
     - Schema.Table: prd.pd_mm.factpol
     - Column: openpoqty
     - Logic: SUM(openpoqty) by materialnum, plantcd

4. Gross Demand
   - 13 weeks
     - Schema.Table: prd.ud_gsco.vw_active_items
     - Column: demand13wk
   - 26 weeks
     - Schema.Table: prd.ud_gsco.vw_active_items
     - Column: demand26wk
   - 52 weeks
     - Schema.Table: prd.ud_gsco.vw_active_items
     - Columns: ssg_52_wk_dmd, ags_52_wk_dmd
     - Logic: COALESCE(ssg_52_wk_dmd, 0) + COALESCE(ags_52_wk_dmd, 0)

5. Standard Cost / Inventory Cost
   - Standard Cost USD
     - Schema.Table: prd.ud_gsco.vw_active_items
     - Column: stdcost
     - Fallback: prd.ud_gsco.material_master.stdcost_usd
   - Inventory Cost USD
     - Logic: On Hand Qty * Standard Cost USD
   - Supply Cost USD
     - Logic: (On Hand Qty + On Order Qty) * Standard Cost USD
     - Included because the current UI tab historically treated OH + OO as the
       cost basis.

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
        materialdesc,
        mrpprfl,
        prcrmnttype,
        stdcost_usd
    FROM prd.ud_gsco.material_master
    WHERE materialnum IN ({part_filter})
      AND plantcd IN ({plant_filter})
),
active_items AS (
    SELECT
        materialnum,
        plantcd,
        materialdesc,
        prcrmnttype,
        mrpprfl,
        nettableoh,
        demand13wk,
        demand26wk,
        COALESCE(ssg_52_wk_dmd, 0) + COALESCE(ags_52_wk_dmd, 0) AS demand52wk,
        stdcost
    FROM prd.ud_gsco.vw_active_items
    WHERE materialnum IN ({part_filter})
      AND plantcd IN ({plant_filter})
),
open_po AS (
    SELECT
        materialnum,
        plantcd,
        SUM(COALESCE(openpoqty, 0)) AS open_order_qty
    FROM prd.pd_mm.factpol
    WHERE materialnum IN ({part_filter})
      AND plantcd IN ({plant_filter})
      AND rflg = 1
    GROUP BY materialnum, plantcd
)
SELECT
    grid.part_number                                              AS part_number,
    COALESCE(ai.materialdesc, mm.materialdesc, '')                AS part_description,
    grid.plant                                                    AS plant,
    CASE
        WHEN UPPER(COALESCE(ai.prcrmnttype, '')) IN ('MAKE', 'BUY')
            THEN UPPER(ai.prcrmnttype)
        WHEN COALESCE(mm.prcrmnttype, '') = 'E'
            THEN 'MAKE'
        WHEN COALESCE(mm.prcrmnttype, '') = 'F'
            THEN 'BUY'
        ELSE COALESCE(ai.prcrmnttype, mm.prcrmnttype, '')
    END                                                           AS make_buy,
    COALESCE(ai.mrpprfl, mm.mrpprfl, '')                          AS mrp_profile,
    CAST(COALESCE(ai.nettableoh, 0) AS DECIMAL(18, 3))            AS on_hand_qty,
    CAST(COALESCE(po.open_order_qty, 0) AS DECIMAL(18, 3))        AS on_order_qty,
    CAST(COALESCE(ai.demand13wk, 0) AS DECIMAL(18, 3))            AS gross_demand_13w,
    CAST(COALESCE(ai.demand26wk, 0) AS DECIMAL(18, 3))            AS gross_demand_26w,
    CAST(COALESCE(ai.demand52wk, 0) AS DECIMAL(18, 3))            AS gross_demand_52w,
    CAST(COALESCE(ai.stdcost, mm.stdcost_usd, 0) AS DECIMAL(18, 3)) AS standard_cost_usd,
    CAST(COALESCE(ai.nettableoh, 0) * COALESCE(ai.stdcost, mm.stdcost_usd, 0) AS DECIMAL(18, 3))
                                                                 AS inventory_cost_usd,
    CAST((COALESCE(ai.nettableoh, 0) + COALESCE(po.open_order_qty, 0))
        * COALESCE(ai.stdcost, mm.stdcost_usd, 0) AS DECIMAL(18, 3))
                                                                 AS supply_cost_usd
FROM part_plant_grid grid
LEFT JOIN material_master mm
    ON mm.materialnum = grid.part_number
   AND mm.plantcd = grid.plant
LEFT JOIN active_items ai
    ON ai.materialnum = grid.part_number
   AND ai.plantcd = grid.plant
LEFT JOIN open_po po
    ON po.materialnum = grid.part_number
   AND po.plantcd = grid.plant
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
