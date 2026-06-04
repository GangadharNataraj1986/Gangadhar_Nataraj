"""Databricks SQL helper - Implemented BOM (up to 18 levels).

Connects via the pre-configured ODBC DSN 'Spark-PRD' on Windows.

Public API
----------
fetch_implemented_bom(parts, max_level=18, plant='4070', include_level0=True) -> list[dict]
    Query Databricks and return multi-level BOM rows.

DISPLAY_HEADERS : list[str]
    Ordered display column headers for a table view.

DB_KEYS : list[str]
    Lowercase SQL alias names returned in each row dict.
"""
from __future__ import annotations

import threading
from typing import Any

_DSN = "Spark-PRD"
_THREAD_LOCAL = threading.local()
_IN_CHUNK_SIZE = 400

DB_KEYS: list[str] = [
    "bom_level",
    "part",
    "tool_comments",
    "rev_ln",
    "plant",
    "description",
    "item_status",
    "base_qty",
    "ext_qty",
    "uom",
    "eco_number",
    "effectivity_date",
    "item_seq",
    "kit_code",
    "sparable_flag",
    "designator",
    "option_class",
    "procurement_type",
    "user_item_type",
    "pace_or_dash",
    "mlo_class",
    "input_part",
    "tool_comments_color",
    "tool_comments_bold",
]

DISPLAY_HEADERS: list[str] = [
    "BOM Level",
    "Part",
    "Tool comments",
    "Rev/Ln",
    "Plant",
    "Description",
    "Item Status",
    "Base Qty",
    "Ext Qty",
    "UOM",
    "ECO Number",
]

_SQL_LVL0 = """
SELECT DISTINCT
  m.materialnum                   AS part,
    m.rvnlvl                        AS rev_ln,
  m.plantcd                       AS plant,
  m.materialdesc                  AS description,
  COALESCE(ms.desc, '')           AS item_status,
  CAST(1 AS DECIMAL(13,3))        AS base_qty,
    CAST(1 AS DECIMAL(13,3))        AS ext_qty,
  m.uomcd                         AS uom,
  CAST(NULL AS STRING)            AS eco_number,
  CAST(NULL AS STRING)            AS effectivity_date,
  CAST(NULL AS STRING)            AS item_seq,
  CAST(NULL AS STRING)            AS kit_code,
  CAST(NULL AS STRING)            AS sparable_flag,
    CAST(NULL AS STRING)            AS designator,
    CAST(NULL AS STRING)            AS option_class,
    CASE COALESCE(UPPER(m.prcrmnttype), '')
        WHEN 'E' THEN 'Make'
        WHEN 'F' THEN 'Buy'
        ELSE COALESCE(m.prcrmnttype, '')
    END                             AS procurement_type,
  m.useritemtype                  AS user_item_type,
  m.program_type                  AS pace_or_dash,
  CAST('' AS STRING)              AS mlo_class,
  m.materialnum                   AS input_part
FROM prd.ud_gsco.material_master m
LEFT JOIN prd.pd_gcat.dimmaterialstatus ms
  ON m.crossplantmaterialstatus = ms.materialstatuscd
WHERE m.materialnum IN ({placeholders})
  AND m.plantcd = '{plant}'
"""

_SQL_LVL1 = """
SELECT DISTINCT
  b.cmpntmaterialnum              AS part,
  b.materialnum                   AS parent_part,
    b.rvnlvl                        AS rev_ln,
  b.plantcd                       AS plant,
  COALESCE(m.materialdesc, '')    AS description,
  COALESCE(ms.desc, '')           AS item_status,
  b.cmpntqty                      AS base_qty,
    b.cmpntqty                      AS ext_qty,
  b.cmpntuom                      AS uom,
  b.cmpntchngnum                  AS eco_number,
  b.cmpntvalidfrom                AS effectivity_date,
  CAST(b.cmpntlnnum AS STRING)    AS item_seq,
  b.sortstring                    AS kit_code,
  b.sprprtind                     AS sparable_flag,
    COALESCE(ol.refdesignator, '')  AS designator,
    COALESCE(ol.option_class, '')   AS option_class,
    CASE COALESCE(UPPER(m.prcrmnttype), '')
        WHEN 'E' THEN 'Make'
        WHEN 'F' THEN 'Buy'
        ELSE COALESCE(m.prcrmnttype, '')
    END                             AS procurement_type,
  COALESCE(m.useritemtype, '')    AS user_item_type,
  COALESCE(m.program_type, '')    AS pace_or_dash,
  CAST('' AS STRING)              AS mlo_class,
  b.materialnum                   AS input_part
FROM prd.pd_mm.factbomlvl1 b
LEFT JOIN prd.ud_gsco.material_master m
  ON b.cmpntmaterialnum = m.materialnum
 AND b.plantcd          = m.plantcd
LEFT JOIN prd.pd_gcat.dimmaterialstatus ms
  ON m.crossplantmaterialstatus = ms.materialstatuscd
LEFT JOIN (
    SELECT option, option_class, refdesignator
    FROM (
        SELECT
            option,
            option_class,
            refdesignator,
            ROW_NUMBER() OVER (
                PARTITION BY option
                ORDER BY last_update_tsmp DESC
            ) AS rn
        FROM prd.rd_dsworkbench.datax_bom_option_library
        WHERE rflag = '1'
    ) opt
    WHERE opt.rn = 1
) ol
    ON b.cmpntmaterialnum = ol.option
WHERE b.materialnum IN ({placeholders})
  AND b.plantcd = '{plant}'
  AND b.rflg = 1
  AND COALESCE(UPPER(b.cmpntactn), '') <> 'DISABLE'
  AND COALESCE(UPPER(ms.desc), '') NOT IN ('PLC_OBSOLETE', 'PLC_INACTIVATE')
"""


def _run(cursor: Any, sql: str) -> tuple[list, list[str]]:
    cursor.execute(sql)
    rows = cursor.fetchall()
    col_names = [d[0].lower() for d in cursor.description]
    return rows, col_names


def _set_designator(rec: dict[str, Any]) -> None:
    """Normalize reference-designator variants to rec['designator']."""
    if rec.get("designator"):
        rec["designator"] = str(rec.get("designator", "")).strip()
        return
    for key in ("reference_designator", "ref_designator", "reference designator"):
        val = rec.get(key)
        if val is not None and str(val).strip():
            rec["designator"] = str(val).strip()
            return
    rec["designator"] = ""


def _in_list(parts: list[str]) -> str:
    return ", ".join(f"'{p}'" for p in parts)


def _chunked(items: list[str], chunk_size: int) -> list[list[str]]:
    return [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]


def _get_conn(pyodbc: Any) -> Any:
    if getattr(_THREAD_LOCAL, "conn", None) is None:
        _THREAD_LOCAL.conn = pyodbc.connect(f"DSN={_DSN}", autocommit=True)
    return _THREAD_LOCAL.conn


def _reset_conn() -> None:
    conn = getattr(_THREAD_LOCAL, "conn", None)
    if conn is not None:
        try:
            conn.close()
        except Exception:
            pass
        _THREAD_LOCAL.conn = None


def _as_float(value: Any, default: float = 1.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def fetch_implemented_bom(
    parts: list[str],
    max_level: int = 18,
    plant: str = "4070",
    include_level0: bool = True,
) -> list[dict[str, Any]]:
    """Fetch Implemented BOM rows for 1..18 levels.

    Parameters
    ----------
    parts : list[str]
        Input part numbers to query.
    max_level : int
        Maximum BOM depth to retrieve (1 to 18).
    plant : str
        Plant code filter (one of 4020/4055/4060/4070/4080/4090).
    include_level0 : bool
        When True, include a synthetic level-0/input row for each input part.

    Returns
    -------
    list[dict[str, Any]]
        Rows keyed by DB_KEYS. Includes level-0 (optional) and level 1..N rows.
    """
    if not parts:
        raise ValueError("parts list must not be empty.")
    if not (1 <= max_level <= 18):
        raise ValueError(f"max_level must be between 1 and 18, got {max_level}.")

    known_plants = {"4020", "4055", "4060", "4070", "4080", "4090"}
    safe_plant = str(plant).strip().replace("'", "")
    if safe_plant not in known_plants:
        raise ValueError(f"plant must be one of {sorted(known_plants)}, got {plant!r}.")

    safe_parts = list(dict.fromkeys(
        str(p).strip().replace("'", "") for p in parts if str(p).strip()
    ))
    if not safe_parts:
        raise ValueError("parts contains only blank values after stripping.")

    try:
        import pyodbc  # type: ignore[import]
    except ImportError as exc:
        raise RuntimeError("pyodbc is not installed. Install it with: pip install pyodbc") from exc

    records: list[dict[str, Any]] = []

    conn = _get_conn(pyodbc)
    try:
        cursor = conn.cursor()
        try:
            if include_level0:
                l0_seen: set[str] = set()
                for chunk in _chunked(safe_parts, _IN_CHUNK_SIZE):
                    rows0, cols0 = _run(
                        cursor,
                        _SQL_LVL0.format(placeholders=_in_list(chunk), plant=safe_plant),
                    )
                    for row in rows0:
                        rec = {
                            col: ("" if val is None else str(val).strip())
                            for col, val in zip(cols0, row)
                        }
                        _set_designator(rec)
                        rec["bom_level"] = "0"
                        rec["tool_comments"] = ""
                        rec["tool_comments_color"] = ""
                        rec["tool_comments_bold"] = "0"
                        rec["__path"] = (str(rec.get("input_part", rec.get("part", ""))).upper(),)
                        records.append(rec)
                        l0_seen.add(str(rec.get("part", "")).upper())

                # Keep root rows even when material master has no entry at the selected plant.
                for p in safe_parts:
                    if p.upper() in l0_seen:
                        continue
                    records.append({
                        "bom_level": "0",
                        "part": p,
                        "tool_comments": "",
                        "rev_ln": "",
                        "plant": safe_plant,
                        "description": "",
                        "item_status": "",
                        "base_qty": "1",
                        "ext_qty": "1",
                        "uom": "",
                        "eco_number": "",
                        "effectivity_date": "",
                        "item_seq": "",
                        "kit_code": "",
                        "sparable_flag": "",
                        "designator": "",
                        "option_class": "",
                        "procurement_type": "",
                        "user_item_type": "",
                        "pace_or_dash": "",
                        "mlo_class": "",
                        "input_part": p,
                        "tool_comments_color": "",
                        "tool_comments_bold": "0",
                        "__path": (p.upper(),),
                    })

            # Traverse child BOM depth-first by level (parent -> children).
            current_nodes: list[dict[str, Any]] = [
                {
                    "current_part": p,
                    "input_part": p,
                    "path": (p.upper(),),
                    "cum_qty": 1.0,
                }
                for p in safe_parts
            ]

            for level in range(1, max_level + 1):
                if not current_nodes:
                    break

                lookup_parts = list(dict.fromkeys(node["current_part"] for node in current_nodes))
                by_parent: dict[str, list[dict[str, Any]]] = {}

                for chunk in _chunked(lookup_parts, _IN_CHUNK_SIZE):
                    rows_n, cols_n = _run(
                        cursor,
                        _SQL_LVL1.format(placeholders=_in_list(chunk), plant=safe_plant),
                    )
                    for row in rows_n:
                        rec = {
                            col: ("" if val is None else str(val).strip())
                            for col, val in zip(cols_n, row)
                        }
                        _set_designator(rec)
                        pkey = str(rec.get("parent_part", "")).upper()
                        by_parent.setdefault(pkey, []).append(rec)

                # Aggregate traversal nodes by (input_part, full path) so repeated
                # child lines under the same parent contribute cumulative quantity
                # without duplicating entire deeper-level subtrees.
                next_nodes_map: dict[tuple[str, tuple[str, ...]], dict[str, Any]] = {}
                for source in current_nodes:
                    parent_part = str(source.get("current_part", ""))
                    parent_key = parent_part.upper()
                    # De-duplicate exact component lines only (not just child part),
                    # so valid repeated child part rows with different line attributes
                    # are preserved in levels 2..6.
                    seen_component_for_parent: set[tuple[str, str, str, str, str, str, str]] = set()

                    for rec in by_parent.get(parent_key, []):
                        child_part = str(rec.get("part", "")).strip()
                        child_key = child_part.upper()
                        component_key = (
                            child_key,
                            str(rec.get("item_seq", "")),
                            str(rec.get("kit_code", "")),
                            str(rec.get("effectivity_date", "")),
                            str(rec.get("eco_number", "")),
                            str(rec.get("base_qty", "")),
                            str(rec.get("uom", "")),
                        )
                        if not child_key or component_key in seen_component_for_parent:
                            continue
                        seen_component_for_parent.add(component_key)

                        base_qty = _as_float(rec.get("base_qty"), default=1.0)
                        ext_qty = round(_as_float(source.get("cum_qty"), default=1.0) * base_qty, 6)

                        rec_out = dict(rec)
                        rec_out["bom_level"] = str(level)
                        rec_out["input_part"] = str(source.get("input_part", ""))
                        rec_out["ext_qty"] = str(ext_qty)
                        rec_out["tool_comments"] = "Removed BOM Item"
                        rec_out.pop("parent_part", None)
                        rec_out["__path"] = source["path"] + (child_key,)
                        # UI can apply these hints directly for orange+bold style.
                        rec_out["tool_comments_color"] = "#FFA500"
                        rec_out["tool_comments_bold"] = "1"
                        records.append(rec_out)

                        if child_key not in source["path"]:
                            next_path = source["path"] + (child_key,)
                            next_key = (str(source.get("input_part", "")), next_path)
                            existing = next_nodes_map.get(next_key)
                            if existing is None:
                                next_nodes_map[next_key] = {
                                    "current_part": child_part,
                                    "input_part": source["input_part"],
                                    "path": next_path,
                                    "cum_qty": ext_qty,
                                }
                            else:
                                existing["cum_qty"] = round(
                                    _as_float(existing.get("cum_qty"), default=0.0) + ext_qty,
                                    6,
                                )

                current_nodes = list(next_nodes_map.values())
        finally:
            cursor.close()
    except Exception:
        _reset_conn()
        raise

    # Keep order grouped by requested input part, then BOM traversal path
    # so deeper levels stay visually under their correct parents.
    order = {p.upper(): i for i, p in enumerate(safe_parts)}

    def _sort_key(rec: dict[str, Any]) -> tuple[int, tuple[str, ...], str, str]:
        input_part = str(rec.get("input_part", "")).upper()
        path = rec.get("__path")
        if isinstance(path, tuple):
            path_key = tuple(str(x) for x in path)
        else:
            path_key = (input_part, str(rec.get("part", "")).upper())
        item_seq = str(rec.get("item_seq", ""))
        part = str(rec.get("part", "")).upper()
        return (order.get(input_part, 10**9), path_key, item_seq, part)

    records.sort(key=_sort_key)
    for rec in records:
        rec.pop("__path", None)
    return records
