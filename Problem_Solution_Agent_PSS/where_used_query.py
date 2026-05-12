"""Databricks SQL helper – Where Used of OBS Parts (multi-level BOM).

Connects via the pre-configured ODBC DSN 'Spark-PRD' on Windows.

Public API
----------
fetch_where_used(obs_parts, max_level) -> list[dict]
    Query Databricks and return one dict per result row.

DISPLAY_HEADERS : list[str]
    Ordered display column headers for the Where Used table in the UI.
    "Replacement" (index 2) is a UI-only column auto-filled from the OBS map;
    it is NOT a Databricks column.

DB_KEYS : list[str]
    Lowercase SQL alias names returned in each row dict.
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import threading
from typing import Any

_DSN = "Spark-PRD"
_THREAD_LOCAL = threading.local()  # per-thread ODBC connection; safe for parallel workers
_IN_CHUNK_SIZE = 1000              # max parts per SQL IN list before splitting into chunks
_PARALLEL_SHARD_SIZE = 120         # per-query part count when parallelizing large runs
_MAX_PARALLEL_WORKERS = 4          # safe upper bound for concurrent ODBC query threads

# ---------------------------------------------------------------------------
# Persistent shared connection for the Orphan Analysis WU import flow.
# Unlike _THREAD_LOCAL (which creates a new connection per thread), this
# connection is module-level and reused across multiple button clicks and
# threads.  Only one non-parallel query uses it at a time.
# ---------------------------------------------------------------------------
_PERSISTENT_CONN: Any = None
_PERSISTENT_CONN_LOCK = threading.Lock()

# Lowercase SQL aliases returned in each row dict (matches the SELECT aliases below).
DB_KEYS: list[str] = [
    "wu_level",
    "part",
    "rev_ln",
    "plant",
    "description",
    "item_status",
    "base_qty",
    "ext_qty",
    "uom",
    "eco_number",
    "procurement_type",
    "effectivity_date",
    "user_item_type",
    "item_seq",
    "kit_code",
    "sparable_flag",
    "designator",
    "option_class",
    "pace_or_dash",
    "mlo_class",
    "input_part",   # internal – not shown as its own column in the UI
]

# Display header labels that appear in the Where Used table.
# Order: Select(col 0 – injected by UI), WU Level(1), Part(2), Replacement(3 – UI-only),
# then the remaining Databricks columns from col 4 onwards.
DISPLAY_HEADERS: list[str] = [
    "WU Level",
    "Part",
    "Replacement",          # UI-only, auto-filled from OBS Parts map
    "Rev/Ln",
    "Plant",
    "Description",
    "Item Status",
    "Base Qty",
    "Ext Qty",
    "UOM",
    "ECO Number",
    "Procurement Type",
    "Effectivity Date",
    "User Item Type",
    "Item Seq",
    "Kit Code",
    "Sparable flag",
    "Designator",
    "Option Class",
    "PACE",
    "MLO Class",
]

# ---------------------------------------------------------------------------
# Iterative per-level query approach
# ---------------------------------------------------------------------------
# Why iterative instead of one big UNION ALL / recursive CTE?
#
#   • UNION ALL with N join-chains: Databricks scans factbomlvl1 N*(N+1)/2 times
#     in a single query plan → very slow for deep levels.
#   • WITH RECURSIVE: hits the 1 M-row recursion limit on large BOM data, and
#     the SET spark.sql.recursivelimit config is not available via ODBC.
#   • Iterative: one small focused query per level, each scanning only the rows
#     that match the previous level's parents → fast, bounded, no limits.
# ---------------------------------------------------------------------------

# SQL for level 0: input parts' own master-data rows.
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
  m.prcrmnttype                   AS procurement_type,
  CAST(NULL AS STRING)            AS effectivity_date,
  m.useritemtype                  AS user_item_type,
  CAST(NULL AS STRING)            AS item_seq,
  CAST(NULL AS STRING)            AS kit_code,
  CAST(NULL AS STRING)            AS sparable_flag,
    CAST(NULL AS STRING)            AS designator,
    CAST(NULL AS STRING)            AS option_class,
  m.program_type                  AS pace_or_dash,
  CAST('' AS STRING)              AS mlo_class,
  m.materialnum                   AS input_part
FROM prd.ud_gsco.material_master m
LEFT JOIN prd.pd_gcat.dimmaterialstatus ms
  ON m.crossplantmaterialstatus = ms.materialstatuscd
WHERE m.materialnum IN ({placeholders})
  AND m.plantcd = '{plant}'
"""

# SQL for each level 1-N: find parent assemblies of the current part set.
# {placeholders} = IN-list of the current level's component part numbers.
# Note: b.sprprtind (sparable_flag) is from factbomlvl1 table and indicates
#       whether a component part is sparable (Y/N or 1/0 indicator).
_SQL_LEVEL_N = """
SELECT DISTINCT
  b.materialnum                       AS part,
  b.cmpntmaterialnum                  AS child_part,
  b.rvnlvl                            AS rev_ln,
  b.plantcd                           AS plant,
  b.parentmaterialdesc                AS description,
  COALESCE(ms.desc, '')               AS item_status,
  b.cmpntqty                          AS base_qty,
  b.cmpntuom                          AS uom,
  b.cmpntchngnum                      AS eco_number,
  b.cmpntvalidfrom                    AS effectivity_date,
  CAST(b.cmpntlnnum AS STRING)        AS item_seq,
  b.sortstring                        AS kit_code,
  b.sprprtind                         AS sparable_flag,
    COALESCE(ol.refdesignator, '')      AS designator,
    COALESCE(ol.option_class, '')       AS option_class,
  COALESCE(m.prcrmnttype, '')         AS procurement_type,
  COALESCE(m.useritemtype, '')        AS user_item_type,
  COALESCE(m.program_type, '')        AS pace_or_dash,
  CAST('' AS STRING)              AS mlo_class
FROM prd.pd_mm.factbomlvl1 b
LEFT JOIN prd.ud_gsco.material_master m
  ON b.materialnum = m.materialnum
 AND b.plantcd     = m.plantcd
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
WHERE b.cmpntmaterialnum IN ({placeholders})
  AND b.plantcd = '{plant}'
        AND b.rflg = 1
        AND COALESCE(UPPER(b.cmpntactn), '') <> 'DISABLE'
        AND (
            (UPPER(b.materialnum) LIKE 'ESW%' AND LENGTH(b.materialnum) > 10)
            OR (LENGTH(b.materialnum) = 10 AND SUBSTRING(b.materialnum, 5, 1) = '-')
        )
"""

# Traversal-only variant of _SQL_LEVEL_N: scans only factbomlvl1 (no metadata JOINs)
# so each hop is faster.  Metadata (item_status, procurement_type, etc.) is fetched
# in a single batch query after all traversal levels complete (_SQL_METADATA_BATCH).
_SQL_LEVEL_N_TRAVERSE = """
SELECT DISTINCT
  b.materialnum                       AS part,
  b.cmpntmaterialnum                  AS child_part,
  b.rvnlvl                            AS rev_ln,
  b.plantcd                           AS plant,
  b.parentmaterialdesc                AS description,
  b.cmpntqty                          AS base_qty,
  b.cmpntuom                          AS uom,
  b.cmpntchngnum                      AS eco_number,
  b.cmpntvalidfrom                    AS effectivity_date,
  CAST(b.cmpntlnnum AS STRING)        AS item_seq,
  b.sortstring                        AS kit_code,
  b.sprprtind                         AS sparable_flag
    , COALESCE(ol.refdesignator, '')    AS designator
    , COALESCE(ol.option_class, '')     AS option_class
FROM prd.pd_mm.factbomlvl1 b
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
WHERE b.cmpntmaterialnum IN ({placeholders})
  AND b.plantcd = '{plant}'
  AND b.rflg = 1
  AND COALESCE(UPPER(b.cmpntactn), '') <> 'DISABLE'
    AND (
        (UPPER(b.materialnum) LIKE 'ESW%' AND LENGTH(b.materialnum) > 10)
        OR (LENGTH(b.materialnum) = 10 AND SUBSTRING(b.materialnum, 5, 1) = '-')
    )
"""

# Batch metadata enrichment for all unique parent parts discovered during traversal.
# Replaces N per-level LEFT JOINs to material_master + dimmaterialstatus with one query.
_SQL_METADATA_BATCH = """
SELECT DISTINCT
  m.materialnum                          AS part,
  COALESCE(ms.desc, '')                  AS item_status,
  COALESCE(m.prcrmnttype, '')            AS procurement_type,
  COALESCE(m.useritemtype, '')           AS user_item_type,
  COALESCE(m.program_type, '')           AS pace_or_dash,
  CAST('' AS STRING)                     AS mlo_class
FROM prd.ud_gsco.material_master m
LEFT JOIN prd.pd_gcat.dimmaterialstatus ms
  ON m.crossplantmaterialstatus = ms.materialstatuscd
WHERE m.materialnum IN ({placeholders})
  AND m.plantcd = '{plant}'
"""


def _run(cursor: Any, sql: str) -> tuple[list, list[str]]:
    """Execute *sql* and return (rows, lowercase_col_names)."""
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


def _in_list(parts: "set[str] | list[str]") -> str:
    """Return a safe SQL IN-list string for the given part numbers."""
    return ", ".join(f"'{p}'" for p in parts)


def _run_query_in_chunks(
    cursor: Any,
    sql_template: str,
    parts: list[str],
    plant: str,
) -> tuple[list, list[str]]:
    """Execute the same IN-list SQL in chunks and combine all rows.

    Large IN-lists can be slow to parse/plan in Databricks over ODBC. Splitting
    very large inputs into bounded chunks gives steadier query times and avoids
    oversized SQL text.
    """
    if not parts:
        return [], []

    safe_parts = list(dict.fromkeys(parts))
    all_rows: list = []
    all_cols: list[str] = []

    for i in range(0, len(safe_parts), _IN_CHUNK_SIZE):
        chunk = safe_parts[i:i + _IN_CHUNK_SIZE]
        rows, cols = _run(
            cursor,
            sql_template.format(placeholders=_in_list(chunk), plant=plant),
        )
        if not all_cols:
            all_cols = cols
        all_rows.extend(rows)

    return all_rows, all_cols


def _run_query_parallel_shards(
    sql_template: str,
    parts: list[str],
    plant: str,
    shard_size: int = _PARALLEL_SHARD_SIZE,
    max_workers: int = _MAX_PARALLEL_WORKERS,
) -> tuple[list, list[str]]:
    """Execute IN-list SQL across part shards concurrently.

    Useful when one large query is slower than several smaller queries run in
    parallel. Each worker uses its own thread-local ODBC connection.
    """
    safe_parts = list(dict.fromkeys(parts))
    if not safe_parts:
        return [], []

    shards = [safe_parts[i:i + shard_size] for i in range(0, len(safe_parts), shard_size)]
    if len(shards) <= 1:
        # Keep single-shard behavior simple and deterministic.
        import pyodbc  # type: ignore[import]
        conn = _get_conn(pyodbc)
        cursor = conn.cursor()
        try:
            return _run(
                cursor,
                sql_template.format(placeholders=_in_list(safe_parts), plant=plant),
            )
        finally:
            cursor.close()

    def _query_one(chunk: list[str]) -> tuple[list, list[str]]:
        import pyodbc  # type: ignore[import]
        conn = _get_conn(pyodbc)
        cursor = conn.cursor()
        try:
            return _run(
                cursor,
                sql_template.format(placeholders=_in_list(chunk), plant=plant),
            )
        finally:
            cursor.close()

    all_rows: list = []
    all_cols: list[str] = []
    worker_count = min(max_workers, len(shards))
    with ThreadPoolExecutor(max_workers=worker_count) as ex:
        for rows, cols in ex.map(_query_one, shards):
            if not all_cols:
                all_cols = cols
            all_rows.extend(rows)

    return all_rows, all_cols


def _get_conn(pyodbc: Any) -> Any:
    """Return the thread-local cached ODBC connection, creating one if needed."""
    if getattr(_THREAD_LOCAL, 'conn', None) is None:
        _THREAD_LOCAL.conn = pyodbc.connect(f"DSN={_DSN}", autocommit=True)
    return _THREAD_LOCAL.conn


def _reset_conn() -> None:
    """Close and discard the current thread's cached connection so the next call reconnects."""
    conn = getattr(_THREAD_LOCAL, 'conn', None)
    if conn is not None:
        try:
            conn.close()
        except Exception:
            pass
        _THREAD_LOCAL.conn = None


def _get_persistent_conn(pyodbc: Any) -> Any:
    """Return the module-level persistent ODBC connection, creating it once if needed.

    Unlike _get_conn(), this connection is NOT thread-local.  It is created once
    and reused across all threads and button clicks.  _PERSISTENT_CONN_LOCK must
    be held for the full duration of a query to prevent concurrent use.
    """
    global _PERSISTENT_CONN
    if _PERSISTENT_CONN is None:
        _PERSISTENT_CONN = pyodbc.connect(f"DSN={_DSN}", autocommit=True)
    return _PERSISTENT_CONN


def _reset_persistent_conn() -> None:
    """Close and clear the persistent connection (e.g. after a connection error)."""
    global _PERSISTENT_CONN
    with _PERSISTENT_CONN_LOCK:
        if _PERSISTENT_CONN is not None:
            try:
                _PERSISTENT_CONN.close()
            except Exception:
                pass
            _PERSISTENT_CONN = None


def pre_warm_connection() -> None:
    """Pre-establish the persistent Databricks connection on the calling thread.

    Call this once at application startup (in a background thread) so that the
    first real query does not pay the 20-120s cluster cold-start cost.
    """
    try:
        import pyodbc  # type: ignore[import]
    except ImportError:
        return
    with _PERSISTENT_CONN_LOCK:
        _get_persistent_conn(pyodbc)


def _is_valid_part(part: str) -> bool:
    """Return True if *part* looks like a real part number.

    Rules (same as the existing _apply_cleanup_rules in the UI file-import flow):
      1. Starts with 'ESW' (case-insensitive) and is longer than 10 characters
         e.g. ESW0241-00001, ESW1234-ABCDE
      2. Exactly 10 characters with a dash at position 4
         e.g. 0190-12345, ABCD-EFGHI

    Everything else – fabrication codes, process tags, build codes like
    'C904172TSMC26-MAKE' – is treated as noise and excluded.
    """
    p = part.strip()
    if not p:
        return False
    # Rule 1: ESW part numbers
    if p.upper().startswith('ESW') and len(p) > 10:
        return True
    # Rule 2: standard 10-char dash format  (XXXX-XXXXX)
    if len(p) == 10 and p[4] == '-':
        return True
    return False


def _to_tree_order(records: list[dict], root_parts: list[str]) -> list[dict]:
    """Reorder flat records into depth-first bottom-up tree order.

    Uses *root_parts* (the original safe_parts list) as the definitive L0
    anchors, in input order.  This means even if an OBS part has no row in
    material_master (and therefore no L0 SQL result), its parent records are
    still emitted correctly.

    For each root part:
      1. Emit its L0 record (if one was returned by the SQL), otherwise skip.
      2. Immediately recurse into its L1 parents, then each parent's parents,
         giving the depth-first tree: root → L1 → L2 → next L1 → its L2 → …

    Key-matching is case-insensitive so Databricks case variations don't break
    the parent-child links.
    """
    from collections import defaultdict

    # Build lookups with upper-cased keys for case-insensitive matching.
    # Parent-child links are scoped by input_part + chain path key so repeated
    # child part numbers in different branches do not get cross-mapped.
    l0_by_part: dict[str, dict] = {}
    by_child_path: dict[tuple[str, str], list] = defaultdict(list)

    for r in records:
        input_part = r.get("input_part", "").upper()
        if r.get("wu_level", "0") == "0":
            l0_by_part[r.get("part", "").upper()] = r
        else:
            child_path_key = r.get("_child_path_key", "")
            by_child_path[(input_part, child_path_key)].append(r)

    def _sibling_order_key(rec: dict) -> tuple[int, str]:
        part = str(rec.get("part", "")).strip().upper()
        return (1 if part.startswith("9024") else 0, part)

    for child_key, siblings in by_child_path.items():
        by_child_path[child_key] = sorted(siblings, key=_sibling_order_key)

    result: list[dict] = []

    def _dfs(input_part_upper: str, node_path_key: str) -> None:
        for parent_rec in by_child_path.get((input_part_upper, node_path_key), []):
            result.append(parent_rec)
            _dfs(input_part_upper, parent_rec.get("_node_path_key", ""))

    for p in root_parts:
        p_upper = p.upper()
        l0 = l0_by_part.get(p_upper)
        if l0 is not None:
            result.append(l0)
        _dfs(p_upper, p_upper)

    return result


def fetch_where_used(obs_parts: list[str], max_level: int, plant: str = "4070") -> list[dict[str, Any]]:
    """Query Databricks for multi-level Where Used data for the given OBS parts.

    Uses one focused query per BOM level (iterative) instead of a single large
    UNION ALL or recursive CTE, which keeps each query fast regardless of depth.

    Parameters
    ----------
    obs_parts : list[str]
        Non-empty list of OBS part numbers from the OBS Parts column in the UI.
        Values are stripped and sanitised (single-quotes removed) before use in SQL.
    max_level : int
        Maximum WU level depth to retrieve (inclusive). Must be 1 to 6.
    plant : str
        Plant code to filter BOM and material master rows (e.g. "4070").
        Must be one of the known plant codes; defaults to "4070".

    Returns
    -------
    list[dict]
        One dict per result row.  Keys match DB_KEYS (all lowercase).
        ``wu_level`` is the numeric level as a string.
        ``part`` is the raw part number without any display indentation
        (the UI applies indentation when rendering the table).
        ``ext_qty`` is the running product of cmpntqty values along the path
        from the input part to this row.

    Raises
    ------
    ValueError
        ``obs_parts`` is empty / ``max_level`` is outside 1–6 / ``plant`` is invalid.
    RuntimeError
        pyodbc is not installed, or a Databricks query failed.
    """
    if not obs_parts:
        raise ValueError("obs_parts list must not be empty.")
    if not (1 <= max_level <= 6):
        raise ValueError(f"max_level must be between 1 and 6, got {max_level}.")
    _KNOWN_PLANTS = {"4020", "4055", "4060", "4070", "4080", "4090"}
    safe_plant = str(plant).strip().replace("'", "")
    if safe_plant not in _KNOWN_PLANTS:
        raise ValueError(f"plant must be one of {sorted(_KNOWN_PLANTS)}, got {plant!r}.")

    try:
        import pyodbc  # type: ignore[import]
    except ImportError as exc:
        raise RuntimeError(
            "pyodbc is not installed.  Install it with: pip install pyodbc"
        ) from exc

    # Sanitise: strip + remove single-quotes (no ? params in Databricks Hive SQL).
    safe_parts = list(dict.fromkeys(
        str(p).strip().replace("'", "") for p in obs_parts if str(p).strip()
    ))
    if not safe_parts:
        raise ValueError("obs_parts contains only blank values after stripping.")

    records: list[dict[str, Any]] = []

    conn = _get_conn(pyodbc)
    try:
        cursor = conn.cursor()
        try:
            # ── Level 0: input parts' own master-data rows ────────────────────
            rows0, cols0 = _run_query_in_chunks(
                cursor,
                _SQL_LVL0,
                safe_parts,
                safe_plant,
            )
            for row in rows0:
                rec: dict[str, Any] = {
                    col: ("" if val is None else str(val).strip())
                    for col, val in zip(cols0, row)
                }
                _set_designator(rec)
                rec["wu_level"] = "0"
                rec["child_part"] = ""   # L0 rows have no child; needed for tree traversal
                rec["_node_path_key"] = rec.get("input_part", rec.get("part", "")).upper()
                # input_part is already set to the part itself in the SQL
                records.append(rec)

            # Rule 1: Level 0 is always the queried child part, even if there is
            # no material master row for the selected plant.
            existing_l0 = {r.get("part", "").upper() for r in records if r.get("wu_level") == "0"}
            for part in safe_parts:
                if part.upper() in existing_l0:
                    continue
                records.append({
                    "part": part,
                    "rev_ln": "",
                    "plant": safe_plant,
                    "description": "",
                    "item_status": "",
                    "base_qty": "1",
                    "ext_qty": "1",
                    "uom": "",
                    "eco_number": "",
                    "procurement_type": "",
                    "effectivity_date": "",
                    "user_item_type": "",
                    "item_seq": "",
                    "kit_code": "",
                    "sparable_flag": "",
                    "designator": "",
                    "option_class": "",
                    "pace_or_dash": "",
                    "mlo_class": "",
                    "input_part": part,
                    "wu_level": "0",
                    "child_part": "",
                    "_node_path_key": part.upper(),
                })

            # ── Levels 1..max_level: strict structural traversal ───────────────
            # Each node is a single parent chain state. We expand each chain by
            # exactly one parent-child hop per level, dedupe only duplicate parents
            # for the same child within that chain.
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
                    break  # nothing to look up at this level

                lookup_parts = list(dict.fromkeys(node["current_part"] for node in current_nodes))
                rows_n, cols_n = _run_query_in_chunks(
                    cursor,
                    _SQL_LEVEL_N,
                    lookup_parts,
                    safe_plant,
                )

                rows_by_child: dict[str, list[dict[str, Any]]] = {}
                for row in rows_n:
                    rec = {
                        col: ("" if val is None else str(val).strip())
                        for col, val in zip(cols_n, row)
                    }
                    _set_designator(rec)
                    child_key = rec.get("child_part", "").upper()
                    rows_by_child.setdefault(child_key, []).append(rec)

                next_nodes: list[dict[str, Any]] = []

                for source in current_nodes:
                    child = source["current_part"]
                    child_key = child.upper()
                    child_path_key = "->".join(source["path"])
                    seen_parent_for_source: set[str] = set()
                    for rec in rows_by_child.get(child_key, []):
                        parent = rec.get("part", "")
                        parent_key = parent.upper()
                        if not _is_valid_part(parent):
                            continue
                        if not parent_key or parent_key in seen_parent_for_source:
                            continue
                        seen_parent_for_source.add(parent_key)

                        # Running extended qty = parent's cumulative qty × this row's base qty.
                        try:
                            cum = float(source.get("cum_qty") or 1.0)
                            bq = float(rec.get("base_qty") or 1.0)
                            ext = round(cum * bq, 6)
                        except (ValueError, TypeError):
                            ext = float(rec.get("base_qty") or 1.0)

                        rec_out = dict(rec)
                        rec_out["wu_level"] = str(level)
                        rec_out["ext_qty"] = str(ext)
                        rec_out["input_part"] = source.get("input_part", child)
                        rec_out["_child_path_key"] = child_path_key
                        rec_out["_node_path_key"] = "->".join(source["path"] + (parent_key,))
                        records.append(rec_out)

                        # Continue only one direct hop upward. Allow the same parent to
                        # appear again if it is reached through a different chain, but do
                        # not create a cycle inside the same chain path.
                        if parent_key not in source["path"]:
                            next_nodes.append({
                                "current_part": parent,
                                "input_part": rec_out["input_part"],
                                "path": source["path"] + (parent_key,),
                                "cum_qty": ext,
                            })

                current_nodes = next_nodes

        finally:
            cursor.close()
    except Exception:
        _reset_conn()
        raise

    # Reorder into depth-first tree order using safe_parts as the definitive
    # root order.  This works even if some OBS parts have no material_master row
    # (no L0 SQL result) – their parent records are still reached via DFS.
    return _to_tree_order(records, safe_parts)


def fetch_where_used_parents_only(
    obs_parts: list[str],
    max_level: int,
    plant: str = "4070",
    log_callback=None,
) -> list[dict[str, Any]]:
    """Like fetch_where_used() but skips the level-0 Databricks lookup.

    Level-0 rows are synthesised locally (blank metadata fields) instead of
    being fetched from material_master.  This saves one full Databricks
    round-trip — for level 1 that halves the number of queries from 2 to 1.

    Use this for the 'WU of Removed BOM Items' tab where:
      • The removed-child parts are already known from the Imp BOM table.
      • Level-0 rows are needed only as block anchors for orphan analysis,
        not for their metadata (description, item_status, etc.).

    Parameters and return value are identical to fetch_where_used().
    """
    if not obs_parts:
        raise ValueError("obs_parts list must not be empty.")
    if not (1 <= max_level <= 6):
        raise ValueError(f"max_level must be between 1 and 6, got {max_level}.")
    _KNOWN_PLANTS = {"4020", "4055", "4060", "4070", "4080", "4090"}
    safe_plant = str(plant).strip().replace("'", "")
    if safe_plant not in _KNOWN_PLANTS:
        raise ValueError(f"plant must be one of {sorted(_KNOWN_PLANTS)}, got {plant!r}.")

    try:
        import pyodbc  # type: ignore[import]
    except ImportError as exc:
        raise RuntimeError(
            "pyodbc is not installed.  Install it with: pip install pyodbc"
        ) from exc

    safe_parts = list(dict.fromkeys(
        str(p).strip().replace("'", "") for p in obs_parts if str(p).strip()
    ))
    if not safe_parts:
        raise ValueError("obs_parts contains only blank values after stripping.")

    records: list[dict[str, Any]] = []

    # ── Level 0: build synthetic root rows locally (no Databricks query) ──────
    # These blank rows anchor each removed-child part in the tree so that the
    # orphan-analysis logic in the UI can find WU Level 0 block boundaries.
    for part in safe_parts:
        records.append({
            "part": part,
            "rev_ln": "",
            "plant": safe_plant,
            "description": "",
            "item_status": "",
            "base_qty": "1",
            "ext_qty": "1",
            "uom": "",
            "eco_number": "",
            "procurement_type": "",
            "effectivity_date": "",
            "user_item_type": "",
            "item_seq": "",
            "kit_code": "",
            "sparable_flag": "",
            "designator": "",
            "option_class": "",
            "pace_or_dash": "",
            "mlo_class": "",
            "input_part": part,
            "wu_level": "0",
            "child_part": "",
            "_node_path_key": part.upper(),
        })

    # ── Levels 1..max_level: traversal-only (no metadata JOINs per level) ─────
    # Using _SQL_LEVEL_N_TRAVERSE scans only factbomlvl1, removing the per-level
    # LEFT JOINs to material_master and dimmaterialstatus.  After all traversal
    # levels complete, one _SQL_METADATA_BATCH query fetches metadata for all
    # unique discovered parents in a single round-trip.
    import time as _time

    def _log(msg: str) -> None:
        if log_callback:
            log_callback(msg)

    _log(f"[{_time.strftime('%H:%M:%S')}] Connecting to Databricks (Spark-PRD)...")
    _t_conn = _time.perf_counter()
    with _PERSISTENT_CONN_LOCK:
        try:
            conn = _get_persistent_conn(pyodbc)
            _already = _time.perf_counter() - _t_conn < 0.5
            if _already:
                _log(f"[{_time.strftime('%H:%M:%S')}] Reusing existing connection (instant).")
            else:
                _log(f"[{_time.strftime('%H:%M:%S')}] Connected in {_time.perf_counter() - _t_conn:.1f}s.")
        except Exception:
            _reset_persistent_conn()
            raise
        try:
            cursor = conn.cursor()
            try:
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
                    _log(f"[{_time.strftime('%H:%M:%S')}] Level {level}: querying {len(lookup_parts)} part(s)...")
                    _tq = _time.perf_counter()
                    rows_n, cols_n = _run_query_in_chunks(
                        cursor,
                        _SQL_LEVEL_N_TRAVERSE,
                        lookup_parts,
                        safe_plant,
                    )
                    _log(f"[{_time.strftime('%H:%M:%S')}] Level {level} done in {_time.perf_counter() - _tq:.1f}s — {len(rows_n)} rows returned.")

                    rows_by_child: dict[str, list[dict[str, Any]]] = {}
                    for row in rows_n:
                        rec = {
                            col: ("" if val is None else str(val).strip())
                            for col, val in zip(cols_n, row)
                        }
                        _set_designator(rec)
                        child_key = rec.get("child_part", "").upper()
                        rows_by_child.setdefault(child_key, []).append(rec)

                    next_nodes: list[dict[str, Any]] = []

                    for source in current_nodes:
                        child = source["current_part"]
                        child_key = child.upper()
                        child_path_key = "->".join(source["path"])
                        seen_parent_for_source: set[str] = set()
                        for rec in rows_by_child.get(child_key, []):
                            parent = rec.get("part", "")
                            parent_key = parent.upper()
                            if not _is_valid_part(parent):
                                continue
                            if not parent_key or parent_key in seen_parent_for_source:
                                continue
                            seen_parent_for_source.add(parent_key)

                            try:
                                cum = float(source.get("cum_qty") or 1.0)
                                bq = float(rec.get("base_qty") or 1.0)
                                ext = round(cum * bq, 6)
                            except (ValueError, TypeError):
                                ext = float(rec.get("base_qty") or 1.0)

                            rec_out = dict(rec)
                            rec_out["wu_level"] = str(level)
                            rec_out["ext_qty"] = str(ext)
                            rec_out["input_part"] = source.get("input_part", child)
                            rec_out["_child_path_key"] = child_path_key
                            rec_out["_node_path_key"] = "->".join(source["path"] + (parent_key,))
                            rec_out.setdefault("item_status", "")
                            rec_out.setdefault("procurement_type", "")
                            rec_out.setdefault("user_item_type", "")
                            rec_out.setdefault("pace_or_dash", "")
                            rec_out.setdefault("mlo_class", "")
                            rec_out.setdefault("designator", "")
                            rec_out.setdefault("option_class", "")
                            records.append(rec_out)

                            if parent_key not in source["path"]:
                                next_nodes.append({
                                    "current_part": parent,
                                    "input_part": rec_out["input_part"],
                                    "path": source["path"] + (parent_key,),
                                    "cum_qty": ext,
                                })

                    current_nodes = next_nodes

                # ── Single metadata batch for all discovered parents ───────────
                parent_parts = list(dict.fromkeys(
                    r["part"] for r in records if r.get("wu_level", "0") != "0" and r.get("part")
                ))
                if parent_parts:
                    _log(f"[{_time.strftime('%H:%M:%S')}] Fetching metadata for {len(parent_parts)} unique parent part(s)...")
                    _tm = _time.perf_counter()
                    meta_map: dict[str, dict] = {}
                    mr, mc = _run_query_in_chunks(
                        cursor,
                        _SQL_METADATA_BATCH,
                        parent_parts,
                        safe_plant,
                    )
                    _log(f"[{_time.strftime('%H:%M:%S')}] Metadata fetch done in {_time.perf_counter() - _tm:.1f}s.")
                    for mrow in mr:
                        md = {
                            col: ("" if val is None else str(val).strip())
                            for col, val in zip(mc, mrow)
                        }
                        meta_map[md["part"].upper()] = md
                    for rec in records:
                        if rec.get("wu_level", "0") == "0":
                            continue
                        meta = meta_map.get(rec["part"].upper(), {})
                        rec["item_status"]      = meta.get("item_status", "")
                        rec["procurement_type"] = meta.get("procurement_type", "")
                        rec["user_item_type"]   = meta.get("user_item_type", "")
                        rec["pace_or_dash"]     = meta.get("pace_or_dash", "")
                        rec["mlo_class"]        = meta.get("mlo_class", "")

            finally:
                cursor.close()
        except Exception:
            _reset_persistent_conn()
            raise

    _log(f"[{_time.strftime('%H:%M:%S')}] Sorting {len(records)} record(s) into tree order...")
    return _to_tree_order(records, safe_parts)


def fetch_where_used_level1_fast(
    obs_parts: list[str],
    plant: str = "4070",
    log_callback=None,
) -> list[dict[str, Any]]:
    """Fast path for single-level where-used (L1 only) with minimal overhead.

    This is optimized for the "Orphan Child alone" flow:
      - Builds L0 synthetic rows locally.
      - Runs only one traversal query on factbomlvl1.
      - Skips metadata enrichment query entirely.

    Returned rows still follow the same dict schema used by the UI table.
    Metadata-heavy fields (item_status, procurement_type, etc.) are returned
    as blank strings to prioritize speed.
    """
    if not obs_parts:
        raise ValueError("obs_parts list must not be empty.")

    _KNOWN_PLANTS = {"4020", "4055", "4060", "4070", "4080", "4090"}
    safe_plant = str(plant).strip().replace("'", "")
    if safe_plant not in _KNOWN_PLANTS:
        raise ValueError(f"plant must be one of {sorted(_KNOWN_PLANTS)}, got {plant!r}.")

    try:
        import pyodbc  # type: ignore[import]
    except ImportError as exc:
        raise RuntimeError(
            "pyodbc is not installed.  Install it with: pip install pyodbc"
        ) from exc

    safe_parts = list(dict.fromkeys(
        str(p).strip().replace("'", "") for p in obs_parts if str(p).strip()
    ))
    if not safe_parts:
        raise ValueError("obs_parts contains only blank values after stripping.")

    records: list[dict[str, Any]] = []

    # Level-0 anchors (no Databricks query).
    for part in safe_parts:
        records.append({
            "part": part,
            "rev_ln": "",
            "plant": safe_plant,
            "description": "",
            "item_status": "",
            "base_qty": "1",
            "ext_qty": "1",
            "uom": "",
            "eco_number": "",
            "procurement_type": "",
            "effectivity_date": "",
            "user_item_type": "",
            "item_seq": "",
            "kit_code": "",
            "sparable_flag": "",
            "designator": "",
            "option_class": "",
            "pace_or_dash": "",
            "mlo_class": "",
            "input_part": part,
            "wu_level": "0",
            "child_part": "",
            "_node_path_key": part.upper(),
        })

    import time as _time

    def _log(msg: str) -> None:
        if log_callback:
            log_callback(msg)

    _log(f"[{_time.strftime('%H:%M:%S')}] Connecting to Databricks (Spark-PRD)...")
    _t_conn = _time.perf_counter()
    with _PERSISTENT_CONN_LOCK:
        try:
            conn = _get_persistent_conn(pyodbc)
            _already = _time.perf_counter() - _t_conn < 0.5
            if _already:
                _log(f"[{_time.strftime('%H:%M:%S')}] Reusing existing connection (instant).")
            else:
                _log(f"[{_time.strftime('%H:%M:%S')}] Connected in {_time.perf_counter() - _t_conn:.1f}s. Querying {len(safe_parts)} part(s) at WU Level 1...")
        except Exception:
            _reset_persistent_conn()
            raise
        try:
            cursor = conn.cursor()
            try:
                _tq = _time.perf_counter()
                if len(safe_parts) >= 200:
                    _log(f"[{_time.strftime('%H:%M:%S')}] Large set ({len(safe_parts)} parts) — using parallel shards...")
                    rows_l1, cols_l1 = _run_query_parallel_shards(
                        _SQL_LEVEL_N_TRAVERSE,
                        safe_parts,
                        safe_plant,
                    )
                else:
                    rows_l1, cols_l1 = _run_query_in_chunks(
                        cursor,
                        _SQL_LEVEL_N_TRAVERSE,
                        safe_parts,
                        safe_plant,
                    )
                _log(f"[{_time.strftime('%H:%M:%S')}] Query done in {_time.perf_counter() - _tq:.1f}s — {len(rows_l1)} raw rows returned.")

                rows_by_child: dict[str, list[dict[str, Any]]] = {}
                for row in rows_l1:
                    rec = {
                        col: ("" if val is None else str(val).strip())
                        for col, val in zip(cols_l1, row)
                    }
                    _set_designator(rec)
                    child_key = rec.get("child_part", "").upper()
                    rows_by_child.setdefault(child_key, []).append(rec)

                for child in safe_parts:
                    child_key = child.upper()
                    seen_parent_for_child: set[str] = set()
                    for rec in rows_by_child.get(child_key, []):
                        parent = rec.get("part", "")
                        parent_key = parent.upper()
                        if not _is_valid_part(parent):
                            continue
                        if not parent_key or parent_key in seen_parent_for_child:
                            continue
                        seen_parent_for_child.add(parent_key)

                        rec_out = dict(rec)
                        rec_out["wu_level"] = "1"
                        rec_out["ext_qty"] = str(rec.get("base_qty", ""))
                        rec_out["input_part"] = child
                        rec_out["_child_path_key"] = child_key
                        rec_out["_node_path_key"] = f"{child_key}->{parent_key}"
                        rec_out.setdefault("item_status", "")
                        rec_out.setdefault("procurement_type", "")
                        rec_out.setdefault("user_item_type", "")
                        rec_out.setdefault("pace_or_dash", "")
                        rec_out.setdefault("mlo_class", "")
                        rec_out.setdefault("designator", "")
                        rec_out.setdefault("option_class", "")
                        records.append(rec_out)

            finally:
                cursor.close()
        except Exception:
            _reset_persistent_conn()
            raise

    _log(f"[{_time.strftime('%H:%M:%S')}] Processing complete. {len(records)} total record(s) ready.")
    return _to_tree_order(records, safe_parts)
