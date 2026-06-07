"""Watch List logic and UI tab.

Rules:
1) Include part when (part in OBS list AND part is PACE using where_used_query data)
2) OR include part when Cholesterol == Bad Cholesterol.

This module is standalone and does not modify where_used_query.py.
"""

from __future__ import annotations

import re
from typing import Any

import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QBrush, QColor
from PyQt6.QtWidgets import QHeaderView
from PyQt6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


REASON_COL = "Reason to add in Watch List"


def _norm_part(value: Any) -> str:
    return str(value or "").strip().upper()


def read_inventory_data(main_window: Any) -> pd.DataFrame:
    """Read Inventory_Cost dataframe from main window and ensure Cholesterol is present."""
    inv_tab = getattr(main_window, "inventory_cost_tab", None)
    src = getattr(inv_tab, "df", None)
    if src is None:
        return pd.DataFrame()
    if not isinstance(src, pd.DataFrame) or src.empty:
        return pd.DataFrame()

    out = src.copy()
    if "Cholesterol" not in out.columns:
        out["Cholesterol"] = out.apply(_compute_cholesterol, axis=1)
    return out


def read_obs_parts(obs_table: QTableWidget, obs_col: int = 1) -> set[str]:
    """Read OBS part numbers from OBS tab table column 1."""
    parts: set[str] = set()
    if obs_table is None:
        return parts
    for r in range(obs_table.rowCount()):
        it = obs_table.item(r, obs_col)
        key = _norm_part(it.text() if it else "")
        if key:
            parts.add(key)
    return parts


def evaluate_pace(parts: set[str], plant: str = "4070") -> dict[str, bool]:
    """Evaluate PACE using where_used_query output (pace_or_dash column on level 0 rows)."""
    pace_map = {p: False for p in parts}
    if not parts:
        return pace_map

    try:
        from where_used_query import fetch_where_used  # type: ignore[import]
    except Exception:
        return pace_map

    try:
        records = fetch_where_used(sorted(parts), 1, plant=str(plant or "4070"))
    except Exception:
        return pace_map

    for rec in records:
        if str(rec.get("wu_level", "")).strip() != "0":
            continue
        part_key = _norm_part(rec.get("part") or rec.get("input_part"))
        if part_key not in pace_map:
            continue
        pace_text = str(rec.get("pace_or_dash", "") or "").strip().upper()
        if "PACE" in pace_text:
            pace_map[part_key] = True
    return pace_map


def apply_watch_list_filter(
    inventory_df: pd.DataFrame,
    obs_parts: set[str],
    pace_map: dict[str, bool],
    part_col: str = "Material Number",
    cholesterol_col: str = "Cholesterol",
) -> pd.DataFrame:
    """Apply watch-list inclusion rules and return deduplicated output rows."""
    if inventory_df is None or inventory_df.empty:
        return pd.DataFrame()

    rows = []
    seen: set[str] = set()

    for _, row in inventory_df.iterrows():
        part_key = _norm_part(row.get(part_col, ""))
        if not part_key or part_key in seen:
            continue

        is_bad_chol = str(row.get(cholesterol_col, "")).strip().lower() == "bad cholesterol"
        in_obs_and_pace = (part_key in obs_parts) and bool(pace_map.get(part_key, False))

        if in_obs_and_pace or is_bad_chol:
            out_row = row.copy()
            out_row[REASON_COL] = (
                "OBS + PACE" if in_obs_and_pace and not is_bad_chol else
                "Bad Cholesterol" if is_bad_chol and not in_obs_and_pace else
                "OBS + PACE; Bad Cholesterol"
            )
            rows.append(out_row)
            seen.add(part_key)

    if not rows:
        return pd.DataFrame(columns=list(inventory_df.columns) + [REASON_COL])
    return pd.DataFrame(rows)


def _write_watch_list_output_plain(table: QTableWidget, output_df: pd.DataFrame) -> None:
    """Write dataframe rows to watch-list table (plain fallback mode)."""
    table.clear()
    if output_df is None or output_df.empty:
        table.setRowCount(0)
        table.setColumnCount(0)
        return

    headers = [str(c) for c in output_df.columns.tolist()]
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(output_df))

    for r in range(len(output_df)):
        row = output_df.iloc[r]
        for c, col in enumerate(headers):
            val = row.get(col, "")
            txt = "" if pd.isna(val) else str(val)
            item = QTableWidgetItem(txt)
            if col == REASON_COL:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(r, c, item)

    table.resizeColumnsToContents()


def write_watch_list_output(table: QTableWidget, output_df: pd.DataFrame, main_window: Any | None = None) -> None:
    """Write watch-list rows using Inventory_Cost-like headers/colors/orientation."""
    if output_df is None or output_df.empty:
        _write_watch_list_output_plain(table, output_df)
        return

    inv_tab = getattr(main_window, "inventory_cost_tab", None) if main_window is not None else None
    inv_table = getattr(inv_tab, "table", None) if inv_tab is not None else None

    # If Inventory_Cost table context is unavailable, fall back safely.
    if inv_table is None or not isinstance(inv_table, QTableWidget):
        _write_watch_list_output_plain(table, output_df)
        return

    try:
        # Keep table look-and-feel aligned with Inventory_Cost.
        table.setStyleSheet(inv_table.styleSheet())
    except Exception:
        pass

    view_df = output_df.copy()
    headers = view_df.columns.tolist()

    summary_cols = [
        "Total On Order Quantity",
        "Total Onhand",
        "Gross Demand-13",
        "Gross Demand-26",
        "Gross Demand-52",
        "Inventory Cost",
    ]

    display_headers = []
    rotated_cols = []
    plant_groups = []
    header_texts = {}
    summary_indices = []
    watch_rule_indices = []

    for idx, name in enumerate(headers):
        h = str(name)
        m = re.match(r"^(\d{4})\s+(.+)$", h)
        if m:
            plant = m.group(1)
            metric = m.group(2)
            display_label = f"{metric} ({plant})"
            display_headers.append(display_label)
            header_texts[idx] = display_label
            rotated_cols.append(idx)
            if not plant_groups or plant_groups[-1][0] != plant:
                plant_groups.append((plant, [idx]))
            else:
                plant_groups[-1][1].append(idx)
        else:
            display_headers.append(h)
            header_texts[idx] = h
            if h.strip() in summary_cols:
                summary_indices.append(idx)
            if h.strip().lower() == REASON_COL.lower():
                watch_rule_indices.append(idx)

    rotated_cols.extend(i for i in summary_indices if i not in rotated_cols)

    # Put totals in their own group; keep watch rule as a separate group if present.
    if summary_indices:
        plant_groups.append(("Total", summary_indices))
    if watch_rule_indices:
        plant_groups.append(("Watch List", watch_rule_indices))

    table.clear()
    table.setRowCount(len(view_df))
    table.setColumnCount(len(view_df.columns))

    # Reuse the same custom header class used by Inventory_Cost (for rotated/grouped headers).
    try:
        inv_hdr = inv_table.horizontalHeader()
        hdr_cls = type(inv_hdr)
        hdr = hdr_cls(Qt.Orientation.Horizontal, rotated_columns=rotated_cols, parent=table)
        if hasattr(hdr, "set_header_texts"):
            hdr.set_header_texts(header_texts)
        if hasattr(hdr, "set_group_spans"):
            hdr.set_group_spans(plant_groups)
        table.setHorizontalHeader(hdr)
    except Exception:
        pass

    table.setHorizontalHeaderLabels(display_headers)

    # Try to preserve Inventory_Cost header height / orientation feel.
    try:
        table.horizontalHeader().setVisible(True)
        table.horizontalHeader().setFixedHeight(inv_table.horizontalHeader().height())
    except Exception:
        pass

    col_to_group_idx = {}
    for gi, (_gname, idxs) in enumerate(plant_groups):
        for ci in idxs:
            col_to_group_idx[ci] = gi

    group_colors = [
        QColor("#D6EAF8"), QColor("#D5F5E3"), QColor("#FDEBD0"),
        QColor("#E8DAEF"), QColor("#D6EAF8"), QColor("#D5F5E3"),
        QColor("#FDEBD0"), QColor("#E8DAEF"),
    ]
    total_color = QColor("#D9E1F2")
    watch_color = QColor("#FFF2CC")

    for r in range(len(view_df)):
        row_data = view_df.iloc[r]
        for c, col in enumerate(view_df.columns):
            v = row_data[col]
            txt = "" if pd.isna(v) else str(v)

            try:
                num_v = float(v)
            except Exception:
                num_v = None

            if num_v == 0:
                txt = ""
            if num_v is not None and "Cost" in str(col) and num_v != 0:
                txt = f"${num_v:,.2f}"

            item = QTableWidgetItem(txt)

            if c in col_to_group_idx:
                gi = col_to_group_idx[c]
                gname = plant_groups[gi][0]
                if gname == "Total":
                    item.setBackground(QBrush(total_color))
                elif gname == "Watch List":
                    if str(col).strip().lower() == REASON_COL.lower():
                        if txt == "Bad Cholesterol":
                            item.setBackground(QBrush(QColor(220, 50, 50)))
                            item.setForeground(QBrush(QColor(255, 255, 255)))
                        elif txt == "OBS + PACE":
                            item.setBackground(QBrush(QColor(50, 180, 50)))
                            item.setForeground(QBrush(QColor(255, 255, 255)))
                        else:
                            item.setBackground(QBrush(watch_color))
                    else:
                        item.setBackground(QBrush(watch_color))
                else:
                    item.setBackground(QBrush(group_colors[gi % len(group_colors)]))

            table.setItem(r, c, item)

    table.resizeColumnsToContents()

    for c in rotated_cols:
        try:
            table.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeMode.Fixed)
            col_name = str(view_df.columns[c])
            if col_name.strip() in ("Total On Order Quantity", "Total Onhand", "Gross Demand-13", "Gross Demand-26", "Gross Demand-52"):
                table.setColumnWidth(c, 36)
            elif col_name.strip() == "Inventory Cost":
                table.setColumnWidth(c, 60)
            elif "Cost" in col_name:
                table.setColumnWidth(c, 50)
            else:
                table.setColumnWidth(c, 30)
        except Exception:
            pass


def _compute_cholesterol(row: pd.Series) -> str:
    try:
        total_inv = float(row.get("Total Onhand", 0) or 0) + float(row.get("Total On Order Quantity", 0) or 0)
        total_dem = (
            float(row.get("Gross Demand-13", 0) or 0)
            + float(row.get("Gross Demand-26", 0) or 0)
            + float(row.get("Gross Demand-52", 0) or 0)
        )
        if total_inv <= 0:
            return ""
        return "Good Cholesterol" if (total_inv - total_dem) < 0 else "Bad Cholesterol"
    except Exception:
        return ""


class WatchListTab(QWidget):
    """Watch_List tab populated from Inventory_Cost + OBS + where_used PACE data."""

    def __init__(self, main_window: Any):
        super().__init__()
        self._main_window = main_window
        self.df = pd.DataFrame()

        outer = QVBoxLayout(self)
        row = QHBoxLayout()

        title = QLabel("Watch_List")
        row.addWidget(title)
        row.addStretch(1)

        self.btn_refresh = QPushButton("Refresh Watch List")
        self.btn_reset = QPushButton("Reset")
        row.addWidget(self.btn_refresh)
        row.addWidget(self.btn_reset)
        outer.addLayout(row)

        self.status = QLabel("Import Inventory_Cost from Databricks to populate Watch_List.")
        self.status.setWordWrap(True)
        outer.addWidget(self.status)

        self.table = QTableWidget(0, 0)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        outer.addWidget(self.table)

        self.btn_refresh.clicked.connect(lambda: self.refresh_from_sources(show_message=True))
        self.btn_reset.clicked.connect(self.reset_tab)

    def reset_tab(self) -> None:
        self.df = pd.DataFrame()
        write_watch_list_output(self.table, self.df, self._main_window)
        self.status.setText("Watch_List reset.")

    def refresh_from_sources(self, show_message: bool = False) -> None:
        """Rebuild watch-list rows from current Inventory_Cost + OBS data."""
        try:
            inventory_df = read_inventory_data(self._main_window)
            if inventory_df.empty:
                self.reset_tab()
                self.status.setText("No Inventory_Cost data available.")
                if show_message:
                    QMessageBox.information(self, "Watch_List", "No Inventory_Cost data available.")
                return

            obs_tab = getattr(self._main_window, "obs_tab", None)
            obs_table = getattr(obs_tab, "table", None)
            obs_parts = read_obs_parts(obs_table)

            candidate_parts = {
                _norm_part(v)
                for v in inventory_df.get("Material Number", pd.Series(dtype=str)).tolist()
                if _norm_part(v) in obs_parts
            }

            plant = "4070"
            where_tab = getattr(self._main_window, "whereused_tab", None)
            if where_tab is not None and hasattr(where_tab, "cmb_plant"):
                try:
                    plant = str(where_tab.cmb_plant.currentText() or "4070").strip()
                except Exception:
                    plant = "4070"

            pace_map = evaluate_pace(candidate_parts, plant=plant)

            self.df = apply_watch_list_filter(
                inventory_df=inventory_df,
                obs_parts=obs_parts,
                pace_map=pace_map,
            )
            write_watch_list_output(self.table, self.df, self._main_window)

            self.status.setText(
                f"Watch_List loaded: {len(self.df)} part(s). "
                f"OBS candidates checked for PACE: {len(candidate_parts)}."
            )
            if show_message:
                QMessageBox.information(self, "Watch_List", f"Loaded {len(self.df)} Watch_List part(s).")
        except Exception as exc:
            self.status.setText(f"Watch_List refresh failed: {exc}")
            if show_message:
                QMessageBox.warning(self, "Watch_List", f"Failed to refresh Watch_List: {exc}")


