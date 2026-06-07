from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any, Iterable, Mapping, Sequence

import pandas as pd


REPORT_COLUMNS = [
    "Part",
    "Part Description",
    "Plant",
    "Inventory Status",
    "Total Inventory",
    "Total Demand",
    "Excess Quantity",
    "Recommended Action",
    "Action Details",
    "Source Plant",
    "Destination Plant",
    "Transfer Quantity",
    "PO Number",
    "PO Qty",
    "Supplier Name",
    "Buyer Name",
    "Delivery Date",
    "Kit Code",
    "Kit Description",
]


@dataclass(frozen=True)
class InventorySnapshot:
    part: str
    part_description: str
    plant: str
    total_inventory: float
    total_demand: float
    inventory_status: str
    excess_quantity: float
    shortage_quantity: float
    on_hand_qty: float
    on_order_qty: float
    demand_13: float
    demand_26: float
    demand_52: float
    plant_inventory_cost: float
    total_inventory_cost: float
    ags_inventory: float
    ags_demand: float


def build_supply_chain_report(
    inventory_df: pd.DataFrame | None,
    where_used_records: Sequence[Mapping[str, Any]] | None = None,
    po_details: pd.DataFrame | Sequence[Mapping[str, Any]] | None = None,
    kit_descriptions: Mapping[str, str] | None = None,
    plants: Sequence[str | int] | None = None,
) -> pd.DataFrame:
    if inventory_df is None or inventory_df.empty:
        return pd.DataFrame(columns=REPORT_COLUMNS)

    kit_map = {
        str(key or "").strip().upper(): str(value or "").strip()
        for key, value in (kit_descriptions or {}).items()
        if str(key or "").strip()
    }
    part_to_kits = _extract_part_kit_codes(where_used_records or ())
    po_df = _normalize_po_details(po_details)
    snapshots = _build_inventory_snapshots(inventory_df, plants=plants)

    report_rows: list[dict[str, Any]] = []
    for part, part_snapshots in snapshots.items():
        # Aggregate across all plants for this part
        total_inventory = sum(s.total_inventory for s in part_snapshots)
        total_demand = sum(s.total_demand for s in part_snapshots)
        total_excess = max(total_inventory - total_demand, 0.0)
        total_cost = sum(s.total_inventory_cost for s in part_snapshots)
        part_description = part_snapshots[0].part_description if part_snapshots else ""

        # Only recommend if there is excess and cost
        if total_cost <= 0 or total_excess <= 0:
            continue

        kits_for_part = part_to_kits.get(part, [])
        outsourced_kits = [
            (kit_code, kit_map.get(kit_code, ""))
            for kit_code in kits_for_part
            if "OUTSOURCED" in kit_map.get(kit_code, "").upper()
        ]

        # For each outsourced kit, recommend sell back for the total excess qty
        for kit_code, kit_description in outsourced_kits:
            report_rows.append({
                "Part": part,
                "Part Description": part_description,
                "Plant": "Total",
                "Inventory Status": "Bad Cholesterol" if total_excess > 0 else "Balanced",
                "Total Inventory": round(total_inventory, 3),
                "Total Demand": round(total_demand, 3),
                "Excess Quantity": round(total_excess, 3),
                "Recommended Action": "Sell back to CM",
                "Action Details": "Sell back excess inventory to Contract Manufacturer (CM).",
                "Source Plant": "",
                "Destination Plant": "",
                "Transfer Quantity": "",
                "PO Number": "",
                "PO Qty": "",
                "Supplier Name": "",
                "Buyer Name": "",
                "Delivery Date": "",
                "Kit Code": kit_code,
                "Kit Description": kit_description,
            })

        # If no outsourced kits, still add a summary row for the part
        if not outsourced_kits:
            report_rows.append({
                "Part": part,
                "Part Description": part_description,
                "Plant": "Total",
                "Inventory Status": "Bad Cholesterol" if total_excess > 0 else "Balanced",
                "Total Inventory": round(total_inventory, 3),
                "Total Demand": round(total_demand, 3),
                "Excess Quantity": round(total_excess, 3),
                "Recommended Action": "Review excess inventory",
                "Action Details": "Excess inventory detected with no outsourced action.",
                "Source Plant": "",
                "Destination Plant": "",
                "Transfer Quantity": "",
                "PO Number": "",
                "PO Qty": "",
                "Supplier Name": "",
                "Buyer Name": "",
                "Delivery Date": "",
                "Kit Code": "",
                "Kit Description": "",
            })

    if not report_rows:
        return pd.DataFrame(columns=REPORT_COLUMNS)

    return pd.DataFrame(report_rows, columns=REPORT_COLUMNS)


def _build_inventory_snapshots(
    inventory_df: pd.DataFrame,
    plants: Sequence[str | int] | None = None,
) -> dict[str, list[InventorySnapshot]]:
    plant_codes = [str(plant).strip() for plant in (plants or _infer_plants(inventory_df))]
    snapshots: dict[str, list[InventorySnapshot]] = {}

    for _, row in inventory_df.iterrows():
        total_inventory_cost = _safe_float(row.get("Inventory Cost"))
        if total_inventory_cost <= 0:
            continue

        part = str(row.get("Material Number", "")).strip().upper()
        if not part:
            continue

        part_description = str(row.get("Material Description", "")).strip()
        for plant in plant_codes:
            on_hand = _safe_float(row.get(f"{plant} Onhand Qty"))
            on_order = _safe_float(row.get(f"{plant} On Order Qty"))
            demand_13 = _safe_float(row.get(f"{plant} Gross Demand-13"))
            demand_26 = _safe_float(row.get(f"{plant} Gross Demand-26"))
            demand_52 = _safe_float(row.get(f"{plant} Gross Demand-52"))
            total_inventory = on_hand + on_order
            total_demand = demand_13 + demand_26 + demand_52
            status = (
                "Good Cholesterol"
                if total_demand > total_inventory
                else "Bad Cholesterol"
                if total_demand < total_inventory
                else "Balanced"
            )
            excess_quantity = max(total_inventory - total_demand, 0.0)
            shortage_quantity = max(total_demand - total_inventory, 0.0)
            plant_inventory_cost = total_inventory * _safe_float(row.get(f"{plant} Standard Cost USD"))
            ags_inventory = _safe_float(row.get(f"{plant} AGS On-Hand")) + _safe_float(
                row.get(f"{plant} AGS On-Order")
            )
            ags_demand = _safe_float(row.get(f"{plant} AGS 6M Gross Demand"))

            snapshots.setdefault(part, []).append(
                InventorySnapshot(
                    part=part,
                    part_description=part_description,
                    plant=plant,
                    total_inventory=total_inventory,
                    total_demand=total_demand,
                    inventory_status=status,
                    excess_quantity=excess_quantity,
                    shortage_quantity=shortage_quantity,
                    on_hand_qty=on_hand,
                    on_order_qty=on_order,
                    demand_13=demand_13,
                    demand_26=demand_26,
                    demand_52=demand_52,
                    plant_inventory_cost=plant_inventory_cost,
                    total_inventory_cost=total_inventory_cost,
                    ags_inventory=ags_inventory,
                    ags_demand=ags_demand,
                )
            )

    return snapshots


def _build_transfer_targets(
    part_snapshots: Sequence[InventorySnapshot],
) -> dict[str, list[dict[str, Any]]]:
    targets: dict[str, list[dict[str, Any]]] = {}
    for snapshot in part_snapshots:
        if snapshot.total_demand > 0 and snapshot.total_inventory == 0:
            targets.setdefault(snapshot.part, []).append(
                {
                    "source_plant": snapshot.plant,
                    "destination": snapshot.plant,
                    "shortage": snapshot.total_demand,
                }
            )
        if snapshot.ags_demand > 0 and snapshot.ags_inventory == 0:
            targets.setdefault(snapshot.part, []).append(
                {
                    "source_plant": snapshot.plant,
                    "destination": "AGS",
                    "shortage": snapshot.ags_demand,
                }
            )
    return targets


def _extract_part_kit_codes(
    where_used_records: Sequence[Mapping[str, Any]],
) -> dict[str, list[str]]:
    part_to_kits: dict[str, list[str]] = {}
    current_part = ""
    current_kits: list[str] = []

    def flush() -> None:
        if not current_part:
            return
        deduped: list[str] = []
        seen: set[str] = set()
        for code in current_kits:
            if code not in seen:
                seen.add(code)
                deduped.append(code)
        part_to_kits[current_part] = deduped

    for record in where_used_records:
        level = _safe_int(record.get("wu_level"), default=-1)
        part = str(record.get("part", "")).strip().upper()
        kit_code = str(record.get("kit_code", "")).strip().upper()
        if level == 0:
            flush()
            current_part = part
            current_kits = []
        if current_part and kit_code:
            current_kits.append(kit_code)
    flush()
    return part_to_kits


def _normalize_po_details(
    po_details: pd.DataFrame | Sequence[Mapping[str, Any]] | None,
) -> pd.DataFrame:
    if po_details is None:
        return pd.DataFrame(
            columns=[
                "part_number",
                "plant",
                "po_number",
                "open_qty",
                "supplier_name",
                "buyer_name",
                "delivery_date",
            ]
        )
    if isinstance(po_details, pd.DataFrame):
        df = po_details.copy()
    else:
        df = pd.DataFrame(list(po_details))

    rename_map = {
        "material_number": "part_number",
        "plant_code": "plant",
        "po_qty": "open_qty",
        "quantity": "open_qty",
        "supplier": "supplier_name",
        "buyer": "buyer_name",
        "expected_delivery_date": "delivery_date",
    }
    df = df.rename(columns={key: value for key, value in rename_map.items() if key in df.columns})
    for column in [
        "part_number",
        "plant",
        "po_number",
        "open_qty",
        "supplier_name",
        "buyer_name",
        "delivery_date",
    ]:
        if column not in df.columns:
            df[column] = ""

    df["part_number"] = df["part_number"].astype(str).str.strip().str.upper()
    df["plant"] = df["plant"].astype(str).str.strip()
    df["open_qty"] = df["open_qty"].apply(_safe_float)
    return df


def _select_latest_po(source_po_rows: pd.DataFrame) -> dict[str, Any] | None:
    if source_po_rows.empty:
        return None
    ranked = source_po_rows.copy()
    ranked["_delivery_sort"] = ranked["delivery_date"].apply(_parse_date)
    ranked = ranked.sort_values(by=["_delivery_sort", "po_number"], ascending=[False, False])
    return ranked.iloc[0].drop(labels=["_delivery_sort"]).to_dict()


def _build_report_row(
    source: InventorySnapshot,
    recommended_action: str,
    action_details: str,
    source_plant: str,
    destination_plant: str = "",
    transfer_quantity: float = 0.0,
    po_row: Mapping[str, Any] | None = None,
    kit_code: str = "",
    kit_description: str = "",
) -> dict[str, Any]:
    po_row = po_row or {}
    delivery_date = po_row.get("delivery_date", "")
    if isinstance(delivery_date, datetime):
        delivery_date = delivery_date.strftime("%Y-%m-%d")
    return {
        "Part": source.part,
        "Part Description": source.part_description,
        "Plant": source.plant,
        "Inventory Status": source.inventory_status,
        "Total Inventory": round(source.total_inventory, 3),
        "Total Demand": round(source.total_demand, 3),
        "Excess Quantity": round(source.excess_quantity, 3),
        "Recommended Action": recommended_action,
        "Action Details": action_details,
        "Source Plant": source_plant,
        "Destination Plant": destination_plant,
        "Transfer Quantity": round(transfer_quantity, 3) if transfer_quantity else "",
        "PO Number": po_row.get("po_number", ""),
        "PO Qty": round(_safe_float(po_row.get("open_qty")), 3) if po_row else "",
        "Supplier Name": po_row.get("supplier_name", ""),
        "Buyer Name": po_row.get("buyer_name", ""),
        "Delivery Date": delivery_date,
        "Kit Code": kit_code,
        "Kit Description": kit_description,
    }


def _build_cancel_action_details(cancel_qty: float, po_row: Mapping[str, Any] | None) -> str:
    if not po_row:
        return f"Cancel open PO quantity of {cancel_qty:,.3f} based on excess inventory."
    po_number = po_row.get("po_number", "")
    supplier_name = po_row.get("supplier_name", "")
    buyer_name = po_row.get("buyer_name", "")
    delivery_date = po_row.get("delivery_date", "")
    return (
        f"Cancel {cancel_qty:,.3f} on PO {po_number} with supplier {supplier_name} "
        f"and buyer {buyer_name}; latest delivery date {delivery_date}."
    ).strip()


def _infer_plants(inventory_df: pd.DataFrame) -> list[str]:
    plants: list[str] = []
    seen: set[str] = set()
    for column in inventory_df.columns:
        text = str(column)
        if not text.endswith(" Onhand Qty"):
            continue
        plant = text.split(" ", 1)[0].strip()
        if plant and plant not in seen:
            seen.add(plant)
            plants.append(plant)
    return plants


def _parse_date(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value
    if hasattr(value, "to_pydatetime"):
        return value.to_pydatetime()
    text = str(value or "").strip()
    if not text:
        return datetime.min
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return datetime.min
    return parsed.to_pydatetime()


def _safe_float(value: Any) -> float:
    try:
        if value is None or value == "":
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _safe_int(value: Any, default: int = 0) -> int:
    try:
        if value is None or value == "":
            return default
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default