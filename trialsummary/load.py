"""Load the Citrus CSV exports, the LabConnect inventory report and the site lookup."""
from __future__ import annotations

import csv
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl

CITRUS_FILES = {
    "participants": "PARTICIPANTS_OVERVIEW",
    "crf": "DATA_COMPLETENESS__CRF",
    "epro": "DATA_COMPLETENESS__EPRO",
    "queries": "DATA_QUERY",
    "psf": "PostSurgical_Findings",
}


def ts(value: str) -> datetime | None:
    """Parse a Citrus timestamp such as '2024-08-07 10:11:00+00:00'. 'NaT' and '' become None."""
    value = (value or "").strip()
    if value in ("", "NaT", "nan"):
        return None
    return datetime.fromisoformat(value)


def lab_date(value: str) -> date | None:
    """Parse a LabConnect date such as '22 Nov 2024'."""
    value = (value or "").strip()
    return datetime.strptime(value, "%d %b %Y").date() if value else None


def read_csv(path: Path) -> list[dict]:
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def find_one(folder: Path, pattern: str) -> Path:
    hits = sorted(folder.glob(pattern))
    if not hits:
        raise FileNotFoundError(f"No file matching {pattern!r} in {folder}")
    return hits[-1]  # latest export if several


def export_time(path: Path) -> datetime:
    m = re.search(r"(\d{4}-\d{2}-\d{2})-(\d{2})-(\d{2})-UTC", path.name)
    return datetime.strptime(f"{m.group(1)} {m.group(2)}:{m.group(3)}", "%Y-%m-%d %H:%M")


DROPPED_INVENTORY_COLUMNS = ("Date of Birth",)  # personal data the summary never needs


def read_inventory(path: Path) -> tuple[list[dict], str]:
    """Rows from the inventory report. Row 3 holds the report date, row 4 the header.
    Date of birth is dropped on read and never reaches the metrics or outputs."""
    ws = openpyxl.load_workbook(path, read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    report_date = str(rows[2][0] or "").replace("Date:", "").strip()
    header = [str(h) for h in rows[3]]
    data = [{h: ("" if v is None else str(v)) for h, v in zip(header, r) if h not in DROPPED_INVENTORY_COLUMNS}
            for r in rows[4:] if any(r)]
    return data, report_date


def load_inputs(inputs: Path, site_lookup: Path) -> dict:
    data, sources = {}, {}
    for key, prefix in CITRUS_FILES.items():
        path = find_one(inputs, f"{prefix}-CleoUSA-*.csv")
        data[key] = read_csv(path)
        sources[key] = path
    inv_path = find_one(inputs, "*InventoryReport_*.xlsx")
    data["inventory"], inventory_date = read_inventory(inv_path)
    sources["inventory"] = inv_path
    data["sites"] = read_csv(site_lookup)
    sources["sites"] = site_lookup
    data["meta"] = {
        "citrus_export": export_time(sources["participants"]),
        "inventory_date": inventory_date,
        "sources": sources,
    }
    return data
