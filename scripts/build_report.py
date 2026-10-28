#!/usr/bin/env python3
"""Build the CleoUSA monthly trial summary from Citrus exports and the central lab inventory report.

Outputs (in --out):
  CleoUSA_Monthly_Summary_<Mon><YYYY>_Sponsor.docx/.xlsx    sponsor version: no data completeness, no participant data
  CleoUSA_Monthly_Summary_<Mon><YYYY>_Internal.docx/.xlsx   internal version: every section and the participant lists
  CleoUSA_Monthly_Totals_<Mon><YYYY>.csv     totals only, no participant-level rows
  charts/                                     chart images used in the Word file

Usage:
  .venv/bin/python scripts/build_report.py --month 2026-07 [--audience both|sponsor|internal]
"""
import argparse
import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from trialsummary.excel import write_workbook  # noqa: E402
from trialsummary.load import load_inputs  # noqa: E402
from trialsummary.metrics import build  # noqa: E402
from trialsummary.word import write_document  # noqa: E402


def write_totals(path, sections):
    """Aggregate tables only. Detail tables (participant or sample lists) are left out."""
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["section", "table", "row", "column", "value"])
        for s in sections.values():
            for t in s.tables:
                if t.detail:
                    continue
                for r in t.rows:
                    for col, v in zip(t.cols[1:], r[1:]):
                        if isinstance(v, tuple):
                            v = f"{v[0]:.1f}-{v[1]:.1f}"
                        elif isinstance(v, float):
                            v = round(v, 4)
                        w.writerow([s.title, t.title, r[0], col.header, "" if v is None else v])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--month", default="2026-07", help="Reporting month as YYYY-MM")
    ap.add_argument("--inputs", type=Path, default=ROOT / "data" / "dummy_inputs")
    ap.add_argument("--sites", type=Path, default=ROOT / "data" / "reference" / "site_lookup.csv")
    ap.add_argument("--out", type=Path, default=None)
    ap.add_argument("--audience", choices=["both", "sponsor", "internal"], default="both")
    a = ap.parse_args()

    data = load_inputs(a.inputs, a.sites)
    period, sections, head = build(data, a.month)
    tag = period.start.strftime("%b%Y")
    out = a.out or ROOT / "outputs" / a.month
    out.mkdir(parents=True, exist_ok=True)

    written = []
    for audience in (["sponsor", "internal"] if a.audience == "both" else [a.audience]):
        name = f"CleoUSA_Monthly_Summary_{tag}_{audience.capitalize()}"
        written.append(write_workbook(out / f"{name}.xlsx", period, sections, head, data["meta"], audience))
        written.append(write_document(out / f"{name}.docx", period, sections, head, data["meta"], out / "charts", audience))
    totals = out / f"CleoUSA_Monthly_Totals_{tag}.csv"
    write_totals(totals, sections)
    for p in written + [totals]:
        print(p.relative_to(ROOT))


if __name__ == "__main__":
    main()
