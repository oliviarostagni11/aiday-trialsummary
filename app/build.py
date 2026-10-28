#!/usr/bin/env python3
"""Assemble the Trial Summary Builder artifact page from app/src.

Inlines engine.js, model.js, writers.js and ui.js, embeds the site list, and stages the example files.
Output: app/dist/trial-summary-builder.html and app/dist/examples/*
Usage:  python3 app/build.py
"""
import base64
import csv
import re
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC, DIST = ROOT / "app" / "src", ROOT / "app" / "dist"
EXAMPLE_SRC = ROOT / "data" / "dummy_inputs"
SITE_COLS = ["site_id", "lab_site_code", "site", "site_type", "academic_status", "activation_date", "terminated", "termination_date"]


def main():
    DIST.mkdir(parents=True, exist_ok=True)
    ex_dir = DIST / "examples"
    shutil.rmtree(ex_dir, ignore_errors=True)
    ex_dir.mkdir()
    examples = sorted(p for p in EXAMPLE_SRC.glob("*") if p.suffix == ".csv" and "InventoryReport" not in p.name)
    for p in examples:
        shutil.copy2(p, ex_dir / p.name)
    # Artifacts cannot serve .xlsx, so the example inventory report ships as base64 inside a JSON file.
    for x in sorted(EXAMPLE_SRC.glob("*InventoryReport*.xlsx")):
        wrapped = ex_dir / "inventory-report.b64.json"
        wrapped.write_text(json.dumps({"name": x.name, "base64": base64.b64encode(x.read_bytes()).decode()}))
        examples.append(wrapped)

    with open(ROOT / "data" / "reference" / "site_lookup.csv", encoding="utf-8-sig") as f:
        sites = [{k: r[k] for k in SITE_COLS} for r in csv.DictReader(f)]

    page = (SRC / "page.html").read_text()
    data_js = ("window.TS_SITES = " + json.dumps(sites, ensure_ascii=False) + ";\n" +
               "window.TS_EXAMPLES = " + json.dumps(["examples/" + p.name for p in examples]) + ";")
    # Keep the page pure ASCII so it renders the same with or without a declared charset.
    page = re.sub(r"[^\x00-\x7f]", lambda m: f"&#x{ord(m.group()):04x};", page)
    data_js = re.sub(r"[^\x00-\x7f]", lambda m: f"\\u{ord(m.group()):04x}", data_js)
    for name in ("engine.js", "model.js", "writers.js", "ui.js"):
        code = re.sub(r"[^\x00-\x7f]", lambda m: f"\\u{ord(m.group()):04x}", (SRC / name).read_text())
        assert "</script" not in code.lower(), f"{name} contains a closing script tag"
        page = page.replace(f"<!-- INLINE:{name} -->", f"<script>\n{code}\n</script>")
    page = page.replace("<!-- INLINE:sites -->", f"<script>\n{data_js}\n</script>")
    assert "<!-- INLINE:" not in page
    out = DIST / "trial-summary-builder.html"
    out.write_text(page)
    print(out.relative_to(ROOT), f"{out.stat().st_size / 1024:.0f} KB;", len(examples), "example files,",
          f"{sum(p.stat().st_size for p in examples) / 1024 / 1024:.1f} MB")


if __name__ == "__main__":
    main()
