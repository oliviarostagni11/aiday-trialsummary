"""Story charts for the Word and Excel outputs.

Two layers, mirrored line for line in app/src/writers.js so both tools draw identical charts:
  chart_specs()  the data, colours and takeaway titles for each chart (compared across tools by a test)
  render()       draws a spec in pixel coordinates on a 1400 px wide canvas, then saves a PNG

Colour rules follow the data-viz reference palette, validated on a white surface:
  emphasis    one accent bar (#2a78d6), context bars in the baseline gray (#c3c2b7)
  categorical fixed slot order #2a78d6 #eb6834 #1baf7a #eda100 #e87ba4 (legend always carries the values)
  ordinal     one-hue ramp #86b6ef #3987e5 #1c5cab #104281 for ordered day bands
"""
from __future__ import annotations

import math
from datetime import date, timedelta
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import FancyBboxPatch, Rectangle  # noqa: E402

W, PADX = 1400, 16
INK, INK_2, GRID, AXIS, CONTEXT = "#0b0b0b", "#52514e", "#e1e0d9", "#c3c2b7", "#c3c2b7"
ACCENT = "#2a78d6"
CATEGORICAL = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"]
ORDINAL = ["#86b6ef", "#3987e5", "#1c5cab", "#104281"]
BAR, BAR_H, RADIUS, GAP = 48, 32, 8, 4  # image px; the image shows at about half size in Word
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ------------------------------------------------------------------ shared text helpers
def pct_text(x):
    if x is None:
        return "–"
    if 0 < x < 0.005:
        return "<1%"
    return f"{x * 100:.0f}%"


def prev_month_label(period) -> str:
    return (period.start - timedelta(days=1)).strftime("%B %Y")


def wrap(text: str, limit: int = 70) -> list[str]:
    """Split a title into lines of at most `limit` characters at word boundaries (same rule in JS)."""
    lines, cur = [], ""
    for word in text.split(" "):
        if cur and len(cur) + 1 + len(word) > limit:
            lines.append(cur)
            cur = word
        else:
            cur = f"{cur} {word}" if cur else word
    if cur:
        lines.append(cur)
    return lines


def month_labels(labels: list[str]) -> tuple[list[str], list[str]]:
    """'Aug 2025' -> line 1 'Aug'; line 2 shows the year on the first column and on January."""
    l1, l2 = [], []
    for i, lab in enumerate(labels):
        mon, yr = lab.split(" ")
        l1.append(mon)
        l2.append(f"'{yr[2:]}" if i == 0 or mon == "Jan" else "")
    return l1, l2


# ------------------------------------------------------------------ chart specs (data + takeaway titles)
def chart_specs(sections: dict, period, head: dict) -> dict:
    T = {t.key: t for s in sections.values() for t in s.tables}
    m, specs = period.label, {}

    # 1. Monthly enrollment, emphasis on the reporting month
    rows = T["enrol_trend"].rows[-12:]
    vals = [r[2] for r in rows]
    n, prev = vals[-1], (vals[-2] if len(vals) >= 2 else None)
    trend = ("" if prev is None else f", up from {prev} in {prev_month_label(period)}" if n > prev else
             f", down from {prev} in {prev_month_label(period)}" if n < prev else f", the same as in {prev_month_label(period)}")
    l1, l2 = month_labels([r[0] for r in rows])
    specs["enrollment_trend"] = {
        "key": "enrollment_trend", "kind": "columns", "after": "enrol_trend", "replaces_table": True,
        "title": f"{n} participants enrolled in {m}{trend}",
        "subtitle": "Participants enrolled per month, last 12 months",
        "labels": l1, "labels2": l2, "values": vals,
        "colors": [CONTEXT] * (len(vals) - 1) + [ACCENT], "label_idx": [len(vals) - 1],
        "value_text": {str(len(vals) - 1): str(n)},
    }

    # 2. Enrollment this month by site, ranked
    t = T["enrol_by_site_week"]
    mcol = next(i for i, c in enumerate(t.cols) if c.header == f"{period.short} total")
    site_rows = [r for i, r in enumerate(t.rows) if i not in t.total_rows and r[mcol] > 0]
    site_rows.sort(key=lambda r: (-r[mcol], r[0]))
    n_active = head["active_sites"]
    if site_rows:
        top = site_rows[0]
        title = f"{len(site_rows)} of {n_active} active sites enrolled participants in {m}; {top[0].replace(' (terminated)', '')} led with {top[mcol]}"
    else:
        title = f"No participants enrolled in {m}"
    specs["enrollment_by_site"] = {
        "key": "enrollment_by_site", "kind": "hbars", "before": "enrol_by_site_week",
        "title": title, "subtitle": f"Participants enrolled in {m}, by site",
        "labels": [r[0].replace(" (terminated)", "") for r in site_rows], "values": [r[mcol] for r in site_rows],
        "colors": [ACCENT] * len(site_rows),
    }

    # 3. Participant status at month end, part to whole
    order = ["Completed", "Enrolled, active", "Consented, not yet enrolled", "Withdrawn", "Screen failed"]
    by_status = {r[0]: r for i, r in enumerate(T["status_eom"].rows) if i not in T["status_eom"].total_rows}
    st = [by_status[k] for k in order if k in by_status]
    consented = sum(r[1] for r in st)
    completed = next(r[1] for r in st if r[0] == "Completed")
    specs["participant_status"] = {
        "key": "participant_status", "kind": "stack100", "after": "status_eom",
        "title": f"{completed:,} of {consented:,} consented participants have completed the study",
        "subtitle": f"Where each consented participant stands at month end ({period.last_day.day} {period.last_day:%b %Y})",
        "labels": [r[0] for r in st], "values": [r[1] for r in st], "colors": CATEGORICAL[:len(st)],
        "legend_text": [f"{r[0]} · {r[1]:,} ({pct_text(r[2])})" for r in st],
    }

    # 4. Pathology reports still awaited, by days since surgery
    days = [r[-1] for r in T["psf_awaiting"].rows]
    bands = [("0–30 days", 0, 30), ("31–60 days", 31, 60), ("61–90 days", 61, 90), ("Over 90 days", 91, 10 ** 9)]
    counts = [sum(lo <= d <= hi for d in days) for _, lo, hi in bands]
    over60 = counts[2] + counts[3]
    specs["pathology_backlog"] = {
        "key": "pathology_backlog", "kind": "columns", "after": "psf_summary",
        "title": (f"{over60} of {len(days)} reports still awaited are more than 60 days past surgery" if days
                  else "No pathology reports are awaited"),
        "subtitle": "Participants with surgery complete and no pathology report, by days since surgery",
        "labels": [b[0] for b in bands], "labels2": ["", "", "", ""], "values": counts, "colors": ORDINAL,
        "label_idx": [0, 1, 2, 3], "value_text": {str(i): str(c) for i, c in enumerate(counts)},
    }

    # 5. Sample registration timeliness by month
    tr = T["blood_trend"].rows
    l1, l2 = month_labels([r[0] for r in tr])
    cur, prv = tr[-1], (tr[-2] if len(tr) >= 2 else None)
    comp = f", compared with {pct_text(prv[4])} in {prev_month_label(period)}" if prv and prv[1] else ""
    specs["sample_timeliness"] = {
        "key": "sample_timeliness", "kind": "stacked_columns", "after": "blood_trend", "replaces_table": True,
        "title": (f"{pct_text(cur[4])} of samples registered in {m} met the registration window{comp}" if cur[1]
                  else f"No samples were registered in {m}"),
        "subtitle": "Samples registered per month at the central lab",
        "labels": l1, "labels2": l2,
        "series": [{"name": "Within window", "color": CATEGORICAL[0], "values": [r[2] for r in tr]},
                   {"name": "Outside window", "color": CATEGORICAL[1], "values": [r[3] for r in tr]}],
        "label_idx": [len(tr) - 1], "value_text": {str(len(tr) - 1): pct_text(cur[4])},
    }

    # 6. Unresolved queries by site
    q = [r for i, r in enumerate(T["query_by_site"].rows) if i not in T["query_by_site"].total_rows and r[2] + r[4] > 0]
    q.sort(key=lambda r: (-(r[2] + r[4]), r[0]))
    total = sum(r[2] + r[4] for r in q)
    lead = q[0] if q else None
    subtitle = "Open and reviewing queries by site"
    if len(q) > 10:  # keep the chart short and the scale honest: top 10 sites, the rest named in the subtitle
        rest = q[10:]
        q = q[:10]
        subtitle = (f"Open and reviewing queries, the 10 sites with the most; "
                    f"{len(rest)} other sites have {sum(r[2] + r[4] for r in rest):,} more")
    specs["queries_by_site"] = {
        "key": "queries_by_site", "kind": "stacked_hbars", "after": "query_summary",
        "title": (f"{total:,} queries were unresolved at month end; {lead[0]} had the most ({lead[2] + lead[4]})" if q
                  else "No queries were unresolved at month end"),
        "subtitle": subtitle,
        "labels": [r[0] for r in q],
        "series": [{"name": "Open", "color": CATEGORICAL[0], "values": [r[2] for r in q]},
                   {"name": "Reviewing", "color": CATEGORICAL[1], "values": [r[4] for r in q]}],
    }
    return specs


# ------------------------------------------------------------------ pixel renderer
def nice_max(v: float, ticks: int = 4) -> float:
    """Smallest 'nice' axis maximum >= v that divides into `ticks` clean steps (same rule in JS)."""
    if v <= 0:
        return ticks
    raw = v / ticks
    mag = 10 ** math.floor(math.log10(raw))
    for f in (1, 2, 2.5, 5, 10):
        if f * mag >= raw:
            return f * mag * ticks
    return 10 * mag * ticks


def fit_label(s: str, limit: int = 38) -> str:
    return s if len(s) <= limit else s[: limit - 1] + "…"


class Canvas:
    """Minimal drawing surface in image pixels, y pointing down. Font sizes are in pixels."""

    def __init__(self, h: int):
        self.h = h
        self.fig = plt.figure(figsize=(W / 200, h / 200), dpi=200)
        self.ax = self.fig.add_axes([0, 0, 1, 1])
        self.ax.set_xlim(0, W)
        self.ax.set_ylim(h, 0)
        self.ax.axis("off")
        self.fig.patch.set_facecolor("white")

    def rect(self, x, y, w, h, color):
        if w > 0 and h > 0:
            self.ax.add_patch(Rectangle((x, y), w, h, facecolor=color, edgecolor="none", linewidth=0))

    def bar(self, x, y, w, h, color, end):
        """Rounded data end ('top' or 'right'), square at the baseline."""
        long_side = h if end == "top" else w
        if w <= 0 or h <= 0:
            return
        if end == "none" or long_side <= 2 * RADIUS or min(w, h) < 2 * RADIUS:
            return self.rect(x, y, w, h, color)
        self.ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={RADIUS}",
                                         facecolor=color, edgecolor="none", linewidth=0, mutation_aspect=1))
        if end == "top":
            self.rect(x, y + h / 2, w, h / 2, color)
        else:
            self.rect(x, y, w / 2, h, color)

    def line(self, x1, y1, x2, y2, color, width=2, z=0.5):
        """Gridlines sit behind the bars (z 0.5); pass z=3 for the baseline, which sits on top."""
        self.ax.plot([x1, x2], [y1, y2], color=color, linewidth=width * 0.36, solid_capstyle="butt", zorder=z)

    def text(self, x, y, s, size, color=INK_2, bold=False, align="left", base="middle"):
        va = {"middle": "center", "top": "top", "alphabetic": "baseline", "bottom": "bottom"}[base]
        self.ax.text(x, y, s, fontsize=size * 0.36, color=color, fontweight="bold" if bold else "normal",
                     ha={"left": "left", "right": "right", "center": "center"}[align], va=va, family="Arial")

    def save(self, path: Path):
        self.fig.savefig(path, dpi=200, facecolor="white")
        plt.close(self.fig)
        return path


def header_height(spec) -> int:
    lines = len(wrap(spec["title"]))
    legend = spec["kind"] in ("stacked_columns", "stacked_hbars")
    return 62 + 40 * lines + (44 if legend else 0)


def draw_header(c: Canvas, spec):
    y = 40
    for line in wrap(spec["title"]):
        c.text(PADX, y, line, 34, INK, bold=True, base="alphabetic")
        y += 40
    c.text(PADX, y, spec["subtitle"], 25, INK_2, base="alphabetic")
    if spec["kind"] in ("stacked_columns", "stacked_hbars"):
        x = PADX
        for s in spec["series"]:
            c.rect(x, y + 22, 22, 22, s["color"])
            c.text(x + 32, y + 33, s["name"], 24, INK_2)
            x += 32 + 14 * len(s["name"]) + 40


def plan(spec) -> int:
    """Image height for a spec (same arithmetic in JS)."""
    hh = header_height(spec)
    k = spec["kind"]
    if k in ("columns", "stacked_columns"):
        return hh + 380
    if k in ("hbars", "stacked_hbars"):
        return hh + 16 + 52 * max(len(spec["labels"]), 1) + 20
    if k == "stack100":
        rows = math.ceil(len(spec["labels"]) / 3)
        return hh + 30 + 48 + 34 + 44 * rows + 10
    raise ValueError(k)


def render(spec, path: Path) -> Path:
    H = plan(spec)
    c = Canvas(H)
    draw_header(c, spec)
    hh, k = header_height(spec), spec["kind"]

    if k in ("columns", "stacked_columns"):
        x0, x1, y0, y1 = PADX + 64, W - PADX, hh + 30, H - 80
        stacks = (spec["values"] and [[v] for v in spec["values"]]) if k == "columns" else \
            [[s["values"][i] for s in spec["series"]] for i in range(len(spec["labels"]))]
        top = nice_max(max([sum(s) for s in stacks] + [1]) * 1.12)
        for t in range(5):
            y = y1 - (y1 - y0) * t / 4
            if t:
                c.line(x0, y, x1, y, GRID, 2)
            label = f"{top * t / 4:,.0f}" if (top * t / 4).is_integer() else f"{top * t / 4:,.1f}"
            c.text(x0 - 12, y, label, 22, INK_2, align="right")
        slot = (x1 - x0) / len(stacks)
        bw = min(BAR, slot * 0.6)
        for i, st in enumerate(stacks):
            x = x0 + i * slot + (slot - bw) / 2
            yb = y1
            last = max((j for j, v in enumerate(st) if v > 0), default=-1)
            for j, v in enumerate(st):
                if v <= 0:
                    continue
                h = (y1 - y0) * v / top
                color = spec["colors"][i] if k == "columns" else spec["series"][j]["color"]
                gap = GAP if j < last else 0
                c.bar(x, yb - h + gap, bw, h - gap, color, "top" if j == last else "none")
                yb -= h
            if str(i) in spec.get("value_text", {}):
                c.text(x + bw / 2, yb - 12, spec["value_text"][str(i)], 26, INK, bold=True, align="center", base="bottom")
            c.text(x + bw / 2, y1 + 14, spec["labels"][i], 22, INK_2, align="center", base="top")
            if spec["labels2"][i]:
                c.text(x + bw / 2, y1 + 42, spec["labels2"][i], 22, INK_2, align="center", base="top")
        c.line(x0, y1, x1, y1, AXIS, 2, z=3)

    elif k in ("hbars", "stacked_hbars"):
        lw, y = 520, hh + 16
        x0, x1 = PADX + lw, W - PADX - 90
        rows = [[v] for v in spec["values"]] if k == "hbars" else \
            [[s["values"][i] for s in spec["series"]] for i in range(len(spec["labels"]))]
        top = max([sum(r) for r in rows] + [1])
        for i, r in enumerate(rows):
            cy = y + 52 * i + 26
            c.text(x0 - 14, cy, fit_label(spec["labels"][i]), 24, INK_2, align="right")
            x = x0
            last = max((j for j, v in enumerate(r) if v > 0), default=-1)
            for j, v in enumerate(r):
                if v <= 0:
                    continue
                w = (x1 - x0) * v / top
                color = spec["colors"][i] if k == "hbars" else spec["series"][j]["color"]
                gap = GAP if j < last else 0
                c.bar(x, cy - BAR_H / 2, w - gap, BAR_H, color, "right" if j == last else "none")
                x += w
            c.text(x + 12, cy, f"{sum(r):,}", 24, INK)
        if rows:
            c.line(x0, y, x0, y + 52 * len(rows), AXIS, 2, z=3)

    elif k == "stack100":
        y, x0, x1 = hh + 30, PADX, W - PADX
        total = sum(spec["values"]) or 1
        x = x0
        nz = [i for i, v in enumerate(spec["values"]) if v > 0]
        for i, v in enumerate(spec["values"]):
            if v <= 0:
                continue
            w = (x1 - x0) * v / total
            gap = GAP if i != nz[-1] else 0
            # rounded only at the two outer ends of the whole bar
            if i == nz[0] and i == nz[-1]:
                c.bar(x, y, w, 48, spec["colors"][i], "right")
            elif i == nz[-1]:
                c.bar(x, y, w - gap, 48, spec["colors"][i], "right")
            else:
                c.rect(x, y, w - gap, 48, spec["colors"][i])
            x += w
        colw = (x1 - x0) / 3
        for i, txt in enumerate(spec["legend_text"]):
            lx, ly = x0 + (i % 3) * colw, y + 48 + 34 + 44 * (i // 3)
            c.rect(lx, ly - 11, 22, 22, spec["colors"][i])
            c.text(lx + 32, ly, txt, 24, INK_2)
    return c.save(path)


def render_all(specs: dict, out_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = {}
    for key, spec in specs.items():
        if spec["kind"] in ("hbars", "stacked_hbars") and not spec["labels"]:
            continue
        paths[key] = render(spec, out_dir / f"{key}.png")
    return paths
