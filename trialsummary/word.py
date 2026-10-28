"""Write the client-facing Word summary: key figures, highlights, one part per section, definitions."""
from __future__ import annotations

from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from . import charts, narrative
from .audience import AUDIENCES

NAVY, BLUE, ACCENT, INK, INK_2, RED = "184F95", "1C5CAB", "2A78D6", "0B0B0B", "52514E", "D03B3B"
TILE, ZEBRA, TOTAL, RULE = "EEF4FC", "F4F7FB", "E6EEF8", "D9D8D0"
PALE_BLUE, MIST = "CDE2FB", "B7D3F6"           # text on the navy band
WARN, WARN_FILL = "FAB219", "FFF7E6"           # status 'warning': the Needs attention callout (always labelled)
CONTENT_WIDTH = 7.1  # inches on US Letter with 0.7 inch margins

# Which tables go in the Word summary, and which columns (None = all). Everything else stays in the workbook.
WORD_LAYOUT = {
    "enrollment": [("funnel", None), ("conversion", None), ("enrol_trend", None), ("status_eom", None),
                   ("enrol_by_site_week", [0, 2, 3, 4, 5, 6, 7, 8, 9, 11]), ("enrol_by_type", None)],
    "completeness": [("crf_by_site", [0, 1, 2, 3, 4, 5]), ("epro_by_form", [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]),
                     ("crf_by_form", [0, 1, 2, 3, 4, 5])],
    "post_surgical": [("psf_summary", None), ("psf_turnaround", None), ("psf_by_site", None)],
    "blood": [("blood_summary", None), ("blood_trend", None), ("blood_pairs", None), ("blood_by_site", None)],
    "queries": [("query_summary", None), ("query_days_open", None),
                ("query_by_site", [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]), ("query_by_type", [0, 1, 2, 3, 4, 5, 6, 7, 8, 9])],
}


# ---------------------------------------------------------------- low-level helpers
def _rgb(h):
    return RGBColor.from_string(h)


def _shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _borders(table, color=RULE):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "bottom", "insideH"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:color"), color)
        b.append(e)
    for edge in ("left", "right", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "nil")
        b.append(e)
    tblPr.append(b)


def _cell_margins(table, top=40, bottom=40, left=70, right=70):
    tblPr = table._tbl.tblPr
    m = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        e = OxmlElement(f"w:{k}")
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        m.append(e)
    tblPr.append(m)


def _repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    e = OxmlElement("w:tblHeader")
    e.set(qn("w:val"), "true")
    trPr.append(e)


def _keep_with_next(p):
    p.paragraph_format.keep_with_next = True


def _text(cell, text, size=8.5, bold=False, color=INK, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run(text)
    r.font.size, r.font.bold, r.font.color.rgb = Pt(size), bold, _rgb(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return r


def fmt(col, v):
    if v is None:
        return "–"
    k = col.kind
    if k == "pct":
        return "<1%" if 0 < v < 0.005 else f"{v:.0%}"
    if k == "rate":
        return f"{v:.2f}"
    if k == "dec":
        return f"{v:.1f}"
    if k == "int":
        return f"{v:,.2f}" if isinstance(v, float) else f"{v:,}"
    if k == "date":
        return f"{v.day} {v:%b %Y}"
    if k == "range":
        return f"{v[0]:.0f}–{v[1]:.0f}"
    return str(v)


# ---------------------------------------------------------------- building blocks
def _para(doc, text="", size=10, bold=False, color=INK, after=4, before=0, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after, p.paragraph_format.space_before = Pt(after), Pt(before)
    if text:
        r = p.add_run(text)
        r.font.size, r.font.bold, r.font.italic, r.font.color.rgb = Pt(size), bold, italic, _rgb(color)
    return p


def _rule(doc, color=NAVY, size=8):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", str(size)), ("w:space", "1"), ("w:color", color)):
        bottom.set(qn(k), v)
    bdr.append(bottom)
    pPr.append(bdr)


def _table(doc, t, cols_idx=None):
    cols_idx = cols_idx or list(range(len(t.cols)))
    cols = [t.cols[i] for i in cols_idx]
    h = doc.add_paragraph()
    h.paragraph_format.space_before, h.paragraph_format.space_after = Pt(8), Pt(3)
    r = h.add_run(t.title)
    r.font.size, r.font.bold, r.font.color.rgb = Pt(10.5), True, _rgb(NAVY)
    _keep_with_next(h)
    widths = [Inches(w) for w in _col_widths(t, cols_idx)]
    tbl = doc.add_table(rows=1, cols=len(cols))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    _borders(tbl)
    _cell_margins(tbl)
    hdr = tbl.rows[0]
    _repeat_header(hdr)
    for j, c in enumerate(cols):
        cell = hdr.cells[j]
        cell.width = widths[j]
        _shade(cell, BLUE)
        _text(cell, c.header, size=8, bold=True, color="FFFFFF",
              align=WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER)
    rows = t.rows or [["None"] + [None] * (len(t.cols) - 1)]
    for i, row in enumerate(rows):
        cells = tbl.add_row().cells
        total = i in t.total_rows
        for j, ci in enumerate(cols_idx):
            col = t.cols[ci]
            cells[j].width = widths[j]
            _text(cells[j], fmt(col, row[ci]), size=8.5, bold=total,
                  align=WD_ALIGN_PARAGRAPH.LEFT if col.kind == "text" else WD_ALIGN_PARAGRAPH.RIGHT)
            if total:
                _shade(cells[j], TOTAL)
            elif i % 2 == 1:
                _shade(cells[j], ZEBRA)
    _keep_rows_together(tbl, whole=len(tbl.rows) <= 16)
    if t.note:
        _para(doc, t.note, size=7.5, color=INK_2, italic=True, before=2, after=6)


def _keep_rows_together(tbl, whole):
    """Rows never split across pages. Short tables stay on one page; long ones keep the header with 3 rows."""
    n = len(tbl.rows)
    for i, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        e = OxmlElement("w:cantSplit")
        e.set(qn("w:val"), "true")
        trPr.append(e)
        if i < n - 1 and (whole or i < 3):
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_with_next = True


def _col_widths(t, cols_idx):
    """Inches per column. Numeric columns fit their longest header word and longest value.
    The first column takes the space left over."""
    char = 0.068  # inches per character at 8 to 8.5 pt Arial
    out = []
    for n, ci in enumerate(cols_idx):
        col = t.cols[ci]
        if n == 0:
            out.append(None)
            continue
        longest_word = max(len(w) for w in col.header.split())
        longest_val = max([len(fmt(col, r[ci])) for r in t.rows] or [1])
        if col.kind == "text":
            out.append(max((col.width or 12) * 0.06, longest_word * char + 0.2))
        else:
            out.append(max(0.5, longest_word * char + 0.22, longest_val * char + 0.22,
                           len(col.header) * char / 2 + 0.22))  # header fits on two lines
    rest = sum(out[1:])
    first = CONTENT_WIDTH - rest
    if first < 1.7:  # too wide: shrink the other columns proportionally
        k = (CONTENT_WIDTH - 1.7) / rest
        out = [None] + [w * k for w in out[1:]]
        first = 1.7
    out[0] = first
    return out


def _cell_borders(cell, **edges):
    """Per-cell borders: edges as name=(size_eighths, hex) or name=None for no border."""
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{edge}")
        spec = edges.get(edge)
        if spec is None:
            e.set(qn("w:val"), "nil")
        else:
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), str(spec[0]))
            e.set(qn("w:color"), spec[1])
        b.append(e)
    tcPr.append(b)


def _no_table_borders(tbl):
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "nil")
        b.append(e)
    tbl._tbl.tblPr.append(b)


def _grid(doc, rows, cols, widths, margins=(90, 90, 120, 120)):
    tbl = doc.add_table(rows=rows, cols=cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    _no_table_borders(tbl)
    _cell_margins(tbl, *margins)
    for row in tbl.rows:
        for j, cell in enumerate(row.cells):
            cell.width = Inches(widths[j])
    return tbl


def _add(cell, text, size, bold=False, color=INK, first=False, after=0, align=WD_ALIGN_PARAGRAPH.LEFT, italic=False):
    p = cell.paragraphs[0] if first else cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(0), Pt(after)
    r = p.add_run(text)
    r.font.size, r.font.bold, r.font.italic, r.font.color.rgb = Pt(size), bold, italic, _rgb(color)
    return p


def _spacer(doc, pts=6, page_break=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(0), Pt(0)
    p.paragraph_format.page_break_before = page_break
    r = p.add_run("")
    r.font.size = Pt(1)
    p.paragraph_format.line_spacing = Pt(pts)
    return p


def _title_band(doc, period, meta, audience):
    tbl = _grid(doc, 1, 1, [CONTENT_WIDTH], margins=(160, 160, 240, 240))
    cell = tbl.rows[0].cells[0]
    _shade(cell, NAVY)
    _add(cell, "CLEO DIAGNOSTICS · CSP-01-001", 8.5, True, MIST, first=True, after=4)
    _add(cell, "Monthly trial summary", 22, True, "FFFFFF")
    _add(cell, period.label, 15, False, "FFFFFF", after=4)
    _add(cell, (f"Data to {period.last_day.day} {period.last_day:%B %Y} · Citrus export {meta['citrus_export']:%d %b %Y} · "
                f"Central lab inventory report {meta['inventory_date']}"), 8.5, False, PALE_BLUE)
    p = _para(doc, "", after=0, before=6)
    r = p.add_run("●  " + ("Sponsor version · Draft for review. Check every figure before sending to the sponsor." if audience == "sponsor"
                               else "Internal version · Not for the sponsor. Contains participant-level references."))
    r.font.size, r.font.bold, r.font.color.rgb = Pt(8.5), True, _rgb(RED)


def _tiles(doc, tiles):
    """Key figures: 3 per row, a blue accent edge on top, white gutters between tiles."""
    rows = (len(tiles) + 2) // 3
    tbl = _grid(doc, rows, 3, [CONTENT_WIDTH / 3] * 3, margins=(80, 80, 140, 140))
    gutter = (36, "FFFFFF")
    for k in range(rows * 3):
        cell = tbl.rows[k // 3].cells[k % 3]
        if k >= len(tiles):
            continue
        value, label, sub = tiles[k]
        _shade(cell, TILE)
        _cell_borders(cell, top=gutter, left=gutter, right=gutter, bottom=gutter)
        _add(cell, value, 18, True, NAVY, first=True)
        _add(cell, label, 9, True, INK)
        _add(cell, sub, 8, False, INK_2)


def _callout(doc, title, items, edge, fill, note=None):
    """A shaded box with a coloured left edge, a heading and bullet points."""
    tbl = _grid(doc, 1, 1, [CONTENT_WIDTH], margins=(100, 100, 220, 200))
    cell = tbl.rows[0].cells[0]
    _shade(cell, fill)
    _cell_borders(cell, left=(36, edge))
    _add(cell, title, 10.5, True, NAVY, first=True, after=2)
    for it in items:
        p = cell.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(1.5)
        r = p.add_run(it)
        r.font.size, r.font.color.rgb = Pt(9), _rgb(INK)
    if note:
        _add(cell, note, 8, False, INK_2, italic=True, after=0)
    _spacer(doc, 6)


def _inside(doc, rows):
    """'Inside this report': one row per section with its main finding."""
    _para(doc, "Inside this report", size=10.5, bold=True, color=NAVY, before=2, after=3)
    tbl = _grid(doc, len(rows), 3, [0.4, 1.7, CONTENT_WIDTH - 2.1], margins=(45, 45, 100, 100))
    for i, (name, finding) in enumerate(rows):
        c0, c1, c2 = tbl.rows[i].cells
        _shade(c0, NAVY)
        _add(c0, str(i + 1), 10, True, "FFFFFF", first=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        for c in (c1, c2):
            _shade(c, ZEBRA if i % 2 else "FFFFFF")
            _cell_borders(c, bottom=(4, RULE))
        _add(c1, name, 9.5, True, NAVY, first=True)
        _add(c2, finding, 8.5, False, INK, first=True)
        for c in (c0, c1, c2):
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _section_opener(doc, number, title, intro):
    _spacer(doc, 2, page_break=True)
    tbl = _grid(doc, 1, 2, [0.6, CONTENT_WIDTH - 0.6], margins=(150, 150, 160, 160))
    c0, c1 = tbl.rows[0].cells
    _shade(c0, NAVY)
    _shade(c1, TILE)
    _add(c0, str(number), 22, True, "FFFFFF", first=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    c0.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _add(c1, title, 16, True, NAVY, first=True, after=2)
    _add(c1, intro, 9, False, INK_2)
    for p in c1.paragraphs + c0.paragraphs:
        p.paragraph_format.keep_with_next = True


def _figure_strip(doc, figures):
    """Four key figures for one section, in a row of light cards with a blue left edge."""
    _spacer(doc, 6)
    n = len(figures)
    tbl = _grid(doc, 1, n, [CONTENT_WIDTH / n] * n, margins=(90, 90, 130, 100))
    for j, (value, label, sub) in enumerate(figures):
        cell = tbl.rows[0].cells[j]
        _shade(cell, ZEBRA)
        _cell_borders(cell, left=(24, ACCENT), right=(36, "FFFFFF"))
        _add(cell, value, 15, True, NAVY, first=True)
        _add(cell, label, 8.5, True, INK)
        _add(cell, sub, 7.5, False, INK_2)
    _spacer(doc, 6)


def _subhead(doc, text):
    p = _para(doc, text.upper(), size=8, bold=True, color=INK_2, before=10, after=2)
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "4"), ("w:space", "2"), ("w:color", RULE)):
        bottom.set(qn(k), v)
    bdr.append(bottom)
    pPr.append(bdr)
    p.paragraph_format.keep_with_next = True


def _chart(doc, path):
    """A story chart: the takeaway title is drawn inside the image, so no caption is needed."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(6), Pt(6)
    p.paragraph_format.keep_together = True
    p.add_run().add_picture(str(path), width=Inches(CONTENT_WIDTH))


def _field(run, code):
    for tag, attr in (("w:fldChar", {"w:fldCharType": "begin"}), ("w:instrText", None), ("w:fldChar", {"w:fldCharType": "end"})):
        e = OxmlElement(tag)
        if attr:
            for k, v in attr.items():
                e.set(qn(k), v)
        else:
            e.set(qn("xml:space"), "preserve")
            e.text = f" {code} "
        run._r.append(e)


def _header_footer(sec, period, aud_label):
    """Running header from page 2 on; footer on every page with 'Page X of Y'."""
    sec.different_first_page_header_footer = True
    hp = sec.header.paragraphs[0]
    r = hp.add_run(f"Cleo Diagnostics CSP-01-001  ·  Monthly trial summary  ·  {period.label}")
    r.font.size, r.font.bold, r.font.color.rgb = Pt(7.5), True, _rgb(NAVY)
    pPr = hp._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "4"), ("w:space", "4"), ("w:color", RULE)):
        bottom.set(qn(k), v)
    bdr.append(bottom)
    pPr.append(bdr)
    for footer in (sec.footer, sec.first_page_footer):
        p = footer.paragraphs[0]
        r = p.add_run(f"{aud_label} version  ·  Prepared by Lindus Health  ·  Page ")
        r.font.size, r.font.color.rgb = Pt(7.5), _rgb(INK_2)
        for code, tail in (("PAGE", " of "), ("NUMPAGES", "")):
            run = p.add_run()
            run.font.size, run.font.color.rgb = Pt(7.5), _rgb(INK_2)
            _field(run, code)
            if tail:
                t = p.add_run(tail)
                t.font.size, t.font.color.rgb = Pt(7.5), _rgb(INK_2)


# ---------------------------------------------------------------- document
def write_document(path: Path, period, sections, head, meta, chart_dir: Path, audience="internal"):
    aud = AUDIENCES[audience]
    completeness = "completeness" not in aud["exclude_sections"]
    T = {t.key: t for s in sections.values() for t in s.tables}
    specs = charts.chart_specs(sections, period, head)
    images = charts.render_all(specs, chart_dir)
    included = [k for k in WORD_LAYOUT if k not in aud["exclude_sections"]]
    dig = narrative.digest(sections, head, period, included)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.orientation = WD_ORIENT.PORTRAIT
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Inches(0.7))
    sec.top_margin, sec.bottom_margin = Inches(0.6), Inches(0.6)
    sec.header_distance, sec.footer_distance = Inches(0.3), Inches(0.3)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Arial", Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    _header_footer(sec, period, aud["label"])

    # ---- page 1: title band, key figures, highlights, needs attention, inside this report
    _title_band(doc, period, meta, audience)
    _spacer(doc, 8)
    h = head
    _tiles(doc, [
        (f"{h['enrolled_month']}", f"Enrolled in {period.short}", f"{h['enrolled_to_date']:,} enrolled to date"),
        (f"{h['per_site_month']:.2f}", "Enrolled per active site", f"in {period.short} · {h['per_site_week']:.2f} per site per week"),
        (f"{h['crf_pct']:.0%}", "CRF forms complete", f"{h['crf_done']:,} of {h['crf_expected']:,} expected forms") if completeness
        else (f"{h['surgery_month']}", f"Surgeries in {period.short}", f"{h['surgery_to_date']:,} surgeries complete to date"),
        (f"{h['reports_pct']:.0%}", "Pathology reports uploaded", f"{h['reports_to_date']:,} of {h['surgery_to_date']:,} surgeries complete"),
        (f"{h['window_pct_month']:.0%}", "Samples inside window", f"{h['window_n_month']} of {h['window_d_month']} registered in {period.short}"),
        (f"{h['q_open'] + h['q_reviewing']}", "Unresolved queries", f"median {h['q_median_days']:.0f} days open"),
    ])
    _spacer(doc, 8)
    _callout(doc, "Highlights", narrative.highlights(head, period, completeness, audience), ACCENT, ZEBRA)
    acts = narrative.actions(head, sections, period, audience)
    if acts:
        _callout(doc, "Needs attention", acts, WARN, WARN_FILL,
                 note="Participant-level lists for each item are in the Internal Excel workbook." if aud["detail_tables"] else None)
    _inside(doc, dig["inside"])

    # ---- sections: opener, key figures, charts, then the detail tables
    for n, key in enumerate(included, start=1):
        s, layout = sections[key], WORD_LAYOUT[key]
        _section_opener(doc, n, s.title, s.intro)
        _figure_strip(doc, dig["figures"][key])
        order = [t for t, _ in layout]
        story = sorted((k for k, sp in specs.items() if k in images and (sp.get("before") in order or sp.get("after") in order)),
                       key=lambda k: order.index(specs[k].get("before") or specs[k].get("after")))
        for k in story:
            _chart(doc, images[k])
        replaced = {specs[k]["after"] for k in story if specs[k].get("replaces_table")}
        _subhead(doc, "Detail tables")
        for item, cols in layout:
            if item not in replaced:
                _table(doc, T[item], cols)

    # ---- definitions
    _spacer(doc, 2, page_break=True)
    _para(doc, "Definitions and sources", size=16, bold=True, color=NAVY, after=6)
    tbl = doc.add_table(rows=0, cols=2)
    tbl.autofit = False
    _borders(tbl)
    _cell_margins(tbl)
    for i, (term, text) in enumerate(narrative.definitions(period, meta, completeness)):
        cells = tbl.add_row().cells
        cells[0].width, cells[1].width = Inches(1.7), Inches(CONTENT_WIDTH - 1.7)
        _text(cells[0], term, size=8.5, bold=True, color=NAVY)
        _text(cells[1], text, size=8.5)
        if i % 2:
            _shade(cells[0], ZEBRA)
            _shade(cells[1], ZEBRA)
    doc.save(path)
    return path
