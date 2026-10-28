/* Trial Summary Builder: output writers (Word, Excel, totals CSV) and chart drawing.
 * Mirrors trialsummary/word.py, excel.py and build_report.py. Uses the global `docx` (docx 8.5.0) and `ExcelJS` (4.4.0).
 * `sel` is the user's content selection, resolved by TSWriters.resolve() into the tables and columns to show.
 */
(function (root) {
  "use strict";
  var E = root.TSEngine;
  var NAVY = "184F95", BLUE = "1C5CAB", INK = "0B0B0B", INK_2 = "52514E", RED = "D03B3B";
  var TILE = "EEF4FC", ZEBRA = "F4F7FB", TOTAL = "E6EEF8", RULE = "D9D8D0";
  var CONTENT_IN = 7.1, DXA = 1440;

  // Columns that fit a portrait Word page. Excel always gets every column the user keeps.
  var WORD_COLS = {
    enrol_by_site_week: function (t) { var n = t.cols.length; return t.cols.map(function (_, i) { return i; }).filter(function (i) { return i !== 1 && i !== n - 2; }); },
    crf_by_site: [0, 1, 2, 3, 4, 5], crf_by_form: [0, 1, 2, 3, 4, 5], epro_by_form: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9], epro_by_site: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9],
    query_by_site: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9], query_by_type: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9], query_by_form: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
  };

  var WORD_ORDER = { enrollment: ["funnel", "conversion", "enrol_trend", "status_eom", "enrol_by_site_week", "enrol_by_type"] };

  // ------------------------------------------------------------ value formatting (matches word.py fmt)
  function fmt(col, v) {
    if (v == null) return "–";
    switch (col.kind) {
      case "pct": return E.fmtPct(v);
      case "rate": return v.toFixed(2);
      case "dec": return E.roundHalfEven(v, 1).toFixed(1);
      case "int": return Number.isInteger(v) ? E.fmtInt(v) : v.toFixed(2);
      case "date": return E.fmtDate(v);
      case "range": return E.fmtD0(v[0]) + "–" + E.fmtD0(v[1]);
      default: return String(v);
    }
  }

  // ------------------------------------------------------------ selection
  // sel = { audience, sections:{key:bool}, tables:{key:bool}, hiddenCols:{tableKey:[colIndex]}, page1:{tiles,highlights,attention} }
  function resolve(result, sel) {
    var out = [];
    Object.keys(result.sections).forEach(function (k) {
      var s = result.sections[k];
      if (!sel.sections[k]) return;
      var tables = s.tables.filter(function (t) {
        if (t.detail) return sel.audience === "internal" && sel.sections.lists && sel.tables[t.key];
        return sel.tables[t.key] !== false;
      });
      if (tables.length) out.push({ key: k, title: s.title, intro: s.intro, tables: tables });
    });
    return out;
  }
  function keptCols(t, sel, forWord) {
    var hidden = (sel.hiddenCols && sel.hiddenCols[t.key]) || [];
    var idx = t.cols.map(function (_, i) { return i; });
    if (forWord && WORD_COLS[t.key]) idx = typeof WORD_COLS[t.key] === "function" ? WORD_COLS[t.key](t) : WORD_COLS[t.key];
    return idx.filter(function (i) { return hidden.indexOf(i) < 0; });
  }
  function tableOn(result, sel, key) {
    var found = null;
    Object.keys(result.sections).forEach(function (k) { result.sections[k].tables.forEach(function (t) { if (t.key === key) found = k; }); });
    return found != null && sel.sections[found] && sel.tables[key] !== false;
  }

  // Page-1 content that follows the selection: tiles, highlights and action items drop out with their tables.
  function pageOne(result, sel) {
    var h = result.head, p = result.period, on = function (k) { return tableOn(result, sel, k); };
    var tiles = [];
    if (on("funnel")) tiles.push([String(h.enrolled_month), "Enrolled in " + p.short, E.fmtInt(h.enrolled_to_date) + " enrolled to date"]);
    if (on("enrol_by_site_week")) tiles.push([h.per_site_month.toFixed(2), "Enrolled per active site", "in " + p.short + " · " + h.per_site_week.toFixed(2) + " per site per week"]);
    if (on("crf_by_site")) tiles.push([E.fmtPct(h.crf_pct), "CRF forms complete", E.fmtInt(h.crf_done) + " of " + E.fmtInt(h.crf_expected) + " expected forms"]);
    else if (on("psf_summary")) tiles.push([String(h.surgery_month), "Surgeries in " + p.short, E.fmtInt(h.surgery_to_date) + " surgeries complete to date"]);
    if (on("psf_summary")) tiles.push([E.fmtPct(h.reports_pct), "Pathology reports uploaded", E.fmtInt(h.reports_to_date) + " of " + E.fmtInt(h.surgery_to_date) + " surgeries complete"]);
    if (on("blood_summary")) tiles.push([E.fmtPct(h.window_pct_month), "Samples inside window", h.window_n_month + " of " + h.window_d_month + " registered in " + p.short]);
    if (on("query_summary")) tiles.push([String(h.q_open + h.q_reviewing), "Unresolved queries", "median " + E.fmtD0(h.q_median_days) + " days open"]);
    var hl, acts;
    if (sel.audience === "sponsor") {
      // Sponsors get high-level progress and site-facing items only; each line names the table it depends on.
      var pick = function (list) { return list.filter(function (x) { return on(x.key); }).map(function (x) { return x.text; }); };
      hl = pick(E.sponsorHighlights(h, p));
      acts = pick(E.sponsorActions(h, p));
    } else {
      var hlKeys = ["funnel", "enrol_by_site_week", "crf_by_site", "psf_summary", "blood_summary", "query_summary"];
      hl = E.highlights(h, p, true).filter(function (_, i) { return on(hlKeys[i]); });
      var actDrive = { pairs: "blood_pairs", nosample: "blood_pairs", late: "blood_summary", sameday: "blood_summary", path: "psf_summary", old: "query_days_open" };
      acts = E.actions(h, result.sections, p).filter(function (a) { return on(actDrive[a.key]); }).map(function (a) { return a.text; });
    }
    return {
      tiles: sel.page1.tiles ? tiles : [], highlights: sel.page1.highlights ? hl : [], actions: sel.page1.attention ? acts : []
    };
  }

  // ------------------------------------------------------------ story charts (canvas -> PNG bytes)
  // Mirror of trialsummary/charts.py render(): same 1400 px canvas, same geometry, same colours, so both tools match.
  var CW = 1400, PADX = 16, BAR = 48, BAR_H = 32, RADIUS = 8, GAP = 4;
  var C = { ink: "#0b0b0b", ink2: "#52514e", grid: "#e1e0d9", axis: "#c3c2b7" };
  function wrapTitle(text, limit) {
    limit = limit || 70;
    var lines = [], cur = "";
    text.split(" ").forEach(function (w) {
      if (cur && cur.length + 1 + w.length > limit) { lines.push(cur); cur = w; } else cur = cur ? cur + " " + w : w;
    });
    if (cur) lines.push(cur);
    return lines;
  }
  function niceMax(v, ticks) {
    ticks = ticks || 4;
    if (v <= 0) return ticks;
    var raw = v / ticks, mag = Math.pow(10, Math.floor(Math.log10(raw))), f = [1, 2, 2.5, 5, 10];
    for (var i = 0; i < f.length; i++) if (f[i] * mag >= raw) return f[i] * mag * ticks;
    return 10 * mag * ticks;
  }
  function fitLabel(s, limit) { limit = limit || 38; return s.length <= limit ? s : s.slice(0, limit - 1) + "…"; }
  var hasLegend = function (spec) { return spec.kind === "stacked_columns" || spec.kind === "stacked_hbars"; };
  function headerHeight(spec) { return 62 + 40 * wrapTitle(spec.title).length + (hasLegend(spec) ? 44 : 0); }
  function planHeight(spec) {
    var hh = headerHeight(spec), k = spec.kind;
    if (k === "columns" || k === "stacked_columns") return hh + 380;
    if (k === "hbars" || k === "stacked_hbars") return hh + 16 + 52 * Math.max(spec.labels.length, 1) + 20;
    return hh + 30 + 48 + 34 + 44 * Math.ceil(spec.labels.length / 3) + 10;
  }
  function Pen(g) {
    this.rect = function (x, y, w, h, color) { if (w > 0 && h > 0) { g.fillStyle = color; g.fillRect(x, y, w, h); } };
    this.bar = function (x, y, w, h, color, end) {
      if (w <= 0 || h <= 0) return;
      var longSide = end === "top" ? h : w;
      if (end === "none" || longSide <= 2 * RADIUS || Math.min(w, h) < 2 * RADIUS) return this.rect(x, y, w, h, color);
      var r = RADIUS;
      g.fillStyle = color; g.beginPath();
      if (end === "top") {
        g.moveTo(x, y + h); g.lineTo(x, y + r); g.arcTo(x, y, x + r, y, r); g.lineTo(x + w - r, y); g.arcTo(x + w, y, x + w, y + r, r); g.lineTo(x + w, y + h);
      } else {
        g.moveTo(x, y); g.lineTo(x + w - r, y); g.arcTo(x + w, y, x + w, y + r, r); g.lineTo(x + w, y + h - r); g.arcTo(x + w, y + h, x + w - r, y + h, r); g.lineTo(x, y + h);
      }
      g.closePath(); g.fill();
    };
    this.line = function (x1, y1, x2, y2, color, width) { g.strokeStyle = color; g.lineWidth = width || 2; g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke(); };
    this.text = function (x, y, s, size, color, bold, align, base) {
      g.font = (bold ? "bold " : "") + size + "px Arial"; g.fillStyle = color || C.ink2;
      g.textAlign = align || "left"; g.textBaseline = base || "middle"; g.fillText(s, x, y);
    };
  }
  function renderSpec(spec) {
    var H = planHeight(spec), cv = document.createElement("canvas");
    cv.width = CW; cv.height = H;
    var g = cv.getContext("2d"), c = new Pen(g), k = spec.kind, hh = headerHeight(spec);
    g.fillStyle = "#ffffff"; g.fillRect(0, 0, CW, H);
    var y = 40;
    wrapTitle(spec.title).forEach(function (line) { c.text(PADX, y, line, 34, C.ink, true, "left", "alphabetic"); y += 40; });
    c.text(PADX, y, spec.subtitle, 25, C.ink2, false, "left", "alphabetic");
    if (hasLegend(spec)) {
      var lx = PADX;
      spec.series.forEach(function (s) { c.rect(lx, y + 22, 22, 22, s.color); c.text(lx + 32, y + 33, s.name, 24, C.ink2); lx += 32 + 14 * s.name.length + 40; });
    }
    var vt = spec.value_text || {};
    if (k === "columns" || k === "stacked_columns") {
      var x0 = PADX + 64, x1 = CW - PADX, y0 = hh + 30, y1 = H - 80;
      var stacks = k === "columns" ? spec.values.map(function (v) { return [v]; })
        : spec.labels.map(function (_, i) { return spec.series.map(function (s) { return s.values[i]; }); });
      var top = niceMax(Math.max.apply(null, stacks.map(function (st) { return st.reduce(function (a, b) { return a + b; }, 0); }).concat([1])) * 1.12);
      for (var t = 0; t < 5; t++) {
        var gy = y1 - (y1 - y0) * t / 4, tv = top * t / 4;
        if (t) c.line(x0, gy, x1, gy, C.grid, 2);
        c.text(x0 - 12, gy, Number.isInteger(tv) ? tv.toLocaleString("en-US") : tv.toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 }), 22, C.ink2, false, "right");
      }
      var slot = (x1 - x0) / stacks.length, bw = Math.min(BAR, slot * 0.6);
      stacks.forEach(function (st, i) {
        var x = x0 + i * slot + (slot - bw) / 2, yb = y1, last = -1;
        st.forEach(function (v, j) { if (v > 0) last = j; });
        st.forEach(function (v, j) {
          if (v <= 0) return;
          var h = (y1 - y0) * v / top, color = k === "columns" ? spec.colors[i] : spec.series[j].color, gap = j < last ? GAP : 0;
          c.bar(x, yb - h + gap, bw, h - gap, color, j === last ? "top" : "none");
          yb -= h;
        });
        if (vt[String(i)] != null) c.text(x + bw / 2, yb - 12, vt[String(i)], 26, C.ink, true, "center", "bottom");
        c.text(x + bw / 2, y1 + 14, spec.labels[i], 22, C.ink2, false, "center", "top");
        if (spec.labels2[i]) c.text(x + bw / 2, y1 + 42, spec.labels2[i], 22, C.ink2, false, "center", "top");
      });
      c.line(x0, y1, x1, y1, C.axis, 2);
    } else if (k === "hbars" || k === "stacked_hbars") {
      var lw = 520, yy = hh + 16, hx0 = PADX + lw, hx1 = CW - PADX - 90;
      var rows = k === "hbars" ? spec.values.map(function (v) { return [v]; })
        : spec.labels.map(function (_, i) { return spec.series.map(function (s) { return s.values[i]; }); });
      var htop = Math.max.apply(null, rows.map(function (r) { return r.reduce(function (a, b) { return a + b; }, 0); }).concat([1]));
      rows.forEach(function (r, i) {
        var cy = yy + 52 * i + 26, x = hx0, last = -1, sum = r.reduce(function (a, b) { return a + b; }, 0);
        c.text(hx0 - 14, cy, fitLabel(spec.labels[i]), 24, C.ink2, false, "right");
        r.forEach(function (v, j) { if (v > 0) last = j; });
        r.forEach(function (v, j) {
          if (v <= 0) return;
          var w = (hx1 - hx0) * v / htop, color = k === "hbars" ? spec.colors[i] : spec.series[j].color, gap = j < last ? GAP : 0;
          c.bar(x, cy - BAR_H / 2, w - gap, BAR_H, color, j === last ? "right" : "none");
          x += w;
        });
        c.text(x + 12, cy, sum.toLocaleString("en-US"), 24, C.ink);
      });
      if (rows.length) c.line(hx0, yy, hx0, yy + 52 * rows.length, C.axis, 2);
    } else if (k === "stack100") {
      var sy = hh + 30, sx0 = PADX, sx1 = CW - PADX, total = spec.values.reduce(function (a, b) { return a + b; }, 0) || 1, sx = sx0;
      var nz = spec.values.map(function (v, i) { return v > 0 ? i : -1; }).filter(function (i) { return i >= 0; });
      spec.values.forEach(function (v, i) {
        if (v <= 0) return;
        var w = (sx1 - sx0) * v / total, gap = i !== nz[nz.length - 1] ? GAP : 0;
        if (i === nz[nz.length - 1]) c.bar(sx, sy, w - gap, 48, spec.colors[i], "right"); else c.rect(sx, sy, w - gap, 48, spec.colors[i]);
        sx += w;
      });
      var colw = (sx1 - sx0) / 3;
      spec.legend_text.forEach(function (txt, i) {
        var lx2 = sx0 + (i % 3) * colw, ly = sy + 48 + 34 + 44 * Math.floor(i / 3);
        c.rect(lx2, ly - 11, 22, 22, spec.colors[i]); c.text(lx2 + 32, ly, txt, 24, C.ink2);
      });
    }
    return new Promise(function (res) { cv.toBlob(function (b) { b.arrayBuffer().then(function (ab) { res({ png: new Uint8Array(ab), w: CW, h: H }); }); }, "image/png"); });
  }
  // Renders every chart spec; a ranked chart with no rows is skipped.
  function drawCharts(result) {
    var specs = E.chartSpecs(result.sections, result.period, result.head), keys = Object.keys(specs).filter(function (k) {
      return !((specs[k].kind === "hbars" || specs[k].kind === "stacked_hbars") && !specs[k].labels.length);
    });
    return Promise.all(keys.map(function (k) { return renderSpec(specs[k]); })).then(function (imgs) {
      var out = { _specs: specs };
      keys.forEach(function (k, i) { out[k] = imgs[i]; });
      return out;
    });
  }
  // Charts placed around a table, honouring the queries chart toggle.
  function chartsFor(tableKey, charts, sel) {
    var out = { before: [], after: [], replaces: false };
    if (!charts || !charts._specs) return out;
    Object.keys(charts._specs).forEach(function (k) {
      var sp = charts._specs[k];
      if (!charts[k] || (k === "queries_by_site" && sel.tables.query_chart === false)) return;
      if (sp.before === tableKey) out.before.push(k);
      if (sp.after === tableKey) { out.after.push(k); if (sp.replaces_table) out.replaces = true; }
    });
    return out;
  }

  // ------------------------------------------------------------ Word
  function buildDocx(result, meta, sel, charts, opts) {
    var D = root.docx, p = result.period, aud = sel.audience, one = pageOne(result, sel), secs = resolve(result, sel);
    opts = opts || {};
    var T = {};
    Object.keys(result.sections).forEach(function (k) { result.sections[k].tables.forEach(function (t) { T[t.key] = t; }); });
    var run = function (text, o) { o = o || {}; return new D.TextRun({ text: text, font: "Arial", size: Math.round((o.size || 10) * 2), bold: !!o.bold, italics: !!o.italic, color: o.color || INK }); };
    var para = function (text, o) {
      o = o || {};
      return new D.Paragraph({ children: text ? [run(text, o)] : [], spacing: { before: (o.before || 0) * 20, after: (o.after == null ? 4 : o.after) * 20 },
        keepNext: !!o.keepNext, alignment: o.align, pageBreakBefore: !!o.pageBreak, heading: o.heading });
    };
    var noBorder = { style: D.BorderStyle.NONE, size: 0, color: "FFFFFF" }, hair = { style: D.BorderStyle.SINGLE, size: 4, color: RULE };
    var cellMargins = { top: 40, bottom: 40, left: 70, right: 70 };
    var children = [];

    function heading(text, level, pageBreak) {
      return new D.Paragraph({ heading: level === 1 ? D.HeadingLevel.HEADING_1 : D.HeadingLevel.HEADING_2, keepNext: true, pageBreakBefore: !!pageBreak,
        spacing: { before: (level === 1 ? 14 : 8) * 20, after: 80 }, children: [new D.TextRun({ text: text, font: "Arial", bold: true, size: level === 1 ? 28 : 22, color: NAVY })] });
    }
    function colWidths(t, idx) {
      var ch = 0.068, out = idx.map(function (ci, n) {
        if (n === 0) return null;
        var c = t.cols[ci], word = Math.max.apply(null, c.header.split(" ").map(function (w) { return w.length; }));
        var val = Math.max.apply(null, t.rows.map(function (r) { return fmt(c, r[ci]).length; }).concat([1]));
        if (c.kind === "text") return Math.max((c.width || 12) * 0.06, word * ch + 0.2);
        return Math.max(0.5, word * ch + 0.22, val * ch + 0.22, c.header.length * ch / 2 + 0.22);
      });
      var rest = out.slice(1).reduce(function (a, b) { return a + b; }, 0), first = CONTENT_IN - rest;
      if (first < 1.7) { var k = (CONTENT_IN - 1.7) / rest; out = out.map(function (w, i) { return i ? w * k : w; }); first = 1.7; }
      out[0] = first;
      return out.map(function (w) { return Math.round(w * DXA); });
    }
    function table(t) {
      var idx = keptCols(t, sel, true), widths = colWidths(t, idx), total = widths.reduce(function (a, b) { return a + b; }, 0);
      var rows = t.rows.length ? t.rows : [["None"].concat(t.cols.slice(1).map(function () { return null; }))];
      var whole = rows.length + 1 <= 16;
      var mkCell = function (text, w, o) {
        return new D.TableCell({ width: { size: w, type: D.WidthType.DXA }, margins: cellMargins, verticalAlign: D.VerticalAlign.CENTER,
          shading: o.fill ? { fill: o.fill, type: D.ShadingType.CLEAR, color: "auto" } : undefined,
          children: [new D.Paragraph({ alignment: o.align, keepNext: !!o.keepNext, spacing: { before: 0, after: 0 },
            children: [run(text, { size: o.size || 8.5, bold: o.bold, color: o.color })] })] });
      };
      var trs = [new D.TableRow({ tableHeader: true, cantSplit: true, children: idx.map(function (ci, n) {
        return mkCell(t.cols[ci].header, widths[n], { fill: BLUE, color: "FFFFFF", bold: true, size: 8, align: n ? D.AlignmentType.CENTER : D.AlignmentType.LEFT, keepNext: true });
      }) })];
      rows.forEach(function (r, i) {
        var tot = t.totalRows.has(i), keep = i < rows.length - 1 && (whole || i < 3);
        trs.push(new D.TableRow({ cantSplit: true, children: idx.map(function (ci, n) {
          var c = t.cols[ci];
          return mkCell(fmt(c, r[ci]), widths[n], { bold: tot, fill: tot ? TOTAL : i % 2 ? ZEBRA : null, keepNext: keep,
            align: c.kind === "text" ? D.AlignmentType.LEFT : D.AlignmentType.RIGHT });
        }) }));
      });
      var out = [para(t.title, { size: 10.5, bold: true, color: NAVY, before: 8, after: 3, keepNext: true }),
        new D.Table({ rows: trs, columnWidths: widths, width: { size: total, type: D.WidthType.DXA }, alignment: D.AlignmentType.CENTER,
          layout: D.TableLayoutType.FIXED,
          borders: { top: hair, bottom: hair, insideHorizontal: hair, left: noBorder, right: noBorder, insideVertical: noBorder } })];
      if (t.note) out.push(para(t.note, { size: 7.5, italic: true, color: INK_2, before: 2, after: 6 }));
      return out;
    }
    function chart(key) {  // the takeaway title is drawn inside the image, so no caption
      var c = charts[key], w = 680, h = Math.round(w * c.h / c.w);
      return [new D.Paragraph({ spacing: { before: 200, after: 160 }, keepLines: true,
        children: [new D.ImageRun({ data: c.png, transformation: { width: w, height: h }, type: "png" })] })];
    }

    // ---- layout helpers (mirror of trialsummary/word.py)
    var ACC = "2A78D6", MIST = "B7D3F6", PALE = "CDE2FB", WARN = "FAB219", WARN_FILL = "FFF7E6";
    var tw = function (inches) { return Math.round(inches * DXA); };
    var shade = function (hex) { return { fill: hex, type: D.ShadingType.CLEAR, color: "auto" }; };
    var edge = function (size, color) { return { style: D.BorderStyle.SINGLE, size: size, color: color }; };
    var noTable = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder };
    var line = function (text, size, o) {
      o = o || {};
      return new D.Paragraph({ alignment: o.align, keepNext: !!o.keepNext, spacing: { before: 0, after: (o.after || 0) * 20 },
        children: [run(text, { size: size, bold: o.bold, color: o.color, italic: o.italic })] });
    };
    var cellOf = function (width, kids, o) {
      o = o || {};
      return new D.TableCell({ width: { size: tw(width), type: D.WidthType.DXA }, margins: o.margins, shading: o.fill ? shade(o.fill) : undefined,
        borders: o.borders, verticalAlign: o.valign, children: kids });
    };
    var grid = function (widths, rows) {
      var total = widths.reduce(function (a, b) { return a + b; }, 0);
      return new D.Table({ columnWidths: widths.map(tw), width: { size: tw(total), type: D.WidthType.DXA }, layout: D.TableLayoutType.FIXED,
        alignment: D.AlignmentType.CENTER, borders: noTable, rows: rows });
    };
    var spacer = function (pts, pageBreak) {
      return new D.Paragraph({ pageBreakBefore: !!pageBreak, spacing: { before: 0, after: 0, line: Math.round(pts * 20), lineRule: D.LineRuleType.EXACT }, children: [] });
    };
    var included = secs.map(function (s) { return s.key; }), dig = E.digest(result.sections, result.head, p, included);

    // ---- page 1: title band
    var lastLong = new Date(p.lastDay * E.DAY).getUTCDate() + " " + p.label;
    children.push(grid([CONTENT_IN], [new D.TableRow({ children: [cellOf(CONTENT_IN, [
      line((opts.example ? "EXAMPLE DATA · " : "") + "CLEO DIAGNOSTICS · CSP-01-001", 8.5, { bold: true, color: MIST, after: 4 }),
      line("Monthly trial summary", 22, { bold: true, color: "FFFFFF" }),
      line(p.label, 15, { color: "FFFFFF", after: 4 }),
      line("Data to " + lastLong + (meta.citrusExport ? " · Citrus export " + E.fmtDate2(E.dayOf(meta.citrusExport)) : "") +
        (meta.inventoryDate ? " · Central lab inventory report " + meta.inventoryDate : ""), 8.5, { color: PALE })
    ], { fill: NAVY, margins: { top: 160, bottom: 160, left: 240, right: 240 } })] })]));
    children.push(new D.Paragraph({ spacing: { before: 120, after: 0 }, children: [run("●  " + (aud === "sponsor"
      ? "Sponsor version · Draft for review. Check every figure before sending to the sponsor."
      : "Internal version · Not for the sponsor. Contains participant-level references."), { size: 8.5, bold: true, color: RED })] }));
    children.push(spacer(8));

    // ---- key figure tiles, 3 per row with white gutters
    if (one.tiles.length) {
      var gut = edge(36, "FFFFFF"), trs = [];
      for (var i = 0; i < one.tiles.length; i += 3) {
        var cells = [];
        for (var j = i; j < i + 3; j++) {
          var t = one.tiles[j];
          cells.push(cellOf(CONTENT_IN / 3, t ? [line(t[0], 18, { bold: true, color: NAVY }), line(t[1], 9, { bold: true }), line(t[2], 8, { color: INK_2 })]
            : [new D.Paragraph({ children: [] })], { fill: t ? TILE : null, margins: { top: 80, bottom: 80, left: 140, right: 140 },
            borders: { top: gut, bottom: gut, left: gut, right: gut } }));
        }
        trs.push(new D.TableRow({ cantSplit: true, children: cells }));
      }
      children.push(grid([CONTENT_IN / 3, CONTENT_IN / 3, CONTENT_IN / 3], trs));
      children.push(spacer(8));
    }

    // ---- callouts: highlights (blue edge) and needs attention (amber edge, always labelled)
    var callout = function (title, items, edgeColor, fill, note) {
      var kids = [line(title, 10.5, { bold: true, color: NAVY, after: 2 })].concat(items.map(function (it) {
        return new D.Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 30 }, children: [run(it, { size: 9 })] });
      }));
      if (note) kids.push(line(note, 8, { italic: true, color: INK_2 }));
      children.push(grid([CONTENT_IN], [new D.TableRow({ cantSplit: true, children: [cellOf(CONTENT_IN, kids, { fill: fill,
        margins: { top: 100, bottom: 100, left: 220, right: 200 }, borders: { left: edge(36, edgeColor), top: noBorder, bottom: noBorder, right: noBorder } })] })]));
      children.push(spacer(6));
    };
    if (one.highlights.length) callout("Highlights", one.highlights, ACC, ZEBRA);
    if (one.actions.length) callout("Needs attention", one.actions, WARN, WARN_FILL,
      aud === "internal" && sel.sections.lists ? "Participant-level lists for each item are in the Internal Excel workbook." : null);

    // ---- inside this report
    if (dig.inside.length) {
      children.push(new D.Paragraph({ keepNext: true, spacing: { before: 40, after: 60 }, children: [run("Inside this report", { size: 10.5, bold: true, color: NAVY })] }));
      var iw = [0.4, 1.7, CONTENT_IN - 2.1], im = { top: 45, bottom: 45, left: 100, right: 100 };
      children.push(grid(iw, dig.inside.map(function (r, n) {
        var f = n % 2 ? ZEBRA : "FFFFFF", b = { bottom: hair, top: noBorder, left: noBorder, right: noBorder };
        return new D.TableRow({ cantSplit: true, children: [
          cellOf(iw[0], [line(String(n + 1), 10, { bold: true, color: "FFFFFF", align: D.AlignmentType.CENTER })], { fill: NAVY, margins: im, valign: D.VerticalAlign.CENTER }),
          cellOf(iw[1], [line(r[0], 9.5, { bold: true, color: NAVY })], { fill: f, margins: im, borders: b, valign: D.VerticalAlign.CENTER }),
          cellOf(iw[2], [line(r[1], 8.5)], { fill: f, margins: im, borders: b, valign: D.VerticalAlign.CENTER })] });
      })));
    }

    // ---- sections: numbered opener, key figures, charts, then the detail tables
    secs.forEach(function (s, n) {
      children.push(spacer(2, true));
      children.push(grid([0.6, CONTENT_IN - 0.6], [new D.TableRow({ cantSplit: true, children: [
        cellOf(0.6, [line(String(n + 1), 22, { bold: true, color: "FFFFFF", align: D.AlignmentType.CENTER, keepNext: true })],
          { fill: NAVY, margins: { top: 150, bottom: 150, left: 160, right: 160 }, valign: D.VerticalAlign.CENTER }),
        cellOf(CONTENT_IN - 0.6, [line(s.title, 16, { bold: true, color: NAVY, after: 2, keepNext: true }), line(s.intro, 9, { color: INK_2, keepNext: true })],
          { fill: TILE, margins: { top: 150, bottom: 150, left: 160, right: 160 } })] })]));
      var figs = dig.figures[s.key] || [];
      if (figs.length) {
        children.push(spacer(6));
        var fw = CONTENT_IN / figs.length;
        children.push(grid(figs.map(function () { return fw; }), [new D.TableRow({ cantSplit: true, children: figs.map(function (f) {
          return cellOf(fw, [line(f[0], 15, { bold: true, color: NAVY }), line(f[1], 8.5, { bold: true }), line(f[2], 7.5, { color: INK_2 })],
            { fill: ZEBRA, margins: { top: 90, bottom: 90, left: 130, right: 100 },
              borders: { left: edge(24, ACC), right: edge(36, "FFFFFF"), top: noBorder, bottom: noBorder } });
        }) })]));
        children.push(spacer(6));
      }
      var order = WORD_ORDER[s.key];
      var tables = (order ? s.tables.slice().sort(function (a, b) {
        var ia = order.indexOf(a.key), ib = order.indexOf(b.key);
        return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
      }) : s.tables).filter(function (t) { return !t.detail; });  // participant lists go to Excel only
      var story = [], replaced = {};
      tables.forEach(function (t) {
        var place = chartsFor(t.key, charts, sel);
        place.before.concat(place.after).forEach(function (k) { if (story.indexOf(k) < 0) story.push(k); });
        if (place.replaces) replaced[t.key] = true;
      });
      story.forEach(function (k) { children.push.apply(children, chart(k)); });
      children.push(new D.Paragraph({ keepNext: true, spacing: { before: 200, after: 40 }, border: { bottom: { style: D.BorderStyle.SINGLE, size: 4, color: RULE, space: 2 } },
        children: [run("DETAIL TABLES", { size: 8, bold: true, color: INK_2 })] }));
      tables.forEach(function (t) { if (!replaced[t.key]) children.push.apply(children, table(t)); });
    });

    // ---- definitions
    var defs = E.definitions(p, meta, !!sel.sections.completeness);
    children.push(spacer(2, true));
    children.push(new D.Paragraph({ spacing: { after: 120 }, children: [run("Definitions and sources", { size: 16, bold: true, color: NAVY })] }));
    var dw = [1.7, CONTENT_IN - 1.7];
    children.push(new D.Table({ columnWidths: dw.map(tw), width: { size: tw(CONTENT_IN), type: D.WidthType.DXA }, layout: D.TableLayoutType.FIXED,
      borders: { top: hair, bottom: hair, insideHorizontal: hair, left: noBorder, right: noBorder, insideVertical: noBorder },
      rows: defs.map(function (d, n) {
        var f = n % 2 ? ZEBRA : null;
        return new D.TableRow({ cantSplit: true, children: [
          cellOf(dw[0], [line(d[0], 8.5, { bold: true, color: NAVY })], { fill: f, margins: cellMargins }),
          cellOf(dw[1], [line(d[1], 8.5)], { fill: f, margins: cellMargins })] });
      }) }));

    // ---- running header from page 2, footer with Page X of Y on every page
    var audLabel = aud === "sponsor" ? "Sponsor" : "Internal";
    var footer = function () {
      var small = function (x) { return new D.TextRun({ children: [x], font: "Arial", size: 15, color: INK_2 }); };
      return new D.Footer({ children: [new D.Paragraph({ children: [run(audLabel + " version  ·  Prepared by Lindus Health  ·  Page ", { size: 7.5, color: INK_2 }),
        small(D.PageNumber.CURRENT), run(" of ", { size: 7.5, color: INK_2 }), small(D.PageNumber.TOTAL_PAGES)] })] });
    };
    var header = new D.Header({ children: [new D.Paragraph({ border: { bottom: { style: D.BorderStyle.SINGLE, size: 4, color: RULE, space: 4 } },
      children: [run("Cleo Diagnostics CSP-01-001  ·  Monthly trial summary  ·  " + p.label, { size: 7.5, bold: true, color: NAVY })] })] });
    return new D.Document({
      creator: "Lindus Health", title: "Cleo Diagnostics monthly trial summary " + p.label,
      styles: { default: { document: { run: { font: "Arial", size: 20, color: INK } } } },
      numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: D.LevelFormat.BULLET, text: "•", alignment: D.AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
      sections: [{
        properties: { titlePage: true, page: { size: { width: 12240, height: 15840 }, margin: { top: 864, bottom: 864, left: 1008, right: 1008, header: 432, footer: 432 } } },
        headers: { default: header, first: new D.Header({ children: [new D.Paragraph({ children: [] })] }) },
        footers: { default: footer(), first: footer() },
        children: children
      }]
    });
  }

  // ------------------------------------------------------------ Excel
  var XFMT = { int: "#,##0", pct: '[=0]0%;[<0.005]"<1%";0%', dec: "0.0", rate: "0.00", date: "d mmm yyyy" };
  function xValue(c, v) {
    if (v == null) return ["–", null];
    if (c.kind === "range") return [E.fmtD0(v[0]) + "–" + E.fmtD0(v[1]), null];
    if (c.kind === "date") return [new Date(v * E.DAY), XFMT.date];
    if (c.kind === "int" && !Number.isInteger(v)) return [v, "0.00"];
    return [v, XFMT[c.kind] || null];
  }
  var font = function (o) { o = o || {}; return { name: "Arial", size: o.size || 10, bold: !!o.bold, italic: !!o.italic, color: { argb: "FF" + (o.color || INK) } }; };
  var fill = function (hex) { return { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } }; };
  var thin = { style: "thin", color: { argb: "FFE1E0D9" } }, topRule = { style: "thin", color: { argb: "FFC3C2B7" } };

  function buildXlsx(result, meta, sel, charts, opts) {
    var X = root.ExcelJS, p = result.period, aud = sel.audience, one = pageOne(result, sel), secs = resolve(result, sel), h = result.head;
    opts = opts || {};
    var wb = new X.Workbook();
    wb.creator = "Lindus Health";
    var subtitle = (opts.example ? "EXAMPLE DATA · " : "") + "Reporting month " + p.label + " · Data to " + E.fmtDate2(p.lastDay) +
      (meta.citrusExport ? " · Citrus export " + E.fmtDate2(E.dayOf(meta.citrusExport)) : "") + (meta.inventoryDate ? " · Inventory report " + meta.inventoryDate : "");
    function sheet(name, title) {
      var ws = wb.addWorksheet(name, { views: [{ showGridLines: false }],
        pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
      ws.getCell("A1").value = title; ws.getCell("A1").font = font({ size: 16, bold: true, color: NAVY }); ws.getRow(1).height = 24;
      ws.getCell("A2").value = subtitle; ws.getCell("A2").font = font({ size: 9, color: INK_2 });
      return ws;
    }

    // Summary sheet
    var ws = sheet("Summary", "Cleo Diagnostics CSP-01-001 | Monthly trial summary | " + p.label);
    ws.getCell("A3").value = aud === "sponsor" ? "Sponsor version, draft for review. A person checks every figure before this goes to the sponsor."
      : "Internal version. Contains participant-level lists. Not for the sponsor.";
    ws.getCell("A3").font = font({ size: 9, bold: true, color: RED });
    var on = function (k) { return tableOn(result, sel, k); };
    var kpis = [
      ["Enrollment", "Participants enrolled", h.enrolled_month, h.enrolled_to_date, "", "funnel"],
      ["Enrollment", "Enrolled per active site", Math.round(h.per_site_month * 100) / 100, null, h.per_site_week.toFixed(2) + " per active site per week; " + h.active_sites + " active sites", "enrol_by_site_week"],
      ["Data completeness", "CRF forms complete", h.crf_pct, null, E.fmtInt(h.crf_done) + " of " + E.fmtInt(h.crf_expected) + " expected forms", "crf_by_site"],
      ["Data completeness", "ePRO forms completed", h.epro_pct, null, E.fmtInt(h.epro_done) + " of " + E.fmtInt(h.epro_assigned) + " assigned; " + h.epro_open + " open", "epro_by_form"],
      ["Post-surgical", "Surgeries complete", h.surgery_month, h.surgery_to_date, "", "psf_summary"],
      ["Post-surgical", "Pathology reports uploaded", h.reports_month, h.reports_to_date, E.fmtPct(h.reports_pct) + " of surgeries complete; median " + E.fmtD0(h.turn_median) + " days from enrollment", "psf_summary"],
      ["Blood samples", "Samples registered", h.samples_month, null, "", "blood_summary"],
      ["Blood samples", "Samples inside registration window", h.window_pct_month, h.window_pct_to_date, h.window_n_month + " of " + h.window_d_month + " in " + p.short, "blood_summary"],
      ["Data queries", "Queries raised", h.q_raised_month, h.q_total, h.q_closed_month + " closed in " + p.short, "query_summary"],
      ["Data queries", "Unresolved queries at month end", null, h.q_open + h.q_reviewing, h.q_open + " open, " + h.q_reviewing + " reviewing; median " + E.fmtD0(h.q_median_days) + " days open", "query_summary"]
    ].filter(function (k) { return on(k[5]); });
    var r = 5;
    ws.getCell(r, 1).value = "Key figures"; ws.getCell(r, 1).font = font({ size: 12, bold: true, color: NAVY }); r++;
    ["Area", "Measure", p.label, "To date", "Detail"].forEach(function (t, j) {
      var c = ws.getCell(r, j + 1); c.value = t; c.font = font({ size: 9, bold: true, color: "FFFFFF" }); c.fill = fill(BLUE);
      c.alignment = { horizontal: j === 2 || j === 3 ? "center" : "left", vertical: "middle" };
    });
    ws.getRow(r).height = 22;
    kpis.forEach(function (k, i) {
      r++;
      [k[0], k[1], k[2], k[3], k[4]].forEach(function (v, j) {
        var c = ws.getCell(r, j + 1);
        c.value = v == null ? "–" : v;
        c.font = font({ bold: j === 2 || j === 3, color: j === 2 || j === 3 ? NAVY : INK });
        c.border = { bottom: thin };
        c.alignment = { horizontal: (j === 2 || j === 3) ? (v == null ? "center" : "right") : "left", vertical: "middle", wrapText: j === 4 };
        if ((j === 2 || j === 3) && typeof v === "number" && v <= 1 && k[1].indexOf("per active site") < 0 && !Number.isInteger(v)) c.numFmt = "0%";
        else if ((j === 2 || j === 3) && Number.isInteger(v)) c.numFmt = "#,##0";
        if (i % 2) c.fill = fill(ZEBRA);
      });
      ws.getRow(r).height = 18;
    });
    r += 3;
    [["Highlights", one.highlights], ["Needs attention", one.actions]].forEach(function (blk) {
      if (!blk[1].length) return;
      ws.getCell(r, 1).value = blk[0]; ws.getCell(r, 1).font = font({ size: 12, bold: true, color: NAVY });
      blk[1].forEach(function (it) {
        r++; ws.mergeCells(r, 1, r, 5);
        var c = ws.getCell(r, 1); c.value = "•  " + it; c.font = font(); c.alignment = { wrapText: true, vertical: "top" }; ws.getRow(r).height = 28;
      });
      r += 3;
    });
    [20, 34, 14, 12, 62].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });

    // Section sheets
    var TAB = { enrollment: "Enrollment", completeness: "Data completeness", post_surgical: "Post-surgical", blood: "Blood samples", queries: "Data queries" };
    secs.forEach(function (s) {
      var ws2 = sheet(TAB[s.key], s.title);
      ws2.getCell("A3").value = s.intro; ws2.getCell("A3").font = font({ size: 9.5, color: INK_2 });
      var widths = {};
      s.tables.forEach(function (t) {
        keptCols(t, sel, false).forEach(function (ci, j) { var w = t.cols[ci].width || (j ? 12 : 30); widths[j + 1] = Math.max(widths[j + 1] || 0, w); });
      });
      Object.keys(widths).forEach(function (j) { ws2.getColumn(+j).width = Math.min(Math.max(widths[j], 11), 46); });
      var row = 4;
      var image = function (k) {  // story chart below the current row; returns the next free row
        var c = charts[k], hpx = Math.round(760 * c.h / c.w);
        ws2.addImage(wb.addImage({ buffer: c.png, extension: "png" }), { tl: { col: 0, row: row - 1 }, ext: { width: 760, height: hpx } });
        row += Math.floor(hpx / 20) + 3;
      };
      s.tables.forEach(function (t) {
        var place = chartsFor(t.key, charts, sel);
        place.before.forEach(image);
        var idx = keptCols(t, sel, false);
        ws2.getCell(row, 1).value = t.title; ws2.getCell(row, 1).font = font({ size: 12, bold: true, color: NAVY }); row++;
        if (t.note) {
          ws2.mergeCells(row, 1, row, Math.max(idx.length, 6));
          var nc = ws2.getCell(row, 1); nc.value = t.note; nc.font = font({ size: 8.5, italic: true, color: INK_2 }); nc.alignment = { wrapText: true, vertical: "top" };
          ws2.getRow(row).height = 13 * (Math.floor(t.note.length / 150) + 1); row++;
        }
        var lines = 1;
        idx.forEach(function (ci, j) {
          var c = ws2.getCell(row, j + 1); c.value = t.cols[ci].header; c.font = font({ size: 9, bold: true, color: "FFFFFF" }); c.fill = fill(BLUE);
          c.alignment = { horizontal: j ? "center" : "left", vertical: "middle", wrapText: true };
          var w = ws2.getColumn(j + 1).width || 12; lines = Math.max(lines, Math.ceil(t.cols[ci].header.length / Math.max(Math.floor(w * 0.95), 1)));
        });
        ws2.getRow(row).height = 13 * lines + 6;
        var bodyStart = row + 1;
        var rows = t.rows.length ? t.rows : null;
        if (!rows) { row++; ws2.getCell(row, 1).value = "None"; ws2.getCell(row, 1).font = font({ size: 9, italic: true, color: INK_2 }); row += 3; return; }
        rows.forEach(function (rr, i) {
          row++;
          var tot = t.totalRows.has(i);
          idx.forEach(function (ci, j) {
            var col = t.cols[ci], xv = xValue(col, rr[ci]), c = ws2.getCell(row, j + 1);
            c.value = xv[0]; if (xv[1]) c.numFmt = xv[1];
            c.font = font({ size: 9.5, bold: tot });
            c.alignment = { horizontal: col.kind === "text" ? "left" : (xv[0] === "–" ? "center" : "right"), vertical: "middle" };
            c.border = tot ? { bottom: thin, top: topRule } : { bottom: thin };
            if (tot) c.fill = fill(TOTAL); else if (i % 2) c.fill = fill(ZEBRA);
          });
        });
        var bodyRows = rows.map(function (_, i) { return i; }).filter(function (i) { return !t.totalRows.has(i); });
        idx.forEach(function (ci, j) {
          if (t.cols[ci].kind !== "pct" || !bodyRows.length) return;
          var L = ws2.getColumn(j + 1).letter;
          ws2.addConditionalFormatting({ ref: L + (bodyStart + bodyRows[0]) + ":" + L + (bodyStart + bodyRows[bodyRows.length - 1]),
            rules: [{ type: "dataBar", minLength: 0, maxLength: 100, gradient: false, border: false, cfvo: [{ type: "num", value: 0 }, { type: "num", value: 1 }], color: { argb: "FF86B6EF" } }] });
        });
        row += 3;
        place.after.forEach(image);
      });
    });

    // Definitions
    var wd = sheet("Definitions", "Definitions and sources");
    var dr = 4;
    E.definitions(p, meta, !!sel.sections.completeness).forEach(function (d) {
      var a = wd.getCell(dr, 1), b = wd.getCell(dr, 2);
      a.value = d[0]; a.font = font({ bold: true }); a.alignment = { vertical: "top" };
      b.value = d[1]; b.font = font(); b.alignment = { wrapText: true, vertical: "top" };
      wd.getRow(dr).height = 15 * (Math.floor(d[1].length / 95) + 1); dr++;
    });
    wd.getColumn(1).width = 28; wd.getColumn(2).width = 100;
    return wb;
  }

  // ------------------------------------------------------------ totals CSV (no participant rows)
  function totalsCsv(result, sel) {
    var esc = function (v) { v = String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
    var lines = ["section,table,row,column,value"];
    resolve(result, sel).forEach(function (s) {
      s.tables.forEach(function (t) {
        if (t.detail || t.chart) return;
        t.rows.forEach(function (r) {
          t.cols.slice(1).forEach(function (c, i) {
            var v = r[i + 1];
            if (Array.isArray(v)) v = v[0].toFixed(1) + "-" + v[1].toFixed(1);
            else if (typeof v === "number" && !Number.isInteger(v)) v = Math.round(v * 1e4) / 1e4;
            lines.push([s.title, t.title, r[0], c.header, v == null ? "" : v].map(esc).join(","));
          });
        });
      });
    });
    return "﻿" + lines.join("\n") + "\n";
  }

  root.TSWriters = { fmt: fmt, resolve: resolve, keptCols: keptCols, pageOne: pageOne, drawCharts: drawCharts,
    buildDocx: buildDocx, buildXlsx: buildXlsx, totalsCsv: totalsCsv };
})(typeof globalThis !== "undefined" ? globalThis : this);
