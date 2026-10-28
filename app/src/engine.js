/* Trial Summary Builder: calculation engine.
 * A line-for-line port of trialsummary/metrics.py and narrative.py, with file recognition and input checks.
 * No libraries and no DOM, so the same file runs in the browser and in the Node parity test.
 * Dates are UTC. Timestamps are milliseconds. Calendar dates are whole day numbers since 1970-01-01.
 */
(function (root) {
  "use strict";
  var DAY = 86400000;
  var MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var MONTHS_LONG = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
    "October", "November", "December"];
  var SAMPLE_WINDOW_DAYS = { CLEODX1: 14, CLEODX2: 28 };
  var NO_SAMPLE_REASON = "Unable to Obtain Study Blood Sample";
  var CRF_FORM_ORDER = ["Informed Consent", "Inclusion Criteria", "Exclusion Criteria", "Demographics", "Medical History",
    "Health Information Collection", "CA 125 Test, Pre-Surgery", "Enrollment",
    "Study Blood Sample Collection & Processing", "Post-Surgical Findings", "Study Completion"];
  var DROPPED_COLUMNS = ["Date of Birth"];

  // ------------------------------------------------------------ small helpers
  function cmp(a, b) { return a < b ? -1 : a > b ? 1 : 0; }  // code-point order, like Python sorted()
  function sortBy(arr, key) { return arr.slice().sort(function (a, b) { return cmp(key(a), key(b)); }); }
  function pct(n, d) { return d ? n / d : null; }
  function stats(values) {
    if (!values.length) return [null, null, null, null];
    var s = values.slice().sort(function (a, b) { return a - b; });
    var mean = s.reduce(function (a, b) { return a + b; }, 0) / s.length;
    var mid = Math.floor(s.length / 2);
    var med = s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
    return [mean, med, s[0], s[s.length - 1]];
  }
  function counter(list, key) {
    var c = new Map();
    list.forEach(function (x) { var k = key ? key(x) : x; c.set(k, (c.get(k) || 0) + 1); });
    return c;
  }
  function get(map, k) { return map.get(k) || 0; }
  function groupBy(list, key) {
    var g = new Map();
    list.forEach(function (x) { var k = key(x); if (!g.has(k)) g.set(k, []); g.get(k).push(x); });
    return g;
  }

  // ------------------------------------------------------------ dates
  function ts(v) {
    v = (v == null ? "" : String(v)).trim();
    if (!v || v === "NaT" || v === "nan") return null;
    var ms = Date.parse(v.replace(" ", "T"));
    return isNaN(ms) ? null : ms;
  }
  function labDate(v) {
    var m = /^(\d{1,2}) ([A-Za-z]{3}) (\d{4})$/.exec((v || "").trim());
    if (!m) return null;
    return Date.UTC(+m[3], MONTHS.indexOf(m[2].charAt(0).toUpperCase() + m[2].slice(1, 3).toLowerCase()), +m[1]) / DAY;
  }
  function isoDay(v) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec((v || "").trim());
    return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / DAY : null;
  }
  function dayOf(ms) { return Math.floor(ms / DAY); }
  function dparts(day) { var d = new Date(day * DAY); return [d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()]; }
  function fmtDay(day) { var p = dparts(day); return p[2] + " " + MONTHS[p[1]]; }
  function fmtDate(day) { var p = dparts(day); return p[2] + " " + MONTHS[p[1]] + " " + p[0]; }
  function fmtDate2(day) { var p = dparts(day); return (p[2] < 10 ? "0" : "") + p[2] + " " + MONTHS[p[1]] + " " + p[0]; }
  function ym(ms) { var d = new Date(ms); return d.getUTCFullYear() * 12 + d.getUTCMonth(); }

  function Period(month) {
    var y = +month.slice(0, 4), m = +month.slice(5, 7);
    this.key = month;
    this.start = Date.UTC(y, m - 1, 1);
    this.end = Date.UTC(y, m, 1);
    this.firstDay = this.start / DAY;
    this.lastDay = this.end / DAY - 1;
    this.year = y; this.month = m;
    this.label = MONTHS_LONG[m - 1] + " " + y;
    this.short = MONTHS[m - 1] + " " + y;
    this.days = this.lastDay - this.firstDay + 1;
    this.weeksInMonth = this.days / 7;
    this.tag = MONTHS[m - 1] + y;
  }
  Period.prototype.inMonth = function (ms) { return ms != null && ms >= this.start && ms < this.end; };
  Period.prototype.byEnd = function (ms) { return ms != null && ms < this.end; };
  Period.prototype.inMonthDay = function (d) { return d != null && d >= this.firstDay && d <= this.lastDay; };
  Period.prototype.byEndDay = function (d) { return d != null && d <= this.lastDay; };
  Period.prototype.endLabel = function () { return fmtDay(this.lastDay) + " " + this.year; };
  Period.prototype.weeks = function () {
    var wd = (new Date(this.firstDay * DAY).getUTCDay() + 6) % 7;
    var monday = this.firstDay - wd, out = [];
    while (monday <= this.lastDay) {
      var sunday = monday + 6, lo = Math.max(monday, this.firstDay), hi = Math.min(sunday, this.lastDay);
      var partial = lo !== monday || hi !== sunday;
      out.push([lo, hi, fmtDay(monday) + "–" + fmtDay(sunday) + (partial ? "*" : "")]);
      monday += 7;
    }
    return out;
  };

  // ------------------------------------------------------------ table model
  function Col(header, kind, width) { return { header: header, kind: kind || "int", width: width || null }; }
  function Table(key, title, cols, opts) {
    opts = opts || {};
    return { key: key, title: title, cols: cols, rows: [], note: opts.note || "", totalRows: new Set(), detail: !!opts.detail,
      add: function (row, total) { if (total) this.totalRows.add(this.rows.length); this.rows.push(row); } };
  }
  function Section(key, title, intro, tables) { return { key: key, title: title, intro: intro, tables: tables }; }

  // ------------------------------------------------------------ CSV parsing and file recognition
  function parseCSV(text) {
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    var rows = [], row = [], field = "", i = 0, q = false, n = text.length;
    while (i < n) {
      var c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } q = false; i++; continue; }
        field += c; i++; continue;
      }
      if (c === '"') { q = true; i++; continue; }
      if (c === ",") { row.push(field); field = ""; i++; continue; }
      if (c === "\r") { i++; continue; }
      if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
      field += c; i++;
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    return rows;
  }
  function rowsToObjects(header, rows) {
    var keep = header.map(function (h) { return DROPPED_COLUMNS.indexOf(h) < 0; });
    return rows.filter(function (r) { return r.some(function (v) { return v !== "" && v != null; }); }).map(function (r) {
      var o = {};
      header.forEach(function (h, i) { if (keep[i]) o[h] = r[i] == null ? "" : String(r[i]); });
      return o;
    });
  }

  var FILE_TYPES = [
    { key: "inventory", label: "Central lab inventory", test: function (h) { return h.indexOf("Requisition Number") >= 0; },
      need: ["Requisition Number", "LabMatrix Barcode", "Site", "Subject ID", "Date of Collection", "Date Registered", "Test Code", "Current Status"] },
    { key: "participants", label: "Participants overview", test: function (h) { return h.indexOf("Consent at") >= 0 && h.indexOf("Enrolled at") >= 0; },
      need: ["Participant ID", "Participant Status", "Site ID", "Site", "Screen failure at", "Withdrawal at", "Withdrawal reason", "Consent at", "Enrolled at", "Completed at"] },
    { key: "epro", label: "ePRO completeness", test: function (h) { return h.indexOf("num_completed") >= 0; },
      need: ["participant_id", "site_id", "form_title", "num_available", "num_expired", "num_completed", "num_skipped", "supp_num_overdue"] },
    { key: "crf", label: "CRF completeness", test: function (h) { return h.indexOf("form_variable") >= 0 && h.indexOf("status") >= 0; },
      need: ["participant_id", "site_id", "form_title", "status"] },
    { key: "queries", label: "Data queries", test: function (h) { return h.indexOf("Query content") >= 0; },
      need: ["Participant ID", "Site ID", "Created At", "Updated at", "Type", "Status", "Form Title"] },
    { key: "psf", label: "Post-surgical findings", test: function (h) { return h.some(function (x) { return x.indexOf("(Surgery Date)") >= 0; }); },
      need: ["Participant ID", "Form Status", "Completed At", "*(Surgery Date)", "*(Pathology Report)"] },
    { key: "inclusion", label: "Inclusion criteria", test: function (h) { return h.some(function (x) { return x.indexOf("(IC#1)") >= 0; }); },
      need: ["Participant ID", "Form Status"] },
    { key: "exclusion", label: "Exclusion criteria", test: function (h) { return h.some(function (x) { return x.indexOf("(EC#1)") >= 0; }); },
      need: ["Participant ID", "Form Status"] }
  ];
  function identify(header) {
    for (var i = 0; i < FILE_TYPES.length; i++) if (FILE_TYPES[i].test(header)) return FILE_TYPES[i];
    return null;
  }
  function missingColumns(type, header) {
    return type.need.filter(function (c) {
      if (c.charAt(0) === "*") { var part = c.slice(1); return !header.some(function (h) { return h.indexOf(part) >= 0; }); }
      return header.indexOf(c) < 0;
    });
  }
  function exportTimeFromName(name) {
    var m = /(\d{4}-\d{2}-\d{2})-(\d{2})-(\d{2})-UTC/.exec(name || "");
    return m ? Date.parse(m[1] + "T" + m[2] + ":" + m[3] + ":00Z") : null;
  }
  function inventoryDateDay(text) {
    var m = /(\d{2})([A-Z]{3})(\d{4})/.exec((text || "").toUpperCase());
    if (!m) return null;
    var mi = MONTHS.map(function (x) { return x.toUpperCase(); }).indexOf(m[2]);
    return mi < 0 ? null : Date.UTC(+m[3], mi, +m[1]) / DAY;
  }

  // ------------------------------------------------------------ shared participant view
  function siteActive(s, period) {
    if (s.activation_date && isoDay(s.activation_date) > period.lastDay) return false;
    if (s.terminated === "Yes") return !!s.termination_date && isoDay(s.termination_date) >= period.firstDay;
    return true;
  }
  function prepare(data, period) {
    var sites = new Map(data.sites.map(function (s) { return [s.site_id, s]; }));
    var people = data.participants.map(function (r) {
      var p = { pid: r["Participant ID"], site_id: r["Site ID"], site: r["Site"], consent: ts(r["Consent at"]),
        enrolled: ts(r["Enrolled at"]), completed: ts(r["Completed at"]), sf: ts(r["Screen failure at"]),
        wd: ts(r["Withdrawal at"]), wd_reason: r["Withdrawal reason"] || "" };
      var e = function (x) { return period.byEnd(x); };
      p.status_eom = e(p.completed) ? "Completed" : e(p.wd) ? "Withdrawn" : e(p.sf) ? "Screen failed" :
        e(p.enrolled) ? "Enrolled, active" : e(p.consent) ? "Consented, not yet enrolled" : null;
      return p;
    });
    var active = data.sites.filter(function (s) { return siteActive(s, period); });
    return { sites: sites, active_sites: active, people: people };
  }
  function siteName(ctx, id) { var s = ctx.sites.get(id); return s ? s.site : (id || "(no site)"); }

  // ------------------------------------------------------------ 1. enrollment
  function enrollment(data, ctx, period) {
    var people = ctx.people, m = function (x) { return period.inMonth(x); };
    var consented = people.filter(function (p) { return period.byEnd(p.consent); });
    var enrolled = people.filter(function (p) { return period.byEnd(p.enrolled); });

    var funnel = Table("funnel", "Recruitment summary", [Col("Measure", "text", 40), Col(period.label), Col("To date (to " + period.endLabel() + ")")],
      { note: "Counts use the date of each event. Citrus holds no pre-screening date, so pre-screening is not shown." });
    [["Consented", "consent"], ["Enrolled", "enrolled"], ["Screen failures", "sf"], ["Withdrawals", "wd"], ["Completed study", "completed"]]
      .forEach(function (x) {
        funnel.add([x[0], people.filter(function (p) { return m(p[x[1]]); }).length,
          people.filter(function (p) { return period.byEnd(p[x[1]]); }).length]);
      });

    var conv = Table("conversion", "Consent to enrollment", [Col("Measure", "text", 40), Col("n"), Col("Denominator", "text", 28), Col("%", "pct")]);
    var cMonth = people.filter(function (p) { return m(p.consent); });
    var n = cMonth.filter(function (p) { return p.enrolled != null && period.byEnd(p.enrolled); }).length;
    conv.add(["Consented in " + period.short + " and enrolled by month end", n, cMonth.length + " consented in " + period.short, pct(n, cMonth.length)]);
    conv.add(["Enrolled to date", enrolled.length, consented.length + " consented to date", pct(enrolled.length, consented.length)]);
    var sf = people.filter(function (p) { return period.byEnd(p.sf); }).length;
    conv.add(["Screen failures to date", sf, consented.length + " consented to date", pct(sf, consented.length)]);

    var status = Table("status_eom", "Participant status at month end (" + period.endLabel() + ")",
      [Col("Status", "text", 40), Col("Participants"), Col("% of consented", "pct")],
      { note: "Denominator: " + consented.length + " participants consented to date." });
    var counts = counter(consented, function (p) { return p.status_eom; });
    ["Consented, not yet enrolled", "Enrolled, active", "Completed", "Withdrawn", "Screen failed"].forEach(function (s) {
      status.add([s, get(counts, s), pct(get(counts, s), consented.length)]);
    });
    status.add(["Total consented", consented.length, pct(consented.length, consented.length)], true);

    var weeks = period.weeks();
    var bySite = Table("enrol_by_site_week", "Enrollments by site and week, " + period.label,
      [Col("Site", "text", 38), Col("Type", "text", 10)].concat(weeks.map(function (w) { return Col(w[2]); })).concat(
        [Col(period.short + " total"), Col("Per week", "rate"), Col("Enrolled to date"), Col("Months active", "dec"), Col("Per month since activation", "rate")]),
      { note: "Weeks run Monday to Sunday. * Part week: only days inside the month count. Per week = month total ÷ " +
        period.weeksInMonth.toFixed(2) + " weeks. Per month since activation = enrolled to date ÷ months since the site activation date. " +
        "Activation dates and site types come from the site lookup table." });
    var siteEnrol = new Map();
    enrolled.forEach(function (p) { if (!siteEnrol.has(p.site_id)) siteEnrol.set(p.site_id, []); siteEnrol.get(p.site_id).push(dayOf(p.enrolled)); });
    var activeIds = new Set(ctx.active_sites.map(function (s) { return s.site_id; }));
    var shown = sortBy(data.sites.filter(function (s) { return activeIds.has(s.site_id) || (siteEnrol.get(s.site_id) || []).length; }),
      function (s) { return s.site; });
    var totWeek = weeks.map(function () { return 0; });
    shown.forEach(function (s) {
      var ds = siteEnrol.get(s.site_id) || [];
      var wk = weeks.map(function (w) { return ds.filter(function (d) { return d >= w[0] && d <= w[1]; }).length; });
      totWeek = totWeek.map(function (a, i) { return a + wk[i]; });
      var monthN = ds.filter(function (d) { return period.inMonthDay(d); }).length;
      var monthsActive = s.activation_date ? (period.lastDay - isoDay(s.activation_date)) / 30.4375 : null;
      var label = s.site + (activeIds.has(s.site_id) ? "" : " (terminated)");
      bySite.add([label, s.site_type].concat(wk).concat([monthN, monthN / period.weeksInMonth, ds.length, monthsActive,
        monthsActive ? ds.length / monthsActive : null]));
    });
    var nActive = ctx.active_sites.length;
    var monthTotal = totWeek.reduce(function (a, b) { return a + b; }, 0);
    bySite.add(["All sites", ""].concat(totWeek).concat([monthTotal, monthTotal / period.weeksInMonth, enrolled.length, null, null]), true);
    bySite.add(["Average per active site (" + nActive + " sites)", ""].concat(totWeek.map(function (w) { return w / nActive; }))
      .concat([monthTotal / nActive, monthTotal / period.weeksInMonth / nActive, enrolled.length / nActive, null, null]), true);

    var byType = Table("enrol_by_type", "Enrollment by site type",
      [Col("Group", "text", 22), Col("Active sites"), Col("Enrolled " + period.short), Col("Per site per month", "rate"),
        Col("Per site per week", "rate"), Col("Enrolled to date"), Col("% of all enrolled to date", "pct")],
      { note: "Denominator for %: " + enrolled.length + " participants enrolled to date. Month rates use active sites only. " +
        "Enrolled to date includes terminated sites. Site types come from the CLEO-001 site list." });
    var types = [];
    data.sites.forEach(function (s) { if (types.indexOf(s.site_type) < 0) types.push(s.site_type); });
    types.map(function (t) { return [t, "site_type", t]; }).concat([["Academic", "academic_status", "Academic"],
      ["Non-academic", "academic_status", "Non-academic"]]).forEach(function (g) {
      var ids = new Set(ctx.active_sites.filter(function (s) { return s[g[1]] === g[2]; }).map(function (s) { return s.site_id; }));
      var allIds = new Set(data.sites.filter(function (s) { return s[g[1]] === g[2]; }).map(function (s) { return s.site_id; }));
      var mo = enrolled.filter(function (p) { return ids.has(p.site_id) && m(p.enrolled); }).length;
      var td = enrolled.filter(function (p) { return allIds.has(p.site_id); }).length;
      byType.add([g[0], ids.size, mo, pct(mo, ids.size), ids.size ? mo / ids.size / period.weeksInMonth : null, td, pct(td, enrolled.length)]);
    });

    var trend = Table("enrol_trend", "Enrollments by month", [Col("Month", "text", 12), Col("Consented"), Col("Enrolled"), Col("Cumulative enrolled")]);
    var first = enrolled.length ? Math.min.apply(null, enrolled.map(function (p) { return p.enrolled; })) : period.start;
    var cum = 0, eBy = counter(enrolled, function (p) { return ym(p.enrolled); }), cBy = counter(consented, function (p) { return ym(p.consent); });
    for (var k = ym(first); k <= ym(period.start); k++) {
      cum += get(eBy, k);
      trend.add([MONTHS[k % 12] + " " + Math.floor(k / 12), get(cBy, k), get(eBy, k), cum]);
    }
    return Section("enrollment", "Enrollment", "Enrollment activity in " + period.label + " and cumulative to month end.",
      [funnel, conv, status, bySite, byType, trend]);
  }

  // ------------------------------------------------------------ 2. completeness
  function completeness(data, ctx, period) {
    function crfRow(label, rows) {
      var expected = rows.filter(function (r) { return r.status !== "Unavailable"; });
      var c = counter(expected, function (r) { return r.status; });
      var done = get(c, "Complete") + get(c, "Locked");
      return [label, expected.length, done, pct(done, expected.length), get(c, "Available"), get(c, "Open Query"), get(c, "Skipped")];
    }
    function cols(first) {
      return [Col(first, "text", 38), Col("Forms expected"), Col("Complete"), Col("% complete", "pct"), Col("Outstanding"), Col("Open query"), Col("Skipped")];
    }
    var note = "Forms expected = all forms open to the participant (status not 'Unavailable'). Complete = 'Complete' or 'Locked'. " +
      "Outstanding = 'Available', not yet completed.";
    var bySite = Table("crf_by_site", "CRF completeness by site", cols("Site"), { note: note });
    var g = groupBy(data.crf, function (r) { return r.site_id; });
    sortBy(Array.from(g.keys()), function (s) { return siteName(ctx, s); }).forEach(function (sid) { bySite.add(crfRow(siteName(ctx, sid), g.get(sid))); });
    bySite.add(crfRow("All sites", data.crf), true);

    var byForm = Table("crf_by_form", "CRF completeness by form", cols("Form"), { note: note });
    var f = groupBy(data.crf, function (r) { return r.form_title; });
    Array.from(f.keys()).sort(function (a, b) {
      var ia = CRF_FORM_ORDER.indexOf(a), ib = CRF_FORM_ORDER.indexOf(b);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
    }).forEach(function (t) { byForm.add(crfRow(t, f.get(t))); });
    byForm.add(crfRow("All forms", data.crf), true);

    function eproRow(label, rows) {
      var t = { num_available: 0, num_expired: 0, num_completed: 0, num_skipped: 0, supp_num_overdue: 0 };
      rows.forEach(function (r) { Object.keys(t).forEach(function (k) { t[k] += parseInt(r[k] || "0", 10) || 0; }); });
      var a = t.num_available + t.num_expired + t.num_completed + t.num_skipped;
      return [label, a, t.num_completed, pct(t.num_completed, a), t.num_available, pct(t.num_available, a), t.num_expired,
        pct(t.num_expired, a), t.num_skipped, pct(t.num_skipped, a), t.supp_num_overdue];
    }
    function ecols(first) {
      return [Col(first, "text", 38), Col("Assigned"), Col("Completed"), Col("% completed", "pct"), Col("Open"), Col("% open", "pct"),
        Col("Expired"), Col("% expired", "pct"), Col("Skipped"), Col("% skipped", "pct"), Col("Overdue (open)")];
    }
    var enote = "Denominator: ePRO forms assigned. Open = still available to the participant. Overdue is a subset of open.";
    var eForm = Table("epro_by_form", "ePRO completeness by form", ecols("Form"), { note: enote });
    var ef = groupBy(data.epro, function (r) { return r.form_title; });
    Array.from(ef.keys()).sort(cmp).forEach(function (t) { eForm.add(eproRow(t, ef.get(t))); });
    eForm.add(eproRow("All ePRO forms", data.epro), true);
    var eSite = Table("epro_by_site", "ePRO completeness by site", ecols("Site"), { note: enote });
    var es = groupBy(data.epro, function (r) { return r.site_id; });
    sortBy(Array.from(es.keys()), function (s) { return siteName(ctx, s); }).forEach(function (sid) { eSite.add(eproRow(siteName(ctx, sid), es.get(sid))); });
    eSite.add(eproRow("All sites", data.epro), true);

    var snap = data.meta.citrusExport;
    return Section("completeness", "Data completeness",
      "Completeness is a snapshot from the Citrus export" + (snap ? " on " + fmtDate2(dayOf(snap)) : "") +
      ". Citrus does not date form status changes, so month-end values cannot be rebuilt.", [bySite, byForm, eForm, eSite]);
  }

  // ------------------------------------------------------------ 3. post-surgical findings
  function postSurgical(data, ctx, period) {
    var psf = new Map();
    if (data.psf.length) {
      var keys = Object.keys(data.psf[0]);
      var sdCol = keys.filter(function (k) { return k.indexOf("(Surgery Date)") >= 0; })[0];
      var rptCol = keys.filter(function (k) { return k.indexOf("(Pathology Report)") >= 0; })[0];
      data.psf.forEach(function (r) {
        var rep = r[rptCol] || "";
        var hasReport = rep.indexOf("permanent/") === 0 || rep.indexOf("temporary/") === 0;
        psf.set(r["Participant ID"], { surgery: r[sdCol] ? isoDay(r[sdCol]) : null,
          upload: r["Form Status"] === "COMPLETE" && hasReport ? ts(r["Completed At"]) : null });
      });
    }
    var enrolled = ctx.people.filter(function (p) { return period.byEnd(p.enrolled); });
    var rows = enrolled.map(function (p) {
      var f = psf.get(p.pid) || {};
      var surgery = period.byEndDay(f.surgery) ? f.surgery : null;
      var upload = period.byEnd(f.upload) ? f.upload : null;
      return Object.assign({}, p, { surgery: surgery, upload: upload, days: upload ? (upload - p.enrolled) / DAY : null });
    });
    var surg = rows.filter(function (r) { return r.surgery != null; });
    var upl = rows.filter(function (r) { return r.upload != null; });
    var waiting = surg.filter(function (r) { return r.upload == null; });

    var summ = Table("psf_summary", "Surgery and pathology reports",
      [Col("Measure", "text", 40), Col(period.label), Col("To date"), Col("Denominator (to date)", "text", 30), Col("% (to date)", "pct")],
      { note: "Surgery complete = a surgery date is entered. Report uploaded = Post-Surgical Findings form complete with a report file." });
    summ.add(["Participants enrolled", rows.filter(function (r) { return period.inMonth(r.enrolled); }).length, rows.length, "", null]);
    summ.add(["Surgery complete", surg.filter(function (r) { return period.inMonthDay(r.surgery); }).length, surg.length,
      rows.length + " enrolled", pct(surg.length, rows.length)]);
    summ.add(["Pathology report uploaded", upl.filter(function (r) { return period.inMonth(r.upload); }).length, upl.length,
      surg.length + " with surgery complete", pct(upl.length, surg.length)]);
    summ.add(["Awaiting pathology report", null, waiting.length, surg.length + " with surgery complete", pct(waiting.length, surg.length)]);

    var turn = Table("psf_turnaround", "Days from enrollment to pathology report upload",
      [Col("Reports uploaded", "text", 40), Col("Reports"), Col("Mean days", "dec"), Col("Median days", "dec"), Col("Shortest", "dec"), Col("Longest", "dec")],
      { note: "Upload date = 'Completed At' on the Post-Surgical Findings form." });
    [["In " + period.label, upl.filter(function (r) { return period.inMonth(r.upload); })], ["To date", upl]].forEach(function (x) {
      turn.add([x[0], x[1].length].concat(stats(x[1].map(function (r) { return r.days; }))));
    });

    var bySite = Table("psf_by_site", "Surgery and pathology reports by site",
      [Col("Site", "text", 38), Col("Enrolled"), Col("Surgery complete"), Col("% of enrolled", "pct"), Col("Report uploaded"),
        Col("% of surgeries", "pct"), Col("Awaiting report"), Col("Median days to upload", "dec"), Col("Mean days to upload", "dec")]);
    var g = groupBy(rows, function (r) { return r.site_id; });
    sortBy(Array.from(g.keys()), function (s) { return siteName(ctx, s); }).forEach(function (sid) {
      var rs = g.get(sid), s_ = rs.filter(function (r) { return r.surgery != null; }), u_ = rs.filter(function (r) { return r.upload != null; });
      var st = stats(u_.map(function (r) { return r.days; }));
      bySite.add([siteName(ctx, sid), rs.length, s_.length, pct(s_.length, rs.length), u_.length, pct(u_.length, s_.length), s_.length - u_.length, st[1], st[0]]);
    });
    var st = stats(upl.map(function (r) { return r.days; }));
    bySite.add(["All sites", rows.length, surg.length, pct(surg.length, rows.length), upl.length, pct(upl.length, surg.length), waiting.length, st[1], st[0]], true);

    var wl = Table("psf_awaiting", "Participants awaiting a pathology report",
      [Col("Participant", "text", 12), Col("Site", "text", 38), Col("Enrolled", "date"), Col("Surgery date", "date"), Col("Days since surgery")],
      { detail: true, note: "Days since surgery counted to " + period.endLabel() + "." });
    waiting.slice().sort(function (a, b) { return a.surgery - b.surgery; }).forEach(function (r) {
      wl.add([r.pid, siteName(ctx, r.site_id), dayOf(r.enrolled), r.surgery, period.lastDay - r.surgery]);
    });
    return Section("post_surgical", "Post-surgical findings", "Surgeries, pathology report uploads and turnaround for enrolled participants.",
      [summ, turn, bySite, wl]);
  }

  // ------------------------------------------------------------ 4. blood samples
  function bloodSamples(data, ctx, period) {
    var samples = [];
    data.inventory.forEach(function (r) {
      var coll = labDate(r["Date of Collection"]), reg = labDate(r["Date Registered"]);
      if (!period.byEndDay(reg) || coll == null) return;
      var code = r["Test Code"], limit = SAMPLE_WINDOW_DAYS[code];
      if (!limit) return;
      samples.push({ req: r["Requisition Number"], barcode: r["LabMatrix Barcode"], pid: r["Subject ID"], site: r["Site"], code: code,
        coll: coll, reg: reg, days: reg - coll, limit: limit, inWindow: reg - coll <= limit, status: r["Current Status"] });
    });
    var codeToSite = new Map(data.sites.map(function (s) { return [s.lab_site_code, s.site]; }));
    var siteOf = function (code) { return codeToSite.has(code) ? codeToSite.get(code) : code; };
    var reqs = new Map();
    samples.forEach(function (s) { if (!reqs.has(s.req)) reqs.set(s.req, {}); reqs.get(s.req)[s.code] = s; });
    var reqList = Array.from(reqs.values());
    var both = function (v) { return v.CLEODX1 && v.CLEODX2; };
    var sameDay = reqList.filter(function (v) { return both(v) && v.CLEODX1.reg === v.CLEODX2.reg; });

    function windowRows(subset, samePairs) {
      var d1 = subset.filter(function (s) { return s.code === "CLEODX1"; }), d2 = subset.filter(function (s) { return s.code === "CLEODX2"; });
      var ok = function (xs) { return xs.filter(function (s) { return s.inWindow; }).length; };
      return [
        ["Samples registered (DX1 + DX2)", subset.length, null, null],
        ["DX1 registered within 14 days of collection", ok(d1), d1.length, pct(ok(d1), d1.length)],
        ["DX2 registered within 28 days of collection", ok(d2), d2.length, pct(ok(d2), d2.length)],
        ["All samples within their window", ok(subset), subset.length, pct(ok(subset), subset.length)],
        ["Samples outside their window", subset.length - ok(subset), subset.length, pct(subset.length - ok(subset), subset.length)],
        ["DX1 and DX2 registered on the same day (pairs)", samePairs.length, null, null]
      ];
    }
    var month = samples.filter(function (s) { return period.inMonthDay(s.reg); });
    var monthSame = sameDay.filter(function (v) { return period.inMonthDay(v.CLEODX1.reg); });
    var summ = Table("blood_summary", "Sample registration at the central lab",
      [Col("Measure", "text", 46), Col(period.short + " n"), Col(period.short + " denominator"), Col(period.short + " %", "pct"),
        Col("To date n"), Col("To date denominator"), Col("To date %", "pct")],
      { note: period.label + " counts samples registered in the month. Each DX1 must be registered within 14 days of collection and each DX2 within 28 days. " +
        "DX1 and DX2 must ship separately, so a same-day registration is a deviation." });
    var a = windowRows(month, monthSame), b = windowRows(samples, sameDay);
    a.forEach(function (x, i) { summ.add([x[0], x[1], x[2], x[3], b[i][1], b[i][2], b[i][3]]); });

    var enrolled = ctx.people.filter(function (p) { return period.byEnd(p.enrolled); });
    var pidReqs = new Map();
    samples.forEach(function (s) { if (!pidReqs.has(s.pid)) pidReqs.set(s.pid, []); if (pidReqs.get(s.pid).indexOf(s.req) < 0) pidReqs.get(s.pid).push(s.req); });
    var pairs = Table("blood_pairs", "Sample pair status at month end", [Col("Status", "text", 46), Col("Participants"), Col("% of enrolled", "pct")],
      { note: "Denominator: " + enrolled.length + " participants enrolled to date. Overdue = the registration window for the missing sample has passed. " +
        "Pending = still inside the window." });
    var missingList = Table("blood_missing_pair", "Samples missing their pair",
      [Col("Requisition", "text", 12), Col("Participant", "text", 12), Col("Site", "text", 34), Col("Collected", "date"), Col("Registered", "text", 10),
        Col("Missing", "text", 10), Col("Days since collection"), Col("Status", "text", 10)], { detail: true });
    var noSample = Table("blood_no_sample", "Enrolled participants with no sample registered",
      [Col("Participant", "text", 12), Col("Site", "text", 34), Col("Enrolled", "date"), Col("Days since enrollment"), Col("Status", "text", 34)], { detail: true });
    var c = new Map(), inc = function (k) { c.set(k, get(c, k) + 1); };
    enrolled.forEach(function (p) {
      var list = pidReqs.get(p.pid);
      if (!list || !list.length) {
        var age = period.lastDay - dayOf(p.enrolled), st;
        if (p.wd_reason === NO_SAMPLE_REASON) { inc("No sample: withdrawn, sample not obtained"); st = "Withdrawn: sample not obtained"; }
        else if (age > SAMPLE_WINDOW_DAYS.CLEODX1) { inc("No sample registered: overdue"); st = "Overdue"; }
        else { inc("No sample registered: pending"); st = "Pending"; }
        noSample.add([p.pid, siteName(ctx, p.site_id), dayOf(p.enrolled), age, st]);
        return;
      }
      list.forEach(function (req) {
        var got = reqs.get(req);
        if (both(got)) { inc("Both samples registered"); return; }
        var have = got.CLEODX1 || got.CLEODX2, missing = have.code === "CLEODX1" ? "CLEODX2" : "CLEODX1";
        var age2 = period.lastDay - have.coll, overdue = age2 > SAMPLE_WINDOW_DAYS[missing];
        inc(missing.slice(-3) + " missing: " + (overdue ? "overdue" : "pending"));
        missingList.add([req, p.pid, siteOf(have.site), have.coll, have.code.slice(-3), missing.slice(-3), age2, overdue ? "Overdue" : "Pending"]);
      });
    });
    ["Both samples registered", "DX2 missing: pending", "DX2 missing: overdue", "DX1 missing: pending", "DX1 missing: overdue",
      "No sample registered: pending", "No sample registered: overdue", "No sample: withdrawn, sample not obtained"].forEach(function (l) {
      pairs.add([l, get(c, l), pct(get(c, l), enrolled.length)]);
    });
    pairs.add(["Total enrolled", enrolled.length, pct(enrolled.length, enrolled.length)], true);
    missingList.rows.sort(function (x, y) { return ((x[7] !== "Overdue") - (y[7] !== "Overdue")) || (y[6] - x[6]); });

    var bySite = Table("blood_by_site", "Sample registration by site, to date",
      [Col("Site", "text", 34), Col("Samples registered"), Col("DX1 registered"), Col("DX1 within 14 days"), Col("% DX1 in window", "pct"),
        Col("DX2 registered"), Col("DX2 within 28 days"), Col("% DX2 in window", "pct"), Col("Same-day pairs"), Col("Missing pair: overdue")]);
    var overdueBySite = counter(missingList.rows.filter(function (r) { return r[7] === "Overdue"; }), function (r) { return r[2]; });
    var sameBySite = counter(sameDay, function (v) { return codeToSite.get(v.CLEODX1.site); });
    function siteRow(name, xs, same, od) {
      var d1 = xs.filter(function (s) { return s.code === "CLEODX1"; }), d2 = xs.filter(function (s) { return s.code === "CLEODX2"; });
      var o1 = d1.filter(function (s) { return s.inWindow; }).length, o2 = d2.filter(function (s) { return s.inWindow; }).length;
      return [name, xs.length, d1.length, o1, pct(o1, d1.length), d2.length, o2, pct(o2, d2.length), same, od];
    }
    var gs = groupBy(samples, function (s) { return siteOf(s.site); });
    Array.from(gs.keys()).sort(cmp).forEach(function (name) { bySite.add(siteRow(name, gs.get(name), get(sameBySite, name), get(overdueBySite, name))); });
    var odTotal = 0; overdueBySite.forEach(function (v) { odTotal += v; });
    bySite.add(siteRow("All sites", samples, sameDay.length, odTotal), true);

    var trend = Table("blood_trend", "Sample registration by month",
      [Col("Month", "text", 12), Col("Registered"), Col("Within window"), Col("Outside window"), Col("% within window", "pct")],
      { note: "Samples registered at the central lab in each month, last 12 months. Denominator for %: samples registered that month." });
    var endK = period.year * 12 + period.month - 1;
    for (var k = endK - 11; k <= endK; k++) {
      var xs = samples.filter(function (s) { var d = new Date(s.reg * DAY); return d.getUTCFullYear() * 12 + d.getUTCMonth() === k; });
      var okN = xs.filter(function (s) { return s.inWindow; }).length;
      trend.add([MONTHS[k % 12] + " " + Math.floor(k / 12), xs.length, okN, xs.length - okN, pct(okN, xs.length)]);
    }

    var late = Table("blood_out_of_window", "Samples registered outside their window, " + period.label,
      [Col("Barcode", "text", 18), Col("Participant", "text", 12), Col("Site", "text", 34), Col("Collected", "date"), Col("Registered", "date"),
        Col("Days to register"), Col("Limit (days)")], { detail: true });
    month.filter(function (s) { return !s.inWindow; }).sort(function (x, y) { return y.days - x.days; }).forEach(function (s) {
      late.add([s.barcode, s.pid, siteOf(s.site), s.coll, s.reg, s.days, s.limit]);
    });
    var sameList = Table("blood_same_day", "Pairs registered on the same day, to date",
      [Col("Requisition", "text", 12), Col("Participant", "text", 12), Col("Site", "text", 34), Col("Collected", "date"), Col("Registered", "date")], { detail: true });
    sameDay.slice().sort(function (x, y) { return y.CLEODX1.reg - x.CLEODX1.reg; }).forEach(function (v) {
      var s = v.CLEODX1; sameList.add([s.req, s.pid, siteOf(s.site), s.coll, s.reg]);
    });
    var loc = Table("blood_location", "Where registered samples are now", [Col("Current status", "text", 30), Col("Samples"), Col("% of registered", "pct")],
      { note: "Status from the inventory report dated " + (data.meta.inventoryDate || "(unknown)") + ". Denominator: " + samples.length + " samples registered to month end." });
    var lc = counter(samples, function (s) { return s.status; });
    ["In Inventory", "Transferred", "Destroyed"].forEach(function (k) { loc.add([k, get(lc, k), pct(get(lc, k), samples.length)]); });
    return Section("blood", "Blood samples", "Central lab registration from the LabConnect inventory report dated " + (data.meta.inventoryDate || "(unknown)") + ".",
      [summ, trend, pairs, bySite, loc, missingList, noSample, late, sameList]);
  }

  // ------------------------------------------------------------ 5. data queries
  function queries(data, ctx, period) {
    var qs = [];
    data.queries.forEach(function (r) {
      var created = ts(r["Created At"]), updated = ts(r["Updated at"]);
      if (!period.byEnd(created)) return;
      var closedByEnd = r.Status === "Closed" && period.byEnd(updated);
      var state = closedByEnd ? "Closed" : (r.Status === "Reviewing" && period.byEnd(updated)) ? "Reviewing" : "Open";
      qs.push({ site_id: r["Site ID"], type: r.Type, form: r["Form Title"], state: state, created: created,
        closedInMonth: closedByEnd && period.inMonth(updated), daysOpen: state === "Closed" ? null : (period.end - created) / DAY });
    });
    var total = qs.length, st = counter(qs, function (q) { return q.state; });
    var summ = Table("query_summary", "Query status at month end", [Col("Measure", "text", 40), Col("Queries"), Col("% of all queries raised", "pct")],
      { note: "Denominator: " + total + " queries raised to month end. Reviewing = the site responded and a reviewer is checking it." });
    ["Open", "Reviewing", "Closed"].forEach(function (k) { summ.add([k, get(st, k), pct(get(st, k), total)]); });
    summ.add(["Total raised to date", total, pct(total, total)], true);
    summ.add(["Raised in " + period.label, qs.filter(function (q) { return period.inMonth(q.created); }).length, null]);
    summ.add(["Closed in " + period.label, qs.filter(function (q) { return q.closedInMonth; }).length, null]);
    summ.add(["Open or reviewing for more than 90 days", qs.filter(function (q) { return (q.daysOpen || 0) > 90; }).length, null]);

    var days = Table("query_days_open", "Days open for unresolved queries",
      [Col("Status at month end", "text", 40), Col("Queries"), Col("Mean days open", "dec"), Col("Median days open", "dec"), Col("Shortest", "dec"), Col("Longest", "dec")],
      { note: "Days open = days from creation to the end of " + period.endLabel() + "." });
    [["Open", ["Open"]], ["Reviewing", ["Reviewing"]], ["Open and reviewing", ["Open", "Reviewing"]]].forEach(function (x) {
      var xs = qs.filter(function (q) { return x[1].indexOf(q.state) >= 0; }).map(function (q) { return q.daysOpen; });
      days.add([x[0], xs.length].concat(stats(xs)), x[0] === "Open and reviewing");
    });

    function grouped(field, key, title, label, namer) {
      var t = Table("query_by_" + key, title, [Col(label, "text", 38), Col("Total"), Col("Open"), Col("% open", "pct"), Col("Reviewing"),
        Col("% reviewing", "pct"), Col("Closed"), Col("% closed", "pct"), Col("Mean days open", "dec"), Col("Median days open", "dec"),
        Col("Range of days open", "range", 14)], { note: "Percentages use the group's total queries. Days open covers open and reviewing queries." });
      var g = groupBy(qs, function (q) { return q[field]; });
      function row(name, xs) {
        var c = counter(xs, function (q) { return q.state; });
        var s = stats(xs.filter(function (q) { return q.daysOpen != null; }).map(function (q) { return q.daysOpen; }));
        return [name, xs.length, get(c, "Open"), pct(get(c, "Open"), xs.length), get(c, "Reviewing"), pct(get(c, "Reviewing"), xs.length),
          get(c, "Closed"), pct(get(c, "Closed"), xs.length), s[0], s[1], s[2] != null ? [s[2], s[3]] : null];
      }
      sortBy(Array.from(g.keys()), namer).forEach(function (k) { t.add(row(namer(k), g.get(k))); });
      t.add(row("All", qs), true);
      return t;
    }
    var id = function (k) { return k; };
    return Section("queries", "Data queries", "Query status at month end, built from each query's creation and last-update dates.",
      [summ, days, grouped("site_id", "site", "Queries by site", "Site", function (s) { return siteName(ctx, s); }),
        grouped("type", "type", "Queries by type", "Query type", id), grouped("form", "form", "Queries by form", "Form", id)]);
  }

  // ------------------------------------------------------------ headline, narrative, definitions
  function headline(sections, period, nActive) {
    var T = {};
    Object.keys(sections).forEach(function (k) { sections[k].tables.forEach(function (t) { T[t.key] = t; }); });
    var by = function (key) { var o = {}; T[key].rows.forEach(function (r) { o[r[0]] = r; }); return o; };
    var funnel = by("funnel"), psf = by("psf_summary"), blood = by("blood_summary"), pairs = by("blood_pairs"), qsum = by("query_summary");
    var crfAll = T.crf_by_site.rows[T.crf_by_site.rows.length - 1], eproAll = T.epro_by_form.rows[T.epro_by_form.rows.length - 1];
    var turn = T.psf_turnaround.rows[T.psf_turnaround.rows.length - 1], qdays = T.query_days_open.rows[T.query_days_open.rows.length - 1];
    var avg = T.enrol_by_site_week.rows[T.enrol_by_site_week.rows.length - 1], L = avg.length;
    var inWin = blood["All samples within their window"];
    return {
      enrolled_month: funnel.Enrolled[1], enrolled_to_date: funnel.Enrolled[2], consented_month: funnel.Consented[1],
      consented_to_date: funnel.Consented[2], completed_month: funnel["Completed study"][1],
      per_site_month: avg[L - 5], per_site_week: avg[L - 4], active_sites: nActive,
      crf_pct: crfAll[3], crf_done: crfAll[2], crf_expected: crfAll[1],
      epro_pct: eproAll[3], epro_done: eproAll[2], epro_assigned: eproAll[1], epro_open: eproAll[4],
      surgery_to_date: psf["Surgery complete"][2], surgery_month: psf["Surgery complete"][1],
      reports_to_date: psf["Pathology report uploaded"][2], reports_month: psf["Pathology report uploaded"][1],
      reports_pct: psf["Pathology report uploaded"][4], awaiting: psf["Awaiting pathology report"][2],
      turn_median: turn[3], turn_mean: turn[2], turn_min: turn[4], turn_max: turn[5],
      samples_month: blood["Samples registered (DX1 + DX2)"][1], window_pct_month: inWin[3], window_n_month: inWin[1],
      window_d_month: inWin[2], window_pct_to_date: inWin[6], same_day_month: blood["DX1 and DX2 registered on the same day (pairs)"][1],
      missing_overdue: pairs["DX2 missing: overdue"][1] + pairs["DX1 missing: overdue"][1],
      no_sample_overdue: pairs["No sample registered: overdue"][1],
      q_open: qsum.Open[1], q_reviewing: qsum.Reviewing[1], q_total: qsum["Total raised to date"][1],
      q_raised_month: qsum["Raised in " + period.label][1], q_closed_month: qsum["Closed in " + period.label][1],
      q_median_days: qdays[3], q_mean_days: qdays[2], q_over_90: qsum["Open or reviewing for more than 90 days"][1],
      completed_to_date: funnel["Completed study"][2],
      enrolled_prev_month: T.enrol_trend.rows.length >= 2 ? T.enrol_trend.rows[T.enrol_trend.rows.length - 2][2] : null,
      late_month: T.blood_out_of_window.rows.length,
      path_over_60: T.psf_awaiting.rows.filter(function (r) { return r[r.length - 1] > 60; }).length,
      samples_to_date: blood["Samples registered (DX1 + DX2)"][4],
      crf_outstanding: crfAll[4], crf_open_query: crfAll[5]
    };
  }

  // Page-1 'Inside this report' rows and the key-figure strip that opens each section (mirror of narrative.digest).
  var SECTION_ORDER = ["enrollment", "completeness", "post_surgical", "blood", "queries"];
  function digest(sections, h, period, included) {
    var specs = chartSpecs(sections, period, h), m = period.short, p = fmtPct, n = fmtInt;
    var completedPct = h.enrolled_to_date ? h.completed_to_date / h.enrolled_to_date : null;
    var findings = {
      enrollment: specs.enrollment_trend.title,
      completeness: p(h.crf_pct) + " of expected CRF forms and " + p(h.epro_pct) + " of assigned ePRO forms are complete",
      post_surgical: "Pathology reports are " + p(h.reports_pct) + " complete; " + h.path_over_60 + " of " + h.awaiting + " still awaited are more than 60 days past surgery",
      blood: specs.sample_timeliness.title,
      queries: specs.queries_by_site.title
    };
    var figures = {
      enrollment: [[n(h.enrolled_month), "Enrolled in " + m, n(h.consented_month) + " consented in " + m],
        [n(h.enrolled_to_date), "Enrolled to date", n(h.consented_to_date) + " consented to date"],
        [n(h.completed_to_date), "Completed the study", p(completedPct) + " of enrolled"],
        [String(h.active_sites), "Active sites", h.per_site_month.toFixed(2) + " enrolled per site in " + m]],
      completeness: [[p(h.crf_pct), "CRF forms complete", n(h.crf_done) + " of " + n(h.crf_expected) + " expected"],
        [p(h.epro_pct), "ePRO forms completed", n(h.epro_done) + " of " + n(h.epro_assigned) + " assigned"],
        [n(h.crf_outstanding), "CRF forms outstanding", "open, not yet complete"],
        [n(h.crf_open_query), "CRF forms with an open query", "snapshot at export"]],
      post_surgical: [[n(h.surgery_month), "Surgeries in " + m, n(h.surgery_to_date) + " to date"],
        [n(h.reports_month), "Reports uploaded in " + m, n(h.reports_to_date) + " to date"],
        [p(h.reports_pct), "Reports uploaded", "of " + n(h.surgery_to_date) + " surgeries complete"],
        [fmtD0(h.turn_median) + " days", "Median enrollment to upload", "range " + fmtD0(h.turn_min) + "–" + fmtD0(h.turn_max) + " days"]],
      blood: [[n(h.samples_month), "Samples registered in " + m, n(h.samples_to_date) + " to date"],
        [p(h.window_pct_month), "Inside registration window", n(h.window_n_month) + " of " + n(h.window_d_month) + " in " + m],
        [n(h.missing_overdue + h.no_sample_overdue), "Participants missing samples", "past the registration deadline"],
        [n(h.same_day_month), "Pairs shipped together", "registered on the same day in " + m]],
      queries: [[n(h.q_open + h.q_reviewing), "Unresolved queries", n(h.q_open) + " open, " + n(h.q_reviewing) + " reviewing"],
        [n(h.q_raised_month), "Raised in " + m, n(h.q_closed_month) + " closed in " + m],
        [fmtD0(h.q_median_days) + " days", "Median days open", "open and reviewing queries"],
        [n(h.q_over_90), "Open more than 90 days", "open or in review"]]
    };
    var keys = SECTION_ORDER.filter(function (k) { return included.indexOf(k) >= 0; }), figs = {};
    keys.forEach(function (k) { figs[k] = figures[k]; });
    return { inside: keys.map(function (k) { return [sections[k].title, findings[k]]; }), figures: figs };
  }

  // Python-style rounding for display: exact ties go to the even neighbour, like format(x, ".0f").
  function roundHalfEven(x, nd) {
    var f = Math.pow(10, nd || 0), v = x * f, r = Math.round(v);
    if (Math.abs(v % 1) === 0.5) r = 2 * Math.round(v / 2);
    return r / f;
  }
  function fmtPct(x) { if (x == null) return "–"; if (x > 0 && x < 0.005) return "<1%"; return roundHalfEven(x * 100, 0).toFixed(0) + "%"; }
  function fmtInt(x) { return Math.round(x).toLocaleString("en-US"); }
  function fmtD0(x) { return x == null ? "–" : roundHalfEven(x, 0).toFixed(0); }

  // Sponsor text: high-level progress and site-facing items only. Mirrors sponsor_highlights/sponsor_actions in narrative.py.
  function prevMonthLabel(period) { var k = period.year * 12 + period.month - 2; return MONTHS_LONG[k % 12] + " " + Math.floor(k / 12); }
  function sponsorHighlights(h, period) {
    var m = period.label, prev = h.enrolled_prev_month, trend = "";
    if (prev != null) trend = h.enrolled_month > prev ? ", up from " + prev + " in " + prevMonthLabel(period)
      : h.enrolled_month < prev ? ", down from " + prev + " in " + prevMonthLabel(period) : ", the same as in " + prevMonthLabel(period);
    var completedPct = h.enrolled_to_date ? h.completed_to_date / h.enrolled_to_date : null;
    return [
      { key: "funnel", text: h.enrolled_month + " participants enrolled in " + m + trend + ". Total enrollment is now " + fmtInt(h.enrolled_to_date) + "." },
      { key: "enrol_by_site_week", text: h.active_sites + " sites are active and enrolled an average of " + roundHalfEven(h.per_site_month, 1).toFixed(1) + " participants each in " + m + "." },
      { key: "status_eom", text: fmtInt(h.completed_to_date) + " participants have completed the study, " + fmtPct(completedPct) + " of those enrolled." },
      { key: "psf_summary", text: "Pathology reports are " + fmtPct(h.reports_pct) + " complete: " + fmtInt(h.reports_to_date) + " of " + fmtInt(h.surgery_to_date) +
        " participants with surgery have a report uploaded." },
      { key: "blood_summary", text: fmtPct(h.window_pct_month) + " of blood samples registered in " + m + " reached the central lab within the required window." }
    ];
  }
  function sponsorActions(h, period) {
    var m = period.label, out = [];
    if (h.late_month || h.same_day_month) {
      var parts = [];
      if (h.late_month) parts.push(h.late_month + " of " + h.samples_month + " samples registered in " + m + " arrived outside the required window");
      if (h.same_day_month) parts.push(h.same_day_month + " sample pairs were shipped together instead of separately");
      out.push({ key: "blood_summary", text: "Sample shipping: " + parts.join(", and ") + ". Sites should ship each sample promptly after collection." });
    }
    var missing = h.missing_overdue + h.no_sample_overdue;
    if (missing) out.push({ key: "blood_pairs", text: "Missing samples: " + missing + " participants are missing one or both blood samples past the registration deadline." });
    if (h.path_over_60) out.push({ key: "psf_summary", text: "Pathology reports: " + h.path_over_60 + " participants are more than 60 days past surgery without a report uploaded." });
    return out;
  }

  function highlights(h, period, completeness) {
    var m = period.label;
    var out = [
      h.enrolled_month + " participants enrolled in " + m + ", and " + h.consented_month + " consented. " + h.enrolled_to_date +
        " participants are enrolled to date across " + h.active_sites + " active sites.",
      "The " + m + " enrollment rate was " + h.per_site_month.toFixed(2) + " per active site per month, or " + h.per_site_week.toFixed(2) + " per active site per week.",
      "CRF completeness is " + fmtPct(h.crf_pct) + ": " + fmtInt(h.crf_done) + " of " + fmtInt(h.crf_expected) + " expected forms are complete. " +
        "ePRO completion is " + fmtPct(h.epro_pct) + ": " + fmtInt(h.epro_done) + " of " + fmtInt(h.epro_assigned) + " assigned forms, with " + fmtInt(h.epro_open) + " still open.",
      "Pathology reports are uploaded for " + h.reports_to_date + " of " + h.surgery_to_date + " participants with surgery complete (" + fmtPct(h.reports_pct) +
        "). The median time from enrollment to upload is " + fmtD0(h.turn_median) + " days.",
      fmtPct(h.window_pct_month) + " of samples registered in " + m + " were inside their registration window (" + h.window_n_month + " of " + h.window_d_month + ").",
      (h.q_open + h.q_reviewing) + " queries were unresolved at month end: " + h.q_open + " open and " + h.q_reviewing + " in review. " +
        "The median time open is " + fmtD0(h.q_median_days) + " days."
    ];
    if (completeness === false) out = out.filter(function (x) { return x.indexOf("CRF completeness") !== 0; });
    return out;
  }
  function actions(h, sections, period) {
    var T = {};
    Object.keys(sections).forEach(function (k) { sections[k].tables.forEach(function (t) { T[t.key] = t; }); });
    var out = [];
    if (h.missing_overdue) out.push({ key: "pairs", text: h.missing_overdue + " participants have a sample with no pair registered after the window closed." });
    if (h.no_sample_overdue) out.push({ key: "nosample", text: h.no_sample_overdue + " enrolled participants have no sample registered at the central lab after 14 days." });
    var late = T.blood_out_of_window.rows.length;
    if (late) out.push({ key: "late", text: late + " samples registered in " + period.label + " were outside their registration window." });
    if (h.same_day_month) out.push({ key: "sameday", text: h.same_day_month + " DX1 and DX2 pairs were registered on the same day in " + period.label + "." });
    var longWait = T.psf_awaiting.rows.filter(function (r) { return r[r.length - 1] > 60; }).length;
    if (longWait) out.push({ key: "path", text: longWait + " participants are still waiting for a pathology report more than 60 days after surgery." });
    if (h.q_over_90) out.push({ key: "old", text: h.q_over_90 + " queries have been open or in review for more than 90 days." });
    return out;
  }
  function definitions(period, meta, completeness) {
    var lastLong = new Date(period.lastDay * DAY).getUTCDate() + " " + MONTHS_LONG[period.month - 1] + " " + period.year;
    var exp = meta.citrusExport ? new Date(meta.citrusExport) : null;
    var expText = exp ? fmtDate2(dayOf(meta.citrusExport)) + " " + ("0" + exp.getUTCHours()).slice(-2) + ":" + ("0" + exp.getUTCMinutes()).slice(-2) + " UTC" : "(date unknown)";
    var out = [
      ["Reporting month", period.label + ". Month counts use events dated 1 to " + lastLong + " (UTC)."],
      ["To date", "All events up to the end of " + lastLong + "."],
      ["Sources", "Citrus exports from " + expText + ". LabConnect inventory report dated " + (meta.inventoryDate || "(unknown)") +
        ". CLEO-001 site list for site type, academic status and activation date."],
      ["Weeks", "Monday to Sunday. Part weeks at the start or end of the month count only days inside the month."],
      ["Enrollment rate", "Per site per week = enrollments in the month ÷ " + period.weeksInMonth.toFixed(2) + " weeks. Per site per month = enrollments in the month. " +
        "The 'per active site' rates divide by the number of sites activated by month end and not terminated. Per month since activation = enrolled to date ÷ months since activation."],
      ["CRF completeness", "Complete or Locked forms ÷ forms expected. Forms expected excludes forms marked 'Unavailable', which are not yet open to the participant. Snapshot at the Citrus export time."],
      ["ePRO completeness", "Each assigned ePRO form is Completed, Open, Expired or Skipped. Percentages use forms assigned. Overdue forms are open forms past their due date."],
      ["Surgery complete", "A surgery date is entered on the Post-Surgical Findings form."],
      ["Pathology report uploaded", "The Post-Surgical Findings form is complete with a report file. The upload date is the form's 'Completed At' time."],
      ["Sample registration window", "DX1 within 14 days of collection. DX2 within 28 days of collection. DX1 and DX2 must be registered on different days because they ship separately."],
      ["Missing pair", "One sample of the pair is registered and the other is not. Overdue when the missing sample's window has passed."],
      ["Blood sample month view", "Counts samples registered at the central lab in the month."],
      ["Query status at month end", "Closed if closed by month end. Reviewing if the last update before month end moved it to review. Open otherwise."],
      ["Days open", "Days from query creation to the end of the month, for open and reviewing queries only."]
    ];
    if (completeness === false) out = out.filter(function (d) { return d[0] !== "CRF completeness" && d[0] !== "ePRO completeness"; });
    return out;
  }

  // ------------------------------------------------------------ story chart specs (mirror of trialsummary/charts.py chart_specs)
  var CH = { CONTEXT: "#c3c2b7", ACCENT: "#2a78d6", CATEGORICAL: ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"],
    ORDINAL: ["#86b6ef", "#3987e5", "#1c5cab", "#104281"] };
  function monthLabels(labels) {
    var l1 = [], l2 = [];
    labels.forEach(function (lab, i) { var p = lab.split(" "); l1.push(p[0]); l2.push(i === 0 || p[0] === "Jan" ? "'" + p[1].slice(2) : ""); });
    return [l1, l2];
  }
  function chartSpecs(sections, period, h) {
    var T = {};
    Object.keys(sections).forEach(function (k) { sections[k].tables.forEach(function (t) { T[t.key] = t; }); });
    var m = period.label, specs = {}, prevL = prevMonthLabel(period);
    var body = function (t) { return t.rows.filter(function (_, i) { return !t.totalRows.has(i); }); };

    var rows = T.enrol_trend.rows.slice(-12), vals = rows.map(function (r) { return r[2]; });
    var n = vals[vals.length - 1], prev = vals.length >= 2 ? vals[vals.length - 2] : null;
    var trend = prev == null ? "" : n > prev ? ", up from " + prev + " in " + prevL : n < prev ? ", down from " + prev + " in " + prevL : ", the same as in " + prevL;
    var ml = monthLabels(rows.map(function (r) { return r[0]; })), last = vals.length - 1, vt = {};
    vt[String(last)] = String(n);
    specs.enrollment_trend = { key: "enrollment_trend", kind: "columns", after: "enrol_trend", replaces_table: true,
      title: n + " participants enrolled in " + m + trend, subtitle: "Participants enrolled per month, last 12 months",
      labels: ml[0], labels2: ml[1], values: vals, colors: vals.map(function (_, i) { return i === last ? CH.ACCENT : CH.CONTEXT; }),
      label_idx: [last], value_text: vt };

    var t = T.enrol_by_site_week, mcol = t.cols.map(function (c) { return c.header; }).indexOf(period.short + " total");
    var sr = body(t).filter(function (r) { return r[mcol] > 0; }).sort(function (a, b) { return (b[mcol] - a[mcol]) || cmp(a[0], b[0]); });
    var strip = function (s) { return s.replace(" (terminated)", ""); };
    specs.enrollment_by_site = { key: "enrollment_by_site", kind: "hbars", before: "enrol_by_site_week",
      title: sr.length ? sr.length + " of " + h.active_sites + " active sites enrolled participants in " + m + "; " + strip(sr[0][0]) + " led with " + sr[0][mcol]
        : "No participants enrolled in " + m,
      subtitle: "Participants enrolled in " + m + ", by site",
      labels: sr.map(function (r) { return strip(r[0]); }), values: sr.map(function (r) { return r[mcol]; }), colors: sr.map(function () { return CH.ACCENT; }) };

    var byStatus = {};
    body(T.status_eom).forEach(function (r) { byStatus[r[0]] = r; });
    var st = ["Completed", "Enrolled, active", "Consented, not yet enrolled", "Withdrawn", "Screen failed"].filter(function (k) { return byStatus[k]; }).map(function (k) { return byStatus[k]; });
    var consented = st.reduce(function (a, r) { return a + r[1]; }, 0), completed = byStatus.Completed ? byStatus.Completed[1] : 0;
    specs.participant_status = { key: "participant_status", kind: "stack100", after: "status_eom",
      title: fmtInt(completed) + " of " + fmtInt(consented) + " consented participants have completed the study",
      subtitle: "Where each consented participant stands at month end (" + new Date(period.lastDay * DAY).getUTCDate() + " " + period.short + ")",
      labels: st.map(function (r) { return r[0]; }), values: st.map(function (r) { return r[1]; }), colors: CH.CATEGORICAL.slice(0, st.length),
      legend_text: st.map(function (r) { return r[0] + " · " + fmtInt(r[1]) + " (" + fmtPct(r[2]) + ")"; }) };

    var days = T.psf_awaiting.rows.map(function (r) { return r[r.length - 1]; });
    var bands = [["0–30 days", 0, 30], ["31–60 days", 31, 60], ["61–90 days", 61, 90], ["Over 90 days", 91, 1e9]];
    var counts = bands.map(function (b) { return days.filter(function (d) { return d >= b[1] && d <= b[2]; }).length; });
    var vt4 = {};
    counts.forEach(function (c, i) { vt4[String(i)] = String(c); });
    specs.pathology_backlog = { key: "pathology_backlog", kind: "columns", after: "psf_summary",
      title: days.length ? (counts[2] + counts[3]) + " of " + days.length + " reports still awaited are more than 60 days past surgery" : "No pathology reports are awaited",
      subtitle: "Participants with surgery complete and no pathology report, by days since surgery",
      labels: bands.map(function (b) { return b[0]; }), labels2: ["", "", "", ""], values: counts, colors: CH.ORDINAL, label_idx: [0, 1, 2, 3], value_text: vt4 };

    var tr = T.blood_trend.rows, ml5 = monthLabels(tr.map(function (r) { return r[0]; })), cur = tr[tr.length - 1], prv = tr.length >= 2 ? tr[tr.length - 2] : null;
    var comp = prv && prv[1] ? ", compared with " + fmtPct(prv[4]) + " in " + prevL : "", vt5 = {};
    vt5[String(tr.length - 1)] = fmtPct(cur[4]);
    specs.sample_timeliness = { key: "sample_timeliness", kind: "stacked_columns", after: "blood_trend", replaces_table: true,
      title: cur[1] ? fmtPct(cur[4]) + " of samples registered in " + m + " met the registration window" + comp : "No samples were registered in " + m,
      subtitle: "Samples registered per month at the central lab", labels: ml5[0], labels2: ml5[1],
      series: [{ name: "Within window", color: CH.CATEGORICAL[0], values: tr.map(function (r) { return r[2]; }) },
        { name: "Outside window", color: CH.CATEGORICAL[1], values: tr.map(function (r) { return r[3]; }) }],
      label_idx: [tr.length - 1], value_text: vt5 };

    var q = body(T.query_by_site).filter(function (r) { return r[2] + r[4] > 0; }).sort(function (a, b) { return ((b[2] + b[4]) - (a[2] + a[4])) || cmp(a[0], b[0]); });
    var total = q.reduce(function (a, r) { return a + r[2] + r[4]; }, 0), lead = q[0];
    var qSub = "Open and reviewing queries by site";
    if (q.length > 10) {  // top 10 sites keeps the scale honest; the rest are named in the subtitle
      var rest = q.slice(10);
      q = q.slice(0, 10);
      qSub = "Open and reviewing queries, the 10 sites with the most; " + rest.length + " other sites have " +
        fmtInt(rest.reduce(function (a, r) { return a + r[2] + r[4]; }, 0)) + " more";
    }
    specs.queries_by_site = { key: "queries_by_site", kind: "stacked_hbars", after: "query_summary",
      title: q.length ? fmtInt(total) + " queries were unresolved at month end; " + lead[0] + " had the most (" + (lead[2] + lead[4]) + ")" : "No queries were unresolved at month end",
      subtitle: qSub, labels: q.map(function (r) { return r[0]; }),
      series: [{ name: "Open", color: CH.CATEGORICAL[0], values: q.map(function (r) { return r[2]; }) },
        { name: "Reviewing", color: CH.CATEGORICAL[1], values: q.map(function (r) { return r[4]; }) }] };
    return specs;
  }

  function build(data, month) {
    var period = new Period(month);
    var ctx = prepare(data, period);
    var list = [enrollment(data, ctx, period), completeness(data, ctx, period), postSurgical(data, ctx, period),
      bloodSamples(data, ctx, period), queries(data, ctx, period)];
    var sections = {};
    list.forEach(function (s) { sections[s.key] = s; });
    return { period: period, sections: sections, head: headline(sections, period, ctx.active_sites.length), ctx: ctx };
  }

  // ------------------------------------------------------------ input checks
  function monthOptions(data) {
    var exp = data.meta.citrusExport;
    var times = data.participants.map(function (r) { return ts(r["Enrolled at"]); }).filter(function (x) { return x != null; });
    if (!times.length) return { options: [], def: null };
    var first = Math.min.apply(null, times), lastK = exp ? ym(exp) - 1 : Math.max.apply(null, times.map(ym));
    var out = [];
    for (var k = lastK; k >= ym(first); k--) {
      var y = Math.floor(k / 12), m = k % 12 + 1;
      out.push({ value: y + "-" + (m < 10 ? "0" : "") + m, label: MONTHS_LONG[m - 1] + " " + y });
    }
    return { options: out, def: out.length ? out[0].value : null };
  }

  function checks(data, files, period) {
    var out = [];
    var need = FILE_TYPES.map(function (t) { return t.key; });
    var missing = need.filter(function (k) { return !files[k]; });
    out.push(missing.length
      ? { level: "bad", title: missing.length + " of 8 files missing", text: "Add: " + missing.map(function (k) { return FILE_TYPES.filter(function (t) { return t.key === k; })[0].label; }).join(", ") + "." }
      : { level: "ok", title: "All 8 files found", text: "Each file matched by its column layout." });
    var badCols = need.filter(function (k) { return files[k] && files[k].missingCols.length; });
    out.push(badCols.length
      ? { level: "bad", title: "Columns missing in " + badCols.length + " file(s)", text: badCols.map(function (k) { return files[k].label + ": " + files[k].missingCols.join(", "); }).join(". ") + ". Download the export again from Citrus or the lab email." }
      : { level: "ok", title: "Columns match the expected layout", text: "No renamed or missing columns in any export." });
    need.forEach(function (k) { if (files[k] && files[k].duplicates) out.push({ level: "warn", title: "Two " + files[k].label + " files added", text: "The tool uses " + files[k].name + ", the most recent export." }); });
    if (missing.length || badCols.length || !period) return out;

    var exp = data.meta.citrusExport, inv = data.meta.inventoryDay;
    if (!exp) out.push({ level: "warn", title: "Export date not found", text: "The Citrus file names carry no export date. Check that the exports are from after " + period.endLabel() + "." });
    else if (exp < period.end) out.push({ level: "bad", title: "Exports dated before month end", text: "Citrus exports are dated " + fmtDate2(dayOf(exp)) + ", before the end of " + period.label + ". Export again after the month ends." });
    else out.push({ level: "ok", title: "Exports dated after month end", text: "Citrus exports dated " + fmtDate2(dayOf(exp)) + (inv != null ? ", inventory report dated " + fmtDate2(inv) : "") + ", after " + period.endLabel() + "." });
    if (inv != null && exp && Math.abs(inv - dayOf(exp)) > 3) out.push({ level: "warn", title: "Inventory report and Citrus exports are from different days", text: "Inventory report " + fmtDate2(inv) + ", Citrus " + fmtDate2(dayOf(exp)) + ". Sample counts and participant counts may not line up." });
    if (inv != null && inv <= period.lastDay) out.push({ level: "bad", title: "Inventory report dated before month end", text: "Use an inventory report from after " + period.endLabel() + "." });

    var known = new Set(data.sites.map(function (s) { return s.site_id; }));
    var dataSites = new Set(data.participants.map(function (r) { return r["Site ID"]; }).filter(Boolean));
    var unknown = Array.from(dataSites).filter(function (s) { return !known.has(s); });
    out.push(unknown.length
      ? { level: "warn", title: unknown.length + " site(s) not in the site list", text: unknown.join(", ") + ". Their participants count in totals but have no site type." }
      : { level: "ok", title: "Every site is in the site list", text: dataSites.size + " sites in the data. All are in the CLEO-001 site list of " + data.sites.length + " sites." });
    var labCodes = new Set(data.sites.map(function (s) { return s.lab_site_code; }));
    var unknownLab = Array.from(new Set(data.inventory.map(function (r) { return r.Site; }))).filter(function (c) { return c && !labCodes.has(c); });
    if (unknownLab.length) out.push({ level: "warn", title: "Lab site codes not in the site list", text: unknownLab.join(", ") + ". These samples show under the code, not a site name." });

    var pids = new Set(data.participants.map(function (r) { return r["Participant ID"]; }));
    var orphans = data.inventory.filter(function (r) { return !pids.has(r["Subject ID"]); });
    out.push(orphans.length
      ? { level: "warn", title: fmtInt(orphans.length) + " samples have no matching Citrus participant", text: "Their Subject ID is blank or not in the participants export. They count in sample totals but not in pair status." }
      : { level: "ok", title: "Every sample matches a Citrus participant", text: "All " + fmtInt(data.inventory.length) + " samples carry a Subject ID found in the participants export." });
    if (data.meta.droppedDob) out.push({ level: "info", title: "Date of birth removed", text: "The inventory report's Date of Birth column was dropped as the file opened. It never reaches the summary." });

    var after = function (rows, col, parse) { return rows.filter(function (r) { var v = parse(r[col]); return v != null && v >= period.end; }).length; };
    var lateEnrol = after(data.participants, "Enrolled at", ts), lateCons = after(data.participants, "Consent at", ts);
    var lateReg = data.inventory.filter(function (r) { var d = labDate(r["Date Registered"]); return d != null && d > period.lastDay; }).length;
    if (lateEnrol + lateCons + lateReg) out.push({ level: "info", title: "Events after " + period.endLabel() + " are left out",
      text: [lateCons ? lateCons + " consents" : "", lateEnrol ? lateEnrol + " enrollments" : "", lateReg ? lateReg + " sample registrations" : ""].filter(Boolean).join(", ") + " fall after month end and do not count." });
    return out;
  }

  root.TSEngine = {
    DAY: DAY, MONTHS: MONTHS, FILE_TYPES: FILE_TYPES, DROPPED_COLUMNS: DROPPED_COLUMNS,
    parseCSV: parseCSV, rowsToObjects: rowsToObjects, identify: identify, missingColumns: missingColumns,
    exportTimeFromName: exportTimeFromName, inventoryDateDay: inventoryDateDay,
    Period: Period, build: build, highlights: highlights, actions: actions, definitions: definitions,
    sponsorHighlights: sponsorHighlights, sponsorActions: sponsorActions, chartSpecs: chartSpecs, digest: digest,
    monthOptions: monthOptions, checks: checks, fmtPct: fmtPct, fmtInt: fmtInt, fmtD0: fmtD0, fmtDate: fmtDate, fmtDate2: fmtDate2,
    roundHalfEven: roundHalfEven, dayOf: dayOf
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
