/* Trial Summary Builder: page behaviour. Reads files in the browser, runs the engine, renders checks,
 * toggles and the page-1 preview, and hands finished files to the viewer through the downloads capability. */
(function () {
  "use strict";
  var E = window.TSEngine, M = window.TSModel, W = window.TSWriters;
  var SITES = window.TS_SITES || [], EXAMPLES = window.TS_EXAMPLES || [];
  var STORE_KEY = "tsb-choices-v1";
  var app = { files: {}, unknown: [], example: false, month: null, data: null, period: null, result: null, checks: [], state: null, downloads: null, inViewer: false };
  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); };

  // ------------------------------------------------------------ saved choices (per browser, never data)
  function loadSaved(aud) { try { var s = JSON.parse(localStorage.getItem(STORE_KEY) || "{}"); return s[aud] || null; } catch (e) { return null; } }
  function saveChoices() {
    try { var s = JSON.parse(localStorage.getItem(STORE_KEY) || "{}"); s[app.state.audience] = app.state; localStorage.setItem(STORE_KEY, JSON.stringify(s));
      $("choices-status").textContent = "Saved for the " + (app.state.audience === "sponsor" ? "Sponsor" : "Internal") + " audience in this browser."; }
    catch (e) { $("choices-status").textContent = "This browser does not allow saving choices."; }
  }
  function stateFor(aud) {
    var d = M.defaultState(aud), s = loadSaved(aud);
    if (!s) return d;
    ["groups", "items", "cols"].forEach(function (k) { Object.keys(d[k]).forEach(function (x) { if (s[k] && typeof s[k][x] === "boolean") d[k][x] = s[k][x]; }); });
    return d;
  }
  app.state = stateFor("sponsor");

  // ------------------------------------------------------------ reading files
  function cellText(v) {
    if (v == null) return "";
    if (v instanceof Date) return ("0" + v.getUTCDate()).slice(-2) + " " + E.MONTHS[v.getUTCMonth()] + " " + v.getUTCFullYear();
    if (typeof v === "object") {
      if (v.richText) return v.richText.map(function (r) { return r.text; }).join("");
      if ("result" in v) return cellText(v.result);
      if (v.text != null) return String(v.text);
      return "";
    }
    return String(v);
  }
  function locate(rows) {
    for (var i = 0; i < Math.min(rows.length, 12); i++) {
      var h = (rows[i] || []).map(function (s) { return String(s == null ? "" : s).trim(); });
      var t = E.identify(h);
      if (t) return { type: t, header: h, body: rows.slice(i + 1), pre: rows.slice(0, i) };
    }
    return null;
  }
  function readOne(name, buf) {
    var lower = name.toLowerCase();
    if (lower.endsWith(".xlsx")) {
      var wb = new window.ExcelJS.Workbook();
      return wb.xlsx.load(buf).then(function () {
        var ws = wb.worksheets[0], rows = [];
        if (!ws) return null;
        ws.eachRow({ includeEmpty: true }, function (row, n) {
          var vals = [];
          for (var c = 1; c <= ws.columnCount; c++) vals.push(cellText(row.getCell(c).value));
          rows[n - 1] = vals;
        });
        for (var i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
        return locate(rows);
      });
    }
    var text = new TextDecoder("utf-8").decode(buf);
    return Promise.resolve(locate(E.parseCSV(text)));
  }
  function addBuffers(list) {
    status("Reading " + list.length + (list.length === 1 ? " file…" : " files…"));
    var jobs = list.map(function (f) {
      return readOne(f.name, f.buf).then(function (found) { return { f: f, found: found }; }, function (err) { return { f: f, error: err }; });
    });
    return Promise.all(jobs).then(function (results) {
      results.forEach(function (r) {
        if (!r.found) { app.unknown.push(r.f.name + (r.error ? " (could not be read)" : "")); return; }
        var t = r.found.type, header = r.found.header;
        var preText = r.found.pre.map(function (row) { return row.join(" "); }).join(" ");
        var entry = { key: t.key, label: t.label, name: r.f.name, header: header, rows: E.rowsToObjects(header, r.found.body),
          missingCols: E.missingColumns(t, header), exportTime: E.exportTimeFromName(r.f.name),
          droppedDob: E.DROPPED_COLUMNS.some(function (c) { return header.indexOf(c) >= 0; }) };
        if (t.key === "inventory") {
          var m = /Date:\s*(\d{2}[A-Za-z]{3}\d{4})/.exec(preText) || /_(\d{2}[A-Z]{3}\d{4})/.exec(r.f.name);
          entry.inventoryDate = m ? m[1].toUpperCase() : null;
          entry.inventoryDay = entry.inventoryDate ? E.inventoryDateDay(entry.inventoryDate) : null;
        }
        var prev = app.files[t.key];
        if (prev) {
          var newer = (entry.exportTime || 0) >= (prev.exportTime || 0) ? entry : prev;
          newer.duplicates = true;
          app.files[t.key] = newer;
        } else app.files[t.key] = entry;
      });
      status(null);
      refresh();
    });
  }
  function addFileList(fileList) {
    if (app.example) clearFiles();
    var arr = Array.prototype.slice.call(fileList);
    Promise.all(arr.map(function (f) { return f.arrayBuffer().then(function (b) { return { name: f.name, buf: b }; }); })).then(addBuffers);
  }
  function loadExample() {
    clearFiles();
    status("Loading example data…");
    return Promise.all(EXAMPLES.map(function (p) {
      // Artifacts cannot serve .xlsx files, so the example inventory report travels as base64 inside JSON.
      if (/\.b64\.json$/.test(p)) {
        return fetch(encodeURI(p)).then(function (r) { if (!r.ok) throw new Error(p); return r.json(); }).then(function (j) {
          var bin = atob(j.base64), bytes = new Uint8Array(bin.length);
          for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          return { name: j.name, buf: bytes.buffer };
        });
      }
      return fetch(encodeURI(p)).then(function (r) { if (!r.ok) throw new Error(p); return r.arrayBuffer(); })
        .then(function (b) { return { name: p.split("/").pop(), buf: b }; });
    })).then(function (list) { app.example = true; return addBuffers(list); })
      .catch(function () { status("The example data could not be loaded. Add your own files instead.", true); });
  }
  function clearFiles() {
    app.files = {}; app.unknown = []; app.example = false; app.result = null; app.data = null;
    refresh();
  }
  function status(msg, bad) {
    var el = $("file-status");
    el.hidden = !msg; el.textContent = msg || ""; el.className = "notice" + (bad ? " bad" : "");
  }

  // ------------------------------------------------------------ compute
  function refresh() {
    var f = app.files, all = E.FILE_TYPES.every(function (t) { return f[t.key]; });
    var data = { sites: SITES, meta: {} };
    E.FILE_TYPES.forEach(function (t) { data[t.key] = f[t.key] ? f[t.key].rows : []; });
    var citrus = ["participants", "crf", "epro", "queries", "psf", "inclusion", "exclusion"].map(function (k) { return f[k] && f[k].exportTime; }).filter(Boolean);
    data.meta.citrusExport = f.participants && f.participants.exportTime ? f.participants.exportTime : (citrus.length ? Math.max.apply(null, citrus) : null);
    if (f.inventory) { data.meta.inventoryDate = f.inventory.inventoryDate; data.meta.inventoryDay = f.inventory.inventoryDay; data.meta.droppedDob = f.inventory.droppedDob; }
    app.data = data;

    var sel = $("f-month");
    var opts = f.participants && !f.participants.missingCols.length ? E.monthOptions(data) : { options: [], def: null };
    if (opts.options.length) {
      if (!opts.options.some(function (o) { return o.value === app.month; })) app.month = opts.def;
      sel.innerHTML = opts.options.map(function (o) { return '<option value="' + o.value + '"' + (o.value === app.month ? " selected" : "") + ">" + esc(o.label) + "</option>"; }).join("");
      sel.disabled = false;
    } else { app.month = null; sel.innerHTML = "<option>Add files first</option>"; sel.disabled = true; }
    app.period = app.month ? new E.Period(app.month) : null;

    app.checks = Object.keys(f).length ? E.checks(data, f, app.period) : [];
    var blocked = app.checks.some(function (c) { return c.level === "bad"; });
    app.result = null;
    if (all && app.period && !blocked) {
      try { app.result = E.build(data, app.month); }
      catch (err) { app.checks.push({ level: "bad", title: "The files could not be summarised", text: String(err && err.message || err) }); }
    }
    renderAll();
  }

  // ------------------------------------------------------------ render
  var LABEL = { ok: "Ready", warn: "Note", bad: "Fix", info: "Info", wait: "Waiting" };
  var CHECK_LABEL = { ok: "Pass", warn: "Note", bad: "Fix", info: "Info" };
  function renderFiles() {
    var rows = E.FILE_TYPES.map(function (t) {
      var e = app.files[t.key];
      if (!e) return "<tr><td>" + esc(t.label) + "</td><td class='muted'>Not added yet</td><td class='num-r'>–</td><td>–</td><td><span class='chip wait'>Waiting</span></td></tr>";
      var lvl = e.missingCols.length ? "bad" : e.duplicates ? "warn" : "ok";
      var dated = e.key === "inventory" ? (e.inventoryDay != null ? E.fmtDate2(e.inventoryDay) : "–") : (e.exportTime ? E.fmtDate2(E.dayOf(e.exportTime)) : "–");
      return "<tr><td>" + esc(t.label) + "</td><td class='mono'>" + esc(e.name) + "</td><td class='num-r'>" + E.fmtInt(e.rows.length) +
        "</td><td>" + dated + "</td><td><span class='chip " + lvl + "'>" + (e.missingCols.length ? "Fix" : e.duplicates ? "Replaced" : LABEL.ok) + "</span></td></tr>";
    });
    app.unknown.forEach(function (n) { rows.push("<tr><td class='muted'>Not recognised</td><td class='mono'>" + esc(n) + "</td><td class='num-r'>–</td><td>–</td><td><span class='chip warn'>Ignored</span></td></tr>"); });
    $("file-rows").innerHTML = rows.join("");
    $("example-banner").hidden = !app.example;
  }
  function renderChecks() {
    var el = $("check-list");
    if (!app.checks.length) { el.innerHTML = "<li class='empty-note' style='display:block'>Checks appear here once you add files.</li>"; return; }
    el.innerHTML = app.checks.map(function (c) {
      return "<li><span><span class='chip " + c.level + "'>" + CHECK_LABEL[c.level] + "</span></span><span><b>" + esc(c.title) + "</b><br><span class='muted'>" + esc(c.text) + "</span></span></li>";
    }).join("");
  }
  function custom() { app.state.edited = true; }
  function renderTree() {
    var tree = $("tree"), st = app.state;
    tree.innerHTML = "";
    M.MODEL.forEach(function (g) {
      var box = document.createElement("div");
      box.className = "group" + (g.warn ? " lists" : "");
      var head = document.createElement("div");
      head.className = "group-head";
      head.innerHTML = "<span class='t'><h3>" + esc(g.name) + "</h3><small>" + esc(g.note) + "</small></span>";
      var locked = g.internalOnly && st.audience === "sponsor";
      var on = locked ? false : st.groups[g.id];
      var sw = document.createElement("button");
      sw.type = "button"; sw.className = "switch"; sw.setAttribute("role", "switch");
      sw.setAttribute("aria-checked", String(!!on)); sw.setAttribute("aria-label", "Include " + g.name); sw.disabled = locked;
      sw.addEventListener("click", function () { st.groups[g.id] = !st.groups[g.id]; custom(); renderContent(); });
      head.appendChild(sw); box.appendChild(head);
      if (locked) {
        var lk = document.createElement("p"); lk.className = "lock";
        lk.textContent = "Locked off for sponsors. Choose the Internal audience to include participant lists.";
        box.appendChild(lk);
      }
      var items = document.createElement("div");
      items.className = "items"; items.hidden = !on;
      g.items.forEach(function (it) {
        var row = document.createElement("div");
        row.className = "item" + (st.items[it.id] ? "" : " off");
        var cb = document.createElement("input");
        cb.type = "checkbox"; cb.id = "i-" + it.id; cb.checked = !!st.items[it.id];
        cb.addEventListener("change", function () { st.items[it.id] = cb.checked; custom(); renderContent(); });
        var span = document.createElement("span"), lab = document.createElement("label");
        lab.htmlFor = cb.id; lab.textContent = it.name; span.appendChild(lab);
        if (it.cols && st.items[it.id]) {
          var cols = document.createElement("div"); cols.className = "cols";
          it.cols.forEach(function (c) {
            var key = it.id + ":" + c.id, l = document.createElement("label"), ccb = document.createElement("input");
            ccb.type = "checkbox"; ccb.id = "c-" + key.replace(/\W+/g, "-"); ccb.checked = !!st.cols[key];
            ccb.addEventListener("change", function () { st.cols[key] = ccb.checked; custom(); renderContent(); });
            l.appendChild(ccb); l.appendChild(document.createTextNode(c.label)); cols.appendChild(l);
          });
          span.appendChild(cols);
        }
        row.appendChild(cb); row.appendChild(span); items.appendChild(row);
      });
      box.appendChild(items); tree.appendChild(box);
    });
    $("a-sponsor").setAttribute("aria-pressed", String(st.audience === "sponsor"));
    $("a-internal").setAttribute("aria-pressed", String(st.audience === "internal"));
    $("edited").hidden = !st.edited;
  }
  function renderPreview() {
    var paper = $("paper"), r = app.result, st = app.state;
    if (!r) {
      paper.innerHTML = "<div class='p-eyebrow'>Cleo Diagnostics · CSP-01-001</div><div class='p-title'>Monthly trial summary</div>" +
        "<p class='empty'>Page 1 appears here once all 8 files are added and the checks pass.</p>";
      $("counts").innerHTML = "";
      return;
    }
    var sel = M.toSelection(st, r), one = W.pageOne(r, sel), p = r.period, meta = app.data.meta;
    var html = "<div class='p-eyebrow'>" + (app.example ? "Example data · " : "") + "Cleo Diagnostics · CSP-01-001</div>" +
      "<div class='p-title'>Monthly trial summary: " + esc(p.label) + "</div>" +
      "<div class='p-meta'>Data to " + E.fmtDate(p.lastDay) + "." + (meta.citrusExport ? " Citrus export " + E.fmtDate2(E.dayOf(meta.citrusExport)) + "." : "") +
      (meta.inventoryDate ? " Central lab inventory report " + esc(meta.inventoryDate) + "." : "") + "</div>" +
      "<div class='p-draft'>" + (st.audience === "sponsor" ? "DRAFT FOR REVIEW. Check every figure before sending to the sponsor." : "INTERNAL VERSION. Not for the sponsor.") + "</div><hr>";
    if (one.tiles.length) html += "<div class='tiles'>" + one.tiles.map(function (t) {
      return "<div class='tile'><b>" + esc(t[0]) + "</b><span>" + esc(t[1]) + "</span><small>" + esc(t[2]) + "</small></div>"; }).join("") + "</div>";
    if (sel.page1.highlights) html += "<h4>Highlights</h4>" + (one.highlights.length ? "<ul>" + one.highlights.map(function (x) { return "<li>" + esc(x) + "</li>"; }).join("") + "</ul>" :
      "<p class='empty'>No highlights: every linked table is off.</p>");
    if (one.actions.length) html += "<h4>Needs attention</h4><ul>" + one.actions.map(function (x) { return "<li>" + esc(x) + "</li>"; }).join("") + "</ul>";
    var secs = W.resolve(r, sel);
    html += "<h4>In this summary</h4><div class='contents'>" + (secs.length ? secs.map(function (s) {
      var word = s.tables.filter(function (t) { return !t.detail; }).length, lists = s.tables.length - word;
      return "<div><span>" + esc(s.title) + "</span><span>" + word + (word === 1 ? " table" : " tables") + (lists ? " + " + lists + (lists === 1 ? " list" : " lists") + " in Excel" : "") + "</span></div>";
    }).join("") : "<p class='empty'>No sections selected.</p>") + "</div>";
    paper.innerHTML = html;
    var shown = 0, total = 0;
    secs.forEach(function (s) { shown += s.tables.length; });
    M.MODEL.forEach(function (g) { if (!g.page1) total += g.items.length; });
    var listsOn = st.audience === "internal" && sel.sections.lists && M.MODEL.filter(function (g) { return g.internalOnly; })[0].items.some(function (it) { return sel.tables[it.id]; });
    $("counts").innerHTML = "<b>" + shown + " of " + total + "</b> tables and charts included · Participant lists <b>" + (listsOn ? "on" : "off") + "</b>";
  }
  function renderDownloads() {
    var ready = !!app.result, blocked = app.checks.some(function (c) { return c.level === "bad"; });
    ["dl-docx", "dl-xlsx", "dl-csv"].forEach(function (id) { $(id).disabled = !ready; });
    $("dl-blocked").hidden = !blocked;
  }
  function renderSteps() {
    var mark = function (id, done, bad) { var el = $(id); el.className = bad ? "blocked" : done ? "done" : ""; };
    var all = E.FILE_TYPES.every(function (t) { return app.files[t.key]; }), bad = app.checks.some(function (c) { return c.level === "bad"; });
    mark("nav-study", !!app.month);
    mark("nav-files", all);
    mark("nav-checks", all && !bad, bad);
    mark("nav-content", !!app.result);
    $("site-list-label").textContent = "CLEO-001 site list · " + SITES.length + " sites" + (app.period ? " · " +
      SITES.filter(siteActive).length + " active in " + app.period.short : "");
  }
  function siteActive(s) {
    var p = app.period;
    var iso = function (v) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(v || ""); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / E.DAY : null; };
    if (s.activation_date && iso(s.activation_date) > p.lastDay) return false;
    if (s.terminated === "Yes") return !!s.termination_date && iso(s.termination_date) >= p.firstDay;
    return true;
  }
  function renderContent() { renderTree(); renderPreview(); }
  function renderAll() { renderFiles(); renderChecks(); renderContent(); renderDownloads(); renderSteps(); }

  // ------------------------------------------------------------ downloads
  function fileName(ext) {
    var aud = app.state.audience === "sponsor" ? "Sponsor" : "Internal";
    return (app.example ? "EXAMPLE_" : "") + "CleoUSA_Monthly_" + (ext === "csv" ? "Totals_" : "Summary_") + app.result.period.tag + (ext === "csv" ? "" : "_" + aud) + "." + ext;
  }
  function setStatus(kind, text, cls) { var el = $("st-" + kind); el.textContent = text || ""; el.className = "status" + (cls ? " " + cls : ""); }
  var ERR = { declined: "Not saved.", rate_limited: "Another save prompt is open. Try again in a moment.", rejected_extension: "This file type cannot be saved here.",
    extension_not_enabled: "This file type cannot be saved in this view." };
  function deliver(kind, blob, name) {
    if (app.downloads) {
      return app.downloads.save({ filename: name, data: blob }).then(function () { setStatus(kind, "Saved " + name, "ok"); markDownloaded(); },
        function (e) { setStatus(kind, ERR[e && e.code] || "Downloads are not available in this view.", (e && e.code) === "declined" ? "" : "bad"); });
    }
    if (app.inViewer) { setStatus(kind, "Downloads are not available in this view. Open the page in claude.ai.", "bad"); return Promise.resolve(); }
    var a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
    setStatus(kind, "Saved " + name, "ok"); markDownloaded();
    return Promise.resolve();
  }
  function markDownloaded() { $("nav-download").className = "done"; }
  function buildBlob(kind) {
    var r = app.result, sel = M.toSelection(app.state, r), meta = app.data.meta, opts = { example: app.example };
    if (kind === "csv") return Promise.resolve(new Blob([W.totalsCsv(r, sel)], { type: "text/csv" }));
    return W.drawCharts(r).then(function (charts) {
      if (kind === "docx") return window.docx.Packer.toBlob(W.buildDocx(r, meta, sel, charts, opts));
      return W.buildXlsx(r, meta, sel, charts, opts).xlsx.writeBuffer().then(function (buf) {
        return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      });
    });
  }
  function makeFile(kind) {
    if (!app.result) return;
    var btn = $("dl-" + kind);
    btn.disabled = true; setStatus(kind, "Building…", "busy");
    buildBlob(kind).then(function (blob) { return deliver(kind, blob, fileName(kind)); })
      .catch(function (err) { setStatus(kind, "The file could not be built: " + (err && err.message || err), "bad"); })
      .then(function () { btn.disabled = !app.result; });
  }

  // ------------------------------------------------------------ wiring
  var drop = $("drop"), input = $("file-input");
  $("btn-choose").addEventListener("click", function () { input.click(); });
  input.addEventListener("change", function () { if (input.files.length) addFileList(input.files); input.value = ""; });
  ["dragenter", "dragover"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add("over"); }); });
  ["dragleave", "drop"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove("over"); }); });
  drop.addEventListener("drop", function (e) { if (e.dataTransfer && e.dataTransfer.files.length) addFileList(e.dataTransfer.files); });
  $("btn-example").addEventListener("click", loadExample);
  $("btn-clear").addEventListener("click", clearFiles);
  $("btn-clear-example").addEventListener("click", clearFiles);
  $("f-month").addEventListener("change", function () { app.month = this.value; refresh(); });
  $("a-sponsor").addEventListener("click", function () { app.state = stateFor("sponsor"); renderContent(); $("choices-status").textContent = ""; });
  $("a-internal").addEventListener("click", function () { app.state = stateFor("internal"); renderContent(); $("choices-status").textContent = ""; });
  $("btn-save-choices").addEventListener("click", function () { app.state.edited = false; saveChoices(); renderContent(); });
  $("btn-reset-choices").addEventListener("click", function () {
    try { var s = JSON.parse(localStorage.getItem(STORE_KEY) || "{}"); delete s[app.state.audience]; localStorage.setItem(STORE_KEY, JSON.stringify(s)); } catch (e) { /* storage blocked */ }
    app.state = M.defaultState(app.state.audience); $("choices-status").textContent = "Back to the default choices."; renderContent();
  });
  ["docx", "xlsx", "csv"].forEach(function (k) { $("dl-" + k).addEventListener("click", function () { makeFile(k); }); });

  app.inViewer = !!(window.claude && window.claude.use);
  if (app.inViewer) window.claude.use("downloads").then(function (d) { app.downloads = d; }, function () { app.downloads = null; });
  if (!window.ExcelJS || !window.docx) {
    var n = $("dl-note"); n.hidden = false; n.className = "notice bad";
    n.textContent = "The Word or Excel library did not load. Check your connection and reload the page.";
  }
  renderAll();

  // Self-test, only when the address ends in #selftest: build every output from the example data through the
  // same code the download buttons use, and print them as base64 for an external check. It never runs otherwise.
  function selfTest() {
    var out = { checks: app.checks.map(function (c) { return c.level + ": " + c.title; }), files: {} };
    var toB64 = function (blob) { return new Promise(function (res) { var fr = new FileReader(); fr.onload = function () { res(String(fr.result).split(",")[1]); }; fr.readAsDataURL(blob); }); };
    var seq = Promise.resolve();
    ["sponsor", "internal"].forEach(function (aud) {
      ["docx", "xlsx", "csv"].forEach(function (kind) {
        seq = seq.then(function () { app.state = M.defaultState(aud); return buildBlob(kind); }).then(toB64)
          .then(function (b64) { out.files[aud + "." + kind] = b64; });
      });
    });
    return seq.then(function () { out.preview = $("paper").textContent; },
      function (err) { out.error = String(err && err.stack || err); })
      .then(function () { var pre = document.createElement("pre"); pre.id = "selftest-out"; pre.hidden = true; pre.textContent = JSON.stringify(out); document.body.appendChild(pre); });
  }
  if (EXAMPLES.length) {
    var loading = loadExample();
    if (location.hash === "#selftest" && loading) loading.then(selfTest);
  }
})();
