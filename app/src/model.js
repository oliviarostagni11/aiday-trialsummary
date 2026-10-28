/* Trial Summary Builder: the content model behind the toggles, and the sponsor and internal defaults. */
(function (root) {
  "use strict";
  var MEAN_MEDIAN_RANGE = [
    { id: "mean", label: "Mean", match: function (h) { return /^Mean/.test(h); } },
    { id: "median", label: "Median", match: function (h) { return /^Median/.test(h); } },
    { id: "range", label: "Shortest and longest", match: function (h) { return h === "Shortest" || h === "Longest" || /^Range/.test(h); } }
  ];
  var MODEL = [
    { id: "page1", name: "Page 1", note: "Opening page of the Word summary and the Excel Summary tab", page1: true, items: [
      { id: "tiles", name: "Key figure tiles" }, { id: "highlights", name: "Highlights" }, { id: "attention", name: "Needs attention" }] },
    { id: "enrollment", name: "Enrollment", note: "Funnel, weekly site rates, site types", items: [
      { id: "funnel", name: "Recruitment summary" },
      { id: "conversion", name: "Consent to enrollment" },
      { id: "status_eom", name: "Participant status at month end" },
      { id: "enrol_trend", name: "Enrollments by month (chart in Word)" },
      { id: "enrol_by_site_week", name: "Enrollments by site and week", cols: [
        { id: "weeks", label: "Weekly counts", match: function (h) { return h.indexOf("–") >= 0; } },
        { id: "perweek", label: "Per week", match: function (h) { return h === "Per week"; } },
        { id: "todate", label: "Enrolled to date", match: function (h) { return h === "Enrolled to date"; } },
        { id: "permonth", label: "Per month since activation", match: function (h) { return h === "Per month since activation" || h === "Months active"; } }] },
      { id: "enrol_by_type", name: "Enrollment by site type" }] },
    { id: "completeness", name: "Data completeness", note: "Snapshot at the export time. Off for sponsors by default.", sponsorOff: true, items: [
      { id: "crf_by_site", name: "CRF completeness by site" },
      { id: "crf_by_form", name: "CRF completeness by form" },
      { id: "epro_by_form", name: "ePRO completeness by form" },
      { id: "epro_by_site", name: "ePRO completeness by site" }] },
    { id: "post_surgical", name: "Post-surgical findings", note: "Surgeries, pathology reports, turnaround", items: [
      { id: "psf_summary", name: "Surgery and pathology reports" },
      { id: "psf_turnaround", name: "Days from enrollment to report upload", cols: MEAN_MEDIAN_RANGE },
      { id: "psf_by_site", name: "Surgery and pathology reports by site" }] },
    { id: "blood", name: "Blood samples", note: "Central lab registration windows and pairs", items: [
      { id: "blood_summary", name: "Sample registration at the central lab" },
      { id: "blood_trend", name: "Sample registration by month (chart in Word)" },
      { id: "blood_pairs", name: "Sample pair status at month end" },
      { id: "blood_by_site", name: "Sample registration by site" },
      { id: "blood_location", name: "Where registered samples are now", sponsorOff: true }] },
    { id: "queries", name: "Data queries", note: "Open, reviewing and closed, with days open", items: [
      { id: "query_summary", name: "Query status at month end" },
      { id: "query_days_open", name: "Days open for unresolved queries", cols: MEAN_MEDIAN_RANGE },
      { id: "query_chart", name: "Unresolved queries by site chart" },
      { id: "query_by_site", name: "Queries by site", cols: MEAN_MEDIAN_RANGE },
      { id: "query_by_type", name: "Queries by type" },
      { id: "query_by_form", name: "Queries by form", sponsorOff: true }] },
    { id: "lists", name: "Participant lists", note: "Names participants and samples. Internal Excel workbook only.", warn: true, internalOnly: true, items: [
      { id: "psf_awaiting", name: "Participants awaiting a pathology report" },
      { id: "blood_missing_pair", name: "Samples missing their pair" },
      { id: "blood_no_sample", name: "Enrolled participants with no sample registered" },
      { id: "blood_out_of_window", name: "Samples registered outside their window" },
      { id: "blood_same_day", name: "Pairs registered on the same day" }] }
  ];

  // UI state -> selection for the writers
  function defaultState(audience) {
    var st = { audience: audience, groups: {}, items: {}, cols: {}, edited: false };
    var internal = audience === "internal";
    MODEL.forEach(function (g) {
      st.groups[g.id] = internal ? true : !(g.sponsorOff || g.internalOnly);
      g.items.forEach(function (it) {
        st.items[it.id] = internal ? true : !it.sponsorOff;
        (it.cols || []).forEach(function (c) { st.cols[it.id + ":" + c.id] = true; });
      });
    });
    return st;
  }
  function toSelection(st, result) {
    var sel = { audience: st.audience, sections: {}, tables: {}, hiddenCols: {}, page1: {} };
    MODEL.forEach(function (g) {
      var groupOn = st.groups[g.id] && !(g.internalOnly && st.audience === "sponsor");
      if (g.page1) { g.items.forEach(function (it) { sel.page1[it.id] = groupOn && st.items[it.id]; }); return; }
      sel.sections[g.id] = groupOn;
      g.items.forEach(function (it) { sel.tables[it.id] = !!st.items[it.id]; });
    });
    // The participant lists live inside the post-surgical and blood sections, so those sections carry them.
    if (result) {
      Object.keys(result.sections).forEach(function (k) {
        result.sections[k].tables.forEach(function (t) {
          var item = null;
          MODEL.forEach(function (g) { g.items.forEach(function (it) { if (it.id === t.key) item = it; }); });
          if (!item || !item.cols) return;
          var hidden = [];
          item.cols.forEach(function (c) {
            if (st.cols[t.key + ":" + c.id]) return;
            t.cols.forEach(function (col, i) { if (i > 0 && c.match(col.header)) hidden.push(i); });
          });
          sel.hiddenCols[t.key] = hidden;
        });
      });
    }
    return sel;
  }
  root.TSModel = { MODEL: MODEL, defaultState: defaultState, toSelection: toSelection };
})(typeof globalThis !== "undefined" ? globalThis : this);
