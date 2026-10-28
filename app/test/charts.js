// The chart specs (data, colours, takeaway titles) must be identical in the browser tool and the Python report.
// Usage: node app/test/charts.js <python-specs.json> [month]
const fs = require("fs"), path = require("path");
require(path.join(__dirname, "../src/engine.js"));
const E = globalThis.TSEngine, IN = "data/dummy_inputs", data = { meta: {} };
for (const f of fs.readdirSync(IN).filter(f => f.endsWith(".csv"))) {
  const rows = E.parseCSV(fs.readFileSync(path.join(IN, f), "utf8")), t = E.identify(rows[0]);
  if (!t) continue;
  data[t.key] = E.rowsToObjects(rows[0], rows.slice(1));
  if (t.key === "participants") data.meta.citrusExport = E.exportTimeFromName(f);
}
const sr = E.parseCSV(fs.readFileSync("data/reference/site_lookup.csv", "utf8"));
data.sites = E.rowsToObjects(sr[0], sr.slice(1));
const res = E.build(data, process.argv[3] || "2026-07");
const js = E.chartSpecs(res.sections, res.period, res.head), py = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const norm = o => JSON.stringify(o, (k, v) => (v && typeof v === "object" && !Array.isArray(v)) ? Object.fromEntries(Object.entries(v).sort()) : v);
let bad = 0;
for (const k of new Set([...Object.keys(js), ...Object.keys(py)])) {
  if (norm(js[k]) !== norm(py[k])) { bad++; console.log("DIFF", k, "\n js:", norm(js[k]).slice(0, 600), "\n py:", norm(py[k]).slice(0, 600)); }
}
console.log(bad ? `${bad} chart specs differ` : `all ${Object.keys(js).length} chart specs identical`);
process.exit(bad ? 1 : 0);
