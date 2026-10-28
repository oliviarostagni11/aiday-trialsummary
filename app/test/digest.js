// 'Inside this report' rows and the section key-figure strips must match narrative.digest() in Python.
// Usage: node app/test/digest.js <python-digest.json> [month]
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
const py = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
let bad = 0;
for (const [aud, included] of Object.entries(py)) {
  const js = E.digest(res.sections, res.head, res.period, included.keys);
  if (JSON.stringify(js) !== JSON.stringify(included.digest)) { bad++; console.log("DIFF", aud, "\n js:", JSON.stringify(js), "\n py:", JSON.stringify(included.digest)); }
}
console.log(bad ? "digest differs" : "digest identical for sponsor and internal");
process.exit(bad ? 1 : 0);
