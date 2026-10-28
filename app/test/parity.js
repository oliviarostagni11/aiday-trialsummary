// Parity test: the JavaScript engine must reproduce every total in the Python totals CSV.
// Usage: node app/test/parity.js [month]   (run from the repo root after scripts/build_report.py)
const fs = require("fs"), path = require("path");
require(path.join(__dirname, "../src/engine.js"));
const E = globalThis.TSEngine;
const month = process.argv[2] || "2026-07";
const IN = "data/dummy_inputs";

function load(file) {
  const rows = E.parseCSV(fs.readFileSync(path.join(IN, file), "utf8"));
  return { header: rows[0], objs: E.rowsToObjects(rows[0], rows.slice(1)) };
}
const files = fs.readdirSync(IN).filter(f => f.endsWith(".csv"));
const data = { meta: {} };
for (const f of files) {
  const { header, objs } = load(f);
  const t = E.identify(header);
  if (!t) continue;
  data[t.key] = objs;
  if (t.key === "participants") data.meta.citrusExport = E.exportTimeFromName(f);
  if (t.key === "inventory") { data.meta.inventoryDate = (f.match(/_(\d{2}[A-Z]{3}\d{4})/) || [])[1]; data.meta.inventoryDay = E.inventoryDateDay(f); }
}
const siteRows = E.parseCSV(fs.readFileSync("data/reference/site_lookup.csv", "utf8"));
data.sites = E.rowsToObjects(siteRows[0], siteRows.slice(1));

const res = E.build(data, month);
const py = E.parseCSV(fs.readFileSync(`outputs/${month}/CleoUSA_Monthly_Totals_${res.period.tag}.csv`, "utf8")).slice(1).filter(r => r.length > 1);

const js = [];
for (const s of Object.values(res.sections)) for (const t of s.tables) {
  if (t.detail) continue;
  for (const r of t.rows) t.cols.slice(1).forEach((c, i) => {
    let v = r[i + 1];
    if (Array.isArray(v)) v = `${v[0].toFixed(1)}-${v[1].toFixed(1)}`;
    js.push([s.title, t.title, String(r[0]), c.header, v == null ? "" : v]);
  });
}
let bad = 0;
if (js.length !== py.length) { console.log(`row count differs: js ${js.length}, python ${py.length}`); bad++; }
for (let i = 0; i < Math.min(js.length, py.length); i++) {
  const a = js[i], b = py[i];
  const same = a.slice(0, 4).every((x, k) => x === b[k]) &&
    (typeof a[4] === "number" && b[4] !== "" ? Math.abs(a[4] - Number(b[4])) < 1e-3 : String(a[4]) === b[4]);
  if (!same && bad++ < 15) console.log("DIFF", JSON.stringify(a), "| python:", JSON.stringify(b));
}
// Detail lists are not in the totals file: compare their row counts with the Python report's own counts instead.
const counts = {};
for (const s of Object.values(res.sections)) for (const t of s.tables) if (t.detail) counts[t.key] = t.rows.length;
console.log(`compared ${js.length} values, ${bad} differences. detail list rows:`, JSON.stringify(counts));
console.log("highlights:\n - " + E.highlights(res.head, res.period).join("\n - "));
process.exit(bad ? 1 : 0);
