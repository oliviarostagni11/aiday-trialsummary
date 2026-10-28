// The sponsor highlights and action items must read the same in the browser tool and the Python report.
// Usage: node app/test/sponsor_text.js <python-output.json>
const fs = require("fs"), path = require("path");
for (const f of ["engine.js", "model.js", "writers.js"]) require(path.join(__dirname, "../src", f));
const E = globalThis.TSEngine, M = globalThis.TSModel, W = globalThis.TSWriters;
const IN = "data/dummy_inputs", data = { meta: {} };
for (const f of fs.readdirSync(IN).filter(f => f.endsWith(".csv"))) {
  const rows = E.parseCSV(fs.readFileSync(path.join(IN, f), "utf8")), t = E.identify(rows[0]);
  if (!t) continue;
  data[t.key] = E.rowsToObjects(rows[0], rows.slice(1));
  if (t.key === "participants") data.meta.citrusExport = E.exportTimeFromName(f);
}
const sr = E.parseCSV(fs.readFileSync("data/reference/site_lookup.csv", "utf8"));
data.sites = E.rowsToObjects(sr[0], sr.slice(1));
const res = E.build(data, process.argv[3] || "2026-07");
const one = W.pageOne(res, M.toSelection(M.defaultState("sponsor"), res));
const py = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
let bad = 0;
[["highlights", one.highlights, py.highlights], ["actions", one.actions, py.actions]].forEach(([name, js, p]) => {
  if (JSON.stringify(js) !== JSON.stringify(p)) { bad++; console.log("DIFF in", name, "\n js:", js, "\n py:", p); }
});
console.log(bad ? "sponsor text differs" : `sponsor text identical: ${one.highlights.length} highlights, ${one.actions.length} action items`);
process.exit(bad ? 1 : 0);
