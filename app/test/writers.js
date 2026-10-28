// Writer test: build sponsor and internal Word/Excel/CSV from the dummy data, then scan the sponsor files.
// Usage: NODE_PATH=<dir with docx@8.5.0 and exceljs@4.4.0> node app/test/writers.js <outdir>
const fs = require("fs"), path = require("path"), zlib = require("zlib");
globalThis.docx = require("docx");
globalThis.ExcelJS = require("exceljs");
for (const f of ["engine.js", "model.js", "writers.js"]) require(path.join(__dirname, "../src", f));
const E = globalThis.TSEngine, M = globalThis.TSModel, W = globalThis.TSWriters;
const OUT = process.argv[2] || ".";
const IN = "data/dummy_inputs";
const data = { meta: {} };
for (const f of fs.readdirSync(IN).filter(f => f.endsWith(".csv"))) {
  const rows = E.parseCSV(fs.readFileSync(path.join(IN, f), "utf8"));
  const t = E.identify(rows[0]);
  if (!t) continue;
  data[t.key] = E.rowsToObjects(rows[0], rows.slice(1));
  if (t.key === "participants") data.meta.citrusExport = E.exportTimeFromName(f);
  if (t.key === "inventory") { data.meta.inventoryDate = "03AUG2026"; data.meta.inventoryDay = E.inventoryDateDay("03AUG2026"); data.meta.droppedDob = rows[0].includes("Date of Birth"); }
}
const sr = E.parseCSV(fs.readFileSync("data/reference/site_lookup.csv", "utf8"));
data.sites = E.rowsToObjects(sr[0], sr.slice(1));
const res = E.build(data, "2026-07");
console.log("checks:", E.checks(data, Object.fromEntries(E.FILE_TYPES.map(t => [t.key, { label: t.label, missingCols: [] }])), res.period).map(c => c.level + ": " + c.title).join(" | "));

function unzipText(buf) {  // minimal zip reader: inflate every .xml entry
  let text = "", i = 0;
  while ((i = buf.indexOf(Buffer.from([0x50, 0x4b, 0x03, 0x04]), i)) >= 0) {
    const method = buf.readUInt16LE(i + 8), csize = buf.readUInt32LE(i + 18), nlen = buf.readUInt16LE(i + 26), xlen = buf.readUInt16LE(i + 28);
    const name = buf.slice(i + 30, i + 30 + nlen).toString(), start = i + 30 + nlen + xlen;
    if (name.endsWith(".xml") && csize) text += method === 8 ? zlib.inflateRawSync(buf.slice(start, start + csize)).toString() : buf.slice(start, start + csize).toString();
    i = start + csize;
  }
  return text;
}
(async () => {
  for (const aud of ["sponsor", "internal"]) {
    const st = M.defaultState(aud), sel = M.toSelection(st, res);
    const name = `TEST_CleoUSA_Monthly_Summary_${res.period.tag}_${aud === "sponsor" ? "Sponsor" : "Internal"}`;
    const docBuf = await globalThis.docx.Packer.toBuffer(W.buildDocx(res, data.meta, sel, null, {}));
    fs.writeFileSync(path.join(OUT, name + ".docx"), docBuf);
    const xBuf = Buffer.from(await W.buildXlsx(res, data.meta, sel, null, {}).xlsx.writeBuffer());
    fs.writeFileSync(path.join(OUT, name + ".xlsx"), xBuf);
    const csv = W.totalsCsv(res, sel);
    for (const [ext, buf] of [["docx", docBuf], ["xlsx", xBuf]]) {
      const t = unzipText(buf);
      console.log(aud, ext, `${(buf.length / 1024).toFixed(0)} KB`, "| participant IDs:", (t.match(/USA-\d{3,4}/g) || []).length,
        "| 'CRF completeness':", (t.match(/CRF completeness/g) || []).length, "| 'Birth':", (t.match(/Birth/g) || []).length);
    }
    console.log(aud, "totals csv lines:", csv.split("\n").length - 2, "| participant IDs:", (csv.match(/USA-\d{3,4}/g) || []).length);
    console.log(aud, "page 1:", JSON.stringify(W.pageOne(res, sel).tiles.map(t => t[1])));
  }
})().catch(e => { console.error(e); process.exit(1); });
