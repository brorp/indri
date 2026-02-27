const ExcelJS = require("exceljs");

const output_name_str = (process.argv[2].match(/(\d{6,8})/) || ["", "output"])[1];
const OUTPUT_FILE = `output-twamp-hourly-${output_name_str}.xlsx`;
const THRESHOLD_LT = 0.7;

const SHEET_DATA = ["DATA Twamp", "Twamp", "TWAMP", "DATA TWAMP", "DATA", "DATA ", "DATA  ", "data"];
const SHEET_SITE = ["SITE LIST TWAMP", "SITE LIST", "SITELIST TWAMP", "SITELIST",
  "site list twamp", "site list", "sitelist twamp", "sitelist"];

// ---------- helpers ----------
function pickWorksheet(workbook, candidates) {
  for (const name of candidates) {
    const ws = workbook.getWorksheet(name);
    if (ws) return ws;
  }
  const lowers = candidates.map(x => x.trim().toLowerCase()).filter(Boolean);
  for (const ws of workbook.worksheets) {
    const n = ws.name.trim().toLowerCase();
    if (lowers.some(c => n.includes(c))) return ws;
  }
  return null;
}

function cellToText(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object" && typeof v.error === "string") return v.error;
  if (typeof v === "object" && v.result !== undefined) return cellToText(v.result);
  if (typeof v === "object" && Array.isArray(v.richText)) return v.richText.map(x => x.text ?? "").join("").trim();
  if (typeof v === "object" && typeof v.text === "string") return v.text.trim();
  return String(v).trim();
}

function excelSerialToDateMs(serial) {
  return Math.round((serial - 25569) * 86400 * 1000);
}

function parseDateStringToMs(s) {
  const str = String(s).trim().replace(/\s+/g, " ");
  const parsed = Date.parse(str);
  if (Number.isFinite(parsed)) return parsed;

  let m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    let dd = Number(m[1]), mm = Number(m[2]), yy = Number(m[3]);
    const hh = m[4] ? Number(m[4]) : 0;
    const mi = m[5] ? Number(m[5]) : 0;
    const ss = m[6] ? Number(m[6]) : 0;
    if (yy < 100) yy = 2000 + yy;
    return new Date(yy, mm - 1, dd, hh, mi, ss).getTime();
  }
  return null;
}

function getDateMs(cellValue) {
  if (!cellValue) return null;
  if (typeof cellValue === "object" && cellValue.result !== undefined) return getDateMs(cellValue.result);
  if (cellValue instanceof Date) return cellValue.getTime();
  if (typeof cellValue === "number" && cellValue > 20000) return excelSerialToDateMs(cellValue);

  const s = cellToText(cellValue);
  if (!s) return null;
  return parseDateStringToMs(s);
}

function floorToHourMs(ms) {
  const d = new Date(ms);
  d.setMinutes(0, 0, 0);
  return d.getTime();
}

function toNumber(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "object" && v.result !== undefined) return toNumber(v.result);
  if (typeof v === "number") return v;

  const s0 = cellToText(v);
  if (!s0) return null;
  if (s0 === "#" || s0 === "-") return null;

  const n = Number(s0.replace("%", "").replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

function buildHeaderMap(ws) {
  const header = ws.getRow(1);
  const m = new Map();
  header.eachCell((cell, col) => {
    const t = cellToText(cell.value);
    if (t) m.set(t.trim().toLowerCase(), col);
  });
  return m;
}

function decideStatus(lastVals) {
  // STRICT: need 5 values AND ALL <= 0.7
  if (!lastVals || lastVals.length < 5) return "open";
  if (lastVals.some(v => v === null || v === undefined)) return "open";
  return lastVals.every(v => v <= THRESHOLD_LT) ? "close" : "open";
}

async function run(inputPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(inputPath);

  const wsData = pickWorksheet(wb, SHEET_DATA);
  if (!wsData) throw new Error(`DATA sheet not found (tried: ${SHEET_DATA.join(", ")})`);

  const wsSite = pickWorksheet(wb, SHEET_SITE);
  if (!wsSite) throw new Error(`SITE LIST sheet not found (tried: ${SHEET_SITE.join(", ")})`);

  // --- 1) Build DATA map: CEK -> Map(hourMs -> max(MAX TWAMP)) ---
  const hmap = buildHeaderMap(wsData);
  const colTime = hmap.get("time");
  const colCek = hmap.get("cek");
  const colVal =
    hmap.get("max twamp") ||
    hmap.get("max_twamp") ||
    hmap.get("average of max twamp") ||
    hmap.get("avg of max twamp");

  if (!colTime || !colCek || !colVal) {
    throw new Error(`DATA headers must include: Time, CEK, MAX TWAMP`);
  }

  const cekHourMax = new Map();   // cek -> Map(hourMs -> max)
  const cekHasAnyNumeric = new Set();

  wsData.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const tMs = getDateMs(row.getCell(colTime).value);
    if (tMs === null) return;

    const cek = cellToText(row.getCell(colCek).value);
    if (!cek) return;

    const cekUpper = cek.toUpperCase();
    if (cekUpper === "#N/A" || cekUpper === "N/A") return;

    const hourMs = floorToHourMs(tMs);

    const v = toNumber(row.getCell(colVal).value);
    if (v === null) return; // # or - or empty => skip

    cekHasAnyNumeric.add(cek);

    if (!cekHourMax.has(cek)) cekHourMax.set(cek, new Map());
    const m = cekHourMax.get(cek);

    const prev = m.get(hourMs);
    if (prev === undefined || v > prev) m.set(hourMs, v);
  });

  // --- 2) CEK list source: SITE LIST -> Entity_ID ---
  const siteH = buildHeaderMap(wsSite);
  const colEntity = siteH.get("entity_id") || siteH.get("entity id") || 1;

  const cekSet = new Set();
  wsSite.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const v = cellToText(row.getCell(colEntity).value);
    if (v) cekSet.add(v);
  });

  const cekList = Array.from(cekSet).sort((a, b) => a.localeCompare(b));

  // --- 3) Output ---
  const outWb = new ExcelJS.Workbook();
  const outWs = outWb.addWorksheet("Sheet1");
  outWs.addRow(["CEK", "STATUS"]);

  for (const cek of cekList) {
    if (!cekHasAnyNumeric.has(cek)) {
      outWs.addRow([cek, "NO DATA"]);
      continue;
    }

    const hourMap = cekHourMax.get(cek) || new Map();
    const last5HourMs = Array.from(hourMap.keys()).sort((a, b) => a - b).slice(-5);
    const lastVals = last5HourMs.map(ms => hourMap.get(ms)); // each CEK's own latest 5

    const status = decideStatus(lastVals);
    outWs.addRow([cek, status]);
  }

  outWs.getColumn(1).width = 32;
  outWs.getColumn(2).width = 12;

  await outWb.xlsx.writeFile(OUTPUT_FILE);

  console.log(`✅ Generated ${OUTPUT_FILE}`);
  console.log(`📌 Sheets: DATA="${wsData.name}", SITE="${wsSite.name}"`);
  console.log(`📌 Rule: CLOSE only if each CEK's own last-5 values <= ${THRESHOLD_LT}; <5 data => OPEN; no numeric => NO DATA`);
  console.log(`📌 Total CEK from SITE LIST: ${cekList.length}`);
}

const inputFile = process.argv[2];
if (!inputFile) {
  console.log('Usage: node twamp-hourly.js "ANALISA TWAMP 23022026.xlsx"');
  process.exit(1);
}

run(inputFile).catch(err => {
  console.error("❌ Error:", err.message);
  process.exit(1);
});
