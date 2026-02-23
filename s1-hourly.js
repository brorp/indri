const ExcelJS = require("exceljs");

const output_name_str = (process.argv[2].match(/(\d{6,8})/) || ["", "output"])[1];
const OUTPUT_FILE = `output-s1-hourly-${output_name_str}.xlsx`;
const THRESHOLD_GT = 99;

const SHEET_GRAFIK = ["GRAFFIK", "GRAFIK", "GRAFFIK ", "GRAFIK "];
const SHEET_DATA = ["DATA", "DATA ", "DATA  "];
const SHEET_SITE = ["SITE LIST S1 SR", "SITE LIST", "SITELIST S1 SR", "SITELIST"];

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

  // DD/MM/YY(YY) + optional time
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

  // handle "100%", "99.5", "99,5"
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
  // STRICT: need 5 values AND ALL > 99
  if (!lastVals || lastVals.length < 5) return "open";
  if (lastVals.some(v => v === null || v === undefined)) return "open";
  return lastVals.every(v => v >= THRESHOLD_GT) ? "close" : "open";
}

async function run(inputPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(inputPath);

  const wsGrafik = pickWorksheet(wb, SHEET_GRAFIK);
  if (!wsGrafik) throw new Error(`GRAFFIK sheet not found (tried: ${SHEET_GRAFIK.join(", ")})`);

  const wsData = pickWorksheet(wb, SHEET_DATA);
  if (!wsData) throw new Error(`DATA sheet not found (tried: ${SHEET_DATA.join(", ")})`);

  const wsSite = pickWorksheet(wb, SHEET_SITE);
  if (!wsSite) throw new Error(`SITE LIST sheet not found (tried: ${SHEET_SITE.join(", ")})`);

  // --- 1) Take LAST 5 Row Labels from GRAFFIK col A ---
  const grafikHourMs = [];
  wsGrafik.eachRow((row) => {
    const ms = getDateMs(row.getCell(1).value);
    if (ms !== null) grafikHourMs.push(floorToHourMs(ms));
  });

  if (grafikHourMs.length < 5) {
    throw new Error(`GRAFFIK col A has only ${grafikHourMs.length} datetime rows; need >= 5.`);
  }

  const last5HourMs = grafikHourMs.slice(-5); // EXACTLY bottom-5 from GRAFFIK

  // --- 2) Build DATA map: CEK -> Map(hourMs -> max(S1 Success Rate)) ---
  const hmap = buildHeaderMap(wsData);
  const colTime = hmap.get("time");
  const colCek = hmap.get("cek");
  const colVal =
    hmap.get("s1 success rate") ||
    hmap.get("average of s1 success rate") || // just in case
    hmap.get("avg of s1 success rate");

  if (!colTime || !colCek || !colVal) {
    throw new Error(`DATA headers must include: Time, CEK, S1 Success Rate`);
  }

  const cekHourMax = new Map(); // cek -> Map(hourMs -> max)
  const cekHasAnyNumeric = new Set();

  wsData.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const tMs = getDateMs(row.getCell(colTime).value);
    if (tMs === null) return;

    const cek = cellToText(row.getCell(colCek).value);
    if (!cek) return;

    const hourMs = floorToHourMs(tMs);

    const v = toNumber(row.getCell(colVal).value);
    if (v === null) return;

    cekHasAnyNumeric.add(cek);

    if (!cekHourMax.has(cek)) cekHourMax.set(cek, new Map());
    const m = cekHourMax.get(cek);

    const prev = m.get(hourMs);
    if (prev === undefined || v > prev) m.set(hourMs, v);
  });

  // --- 3) CEK list source: SITE LIST S1 SR -> Entity_ID ---
  const siteH = buildHeaderMap(wsSite);
  const colEntity = siteH.get("entity_id") || siteH.get("entity id") || 1; // fallback col 1

  const cekSet = new Set();
  wsSite.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const v = cellToText(row.getCell(colEntity).value);
    if (v) cekSet.add(v);
  });

  const cekList = Array.from(cekSet).sort((a, b) => a.localeCompare(b));

  // --- 4) Output ---
  const outWb = new ExcelJS.Workbook();
  const outWs = outWb.addWorksheet("Sheet1");
  outWs.addRow(["CEK", "STATUS"]);

  for (const cek of cekList) {
    if (!cekHasAnyNumeric.has(cek)) {
      outWs.addRow([cek, "NO DATA"]);
      continue;
    }

    const hourMap = cekHourMax.get(cek) || new Map();
    const lastVals = last5HourMs.map(ms => hourMap.get(ms)); // EXACT match by hourMs

    const status = decideStatus(lastVals);
    outWs.addRow([cek, status]);
  }

  outWs.getColumn(1).width = 28;
  outWs.getColumn(2).width = 12;

  await outWb.xlsx.writeFile(OUTPUT_FILE);

  console.log(`✅ Generated ${OUTPUT_FILE}`);
  console.log(`📌 Sheets: DATA="${wsData.name}", GRAFFIK="${wsGrafik.name}", SITE="${wsSite.name}"`);
  console.log(`📌 Last 5 Row Labels (from GRAFFIK): ${last5HourMs.map(ms => new Date(ms).toISOString()).join(" | ")}`);
  console.log(`📌 Rule: CLOSE only if ALL last-5 values > ${THRESHOLD_GT}; missing => OPEN; no numeric => NO DATA`);
  console.log(`📌 Total CEK from SITE LIST: ${cekList.length}`);
}

const inputFile = process.argv[2];
if (!inputFile) {
  console.log('Usage: node agent_s1_pivot_last5_like_twamp.js "ANALISA S1 SR 18022026.xlsx"');
  process.exit(1);
}

run(inputFile).catch(err => {
  console.error("❌ Error:", err.message);
  process.exit(1);
});
