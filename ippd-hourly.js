const ExcelJS = require("exceljs");

const outputNameStr = (process.argv[2]?.match(/(\d{6,8})/) || ["", "output"])[1];
const OUTPUT_FILE = `output-ippd-hourly-${outputNameStr}.xlsx`;
const THRESHOLD_PERCENT = 0.7;

const SHEET_DATA = ["DATA", "DATA ", "DATA  ", "data", "Data"];

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

  const m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    let dd = Number(m[1]);
    let mm = Number(m[2]);
    let yy = Number(m[3]);
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

function toPercentNumber(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "object" && v.result !== undefined) return toPercentNumber(v.result);
  if (typeof v === "number") return v <= 1 ? v * 100 : v;

  const s0 = cellToText(v);
  if (!s0) return null;
  if (s0 === "#" || s0 === "-") return null;

  const hasPercentSign = s0.includes("%");
  const n = Number(s0.replace("%", "").replace(",", "."));
  if (!Number.isFinite(n)) return null;
  if (hasPercentSign) return n;
  return n <= 1 ? n * 100 : n;
}

function buildHeaderMapFromRowValues(values) {
  const m = new Map();
  for (let col = 1; col < values.length; col += 1) {
    const t = cellToText(values[col]);
    if (t) m.set(t.trim().toLowerCase(), col);
  }
  return m;
}

function decideStatus(lastVals) {
  // STRICT: need 5 values AND ALL <= 0.7%
  if (!lastVals || lastVals.length < 5) return "open";
  if (lastVals.some(v => v === null || v === undefined)) return "open";
  return lastVals.every(v => v <= THRESHOLD_PERCENT) ? "close" : "open";
}

function isDataSheetName(name) {
  const normalized = String(name || "").trim().toLowerCase();
  return SHEET_DATA.some(candidate => {
    const c = candidate.trim().toLowerCase();
    return c && (normalized === c || normalized.includes(c));
  });
}

async function run(inputPath) {
  const towerHourMax = new Map();   // towerId -> Map(hourMs -> max packet loss)
  const towerHasAnyNumeric = new Set();
  const towerSet = new Set();

  let parsedSheetName = null;
  let colTime = null;
  let colTower = null;
  let colVal = null;

  const reader = new ExcelJS.stream.xlsx.WorkbookReader(inputPath, {
    sharedStrings: "cache",
    styles: "ignore",
    hyperlinks: "ignore",
    worksheets: "emit",
  });

  for await (const ws of reader) {
    if (!isDataSheetName(ws.name)) continue;

    parsedSheetName = ws.name;

    for await (const row of ws) {
      const rowNumber = row.number;
      const values = row.values;

      if (rowNumber === 1) {
        const hmap = buildHeaderMapFromRowValues(values);
        colTime = hmap.get("begin time") || hmap.get("time");
        colTower =
          hmap.get("tower id") ||
          hmap.get("tower_id") ||
          hmap.get("towerid");
        colVal =
          hmap.get("packet loss rate(%)") ||
          hmap.get("packet loss rate (%)") ||
          hmap.get("package loss rate (%)") ||
          hmap.get("package loss rate(%)");

        if (!colTime || !colTower || !colVal) {
          throw new Error("DATA headers must include: Begin Time/Time, TOWER ID, Packet/Package Loss Rate (%)");
        }
        continue;
      }

      const towerId = cellToText(values[colTower]);
      if (!towerId) continue;

      const towerUpper = towerId.toUpperCase();
      if (towerUpper === "#N/A" || towerUpper === "N/A") continue;
      towerSet.add(towerId);

      const tMs = getDateMs(values[colTime]);
      if (tMs === null) continue;
      const hourMs = floorToHourMs(tMs);

      const v = toPercentNumber(values[colVal]);
      if (v === null) continue;

      towerHasAnyNumeric.add(towerId);

      if (!towerHourMax.has(towerId)) towerHourMax.set(towerId, new Map());
      const hourMap = towerHourMax.get(towerId);

      const prev = hourMap.get(hourMs);
      if (prev === undefined || v > prev) hourMap.set(hourMs, v);
    }

    break;
  }

  if (!parsedSheetName) {
    throw new Error(`DATA sheet not found (tried: ${SHEET_DATA.join(", ")})`);
  }

  const towerList = Array.from(towerSet).sort((a, b) => a.localeCompare(b));

  const outWb = new ExcelJS.Workbook();
  const outWs = outWb.addWorksheet("Sheet1");
  outWs.addRow(["TOWER_ID", "STATUS"]);

  for (const towerId of towerList) {
    if (!towerHasAnyNumeric.has(towerId)) {
      outWs.addRow([towerId, "NO DATA"]);
      continue;
    }

    const hourMap = towerHourMax.get(towerId) || new Map();
    const last5HourMs = Array.from(hourMap.keys()).sort((a, b) => a - b).slice(-5);
    const lastVals = last5HourMs.map(ms => hourMap.get(ms));
    const status = decideStatus(lastVals);
    outWs.addRow([towerId, status]);
  }

  outWs.getColumn(1).width = 32;
  outWs.getColumn(2).width = 12;

  await outWb.xlsx.writeFile(OUTPUT_FILE);

  console.log(`✅ Generated ${OUTPUT_FILE}`);
  console.log(`📌 Sheet DATA="${parsedSheetName}"`);
  console.log(`📌 Rule: CLOSE only if each tower's own last-5 values <= ${THRESHOLD_PERCENT}% (percent mode); <5 data => OPEN; no numeric => NO DATA`);
  console.log(`📌 Total TOWER_ID from DATA: ${towerList.length}`);
}

const inputFile = process.argv[2];
if (!inputFile) {
  console.log('Usage: node ippd-hourly.js "ANALISA IPPD 23022026.xlsx"');
  process.exit(1);
}

run(inputFile).catch(err => {
  console.error("❌ Error:", err.message);
  process.exit(1);
});
