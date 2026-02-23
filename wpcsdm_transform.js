"use strict";

/**
 * wpcsdm_step1_2_3_safe.js
 *
 * EDIT ONLY THESE COLUMNS:
 * - KPI D-1
 * - TAGGING
 * - MOCN Date / MOCN DATE
 * - Status
 *
 * DOES NOT TOUCH M/N/O (or other columns).
 * All VLOOKUP misses => write real Excel error { error: "#N/A" } (so no blanks).
 *
 * Run:
 * node --max-old-space-size=8192 wpcsdm_step1_2_3_safe.js \
 *   --wpc "wpcsdm_wpc_export_20260214144337_default.xlsx" \
 *   --sfxl "NEW SFXL 14022026.xlsx" \
 *   --sitelist "sitelist_mocn_20260212.csv" \
 *   --tagging "TAGGING 13022026.xlsx" \
 *   --out "wpcsdm_out.xlsx"
 */

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

// ================= CLI =================
function getArg(flag) {
  const i = process.argv.indexOf(flag);
  return i >= 0 ? process.argv[i + 1] : null;
}
function abs(p) {
  return p ? path.resolve(process.cwd(), p) : null;
}
function must(p, label) {
  if (!p) throw new Error(`Missing ${label}`);
  if (!fs.existsSync(p)) throw new Error(`File not found ${label}: ${p}`);
  return p;
}

const WPC_PATH = must(abs(getArg("--wpc")), "--wpc");
const SFXL_PATH = must(abs(getArg("--sfxl")), "--sfxl");
const SITELIST_PATH = must(abs(getArg("--sitelist")), "--sitelist");
const TAGGING_PATH = must(abs(getArg("--tagging")), "--tagging");
const OUT_PATH = abs(getArg("--out"));
if (!OUT_PATH) throw new Error("Missing --out");

// ================= CONFIG =================
const SHEET_WPC = "wpcsdm_wpc_export";

const KPI_DATA = new Set([
  "Avg CQI",
  "Avg DL SE",
  "S1 Set up success rate (%)",
  "UE DL IP Throughput",
  "UE UL IP Throughput",
]);

const KPI_IPPD = "IPPD Packet Loss";
const KPI_TWAMP = "TWAMP Packet Loss";

const STEP2_WPC = new Set([
  "Avg CQI",
  "Avg DL SE",
  "UE DL IP Throughput",
  "UE UL IP Throughput",
]);

// ================= Helpers =================
function excelKey(x) {
  return String(x ?? "")
    .replace(/\u00A0/g, " ") // NBSP
    .replace(/[\u200B-\u200D\uFEFF]/g, "") // zero width
    .trim();
}

function cellStr(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") {
    if (v.result !== undefined && v.result !== null) return String(v.result).trim();
    if (typeof v.text === "string") return v.text.trim();
    if (Array.isArray(v.richText)) return v.richText.map(x => x.text || "").join("").trim();
    if (v.error) return String(v.error);
  }
  return String(v).trim();
}

function num(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "object" && v) {
    if (v.result !== undefined) return num(v.result);
    if (v.error) return null;
  }
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function isBlank(v) {
  if (v === null || v === undefined) return true;
  if (typeof v === "string") return v.trim() === "";
  if (typeof v === "object" && v) {
    if (v.result !== undefined) return isBlank(v.result);
    if (v.error) return false; // error is not blank
    if (typeof v.text === "string") return v.text.trim() === "";
    if (Array.isArray(v.richText)) return v.richText.map(x => x.text || "").join("").trim() === "";
  }
  return String(v).trim() === "";
}

function normHeader(h) {
  return cellStr(h).replace(/\s+/g, " ").trim().toUpperCase();
}

function buildHeaderIndex(row) {
  const m = new Map();
  row.eachCell((cell, colNumber) => {
    const h = normHeader(cell.value);
    if (h) m.set(h, colNumber);
  });
  return m;
}

function pickCol(hmap, variants) {
  for (const v of variants) {
    const key = normHeader(v);
    if (hmap.has(key)) return hmap.get(key);
  }
  return null;
}

function setExcelNA(cell) {
  cell.value = { error: "#N/A" }; // real Excel error
}

function parseDateLike(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "number") {
    // excel serial date
    if (v > 20000) {
      const ms = (v - 25569) * 86400 * 1000;
      const d = new Date(ms);
      return isNaN(d.getTime()) ? null : d;
    }
  }
  if (typeof v === "object" && v && v.result !== undefined) return parseDateLike(v.result);

  const s = String(v).trim();
  if (!s) return null;

  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const dd = Number(m[1]);
    const mm = Number(m[2]) - 1;
    let yy = Number(m[3]);
    if (yy < 100) yy += 2000;
    const d2 = new Date(yy, mm, dd);
    if (!isNaN(d2.getTime())) return d2;
  }
  return null;
}

/**
 * Reproduce Column O logic to decide KPI Normalized in Step 1 (without touching O).
 */
function yesOpenByFormula(wpcName, kpiD1, day7) {
  const k = num(kpiD1);
  const d7 = num(day7);
  if (k === null) return "Open";

  const diff = (d7 !== null) ? (k - d7) : null;
  const ratio = (diff !== null && d7 !== null && d7 !== 0) ? (diff / d7) : null;

  if (wpcName === "DL Traffic") {
    if (diff !== null && ratio !== null && diff > -50 && ratio > -0.10) return "Yes";
    return "Open";
  }
  if (wpcName === "S1 Set up success rate (%)") {
    return k > 99 ? "Yes" : "Open";
  }
  if (wpcName === KPI_IPPD) {
    return k <= 0.7 ? "Yes" : "Open";
  }

  const ratioKpis = new Set([
    "Avg CQI",
    "Avg DL SE",
    "UE DL IP Throughput",
    "UE UL IP Throughput",
    "RRC Conn Users",
  ]);

  if (ratio !== null && ratio > -0.10 && ratioKpis.has(wpcName)) return "Yes";
  return "Open";
}

// ================= CSV parser (no dependency) =================
function parseCsvLine(line) {
  const out = [];
  let cur = "";
  let inQ = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];

    if (inQ) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQ = false;
        }
      } else {
        cur += ch;
      }
    } else {
      if (ch === ",") {
        out.push(cur);
        cur = "";
      } else if (ch === '"') {
        inQ = true;
      } else {
        cur += ch;
      }
    }
  }
  out.push(cur);
  return out;
}

function loadSitelistCSV(filePath) {
  const raw = fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "");
  const lines = raw.split(/\r?\n/).filter(l => l.trim() !== "");
  if (!lines.length) return { mapAll: new Map(), mapSf: new Map() };

  const headers = parseCsvLine(lines[0]).map(h => normHeader(h));
  const idx = (name) => headers.indexOf(normHeader(name));

  const iId = idx("New XL ID");
  const iMocn = idx("MOCN Date");
  const iKeep = idx("Keep/Drop");

  if (iId < 0) throw new Error("SITELIST CSV missing column: New XL ID");
  if (iMocn < 0) throw new Error("SITELIST CSV missing column: MOCN Date");
  if (iKeep < 0) throw new Error("SITELIST CSV missing column: Keep/Drop");

  const mapAll = new Map();
  for (let r = 1; r < lines.length; r++) {
    const cols = parseCsvLine(lines[r]);
    const id = excelKey(cols[iId]);
    if (!id) continue;

    const mocnDate = cols[iMocn];
    const keepDrop = cols[iKeep];

    mapAll.set(id, { mocnDate, keepDrop });
  }

  // CSV has no "SF only" sheet => use same
  const mapSf = new Map(mapAll);
  return { mapAll, mapSf };
}

// ================= Load TAGGING map (Tower ID -> Remark) =================
async function loadTaggingMap(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];

  const hmap = buildHeaderIndex(ws.getRow(1));
  const cTower = pickCol(hmap, ["Tower ID", "TowerID"]);
  const cRemark = pickCol(hmap, ["Remark"]);

  if (!cTower || !cRemark) throw new Error("TAGGING XLSX missing Tower ID / Remark");

  const map = new Map();
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === 1) return;
    const tower = excelKey(row.getCell(cTower).value);
    if (!tower) return;
    const remark = cellStr(row.getCell(cRemark).value);
    if (remark) map.set(tower, remark);
  });

  return map;
}

// ================= Stream NEW SFXL maps (DATA/TRAFFIC/TWAMP) =================
async function loadSfxlMaps(sfxlPath, needKeys) {
  const dataMap = new Map();    // MOEntity -> metrics
  const trafficMap = new Map(); // Row Labels -> {payload, rrc}
  const twampMap = new Map();   // Row Labels -> twamp

  const reader = new ExcelJS.stream.xlsx.WorkbookReader(sfxlPath, {
    entries: "emit",
    sharedStrings: "cache", // IMPORTANT for string keys
    styles: "ignore",
    hyperlinks: "ignore",
    worksheets: "emit",
  });

  for await (const ws of reader) {
    const wsName = cellStr(ws.name).toUpperCase();
    if (!["DATA", "TRAFFIC", "TWAMP"].includes(wsName)) continue;

    let hmap = null;

    for await (const row of ws) {
      if (row.number === 1) {
        hmap = buildHeaderIndex(row);
        continue;
      }
      if (!hmap) continue;

      if (wsName === "DATA") {
        const cKey = pickCol(hmap, ["MOEntity"]);
        if (!cKey) continue;

        const key = excelKey(row.getCell(cKey).value);
        if (!key || !needKeys.data.has(key)) continue;

        const cAvgCqi = pickCol(hmap, ["Avg CQI"]);
        const cDlSe = pickCol(hmap, ["DL SE"]);
        const cS1 = pickCol(hmap, ["S1 Setup Success Rate"]);
        const cDlThr = pickCol(hmap, ["DL User Throughput"]);
        const cUlThr = pickCol(hmap, ["UL User Throughput"]);
        const cRank2 = pickCol(hmap, ["Rank2"]);

        dataMap.set(key, {
          avgCqi: cAvgCqi ? num(row.getCell(cAvgCqi).value) : null,
          dlSe: cDlSe ? num(row.getCell(cDlSe).value) : null,
          s1: cS1 ? num(row.getCell(cS1).value) : null,
          dlThr: cDlThr ? num(row.getCell(cDlThr).value) : null,
          ulThr: cUlThr ? num(row.getCell(cUlThr).value) : null,
          rank2: cRank2 ? num(row.getCell(cRank2).value) : null,
        });

        continue;
      }

      if (wsName === "TRAFFIC") {
        const cKey = pickCol(hmap, ["Row Labels"]);
        const cPayload = pickCol(hmap, ["Sum of Payload per PLMN"]);
        const cRrc = pickCol(hmap, ["Sum of RRC User per PLMN"]);
        if (!cKey) continue;

        const key = excelKey(row.getCell(cKey).value);
        if (!key || !needKeys.traffic.has(key)) continue;

        trafficMap.set(key, {
          payload: cPayload ? num(row.getCell(cPayload).value) : null,
          rrc: cRrc ? num(row.getCell(cRrc).value) : null,
        });

        continue;
      }

      if (wsName === "TWAMP") {
        const cKey = pickCol(hmap, ["Row Labels"]);
        const cVal = pickCol(hmap, ["Max of MAX TWAMP"]);
        if (!cKey || !cVal) continue;

        const key = excelKey(row.getCell(cKey).value);
        if (!key || !needKeys.twamp.has(key)) continue;

        twampMap.set(key, num(row.getCell(cVal).value));
        continue;
      }
    }
  }

  return { dataMap, trafficMap, twampMap };
}

// ================= MAIN =================
async function main() {
  console.log("Load WPC...");
  const wpcWb = new ExcelJS.Workbook();
  await wpcWb.xlsx.readFile(WPC_PATH);

  // Force Excel to recalc formulas when opening output
  wpcWb.calcProperties = wpcWb.calcProperties || {};
  wpcWb.calcProperties.fullCalcOnLoad = true;

  const ws = wpcWb.getWorksheet(SHEET_WPC);
  if (!ws) throw new Error(`Missing sheet '${SHEET_WPC}'`);

  const headerRow = ws.getRow(1);
  const hmap = buildHeaderIndex(headerRow);

  const cEntity = pickCol(hmap, ["Entity_ID"]);
  const cWpc = pickCol(hmap, ["WPC Name"]);
  const cDay7 = pickCol(hmap, ["Day-7"]);
  const cKpi = pickCol(hmap, ["KPI D-1"]);
  const cStatus = pickCol(hmap, ["Status"]);

  const cTower = pickCol(hmap, ["Tower ID", "TowerID"]);
  const cTag = pickCol(hmap, ["TAGGING"]);
  const cMocn = pickCol(hmap, ["MOCN DATE", "MOCN Date"]);

  const cDesc2 = pickCol(hmap, ["Description2"]);
  const cPriority = pickCol(hmap, ["Priority"]);
  const cOperator = pickCol(hmap, ["Operator"]);

  if (!cEntity || !cWpc || !cDay7 || !cKpi || !cStatus || !cTower || !cTag || !cMocn) {
    throw new Error("WPC missing required columns (Entity_ID, WPC Name, Day-7, KPI D-1, Status, Tower ID, TAGGING, MOCN Date)");
  }

  // rename header of MOCN to "MOCN DATE" (allowed, still same column)
  headerRow.getCell(cMocn).value = "MOCN DATE";
  ws.getColumn(cMocn).numFmt = "dd/mm/yy";

  // Collect needed keys for SFXL
  const needKeys = { data: new Set(), traffic: new Set(), twamp: new Set() };
  const needTower = new Set();

  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === 1) return;
    const entity = excelKey(row.getCell(cEntity).value);
    const tower = excelKey(row.getCell(cTower).value);
    const wpcName = cellStr(row.getCell(cWpc).value);

    if (tower) needTower.add(tower);

    if (entity && (KPI_DATA.has(wpcName) || wpcName === KPI_IPPD)) needKeys.data.add(entity);
    if (entity && (wpcName === "DL Traffic" || wpcName === "RRC Conn Users")) needKeys.traffic.add(entity);
    if (entity && wpcName === KPI_TWAMP) needKeys.twamp.add(entity);
  });

  console.log("Load sitelist...");
  const { mapAll: sitelistAll, mapSf: sitelistSf } = loadSitelistCSV(SITELIST_PATH);

  console.log("Load tagging...");
  const taggingMap = await loadTaggingMap(TAGGING_PATH);

  console.log("Stream NEW SFXL...");
  const { dataMap, trafficMap, twampMap } = await loadSfxlMaps(SFXL_PATH, needKeys);

  console.log("Maps loaded:", {
    DATA: dataMap.size,
    TRAFFIC: trafficMap.size,
    TWAMP: twampMap.size,
    SITELIST: sitelistAll.size,
    TAGGING: taggingMap.size,
  });

  // ===================== STEP 1 (KPI D-1 + Status) =====================
  console.log("STEP 1...");
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === 1) return;

    const entity = excelKey(row.getCell(cEntity).value);
    const wpcName = cellStr(row.getCell(cWpc).value);
    const day7 = row.getCell(cDay7).value;

    const isTarget =
      KPI_DATA.has(wpcName) ||
      wpcName === KPI_IPPD ||
      wpcName === "DL Traffic" ||
      wpcName === "RRC Conn Users" ||
      wpcName === KPI_TWAMP;

    if (!isTarget) return;

    let newKpi = null;
    let found = false;

    if (KPI_DATA.has(wpcName)) {
      const rec = dataMap.get(entity);
      if (rec) {
        if (wpcName === "Avg CQI") newKpi = rec.avgCqi;
        else if (wpcName === "Avg DL SE") newKpi = rec.dlSe;
        else if (wpcName === "S1 Set up success rate (%)") newKpi = rec.s1;
        else if (wpcName === "UE DL IP Throughput") newKpi = rec.dlThr;
        else if (wpcName === "UE UL IP Throughput") newKpi = rec.ulThr;
        found = true;
      }
    } else if (wpcName === KPI_IPPD) {
      // fallback for your SFXL: Rank2*100
      const rec = dataMap.get(entity);
      if (rec && rec.rank2 !== null && rec.rank2 !== undefined) {
        newKpi = rec.rank2 * 100;
        found = true;
      }
    } else if (wpcName === "DL Traffic") {
      const rec = trafficMap.get(entity);
      if (rec) {
        newKpi = rec.payload;
        found = true;
      }
    } else if (wpcName === "RRC Conn Users") {
      const rec = trafficMap.get(entity);
      if (rec) {
        newKpi = rec.rrc;
        found = true;
      }
    } else if (wpcName === KPI_TWAMP) {
      const v = twampMap.get(entity);
      if (v !== undefined) {
        newKpi = v;
        found = true;
      }
    }

    // KPI D-1 must not be blank: write number if found else #N/A
    if (found && newKpi !== null && newKpi !== undefined) {
      row.getCell(cKpi).value = newKpi;
    } else {
      setExcelNA(row.getCell(cKpi));
    }

    // Status logic (Step 1)
    const kpiVal = (found && newKpi !== null && newKpi !== undefined) ? newKpi : null;

    if (wpcName !== KPI_IPPD && wpcName !== KPI_TWAMP) {
      const yesOpen = yesOpenByFormula(wpcName, kpiVal ?? row.getCell(cKpi).value, day7);
      if (yesOpen === "Yes") row.getCell(cStatus).value = "KPI Normalized";
    } else {
      // IPPD & TWAMP: L < 0.6 and (L - Day-7) < 0.2
      const k = num(kpiVal ?? row.getCell(cKpi).value);
      const d7 = num(day7);
      const diff = (k !== null && d7 !== null) ? (k - d7) : null;
      if (k !== null && diff !== null && k < 0.6 && diff < 0.2) {
        row.getCell(cStatus).value = "KPI Normalized";
      }
    }
  });

  // ===================== STEP 2 (MOCN DATE + Status NY SSH Approval) =====================
  console.log("STEP 2...");
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === 1) return;

    const wpcName = cellStr(row.getCell(cWpc).value);
    const tower = excelKey(row.getCell(cTower).value);

    const desc2 = cDesc2 ? cellStr(row.getCell(cDesc2).value) : "";
    const priority = cPriority ? cellStr(row.getCell(cPriority).value) : "";
    const operator = cOperator ? cellStr(row.getCell(cOperator).value) : "";

    const pass =
      desc2.toUpperCase().includes("SSH") &&
      priority === "P1" &&
      operator === "MOCN" &&
      STEP2_WPC.has(wpcName);

    if (!pass) return;

    // VLOOKUP MOCN DATE: TowerID -> sitelist (New XL ID -> MOCN Date)
    const hit = sitelistAll.get(tower);

    if (hit && !isBlank(hit.mocnDate)) {
      const d = parseDateLike(hit.mocnDate);
      row.getCell(cMocn).value = d ? d : String(hit.mocnDate).trim();
    } else {
      // make sure it "berbuah" even if miss
      setExcelNA(row.getCell(cMocn));
    }

    // NY SSH Approval: only if Status blank AND MOCN DATE year 2025/2026
    if (isBlank(row.getCell(cStatus).value)) {
      const d = parseDateLike(row.getCell(cMocn).value);
      if (d && (d.getFullYear() === 2025 || d.getFullYear() === 2026)) {
        row.getCell(cStatus).value = "NY SSH Approval";
      }
    }
  });

  // ===================== STEP 3 (TAGGING + SF Keep/Drop + Close status) =====================
  console.log("STEP 3...");

  // 3.1 Fill TAGGING for ALL rows: TowerID -> Remark, if miss => #N/A
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === 1) return;

    const tower = excelKey(row.getCell(cTower).value);

    // Always enforce no blank in TAGGING column
    if (isBlank(row.getCell(cTag).value) || cellStr(row.getCell(cTag).value) === "#N/A") {
      const remark = taggingMap.get(tower);
      if (remark) row.getCell(cTag).value = remark;
      else setExcelNA(row.getCell(cTag));
    }
  });

  // 3.2-3.4 SF rule
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === 1) return;

    const tower = excelKey(row.getCell(cTower).value);
    const operator = cOperator ? cellStr(row.getCell(cOperator).value) : "";

    const tagU = cellStr(row.getCell(cTag).value).toUpperCase();
    const matchTag = ["DROP", "DISMANTLE", "DISMANTLED", "NYOA"].includes(tagU);

    // If operator SF and tagging in list => MOCN DATE becomes Keep/Drop from sitelist "SF only"
    if (operator === "SF" && matchTag) {
      const hit = sitelistSf.get(tower);
      if (hit && !isBlank(hit.keepDrop)) {
        row.getCell(cMocn).value = String(hit.keepDrop).trim(); // "Keep" / "Drop"
      } else {
        setExcelNA(row.getCell(cMocn));
      }
    }

    // If MOCN DATE == "Drop" and Status blank => Close due to site already drop
    const mocnText = cellStr(row.getCell(cMocn).value);
    if (mocnText === "Drop" && isBlank(row.getCell(cStatus).value)) {
      row.getCell(cStatus).value = "Close due to site already drop";
    }
  });

  console.log("Write output:", OUT_PATH);
  await wpcWb.xlsx.writeFile(OUT_PATH);
  console.log("DONE ✅ Step1+2+3 (safe columns only)");
}

main().catch((e) => {
  console.error("ERROR:", e.message || e);
  process.exit(1);
});
