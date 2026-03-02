"use strict";

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

function getArg(flag) {
  const i = process.argv.indexOf(flag);
  return i >= 0 ? process.argv[i + 1] : null;
}

function abs(p) {
  return p ? path.resolve(process.cwd(), p) : null;
}

function ensureFile(p, label) {
  if (!p) throw new Error(`Missing ${label}`);
  if (!fs.existsSync(p)) throw new Error(`File not found ${label}: ${p}`);
  return p;
}

function deriveResultPath(wpcPath) {
  const dir = path.dirname(wpcPath);
  const base = path.basename(wpcPath);

  if (/_default\.xlsx$/i.test(base)) {
    return path.join(dir, base.replace(/_default\.xlsx$/i, "_result.xlsx"));
  }

  if (/\.xlsx$/i.test(base)) {
    return path.join(dir, base.replace(/\.xlsx$/i, "_result.xlsx"));
  }

  return path.join(dir, `${base}_result.xlsx`);
}

const WPC_PATH = ensureFile(abs(getArg("--wpc")), "--wpc");
const SFXL_PATH = ensureFile(abs(getArg("--sfxl")), "--sfxl");
const SITELIST_PATH = ensureFile(abs(getArg("--sitelist")), "--sitelist");
const TAGGING_PATH = ensureFile(abs(getArg("--tagging")), "--tagging");
const OUT_PATH = abs(getArg("--out")) || deriveResultPath(WPC_PATH);

const WPC_SHEET_CANDIDATES = ["wpcsdm_wpc_export", "wpc export", "wpcsdm"];

const KPI = {
  AVG_CQI: "AVG_CQI",
  AVG_DL_SE: "AVG_DL_SE",
  S1_SETUP: "S1_SETUP",
  UE_DL_TP: "UE_DL_TP",
  UE_UL_TP: "UE_UL_TP",
  IPPD: "IPPD",
  DL_TRAFFIC: "DL_TRAFFIC",
  RRC_USERS: "RRC_USERS",
  TWAMP: "TWAMP",
  OTHER: "OTHER",
};

const STEP2_KPI_SET = new Set([KPI.AVG_CQI, KPI.AVG_DL_SE, KPI.UE_DL_TP, KPI.UE_UL_TP]);

function normalizeSpaces(s) {
  return String(s || "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function cellText(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return normalizeSpaces(v);
  if (typeof v === "number") return String(v);
  if (v instanceof Date) return v.toISOString();

  if (typeof v === "object") {
    if (v.result !== undefined) return cellText(v.result);
    if (typeof v.text === "string") return normalizeSpaces(v.text);
    if (Array.isArray(v.richText)) return normalizeSpaces(v.richText.map((x) => x.text || "").join(""));
    if (v.error) return String(v.error).trim();
  }

  return normalizeSpaces(String(v));
}

function toNumber(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;

  if (typeof v === "object") {
    if (v.result !== undefined) return toNumber(v.result);
    if (v.error) return null;
  }

  const s = cellText(v);
  if (!s || s === "#" || s === "-") return null;

  const pct = s.includes("%");
  const n = Number(s.replace(/%/g, "").replace(/,/g, "."));
  if (!Number.isFinite(n)) return null;

  return pct ? n : n;
}

function isBlankValue(v) {
  if (v === null || v === undefined) return true;
  if (typeof v === "string") return normalizeSpaces(v) === "";

  if (typeof v === "object") {
    if (v.result !== undefined) return isBlankValue(v.result);
    if (v.error) return false;
    if (typeof v.text === "string") return normalizeSpaces(v.text) === "";
    if (Array.isArray(v.richText)) return normalizeSpaces(v.richText.map((x) => x.text || "").join("")) === "";
  }

  return normalizeSpaces(String(v)) === "";
}

function setExcelNA(cell) {
  cell.value = { error: "#N/A" };
}

function normHeader(v) {
  return normalizeSpaces(cellText(v)).toUpperCase();
}

function buildHeaderMap(row) {
  const map = new Map();
  row.eachCell((cell, col) => {
    const h = normHeader(cell.value);
    if (h) map.set(h, col);
  });
  return map;
}

function pickCol(hmap, variants) {
  for (const v of variants) {
    const key = normHeader(v);
    if (hmap.has(key)) return hmap.get(key);
  }
  return null;
}

function findColContainsAll(hmap, words) {
  const upWords = words.map((w) => String(w).toUpperCase());
  for (const [k, idx] of hmap.entries()) {
    if (upWords.every((w) => k.includes(w))) return idx;
  }
  return null;
}

function parseDateLike(v) {
  if (!v) return null;

  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;

  if (typeof v === "number") {
    if (v > 20000) {
      const ms = Math.round((v - 25569) * 86400 * 1000);
      const d = new Date(ms);
      return Number.isNaN(d.getTime()) ? null : d;
    }
    return null;
  }

  if (typeof v === "object" && v.result !== undefined) {
    return parseDateLike(v.result);
  }

  const s = normalizeSpaces(cellText(v));
  if (!s) return null;

  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const dd = Number(m[1]);
    const mm = Number(m[2]) - 1;
    let yy = Number(m[3]);
    const hh = m[4] ? Number(m[4]) : 0;
    const mi = m[5] ? Number(m[5]) : 0;
    const ss = m[6] ? Number(m[6]) : 0;
    if (yy < 100) yy += 2000;

    const d = new Date(yy, mm, dd, hh, mi, ss);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  const parsed = Date.parse(s);
  if (Number.isFinite(parsed)) return new Date(parsed);

  return null;
}

function normalizeWpcName(raw) {
  const t = normalizeSpaces(cellText(raw)).toUpperCase();

  if (!t) return KPI.OTHER;
  if (t.includes("AVG CQI")) return KPI.AVG_CQI;
  if (t.includes("AVG DL SE")) return KPI.AVG_DL_SE;
  if (t.includes("S1") && t.includes("SUCCESS")) return KPI.S1_SETUP;
  if (t.includes("UE DL") && t.includes("THROUGHPUT")) return KPI.UE_DL_TP;
  if (t.includes("UE UL") && t.includes("THROUGHPUT")) return KPI.UE_UL_TP;
  if (t.includes("IPPD") && t.includes("LOSS")) return KPI.IPPD;
  if (t.includes("DL TRAFFIC")) return KPI.DL_TRAFFIC;
  if (t.includes("RRC") && t.includes("USER")) return KPI.RRC_USERS;
  if (t.includes("TWAMP")) return KPI.TWAMP;

  return KPI.OTHER;
}

function yesForStep1(kpiType, kpiD1, day7) {
  const k = toNumber(kpiD1);
  const d7 = toNumber(day7);

  if (k === null) return false;

  const diff = d7 !== null ? k - d7 : null;
  const ratio = diff !== null && d7 !== 0 ? diff / d7 : null;

  if (kpiType === KPI.DL_TRAFFIC) {
    return diff !== null && ratio !== null && diff > -50 && ratio > -0.10;
  }

  if (kpiType === KPI.S1_SETUP) {
    return k > 99;
  }

  if (
    kpiType === KPI.AVG_CQI ||
    kpiType === KPI.AVG_DL_SE ||
    kpiType === KPI.UE_DL_TP ||
    kpiType === KPI.UE_UL_TP ||
    kpiType === KPI.RRC_USERS
  ) {
    return ratio !== null && ratio > -0.10;
  }

  return false;
}

function parseCsvLine(line) {
  const out = [];
  let cur = "";
  let inQ = false;

  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];

    if (inQ) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i += 1;
        } else {
          inQ = false;
        }
      } else {
        cur += ch;
      }
      continue;
    }

    if (ch === ",") {
      out.push(cur);
      cur = "";
    } else if (ch === '"') {
      inQ = true;
    } else {
      cur += ch;
    }
  }

  out.push(cur);
  return out;
}

function loadSitelistFromCsv(filePath) {
  const raw = fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "");
  const lines = raw.split(/\r?\n/).filter((l) => l.trim() !== "");
  if (!lines.length) return { allMap: new Map(), sfMap: new Map() };

  const headers = parseCsvLine(lines[0]).map((h) => normHeader(h));
  const idx = (name) => headers.indexOf(normHeader(name));

  const iId = idx("New XL ID");
  const iMocn = idx("MOCN Date");
  const iKeep = idx("Keep/Drop");

  if (iId < 0) throw new Error("Sitelist CSV missing column: New XL ID");

  const allMap = new Map();
  for (let r = 1; r < lines.length; r += 1) {
    const cols = parseCsvLine(lines[r]);
    const towerId = normalizeSpaces(cols[iId] || "");
    if (!towerId) continue;

    const mocnDate = iMocn >= 0 ? (cols[iMocn] || "") : "";
    const keepDrop = iKeep >= 0 ? (cols[iKeep] || "") : "";

    allMap.set(towerId, {
      mocnDate: normalizeSpaces(mocnDate),
      keepDrop: normalizeSpaces(keepDrop),
    });
  }

  return { allMap, sfMap: new Map(allMap) };
}

async function extractSitelistMapFromSheet(ws) {
  const hmap = buildHeaderMap(ws.getRow(1));

  const cId = pickCol(hmap, ["New XL ID", "New XLID", "XL ID", "TowerID", "Tower ID"]);
  if (!cId) return null;

  const cMocn = pickCol(hmap, ["MOCN DATE", "MOCN Date"]);
  const cKeep = pickCol(hmap, ["Keep/Drop", "Keep / Drop", "Keep Drop"]);

  const map = new Map();

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const towerId = normalizeSpaces(cellText(row.getCell(cId).value));
    if (!towerId) return;

    const mocnDate = cMocn ? row.getCell(cMocn).value : null;
    const keepDrop = cKeep ? row.getCell(cKeep).value : null;

    map.set(towerId, {
      mocnDate,
      keepDrop: normalizeSpaces(cellText(keepDrop)),
    });
  });

  return {
    map,
    hasMocn: Boolean(cMocn),
    hasKeepDrop: Boolean(cKeep),
  };
}

async function loadSitelistFromXlsx(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  let allCandidate = null;
  let sfCandidate = null;

  for (const ws of wb.worksheets) {
    const extracted = await extractSitelistMapFromSheet(ws);
    if (!extracted || !extracted.map.size) continue;

    const sheetName = normalizeSpaces(ws.name).toUpperCase();

    if (!allCandidate && extracted.hasMocn) {
      allCandidate = extracted.map;
    }

    if (extracted.hasMocn && allCandidate && extracted.map.size > allCandidate.size) {
      allCandidate = extracted.map;
    }

    if (sheetName.includes("SF") && sheetName.includes("ONLY") && extracted.hasKeepDrop) {
      sfCandidate = extracted.map;
    }
  }

  if (!allCandidate) {
    throw new Error("Sitelist XLSX: cannot find sheet with New XL ID + MOCN DATE");
  }

  if (!sfCandidate) sfCandidate = new Map(allCandidate);

  return { allMap: allCandidate, sfMap: sfCandidate };
}

async function loadSitelistMaps(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".csv") return loadSitelistFromCsv(filePath);
  return loadSitelistFromXlsx(filePath);
}

async function loadTaggingMap(filePath) {
  const ext = path.extname(filePath).toLowerCase();

  if (ext === ".csv") {
    const raw = fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "");
    const lines = raw.split(/\r?\n/).filter((l) => l.trim() !== "");
    if (!lines.length) return new Map();

    const headers = parseCsvLine(lines[0]).map((h) => normHeader(h));
    const idx = (name) => headers.indexOf(normHeader(name));

    const iTower = idx("Tower ID") >= 0 ? idx("Tower ID") : idx("TowerID");
    const iRemark = idx("Remark");
    if (iTower < 0 || iRemark < 0) {
      throw new Error("TAGGING CSV missing Tower ID / Remark");
    }

    const map = new Map();
    for (let r = 1; r < lines.length; r += 1) {
      const cols = parseCsvLine(lines[r]);
      const towerId = normalizeSpaces(cols[iTower] || "");
      if (!towerId) continue;
      map.set(towerId, normalizeSpaces(cols[iRemark] || ""));
    }

    return map;
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  for (const ws of wb.worksheets) {
    const hmap = buildHeaderMap(ws.getRow(1));
    const cTower = pickCol(hmap, ["Tower ID", "TowerID"]);
    const cRemark = pickCol(hmap, ["Remark"]);

    if (!cTower || !cRemark) continue;

    const map = new Map();
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;
      const towerId = normalizeSpaces(cellText(row.getCell(cTower).value));
      if (!towerId) return;
      map.set(towerId, normalizeSpaces(cellText(row.getCell(cRemark).value)));
    });

    return map;
  }

  throw new Error("TAGGING XLSX missing Tower ID / Remark");
}

function classifySfxlSheet(name) {
  const n = normalizeSpaces(name).toUpperCase();
  if (!n) return null;

  if (n === "DATA") return "DATA";
  if (n.includes("IPPD")) return "IPPD";
  if (n === "TRAFFIC" || n === "PLMN" || n.includes("TRAFFIC") || n.includes("PLMN")) return "TRAFFIC";
  if (n.includes("TWAMP")) return "TWAMP";
  return null;
}

function ingestSfxlRow(type, row, hmap, needed, maps) {
  const { dataMap, ippdMap, trafficMap, twampMap } = maps;

  if (type === "DATA") {
    const cKey = pickCol(hmap, ["MOEntity"]);
    if (!cKey) return;

    const key = normalizeSpaces(cellText(row.getCell(cKey).value));
    if (!key || !needed.dataKeys.has(key)) return;

    const cAvgCqi = pickCol(hmap, ["Avg CQI"]);
    const cDlSe = pickCol(hmap, ["DL SE"]);
    const cS1 = pickCol(hmap, ["S1 Setup Success Rate"]);
    const cDlThr = pickCol(hmap, ["DL User Throughput"]);
    const cUlThr = pickCol(hmap, ["UL User Throughput"]);

    dataMap.set(key, {
      avgCqi: cAvgCqi ? toNumber(row.getCell(cAvgCqi).value) : null,
      dlSe: cDlSe ? toNumber(row.getCell(cDlSe).value) : null,
      s1: cS1 ? toNumber(row.getCell(cS1).value) : null,
      dlThr: cDlThr ? toNumber(row.getCell(cDlThr).value) : null,
      ulThr: cUlThr ? toNumber(row.getCell(cUlThr).value) : null,
    });

    return;
  }

  if (type === "IPPD") {
    const cKey = pickCol(hmap, ["Row Labels"]);
    if (!cKey) return;

    const key = normalizeSpaces(cellText(row.getCell(cKey).value));
    if (!key || !needed.ippdKeys.has(key)) return;

    let cVal = pickCol(hmap, ["IPPD*100"]);
    if (!cVal) cVal = findColContainsAll(hmap, ["IPPD", "100"]);
    if (!cVal) cVal = findColContainsAll(hmap, ["IPPD", "%"]);
    if (!cVal) return;

    const v = toNumber(row.getCell(cVal).value);
    if (v !== null) ippdMap.set(key, v);
    return;
  }

  if (type === "TRAFFIC") {
    const cKey = pickCol(hmap, ["Row Labels"]);
    if (!cKey) return;

    const key = normalizeSpaces(cellText(row.getCell(cKey).value));
    if (!key || !needed.trafficKeys.has(key)) return;

    const cPayload = pickCol(hmap, ["Sum of Payload per PLMN"]);
    const cRrc = pickCol(hmap, ["Sum of RRC User per PLMN"]);

    const prev = trafficMap.get(key) || { payload: null, rrc: null };
    const nextPayload = cPayload ? toNumber(row.getCell(cPayload).value) : null;
    const nextRrc = cRrc ? toNumber(row.getCell(cRrc).value) : null;

    trafficMap.set(key, {
      payload: nextPayload !== null ? nextPayload : prev.payload,
      rrc: nextRrc !== null ? nextRrc : prev.rrc,
    });
    return;
  }

  if (type === "TWAMP") {
    const cKey = pickCol(hmap, ["Row Labels"]);
    let cVal = pickCol(hmap, ["Max of MAX TWAMP"]);
    if (!cVal) cVal = findColContainsAll(hmap, ["MAX", "TWAMP"]);
    if (!cKey || !cVal) return;

    const key = normalizeSpaces(cellText(row.getCell(cKey).value));
    if (!key || !needed.twampKeys.has(key)) return;

    const v = toNumber(row.getCell(cVal).value);
    if (v !== null) twampMap.set(key, v);
  }
}

async function loadSfxlMapsByWorkbook(sfxlPath, needed, maps) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(sfxlPath);

  for (const ws of wb.worksheets) {
    const type = classifySfxlSheet(ws.name);
    if (!type) continue;

    const hmap = buildHeaderMap(ws.getRow(1));
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;
      ingestSfxlRow(type, row, hmap, needed, maps);
    });
  }
}

async function loadSfxlMaps(sfxlPath, needed) {
  const dataMap = new Map();
  const ippdMap = new Map();
  const trafficMap = new Map();
  const twampMap = new Map();
  const maps = { dataMap, ippdMap, trafficMap, twampMap };

  try {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(sfxlPath, {
      entries: "emit",
      sharedStrings: "cache",
      styles: "ignore",
      hyperlinks: "ignore",
      worksheets: "emit",
    });

    for await (const ws of reader) {
      const type = classifySfxlSheet(ws.name);
      if (!type) continue;

      let hmap = null;

      for await (const row of ws) {
        if (row.number === 1) {
          hmap = buildHeaderMap(row);
          continue;
        }
        if (!hmap) continue;
        ingestSfxlRow(type, row, hmap, needed, maps);
      }
    }
  } catch (err) {
    console.warn(`SFXL stream reader failed (${err.message}). Fallback to workbook reader...`);
    await loadSfxlMapsByWorkbook(sfxlPath, needed, maps);
  }

  return { dataMap, ippdMap, trafficMap, twampMap };
}

function pickWpcWorksheet(workbook) {
  for (const name of WPC_SHEET_CANDIDATES) {
    const ws = workbook.getWorksheet(name);
    if (ws) return ws;
  }

  const lowers = WPC_SHEET_CANDIDATES.map((x) => x.toLowerCase());
  for (const ws of workbook.worksheets) {
    const n = normalizeSpaces(ws.name).toLowerCase();
    if (lowers.some((x) => n.includes(x))) return ws;
  }

  return workbook.worksheets[0] || null;
}

function tagInDropDismantleNyOa(tagText) {
  const t = normalizeSpaces(tagText).toUpperCase();
  if (!t) return false;

  return t.includes("DROP") || t.includes("DISMANTLE") || t.includes("NYOA");
}

async function main() {
  console.log("Load WPC...");
  const wpcWb = new ExcelJS.Workbook();
  await wpcWb.xlsx.readFile(WPC_PATH);

  wpcWb.calcProperties = wpcWb.calcProperties || {};
  wpcWb.calcProperties.fullCalcOnLoad = true;

  const ws = pickWpcWorksheet(wpcWb);
  if (!ws) throw new Error("WPC sheet not found");

  const hmap = buildHeaderMap(ws.getRow(1));

  const cEntity = pickCol(hmap, ["Entity_ID", "Entity ID"]);
  const cWpc = pickCol(hmap, ["WPC Name"]);
  const cDay7 = pickCol(hmap, ["Day-7", "Day - 7"]);
  const cKpi = pickCol(hmap, ["KPI D-1", "KPI D - 1"]);
  const cStatus = pickCol(hmap, ["Status"]);
  const cTower = pickCol(hmap, ["Tower ID", "TowerID"]);
  const cTag = pickCol(hmap, ["Tagging", "TAGGING"]);
  const cMocn = pickCol(hmap, ["MOCN DATE", "MOCN Date"]);

  const cDesc2 = pickCol(hmap, ["Description2", "Description 2"]);
  const cPriority = pickCol(hmap, ["Priority"]);
  const cOperator = pickCol(hmap, ["Operator"]);

  if (!cEntity || !cWpc || !cDay7 || !cKpi || !cStatus || !cTower || !cTag || !cMocn) {
    throw new Error("WPC missing required columns: Entity_ID, WPC Name, Day-7, KPI D-1, Status, Tower ID, TAGGING, MOCN DATE");
  }

  ws.getRow(1).getCell(cMocn).value = "MOCN DATE";
  ws.getColumn(cMocn).numFmt = "dd/mm/yy";

  const needed = {
    dataKeys: new Set(),
    ippdKeys: new Set(),
    trafficKeys: new Set(),
    twampKeys: new Set(),
  };

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const entity = normalizeSpaces(cellText(row.getCell(cEntity).value));
    if (!entity) return;

    const kpiType = normalizeWpcName(row.getCell(cWpc).value);

    if (
      kpiType === KPI.AVG_CQI ||
      kpiType === KPI.AVG_DL_SE ||
      kpiType === KPI.S1_SETUP ||
      kpiType === KPI.UE_DL_TP ||
      kpiType === KPI.UE_UL_TP
    ) {
      needed.dataKeys.add(entity);
    } else if (kpiType === KPI.IPPD) {
      needed.ippdKeys.add(entity);
    } else if (kpiType === KPI.DL_TRAFFIC || kpiType === KPI.RRC_USERS) {
      needed.trafficKeys.add(entity);
    } else if (kpiType === KPI.TWAMP) {
      needed.twampKeys.add(entity);
    }
  });

  console.log("Load sitelist...");
  const { allMap: sitelistAll, sfMap: sitelistSf } = await loadSitelistMaps(SITELIST_PATH);

  console.log("Load tagging...");
  const taggingMap = await loadTaggingMap(TAGGING_PATH);

  console.log("Stream NEW SFXL...");
  const { dataMap, ippdMap, trafficMap, twampMap } = await loadSfxlMaps(SFXL_PATH, needed);

  console.log("Maps loaded:", {
    DATA: dataMap.size,
    IPPD: ippdMap.size,
    TRAFFIC_PLMN: trafficMap.size,
    TWAMP: twampMap.size,
    SITELIST_ALL: sitelistAll.size,
    SITELIST_SF: sitelistSf.size,
    TAGGING: taggingMap.size,
  });

  const count = {
    step1KpiFilled: 0,
    step1KpiNA: 0,
    step1StatusNormalized: 0,
    step2MocnFilled: 0,
    step2MocnNA: 0,
    step2StatusNySsh: 0,
    step3TagFilled: 0,
    step3TagNA: 0,
    step3SfMocnFilled: 0,
    step3SfMocnNA: 0,
    step3StatusClosed: 0,
  };

  console.log("STEP 1...");
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const entity = normalizeSpaces(cellText(row.getCell(cEntity).value));
    const kpiType = normalizeWpcName(row.getCell(cWpc).value);

    if (kpiType === KPI.OTHER || !entity) return;

    let value = null;
    let found = false;

    if (kpiType === KPI.AVG_CQI) {
      const rec = dataMap.get(entity);
      value = rec ? rec.avgCqi : null;
      found = value !== null && value !== undefined;
    } else if (kpiType === KPI.AVG_DL_SE) {
      const rec = dataMap.get(entity);
      value = rec ? rec.dlSe : null;
      found = value !== null && value !== undefined;
    } else if (kpiType === KPI.S1_SETUP) {
      const rec = dataMap.get(entity);
      value = rec ? rec.s1 : null;
      found = value !== null && value !== undefined;
    } else if (kpiType === KPI.UE_DL_TP) {
      const rec = dataMap.get(entity);
      value = rec ? rec.dlThr : null;
      found = value !== null && value !== undefined;
    } else if (kpiType === KPI.UE_UL_TP) {
      const rec = dataMap.get(entity);
      value = rec ? rec.ulThr : null;
      found = value !== null && value !== undefined;
    } else if (kpiType === KPI.IPPD) {
      value = ippdMap.get(entity);
      found = value !== undefined && value !== null;
    } else if (kpiType === KPI.DL_TRAFFIC) {
      const rec = trafficMap.get(entity);
      value = rec ? rec.payload : null;
      found = value !== null && value !== undefined;
    } else if (kpiType === KPI.RRC_USERS) {
      const rec = trafficMap.get(entity);
      value = rec ? rec.rrc : null;
      found = value !== null && value !== undefined;
    } else if (kpiType === KPI.TWAMP) {
      value = twampMap.get(entity);
      found = value !== undefined && value !== null;
    }

    if (found) {
      row.getCell(cKpi).value = value;
      count.step1KpiFilled += 1;
    } else {
      setExcelNA(row.getCell(cKpi));
      count.step1KpiNA += 1;
    }

    if (kpiType !== KPI.IPPD && kpiType !== KPI.TWAMP) {
      if (yesForStep1(kpiType, row.getCell(cKpi).value, row.getCell(cDay7).value)) {
        row.getCell(cStatus).value = "KPI Normalized";
        count.step1StatusNormalized += 1;
      }
      return;
    }

    const k = toNumber(row.getCell(cKpi).value);
    const d7 = toNumber(row.getCell(cDay7).value);
    const diff = k !== null && d7 !== null ? k - d7 : null;

    if (k !== null && diff !== null && k < 0.6 && diff < 0.2) {
      row.getCell(cStatus).value = "KPI Normalized";
      count.step1StatusNormalized += 1;
    }
  });

  console.log("STEP 2...");
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const kpiType = normalizeWpcName(row.getCell(cWpc).value);
    if (!STEP2_KPI_SET.has(kpiType)) return;

    const desc2 = cDesc2 ? normalizeSpaces(cellText(row.getCell(cDesc2).value)).toUpperCase() : "";
    const priority = cPriority ? normalizeSpaces(cellText(row.getCell(cPriority).value)).toUpperCase() : "";
    const operator = cOperator ? normalizeSpaces(cellText(row.getCell(cOperator).value)).toUpperCase() : "";

    if (!desc2.includes("SSH") || priority !== "P1" || operator !== "MOCN") return;

    const towerId = normalizeSpaces(cellText(row.getCell(cTower).value));
    const hit = sitelistAll.get(towerId);

    if (hit && !isBlankValue(hit.mocnDate)) {
      const d = parseDateLike(hit.mocnDate);
      row.getCell(cMocn).value = d || normalizeSpaces(cellText(hit.mocnDate));
      count.step2MocnFilled += 1;
    } else {
      setExcelNA(row.getCell(cMocn));
      count.step2MocnNA += 1;
    }

    if (!isBlankValue(row.getCell(cStatus).value)) return;

    const mocnDate = parseDateLike(row.getCell(cMocn).value);
    if (mocnDate && (mocnDate.getFullYear() === 2025 || mocnDate.getFullYear() === 2026)) {
      row.getCell(cStatus).value = "NY SSH Approval";
      count.step2StatusNySsh += 1;
    }
  });

  console.log("STEP 3...");
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const towerId = normalizeSpaces(cellText(row.getCell(cTower).value));
    if (!towerId) return;

    const remark = taggingMap.get(towerId);
    if (!isBlankValue(remark)) {
      row.getCell(cTag).value = remark;
      count.step3TagFilled += 1;
    } else {
      setExcelNA(row.getCell(cTag));
      count.step3TagNA += 1;
    }
  });

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const operator = cOperator ? normalizeSpaces(cellText(row.getCell(cOperator).value)).toUpperCase() : "";
    if (operator !== "SF") return;

    const tagText = cellText(row.getCell(cTag).value);
    if (!tagInDropDismantleNyOa(tagText)) return;

    const towerId = normalizeSpaces(cellText(row.getCell(cTower).value));
    const hit = sitelistSf.get(towerId);

    if (hit && !isBlankValue(hit.keepDrop)) {
      row.getCell(cMocn).value = normalizeSpaces(cellText(hit.keepDrop));
      count.step3SfMocnFilled += 1;
    } else {
      setExcelNA(row.getCell(cMocn));
      count.step3SfMocnNA += 1;
    }

    const mocnText = normalizeSpaces(cellText(row.getCell(cMocn).value)).toUpperCase();
    if (mocnText === "DROP" && isBlankValue(row.getCell(cStatus).value)) {
      row.getCell(cStatus).value = "Close due to site already drop";
      count.step3StatusClosed += 1;
    }
  });

  console.log("Summary:", count);
  console.log("Write output:", OUT_PATH);
  await wpcWb.xlsx.writeFile(OUT_PATH);
  console.log("DONE ✅ wpcsdm transform");
}

main().catch((err) => {
  console.error("ERROR:", err.message || err);
  process.exit(1);
});
