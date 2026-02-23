"use strict";

/**
 * wpcsdm_step1_strict.js
 *
 * STEP 1 ONLY (Strict VLOOKUP replication)
 * EDIT ONLY:
 * - KPI D-1
 * - Status
 *
 * DO NOT TOUCH:
 * - Column M/N/O (or other columns)
 *
 * Exact lookup (range_lookup=0). Miss => Excel error #N/A.
 *
 * Run:
 * node --max-old-space-size=8192 wpcsdm_step1_strict.js \
 *   --wpc "wpcsdm_wpc_export_20260214144337_default.xlsx" \
 *   --sfxl "NEW SFXL 14022026.xlsx" \
 *   --out "wpcsdm_step1_out.xlsx"
 */

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

// ---------------- CLI ----------------
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
const OUT_PATH = abs(getArg("--out"));
if (!OUT_PATH) throw new Error("Missing --out");

// ---------------- CONFIG ----------------
const SHEET_WPC = "wpcsdm_wpc_export";

// WPC Name normalized to UPPER for matching
const WPC = {
  AVG_CQI: "AVG CQI",
  AVG_DL_SE: "AVG DL SE",
  S1: "S1 SET UP SUCCESS RATE (%)",
  UE_DL: "UE DL IP THROUGHPUT",
  UE_UL: "UE UL IP THROUGHPUT",
  IPPD: "IPPD PACKET LOSS",
  DL_TRAFFIC: "DL TRAFFIC",
  RRC: "RRC CONN USERS",
  TWAMP: "TWAMP PACKET LOSS", // file bisa "TWAMP" saja -> kita handle juga
};

// ---------------- Helpers ----------------
function normText(x) {
  // minimal normalization (safe): trim + remove NBSP/zero-width
  return String(x ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .trim();
}

function cellText(v) {
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
    if (v.error) return false;
    if (typeof v.text === "string") return v.text.trim() === "";
    if (Array.isArray(v.richText)) return v.richText.map(x => x.text || "").join("").trim() === "";
  }
  return String(v).trim() === "";
}

function setExcelNA(cell) {
  cell.value = { error: "#N/A" }; // real Excel error cell
}

function normHeader(h) {
  return cellText(h).replace(/\s+/g, " ").trim().toUpperCase();
}

function buildHeaderIndex(row) {
  const map = new Map();
  row.eachCell((cell, i) => {
    const h = normHeader(cell.value);
    if (h) map.set(h, i);
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

// Find a column by "contains" rules (for weird header variations)
function findColContains(hmap, mustContainWords) {
  // mustContainWords: ["IPPD", "100"] etc
  for (const [h, idx] of hmap.entries()) {
    const ok = mustContainWords.every(w => h.includes(w));
    if (ok) return idx;
  }
  return null;
}

/**
 * Reproduce Column O "Yes/Open" logic AFTER KPI D-1 is updated (without touching col O).
 * This is used for Step 10-11 (filter O == Yes then Status = KPI Normalized) for all except IPPD/TWAMP.
 */
function yesOpenByFormula(wpcNameUpper, kpiD1, day7) {
  const k = num(kpiD1);
  const d7 = num(day7);
  if (k === null) return "OPEN";

  const diff = (d7 !== null) ? (k - d7) : null;
  const ratio = (diff !== null && d7 !== null && d7 !== 0) ? (diff / d7) : null;

  if (wpcNameUpper === WPC.DL_TRAFFIC) {
    // M > -50 AND N > -10%
    if (diff !== null && ratio !== null && diff > -50 && ratio > -0.10) return "YES";
    return "OPEN";
  }

  if (wpcNameUpper === WPC.S1) {
    return (k > 99) ? "YES" : "OPEN";
  }

  // for these KPIs: N > -10%
  const ratioKpis = new Set([WPC.AVG_CQI, WPC.AVG_DL_SE, WPC.UE_DL, WPC.UE_UL, WPC.RRC]);
  if (ratioKpis.has(wpcNameUpper)) {
    if (ratio !== null && ratio > -0.10) return "YES";
    return "OPEN";
  }

  return "OPEN";
}

// ---------------- Stream NEW SFXL lookup tables ----------------
async function loadSfxlLookupsStrict(sfxlPath, needed) {
  const maps = {
    DATA: new Map(),    // MOEntity -> {avgCqi, dlSe, s1, dlThr, ulThr}
    IPPD: new Map(),    // Row Labels -> ippd100
    PLMN: new Map(),    // Row Labels -> {payload, rrc}
    TWAMP: new Map(),   // Row Labels -> twamp
  };

  let hasSheet = { DATA: false, IPPD: false, PLMN: false, TWAMP: false };

  const reader = new ExcelJS.stream.xlsx.WorkbookReader(sfxlPath, {
    entries: "emit",
    sharedStrings: "cache", // IMPORTANT for string keys
    styles: "ignore",
    hyperlinks: "ignore",
    worksheets: "emit",
  });

  for await (const ws of reader) {
    const name = normText(ws.name).toUpperCase();

    // Sheet selection:
    // DATA => "DATA"
    // IPPD => "IPPD"
    // PLMN/TRAFFIC => "PLMN" or "TRAFFIC"
    // TWAMP => "TWAMP"
    let type = null;
    if (name === "DATA") type = "DATA";
    else if (name === "IPPD") type = "IPPD";
    else if (name === "PLMN" || name === "TRAFFIC") type = "PLMN";
    else if (name === "TWAMP") type = "TWAMP";
    else type = null;

    if (!type) continue;
    hasSheet[type] = true;

    let hmap = null;

    for await (const row of ws) {
      if (row.number === 1) {
        hmap = buildHeaderIndex(row);
        continue;
      }
      if (!hmap) continue;

      if (type === "DATA") {
        const cKey = pickCol(hmap, ["MOEntity"]);
        if (!cKey) continue;

        const key = normText(row.getCell(cKey).value); // exact match style
        if (!key || !needed.dataKeys.has(key)) continue;

        // we strictly pick by headers
        const cAvgCqi = pickCol(hmap, ["Avg CQI"]);
        const cDlSe = pickCol(hmap, ["DL SE"]);
        const cS1 = pickCol(hmap, ["S1 Setup Success Rate"]);
        const cDlThr = pickCol(hmap, ["DL User Throughput"]);
        const cUlThr = pickCol(hmap, ["UL User Throughput"]);

        maps.DATA.set(key, {
          avgCqi: cAvgCqi ? num(row.getCell(cAvgCqi).value) : null,
          dlSe: cDlSe ? num(row.getCell(cDlSe).value) : null,
          s1: cS1 ? num(row.getCell(cS1).value) : null,
          dlThr: cDlThr ? num(row.getCell(cDlThr).value) : null,
          ulThr: cUlThr ? num(row.getCell(cUlThr).value) : null,
        });
      }

      if (type === "IPPD") {
        // strict: Row Labels -> IPPD*100
        const cKey = pickCol(hmap, ["Row Labels"]);
        if (!cKey) continue;

        const key = normText(row.getCell(cKey).value);
        if (!key || !needed.ippdKeys.has(key)) continue;

        // find IPPD*100 column
        let cVal = pickCol(hmap, ["IPPD*100"]);
        if (!cVal) cVal = findColContains(hmap, ["IPPD", "100"]);
        if (!cVal) cVal = findColContains(hmap, ["IPPD", "%"]);

        if (!cVal) continue;
        maps.IPPD.set(key, num(row.getCell(cVal).value));
      }

      if (type === "PLMN") {
        const cKey = pickCol(hmap, ["Row Labels"]);
        if (!cKey) continue;

        const key = normText(row.getCell(cKey).value);
        if (!key || !needed.plmnKeys.has(key)) continue;

        const cPayload = pickCol(hmap, ["Sum of Payload per PLMN"]);
        const cRrc = pickCol(hmap, ["Sum of RRC User per PLMN"]);

        maps.PLMN.set(key, {
          payload: cPayload ? num(row.getCell(cPayload).value) : null,
          rrc: cRrc ? num(row.getCell(cRrc).value) : null,
        });
      }

      if (type === "TWAMP") {
        const cKey = pickCol(hmap, ["Row Labels"]);
        const cVal = pickCol(hmap, ["Max of MAX TWAMP"]);
        if (!cKey || !cVal) continue;

        const key = normText(row.getCell(cKey).value);
        if (!key || !needed.twampKeys.has(key)) continue;

        maps.TWAMP.set(key, num(row.getCell(cVal).value));
      }
    }
  }

  return { maps, hasSheet };
}

// ---------------- MAIN (STEP 1) ----------------
async function main() {
  console.log("Load WPC file...");
  const wpcWb = new ExcelJS.Workbook();
  await wpcWb.xlsx.readFile(WPC_PATH);

  // Let Excel recalc formulas on open (M/N/O will update when you open file)
  wpcWb.calcProperties = wpcWb.calcProperties || {};
  wpcWb.calcProperties.fullCalcOnLoad = true;

  const ws = wpcWb.getWorksheet(SHEET_WPC);
  if (!ws) throw new Error(`Missing sheet '${SHEET_WPC}' in WPC`);

  const headerRow = ws.getRow(1);
  const hmap = buildHeaderIndex(headerRow);

  const cEntity = pickCol(hmap, ["Entity_ID"]);
  const cWpc = pickCol(hmap, ["WPC Name"]);
  const cDay7 = pickCol(hmap, ["Day-7"]);
  const cKpi = pickCol(hmap, ["KPI D-1"]);
  const cStatus = pickCol(hmap, ["Status"]);

  if (!cEntity || !cWpc || !cDay7 || !cKpi || !cStatus) {
    throw new Error("WPC missing required columns: Entity_ID, WPC Name, Day-7, KPI D-1, Status");
  }

  // Collect needed keys (exact) for each NEW SFXL sheet lookup
  const needed = {
    dataKeys: new Set(),
    ippdKeys: new Set(),
    plmnKeys: new Set(),
    twampKeys: new Set(),
  };

  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === 1) return;

    const entity = normText(row.getCell(cEntity).value);
    if (!entity) return;

    const wpcNameUpper = normText(row.getCell(cWpc).value).toUpperCase();

    // DATA (MOEntity based)
    if (
      wpcNameUpper === WPC.AVG_CQI ||
      wpcNameUpper === WPC.AVG_DL_SE ||
      wpcNameUpper === WPC.S1 ||
      wpcNameUpper === WPC.UE_DL ||
      wpcNameUpper === WPC.UE_UL
    ) {
      needed.dataKeys.add(entity);
    }

    // IPPD (Row Labels based)
    if (wpcNameUpper === WPC.IPPD) needed.ippdKeys.add(entity);

    // PLMN/TRAFFIC (Row Labels based)
    if (wpcNameUpper === WPC.DL_TRAFFIC || wpcNameUpper === WPC.RRC) needed.plmnKeys.add(entity);

    // TWAMP (Row Labels based)
    if (wpcNameUpper === WPC.TWAMP || wpcNameUpper === "TWAMP") needed.twampKeys.add(entity);
  });

  console.log("Collect keys:", {
    DATA: needed.dataKeys.size,
    IPPD: needed.ippdKeys.size,
    PLMN_TRAFFIC: needed.plmnKeys.size,
    TWAMP: needed.twampKeys.size,
  });

  console.log("Load NEW SFXL lookup tables (stream)...");
  const { maps, hasSheet } = await loadSfxlLookupsStrict(SFXL_PATH, needed);

  console.log("Sheet availability:", hasSheet);
  console.log("Loaded map sizes:", {
    DATA: maps.DATA.size,
    IPPD: maps.IPPD.size,
    PLMN_TRAFFIC: maps.PLMN.size,
    TWAMP: maps.TWAMP.size,
  });

  // Apply Step 1
  let kpiFilled = 0;
  let kpiNA = 0;
  let statusNormalized = 0;

  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === 1) return;

    const entity = normText(row.getCell(cEntity).value);
    const wpcNameUpper = normText(row.getCell(cWpc).value).toUpperCase();
    const day7 = row.getCell(cDay7).value;

    // ---- 1-9: Fill KPI D-1 strictly per WPC Name ----
    let kpiValue = null;
    let found = false;

    if (wpcNameUpper === WPC.AVG_CQI) {
      const rec = maps.DATA.get(entity);
      if (rec && rec.avgCqi !== null && rec.avgCqi !== undefined) {
        kpiValue = rec.avgCqi; found = true;
      }
    } else if (wpcNameUpper === WPC.AVG_DL_SE) {
      const rec = maps.DATA.get(entity);
      if (rec && rec.dlSe !== null && rec.dlSe !== undefined) {
        kpiValue = rec.dlSe; found = true;
      }
    } else if (wpcNameUpper === WPC.S1) {
      const rec = maps.DATA.get(entity);
      if (rec && rec.s1 !== null && rec.s1 !== undefined) {
        kpiValue = rec.s1; found = true;
      }
    } else if (wpcNameUpper === WPC.UE_DL) {
      const rec = maps.DATA.get(entity);
      if (rec && rec.dlThr !== null && rec.dlThr !== undefined) {
        kpiValue = rec.dlThr; found = true;
      }
    } else if (wpcNameUpper === WPC.UE_UL) {
      const rec = maps.DATA.get(entity);
      if (rec && rec.ulThr !== null && rec.ulThr !== undefined) {
        kpiValue = rec.ulThr; found = true;
      }
    } else if (wpcNameUpper === WPC.IPPD) {
      const v = maps.IPPD.get(entity);
      if (v !== undefined && v !== null) {
        kpiValue = v; found = true;
      }
    } else if (wpcNameUpper === WPC.DL_TRAFFIC) {
      const rec = maps.PLMN.get(entity);
      if (rec && rec.payload !== null && rec.payload !== undefined) {
        kpiValue = rec.payload; found = true;
      }
    } else if (wpcNameUpper === WPC.RRC) {
      const rec = maps.PLMN.get(entity);
      if (rec && rec.rrc !== null && rec.rrc !== undefined) {
        kpiValue = rec.rrc; found = true;
      }
    } else if (wpcNameUpper === WPC.TWAMP || wpcNameUpper === "TWAMP") {
      const v = maps.TWAMP.get(entity);
      if (v !== undefined && v !== null) {
        kpiValue = v; found = true;
      }
    } else {
      // Not part of Step 1 lookup list -> leave KPI D-1 unchanged
      return;
    }

    // Write KPI D-1: number if found else #N/A
    if (found) {
      row.getCell(cKpi).value = kpiValue;
      kpiFilled++;
    } else {
      setExcelNA(row.getCell(cKpi));
      kpiNA++;
    }

    // ---- 10-11: Non IPPD/TWAMP -> if O would be YES -> Status = KPI Normalized ----
    // We'll compute YES using the same formula logic based on updated KPI D-1 and Day-7
    if (wpcNameUpper !== WPC.IPPD && wpcNameUpper !== WPC.TWAMP && wpcNameUpper !== "TWAMP") {
      const yesOpen = yesOpenByFormula(wpcNameUpper, found ? kpiValue : null, day7);
      if (yesOpen === "YES") {
        row.getCell(cStatus).value = "KPI Normalized";
        statusNormalized++;
      }
      return;
    }

    // ---- 12-13: IPPD & TWAMP -> KPI D-1 < 0.6 AND (KPI D-1 - Day-7) < 0.2 -> Status = KPI Normalized ----
    const k = found ? num(kpiValue) : null;
    const d7 = num(day7);
    const diff = (k !== null && d7 !== null) ? (k - d7) : null;

    if (k !== null && diff !== null && k < 0.6 && diff < 0.2) {
      row.getCell(cStatus).value = "KPI Normalized";
      statusNormalized++;
    }
  });

  console.log("STEP 1 summary:");
  console.log(" - KPI D-1 filled (number):", kpiFilled);
  console.log(" - KPI D-1 set to #N/A:", kpiNA);
  console.log(" - Status set KPI Normalized:", statusNormalized);

  console.log("Write output:", OUT_PATH);
  await wpcWb.xlsx.writeFile(OUT_PATH);
  console.log("DONE ✅ STEP 1");
}

main().catch((e) => {
  console.error("ERROR:", e.message || e);
  process.exit(1);
});
