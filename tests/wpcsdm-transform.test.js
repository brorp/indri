"use strict";

const assert = require("assert");
const fs = require("fs");
const fsp = require("fs/promises");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const ExcelJS = require("exceljs");

function norm(v) {
  return String(v ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function cellToComparable(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    if (v.error) return String(v.error).trim();
    if (v.result !== undefined) return cellToComparable(v.result);
    if (typeof v.text === "string") return norm(v.text);
    if (Array.isArray(v.richText)) return norm(v.richText.map((x) => x.text || "").join(""));
  }
  return norm(v);
}

function headerMap(ws) {
  const map = new Map();
  ws.getRow(1).eachCell((cell, col) => {
    const h = norm(cell.value).toUpperCase();
    if (h) map.set(h, col);
  });
  return map;
}

async function writeWpcFixture(filePath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("wpcsdm_wpc_export");

  ws.addRow([
    "WPC Ticket ID",
    "Title",
    "Ticket Status",
    "Date WPC",
    "Entity_ID",
    "WPC Name",
    "Priority",
    "Operator",
    "Day-7",
    "Day",
    "Delta(%)",
    "KPI D-1",
    "M",
    "N",
    "O",
    "TAGGING",
    "MOCN Date",
    "Status",
    "Description2",
    "Tower ID",
  ]);

  const rows = [
    ["T001", "a", "Open", "2026-03-02", "E1", "Avg CQI", "P2", "XL", 10, "", "", "", "", "", "", "", "", "", "", "T1"],
    ["T002", "a", "Open", "2026-03-02", "E2", "Avg DL SE", "P2", "XL", 10, "", "", "", "", "", "", "", "", "", "", "T2"],
    ["T003", "a", "Open", "2026-03-02", "E3", "S1 Set up success rate", "P2", "XL", 98, "", "", "", "", "", "", "", "", "", "", "T3"],
    ["T004", "a", "Open", "2026-03-02", "E4", "UE DL IP Throughput", "P2", "XL", 20, "", "", "", "", "", "", "", "", "", "", "T4"],
    ["T005", "a", "Open", "2026-03-02", "E5", "UE UL IP Throughput", "P2", "XL", 20, "", "", "", "", "", "", "", "", "", "", "T5"],
    ["T006", "a", "Open", "2026-03-02", "E6", "IPPD Packet Loss", "P2", "XL", 0.6, "", "", "", "", "", "", "", "", "", "", "T6"],
    ["T007", "a", "Open", "2026-03-02", "E7", "TWAMP Packet loss", "P2", "XL", 0.3, "", "", "", "", "", "", "", "", "", "", "T7"],
    ["T008", "a", "Open", "2026-03-02", "E8", "DL Traffic", "P2", "XL", 1000, "", "", "", "", "", "", "", "", "", "", "T8"],
    ["T009", "a", "Open", "2026-03-02", "E9", "RRC Conn Users", "P2", "XL", 100, "", "", "", "", "", "", "", "", "", "", "T9"],
    ["T010", "a", "Open", "2026-03-02", "E10", "Avg CQI", "P1", "MOCN", 10, "", "", "", "", "", "", "", "", "", "SSH", "TSSH"],
    ["T011", "a", "Open", "2026-03-02", "E11", "Avg CQI", "P1", "MOCN", 10, "", "", "", "", "", "", "", "", "", "SSH", "TSSH2"],
    ["T012", "a", "Open", "2026-03-02", "E12", "Avg DL SE", "P2", "SF", 10, "", "", "", "", "", "", "", "", "", "", "TSF1"],
    ["T013", "a", "Open", "2026-03-02", "E13", "Avg DL SE", "P2", "SF", 10, "", "", "", "", "", "", "", "", "", "", "TSF2"],
  ];

  for (const row of rows) ws.addRow(row);
  await wb.xlsx.writeFile(filePath);
}

async function writeSfxlFixture(filePath) {
  const wb = new ExcelJS.Workbook();

  const wsData = wb.addWorksheet("DATA");
  wsData.addRow([
    "Time",
    "MOEntity",
    "TowerID",
    "UL User Throughput",
    "DL User Throughput",
    "Avg CQI",
    "S1 Setup Success Rate",
    "DL SE",
  ]);

  const t = "2026-03-02 10:00";
  wsData.addRow([t, "E1", "T1", 0, 0, 11, null, null]);
  wsData.addRow([t, "E2", "T2", 0, 0, null, null, 8]);
  wsData.addRow([t, "E3", "T3", 0, 0, null, 99.5, null]);
  wsData.addRow([t, "E4", "T4", 0, 18, null, null, null]);
  wsData.addRow([t, "E10", "TSSH", 0, 0, 8, null, null]);
  wsData.addRow([t, "E11", "TSSH2", 0, 0, 11, null, null]);
  wsData.addRow([t, "E12", "TSF1", 0, 0, null, null, 8]);
  wsData.addRow([t, "E13", "TSF2", 0, 0, null, null, 8]);

  const wsIppd = wb.addWorksheet("IPPD");
  wsIppd.addRow(["Row Labels", "IPPD*100"]);
  wsIppd.addRow(["E6", 0.55]);

  const wsTraffic = wb.addWorksheet("TRAFFIC");
  wsTraffic.addRow(["Row Labels", "Sum of Payload per PLMN", "Sum of RRC User per PLMN"]);
  wsTraffic.addRow(["E8", 980, null]);
  wsTraffic.addRow(["E9", null, 90]);

  const wsTwamp = wb.addWorksheet("TWAMP");
  wsTwamp.addRow(["Row Labels", "Max of MAX TWAMP"]);
  wsTwamp.addRow(["E7", 0.5]);

  await wb.xlsx.writeFile(filePath);
}

async function writeSitelistCsvFixture(filePath) {
  const content = [
    "New XL ID,MOCN Date,Keep/Drop",
    "TSSH,01/05/2025,Keep",
    "TSSH2,15/03/2026,Keep",
    "TSF1,01/01/2024,Drop",
    "TSF2,02/01/2024,Keep",
  ].join("\n");
  await fsp.writeFile(filePath, content, "utf8");
}

async function writeTaggingFixture(filePath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("TAG");
  ws.addRow(["Tower ID", "Remark"]);
  ws.addRow(["TSF1", "Drop & Dismantle"]);
  ws.addRow(["TSF2", "NYOA"]);
  await wb.xlsx.writeFile(filePath);
}

function runTransform({ wpc, sfxl, sitelist, tagging, out }) {
  const script = path.resolve(__dirname, "..", "wpcsdm_transform.js");
  const result = spawnSync(
    process.execPath,
    [script, "--wpc", wpc, "--sfxl", sfxl, "--sitelist", sitelist, "--tagging", tagging, "--out", out],
    { encoding: "utf8" }
  );

  if (result.status !== 0) {
    throw new Error(
      `transform failed (exit=${result.status})\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`
    );
  }
}

async function readOutputMap(outputPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(outputPath);
  const ws = wb.getWorksheet("wpcsdm_wpc_export") || wb.worksheets[0];
  const hm = headerMap(ws);

  const cTicket = hm.get("WPC TICKET ID");
  const cKpi = hm.get("KPI D-1");
  const cStatus = hm.get("STATUS");
  const cTag = hm.get("TAGGING");
  const cMocn = hm.get("MOCN DATE");

  const out = new Map();
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const ticket = norm(row.getCell(cTicket).value);
    if (!ticket) return;
    out.set(ticket, {
      kpi: row.getCell(cKpi).value,
      status: row.getCell(cStatus).value,
      tagging: row.getCell(cTag).value,
      mocn: row.getCell(cMocn).value,
    });
  });
  return out;
}

function assertTextEq(actual, expected, msg) {
  assert.strictEqual(norm(cellToComparable(actual)), norm(expected), msg);
}

async function main() {
  const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "indri-wpc-test-"));
  try {
    const wpc = path.join(tempDir, "wpcsdm_wpc_export_02032026_default.xlsx");
    const sfxl = path.join(tempDir, "NEW SFXL 02032026.xlsx");
    const sitelist = path.join(tempDir, "sitelist_mocn_02032026.csv");
    const tagging = path.join(tempDir, "TAGGING 02032026.xlsx");
    const out = path.join(tempDir, "wpcsdm_wpc_export_02032026_result.xlsx");

    await writeWpcFixture(wpc);
    await writeSfxlFixture(sfxl);
    await writeSitelistCsvFixture(sitelist);
    await writeTaggingFixture(tagging);

    runTransform({ wpc, sfxl, sitelist, tagging, out });
    const rows = await readOutputMap(out);

    assertTextEq(rows.get("T001").kpi, "11", "T001 KPI");
    assertTextEq(rows.get("T001").status, "KPI Normalized", "T001 status");

    assertTextEq(rows.get("T002").kpi, "8", "T002 KPI");
    assertTextEq(rows.get("T002").status, "", "T002 status should stay blank");

    assertTextEq(rows.get("T003").kpi, "99.5", "T003 KPI");
    assertTextEq(rows.get("T003").status, "KPI Normalized", "T003 status");

    assertTextEq(rows.get("T004").kpi, "18", "T004 KPI");
    assertTextEq(rows.get("T004").status, "", "T004 status (ratio -0.1) should blank");

    assertTextEq(rows.get("T005").kpi, "#N/A", "T005 should be #N/A when lookup miss");

    assertTextEq(rows.get("T006").kpi, "0.55", "T006 KPI");
    assertTextEq(rows.get("T006").status, "KPI Normalized", "T006 IPPD status");

    assertTextEq(rows.get("T007").kpi, "0.5", "T007 KPI");
    assertTextEq(rows.get("T007").status, "", "T007 TWAMP status should blank");

    assertTextEq(rows.get("T008").kpi, "980", "T008 KPI");
    assertTextEq(rows.get("T008").status, "KPI Normalized", "T008 status");

    assertTextEq(rows.get("T009").kpi, "90", "T009 KPI");
    assertTextEq(rows.get("T009").status, "", "T009 status should blank");

    const mocn10 = rows.get("T010").mocn;
    assert(mocn10 instanceof Date, "T010 MOCN should be date");
    assert.strictEqual(mocn10.getFullYear(), 2025, "T010 MOCN year");
    assertTextEq(rows.get("T010").status, "NY SSH Approval", "T010 status");

    const mocn11 = rows.get("T011").mocn;
    assert(mocn11 instanceof Date, "T011 MOCN should be date");
    assert.strictEqual(mocn11.getFullYear(), 2026, "T011 MOCN year");
    assertTextEq(rows.get("T011").status, "KPI Normalized", "T011 should keep KPI Normalized");

    assertTextEq(rows.get("T012").tagging, "Drop & Dismantle", "T012 tagging");
    assertTextEq(rows.get("T012").mocn, "Drop", "T012 MOCN keep/drop");
    assertTextEq(rows.get("T012").status, "Close due to site already drop", "T012 status");

    assertTextEq(rows.get("T013").tagging, "NYOA", "T013 tagging");
    assertTextEq(rows.get("T013").mocn, "Keep", "T013 MOCN keep/drop");
    assertTextEq(rows.get("T013").status, "", "T013 status remains blank");

    assertTextEq(rows.get("T001").tagging, "#N/A", "missing tagging should be #N/A");

    console.log("PASS wpcsdm-transform test");
  } finally {
    await fsp.rm(tempDir, { recursive: true, force: true });
  }
}

main().catch((err) => {
  console.error("FAIL wpcsdm-transform test");
  console.error(err && err.stack ? err.stack : String(err));
  process.exit(1);
});
