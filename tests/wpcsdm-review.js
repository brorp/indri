"use strict";

const path = require("path");
const ExcelJS = require("exceljs");

function getArg(flag) {
  const i = process.argv.indexOf(flag);
  return i >= 0 ? process.argv[i + 1] : null;
}

function norm(v) {
  return String(v ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function toComparable(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return `DATE:${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;

  if (typeof v === "object") {
    if (v.error) return norm(v.error);
    if (v.result !== undefined) return toComparable(v.result);
    if (typeof v.text === "string") return norm(v.text);
    if (Array.isArray(v.richText)) return norm(v.richText.map((x) => x.text || "").join(""));
  }

  if (typeof v === "number") return `NUM:${v}`;
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

function pickCol(hm, names) {
  for (const n of names) {
    const key = n.toUpperCase();
    if (hm.has(key)) return hm.get(key);
  }
  return null;
}

async function main() {
  const actualPath = getArg("--actual");
  const expectedPath = getArg("--expected");
  const sheetName = getArg("--sheet") || "wpcsdm_wpc_export";
  const sampleLimit = Number(getArg("--sample") || 30);

  if (!actualPath || !expectedPath) {
    console.error(
      "Usage: node tests/wpcsdm-review.js --actual out.xlsx --expected expected.xlsx [--sheet wpcsdm_wpc_export] [--sample 30]"
    );
    process.exit(1);
  }

  const actualAbs = path.resolve(process.cwd(), actualPath);
  const expectedAbs = path.resolve(process.cwd(), expectedPath);

  const wbA = new ExcelJS.Workbook();
  const wbE = new ExcelJS.Workbook();
  await wbA.xlsx.readFile(actualAbs);
  await wbE.xlsx.readFile(expectedAbs);

  const wsA = wbA.getWorksheet(sheetName) || wbA.worksheets[0];
  const wsE = wbE.getWorksheet(sheetName) || wbE.worksheets[0];

  if (!wsA || !wsE) {
    throw new Error("Sheet not found in one of files.");
  }

  const hA = headerMap(wsA);
  const hE = headerMap(wsE);

  const colsA = {
    ticket: pickCol(hA, ["WPC TICKET ID"]),
    entity: pickCol(hA, ["ENTITY_ID", "ENTITY ID"]),
    wpc: pickCol(hA, ["WPC NAME"]),
    kpi: pickCol(hA, ["KPI D-1"]),
    status: pickCol(hA, ["STATUS"]),
    tagging: pickCol(hA, ["TAGGING"]),
    mocn: pickCol(hA, ["MOCN DATE", "MOCN DATE "]),
  };

  const colsE = {
    ticket: pickCol(hE, ["WPC TICKET ID"]),
    entity: pickCol(hE, ["ENTITY_ID", "ENTITY ID"]),
    wpc: pickCol(hE, ["WPC NAME"]),
    kpi: pickCol(hE, ["KPI D-1"]),
    status: pickCol(hE, ["STATUS"]),
    tagging: pickCol(hE, ["TAGGING"]),
    mocn: pickCol(hE, ["MOCN DATE", "MOCN DATE "]),
  };

  if (!colsA.ticket || !colsE.ticket) {
    throw new Error("Column 'WPC Ticket ID' must exist in both files.");
  }

  const expectedRows = new Map();
  wsE.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const ticket = norm(row.getCell(colsE.ticket).value);
    if (!ticket) return;

    expectedRows.set(ticket, {
      rowNumber,
      entity: colsE.entity ? norm(row.getCell(colsE.entity).value) : "",
      wpc: colsE.wpc ? norm(row.getCell(colsE.wpc).value) : "",
      kpi: colsE.kpi ? toComparable(row.getCell(colsE.kpi).value) : "",
      status: colsE.status ? toComparable(row.getCell(colsE.status).value) : "",
      tagging: colsE.tagging ? toComparable(row.getCell(colsE.tagging).value) : "",
      mocn: colsE.mocn ? toComparable(row.getCell(colsE.mocn).value) : "",
    });
  });

  let totalActual = 0;
  let missingTicket = 0;
  const mismatch = { kpi: 0, status: 0, tagging: 0, mocn: 0 };
  const samples = [];

  wsA.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    totalActual += 1;

    const ticket = norm(row.getCell(colsA.ticket).value);
    if (!ticket) return;

    const expected = expectedRows.get(ticket);
    if (!expected) {
      missingTicket += 1;
      if (samples.length < sampleLimit) {
        samples.push({
          type: "missing-ticket",
          ticket,
          actualRow: rowNumber,
          detail: "ticket does not exist in expected file",
        });
      }
      return;
    }

    const entity = colsA.entity ? norm(row.getCell(colsA.entity).value) : "";
    const wpc = colsA.wpc ? norm(row.getCell(colsA.wpc).value) : "";

    const actualVals = {
      kpi: colsA.kpi ? toComparable(row.getCell(colsA.kpi).value) : "",
      status: colsA.status ? toComparable(row.getCell(colsA.status).value) : "",
      tagging: colsA.tagging ? toComparable(row.getCell(colsA.tagging).value) : "",
      mocn: colsA.mocn ? toComparable(row.getCell(colsA.mocn).value) : "",
    };

    for (const key of ["kpi", "status", "tagging", "mocn"]) {
      if (actualVals[key] !== expected[key]) {
        mismatch[key] += 1;
        if (samples.length < sampleLimit) {
          samples.push({
            type: key,
            ticket,
            actualRow: rowNumber,
            expectedRow: expected.rowNumber,
            entity,
            wpc,
            actual: actualVals[key],
            expected: expected[key],
          });
        }
      }
    }
  });

  console.log("WPCSDM Review Summary");
  console.log("---------------------");
  console.log(`Actual file   : ${actualAbs}`);
  console.log(`Expected file : ${expectedAbs}`);
  console.log(`Sheet         : ${wsA.name} (actual) vs ${wsE.name} (expected)`);
  console.log(`Total actual rows      : ${totalActual}`);
  console.log(`Missing ticket in expected: ${missingTicket}`);
  console.log(`Mismatch KPI D-1       : ${mismatch.kpi}`);
  console.log(`Mismatch Status        : ${mismatch.status}`);
  console.log(`Mismatch Tagging       : ${mismatch.tagging}`);
  console.log(`Mismatch MOCN DATE     : ${mismatch.mocn}`);

  if (samples.length) {
    console.log("\nSample mismatches:");
    for (const s of samples) {
      if (s.type === "missing-ticket") {
        console.log(`- [missing-ticket] ticket=${s.ticket} row=${s.actualRow} :: ${s.detail}`);
        continue;
      }

      console.log(
        `- [${s.type}] ticket=${s.ticket} entity=${s.entity} wpc=${s.wpc} rowA=${s.actualRow} rowE=${s.expectedRow} :: actual="${s.actual}" expected="${s.expected}"`
      );
    }
  } else {
    console.log("\nNo mismatches in sampled columns.");
  }
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : String(err));
  process.exit(1);
});
