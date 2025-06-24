// Define strict cell type for inference
type CellVal = string | number | boolean | Date;

/**
 * Full pipeline:
 * 1. normalizeAB: columns A (Material) & B (Season)
 * 2. normalizeCD: columns C-D (date parsing & formatting)
 * 3. populateEF: copy from GFE into E-F with INACTIVE logic
 * 4. flagGH: compare and flag changes in G-H
 */
function main(workbook: ExcelScript.Workbook): void {
  const SHEET = "Main";
  const GFE_SHEET = "GFE";
  const START_ROW = 2;
  const MAIN_COL_COUNT = 8; // up to H

  const ws = workbook.getWorksheet(SHEET);
  const wsGfe = workbook.getWorksheet(GFE_SHEET);
  if (!ws || !wsGfe) return;

  normalizeAB(ws, START_ROW);
  normalizeCD(ws, START_ROW);
  populateEF(ws, wsGfe, START_ROW);
  flagGH(ws, START_ROW);
}

// 1. Normalize A (Material) & B (Season)
function normalizeAB(ws: ExcelScript.Worksheet, startRow: number): void {
  const used = ws.getUsedRange(); if (!used) return;
  const rows = used.getRowCount();
  const range = ws.getRangeByIndexes(startRow, 0, rows - startRow, 2);
  const vals = range.getValues() as CellVal[][];
  for (let i = 0; i < vals.length; i++) {
    // Material
    let m = vals[i][0];
    if (typeof m === 'string') {
      const t = m.replace(/^'/,'').trim();
      const n = Number(t);
      m = !isNaN(n) ? n : t;
    }
    // Season
    let s = vals[i][1];
    if (typeof s === 'string') s = s.trim().toUpperCase();
    vals[i][0] = m;
    vals[i][1] = s;
  }
  range.setValues(vals);
}

// 2. Normalize & format C-D
function normalizeCD(ws: ExcelScript.Worksheet, startRow: number): void {
  const used = ws.getUsedRange(); if (!used) return;
  const rows = used.getRowCount();
  const range = ws.getRangeByIndexes(startRow, 2, rows - startRow, 2);
  const vals = range.getValues() as CellVal[][];
  const out: string[][] = [];
  for (const row of vals) {
    out.push(row.map(v => formatCellDate(v)));
  }
  const fmt = Array(out.length).fill(["@","@"]); range.setNumberFormatLocal(fmt);
  range.setValues(out);
}
function formatCellDate(v: CellVal): string {
  const d = parseDate(v);
  return d ? formatDate(d) : "";
}
function parseDate(v: CellVal): Date | null {
  if (v instanceof Date) return v;
  if (typeof v === 'number') return excelSerialToDate(v);
  if (typeof v === 'string' && v.trim()) {
    const d = new Date(v.trim());
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}
function excelSerialToDate(n: number): Date {
  const d = n > 60 ? n - 1 : n;
  const e = new Date(Date.UTC(1899,11,31));
  return new Date(e.getTime() + d*86400000);
}
function formatDate(d: Date): string {
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const dd = String(d.getDate()).padStart(2,'0');
  return `${mm}/${dd}/${d.getFullYear()}`;
}

// 3. Populate E-F from GFE
function populateEF(ws: ExcelScript.Worksheet, wsGfe: ExcelScript.Worksheet, startRow: number): void {
  const gfeUsed = wsGfe.getUsedRange(); if (!gfeUsed) return;
  const gfe = gfeUsed.getValues() as CellVal[][];
  const lookup: Record<string,{fa:CellVal,disco:CellVal,status:string}> = {};
  for (let i = 1; i < gfe.length; i++) {
    const key = normalizeKey(gfe[i][0])+'|'+normalizeKey(gfe[i][1]);
    lookup[key] = {fa:gfe[i][2], disco:gfe[i][3], status:String(gfe[i][4]).trim().toUpperCase()};
  }
  const used = ws.getUsedRange(); if (!used) return;
  const rows = used.getRowCount();
  const out: string[][] = [];
  for (let r = startRow; r < rows; r++) {
    const m = normalizeKey(ws.getCell(r,0).getValue() as CellVal);
    const s = normalizeKey(ws.getCell(r,1).getValue() as CellVal);
    const e = lookup[m+'|'+s];
    let fa="",di="";
    if (e) {
      if (e.status!=='ACTIVE') fa=di='INACTIVE'; else {
        const d1=parseDate(e.fa), d2=parseDate(e.disco);
        if (d1) fa=formatDate(d1);
        if (d2) di=formatDate(d2);
      }
    }
    out.push([fa,di]);
  }
  const range = ws.getRangeByIndexes(startRow,4,out.length,2);
  const fmt = Array(out.length).fill(["@","@"]); range.setNumberFormatLocal(fmt);
  range.setValues(out);
}

// 4. Flag G-H
function flagGH(ws: ExcelScript.Worksheet, startRow: number): void {
  const used = ws.getUsedRange(); if (!used) return;
  const rows = used.getRowCount();
  const out: string[][] = [];
  for (let r = startRow; r < rows; r++) {
    const c = toMs(ws.getCell(r,2).getValue() as CellVal);
    const e = toMs(ws.getCell(r,4).getValue() as CellVal);
    const d = toMs(ws.getCell(r,3).getValue() as CellVal);
    const f = toMs(ws.getCell(r,5).getValue() as CellVal);
    out.push([c!==e?'Y':'', d!==f?'Y':'']);
  }
  const range = ws.getRangeByIndexes(startRow,6,out.length,2);
  range.setValues(out);
}

function normalizeKey(val: CellVal): string {
  return String(val).trim().toLowerCase();
}
function toMs(v: CellVal): number|"" {
  if (v==null||v==="") return "";
  if (v instanceof Date) return v.getTime();
  if (typeof v==='number') return v;
  const d = new Date(String(v));
  return isNaN(d.getTime())?"":d.getTime();
}
