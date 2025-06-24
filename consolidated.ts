// Define strict cell type for inference
type CellVal = string | number | boolean | Date;

/**
 * Full Office Script Pipeline:
 * 1. normalizeAB: normalize A-B on Main
 * 2. normalizeCD: format C-D on Main
 * 3. populateEF: copy from GFE into E-F
 * 4. flagGH: flag changes in G-H
 */
function main(workbook: ExcelScript.Workbook): void {
  const SHEET_NAME: string = "Main";
  const GFE_SHEET: string = "GFE";
  const START_ROW: number = 2;

  const wsMain: ExcelScript.Worksheet = workbook.getWorksheet(SHEET_NAME)!;
  const wsGfe: ExcelScript.Worksheet = workbook.getWorksheet(GFE_SHEET)!;

  normalizeAB(wsMain, START_ROW);
  normalizeCD(wsMain, START_ROW);
  populateEF(wsMain, wsGfe, START_ROW);
  flagGH(wsMain, START_ROW);
}

// 1. Normalize columns A & B
function normalizeAB(ws: ExcelScript.Worksheet, startRow: number): void {
  const usedRange: ExcelScript.Range = ws.getUsedRange()!;
  const totalRows: number = usedRange.getRowCount();
  const rng: ExcelScript.Range = ws.getRangeByIndexes(startRow, 0, totalRows - startRow, 2);
  const vals: CellVal[][] = rng.getValues() as CellVal[][];

  for (let i: number = 0; i < vals.length; i++) {
    // Column A
    let matVal: CellVal = vals[i][0];
    if (typeof matVal === 'string') {
      const txt: string = matVal.replace(/^'/, '').trim();
      const num: number = Number(txt);
      matVal = !isNaN(num) ? num : txt;
    }
    // Column B
    let seaVal: CellVal = vals[i][1];
    if (typeof seaVal === 'string') {
      seaVal = seaVal.trim().toUpperCase();
    }
    vals[i][0] = matVal;
    vals[i][1] = seaVal;
  }

  rng.setValues(vals);
}

// 2. Normalize & format columns C-D
function normalizeCD(ws: ExcelScript.Worksheet, startRow: number): void {
  const usedRange: ExcelScript.Range = ws.getUsedRange()!;
  const totalRows: number = usedRange.getRowCount();
  const rng: ExcelScript.Range = ws.getRangeByIndexes(startRow, 2, totalRows - startRow, 2);
  const vals: CellVal[][] = rng.getValues() as CellVal[][];
  const out: string[][] = [];

  for (let i: number = 0; i < vals.length; i++) {
    const row: CellVal[] = vals[i];
    const formattedRow: string[] = ["", ""];
    for (let j: number = 0; j < 2; j++) {
      const v: CellVal = row[j];
      let dt: Date | null = null;
      if (v instanceof Date) {
        dt = v;
      } else if (typeof v === 'number') {
        dt = excelSerialToDate(v);
      } else if (typeof v === 'string' && v.trim() !== "") {
        const parsed: Date = new Date(v.trim());
        if (!isNaN(parsed.getTime())) dt = parsed;
      }
      if (dt) {
        formattedRow[j] = formatDate(dt);
      }
    }
    out.push(formattedRow);
  }

  const fmt: string[][] = Array(out.length).fill(["@", "@"]);
  rng.setNumberFormatLocal(fmt);
  rng.setValues(out);
}

function excelSerialToDate(serial: number): Date {
  const offset: number = serial > 60 ? serial - 1 : serial;
  const epoch: Date = new Date(Date.UTC(1899, 11, 31));
  return new Date(epoch.getTime() + offset * 86400000);
}

function formatDate(d: Date): string {
  const mm: string = String(d.getMonth() + 1).padStart(2, '0');
  const dd: string = String(d.getDate()).padStart(2, '0');
  const yyyy: number = d.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
}

// 3. Populate columns E-F from GFE with INACTIVE logic
function populateEF(
  wsMain: ExcelScript.Worksheet,
  wsGfe: ExcelScript.Worksheet,
  startRow: number
): void {
  const gfeUsed: ExcelScript.Range = wsGfe.getUsedRange()!;
  const gfeData: CellVal[][] = gfeUsed.getValues() as CellVal[][];
  const lookup: Record<string, { fa: CellVal; disco: CellVal; status: string }> = {};

  for (let i: number = 1; i < gfeData.length; i++) {
    const matKey: string = normalizeKey(gfeData[i][0]);
    const seaKey: string = normalizeKey(gfeData[i][1]);
    const statusVal: string = String(gfeData[i][4]).trim().toUpperCase();
    lookup[`${matKey}|${seaKey}`] = {
      fa: gfeData[i][2],
      disco: gfeData[i][3],
      status: statusVal
    };
  }

  const used: ExcelScript.Range = wsMain.getUsedRange()!;
  const totalRows: number = used.getRowCount();
  const out: string[][] = [];

  for (let r: number = startRow; r < totalRows; r++) {
    const matCell: CellVal = wsMain.getCell(r, 0).getValue() as CellVal;
    const seaCell: CellVal = wsMain.getCell(r, 1).getValue() as CellVal;
    const key: string = `${normalizeKey(matCell)}|${normalizeKey(seaCell)}`;
    const entry = lookup[key];
    let faStr: string = "";
    let discoStr: string = "";
    if (entry) {
      if (entry.status !== "ACTIVE") {
        faStr = "INACTIVE";
        discoStr = "INACTIVE";
      } else {
        const d1: Date | null = parseDate(entry.fa);
        const d2: Date | null = parseDate(entry.disco);
        if (d1) faStr = formatDate(d1);
        if (d2) discoStr = formatDate(d2);
      }
    }
    out.push([faStr, discoStr]);
  }

  const rng: ExcelScript.Range = wsMain.getRangeByIndexes(startRow, 4, out.length, 2);
  const fmt: string[][] = Array(out.length).fill(["@", "@"]); rng.setNumberFormatLocal(fmt);
  rng.setValues(out);
}

function normalizeKey(val: CellVal): string {
  return String(val).trim().toLowerCase();
}

// 4. Flag changes in columns G-H
function flagGH(ws: ExcelScript.Worksheet, startRow: number): void {
  const used: ExcelScript.Range = ws.getUsedRange()!;
  const totalRows: number = used.getRowCount();
  const out: string[][] = [];

  for (let r: number = startRow; r < totalRows; r++) {
    const cVal: CellVal = ws.getCell(r, 2).getValue() as CellVal;
    const eVal: CellVal = ws.getCell(r, 4).getValue() as CellVal;
    const dVal: CellVal = ws.getCell(r, 3).getValue() as CellVal;
    const fVal: CellVal = ws.getCell(r, 5).getValue() as CellVal;

    const msC: number | "" = toMs(cVal);
    const msE: number | "" = toMs(eVal);
    const msD: number | "" = toMs(dVal);
    const msF: number | "" = toMs(fVal);

    out.push([
      msC !== msE ? "Y" : "",
      msD !== msF ? "Y" : ""
    ]);
  }

  const rng: ExcelScript.Range = ws.getRangeByIndexes(startRow, 6, out.length, 2);
  rng.setValues(out);
}

function toMs(val: CellVal): number | "" {
  if (val == null || val === "") return "";
  if (val instanceof Date) return val.getTime();
  if (typeof val === 'number') return val;
  const parsed: Date = new Date(String(val));
  return isNaN(parsed.getTime()) ? "" : parsed.getTime();
}
