/**
 * Compare dates between «Main» and «GFE» sheets.
 *
 * 1. Copies GFE dates (cols C–D) → Main (cols E–F) for each “Material + Season”.
 * 2. Writes "Y" in G (FA changed) and H (Disco changed) when dates differ.
 *
 * All variables are explicitly typed; no implicit `any` remains.
 */
function main(workbook: ExcelScript.Workbook): void {
  // ───────── Constants & worksheets ─────────
  const START_ROW: number = 2; // 0‑based index (row 3)
  const wsMain: ExcelScript.Worksheet | undefined = workbook.getWorksheet("Main");
  const wsGfe:  ExcelScript.Worksheet | undefined = workbook.getWorksheet("GFE");
  if (!wsMain || !wsGfe) return;

  // ───────── Read GFE and build lookup ─────────
  const gfeRange: ExcelScript.Range | undefined = wsGfe.getUsedRange();
  if (!gfeRange) return;
  const gfeVals: ExcelScript.CellValue[][] = gfeRange.getValues() as ExcelScript.CellValue[][];

  type DatesPair = { fa: ExcelScript.CellValue; disco: ExcelScript.CellValue };
  const map: Map<string, DatesPair> = new Map<string, DatesPair>();

  for (let r: number = START_ROW; r < gfeVals.length; r++) {
    const mat: ExcelScript.CellValue = gfeVals[r][0];
    const season: ExcelScript.CellValue = gfeVals[r][1];
    if (mat === "" || season === "") continue;
    map.set(`${mat}|${season}`, { fa: gfeVals[r][2], disco: gfeVals[r][3] });
  }

  // ───────── Process Main ─────────
  const mainRange: ExcelScript.Range | undefined = wsMain.getUsedRange();
  if (!mainRange) return;
  const mainVals: ExcelScript.CellValue[][] = mainRange.getValues() as ExcelScript.CellValue[][];

  for (let r: number = START_ROW; r < mainVals.length; r++) {
    const mat: ExcelScript.CellValue = mainVals[r][0];
    const season: ExcelScript.CellValue = mainVals[r][1];
    if (mat === "" || season === "") continue;

    const key: string = `${mat}|${season}`;
    const found: DatesPair | undefined = map.get(key);
    if (!found) continue;

    // Copy GFE dates to Main E/F
    mainVals[r][4] = found.fa;
    mainVals[r][5] = found.disco;

    // Flag changes
    if (!datesEqual(mainVals[r][2], found.fa))   mainVals[r][6] = "Y";
    if (!datesEqual(mainVals[r][3], found.disco)) mainVals[r][7] = "Y";
  }

  // ───────── Write back to sheet ─────────
  mainRange.setValues(mainVals);
}

// ===== Helper functions outside main (explicitly typed) =====

/** Return true when both cell values represent the same date (or both blank). */
function datesEqual(a: ExcelScript.CellValue, b: ExcelScript.CellValue): boolean {
  const isBlank = (v: ExcelScript.CellValue): boolean => v === "" || v === null || v === undefined;
  if (isBlank(a) && isBlank(b)) return true;
  return toMillis(a) === toMillis(b);
}

/** Convert CellValue → milliseconds since epoch (or Excel serial as‑is). */
function toMillis(v: ExcelScript.CellValue): number | undefined {
  if (typeof v === "number") return v; // Excel serial already numeric
  if (v instanceof Date)        return v.getTime();
  if (typeof v === "string") {
    const dt: Date = new Date(v as string);
    return isNaN(dt.getTime()) ? undefined : dt.getTime();
  }
  return undefined;
}
