/**
 * Copy dates from «GFE» → «Main» and flag changes.
 *
 *   • Works from row 3 (index 2).
 *   • Writes «Y» only when FA / Disco dates differ.
 *   • All variables have explicit types (no implicit `any`).
 */
function main(workbook: ExcelScript.Workbook): void {
  // ── Constants & worksheets ──
  const START_ROW: number = 2; // row 3 (0‑based)
  const wsMain: ExcelScript.Worksheet | undefined = workbook.getWorksheet("Main");
  const wsGfe:  ExcelScript.Worksheet | undefined = workbook.getWorksheet("GFE");
  if (!wsMain || !wsGfe) return;

  // ── Build lookup from GFE ──
  const gfeRange: ExcelScript.Range | undefined = wsGfe.getUsedRange();
  if (!gfeRange) return;
  const gfeVals: ExcelScript.CellValue[][] = gfeRange.getValues() as ExcelScript.CellValue[][];

  interface DatesPair { fa: ExcelScript.CellValue; disco: ExcelScript.CellValue }
  const lookup: { [key: string]: DatesPair } = {};

  for (let r: number = START_ROW; r < gfeVals.length; r++) {
    const matCell: ExcelScript.CellValue = gfeVals[r][0];
    const seasonCell: ExcelScript.CellValue = gfeVals[r][1];
    if (matCell === "" || seasonCell === "") continue;
    const key: string = String(matCell) + "|" + String(seasonCell);
    lookup[key] = { fa: gfeVals[r][2], disco: gfeVals[r][3] };
  }

  // ── Process Main rows ──
  const mainRange: ExcelScript.Range | undefined = wsMain.getUsedRange();
  if (!mainRange) return;
  const mainVals: ExcelScript.CellValue[][] = mainRange.getValues() as ExcelScript.CellValue[][];

  for (let r: number = START_ROW; r < mainVals.length; r++) {
    const matCell: ExcelScript.CellValue = mainVals[r][0];
    const seasonCell: ExcelScript.CellValue = mainVals[r][1];
    if (matCell === "" || seasonCell === "") continue;

    const key: string = String(matCell) + "|" + String(seasonCell);
    const ref: DatesPair | undefined = lookup[key];
    if (!ref) continue; // pair absent in GFE

    // Copy dates into E / F
    mainVals[r][4] = ref.fa;
    mainVals[r][5] = ref.disco;

    // Flag changes
    if (!datesEqual(mainVals[r][2], ref.fa))   mainVals[r][6] = "Y";
    if (!datesEqual(mainVals[r][3], ref.disco)) mainVals[r][7] = "Y";
  }

  // ── Write back ──
  mainRange.setValues(mainVals);

  // ── Helper functions ──
  function datesEqual(a: ExcelScript.CellValue, b: ExcelScript.CellValue): boolean {
    return normalise(a) === normalise(b);
  }

  /** Convert value → comparable form (string / number). */
  function normalise(v: ExcelScript.CellValue): string | number {
    if (v === "" || v === null || v === undefined) return ""; // treat blanks as equal
    if (typeof v === "number") return v; // Excel serial
    if (v instanceof Date)      return v.getTime();
    return String(v).trim();
  }
}
