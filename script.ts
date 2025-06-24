/**
 * Compare dates between "Main" and "GFE" sheets.
 *
 *   • Copies dates from GFE → Main (cols E–F, starting row 3).
 *   • Writes "Y" in G/H when dates differ.
 *   • Leaves cells blank if no change or pair not found.
 */
function main(workbook: ExcelScript.Workbook) {
  const MAIN = "Main";
  const GFE = "GFE";
  const START_ROW = 2; // 0‑based → row 3

  const mainWs = workbook.getWorksheet(MAIN);
  const gfeWs = workbook.getWorksheet(GFE);

  // ---------- Build lookup from GFE ----------
  const gfeRange = gfeWs.getUsedRange();
  if (!gfeRange) return; // empty sheet

  const gfeValues = gfeRange.getValues() as ExcelScript.CellValue[][];
  const gfeMap = new Map<string, [ExcelScript.CellValue, ExcelScript.CellValue]>();

  for (let r = START_ROW; r < gfeValues.length; r++) {
    const mat = gfeValues[r][0];
    const season = gfeValues[r][1];
    if (!mat || !season) continue;
    gfeMap.set(`${mat}|${season}`, [gfeValues[r][2], gfeValues[r][3]]);
  }

  // ---------- Process rows in Main ----------
  const mainRange = mainWs.getUsedRange();
  if (!mainRange) return;

  const mainValues = mainRange.getValues() as ExcelScript.CellValue[][];
  const rowCount = mainValues.length;

  for (let r = START_ROW; r < rowCount; r++) {
    const mat = mainValues[r][0];
    const season = mainValues[r][1];
    if (!mat || !season) continue;

    const pair = gfeMap.get(`${mat}|${season}`);
    if (!pair) continue; // not found

    const [gfeFA, gfeDisco] = pair;

    // Copy dates to E (4) & F (5)
    mainValues[r][4] = gfeFA;
    mainValues[r][5] = gfeDisco;

    // Compare and flag
    if (!datesEqual(mainValues[r][2], gfeFA)) mainValues[r][6] = "Y";
    if (!datesEqual(mainValues[r][3], gfeDisco)) mainValues[r][7] = "Y";
  }

  // ---------- Write back ----------
  mainRange.setValues(mainValues);

  // ---------- Helpers ----------
  function datesEqual(a: ExcelScript.CellValue, b: ExcelScript.CellValue): boolean {
    if ((a === "" || a === undefined || a === null) && (b === "" || b === undefined || b === null)) {
      return true;
    }
    return toMillis(a) === toMillis(b);
  }

  function toMillis(v: ExcelScript.CellValue): number | undefined {
    if (v instanceof Date) return v.getTime();
    if (typeof v === "number") return v; // Excel serial number
    if (typeof v === "string") {
      const d = new Date(v);
      if (!isNaN(d.getTime())) return d.getTime();
    }
    return undefined;
  }
}
