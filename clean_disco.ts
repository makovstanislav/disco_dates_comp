/**
 * Office Script: Clean data from "Main" columns A-D into "Clean" sheet.
 * • Reads rows starting at row 3 (index 2).
 * • Removes leading apostrophes, trims strings.
 * • Converts numeric text in Material to number.
 * • Parses dates (string or Excel serial) to Date objects.
 */
function main(workbook: ExcelScript.Workbook): void {
  // Constants
  const START_ROW: number = 2;  // zero-based (row 3)
  const COL_COUNT: number = 4;   // columns A-D

  // Worksheets
  const mainWs: ExcelScript.Worksheet = workbook.getWorksheet("Main")!;
  let cleanWs: ExcelScript.Worksheet | undefined = workbook.getWorksheet("Clean");
  if (!cleanWs) {
    cleanWs = workbook.addWorksheet("Clean");
  }

  // Determine number of rows to process
  const mainRange: ExcelScript.Range = mainWs.getUsedRange()!;
  const rowCount: number = mainRange.getRowCount();
  const processCount: number = rowCount - START_ROW;
  if (processCount <= 0) return;

  // Read raw values from Main A-D
  const sourceRange: ExcelScript.Range = mainWs.getRangeByIndexes(START_ROW, 0, processCount, COL_COUNT);
  const rawValues: ExcelScript.CellValue[][] = sourceRange.getValues() as ExcelScript.CellValue[][];

  // Clean each row
  const cleaned: ExcelScript.CellValue[][] = rawValues.map((row: ExcelScript.CellValue[]): ExcelScript.CellValue[] => {
    // Material Number
    let mat: ExcelScript.CellValue = row[0];
    if (typeof mat === "string") {
      mat = mat.replace(/^'/, "").trim();
      const n: number = Number(mat);
      if (!isNaN(n)) mat = n;
    }
    // Season
    let season: ExcelScript.CellValue = row[1];
    if (typeof season === "string") {
      season = season.trim();
    }
    // First Available Date
    let fa: ExcelScript.CellValue = row[2];
    if (typeof fa === "string") {
      const d: Date = new Date(fa);
      fa = isNaN(d.getTime()) ? "" : d;
    } else if (typeof fa === "number") {
      // Convert Excel serial to JS Date
      const offset: number = fa > 60 ? fa - 1 : fa;
      const epoch: Date = new Date(Date.UTC(1899, 11, 31));
      fa = new Date(epoch.getTime() + offset * 86400000);
    }
    // Discontinue Date
    let disco: ExcelScript.CellValue = row[3];
    if (typeof disco === "string") {
      const d2: Date = new Date(disco);
      disco = isNaN(d2.getTime()) ? "" : d2;
    } else if (typeof disco === "number") {
      const off2: number = disco > 60 ? disco - 1 : disco;
      const epoch2: Date = new Date(Date.UTC(1899, 11, 31));
      disco = new Date(epoch2.getTime() + off2 * 86400000);
    }
    return [mat, season, fa, disco];
  });

  // Write cleaned values into Clean sheet A-D
  const destRange: ExcelScript.Range = cleanWs.getRangeByIndexes(START_ROW, 0, cleaned.length, COL_COUNT);
  destRange.setValues(cleaned);
}
