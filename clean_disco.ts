/**
 * Office Script: Clean data from "Main" columns A-D into "Clean" sheet.
 * • Reads rows starting at row 3 (zero-based index 2).
 * • Removes leading apostrophes and trims.
 * • Converts numeric strings in Material to numbers.
 * • Parses string or serial dates into Date objects.
 * • All variables and functions use explicit types and for-loops to avoid inference issues.
 */
function main(workbook: ExcelScript.Workbook): void {
  // Constants
  const START_ROW: number = 2; // row 3
  const COL_COUNT: number = 4;  // A-D

  // Worksheets
  const mainWs: ExcelScript.Worksheet = workbook.getWorksheet("Main")!;
  let cleanWs: ExcelScript.Worksheet | undefined = workbook.getWorksheet("Clean");
  if (!cleanWs) {
    cleanWs = workbook.addWorksheet("Clean");
  }

  // Read Main used range
  const usedRange: ExcelScript.Range = mainWs.getUsedRange()!;
  const totalRows: number = usedRange.getRowCount();
  const rowsToProcess: number = totalRows - START_ROW;
  if (rowsToProcess <= 0) {
    return;
  }

  // Fetch raw values A-D from Main
  const sourceRange: ExcelScript.Range = mainWs.getRangeByIndexes(START_ROW, 0, rowsToProcess, COL_COUNT);
  const rawValues: ExcelScript.CellValue[][] = sourceRange.getValues() as ExcelScript.CellValue[][];

  // Prepare cleaned data array
  const cleaned: ExcelScript.CellValue[][] = [];
  for (let rowIndex: number = 0; rowIndex < rawValues.length; rowIndex++) {
    const rawRow: ExcelScript.CellValue[] = rawValues[rowIndex];

    // Material Number
    let matVal: ExcelScript.CellValue = rawRow[0];
    if (typeof matVal === "string") {
      const trimmed: string = matVal.replace(/^'/, "").trim();
      const num: number = Number(trimmed);
      if (!isNaN(num)) {
        matVal = num;
      } else {
        matVal = trimmed;
      }
    }

    // Season
    let seasonVal: ExcelScript.CellValue = rawRow[1];
    if (typeof seasonVal === "string") {
      const trimmedSea: string = seasonVal.trim();
      seasonVal = trimmedSea;
    }

    // First Available Date
    let faVal: ExcelScript.CellValue = rawRow[2];
    if (typeof faVal === "string") {
      const dt: Date = new Date(faVal as string);
      faVal = isNaN(dt.getTime()) ? "" : dt;
    } else if (typeof faVal === "number") {
      const serial: number = faVal;
      const offsetDays: number = serial > 60 ? serial - 1 : serial;
      const baseDate: Date = new Date(Date.UTC(1899, 11, 31));
      faVal = new Date(baseDate.getTime() + offsetDays * 86400000);
    }

    // Discontinue Date
    let discoVal: ExcelScript.CellValue = rawRow[3];
    if (typeof discoVal === "string") {
      const dt2: Date = new Date(discoVal as string);
      discoVal = isNaN(dt2.getTime()) ? "" : dt2;
    } else if (typeof discoVal === "number") {
      const serial2: number = discoVal;
      const offset2: number = serial2 > 60 ? serial2 - 1 : serial2;
      const baseDate2: Date = new Date(Date.UTC(1899, 11, 31));
      discoVal = new Date(baseDate2.getTime() + offset2 * 86400000);
    }

    // Append cleaned row
    const cleanRow: ExcelScript.CellValue[] = [matVal, seasonVal, faVal, discoVal];
    cleaned.push(cleanRow);
  }

  // Write cleaned data into Clean sheet A-D starting at row 3
  const destRange: ExcelScript.Range = cleanWs.getRangeByIndexes(START_ROW, 0, cleaned.length, COL_COUNT);
  destRange.setValues(cleaned);
}
