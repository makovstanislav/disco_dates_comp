/**
 * Office Script: Clean data from "Main" columns A-D into "Clean" sheet.
 * • Reads rows starting at row 3 (zero-based index 2).
 * • Removes leading apostrophes and trims.
 * • Converts numeric strings in Material to numbers.
 * • Parses string or serial dates into Date objects.
 * • All variables explicitly typed to avoid inference issues.
 */
function main(workbook: ExcelScript.Workbook): void {
  // Constants
  const START_ROW: number = 2; // row 3
  const COL_COUNT: number = 4;  // columns A-D

  // Worksheets
  const mainWs: ExcelScript.Worksheet = workbook.getWorksheet("Main")!;
  let cleanWs: ExcelScript.Worksheet | undefined = workbook.getWorksheet("Clean");
  if (cleanWs === undefined) {
    cleanWs = workbook.addWorksheet("Clean");
  }

  // Determine rows to process
  const usedRange: ExcelScript.Range = mainWs.getUsedRange()!;
  const totalRows: number = usedRange.getRowCount();
  const rowsToProcess: number = totalRows - START_ROW;
  if (rowsToProcess <= 0) {
    return;
  }

  // Read source values from Main A-D
  const sourceRange: ExcelScript.Range = mainWs.getRangeByIndexes(START_ROW, 0, rowsToProcess, COL_COUNT);
  const rawValues: ExcelScript.CellValue[][] = sourceRange.getValues() as ExcelScript.CellValue[][];

  // Prepare cleaned data array with explicit type
  const cleaned: ExcelScript.CellValue[][] = new Array<ExcelScript.CellValue[]>(rowsToProcess);

  // Loop through each raw row
  for (let rowIndex: number = 0; rowIndex < rowsToProcess; rowIndex++) {
    const rawRow: ExcelScript.CellValue[] = rawValues[rowIndex] as ExcelScript.CellValue[];

    // Clean Material Number
    let matVal: ExcelScript.CellValue = rawRow[0];
    if (typeof matVal === "string") {
      const trimmedMat: string = (matVal as string).replace(/^'/, "").trim();
      const parsedNum: number = Number(trimmedMat);
      if (!isNaN(parsedNum)) {
        matVal = parsedNum;
      } else {
        matVal = trimmedMat;
      }
    }

    // Clean Season
    let seasonVal: ExcelScript.CellValue = rawRow[1];
    if (typeof seasonVal === "string") {
      const trimmedSea: string = (seasonVal as string).trim();
      seasonVal = trimmedSea;
    }

    // Clean First Available Date
    let faVal: ExcelScript.CellValue = rawRow[2];
    if (typeof faVal === "string") {
      const dateObj: Date = new Date(faVal as string);
      faVal = isNaN(dateObj.getTime()) ? "" : dateObj;
    } else if (typeof faVal === "number") {
      const serialNum: number = faVal as number;
      const offsetDays: number = serialNum > 60 ? serialNum - 1 : serialNum;
      const baseDate: Date = new Date(Date.UTC(1899, 11, 31));
      faVal = new Date(baseDate.getTime() + offsetDays * 86400000);
    }

    // Clean Discontinue Date
    let discoVal: ExcelScript.CellValue = rawRow[3];
    if (typeof discoVal === "string") {
      const dateObj2: Date = new Date(discoVal as string);
      discoVal = isNaN(dateObj2.getTime()) ? "" : dateObj2;
    } else if (typeof discoVal === "number") {
      const serialNum2: number = discoVal as number;
      const offsetDays2: number = serialNum2 > 60 ? serialNum2 - 1 : serialNum2;
      const baseDate2: Date = new Date(Date.UTC(1899, 11, 31));
      discoVal = new Date(baseDate2.getTime() + offsetDays2 * 86400000);
    }

    // Assign cleaned row
    cleaned[rowIndex] = [matVal, seasonVal, faVal, discoVal];
  }

  // Write cleaned data into Clean sheet starting at row 3
  const destRange: ExcelScript.Range = cleanWs.getRangeByIndexes(START_ROW, 0, rowsToProcess, COL_COUNT);
  destRange.setValues(cleaned);
}
