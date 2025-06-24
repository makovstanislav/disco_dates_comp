// Define strict cell type for inference
type CellVal = string | number | boolean | Date;

/**
 * Office Script: Export GFE data into GFE_Clean sheet.
 * • Copies columns A–D (Material, Season, First Available, Discontinue) from GFE.
 * • Includes header row.
 * • Creates GFE_Clean sheet if it does not exist.
 */
function main(workbook: ExcelScript.Workbook): void {
  const SOURCE_SHEET: string = "GFE";
  const TARGET_SHEET: string = "GFE_Clean";
  const COL_COUNT: number = 4; // A–D

  // Source worksheet
  const wsSource: ExcelScript.Worksheet = workbook.getWorksheet(SOURCE_SHEET)!;

  // Target worksheet (create if missing)
  let wsTarget: ExcelScript.Worksheet | undefined = workbook.getWorksheet(TARGET_SHEET);
  if (!wsTarget) {
    wsTarget = workbook.addWorksheet(TARGET_SHEET);
  }

  // Read source data A–D
  const usedRange: ExcelScript.Range = wsSource.getUsedRange()!;
  const rowCount: number = usedRange.getRowCount();
  const sourceRange: ExcelScript.Range = wsSource.getRangeByIndexes(0, 0, rowCount, COL_COUNT);
  const data: CellVal[][] = sourceRange.getValues() as CellVal[][];

  // Clear target sheet existing content
  const targetUsed: ExcelScript.Range | undefined = wsTarget.getUsedRange();
  if (targetUsed) {
    wsTarget.getRangeByIndexes(0, 0, targetUsed.getRowCount(), COL_COUNT).clear(ExcelScript.ClearApplyTo.contents);
  }

  // Write data to target sheet A–D
  const targetRange: ExcelScript.Range = wsTarget.getRangeByIndexes(0, 0, rowCount, COL_COUNT);
  targetRange.setValues(data);

  // Optional: adjust column widths
  for (let c: number = 0; c < COL_COUNT; c++) {
    wsTarget.getRangeByIndexes(0, c, rowCount, 1).getFormat().autofitColumns();
  }
}
