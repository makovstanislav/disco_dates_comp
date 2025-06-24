// Define a strict cell value type
type CellVal = string | number | boolean | Date;

/**
 * Office Script: Clean data from "Main" columns A-D into "Clean" sheet with formatted dates.
 * • Reads rows starting at row 3 (zero-based index 2).
 * • Removes leading apostrophes and trims strings.
 * • Converts numeric strings in Material to numbers.
 * • Parses string dates and Excel serials into JS Date objects.
 * • Formats dates as "MM/DD/YYYY" strings in output.
 * • Uses custom CellVal type to satisfy TypeScript inference.
 */
function main(workbook: ExcelScript.Workbook): void {
  const START_ROW: number = 2;
  const COL_COUNT: number = 4;

  const mainWs: ExcelScript.Worksheet = workbook.getWorksheet("Main")!;
  let cleanWs: ExcelScript.Worksheet | undefined = workbook.getWorksheet("Clean");
  if (!cleanWs) cleanWs = workbook.addWorksheet("Clean");

  const usedRange: ExcelScript.Range = mainWs.getUsedRange()!;
  const totalRows: number = usedRange.getRowCount();
  const rowsToProcess: number = totalRows - START_ROW;
  if (rowsToProcess <= 0) return;

  const sourceRange: ExcelScript.Range = mainWs.getRangeByIndexes(START_ROW, 0, rowsToProcess, COL_COUNT);
  const rawValues: CellVal[][] = sourceRange.getValues() as CellVal[][];

  // Helper to format JS Date to MM/DD/YYYY
  function formatDate(d: Date): string {
    const mm: string = String(d.getMonth() + 1).padStart(2, '0');
    const dd: string = String(d.getDate()).padStart(2, '0');
    const yyyy: number = d.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }

  const cleaned: CellVal[][] = [];
  for (let i: number = 0; i < rawValues.length; i++) {
    const rawRow: CellVal[] = rawValues[i];

    // Material Number
    let matVal: CellVal = rawRow[0];
    if (typeof matVal === "string") {
      const strVal: string = matVal.replace(/^'/, "").trim();
      const numVal: number = Number(strVal);
      matVal = isNaN(numVal) ? strVal : numVal;
    }

    // Season
    let seasonVal: CellVal = rawRow[1];
    if (typeof seasonVal === "string") {
      seasonVal = (seasonVal as string).trim();
    }

    // First Available Date
    let faVal: CellVal = rawRow[2];
    if (typeof faVal === "string") {
      const dt: Date = new Date(faVal as string);
      faVal = isNaN(dt.getTime()) ? "" : dt;
    } else if (typeof faVal === "number") {
      const offsetDays: number = (faVal as number) > 60 ? (faVal as number) - 1 : (faVal as number);
      const epoch: Date = new Date(Date.UTC(1899, 11, 31));
      faVal = new Date(epoch.getTime() + offsetDays * 86400000);
    }

    // Discontinue Date
    let discoVal: CellVal = rawRow[3];
    if (typeof discoVal === "string") {
      const dt2: Date = new Date(discoVal as string);
      discoVal = isNaN(dt2.getTime()) ? "" : dt2;
    } else if (typeof discoVal === "number") {
      const offset2: number = (discoVal as number) > 60 ? (discoVal as number) - 1 : (discoVal as number);
      const epoch2: Date = new Date(Date.UTC(1899, 11, 31));
      discoVal = new Date(epoch2.getTime() + offset2 * 86400000);
    }

    // Format dates for output
    const outFa: string | CellVal = faVal instanceof Date ? formatDate(faVal) : faVal;
    const outDisco: string | CellVal = discoVal instanceof Date ? formatDate(discoVal) : discoVal;

    cleaned.push([matVal, seasonVal, outFa, outDisco]);
  }

  const destRange: ExcelScript.Range = cleanWs.getRangeByIndexes(START_ROW, 0, rowsToProcess, COL_COUNT);
  destRange.setValues(cleaned);
}
