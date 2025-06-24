/**
 * Office Script: copy dates from «GFE» → «Main»
 * and mark changes with «Y» (cols G‑H).
 *
 * – Works from row 3 (index 2).
 * – Uses only explicit types; no `any`.
 */
function main(workbook: ExcelScript.Workbook) {
  // ---------- Constants ----------
  const START_ROW = 2; // 0‑based → Excel row 3
  const mainWs = workbook.getWorksheet("Main");
  const gfeWs = workbook.getWorksheet("GFE");
  if (!mainWs || !gfeWs) return;

  // ---------- Build lookup from GFE ----------
  const gfeRange = gfeWs.getUsedRange();
  if (!gfeRange) return;
  const gfeVals = gfeRange.getValues() as ExcelScript.CellValue[][];

  interface DatesPair { fa: ExcelScript.CellValue; disco: ExcelScript.CellValue }
  const lookup = new Map<string, DatesPair>();

  for (let r = START_ROW; r < gfeVals.length; r++) {
    const mat = gfeVals[r][0] as string | number;
    const season = gfeVals[r][1] as string | number;
    if (mat === "" || season === "") continue;
    lookup.set(`${mat}|${season}`, { fa: gfeVals[r][2], disco: gfeVals[r][3] });
  }

  // ---------- Read Main ----------
  const mainRange = mainWs.getUsedRange();
  if (!mainRange) return;
  const mainVals = mainRange.getValues() as ExcelScript.CellValue[][];

  for (let r = START_ROW; r < mainVals.length; r++) {
    const mat = mainVals[r][0] as string | number;
    const season = mainVals[r][1] as string | number;
    if (mat === "" || season === "") continue;

    const key = `${mat}|${season}`;
    const data: DatesPair | undefined = lookup.get(key);
    if (!data) continue; // pair not in GFE

    // Copy dates to E/F
    mainVals[r][4] = data.fa;
    mainVals[r][5] = data.disco;

    // Compare and flag
    if (!sameDate(mainVals[r][2], data.fa)) mainVals[r][6] = "Y";
    if (!sameDate(mainVals[r][3], data.disco)) mainVals[r][7] = "Y";
  }

  // ---------- Write back ----------
  mainRange.setValues(mainVals);

  // ---------- Helpers ----------
  function sameDate(a: ExcelScript.CellValue, b: ExcelScript.CellValue): boolean {
    if ((a === "" || a === null || a === undefined) && (b === "" || b === null || b === undefined)) return true;
    return toMillis(a) === toMillis(b);
  }

  function toMillis(v: ExcelScript.CellValue): number | undefined {
    if (typeof v === "number") return v; // Excel serial
    if (v instanceof Date) return v.getTime();
    if (typeof v === "string") {
      const d = new Date(v);
      return isNaN(d.getTime()) ? undefined : d.getTime();
    }
    return undefined;
  }
}
