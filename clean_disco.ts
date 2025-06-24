function main(workbook) {
  const START_ROW = 2;
  const COL_COUNT = 4;

  const mainWs = workbook.getWorksheet("Main");
  let cleanWs = workbook.getWorksheet("Clean");
  if (!cleanWs) cleanWs = workbook.addWorksheet("Clean");

  const used = mainWs.getUsedRange();
  const totalRows = used.getRowCount();
  const toProc = totalRows - START_ROW;
  if (toProc <= 0) return;

  const raw = mainWs
    .getRangeByIndexes(START_ROW, 0, toProc, COL_COUNT)
    .getValues();

  const cleaned = raw.map(row => {
    let mat = row[0];
    if (typeof mat === "string") {
      mat = mat.replace(/^'/, "").trim();
      const n = Number(mat);
      mat = isNaN(n) ? mat : n;
    }

    let sea = row[1];
    if (typeof sea === "string") sea = sea.trim();

    let fa = row[2];
    if (typeof fa === "string") {
      const d = new Date(fa);
      fa = isNaN(d.getTime()) ? "" : d;
    } else if (typeof fa === "number") {
      const off = fa > 60 ? fa - 1 : fa;
      const epoch = new Date(Date.UTC(1899,11,31));
      fa = new Date(epoch.getTime() + off*86400000);
    }

    let disco = row[3];
    if (typeof disco === "string") {
      const d2 = new Date(disco);
      disco = isNaN(d2.getTime()) ? "" : d2;
    } else if (typeof disco === "number") {
      const off2 = disco > 60 ? disco - 1 : disco;
      const epoch2 = new Date(Date.UTC(1899,11,31));
      disco = new Date(epoch2.getTime() + off2*86400000);
    }

    return [mat, sea, fa, disco];
  });

  cleanWs
    .getRangeByIndexes(START_ROW, 0, toProc, COL_COUNT)
    .setValues(cleaned);
}
