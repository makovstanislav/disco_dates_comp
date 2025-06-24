/**
 * Office Script: Clean data from Main A–D into Clean sheet.
 * - Starts at row 3 (zero-based index 2)
 * - Removes leading apostrophes, trims strings
 * - Parses numeric text → numbers
 * - Parses Excel serials & string dates → JS Date
 */
function main(workbook) {
  var START_ROW = 2;
  var COL_COUNT = 4;

  var mainWs = workbook.getWorksheet("Main");
  var cleanWs = workbook.getWorksheet("Clean");
  if (!cleanWs) {
    cleanWs = workbook.addWorksheet("Clean");
  }

  var used = mainWs.getUsedRange();
  var totalRows = used.getRowCount();
  var toProcess = totalRows - START_ROW;
  if (toProcess <= 0) return;

  // Считаем A–D
  var raw = mainWs
    .getRangeByIndexes(START_ROW, 0, toProcess, COL_COUNT)
    .getValues();

  var cleaned = [];
  for (var r = 0; r < raw.length; r++) {
    var row = raw[r];

    // Материал
    var mat = row[0];
    if (typeof mat === "string") {
      var t = mat.replace(/^'/, "").trim();
      var n = Number(t);
      mat = isNaN(n) ? t : n;
    }

    // Сезон
    var sea = row[1];
    if (typeof sea === "string") {
      sea = sea.trim();
    }

    // First Available
    var fa = row[2];
    if (typeof fa === "string") {
      var d = new Date(fa);
      fa = isNaN(d.getTime()) ? "" : d;
    } else if (typeof fa === "number") {
      var off = fa > 60 ? fa - 1 : fa;
      var ep = new Date(Date.UTC(1899, 11, 31));
      fa = new Date(ep.getTime() + off * 86400000);
    }

    // Discontinue
    var disco = row[3];
    if (typeof disco === "string") {
      var d2 = new Date(disco);
      disco = isNaN(d2.getTime()) ? "" : d2;
    } else if (typeof disco === "number") {
      var off2 = disco > 60 ? disco - 1 : disco;
      var ep2 = new Date(Date.UTC(1899, 11, 31));
      disco = new Date(ep2.getTime() + off2 * 86400000);
    }

    cleaned.push([mat, sea, fa, disco]);
  }

  cleanWs
    .getRangeByIndexes(START_ROW, 0, cleaned.length, COL_COUNT)
    .setValues(cleaned);
}
