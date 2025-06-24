/**
 * Office Script (JavaScript):
 * Копирует данные из Main (A–D), чистит и записывает на лист Clean.
 * • Убирает ведущие апострофы и пробелы.
 * • Преобразует текстовые номера в числа.
 * • Парсит текстовые даты и Excel-сериалы в объекты Date.
 */
function main(workbook) {
  const START_ROW = 2; // начинаем с 3-й строки (0-based index)
  const COL_COUNT = 4; // столбцы A–D

  // получаем Main и Clean (создаём, если нет)
  const mainWs = workbook.getWorksheet("Main");
  let cleanWs = workbook.getWorksheet("Clean");
  if (!cleanWs) {
    cleanWs = workbook.addWorksheet("Clean");
  }

  // сколько строк всего в Main?
  const used = mainWs.getUsedRange();
  const totalRows = used.getRowCount();
  const toProcess = totalRows - START_ROW;
  if (toProcess <= 0) return;

  // читаем A–D из Main
  const raw = mainWs
    .getRangeByIndexes(START_ROW, 0, toProcess, COL_COUNT)
    .getValues();

  // очистка
  const cleaned = [];
  for (const row of raw) {
    // Material
    let mat = row[0];
    if (typeof mat === "string") {
      mat = mat.replace(/^'/, "").trim();
      const nn = Number(mat);
      mat = isNaN(nn) ? mat : nn;
    }

    // Season
    let sea = row[1];
    if (typeof sea === "string") {
      sea = sea.trim();
    }

    // First Available
    let fa = row[2];
    if (typeof fa === "string") {
      const d = new Date(fa);
      fa = isNaN(d.getTime()) ? "" : d;
    } else if (typeof fa === "number") {
      const off = fa > 60 ? fa - 1 : fa;
      const epoch = new Date(Date.UTC(1899, 11, 31));
      fa = new Date(epoch.getTime() + off * 86400000);
    }

    // Discontinue
    let disco = row[3];
    if (typeof disco === "string") {
      const d2 = new Date(disco);
      disco = isNaN(d2.getTime()) ? "" : d2;
    } else if (typeof disco === "number") {
      const off2 = disco > 60 ? disco - 1 : disco;
      const epoch2 = new Date(Date.UTC(1899, 11, 31));
      disco = new Date(epoch2.getTime() + off2 * 86400000);
    }

    cleaned.push([mat, sea, fa, disco]);
  }

  // пишем на лист Clean A–D, начиная с 3-й строки
  cleanWs
    .getRangeByIndexes(START_ROW, 0, cleaned.length, COL_COUNT)
    .setValues(cleaned);
}
