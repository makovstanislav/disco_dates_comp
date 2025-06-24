/**
 * Скрипт для вкладок «Main» и «GFE».
 *
 * Заполняет колонки E–F датами из «GFE»
 * и проставляет «Y» в G и H, если даты изменились.
 * Работает с 3‑й строки (индекс 2).
 */
function main(workbook: ExcelScript.Workbook) {
  const MAIN = "Main";
  const GFE = "GFE";
  const START_ROW = 2; // индекс 0‑based: строка №3 в Excel

  const mainWs = workbook.getWorksheet(MAIN);
  const gfeWs = workbook.getWorksheet(GFE);

  // --- Читаем GFE и строим карту «material|season» → [FA, Disco] ---
  const gfeData = gfeWs.getUsedRange()?.getValues() as (string | number | boolean | Date)[][];
  const gfeMap = new Map<string, [any, any]>();
  if (gfeData) {
    for (let r = START_ROW; r < gfeData.length; r++) {
      const mat = gfeData[r][0];
      const season = gfeData[r][1];
      if (mat === "" || season === "") continue;
      gfeMap.set(`${mat}|${season}`, [gfeData[r][2], gfeData[r][3]]);
    }
  }

  // --- Читаем Main целиком A:H ---
  const lastRow = mainWs.getRange("A:A").getUsedRange().getLastRow();
  const mainRange = mainWs.getRangeByIndexes(0, 0, lastRow + 1, 8); // 8 → кол-во колонок A‑H
  const mainData = mainRange.getValues();

  // --- Обработка строк Main ---
  for (let r = START_ROW; r < mainData.length; r++) {
    const mat = mainData[r][0];
    const season = mainData[r][1];
    if (mat === "" || season === "") continue;

    const key = `${mat}|${season}`;
    const gfeDates = gfeMap.get(key);
    if (!gfeDates) continue; // нет пары в GFE

    const [gfeFA, gfeDisco] = gfeDates;

    // Заполняем E и F
    mainData[r][4] = gfeFA;
    mainData[r][5] = gfeDisco;

    // Сравнения → «Y» при отличии
    if (!datesEqual(mainData[r][2], gfeFA)) mainData[r][6] = "Y";
    if (!datesEqual(mainData[r][3], gfeDisco)) mainData[r][7] = "Y";
  }

  // --- Записываем обратно ---
  mainRange.setValues(mainData);
}

// === Вспомогательные функции ===
function datesEqual(a: any, b: any): boolean {
  // Пустое с пустым считаем равными
  if ((a === "" || a === undefined || a === null) && (b === "" || b === undefined || b === null)) return true;
  return toSerial(a) === toSerial(b);
}

function toSerial(v: any): number {
  if (v instanceof Date) return v.getTime();
  if (typeof v === "number") return v; // Excel serial
  if (typeof v === "string") {
    const d = new Date(v);
    return isNaN(d.getTime()) ? NaN : d.getTime();
  }
  return NaN;
}
