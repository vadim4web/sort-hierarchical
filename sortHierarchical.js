/**
 * =========================================================
 * BIM QA/QC HIERARCHICAL SORT TOOL — v1.5 FINAL
 * Multi-level (hierarchical) sorting like SQL ORDER BY || Багаторівневе сортування як SQL ORDER BY
 * Order: Name → Size → File → Page || Порядок: Ім’я → Габарити → Файл → Сторінка
 * Each next column refines previous grouping || Кожен наступний стовпець уточнює попередній
 * Example / Приклад:
 * B → C → D → E → L → M
 * =========================================================
 */
function sortHierarchical(columnLetters = [
  "B", // Name || Ім’я
  "C", // Height || Висота
  "D", // Width || Ширина
  "E", // Depth || Глибина
  // "L", // File Path || Шлях до файлу
  // "M"  // Page || Сторінка
]) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // active sheet || активний лист
  const startRow = 2; // skip header || пропустити заголовок

  // AUTO DETECT (may hang on bad data) || може зависати на поганих даних
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  // SAFE MODE (uncomment if script freezes) || безпечний режим (якщо зависає)
  // const lastRow = 1000;
  // const lastCol = 20;

  if (lastRow < startRow) return; // no data || немає даних

  const range = sheet.getRange(startRow, 1, lastRow - 1, lastCol); // working range || робочий діапазон

  // GET VALUES + FORMATTING || отримати дані + формат
  const values = range.getValues(); // cell values || значення
  const backgrounds = range.getBackgrounds(); // bg color || фон
  const fontColors = range.getFontColors(); // text color || колір тексту
  const fontWeights = range.getFontWeights(); // bold || жирність
  const fontStyles = range.getFontStyles(); // italic || курсив
  const fontSizes = range.getFontSizes(); // font size || розмір

  const colLetterToIndex = (letter) => // COLUMN LETTER → INDEX || буква → індекс
    letter.toUpperCase().charCodeAt(0) - 65; // A=0, B=1... || A=0, B=1...

  const sortCols = columnLetters.map(colLetterToIndex); // map letters || мапінг

  const clean = (v) => { // CLEAN VALUE || очистка значення
    if (v === null || v === undefined) return ""; // empty → "" || пусто
    return v.toString().trim(); // normalize || нормалізація
  };

  const tryParseNumber = (v) => { // TRY NUMBER || спроба як число
    const n = Number(v);
    return isNaN(n) ? null : n; // number or null || число або null
  };

  const rows = values.map((row, i) => ({ // MERGE DATA + STYLE || об’єднати дані і стиль
    values: row,        // data || дані
    bg: backgrounds[i], // background || фон
    fc: fontColors[i],  // font color || колір
    fw: fontWeights[i], // bold || жирність
    fs: fontStyles[i],  // italic || курсив
    fz: fontSizes[i]    // size || розмір
  }));

  rows.sort((a, b) => { // SORT CORE || ядро сортування

    for (let col of sortCols) { // loop hierarchy || ієрархія

      let A = clean(a.values[col]);
      let B = clean(b.values[col]);

      // empty handling || обробка пустих
      if (!A && !B) continue;
      if (!A) return 1;
      if (!B) return -1;

      // numeric compare || числове порівняння
      const numA = tryParseNumber(A);
      const numB = tryParseNumber(B);

      if (numA !== null && numB !== null) {
        if (numA < numB) return -1;
        if (numA > numB) return 1;
        continue; // equal → next level || рівні → далі
      }

      // string compare || текстове порівняння
      if (A < B) return -1;
      if (A > B) return 1;
    }

    return 0; // fully equal || повністю рівні
  });

  // WRITE BACK VALUES + FORMATTING || записати назад дані + формат
  range.setValues(rows.map(r => r.values)); // values || значення
  range.setBackgrounds(rows.map(r => r.bg)); // bg || фон
  range.setFontColors(rows.map(r => r.fc)); // color || колір
  range.setFontWeights(rows.map(r => r.fw)); // bold || жирність
  range.setFontStyles(rows.map(r => r.fs)); // italic || курсив
  range.setFontSizes(rows.map(r => r.fz)); // size || розмір

  Logger.log(`TABLE IS NOW SORTED IN ORDER: ${columnLetters.join(" → ")}`);
}
