/**
 * =========================================================
 * BIM QA/QC HIERARCHICAL SORT TOOL — v1.5 FINAL
 * Multi-level (hierarchical) sorting like SQL ORDER BY
 * Багаторівневе сортування як SQL ORDER BY
 *
 * Order:
 * Name → Size → File → Page
 * Порядок: Ім’я → Габарити → Файл → Сторінка
 *
 * Each next column refines previous grouping
 * Кожен наступний стовпець уточнює попередній
 *
 * Example / Приклад:
 * B → C → D → E → L → M
 *
 * ⚠️ IMPORTANT ADDITION:
 * ROOT-based grouping is applied BEFORE full comparison
 * Групування по "кореню" імені виконується ПЕРЕД основним сортуванням
 * =========================================================
 */
function sortHierarchical(columnLetters = [
  "B", // Name (Ім’я)
  "C", // Height (Висота)
  "D", // Width (Ширина)
  "E", // Depth (Глибина)
  // "L", // File Path (Шлях до файлу)
  // "M"  // Page (Сторінка)
]) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // active sheet / активний лист
  const startRow = 2; // skip header / пропустити заголовок

  // AUTO DETECT (may hang on bad data)
  // Автовизначення (може зависати на великих або "битих" таблицях)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // SAFE MODE (manual override)
  // Безпечний режим (захардкодити значення якщо скрипт зависає)
  // const lastRow = 1000;
  // const lastCol = 20;

  if (lastRow < startRow) return; // no data / немає даних

  const range = sheet.getRange(startRow, 1, lastRow - 1, lastCol); // working range / робочий діапазон

  // GET VALUES + FORMATTING
  // Отримуємо всі дані + форматування (щоб не втратити стиль після сортування)
  const values = range.getValues(); // значення
  const backgrounds = range.getBackgrounds(); // фон
  const fontColors = range.getFontColors(); // колір тексту
  const fontWeights = range.getFontWeights(); // жирність
  const fontStyles = range.getFontStyles(); // курсив
  const fontSizes = range.getFontSizes(); // розмір

  // COLUMN LETTER → INDEX
  // Перетворення букв колонок у індекси масиву
  const colLetterToIndex = (letter) =>
    letter.toUpperCase().charCodeAt(0) - 65; // A=0, B=1...

  const sortCols = columnLetters.map(colLetterToIndex); // масив індексів колонок

  // CLEAN VALUE
  // Нормалізація значення (прибирає null/undefined/пробіли)
  const clean = (v) => {
    if (v === null || v === undefined) return "";
    return v.toString().trim();
  };

  // TRY NUMBER
  // Пробуємо перетворити в число (для габаритів)
  const tryParseNumber = (v) => {
    const n = Number(v);
    return isNaN(n) ? null : n;
  };

  // ROOT EXTRACTION
  // Витягує "корінь" імені (для групування)
  // Приклад:
  // "CP E2-4-h1-3-d" → "CP"
  // "PB-01-A" → "PB"
  const getRoot = (name) => {
    const v = clean(name);
    const parts = v.split(/[\s\-_]+/); // розділення по пробілу, дефісу, _
    return parts[0] || "";
  };

  // MERGE DATA + STYLE
  // Об’єднуємо значення і формат в один об’єкт (щоб рухались разом)
  const rows = values.map((row, i) => ({
    values: row,
    bg: backgrounds[i],
    fc: fontColors[i],
    fw: fontWeights[i],
    fs: fontStyles[i],
    fz: fontSizes[i]
  }));

  rows.sort((a, b) => { // SORT CORE / ядро сортування

    // =========================
    // 1️⃣ ROOT GROUPING (NEW)
    // =========================
    // Спочатку групуємо по кореню імені (B)
    const nameCol = sortCols[0];

    const rootA = getRoot(a.values[nameCol]);
    const rootB = getRoot(b.values[nameCol]);

    if (rootA < rootB) return -1;
    if (rootA > rootB) return 1;

    // =========================
    // 2️⃣ HIERARCHICAL SORT
    // =========================
    // Далі стандартна ієрархія (B → C → D → E → ...)
    for (let col of sortCols) {

      let A = clean(a.values[col]);
      let B = clean(b.values[col]);

      // empty handling / обробка пустих
      if (!A && !B) continue;
      if (!A) return 1;
      if (!B) return -1;

      // numeric compare (для габаритів)
      const numA = tryParseNumber(A);
      const numB = tryParseNumber(B);

      if (numA !== null && numB !== null) {
        if (numA < numB) return -1;
        if (numA > numB) return 1;
        continue; // якщо рівні → наступний рівень
      }

      // string compare (для тексту)
      if (A < B) return -1;
      if (A > B) return 1;
    }

    return 0; // повністю рівні
  });

  // WRITE BACK VALUES + FORMATTING
  // Записуємо назад і дані і формат (1:1)
  range.setValues(rows.map(r => r.values));
  range.setBackgrounds(rows.map(r => r.bg));
  range.setFontColors(rows.map(r => r.fc));
  range.setFontWeights(rows.map(r => r.fw));
  range.setFontStyles(rows.map(r => r.fs));
  range.setFontSizes(rows.map(r => r.fz));

  // LOG (for debugging / QAQC tracking)
  Logger.log(`TABLE IS NOW SORTED IN ORDER: ${columnLetters.join(" → ")}`);
}
