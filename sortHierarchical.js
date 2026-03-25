/**
 * Hierarchical sorting with style preservation / Ієрархічне сортування зі збереженням стилів
 * @param {string[]} paramLetters - Columns for sorting / Стовпці для сортування
 * @param {boolean} ascending - Sort direction / Напрямок сортування
 */
function sortHierarchical(
  paramLetters = ["B", "C", "D", "E"],
  ascending = false,
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // get sheet / отримати лист
  const startRow = 2; // skip header / пропустити заголовок

  // AUTO DETECT (may hang on bad data) / Автовизначення (може зависати на "битих" даних)
  const lastRow = sheet.getLastRow(); // detect rows / визначити рядки
  const lastCol = sheet.getLastColumn(); // detect cols / визначити колонки

  // SAFE MODE (manual override) / Безпечний режим (ручне керування при зависанні)
  // const lastRow = 1000; / hardcode rows / захардкодити рядки
  // const lastCol = 20; / hardcode cols / захардкодити колонки

  if (lastRow < startRow) return; // empty check / перевірка на порожнечу

  const range = sheet.getRange(startRow, 1, lastRow - 1, lastCol); // target range / цільовий діапазон

  // GET VALUES + FORMATTING / Отримати дані + форматування
  const values = range.getValues(); // cell values / значення
  const backgrounds = range.getBackgrounds(); // bg color / колір фону
  const fontColors = range.getFontColors(); // text color / колір тексту
  const fontWeights = range.getFontWeights(); // bold / жирність
  const fontStyles = range.getFontStyles(); // italic / курсив
  const fontSizes = range.getFontSizes(); // size / розмір шрифту

  // HELPERS / ДОПОМІЖНІ ФУНКЦІЇ
  const colIdx = (l) => l.toUpperCase().charCodeAt(0) - 65; // letter to index / літеру в індекс
  const clean = (v) => (v === null || v === undefined) ? "" : v.toString().trim(); // sanitize / очистка
  const tryNum = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n; }; // num check / перевірка числа
  const getRoot = (name) => clean(name).split(/[\s\-_]+/)[0] || ""; // get prefix / отримати префікс

  const rootIdx = colIdx(paramLetters[0]); // main level / головний рівень
  const sortIdx = paramLetters.slice(1).map(colIdx); // sub levels / підрівні

  // DATA MAPPING / МАПІНГ ДАНИХ
  const rows = values.map((row, i) => ({
    data: row, // values / дані
    style: { // styles / стилі
      bg: backgrounds[i], fc: fontColors[i], fw: fontWeights[i], fs: fontStyles[i], fz: fontSizes[i]
    }
  }));

  // COMPARE LOGIC / ЛОГІКА ПОРІВНЯННЯ
  rows.sort((a, b) => {
    const rootA = getRoot(a.data[rootIdx]); // root a / корінь а
    const rootB = getRoot(b.data[rootIdx]); // root b / корінь б

    if (rootA !== rootB) return rootA.localeCompare(rootB, undefined, { numeric: true }); // root sort / сортування коренів

    let paramResult = 0; // result flag / прапорець результату
    let paramDiff = false; // diff flag / чи є різниця
    let docA = null, docB = null; // doc vars / змінні документа
    let pageA = null, pageB = null; // page vars / змінні сторінки

    for (let i = 0; i < sortIdx.length; i++) {
      const idx = sortIdx[i]; // current col / поточна колонка
      const Araw = a.data[idx], Braw = b.data[idx]; // raw values / сирі дані
      const nA = tryNum(Araw), nB = tryNum(Braw); // numeric check / перевірка чисел
      const bothNumbers = (nA !== null && nB !== null); // num flag / чи обидва числа

      if (i === sortIdx.length - 2) { docA = Araw; docB = Braw; } // capture doc / зафіксувати док
      if (i === sortIdx.length - 1) { pageA = nA; pageB = nB; } // capture page / зафіксувати сторінку

      if (Araw === Braw) continue; // no change / без змін
      if (!paramDiff) {
        paramResult = bothNumbers ? (nA - nB) : clean(Araw).localeCompare(clean(Braw), undefined, { numeric: true }); // sort / сорт
        paramDiff = true; // diff found / різницю знайдено
      }
    }

    if (paramDiff) return paramResult; // return diff / повернення різниці

    // TIEBREAKER / ВИРІШЕННЯ НІЧИЄЇ
    if (docA !== null && docB !== null) {
      const docCompare = clean(docA).localeCompare(clean(docB), undefined, { numeric: true }); // doc compare / порівняння док.
      if (docCompare !== 0) return docCompare; // doc diff / різниця в док.
      if (pageA !== null && pageB !== null && pageA !== pageB) return pageA - pageB; // page diff / різниця в стор.
    }

    return clean(rootA).localeCompare(clean(rootB), undefined, { numeric: true }); // final fallback / остаточне порівняння
  });

  // WRITE BACK / ЗАПИС ДАНИХ
  range.setValues(rows.map(r => r.data)); // write values / запис значень
  range.setBackgrounds(rows.map(r => r.style.bg)); // write bg / запис фону
  range.setFontColors(rows.map(r => r.style.fc)); // write colors / запис кольорів
  range.setFontWeights(rows.map(r => r.style.fw)); // write weight / запис жирності
  range.setFontStyles(rows.map(r => r.style.fs)); // write styles / запис курсиву
  range.setFontSizes(rows.map(r => r.style.fz)); // write sizes / запис розмірів

  Logger.log(`SORT DONE | ROOT: ${paramLetters[0]} | ORDER: ${paramLetters.join(" → ")} | NUMERIC: ${ascending ? "ASC" : "DESC"} | ROWS: ${values.length}`); // log log / логування результату
}