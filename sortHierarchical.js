function sortHierarchical(
  paramLetters = ["B", "C", "D", "E", "F", "H", "L", "M"],
  ascending = false,
) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 2;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < startRow) return;

  const range = sheet.getRange(startRow, 1, lastRow - 1, lastCol);

  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const fontWeights = range.getFontWeights();
  const fontStyles = range.getFontStyles();
  const fontSizes = range.getFontSizes();

  // =========================
  // HELPERS
  // =========================
  const colIdx = (l) => l.toUpperCase().charCodeAt(0) - 65;

  const clean = (v) =>
    (v === null || v === undefined) ? "" : v.toString().trim();

  const tryNum = (v) => {
    const n = parseFloat(v);
    return isNaN(n) ? null : n;
  };

  const getRoot = (name) =>
    clean(name).split(/[\s\-_]+/)[0] || "";

  const rootIdx = colIdx(paramLetters[0]);
  const sortIdx = paramLetters.slice(1).map(colIdx);

  const rows = values.map((row, i) => ({
    data: row,
    style: {
      bg: backgrounds[i],
      fc: fontColors[i],
      fw: fontWeights[i],
      fs: fontStyles[i],
      fz: fontSizes[i]
    }
  }));

  // =========================
  // COMPARE
  // =========================
  rows.sort((a, b) => {

    const rootA = getRoot(a.data[rootIdx]);
    const rootB = getRoot(b.data[rootIdx]);

    if (rootA !== rootB) {
      return rootA.localeCompare(rootB, undefined, { numeric: true });
    }

    let paramResult = 0;
    let paramDiff = false;

    // DOC / PAGE (ONLY FINAL TIEBREAKER INPUTS)
    let docA = null, docB = null;
    let pageA = null, pageB = null;

    // =========================
    // PARAM LOOP
    // =========================
    for (let i = 0; i < sortIdx.length; i++) {

      const idx = sortIdx[i];

      const Araw = a.data[idx];
      const Braw = b.data[idx];

      const nA = tryNum(Araw);
      const nB = tryNum(Braw);

      const bothNumbers = (nA !== null && nB !== null);

      // capture doc/page from last 2 params ONLY
      if (i === sortIdx.length - 2) {
        docA = Araw;
        docB = Braw;
      }

      if (i === sortIdx.length - 1) {
        pageA = nA;
        pageB = nB;
      }

      if (Araw === Braw) continue;

      if (!paramDiff) {

        paramResult = bothNumbers
          ? (nA - nB)
          : clean(Araw).localeCompare(clean(Braw), undefined, { numeric: true });

        paramDiff = true;
      }
    }

    // =========================
    // 1. PARAMS FIRST (ABSOLUTE PRIORITY)
    // =========================
    if (paramDiff) {
      return paramResult;
    }

    // =========================
    // 2. DOC + PAGE (TIEBREAKER ONLY)
    // =========================
    if (docA !== null && docB !== null) {

      const docCompare =
        clean(docA).localeCompare(clean(docB), undefined, { numeric: true });

      if (docCompare !== 0) return docCompare;

      if (pageA !== null && pageB !== null && pageA !== pageB) {
        return pageA - pageB;
      }
    }

    // =========================
    // 3. FINAL FALLBACK (NAME)
    // =========================
    return clean(rootA)
      .localeCompare(clean(rootB), undefined, { numeric: true });
  });

  // =========================
  // WRITE BACK
  // =========================
  range.setValues(rows.map(r => r.data));
  range.setBackgrounds(rows.map(r => r.style.bg));
  range.setFontColors(rows.map(r => r.style.fc));
  range.setFontWeights(rows.map(r => r.style.fw));
  range.setFontStyles(rows.map(r => r.style.fs));
  range.setFontSizes(rows.map(r => r.style.fz));

  Logger.log(`SORT DONE | ROOT: ${paramLetters[0]} | MODE: ${ascending ? "ASC" : "DESC "}`);
}