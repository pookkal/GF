
/**
* ------------------------------------------------------------------
* INSTITUTIONAL TERMINAL — COMPREHENSIVE USER GUIDE
* Industry-grade documentation for dual-mode signal engine
* ------------------------------------------------------------------
*/

function getCleanTickers(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange(3, 1, lastRow - 2, 1)
    .getValues()
    .flat()
    .filter(t => t && t.toString().trim() !== "")
    .map(t => t.toString().toUpperCase().trim());
}
