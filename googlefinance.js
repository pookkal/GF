/**
* ==============================================================================
* BASELINE LABEL: STABLE_MASTER_ALL_CLEAN_v4.0 SUPPORT_RESISTENCE_CHANGES
* ==============================================================================
*/

/**
* ------------------------------------------------------------------
*  Open LOGIC ENGINE (INSERT MENU)
* ------------------------------------------------------------------
*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Institutional Terminal')
    .addItem('🚀 1-CLICK REBUILD ALL', 'FlushAllSheetsAndBuild')
    .addItem('1. Fetch Data Only', 'generateDataSheet')
    .addSeparator()
    .addItem('2. Build Calculations', 'generateCalculationsSheet')
    .addItem('3. Refresh Dashboard Only', 'generateDashboardSheet')
    .addItem('4. Setup Chart Only', 'setupChartSheet')
    .addSeparator()
    .addItem('🤖 Generate AI Narratives', 'runInstitutionalAnalysis')
    .addSeparator()
    .addItem('📖 Open Reference Guide', 'generateReferenceSheet')
    .addSeparator()
    .addItem('🔔 Start Market Monitor', 'startMarketMonitor')
    .addItem('🔕 Stop Monitor', 'stopMarketMonitor')
    .addItem('📩 Test Alert Now', 'checkSignalsAndSendAlerts')
    .addToUi();
}


/**
* ------------------------------------------------------------------
* GLOBAL TRIGGER ENGINE (B1 CHECKBOX CLEANUP + INPUT FILTER REFRESH)
* ------------------------------------------------------------------
*/
// ------------------------------------------------------------
// UPDATED onEdit(e) — watches the NEW CHART control cells
// ------------------------------------------------------------
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const a1 = range.getA1Notation();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ------------------------------------------------------------
  // DASHBOARD update controls:
  // - B1 = Update CALCULATIONS + DASHBOARD
  // - D1 = Update DASHBOARD only
  // ------------------------------------------------------------
  if (sheet.getName() === "DASHBOARD" && (a1 === "B1" || a1 === "D1") && e.value === "TRUE") {
    ss.toast("Refreshing Dashboard...", "⚙️ TERMINAL", 3);
    try {
      if (a1 === "B1") {
        // Full refresh
        generateCalculationsSheet();
      }
      // Dashboard refresh
      generateDashboardSheet();
      ss.toast("Dashboard Synchronized.", "✅ DONE", 2);
    } catch (err) {
      ss.toast("Error: " + err.toString(), "⚠️ FAIL", 6);
    } finally {
      // reset checkbox
      sheet.getRange(a1).setValue(false);
    }
    return;
  }

  // ------------------------------------------------------------
  // INPUT filters -> refresh dashboard
  // ------------------------------------------------------------
  if (sheet.getName() === "INPUT" && (a1 === "B1" || a1 === "C1")) {
    try {
      generateDashboardSheet();
      SpreadsheetApp.flush();
    } catch (err) {
      ss.toast("Dashboard filter refresh error: " + err.toString(), "⚠️ FAIL", 6);
    }
    return;
  }

  // ------------------------------------------------------------
  // CHART controls -> update dynamic chart
  // (keep your existing watch list logic)
  // ------------------------------------------------------------
  if (sheet.getName() === "CHART") {
    const watchList = ["B1", "B2", "B3", "B4", "B6"];
    if (watchList.includes(a1) || (range.getRow() === 1 && range.getColumn() <= 4)) {
      try {
        updateDynamicChart();
      } catch (err) {
        ss.toast("Chart refresh error: " + err.toString(), "⚠️ FAIL", 6);
      }
    }
    return;
  }
}

/**
* ------------------------------------------------------------------
* 1. CUSTOM MATH FUNCTIONS (RSI, MACD, ADX, STOCH)
* ------------------------------------------------------------------
*/
function LIVERSI(history, currentPrice) {
  if (!history || !currentPrice) return 50;

  let raw = history.flat();
  let closes = [];
  for (let i = raw.length - 1; i >= 0; i--) {
    if (typeof raw[i] === 'number' && raw[i] > 0) closes.unshift(raw[i]);
    if (closes.length >= 60) break;
  }

  closes.push(Number(currentPrice));
  if (closes.length < 15) return 50;

  let gains = 0, losses = 0;
  for (let i = 1; i <= 14; i++) {
    let change = closes[i] - closes[i - 1];
    if (change > 0) gains += change; else losses += Math.abs(change);
  }
  let avgGain = gains / 14, avgLoss = losses / 14;

  for (let i = 15; i < closes.length; i++) {
    let change = closes[i] - closes[i - 1];
    let gain = change > 0 ? change : 0;
    let loss = change < 0 ? Math.abs(change) : 0;
    avgGain = ((avgGain * 13) + gain) / 14;
    avgLoss = ((avgLoss * 13) + loss) / 14;
  }

  if (avgLoss === 0) return 100;
  return Number((100 - (100 / (1 + (avgGain / avgLoss)))).toFixed(2));
}

function LIVEMACD(history, currentPrice) {
  if (!history || !currentPrice) return 0;

  let raw = history.flat();
  let closes = [];
  for (let i = raw.length - 1; i >= 0; i--) {
    if (typeof raw[i] === 'number' && raw[i] > 0) closes.unshift(raw[i]);
    if (closes.length >= 160) break;
  }

  closes.push(Number(currentPrice));
  if (closes.length < 26) return 0;

  function calculateEMA(data, period) {
    let k = 2 / (period + 1);
    let ema = data[0];
    let out = [ema];
    for (let i = 1; i < data.length; i++) {
      ema = data[i] * k + ema * (1 - k);
      out.push(ema);
    }
    return out;
  }

  const ema12 = calculateEMA(closes, 12);
  const ema26 = calculateEMA(closes, 26);
  let macdLine = closes.map((_, i) => ema12[i] - ema26[i]);
  const signal = calculateEMA(macdLine, 9);

  return Number((macdLine[macdLine.length - 1] - signal[signal.length - 1]).toFixed(3));
}

// ADX(14) (Wilder)
function LIVEADX(highHist, lowHist, closeHist, currentPrice) {
  try {
    if (!highHist || !lowHist || !closeHist) return 0;

    const Hraw = highHist.flat();
    const Lraw = lowHist.flat();
    const Craw = closeHist.flat();
    const m = Math.min(Hraw.length, Lraw.length, Craw.length);
    if (m < 60) return 0;

    const toNum = (v) => {
      if (typeof v === "number") return v;
      if (v === null || v === undefined) return NaN;
      const s = String(v).trim();
      if (s === "") return NaN;        // IMPORTANT: blank is NaN, not 0
      const n = Number(s);
      return isFinite(n) ? n : NaN;
    };

    // Build aligned OHLC rows only when ALL three are valid
    const h = [], l = [], c = [];
    for (let i = 0; i < m; i++) {
      const hi = toNum(Hraw[i]);
      const lo = toNum(Lraw[i]);
      const cl = toNum(Craw[i]);
      if (isFinite(hi) && isFinite(lo) && isFinite(cl) && hi > 0 && lo > 0 && cl > 0) {
        h.push(hi); l.push(lo); c.push(cl);
      }
    }

    const n = h.length;
    if (n < 40) return 0;

    const take = Math.min(n, 260); // more robust than 90
    const H = h.slice(n - take);
    const L = l.slice(n - take);
    const C = c.slice(n - take);

    const liveClose = toNum(currentPrice);
    if (isFinite(liveClose) && liveClose > 0) C[C.length - 1] = liveClose;

    // --- rest of your existing Wilder ADX math (unchanged) ---
    const period = 14;
    const tr = [], pdm = [], ndm = [];
    for (let i = 1; i < C.length; i++) {
      const upMove = H[i] - H[i - 1];
      const downMove = L[i - 1] - L[i];
      const plusDM = (upMove > downMove && upMove > 0) ? upMove : 0;
      const minusDM = (downMove > upMove && downMove > 0) ? downMove : 0;

      const r1 = H[i] - L[i];
      const r2 = Math.abs(H[i] - C[i - 1]);
      const r3 = Math.abs(L[i] - C[i - 1]);
      const trueRange = Math.max(r1, r2, r3);

      if (!isFinite(trueRange) || trueRange < 0) return 0;
      tr.push(trueRange); pdm.push(plusDM); ndm.push(minusDM);
    }
    if (tr.length < period * 2) return 0;

    const safeDiv = (num, den) => (den > 1e-12 ? (num / den) : 0);

    let atr = tr.slice(0, period).reduce((a, b) => a + b, 0);
    let pDM14 = pdm.slice(0, period).reduce((a, b) => a + b, 0);
    let nDM14 = ndm.slice(0, period).reduce((a, b) => a + b, 0);

    let pDI = 100 * safeDiv(pDM14, atr);
    let nDI = 100 * safeDiv(nDM14, atr);

    const dxArr = [];
    dxArr.push((pDI + nDI > 1e-12) ? (100 * Math.abs(pDI - nDI) / (pDI + nDI)) : 0);

    for (let i = period; i < tr.length; i++) {
      atr = atr - (atr / period) + tr[i];
      pDM14 = pDM14 - (pDM14 / period) + pdm[i];
      nDM14 = nDM14 - (nDM14 / period) + ndm[i];
      if (!isFinite(atr) || atr <= 0) return 0;

      pDI = 100 * safeDiv(pDM14, atr);
      nDI = 100 * safeDiv(nDM14, atr);

      const dx = (pDI + nDI > 1e-12) ? (100 * Math.abs(pDI - nDI) / (pDI + nDI)) : 0;
      dxArr.push(isFinite(dx) ? dx : 0);
    }

    let adx = dxArr.slice(0, period).reduce((a, b) => a + b, 0) / period;
    for (let i = period; i < dxArr.length; i++) adx = ((adx * (period - 1)) + dxArr[i]) / period;

    return Number((isFinite(adx) ? adx : 0).toFixed(2));
  } catch (e) {
    return 0;
  }
}


// Stoch %K(14) in 0..1
function LIVESTOCHK(highHist, lowHist, closeHist, currentPrice) {
  try {
    if (!highHist || !lowHist || !closeHist || !currentPrice) return 0.5;

    const H = highHist.flat().filter(n => typeof n === 'number' && n > 0);
    const L = lowHist.flat().filter(n => typeof n === 'number' && n > 0);
    const C = closeHist.flat().filter(n => typeof n === 'number' && n > 0);

    const n = Math.min(H.length, L.length, C.length);
    if (n < 20) return 0.5;

    const period = 14;
    const h = H.slice(n - period);
    const l = L.slice(n - period);

    const hh = Math.max(...h);
    const ll = Math.min(...l);

    const close = Number(currentPrice);
    if (hh === ll) return 0.5;

    const k = (close - ll) / (hh - ll);
    return Number(Math.max(0, Math.min(1, k)).toFixed(4));
  } catch (e) {
    return 0.5;
  }
}


/**
* ------------------------------------------------------------------
* 2. CORE AUTOMATION
* ------------------------------------------------------------------
*/
function FlushAllSheetsAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["DATA", "CALCULATIONS", "CHART", "DASHBOARD"];
  const ui = SpreadsheetApp.getUi();

  if (ui.alert('🚨 Full Rebuild', 'Rebuild the sheets?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Syncing Global Data..."), "Status");
  generateDataSheet();
  SpreadsheetApp.flush();
  Utilities.sleep(1500);

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>2/4:</b> Integrating Indicators..."), "Status");
  generateCalculationsSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>3/4:</b> Building Dashboard..."), "Status");
  generateDashboardSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>4/4:</b> Constructing Chart..."), "Status");
  setupChartSheet();

  ui.alert('✅ Rebuild Complete', 'Terminal Online. Data links restored.', ui.ButtonSet.OK);
}


/**
* ------------------------------------------------------------------
* 3. DATA ENGINE
* ------------------------------------------------------------------
*/
/**
* ------------------------------------------------------------------
* 3. DATA ENGINE (FULL FUNCTION — ROW 2 TICKER, ROW 3 ATH/PE/EPS IN A..F)
* ------------------------------------------------------------------
* Layout per ticker block (7 columns):
* - Row 2, colStart            : Ticker (bold)
* - Row 3, colStart..colStart+5: Metadata in A..F (ATH / P-E / EPS)
*     A: "ATH:"     B: ATH value
*     C: "P/E:"     D: P/E value
*     E: "EPS:"     F: EPS value
* - Row 4, colStart..colStart+5: GOOGLEFINANCE("all") header row (Date, Open, High, Low, Close, Volume)
* - Row 5+                       : OHLCV data
*
* Impact:
* - DATA consumers that already use OHLCV from row 5 are unchanged.
* - CALCULATIONS that references ATH at DATA!(row 3) remains compatible (A/B of row 3).
* - Adds cached P/E and EPS in DATA row 3 for optional reuse (faster vs repeated GOOGLEFINANCE calls elsewhere).
* ------------------------------------------------------------------
*/
function generateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;

  const tickers = getCleanTickers(inputSheet);
  let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");

  // Clear
  dataSheet.clear({ contentsOnly: true });
  dataSheet.clearFormats();

  // Timestamp
  dataSheet.getRange("A1")
    .setValue("Last Update: " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm"))
    .setFontWeight("bold")
    .setFontColor("blue");

  if (tickers.length === 0) return;

  const colsPer = 7;
  const totalCols = tickers.length * colsPer;

  // Ensure enough columns
  if (dataSheet.getMaxColumns() < totalCols) {
    dataSheet.insertColumnsAfter(dataSheet.getMaxColumns(), totalCols - dataSheet.getMaxColumns());
  }

  // ------------------------------------------------------------
  // Row 2: Tickers
  // ------------------------------------------------------------
  const row2 = new Array(totalCols).fill("");
  for (let i = 0; i < tickers.length; i++) {
    row2[i * colsPer] = tickers[i];
  }
  dataSheet.getRange(2, 1, 1, totalCols)
    .setValues([row2])
    .setNumberFormat("@")
    .setFontWeight("bold");

  // ------------------------------------------------------------
  // Row 3: Formulas first (ATH / P-E / EPS values only)
  // IMPORTANT: do NOT write "" formulas into label cells.
  // We'll write labels AFTER formulas.
  // ------------------------------------------------------------
  const row3Formulas = new Array(totalCols).fill("");
  for (let i = 0; i < tickers.length; i++) {
    const t = tickers[i];
    const b = i * colsPer;

    // value cells only
    row3Formulas[b + 1] =
      `=MAX(QUERY(GOOGLEFINANCE("${t}","high","1/1/2000",TODAY()),"SELECT Col2 LABEL Col2 ''"))`;
    row3Formulas[b + 3] =
      `=IFERROR(GOOGLEFINANCE("${t}","pe"),"")`;
    row3Formulas[b + 5] =
      `=IFERROR(GOOGLEFINANCE("${t}","eps"),"")`;
  }
  dataSheet.getRange(3, 1, 1, totalCols).setFormulas([row3Formulas]);

  // Now write labels (cannot be overwritten now)
  for (let i = 0; i < tickers.length; i++) {
    const c = (i * colsPer) + 1; // 1-based
    dataSheet.getRange(3, c).setValue("ATH:");
    dataSheet.getRange(3, c + 2).setValue("P/E:");
    dataSheet.getRange(3, c + 4).setValue("EPS:");
  }

  // ------------------------------------------------------------
  // Row 4: GOOGLEFINANCE(all)
  // ------------------------------------------------------------
  const row4Formulas = new Array(totalCols).fill("");
  for (let i = 0; i < tickers.length; i++) {
    const t = tickers[i];
    row4Formulas[i * colsPer] =
      `=IFERROR(GOOGLEFINANCE("${t}","all",TODAY()-800,TODAY()),"No Data")`;
  }
  dataSheet.getRange(4, 1, 1, totalCols).setFormulas([row4Formulas]);

  // ------------------------------------------------------------
  // Number formats (row 3 values)
  // ------------------------------------------------------------
  for (let i = 0; i < tickers.length; i++) {
    const c = (i * colsPer) + 1; // 1-based
    dataSheet.getRange(3, c + 1).setNumberFormat("#,##0.00"); // ATH value
    dataSheet.getRange(3, c + 3).setNumberFormat("0.00");     // P/E value
    dataSheet.getRange(3, c + 5).setNumberFormat("0.00");     // EPS value
  }

  // ------------------------------------------------------------
  // Label styling (guaranteed visible)
  // ------------------------------------------------------------
  const LABEL_BG = "#1F2937";
  const LABEL_FG = "#F9FAFB";

  const labelA1s = [];
  for (let i = 0; i < tickers.length; i++) {
    const c = (i * colsPer) + 1; // 1-based
    labelA1s.push(dataSheet.getRange(3, c).getA1Notation());       // ATH label
    labelA1s.push(dataSheet.getRange(3, c + 2).getA1Notation());   // P/E label
    labelA1s.push(dataSheet.getRange(3, c + 4).getA1Notation());   // EPS label
  }

  dataSheet.getRangeList(labelA1s)
    .setBackground(LABEL_BG)
    .setFontColor(LABEL_FG)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  // ------------------------------------------------------------
  // Historical formatting (rows 5+)
  // ------------------------------------------------------------
  for (let i = 0; i < tickers.length; i++) {
    const colStart = (i * colsPer) + 1; // 1-based
    dataSheet.getRange(5, colStart, 1000, 1).setNumberFormat("yyyy-mm-dd");
    dataSheet.getRange(5, colStart + 1, 1000, 5).setNumberFormat("#,##0.00");
  }

  SpreadsheetApp.flush();
}


function getCleanTickers(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange(3, 1, lastRow - 2, 1)
    .getValues()
    .flat()
    .filter(t => t && t.toString().trim() !== "")
    .map(t => t.toString().toUpperCase().trim());
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function forceExpandSheet(sheet, targetCols) {
  if (sheet.getMaxColumns() < targetCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), targetCols - sheet.getMaxColumns());
  }
}

/**
* ------------------------------------------------------------------
* 4. CALCULATION ENGINE (FULL FUNCTION — UPDATED)
* - Fixes: SELL-side decisions (Take Profit / Reduce)
* - Fixes: Locale separator auto-handled (; vs ,)
* - Formatting: LEFT align + WRAP + row height ~4 lines (72px)
* - Preserves: LAST_STATE in AB
* ------------------------------------------------------------------
* Columns (A..AB):
* A  Ticker
* B  SIGNAL
* C  DECISION
* D  FUNDAMENTAL
* E  Price
* F  Change %
* G  Vol Trend
* H  ATH (TRUE)
* I  ATH Diff %
* J  R:R Quality
* K  Trend Score
* L  Trend State
* M  SMA 20
* N  SMA 50
* O  SMA 200
* P  RSI
* Q  MACD Hist
* R  Divergence
* S  ADX (14)
* T  Stoch %K (14)
* U  Support
* V  Resistance
* W  Target (3:1)
* X  ATR (14)
* Y  Bollinger %B
* Z  TECH NOTES
* AA FUND NOTES
* AB LAST_STATE
* ------------------------------------------------------------------
*/
function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet || !inputSheet) return;

  const tickers = getCleanTickers(inputSheet);
  let calc = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");

  // Locale separator: US typically ","; EU typically ";"
  const locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
  const SEP = (/^(en|en_)/.test(locale)) ? "," : ";";

  // Persist LAST_STATE (AB)
  const stateMap = {};
  if (calc.getLastRow() >= 3) {
    const existing = calc.getRange(3, 1, calc.getLastRow() - 2, 28).getValues();
    existing.forEach(r => {
      const t = (r[0] || "").toString().trim().toUpperCase();
      if (t) stateMap[t] = r[27]; // AB
    });
  }

  calc.clear().clearFormats();

  // ------------------------------------------------------------------
  // ROW 1: GROUP HEADERS (MERGED) + timestamp in AB1
  // ------------------------------------------------------------------
  const syncTime = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const styleGroup = (a1, label, bg) => {
    calc.getRange(a1).merge()
      .setValue(label)
      .setBackground(bg)
      .setFontColor("white")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  };

  styleGroup("A1:A1",   "IDENTITY",        "#263238"); // A
  styleGroup("B1:D1",   "SIGNALING",       "#0D47A1"); // B-D
  styleGroup("E1:G1",   "PRICE / VOLUME",  "#1B5E20"); // E-G
  styleGroup("H1:J1",   "PERFORMANCE",     "#004D40"); // H-J
  styleGroup("K1:O1",   "TREND",           "#2E7D32"); // K-O
  styleGroup("P1:T1",   "MOMENTUM",        "#33691E"); // P-T
  styleGroup("U1:Y1",   "LEVELS / RISK",   "#B71C1C"); // U-Y
  styleGroup("Z1:AA1",  "NOTES",           "#212121"); // Z-AA

  calc.getRange("AB1")
    .setValue(syncTime)
    .setBackground("#000000")
    .setFontColor("#00FF00")
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // ------------------------------------------------------------------
  // ROW 2: COLUMN HEADERS
  // ------------------------------------------------------------------
  const headers = [[
    "Ticker","SIGNAL","DECISION","FUNDAMENTAL","Price","Change %","Vol Trend","ATH (TRUE)","ATH Diff %","R:R Quality",
    "Trend Score","Trend State","SMA 20","SMA 50","SMA 200","RSI","MACD Hist","Divergence","ADX (14)","Stoch %K (14)",
    "Support","Resistance","Target (3:1)","ATR (14)","Bollinger %B","TECH NOTES","FUND NOTES","LAST_STATE"
  ]];

  calc.getRange(2, 1, 1, 28)
    .setValues(headers)
    .setBackground("#111111")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  // Write tickers in A3:A
  if (tickers.length > 0) {
    calc.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  }

  // ------------------------------------------------------------------
  // FORMULAS
  // ------------------------------------------------------------------
  const formulas = [];
  const restoredStates = [];

  const BLOCK = 7; // DATA block width (must match generateDataSheet)

  tickers.forEach((ticker, i) => {
    const row = i + 3;
    const t = String(ticker || "").trim().toUpperCase();
    restoredStates.push([stateMap[t] || ""]);

    // DATA block start (each ticker is 7 cols in DATA)
    const tDS = (i * BLOCK) + 1; // colStart
    const dateCol  = columnToLetter(tDS + 0); // Date (row 5+)
    const openCol  = columnToLetter(tDS + 1); // Open
    const highCol  = columnToLetter(tDS + 2); // High
    const lowCol   = columnToLetter(tDS + 3); // Low
    const closeCol = columnToLetter(tDS + 4); // Close
    const volCol   = columnToLetter(tDS + 5); // Volume

    // Cached fundamentals in DATA row 3 (within same block)
    const athCell = `DATA!${columnToLetter(tDS + 1)}3`; // ATH value at colStart+1
    const peCell  = `DATA!${columnToLetter(tDS + 3)}3`; // P/E value at colStart+3
    const epsCell = `DATA!${columnToLetter(tDS + 5)}3`; // EPS value at colStart+5

    // Rolling window anchors (row 5+ only)
    const lastRowCount = `COUNTA(DATA!$${closeCol}$5:$${closeCol})`; // number of data rows
    const lastAbsRow   = `(4+${lastRowCount})`;                      // absolute row index
    const lastRowFormula = "COUNTA(DATA!$A:$A)";                      //used for support /resistence , to stay live

    // SIGNAL (B) — locale-safe + row5-anchored windows
    const fSignal =
      `=IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}"LOADING"${SEP}` +
        `IFS(` +
          `$E${row}<$U${row}${SEP}"Stop-Out"${SEP}` +
          `$E${row}<$O${row}${SEP}"Risk-Off (Below SMA200)"${SEP}` +
          `$X${row}<=MIN(ARRAYFORMULA(` +
            `OFFSET(DATA!$${highCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20)` +
            `-OFFSET(DATA!$${lowCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20)` +
          `))${SEP}"Volatility Squeeze (Coiling)"${SEP}` +
          `$S${row}<15${SEP}"Range-Bound (Low ADX)"${SEP}` +
          `AND($G${row}>=1.5${SEP}$E${row}>=$V${row}*0.995)${SEP}"Breakout (High Volume)"${SEP}` +
          `AND($T${row}<=0.20${SEP}$E${row}>$U${row})${SEP}"Mean Reversion (Oversold)"${SEP}` +
          `AND($E${row}>$O${row}${SEP}$Q${row}>0${SEP}$S${row}>=18)${SEP}"Trend Continuation"${SEP}` +
          `TRUE${SEP}"Hold / Monitor"` +
        `)` +
      `)`;

    // DECISION (C) — unchanged gating pattern (kept stable)
    const fDecision =
      `=IF($A${row}=""${SEP}""${SEP}
        IFS(
          AND(IFERROR(VALUE($E${row})${SEP}0)>0${SEP}IFERROR(VALUE($U${row})${SEP}0)>0${SEP}IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($U${row})${SEP}0))${SEP}"Stop-Out"${SEP}

          AND(
            REGEXMATCH(
              UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP} MATCH($A${row}${SEP} INPUT!$A$3:$A${SEP} 0))${SEP} ""))${SEP}
              "(^|,|\\s)PURCHASED(\\s|,|$)"
            )${SEP}
            IFERROR(VALUE($E${row})${SEP}0)>0${SEP}
            IFERROR(VALUE($W${row})${SEP}0)>0${SEP}
            IFERROR(VALUE($E${row})${SEP}0)>=IFERROR(VALUE($W${row})${SEP}0)
          )${SEP}"Take Profit"${SEP}

          AND(
            REGEXMATCH(
              UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP} MATCH($A${row}${SEP} INPUT!$A$3:$A${SEP} 0))${SEP} ""))${SEP}
              "(^|,|\\s)PURCHASED(\\s|,|$)"
            )${SEP}
            IFERROR(VALUE($V${row})${SEP}0)>0${SEP}
            IFERROR(VALUE($E${row})${SEP}0)>=IFERROR(VALUE($V${row})${SEP}0)*0.995${SEP}
            OR(IFERROR(VALUE($P${row})${SEP}50)>=70${SEP} IFERROR(VALUE($T${row})${SEP}0.5)>=0.8)
          )${SEP}"Take Profit"${SEP}

          AND(
            REGEXMATCH(
              UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP} MATCH($A${row}${SEP} INPUT!$A$3:$A${SEP} 0))${SEP} ""))${SEP}
              "(^|,|\\s)PURCHASED(\\s|,|$)"
            )${SEP}
            IFERROR(VALUE($Q${row})${SEP}0)<0${SEP}
            IFERROR(VALUE($N${row})${SEP}0)>0${SEP}
            IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($N${row})${SEP}0)
          )${SEP}"Reduce (Momentum Weak)"${SEP}

          AND(
            REGEXMATCH(
              UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP} MATCH($A${row}${SEP} INPUT!$A$3:$A${SEP} 0))${SEP} ""))${SEP}
              "(^|,|\\s)PURCHASED(\\s|,|$)"
            )${SEP}
            IFERROR(VALUE($X${row})${SEP}0)>0${SEP}
            IFERROR(VALUE($M${row})${SEP}0)>0${SEP}
            (IFERROR(VALUE($E${row})${SEP}0)-IFERROR(VALUE($M${row})${SEP}0))/IFERROR(VALUE($X${row})${SEP}1) >= 2
          )${SEP}"Reduce (Overextended)"${SEP}

          AND(
            REGEXMATCH(
              UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP} MATCH($A${row}${SEP} INPUT!$A$3:$A${SEP} 0))${SEP} ""))${SEP}
              "(^|,|\\s)PURCHASED(\\s|,|$)"
            )${SEP}
            IFERROR(VALUE($U${row})${SEP}0)>0${SEP}
            IFERROR(VALUE($E${row})${SEP}0)>IFERROR(VALUE($U${row})${SEP}0)${SEP}
            IFERROR(VALUE($T${row})${SEP}0.5)<=0.2${SEP}
            NOT(IFERROR(VALUE($E${row})${SEP}0) < IFERROR(VALUE($O${row})${SEP}0))
          )${SEP}"Add in Dip"${SEP}

          IFERROR(VALUE($E${row})${SEP}0) < IFERROR(VALUE($O${row})${SEP}0)${SEP}"Avoid"${SEP}

          AND($B${row}="Breakout (High Volume)"${SEP}OR($D${row}="VALUE"${SEP}$D${row}="FAIR"))${SEP}"Trade Long"${SEP}
          AND($B${row}="Breakout (High Volume)"${SEP}OR($D${row}="EXPENSIVE"${SEP}$D${row}="PRICED FOR PERFECTION"))${SEP}"Hold"${SEP}
          AND($B${row}="Trend Continuation"${SEP}$D${row}="VALUE")${SEP}"Accumulate"${SEP}
          $B${row}="Trend Continuation"${SEP}"Hold"${SEP}
          TRUE${SEP}"Hold"
        )
      )`;

    // FUNDAMENTAL (D) — reads cached PE/EPS from DATA row 3 (fast)
    const fFund =
      `=IFERROR(` +
        `LET(` +
          `pe${SEP}IFERROR(VALUE(${peCell})${SEP}"" )${SEP}` +
          `eps${SEP}IFERROR(VALUE(${epsCell})${SEP}"" )${SEP}` +
          `IFS(` +
            `OR(pe=""${SEP}eps="")${SEP}"FAIR"${SEP}` +
            `eps<=0${SEP}"ZOMBIE"${SEP}` +
            `pe>=60${SEP}"PRICED FOR PERFECTION"${SEP}` +
            `pe>=35${SEP}"EXPENSIVE"${SEP}` +
            `AND(pe>0${SEP}pe<=25${SEP}eps>=0.5)${SEP}"VALUE"${SEP}` +
            `AND(pe>25${SEP}pe<35${SEP}eps>=0.5)${SEP}"FAIR"${SEP}` +
            `TRUE${SEP}"FAIR"` +
          `)` +
        `)` +
      `${SEP}"FAIR")`;

    // E..Y
    const fPrice  = `=ROUND(IFERROR(GOOGLEFINANCE("${t}"${SEP}"price")${SEP}0)${SEP}2)`;
    const fChg    = `=IFERROR(GOOGLEFINANCE("${t}"${SEP}"changepct")/100${SEP}0)`;

    const fRVOL =
      `=ROUND(` +
        `IFERROR(` +
          `OFFSET(DATA!$${volCol}$5${SEP}${lastRowCount}-1${SEP}0)` +
          `/AVERAGE(OFFSET(DATA!$${volCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))` +
        `${SEP}1)` +
      `${SEP}2)`;

    const fATH    = `=IFERROR(${athCell}${SEP}0)`;
    const fATHPct = `=IFERROR(($E${row}-$H${row})/MAX(0.01${SEP}$H${row})${SEP}0)`;

    const fRR =
      `=IF(OR($E${row}<=$U${row}${SEP}$E${row}=0)${SEP}0${SEP}` +
        `ROUND(MAX(0${SEP}$V${row}-$E${row})/MAX($X${row}*0.5${SEP}$E${row}-$U${row})${SEP}2)` +
      `)`;

    const fStars  = `=REPT("★"${SEP} ($E${row}>$M${row}) + ($E${row}>$N${row}) + ($E${row}>$O${row}))`;
    const fTrend  = `=IF($E${row}>$O${row}${SEP}"BULL"${SEP}"BEAR")`;

    const fSMA20  = `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))${SEP}0)${SEP}2)`;
    const fSMA50  = `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$5${SEP}${lastRowCount}-50${SEP}0${SEP}50))${SEP}0)${SEP}2)`;
    const fSMA200 = `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$5${SEP}${lastRowCount}-200${SEP}0${SEP}200))${SEP}0)${SEP}2)`;

    const fRSI    = `=LIVERSI(DATA!$${closeCol}$5:$${closeCol}${SEP}$E${row})`;
    const fMACD   = `=LIVEMACD(DATA!$${closeCol}$5:$${closeCol}${SEP}$E${row})`;

    const fDiv =
      `=IFERROR(IFS(` +
        `AND($E${row}<INDEX(DATA!$${closeCol}:$${closeCol}${SEP}${lastAbsRow}-14)${SEP}$P${row}>50)${SEP}"BULL DIV"${SEP}` +
        `AND($E${row}>INDEX(DATA!$${closeCol}:$${closeCol}${SEP}${lastAbsRow}-14)${SEP}$P${row}<50)${SEP}"BEAR DIV"${SEP}` +
        `TRUE${SEP}"—")${SEP}"—")`;

    const fADX    = `=IFERROR(LIVEADX(DATA!$${highCol}$5:$${highCol}${SEP}DATA!$${lowCol}$5:$${lowCol}${SEP}DATA!$${closeCol}$5:$${closeCol}${SEP}$E${row})${SEP}0)`;
    const fStoch  = `=LIVESTOCHK(DATA!$${highCol}$5:$${highCol}${SEP}DATA!$${lowCol}$5:$${lowCol}${SEP}DATA!$${closeCol}$5:$${closeCol}${SEP}$E${row})`;

    const fRes = `=LET(window${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}ROUND(IFERROR(AVERAGE(LARGE(OFFSET(DATA!$${closeCol}$5${SEP}${lastRowFormula}-(window+1)${SEP}0${SEP}window)${SEP}{1${SEP}2${SEP}3}))${SEP}$E${row}*1.05)${SEP}2))`;

    const fSup = `=LET(window${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}ROUND(IFERROR(AVERAGE(SMALL(OFFSET(DATA!$${lowCol}$5${SEP}${lastRowFormula}-(window+1)${SEP}0${SEP}window)${SEP}{1${SEP}2${SEP}3}))${SEP}$E${row}*0.95)${SEP}2))`;
    
    // Target: Hybrid Logic (High of Resistance vs. 3:1 Projection)
    const fTgt = `=ROUND(MAX($V${row}${SEP}$E${row}+(($E${row}-$U${row})*3))${SEP}2)`;

    const fATR =
      `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(` +
        `OFFSET(DATA!$${highCol}$5${SEP}${lastRowCount}-14${SEP}0${SEP}14)` +
        `-OFFSET(DATA!$${lowCol}$5${SEP}${lastRowCount}-14${SEP}0${SEP}14)` +
      `))${SEP}0)${SEP}2)`;

    const fBBP =
      `=ROUND(IFERROR((($E${row}-$M${row})/(4*STDEV(OFFSET(DATA!$${closeCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))))+0.5${SEP}0.5)${SEP}2)`;

    // Z TECH NOTES — parse-safe + correct columns + Stoch shown as %
    const fTechNotes =
      `=IF($A${row}=""${SEP}""${SEP}` +
        `"VOL: RVOL "&TEXT(IFERROR(VALUE($G${row})${SEP}0)${SEP}"0.00")&"x; "&` +
          `IF(IFERROR(VALUE($G${row})${SEP}0)<1${SEP}"sub-average (weak sponsorship)."${SEP}"healthy participation.")&CHAR(10)&` +

        `"REGIME: Price "&TEXT(IFERROR(VALUE($E${row})${SEP}0)${SEP}"0.00")&" vs SMA200 "&` +
          `TEXT(IFERROR(VALUE($O${row})${SEP}0)${SEP}"0.00")&"; "&` +
          `IF(IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($O${row})${SEP}0)${SEP}"risk-off below SMA200."${SEP}"risk-on above SMA200.")&CHAR(10)&` +

        `"VOL/STRETCH: ATR(14) "&TEXT(IFERROR(VALUE($X${row})${SEP}0)${SEP}"0.00")&"; stretch "&` +
          `IF(` +
            `OR(IFERROR(VALUE($X${row})${SEP}0)=0${SEP}IFERROR(VALUE($M${row})${SEP}0)=0)${SEP}` +
            `"—"${SEP}` +
            `TEXT((IFERROR(VALUE($E${row})${SEP}0)-IFERROR(VALUE($M${row})${SEP}0))/IFERROR(VALUE($X${row})${SEP}1)${SEP}"0.0")&"x ATR"` +
          `)&" (<= +/-2x)."&CHAR(10)&` +

        `"MOMENTUM: RSI(14) "&TEXT(IFERROR(VALUE($P${row})${SEP}0)${SEP}"0.0")&"; "&` +
          `IF(IFERROR(VALUE($P${row})${SEP}0)<40${SEP}"negative bias."${SEP}"constructive.")&` +
          `" MACD hist "&TEXT(IFERROR(VALUE($Q${row})${SEP}0)${SEP}"0.000")&"; "&` +
          `IF(IFERROR(VALUE($Q${row})${SEP}0)>0${SEP}"improving."${SEP}"weak.")&CHAR(10)&` +

        `"TREND: ADX(14) "&TEXT(IFERROR(VALUE($S${row})${SEP}0)${SEP}"0.0")&"; "&` +
          `IF(IFERROR(VALUE($S${row})${SEP}0)>=25${SEP}"strong."${SEP}"weak.")&` +
          `" Stoch %K "&TEXT(IFERROR(VALUE($T${row})${SEP}0)${SEP}"0.0%")&" — "&` +
          `IF(IFERROR(VALUE($T${row})${SEP}0)<=0.2${SEP}"oversold zone (mean-reversion potential)."${SEP}` +
            `IF(IFERROR(VALUE($T${row})${SEP}0)>=0.8${SEP}"overbought zone (pullback risk)."${SEP}"neutral range (no timing edge)."))&CHAR(10)&` +

        `"R:R: "&TEXT(IFERROR(VALUE($J${row})${SEP}0)${SEP}"0.00")&"x; "&` +
          `IF(IFERROR(VALUE($J${row})${SEP}0)>=3${SEP}"favorable."${SEP}"limited")` +
      `)`;

    // AA FUND NOTES — rewritten to include: FUNDAMENTAL, REASON, SIGNAL (with WHY), Why Not, Action, RISK FLAG
    // AA FUND NOTES — WHY moved to its own line ("Why:")
    const fFundNotes =
      `=IF($A${row}=""${SEP}""${SEP}
      "FUNDAMENTAL: "&IFS(
        $D${row}="VALUE"${SEP}"Positive (tailwind)"
        ${SEP}$D${row}="FAIR"${SEP}"Neutral (not a blocker)"
        ${SEP}$D${row}="EXPENSIVE"${SEP}"Negative (headwind)"
        ${SEP}$D${row}="PRICED FOR PERFECTION"${SEP}"High expectations (fragile)"
        ${SEP}$D${row}="ZOMBIE"${SEP}"High risk (weak earnings)"
        ${SEP}TRUE${SEP}"Neutral"
      )&CHAR(10)&
      "REASON: "&IFS(
        $D${row}="ZOMBIE"${SEP}"EPS <= 0 (loss-making / weak earnings quality)."
        ${SEP}$D${row}="PRICED FOR PERFECTION"${SEP}"PE >= 60 (priced for flawless execution)."
        ${SEP}$D${row}="EXPENSIVE"${SEP}"PE in 35–59 range (valuation premium; lower margin for error)."
        ${SEP}$D${row}="VALUE"${SEP}"EPS >= 0.5 and PE <= 25 (supportive valuation)."
        ${SEP}$D${row}="FAIR"${SEP}"EPS positive but below quality threshold or PE 26–34 (neutral)."
        ${SEP}TRUE${SEP}"Valuation classification unavailable."
      )&CHAR(10)&CHAR(10)&

      "SIGNAL: "&$B${row}&CHAR(10)&

      "Why: "&IFS(
        IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($U${row})${SEP}0)
          ${SEP}"Price below support → Stop-Out."
        ${SEP}IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($O${row})${SEP}0)
          ${SEP}"Price below SMA200 → Risk-Off regime."
        ${SEP}IFERROR(VALUE($X${row})${SEP}0)<=
            MIN(ARRAYFORMULA(
              OFFSET(DATA!$${highCol}$5${SEP}COUNTA(DATA!$${closeCol}:$${closeCol})-21${SEP}0${SEP}20) -
              OFFSET(DATA!$${lowCol}$5${SEP}COUNTA(DATA!$${closeCol}:$${closeCol})-21${SEP}0${SEP}20)
            ))
          ${SEP}"ATR compressed → Volatility squeeze / coiling."
        ${SEP}IFERROR(VALUE($S${row})${SEP}0)<15
          ${SEP}"ADX below 15 → Range-bound market."
        ${SEP}AND(IFERROR(VALUE($G${row})${SEP}0)>=1.5${SEP}
                IFERROR(VALUE($E${row})${SEP}0)>=IFERROR(VALUE($V${row})${SEP}0)*0.995)
          ${SEP}"High volume near resistance → Breakout setup."
        ${SEP}AND(IFERROR(VALUE($T${row})${SEP}0)<=0.20${SEP}
                IFERROR(VALUE($E${row})${SEP}0)>IFERROR(VALUE($U${row})${SEP}0))
          ${SEP}"Stoch oversold above support → Mean-reversion setup."
        ${SEP}AND(IFERROR(VALUE($E${row})${SEP}0)>IFERROR(VALUE($O${row})${SEP}0)${SEP}
                IFERROR(VALUE($Q${row})${SEP}0)>0${SEP}
                IFERROR(VALUE($S${row})${SEP}0)>=18)
          ${SEP}"Above SMA200 with momentum and trend → Trend continuation."
        ${SEP}TRUE${SEP}"No dominant condition → Hold / Monitor."
      )&CHAR(10)&

      "Why Not: "&IFS(
        $B${row}="Stop-Out"
          ${SEP}"All alternatives blocked — price already below support (highest-priority invalidation)."
        ${SEP}AND(IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($O${row})${SEP}0)${SEP}$B${row}<>"Stop-Out")
          ${SEP}"Risk-Off regime — price below SMA200 blocks trend, breakout, and accumulation setups."
        ${SEP}AND(IFERROR(VALUE($S${row})${SEP}0)<15${SEP}$B${row}<>"Range-Bound (Low ADX)")
          ${SEP}"ADX < 15 indicates chop — trend and breakout setups suppressed; expect range behavior."
        ${SEP}AND(
            IFERROR(VALUE($X${row})${SEP}0)<=
              MIN(ARRAYFORMULA(
                OFFSET(DATA!$${highCol}$5${SEP}COUNTA(DATA!$${closeCol}:$${closeCol})-21${SEP}0${SEP}20) -
                OFFSET(DATA!$${lowCol}$5${SEP}COUNTA(DATA!$${closeCol}:$${closeCol})-21${SEP}0${SEP}20)
              ))
            ${SEP}$B${row}<>"Volatility Squeeze (Coiling)"
          )
          ${SEP}"ATR compression detected — volatility squeeze takes priority over directional trades until expansion."
        ${SEP}AND(IFERROR(VALUE($G${row})${SEP}0)<1.5${SEP}$B${row}<>"Breakout (High Volume)")
          ${SEP}"Relative volume < 1.5x — insufficient participation for breakout confirmation."
        ${SEP}AND(IFERROR(VALUE($J${row})${SEP}0)<1.5${SEP}$C${row}<>"Trade Long")
          ${SEP}"R:R below threshold — asymmetry insufficient for new exposure at current levels."
        ${SEP}TRUE
          ${SEP}"No explicit higher-priority blocker detected."
      )&CHAR(10)&CHAR(10)&

      "Action: "&IFS(
        $C${row}="Stop-Out"${SEP}"EXIT / INVALIDATED"
        ${SEP}$C${row}="Avoid"${SEP}"NO TRADE"
        ${SEP}$C${row}="Trade Long"${SEP}"ENTER LONG"
        ${SEP}$C${row}="Accumulate"${SEP}"ADD / SCALE IN"
        ${SEP}TRUE${SEP}$C${row}
      )&

      IF(
        AND(
          OR($B${row}="Breakout (High Volume)"${SEP}$B${row}="Trend Continuation")${SEP}
          OR($D${row}="ZOMBIE"${SEP}$D${row}="PRICED FOR PERFECTION"${SEP}$D${row}="EXPENSIVE")
        )
        ${SEP}CHAR(10)&"RISK FLAG (HIGH): Momentum vs weak / fragile fundamentals."
        ${SEP}IF(
          AND(
            OR($B${row}="Mean Reversion (Oversold)"${SEP}$B${row}="Stop-Out")${SEP}
            $D${row}="VALUE"
          )
          ${SEP}CHAR(10)&"RISK FLAG (MED): Value present, but structure not yet aligned."
          ${SEP}""
        )
      )
      )`;

    formulas.push([
      fSignal,      // B
      fDecision,    // C
      fFund,        // D
      fPrice,       // E
      fChg,         // F
      fRVOL,        // G
      fATH,         // H
      fATHPct,      // I
      fRR,          // J
      fStars,       // K
      fTrend,       // L
      fSMA20,       // M
      fSMA50,       // N
      fSMA200,      // O
      fRSI,         // P
      fMACD,        // Q
      fDiv,         // R
      fADX,         // S
      fStoch,       // T
      fSup,         // U
      fRes,         // V
      fTgt,         // W
      fATR,         // X
      fBBP,         // Y
      fTechNotes,   // Z
      fFundNotes    // AA
    ]);
  });

  if (tickers.length > 0) {
    // B..AA (26 cols)
    calc.getRange(3, 2, formulas.length, 26).setFormulas(formulas);
    // AB LAST_STATE restore
    calc.getRange(3, 28, restoredStates.length, 1).setValues(restoredStates);
  }

  // ------------------------------------------------------------------
  // FORMATTING (kept consistent with your current style)
  // ------------------------------------------------------------------
  const lr = Math.max(calc.getLastRow(), 3);
  calc.setFrozenRows(2);

  if (lr > 2) {
    const dataRows = lr - 2;
    calc.setRowHeights(3, dataRows, 72);
    calc.getRange(3, 1, dataRows, 28)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle")
      .setWrap(true);
  }

  for (let c = 1; c <= 25; c++) calc.setColumnWidth(c, 90);
  calc.setColumnWidth(26, 420); // Z TECH NOTES
  calc.setColumnWidth(27, 420); // AA FUND NOTES
  calc.setColumnWidth(28, 140); // AB LAST_STATE

  calc.getRange("F3:F").setNumberFormat("0.00%");
  calc.getRange("I3:I").setNumberFormat("0.00%");
  calc.getRange("T3:T").setNumberFormat("0.00%"); // Stoch 0..1
  calc.getRange("Y3:Y").setNumberFormat("0.00%");

  const lastRowAll = Math.max(calc.getLastRow(), 2);
  calc.getRange(1, 1, lastRowAll, 28)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  calc.getRange("A1:AB2")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.flush();
}

/**
* ------------------------------------------------------------------
* 5. DASHBOARD ENGINE
* - Signals right after Ticker
* - Formula parse error fixed by simplifying the assembled FILTER() range
* ------------------------------------------------------------------
*/
/**
* ------------------------------------------------------------------
* 5. DASHBOARD ENGINE (OPTIMIZED + SHRINK/GROW SAFE)
* - One-time heavy layout (headers, widths, borders, conditional formats)
* - Fast refresh: only updates controls/timestamp + A4 FILTER/SORT formula
* - Formats only the active rows (based on ticker count)
* - Includes tail cleanup (clears formats below active window when rows shrink)
* ------------------------------------------------------------------
*/
function generateDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = ss.getSheetByName("INPUT");
  if (!input) return;

  const dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD");

  // Determine expected data rows (tickers) so we don’t format 500 every time
  const tickers = getCleanTickers(input);
  const DATA_START_ROW = 4;
  const DATA_ROWS = Math.max(50, Math.min(500, tickers.length + 40)); // cushion for filter spill
  const DATA_END_ROW = DATA_START_ROW + DATA_ROWS - 1;

  // Layout init sentinel
  const SENTINEL = "DASHBOARD_LAYOUT_V1";
  const isInitialized = (dashboard.getRange("A1").getNote() || "").indexOf(SENTINEL) !== -1;

  // ============================================================
  // ONE-TIME LAYOUT BUILD (expensive stuff)
  // ============================================================
  if (!isInitialized) {
    dashboard.clear().clearFormats();

    // ROW 1 — Controls (A1..G1) + checkboxes
    dashboard.getRange("A1")
      .setValue("UPDATE CAL")
      .setBackground("#212121")
      .setFontColor("white")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    dashboard.getRange("B1")
      .insertCheckboxes()
      .setBackground("#212121")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    dashboard.getRange("C1")
      .setValue("UPDATE")
      .setBackground("#212121")
      .setFontColor("white")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    dashboard.getRange("D1")
      .insertCheckboxes()
      .setBackground("#212121")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    dashboard.getRange("E1:G1")
      .merge()
      .setBackground("#000000")
      .setFontColor("#00FF00")
      .setFontWeight("bold")
      .setFontSize(9)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // White border rows 1–3
    dashboard.getRange("A1:AA3")
      .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

    // ROW 2 — Group headers (merged blocks)
    dashboard.getRange("A2:AA2").clearContent();
    const styleGroup = (a1, label, bg) => {
      dashboard.getRange(a1).merge()
        .setValue(label)
        .setBackground(bg)
        .setFontColor("white")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    };
    styleGroup("A2:A2",   "IDENTITY",        "#263238");
    styleGroup("B2:D2",   "SIGNALING",       "#0D47A1");
    styleGroup("E2:G2",   "PRICE / VOLUME",  "#1B5E20");
    styleGroup("H2:J2",   "PERFORMANCE",     "#004D40");
    styleGroup("K2:O2",   "TREND",           "#2E7D32");
    styleGroup("P2:T2",   "MOMENTUM",        "#33691E");
    styleGroup("U2:Y2",   "LEVELS / RISK",   "#B71C1C");
    styleGroup("Z2:AA2",  "NOTES",           "#212121");
    dashboard.getRange("A2:AA2").setWrap(true);

    // ROW 3 — Column headers (Dashboard order)
    const headers = [[
      "Ticker", "SIGNAL", "FUNDAMENTAL", "DECISION", "Price", "Change %", "Vol Trend",
      "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State",
      "SMA 20", "SMA 50", "SMA 200",
      "RSI", "MACD Hist", "Divergence", "ADX (14)", "Stoch %K (14)",
      "Support", "Resistance", "Target (3:1)", "ATR (14)", "Bollinger %B",
      "TECH NOTES", "FUND NOTES"
    ]];
    dashboard.getRange(3, 1, 1, 27)
      .setValues(headers)
      .setBackground("#111111")
      .setFontColor("white")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    // Freeze panes
    dashboard.setFrozenRows(3);
    dashboard.setFrozenColumns(1);

    // Column widths (one time)
    for (let c = 1; c <= 25; c++) dashboard.setColumnWidth(c, 90);
    dashboard.setColumnWidth(26, 420);
    dashboard.setColumnWidth(27, 420);

    // Header alignment
    dashboard.getRange(1, 1, 3, 27).setVerticalAlignment("middle");
    dashboard.getRange("A1:D1").setHorizontalAlignment("center");
    dashboard.getRange("E1:G1").setHorizontalAlignment("center");

    // Borders (one time; bounded but covers likely working area)
    dashboard.getRange(1, 1, 1004, 27)
      .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    dashboard.getRange("A1:AA2")
      .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

    // Conditional formatting (one time, applied to a broad stable window)
    const rules = [];
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setFontColor("#B71C1C")
        .setBold(true)
        .setRanges([
          dashboard.getRange(`F${DATA_START_ROW}:F1000`),
          dashboard.getRange(`I${DATA_START_ROW}:I1000`),
          dashboard.getRange(`Q${DATA_START_ROW}:Q1000`)
        ])
        .build()
    );

    // RSI (P)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$P4>=70")
      .setFontColor("#B71C1C").setBold(true)
      .setRanges([dashboard.getRange(`P${DATA_START_ROW}:P1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$P4<=30")
      .setFontColor("#1B5E20").setBold(true)
      .setRanges([dashboard.getRange(`P${DATA_START_ROW}:P1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND($P4>=50,$P4<70)")
      .setFontColor("#1B5E20")
      .setRanges([dashboard.getRange(`P${DATA_START_ROW}:P1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=AND($P4>30,$P4<50)")
      .setFontColor("#E65100")
      .setRanges([dashboard.getRange(`P${DATA_START_ROW}:P1000`)]).build());

    // ADX (S)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$S4>=25")
      .setFontColor("#1B5E20").setBold(true)
      .setRanges([dashboard.getRange(`S${DATA_START_ROW}:S1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$S4<15")
      .setFontColor("#616161")
      .setRanges([dashboard.getRange(`S${DATA_START_ROW}:S1000`)]).build());

    // Stoch (T)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$T4>=0.8")
      .setFontColor("#B71C1C").setBold(true)
      .setRanges([dashboard.getRange(`T${DATA_START_ROW}:T1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$T4<=0.2")
      .setFontColor("#1B5E20").setBold(true)
      .setRanges([dashboard.getRange(`T${DATA_START_ROW}:T1000`)]).build());

    // %B (Y)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$Y4>=0.8")
      .setFontColor("#B71C1C")
      .setRanges([dashboard.getRange(`Y${DATA_START_ROW}:Y1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$Y4<=0.2")
      .setFontColor("#1B5E20")
      .setRanges([dashboard.getRange(`Y${DATA_START_ROW}:Y1000`)]).build());

    // SIGNAL (B)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($B4,"Breakout|Trend Continuation|RVOL")')
      .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
      .setRanges([dashboard.getRange(`B${DATA_START_ROW}:B1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($B4,"Mean Reversion|Bounce|Oversold|Overbought")')
      .setBackground("#FFF8E1").setFontColor("#E65100").setBold(true)
      .setRanges([dashboard.getRange(`B${DATA_START_ROW}:B1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($B4,"Range|Chop|Hold")')
      .setBackground("#F5F5F5").setFontColor("#616161")
      .setRanges([dashboard.getRange(`B${DATA_START_ROW}:B1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($B4,"Risk-Off|Stop")')
      .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
      .setRanges([dashboard.getRange(`B${DATA_START_ROW}:B1000`)]).build());

    // DECISION (D)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($D4,"Trade|Accumulate|Buy")')
      .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
      .setRanges([dashboard.getRange(`D${DATA_START_ROW}:D1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($D4,"Reduce|Trim|Take Profit")')
      .setBackground("#FFF8E1").setFontColor("#E65100").setBold(true)
      .setRanges([dashboard.getRange(`D${DATA_START_ROW}:D1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($D4,"Hold|Monitor|Wait")')
      .setBackground("#F5F5F5").setFontColor("#616161")
      .setRanges([dashboard.getRange(`D${DATA_START_ROW}:D1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($D4,"Avoid|Stop")')
      .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
      .setRanges([dashboard.getRange(`D${DATA_START_ROW}:D1000`)]).build());

    // FUNDAMENTAL (C)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($C4,"VALUE|GEM|FAIR")')
      .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
      .setRanges([dashboard.getRange(`C${DATA_START_ROW}:C1000`)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($C4,"ZOMBIE|BUBBLE")')
      .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
      .setRanges([dashboard.getRange(`C${DATA_START_ROW}:C1000`)]).build());

    dashboard.setConditionalFormatRules(rules);

    // Sentinel
    dashboard.getRange("A1").setNote(SENTINEL);
  }

  // ============================================================
  // FAST REFRESH PATH (runs every time)
  // ============================================================

  // Clear only content region that changes (keeps layout/CF)
  dashboard.getRange(DATA_START_ROW, 1, 1000, 27).clearContent();

  // Timestamp refresh (E1:G1 already merged in layout)
  dashboard.getRange("E1:G1")
    .setValue(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM dd, yyyy | HH:mm:ss"));

  // ============================================================
  // ROW 4 — Filter formula (always refreshed)
  // ============================================================
  const filterFormula =
    '=IFERROR(' +
      'SORT(' +
        'FILTER({' +
          'CALCULATIONS!$A$3:$A,' +
          'CALCULATIONS!$B$3:$B,' +
          'CALCULATIONS!$D$3:$D,' +
          'CALCULATIONS!$C$3:$C,' +
          'CALCULATIONS!$E$3:$E,' +
          'CALCULATIONS!$F$3:$F,' +
          'CALCULATIONS!$G$3:$G,' +
          'CALCULATIONS!$H$3:$H,' +
          'CALCULATIONS!$I$3:$I,' +
          'CALCULATIONS!$J$3:$J,' +
          'CALCULATIONS!$K$3:$K,' +
          'CALCULATIONS!$L$3:$L,' +
          'CALCULATIONS!$M$3:$M,' +
          'CALCULATIONS!$N$3:$N,' +
          'CALCULATIONS!$O$3:$O,' +
          'CALCULATIONS!$P$3:$P,' +
          'CALCULATIONS!$Q$3:$Q,' +
          'CALCULATIONS!$R$3:$R,' +
          'CALCULATIONS!$S$3:$S,' +
          'CALCULATIONS!$T$3:$T,' +
          'CALCULATIONS!$U$3:$U,' +
          'CALCULATIONS!$V$3:$V,' +
          'CALCULATIONS!$W$3:$W,' +
          'CALCULATIONS!$X$3:$X,' +
          'CALCULATIONS!$Y$3:$Y,' +
          'CALCULATIONS!$Z$3:$Z,' +
          'CALCULATIONS!$AA$3:$AA' +
        '},' +
        'ISNUMBER(MATCH(' +
          'CALCULATIONS!$A$3:$A,' +
          'FILTER(INPUT!$A$3:$A,' +
            'INPUT!$A$3:$A<>"",' +
            '(' +
              'IF(' +
                'OR(' +
                  'INPUT!$B$1="",' +
                  'REGEXMATCH(UPPER(INPUT!$B$1),"(^|,\\s*)ALL(\\s*|,|$)")' +
                '),' +
                'TRUE,' +
                'REGEXMATCH(' +
                  '","&UPPER(TRIM(INPUT!$B$3:$B))&"," ,' +
                  '",\\s*(" & REGEXREPLACE(UPPER(TRIM(INPUT!$B$1)),"\\s*,\\s*","|") & ")\\s*,"' +
                ')' +
              ')' +
            ')' +
            '*' +
            '(' +
              'IF(' +
                'OR(' +
                  'INPUT!$C$1="",' +
                  'REGEXMATCH(UPPER(INPUT!$C$1),"(^|,\\s*)ALL(\\s*|,|$)")' +
                '),' +
                'TRUE,' +
                'REGEXMATCH(' +
                  '","&REGEXREPLACE(UPPER(TRIM(INPUT!$C$3:$C)),"\\s+","")&"," ,' +
                  '",\\s*(" & REGEXREPLACE(REGEXREPLACE(UPPER(TRIM(INPUT!$C$1)),"\\s+",""),"\\s*,\\s*","|") & ")\\s*,"' +
                ')' +
              ')' +
            ')' +
          '),0)' +
        '))' +
        ',6,FALSE' +
      '),' +
    '"No Matches Found")';

  dashboard.getRange("A4").setFormula(filterFormula);

  // ============================================================
  // FAST formatting for active area only
  // ============================================================
  dashboard.getRange(DATA_START_ROW, 1, DATA_ROWS, 25).setWrap(true);
  dashboard.getRange(DATA_START_ROW, 26, DATA_ROWS, 2)
    .setWrap(false)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  dashboard.setRowHeights(DATA_START_ROW, DATA_ROWS, 12);

  // Number formats only for active rows
  dashboard.getRange(`F${DATA_START_ROW}:F${DATA_END_ROW}`).setNumberFormat("0.00%");
  dashboard.getRange(`I${DATA_START_ROW}:I${DATA_END_ROW}`).setNumberFormat("0.00%");
  dashboard.getRange(`T${DATA_START_ROW}:T${DATA_END_ROW}`).setNumberFormat("0.00%");
  dashboard.getRange(`Y${DATA_START_ROW}:Y${DATA_END_ROW}`).setNumberFormat("0.00%");

  // Tail cleanup: clears formats below active window when rows shrink
  const TAIL_START = DATA_END_ROW + 1;
  const TAIL_ROWS = 200; // bounded cleanup for performance
  dashboard.getRange(TAIL_START, 1, TAIL_ROWS, 27).clearFormat();

  SpreadsheetApp.flush();
}


  // ------------------------------------------------------------
  // CHART SHEET setup engine
  // ------------------------------------------------------------

function setupChartSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = ss.getSheetByName("INPUT");
  const calc = ss.getSheetByName("CALCULATIONS");
  if (!input || !calc) throw new Error("Missing INPUT or CALCULATIONS sheet");

  const tickers = getCleanTickers(input);
  let sh = ss.getSheetByName("CHART") || ss.insertSheet("CHART");
  sh.clear().clearFormats();
  forceExpandSheet(sh, 60);

  // ------------------------------------------------------------
  // Column sizing / alignment
  // ------------------------------------------------------------
  sh.setColumnWidth(1, 85);     // A
  sh.setColumnWidth(2, 125);    // B
  sh.setColumnWidth(3, 520);    // C Tech Notes
  sh.setColumnWidth(4, 520);    // D Fund Notes
  sh.setColumnWidth(5, 18);     // E spacer

  sh.getRange("A:A").setHorizontalAlignment("left");
  sh.getRange("B:B").setHorizontalAlignment("left").setWrap(true);

  // Dense top area
  sh.setRowHeights(1, 7, 18);

  // ------------------------------------------------------------
  // Control panel A1:B6
  // ------------------------------------------------------------
  sh.getRange("A1:B6")
    .setBackground("#000000")
    .setFontColor("#FFFF00")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment("middle");

  // Ticker in merged A1:B1 (value lives in A1)
  sh.getRange("A1:B1").merge()
    .setValue(tickers[0] || "")
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setFontColor("#FF80AB")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(tickers.length ? tickers : [""], true)
        .build()
    );

  sh.getRange("A2:A6").setValues([["YEAR"], ["MONTH"], ["DAY"], ["DATE"], ["INTERVAL"]]).setFontWeight("bold");

  const listValidation = (arr) => SpreadsheetApp.newDataValidation().requireValueInList(arr, true).build();

  // B2/B3/B4 start at 0; defaults
  sh.getRange("B2").setDataValidation(listValidation(Array.from({ length: 11 }, (_, i) => i))).setValue(1).setFontColor("#FF80AB");
  sh.getRange("B3").setDataValidation(listValidation(Array.from({ length: 13 }, (_, i) => i))).setValue(0).setFontColor("#FF80AB");
  sh.getRange("B4").setDataValidation(listValidation(Array.from({ length: 32 }, (_, i) => i))).setValue(0).setFontColor("#FF80AB");

  // Date = TODAY() minus (years+months+days)
  sh.getRange("B5").setFormula("=EDATE(TODAY(), -(12*B2+B3)) - B4").setNumberFormat("yyyy-mm-dd").setFontColor("#FF80AB");

  sh.getRange("B6")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"], true).build())
    .setValue("DAILY")
    .setFontWeight("bold")
    .setFontColor("#FF80AB");

  // ------------------------------------------------------------
  // Reasons: C1:C6 and D1:D6
  // CALCULATIONS: Z=TECH NOTES, AA=FUND NOTES
  // ------------------------------------------------------------
  sh.getRange("C1:C6").merge()
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment("top")
    .setHorizontalAlignment("left")
    .setFontSize(10)
    .setFontColor("#FFD54F")
    .setBackground("#0B0B0B")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID)
    .setFormula('=IFERROR(INDEX(CALCULATIONS!$Z$3:$Z, MATCH($A$1, CALCULATIONS!$A$3:$A, 0)), "—")');

  sh.getRange("D1:D6").merge()
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment("top")
    .setHorizontalAlignment("left")
    .setFontSize(10)
    .setFontColor("#FFD54F")
    .setBackground("#0B0B0B")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID)
    .setFormula('=IFERROR(INDEX(CALCULATIONS!$AA$3:$AA, MATCH($A$1, CALCULATIONS!$A$3:$A, 0)), "—")');

  // ------------------------------------------------------------
  // ROW 7: DECISION moved here (A7/B7) + yellow highlight
  // (Do NOT break column mapping: DECISION = CALCULATIONS column C)
  // ------------------------------------------------------------
  const t = "$A$1";
  const IDX = (colLetter, fallback) =>
    `=IFERROR(INDEX(CALCULATIONS!$${colLetter}$3:$${colLetter}, MATCH(${t}, CALCULATIONS!$A$3:$A, 0)), ${fallback})`;

  sh.getRange("A7").setValue("DECISION").setFontWeight("bold");
  sh.getRange("B7").setFormula(IDX("C", '"-"')).setFontWeight("bold");

  sh.getRange("A7:B7")
    .setBackground("#FFEB3B")
    .setFontColor("#111111")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment("middle");

  sh.setRowHeight(7, 18);

  // ------------------------------------------------------------
  // Sidebar (starts row 8)
  // - Add borders
  // - Insert P/E and EPS under [ PERFORMANCE ]
  // - Keep all existing column mappings intact
  // ------------------------------------------------------------
  const startRow = 8;

  // Clear sidebar area (but do not touch chart data region)
  sh.getRange("A8:B200").clearContent();

  const rows = [
    ["SIGNAL",   IDX("B", '"Wait"')],
    ["FUND",     IDX("D", '"-"')],           // FUNDAMENTAL (CALC D)
    // DECISION removed from sidebar because moved to row 7
    ["PRICE",    `=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`],
    ["CHG%",     `=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`],
    ["R:R",      IDX("J", "0")],
    ["", ""],

    ["[ PERFORMANCE ]", ""],
    ["VOL TREND", IDX("G", "0")],
    ["P/E",       `=IFERROR(GOOGLEFINANCE(${t},"pe"), "")`],
    ["EPS",       `=IFERROR(GOOGLEFINANCE(${t},"eps"), "")`],
    ["ATH",       IDX("H", "0")],
    ["ATH %",     IDX("I", "0")],
    ["52W HIGH", `=IFERROR(GOOGLEFINANCE(${t},"high52"), 0)`],
    ["52W LOW",  `=IFERROR(GOOGLEFINANCE(${t},"low52"), 0)`],
    ["", ""],

    ["[ TREND ]", ""],
    ["SMA 20",    IDX("M", "0")],
    ["SMA 50",    IDX("N", "0")],
    ["SMA 200",   IDX("O", "0")],
    ["RSI",       IDX("P", "50")],
    ["MACD",      IDX("Q", "0")],
    ["DIV",       IDX("R", '"-"')],
    ["ADX",       IDX("S", "0")],
    ["STO",       IDX("T", "0")],
    ["", ""],

    ["[ LEVELS ]", ""],
    ["SUPPORT",    IDX("U", "0")],
    ["RESISTANCE", IDX("V", "0")],
    ["TARGET",     IDX("W", "0")],
    ["ATR",        IDX("X", "0")],
    ["%B",         IDX("Y", "0")]
  ];

  sh.getRange(startRow, 1, rows.length, 1).setValues(rows.map(r => [r[0]])).setFontWeight("bold");
  sh.getRange(startRow, 2, rows.length, 1).setFormulas(rows.map(r => [r[1]]));

  // Sidebar borders (requested)
  sh.getRange(startRow, 1, rows.length, 2)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment("middle");

  // Style section headers
  rows.forEach((r, i) => {
    const label = String(r[0] || "");
    if (label.startsWith("[")) {
      sh.getRange(startRow + i, 1, 1, 2)
        .setBackground("#424242")
        .setFontColor("white")
        .setFontWeight("bold");
    }
  });

  sh.setRowHeights(startRow, rows.length, 18);

  // ------------------------------------------------------------
  // Number formats (robust by row numbers in this fixed sidebar)
  // ------------------------------------------------------------
  // Rows are now:
  // 8 SIGNAL
  // 9 FUND
  // 10 PRICE
  // 11 CHG%
  // 12 R:R
  // 13 blank
  // 14 [PERFORMANCE]
  // 15 VOL TREND
  // 16 P/E
  // 17 EPS
  // 18 ATH
  // 19 ATH %
  // 20 blank
  // 21 [TREND]
  // 22 SMA20
  // 23 SMA50
  // 24 SMA200
  // 25 RSI
  // 26 MACD
  // 27 DIV
  // 28 ADX
  // 29 STO
  // 30 blank
  // 31 [LEVELS]
  // 32 SUPPORT
  // 33 RESISTANCE
  // 34 TARGET
  // 35 ATR
  // 36 %B

  sh.getRange("B10").setNumberFormat("#,##0.00"); // PRICE
  sh.getRange("B11").setNumberFormat("0.00%");   // CHG%
  sh.getRange("B12").setNumberFormat("0.00");    // R:R

  sh.getRange("B15").setNumberFormat("0.00");    // VOL TREND
  sh.getRange("B16").setNumberFormat("0.00");    // P/E
  sh.getRange("B17").setNumberFormat("0.00");    // EPS
  sh.getRange("B18").setNumberFormat("#,##0.00");// ATH
  sh.getRange("B19").setNumberFormat("0.00%");   // ATH %

  sh.getRange("B22:B24").setNumberFormat("#,##0.00"); // SMA 20/50/200
  sh.getRange("B25").setNumberFormat("0.00");         // RSI
  sh.getRange("B26").setNumberFormat("0.000");        // MACD
  sh.getRange("B28").setNumberFormat("0.00");         // ADX
  sh.getRange("B29").setNumberFormat("0.00%");        // STO

  sh.getRange("B32:B35").setNumberFormat("#,##0.00"); // SUPPORT/RES/TARGET/ATR
  sh.getRange("B36").setNumberFormat("0.00");        // %B

  SpreadsheetApp.flush();

  updateDynamicChart(); // ensure chart & lines appear
}

/**
* ------------------------------------------------------------------
* updateDynamicChart() — V3_6.1.1 (Live-Stitch + Today's Data)
* ------------------------------------------------------------------
*/
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;

  SpreadsheetApp.flush();

  // 1. Fetch Ticker and Settings
  const ticker = String(sheet.getRange("A1").getValue() || "").trim();
  if (!ticker) return;

  const interval = String(sheet.getRange("B6").getValue() || "DAILY").toUpperCase();
  const isWeekly = interval === "WEEKLY";

  let startDate = sheet.getRange("B5").getValue();
  if (!(startDate instanceof Date)) {
    const now = new Date();
    startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 14);
  }

  // 2. Fetch Sidebar Levels for Chart Lines
  const sidebar = getSidebarValuesByLabels_(sheet, ["PRICE", "SUPPORT", "RESISTANCE", "SUP", "RES"]);
  const livePrice = Number(sidebar["PRICE"]) || 0;
  const supportVal = Number(sidebar["SUPPORT"]) || Number(sidebar["SUP"]) || 0;
  const resistanceVal = Number(sidebar["RESISTANCE"]) || Number(sidebar["RES"]) || 0;

  // 3. Find ticker column in DATA
  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIdx = headers.indexOf(ticker);
  if (colIdx === -1) return;

  // Pull 6 cols: date, open, high, low, close, volume
  const raw = dataSheet.getRange(1, colIdx + 1, dataSheet.getLastRow(), 6).getValues();

  let master = [];
  let vols = [];
  let prices = [];

  // 4. Process Historical Data
  for (let i = 4; i < raw.length; i++) {
    const d = raw[i][0];
    const close = Number(raw[i][4]);
    const vol = Number(raw[i][5]);
    if (!d || !(d instanceof Date) || !isFinite(close) || close < 0.01) continue;
    if (d < startDate) continue;
    if (isWeekly && d.getDay() !== 5) continue;

    // SMA Calculations (Spliced for historical)
    const slice = raw.slice(Math.max(4, i - 200), i + 1).map(r => Number(r[4])).filter(n => isFinite(n) && n > 0);
    const s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
    const s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
    const s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;

    const prevClose = (i > 4) ? Number(raw[i - 1][4]) : close;

    master.push([
      d, close,
      (close >= prevClose) ? vol : null,
      (close < prevClose) ? vol : null,
      s20, s50, s200,
      resistanceVal || null,
      supportVal || null
    ]);

    vols.push(vol);
    prices.push(close);
    if (s20) prices.push(s20);
    if (s50) prices.push(s50);
    if (s200) prices.push(s200);
  }

  // 5. LIVE-STITCH: Add Today's Data point if missing
  const today = new Date();
  const lastDateInMaster = master.length > 0 ? master[master.length - 1][0] : null;

  if (livePrice > 0 && (!lastDateInMaster || lastDateInMaster.toDateString() !== today.toDateString())) {
    const lastHistClose = master.length > 0 ? master[master.length - 1][1] : livePrice;
    
    // For live SMAs, we use the historical slices + current price
    const fullCloses = raw.map(r => Number(r[4])).filter(n => isFinite(n) && n > 0);
    fullCloses.push(livePrice);

    const liveS20 = fullCloses.length >= 20 ? Number((fullCloses.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
    const liveS50 = fullCloses.length >= 50 ? Number((fullCloses.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
    const liveS200 = fullCloses.length >= 200 ? Number((fullCloses.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;

    master.push([
      today, livePrice,
      (livePrice >= lastHistClose) ? (Math.max(...vols) * 0.5) : null, // Proxy Volume for Today
      (livePrice < lastHistClose) ? (Math.max(...vols) * 0.5) : null,
      liveS20, liveS50, liveS200,
      resistanceVal || null,
      supportVal || null
    ]);
    prices.push(livePrice);
  }

  // 6. Write to Data Range (Z3:AH)
  sheet.getRange(3, 26, 2000, 9).clearContent();
  if (master.length === 0) return;

  if (supportVal > 0) prices.push(supportVal);
  if (resistanceVal > 0) prices.push(resistanceVal);
  const cleanPrices = prices.filter(p => typeof p === "number" && isFinite(p) && p > 0);
  const minP = Math.min(...cleanPrices) * 0.98;
  const maxP = Math.max(...cleanPrices) * 1.02;
  const maxVol = Math.max(...vols.filter(v => isFinite(v)), 1);

  sheet.getRange(2, 26, 1, 9).setValues([["Date", "Price", "Bull Vol", "Bear Vol", "SMA 20", "SMA 50", "SMA 200", "Resistance", "Support"]]);
  sheet.getRange(3, 26, master.length, 9).setValues(master);
  sheet.getRange(3, 26, master.length, 1).setNumberFormat("dd/MM/yy");

  // 7. Rebuild COMBO Chart
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(2, 26, master.length + 1, 9))
    .setOption("useFirstRowAsHeaders", true)
    .setOption("series", {
      0: { type: "line", color: "#1A73E8", lineWidth: 3, labelInLegend: "Price" },
      1: { type: "bars", color: "#2E7D32", targetAxisIndex: 1, labelInLegend: "Bull Vol" },
      2: { type: "bars", color: "#C62828", targetAxisIndex: 1, labelInLegend: "Bear Vol" },
      3: { type: "line", color: "#FBBC04", lineWidth: 1.5, labelInLegend: "SMA 20" },
      4: { type: "line", color: "#9C27B0", lineWidth: 1.5, labelInLegend: "SMA 50" },
      5: { type: "line", color: "#FF9800", lineWidth: 2, labelInLegend: "SMA 200" },
      6: { type: "line", color: "#B71C1C", lineDashStyle: [4, 4], labelInLegend: "Resistance" },
      7: { type: "line", color: "#0D47A1", lineDashStyle: [4, 4], labelInLegend: "Support" }
    })
    .setOption("vAxes", {
      0: { viewWindow: { min: minP, max: maxP } },
      1: { viewWindow: { min: 0, max: maxVol * 4 }, format: "short" }
    })
    .setOption("legend", { position: "top" })
    .setPosition(7, 3, 0, 0)
    .setOption("width", 1150)
    .setOption("height", 650)
    .build();

  sheet.insertChart(chart);
}

/**
 * Helper: reads sidebar values by labels (case-insensitive)
 * Scans A8:B200 (your sidebar region)
 */
function getSidebarValuesByLabels_(chartSheet, labels) {
  const want = new Set(labels.map(l => String(l).trim().toUpperCase()));
  const keys = chartSheet.getRange("A8:A200").getValues().flat().map(v => String(v || "").trim().toUpperCase());
  const vals = chartSheet.getRange("B8:B200").getValues().flat();

  const out = {};
  for (let i = 0; i < keys.length; i++) {
    if (want.has(keys[i])) {
      const original = labels.find(l => String(l).trim().toUpperCase() === keys[i]);
      out[original] = vals[i];
    }
  }
  labels.forEach(l => { if (out[l] === undefined) out[l] = 0; });
  return out;
}

function getSidebarLevels_(chartSheet) {
  const labelRange = chartSheet.getRange("A5:A120").getValues().flat();
  const valueRange = chartSheet.getRange("B5:B120").getValues().flat();

  const findValueAny = (labels) => {
    const want = new Set(labels.map(l => String(l).trim().toUpperCase()));
    const idx = labelRange.findIndex(v => want.has(String(v || "").trim().toUpperCase()));
    if (idx === -1) return 0;
    return Number(valueRange[idx]) || 0;
  };

  return {
    support: findValueAny(["SUPPORT", "SUPPORT FLOOR"]),
    resistance: findValueAny(["RESISTANCE", "RESISTANCE CEILING"])
  };
}

/**
* ------------------------------------------------------------------
* 7. AUTOMATED ALERT & MONITOR SYSTEM (LAST_STATE in AB)
* ------------------------------------------------------------------
*/

function startMarketMonitor() {
  stopMarketMonitor();

  ScriptApp.newTrigger('checkSignalsAndSendAlerts')
    .timeBased()
    .everyMinutes(30)
    .create();

  SpreadsheetApp.getUi().alert(
    '🔔 MONITOR ACTIVE',
    'Checking DECISION changes (CALCULATIONS!C) every 30 minutes.\n\n' +
    'You will be emailed only when a DECISION changes, including:\n' +
    '- Trade Long / Accumulate\n' +
    '- Take Profit / Reduce\n' +
    '- Stop-Out / Avoid\n',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
* ------------------------------------------------------------------
* STOP MONITOR (UPDATED TEXT: DECISION monitor)
* ------------------------------------------------------------------
*/
function stopMarketMonitor() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'checkSignalsAndSendAlerts') {
      ScriptApp.deleteTrigger(t);
    }
  });

  SpreadsheetApp.getUi().alert(
    '🔕 MONITOR STOPPED',
    'Automated DECISION checks disabled.\n\n' +
    'No further emails will be sent for DECISION changes (CALCULATIONS!C) until you start the monitor again.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function checkSignalsAndSendAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName("CALCULATIONS");
  if (!calcSheet) return;

  const lastRow = calcSheet.getLastRow();
  if (lastRow < 3) return;

  // A..AB (28 cols)
  const range = calcSheet.getRange(3, 1, lastRow - 2, 28);
  const data = range.getValues();

  const alerts = [];

  data.forEach((r, i) => {
    const ticker = (r[0] || "").toString().trim();     // A
    const decision = (r[2] || "").toString().trim();   // C (DECISION) ✅
    const lastState = (r[27] || "").toString().trim(); // AB (LAST_STATE) ✅

    if (!ticker || !decision || decision === "LOADING") return;
    if (decision === lastState) return;

    // Actionable states: includes SELL/trim/profit + buy/trade + stops/avoid
    const isActionable = /STOP|AVOID|TAKE PROFIT|REDUCE|TRADE LONG|ACCUMULATE/i.test(decision);

    if (isActionable) {
      alerts.push(
        `TICKER: ${ticker}\nNEW DECISION: ${decision}\nPREVIOUS: ${lastState || "Initial Scan"}`
      );
    }

    // Persist the new last notified decision into AB
    calcSheet.getRange(i + 3, 28).setValue(decision);
  });

  if (alerts.length === 0) return;

  const email = Session.getActiveUser().getEmail();
  const subject = `📈 Terminal Alert: ${alerts.length} Decision Change(s)`;
  const body =
    "Institutional Terminal detected DECISION changes (CALCULATIONS!C):\n\n" +
    alerts.join("\n\n") +
    "\n\nView Terminal:\n" + ss.getUrl();

  MailApp.sendEmail(email, subject, body);
}



/**
* ------------------------------------------------------------------
* REFERENCE GUIDE (UPDATED: SELL states + aligned to DECISION/SIGNAL formulas)
* - Keeps your structure; only updates vocabulary tables and explanations.
* ------------------------------------------------------------------
*/
function generateReferenceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "REFERENCE_GUIDE";
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear().clearFormats();

  const rows = [];

  // Title
  rows.push(["INSTITUTIONAL TERMINAL — REFERENCE GUIDE", "", "", ""]);
  rows.push(["Dashboard/Chart vocabulary, column definitions, and action playbook (aligned to current formulas).", "", "", ""]);
  rows.push(["", "", "", ""]);

  // NEW: Position tags policy
  rows.push(["0) POSITION TAGS (INPUT!C) — PRACTICAL USAGE (IMPORTANT)", "", "", ""]);
  rows.push(["RULE", "MEANING", "HOW IT AFFECTS THE ENGINE", "WHAT YOU SHOULD DO"]);
  rows.push([
    "PURCHASED (only behavioral tag)",
    "Indicates you currently hold this ticker (an open position).",
    "Enables sell-side + position-management actions in DECISION (Stop-Out / Take Profit / Reduce / Add in Dip).",
    "When you buy, add PURCHASED to INPUT!C for that ticker. When you fully exit, remove PURCHASED."
  ]);
  rows.push([
    "All other tags (e.g., M7, P0, etc.)",
    "Custom labels for your own grouping and filtering.",
    "NO IMPACT on buy/sell logic. They are ignored by the engine for decisions.",
    "Use them only to filter lists (INPUT filters / watchlists). Do not expect them to alter decisions."
  ]);
  rows.push(["", "", "", ""]);

  // Column definitions
  rows.push(["1) DASHBOARD COLUMN DEFINITIONS (TECHNICAL)", "", "", ""]);
  rows.push(["COLUMN", "WHAT IT IS", "HOW IT IS USED", "USER ACTION"]);
  const cols = [
    ["Ticker", "Symbol (key)", "Join key across DATA/CALCULATIONS/CHART", "Select for chart / review notes."],
    ["SIGNAL", "Technical setup label (rules engine)", "Describes setup type (breakout / trend / mean-rev / risk-off / stop-out)", "Use as setup classification; DECISION is what you act on."],
    ["FUNDAMENTAL", "EPS + P/E risk bucket", "Blocks trades in weak quality/extreme valuation regimes", "Prefer VALUE/FAIR; avoid ZOMBIE/PRICED FOR PERFECTION when momentum is fragile."],
    ["DECISION", "Action label (position-aware)", "Final instruction (trade/accumulate/avoid/stop/trim/profit/add-in-dip)", "Primary action field."],
    ["Price", "Live last price (GOOGLEFINANCE)", "Used for regime tests, distance-to-levels, ATR stretch", "Confirm price vs SMA200 & levels."],
    ["Change %", "Daily % change", "Tape context; not a signal alone", "Do not chase without a setup."],
    ["Vol Trend", "Relative volume proxy (RVOL)", "Conviction filter for breakouts", "Prefer >=1.5x for breakout validation."],
    ["ATH (TRUE)", "All-time high reference", "Context for overhead supply / price discovery", "Avoid chasing into ceilings without RVOL."],
    ["ATH Diff %", "Distance from ATH", "Pullback vs near-ATH classification", "Use with regime + levels."],
    ["R:R Quality", "Reward/Risk ratio proxy", "Trade quality gate", ">=3 excellent; 2–3 acceptable; <2 poor."],
    ["Trend Score", "★ count (Price above SMAs)", "Quick structure strength read", "3★ strongest; <2★ caution."],
    ["Trend State", "Bull/Bear via SMA200", "Defines risk-on vs risk-off playbook", "Below SMA200 = risk-off bias."],
    ["SMA 20", "Short-term mean", "Stretch anchor; mean reversion reference", "Avoid buying when >2x ATR above SMA20."],
    ["SMA 50", "Medium trend line", "Momentum/structure confirmation", "If lost with MACD<0, reduce risk."],
    ["SMA 200", "Long-term regime line", "Primary risk-on/risk-off filter", "Below: avoid trend-chasing."],
    ["RSI", "Momentum oscillator (0–100)", "Overbought/oversold + bias filter", "<30 oversold; >70 overbought; 50 bias."],
    ["MACD Hist", "Impulse (positive/negative)", "Momentum confirmation / deterioration", "Negative impulse with SMA50 loss = reduce."],
    ["Divergence", "Price vs RSI divergence heuristic", "Early reversal warning", "Bull div supports bounce; bear div warns."],
    ["ADX (14)", "Trend strength", "Chop vs trend filter", "<15 range; 15–25 weak; >=25 trend."],
    ["Stoch %K (14)", "Fast oscillator (0–1)", "Timing within regimes", "<0.2 oversold; >0.8 overbought."],
    ["Support", "20-day min low proxy", "Risk line / invalidation reference", "Break below = Stop-Out."],
    ["Resistance", "50-day max high proxy", "Ceiling / profit-taking reference", "Near resistance + overbought = Take Profit."],
    ["Target (3:1)", "Tactical take-profit projection", "Planning exits; not a forecast", "Use for planning only."],
    ["ATR (14)", "Volatility proxy", "Sizing/stops + stretch detection", "Higher ATR = wider stops / smaller size."],
    ["Bollinger %B", "Band position proxy", "Compression/range heuristic", "Low %B + low ADX = chop."],
    ["TECH NOTES", "Narrative (indicator values + rationale)", "Explains what is driving setup and timing", "Read before acting."],
    ["FUND NOTES", "Narrative (fund + signal + action + flags)", "Explains why decision is allowed/blocked", "Respect blockers (risk-off / fragile valuation)."]
  ];
  cols.forEach(r => rows.push(r));

  // SIGNAL vocabulary
  rows.push(["", "", "", ""]);
  rows.push(["2) SIGNAL — FULL VOCABULARY (WHAT IT MEANS + WHAT TO DO)", "", "", ""]);
  rows.push(["SIGNAL VALUE", "TECHNICAL DEFINITION", "WHEN IT TRIGGERS", "EXPECTED USER ACTION"]);
  const signal = [
    ["Stop-Out", "Price < Support", "Breakdown through support floor", "Exit / do not average down. Wait for base."],
    ["Risk-Off (Below SMA200)", "Price < SMA200", "Long-term risk-off regime", "Avoid chasing; only tactical with strict risk."],
    ["Volatility Squeeze (Coiling)", "ATR compressed vs recent lows", "Compression / coiling", "Wait for breakout confirmation (RVOL + levels)."],
    ["Range-Bound (Low ADX)", "ADX < 15", "No trend / chop regime", "Range tactics only; smaller size; tighter targets."],
    ["Breakout (High Volume)", "RVOL high + price near/above resistance", "Breakout attempt with sponsorship", "Only act when DECISION allows (Trade Long)."],
    ["Mean Reversion (Oversold)", "StochK<=0.20 above support", "Oversold timing in structure", "If PURCHASED: Add in Dip (when allowed). If not: Tactical long only if DECISION says Trade Long."],
    ["Trend Continuation", "Above SMA200 with momentum + trend", "Uptrend continuation regime", "Accumulate on pullbacks; avoid chasing stretch."]
  ];
  signal.forEach(r => rows.push(r));

  // FUNDAMENTAL vocabulary
  rows.push(["", "", "", ""]);
  rows.push(["3) FUNDAMENTAL — FULL VOCABULARY (FILTER + RISK)", "", "", ""]);
  rows.push(["FUNDAMENTAL VALUE", "WHAT IT MEANS (IN THIS MODEL)", "RISK PROFILE", "EXPECTED USER ACTION"]);
  const fund = [
    ["VALUE", "EPS positive with supportive P/E", "Lower valuation risk vs others", "Prefer for breakouts/trend setups when tech confirms."],
    ["FAIR", "Neutral valuation bucket (fallback)", "Neutral valuation risk", "Trade only when technical gates pass."],
    ["EXPENSIVE", "Valuation premium (lower margin for error)", "Multiple compression risk", "Be selective; prefer strongest technical setups."],
    ["PRICED FOR PERFECTION", "Very elevated P/E (high expectations)", "Fragile; sharp drawdown risk on misses", "Only best setups; size down; take profits faster."],
    ["ZOMBIE", "EPS <= 0 / weak earnings quality", "High blow-up risk", "Avoid longs; treat as high risk."]
  ];
  fund.forEach(r => rows.push(r));

  // DECISION vocabulary (position-aware)
  rows.push(["", "", "", ""]);
  rows.push(["4) DECISION — FULL VOCABULARY (WHAT TO DO)", "", "", ""]);
  rows.push(["DECISION VALUE", "WHY IT HAPPENS (ENGINE RULE)", "POSITION REQUIREMENT", "EXPECTED USER ACTION"]);
  const decision = [
    ["Stop-Out", "Price < Support (invalidation)", "Purchased or not", "Exit / stand aside."],
    ["Avoid", "Risk-off regime or blocked conditions", "Not required", "No trade; deprioritize."],
    ["Trade Long", "Breakout / setup allowed by gates", "Not purchased", "Enter with stop at Support; plan to Resistance/Target."],
    ["Accumulate", "Trend continuation in acceptable conditions", "Not purchased (or add later manually)", "Scale in on pullbacks; avoid chasing."],
    ["Hold", "No edge or gates not met", "Any", "Do nothing; monitor levels and signals."],
    ["Take Profit", "Target hit OR resistance + overbought", "PURCHASED only", "Trim/sell into strength; do not chase."],
    ["Reduce (Momentum Weak)", "MACD < 0 AND Price < SMA50", "PURCHASED only", "Reduce exposure; tighten risk."],
    ["Reduce (Overextended)", "Stretch >= 2x ATR above SMA20", "PURCHASED only", "Trim; wait for mean reversion."],
    ["Add in Dip", "Oversold timing above support in risk-on regime", "PURCHASED only", "Add small / staged; never add below Support."],
    ["LOADING", "Data not ready", "N/A", "Wait for refresh; do not act."]
  ];
  decision.forEach(r => rows.push(r));

  // Quick playbook
  rows.push(["", "", "", ""]);
  rows.push(["5) QUICK PLAYBOOK (HOW TO USE THE TERMINAL)", "", "", ""]);
  rows.push(["RULE", "WHY", "WHAT TO LOOK FOR", "WHAT TO AVOID"]);
  rows.push(["Position flagging", "Ensures sell logic only triggers for holdings", "INPUT!C contains PURCHASED", "Expecting M7/P0/etc. to change decisions (they will not)."]);
  rows.push(["Trend trades", "Best expectancy in strong regimes", "Above SMA200, ADX>=25, MACD>0, RVOL>=1.5", "Buying in Risk-Off or with ADX<15."]);
  rows.push(["Range trades", "Chop markets are mean-reverting", "ADX<15 and price near Support/Resistance", "Chasing mid-range; poor R:R."]);
  rows.push(["Profit-taking", "Avoid giving back gains", "Take Profit near Resistance/Target", "Adding new longs when stretched/overbought."]);
  rows.push(["Loss avoidance", "Stops define survival", "Stop-Out (Price<Support)", "Averaging down below Support."]);
  rows.push(["R:R gating", "Prevents low-quality trades", "R:R>=2 tactical; >=3 preferred", "R:R<2 unless exceptional setup."]);

  // Write
  sh.getRange(1, 1, rows.length, 4).setValues(rows);

  // Styling (professional)
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 520);
  sh.setColumnWidth(3, 360);
  sh.setColumnWidth(4, 280);
  sh.setRowHeights(1, Math.min(rows.length, 900), 18);
  sh.setFrozenRows(3);

  // Title bars
  sh.getRange("A1:D1").merge()
    .setBackground("#0B5394").setFontColor("white")
    .setFontWeight("bold").setFontSize(13)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.getRange("A2:D2").merge()
    .setBackground("#073763").setFontColor("#FFFF00")
    .setFontWeight("bold").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Section headers
  for (let r = 1; r <= rows.length; r++) {
    const v = String(sh.getRange(r, 1).getValue() || "");
    if (/^\d\)|^0\)/.test(v)) {
      sh.getRange(r, 1, 1, 4).merge()
        .setBackground("#212121").setFontColor("white")
        .setFontWeight("bold").setFontSize(10)
        .setHorizontalAlignment("left");
    }
  }

  // Table header rows
  for (let r = 1; r <= rows.length; r++) {
    const a = String(sh.getRange(r, 1).getValue() || "").trim();
    if (["RULE", "COLUMN", "SIGNAL VALUE", "FUNDAMENTAL VALUE", "DECISION VALUE"].includes(a)) {
      sh.getRange(r, 1, 1, 4)
        .setBackground("#F3F3F3")
        .setFontWeight("bold")
        .setFontColor("#111111")
        .setHorizontalAlignment("center");
    }
  }

  sh.getRange(1, 1, rows.length, 4).setWrap(true).setVerticalAlignment("top");
  sh.getRange(1, 1, rows.length, 4)
    .setBorder(true, true, true, true, true, true, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID);

  const band = sh.getRange(4, 1, Math.max(1, rows.length - 3), 4).applyRowBanding();
  band.setHeaderRowColor("#FFFFFF");
  band.setFirstRowColor("#FFFFFF");
  band.setSecondRowColor("#FAFAFA");

  ss.toast("REFERENCE_GUIDE updated: PURCHASED is the only behavioral tag; others are filter-only.", "✅ DONE", 3);
}