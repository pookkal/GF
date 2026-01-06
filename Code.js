/**
* ==============================================================================
* BASELINE LABEL: STABLE_MASTER_ALL_CLEAN_v4.5.1_GOLDEN
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
    .addItem('🚀 1- FETCH DATA', 'FlushDataSheetAndBuild')
    .addItem('🚀 2. REBUILD ALL SHEETS', 'FlushAllSheetsAndBuild')
    .addSeparator()
    .addItem('3. Build Calculations', 'generateCalculationsSheet')
    .addItem('4. Refresh Dashboard ', 'generateDashboardSheet')
    .addItem('4. Refresh Mobile Dashbaord ', 'setupFormulaBasedReport')
    .addItem('5. Setup Chart', 'setupChartSheet')
    .addSeparator()
    .addItem('🤖 Generate  Narratives', 'runMasterAnalysis')
    .addSeparator()
    .addItem('📖 Open Reference Guide', 'generateReferenceSheet')
    .addSeparator()
    .addItem('🔔 Start Market Monitor', 'startMarketMonitor')
    .addItem('🔕 Stop Monitor', 'stopMarketMonitor')
    .addItem('📩 Test Alert Now', 'checkSignalsAndSendAlerts')
    .addToUi();
}

// ------------------------------------------------------------
// UPDATED onEdit(e) — watches the changes to update shets
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
 
  if (sheet.getName() === "REPORT" && (a1 === "A1" )) {
     try {
      ss.toast("Refreshing Mbile Dashboard...", "⚙️ TERMINAL", 3);
      //generateMobileReport();
      SpreadsheetApp.flush();
    } catch (err) {
      ss.toast('REPORT SHEET A1 select onEdit error: ' + err);
    }
  }

  if (sheet.getName() === "CHART") {
    const watchList = ["A1", "B2", "B3", "B4", "B6"];
   
    // This triggers if B1-B6 are edited OR any cell in Row 1 (Cols 1-4)
    if (watchList.indexOf(a1) !== -1 || (range.getRow() === 1 && range.getColumn() <= 4)) {
      try {
        ss.toast("🔄 Refreshing Chart & Analysis...", "WORKING", 2);
        if (typeof updateDynamicChart === "function") 
          updateDynamicChart();
      } catch (err) {
        ss.toast("Refresh error: " + err.toString(), "⚠️ FAIL", 6);
      }
      return; // Exit after processing CHART
    }
  }
}

function onEditInstall(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  // Trigger ONLY when CHART!A1 is edited
  //if (sheet.getName() === "CHART" && range.getA1Notation() === "A1") {
    //runMasterAnalysis();
  //}
}
/**
* ------------------------------------------------------------------
* 1. CORE AUTOMATION
* ------------------------------------------------------------------
*/
function FlushAllSheetsAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["CALCULATIONS","DASHBOARD",  "CHART", "REPORT"];
  const ui = SpreadsheetApp.getUi();

  if (ui.alert('🚨 Full Rebuild', 'Rebuild the sheets?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Integrating Indicators..."), "Status");
  generateCalculationsSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>2/4:</b> Building Dashboard..."), "Status");
  generateDashboardSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>3/4:</b> Constructing Report..."), "Status");
  setupFormulaBasedReport();

   ui.showModelessDialog(HtmlService.createHtmlOutput("<b>4/4:</b> Constructing Chart..."), "Status");
  setupChartSheet();

  ui.alert('✅ Rebuild Complete', 'Terminal Online. Data links restored.', ui.ButtonSet.OK);
}

function FlushDataSheetAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["DATA"];
  const ui = SpreadsheetApp.getUi();

  if (ui.alert('🚨 Full Rebuild', 'Rebuild Data?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Syncing Global Data..."), "Status");
  generateDataSheet();
  SpreadsheetApp.flush();

  ui.alert('✅ Rebuild Complete', 'Data  rerestored.', ui.ButtonSet.OK);
}


/**
* ------------------------------------------------------------------
* 2. CUSTOM MATH FUNCTIONS (RSI, MACD, ADX, STOCH)
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
        LET(
          tag${SEP}UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}"" ))${SEP}
          purchased${SEP}REGEXMATCH(tag${SEP}"(^|,|\\s)PURCHASED(\\s|,|$)")${SEP}

          IFS(
            AND(IFERROR(VALUE($E${row})${SEP}0)>0${SEP}IFERROR(VALUE($U${row})${SEP}0)>0${SEP}IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($U${row})${SEP}0))${SEP}
              "Stop-Out"${SEP}

            AND(purchased${SEP}
                IFERROR(VALUE($E${row})${SEP}0)>0${SEP}
                IFERROR(VALUE($W${row})${SEP}0)>0${SEP}
                IFERROR(VALUE($E${row})${SEP}0)>=IFERROR(VALUE($W${row})${SEP}0)
            )${SEP}"Take Profit"${SEP}

            AND(purchased${SEP}
                IFERROR(VALUE($V${row})${SEP}0)>0${SEP}
                IFERROR(VALUE($E${row})${SEP}0)>=IFERROR(VALUE($V${row})${SEP}0)*0.995${SEP}
                OR(IFERROR(VALUE($P${row})${SEP}50)>=70${SEP}IFERROR(VALUE($T${row})${SEP}0.5)>=0.8)
            )${SEP}"Take Profit"${SEP}

            AND(purchased${SEP}
                IFERROR(VALUE($Q${row})${SEP}0)<0${SEP}
                IFERROR(VALUE($N${row})${SEP}0)>0${SEP}
                IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($N${row})${SEP}0)
            )${SEP}"Reduce (Momentum Weak)"${SEP}

            AND(purchased${SEP}
                IFERROR(VALUE($X${row})${SEP}0)>0${SEP}
                IFERROR(VALUE($M${row})${SEP}0)>0${SEP}
                (IFERROR(VALUE($E${row})${SEP}0)-IFERROR(VALUE($M${row})${SEP}0))/IFERROR(VALUE($X${row})${SEP}1) >= 2
            )${SEP}"Reduce (Overextended)"${SEP}

            AND(purchased${SEP}IFERROR(VALUE($E${row})${SEP}0) < IFERROR(VALUE($O${row})${SEP}0))${SEP}
              "Risk-Off (Below SMA200)"${SEP}

            AND(NOT(purchased)${SEP}IFERROR(VALUE($E${row})${SEP}0) < IFERROR(VALUE($O${row})${SEP}0))${SEP}
              "Avoid"${SEP}

            AND(purchased${SEP}
                IFERROR(VALUE($U${row})${SEP}0)>0${SEP}
                IFERROR(VALUE($E${row})${SEP}0)>IFERROR(VALUE($U${row})${SEP}0)${SEP}
                IFERROR(VALUE($T${row})${SEP}0.5)<=0.2
            )${SEP}"Add in Dip"${SEP}

            AND($B${row}="Breakout (High Volume)"${SEP}OR($D${row}="VALUE"${SEP}$D${row}="FAIR"))${SEP}"Trade Long"${SEP}
            AND($B${row}="Breakout (High Volume)"${SEP}OR($D${row}="EXPENSIVE"${SEP}$D${row}="PRICED FOR PERFECTION"))${SEP}"Hold"${SEP}
            AND($B${row}="Trend Continuation"${SEP}$D${row}="VALUE")${SEP}"Accumulate"${SEP}
            $B${row}="Trend Continuation"${SEP}"Hold"${SEP}
            TRUE${SEP}"Hold"
          )
        )
      )`;

    // FUNDAMENTAL (D) — reads cached PE/EPS from DATA row 3 (fast)
    const fFund =
  `=IFERROR(` +
    `LET(` +
      `peRaw${SEP}${peCell}${SEP}` +
      `epsRaw${SEP}${epsCell}${SEP}` +
      `athDiffRaw${SEP}$I${row}${SEP}` +  // ATH Diff % column I

      `pe${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(peRaw)${SEP}"[^0-9\\.\\-]"${SEP}"" ))${SEP}"" )${SEP}` +
      `eps${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(epsRaw)${SEP}"[^0-9\\.\\-]"${SEP}"" ))${SEP}"" )${SEP}` +

      // Column I is % (e.g., -14.95%). Normalize to decimal (-0.1495).
      `athDiff${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(athDiffRaw)${SEP}"[^0-9\\.\\-]"${SEP}"" ))/100${SEP}"" )${SEP}` +

      `IFS(` +
        `OR(pe=""${SEP}eps="")${SEP}"FAIR"${SEP}` +
        `eps<=0${SEP}"ZOMBIE"${SEP}` +

        // Priced for perfection = very high PE AND near ATH (within ~8%)
        `AND(pe>=60${SEP}athDiff<>""${SEP}athDiff>=-0.08)${SEP}"PRICED FOR PERFECTION"${SEP}` +

        `pe>=35${SEP}"EXPENSIVE"${SEP}` +
        `AND(pe>0${SEP}pe<=25${SEP}eps>=0.5)${SEP}"VALUE"${SEP}` +
        `AND(pe>25${SEP}pe<35${SEP}eps>=0.5)${SEP}"FAIR"${SEP}` +
        `TRUE${SEP}"FAIR"` +
      `)` +`)` +`${SEP}"FAIR")`;


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

   /**
    * Why this is the correct "Industry" fix:FeatureAverage of Extremes (Trial & Error)Percentile (Industry Standard)Outlier HandlingStill weighted by the outlier (e.g., 362.70).Ignores the outlier entirely.Zone AccuracyRepresents a single point.Represents the Value Area where most trading occurred.StabilityJumps around when a new high/low enters the window.Remains stable as long as the distribution of price is consistent.
    */
    const fRes = `=ROUND(IFERROR(LET(win${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!$${highCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!$${highCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MAX(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.85))${SEP}out)${SEP}0)${SEP}2)`;

    const fSup = `=ROUND(IFERROR(LET(win${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!$${lowCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!$${lowCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MIN(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.15))${SEP}out)${SEP}0)${SEP}2)`;
    
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

    // AA FUND NOTES — Plain English narrative explaining Signal, Fundamental, and Decision
    const fFundNotes =
      `=IF($A${row}=""${SEP}""${SEP}
      "FUNDAMENTAL ANALYSIS: "&IFS(
        $D${row}="VALUE"${SEP}"This stock is attractively priced with strong earnings and reasonable valuation (PE ≤ 25). The fundamentals provide a supportive tailwind for any position."
        ${SEP}$D${row}="FAIR"${SEP}"This stock has decent fundamentals but nothing exceptional. The valuation is neither cheap nor expensive, so fundamentals are neutral to the trade."
        ${SEP}$D${row}="EXPENSIVE"${SEP}"This stock is trading at a premium valuation (PE 35-59). While not prohibitive, there's less margin for error and fundamentals create a headwind."
        ${SEP}$D${row}="PRICED FOR PERFECTION"${SEP}"This stock has extremely high expectations built into the price (PE ≥ 60). Any disappointment could cause significant downside. Fundamentals are fragile."
        ${SEP}$D${row}="ZOMBIE"${SEP}"This company is losing money or has very weak earnings quality (EPS ≤ 0). High risk of permanent capital loss. Fundamentals are severely negative."
        ${SEP}TRUE${SEP}"Fundamental analysis is inconclusive due to missing data."
      )&CHAR(10)&CHAR(10)&

      "TECHNICAL SIGNAL: "&$B${row}&CHAR(10)&
      "Why this signal: "&IFS(
        $B${row}="Stop-Out"${SEP}"Price has broken below the key support level, invalidating the bullish thesis. This is a defensive exit signal to preserve capital."
        ${SEP}$B${row}="Breakout (High Volume)"${SEP}"Price is breaking above resistance with strong volume confirmation, suggesting institutional participation and potential for continued upside momentum."
        ${SEP}$B${row}="Trend Continuation"${SEP}"Price is above the 200-day moving average with positive momentum indicators, suggesting the existing uptrend has room to continue higher."
        ${SEP}$B${row}="Mean Reversion (Oversold)"${SEP}"Price is oversold on short-term indicators but holding above key support, creating a potential bounce opportunity back toward fair value."
        ${SEP}$B${row}="Volatility Squeeze (Coiling)"${SEP}"Price volatility has compressed to extremely low levels, often preceding significant directional moves. Waiting for the breakout direction."
        ${SEP}$B${row}="Range-Bound (Low ADX)"${SEP}"Trend strength is weak with price moving sideways. This environment favors range trading rather than directional bets."
        ${SEP}TRUE${SEP}"Market conditions don't clearly favor any specific technical setup. Monitoring for clearer signals."
      )&CHAR(10)&CHAR(10)&

      "INVESTMENT DECISION: "&$C${row}&CHAR(10)&
      "Why this decision: "&IFS(
        $C${row}="Stop-Out"${SEP}"Price has broken below support level. Exiting to prevent further losses and preserve capital."
        ${SEP}AND($C${row}="Take Profit",IFERROR(VALUE(INDEX(CALCULATIONS!E:E,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),0)>=IFERROR(VALUE(INDEX(CALCULATIONS!W:W,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),0),IFERROR(VALUE(INDEX(CALCULATIONS!W:W,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),0)>0)${SEP}"Price has reached target level. Taking profits while conditions are favorable."
        ${SEP}$C${row}="Take Profit"${SEP}"Price is overbought near resistance. Taking profits to avoid pullback risk from elevated RSI and Stochastic levels."
        ${SEP}$C${row}="Reduce (Momentum Weak)"${SEP}"MACD histogram has turned negative and price is below SMA50. Reducing position size to manage deteriorating momentum."
        ${SEP}$C${row}="Reduce (Overextended)"${SEP}"Price has extended too far above SMA20 relative to average volatility. Taking partial profits to reduce pullback risk."
        ${SEP}$C${row}="Risk-Off (Below SMA200)"${SEP}"Price is below the 200-day moving average indicating risk-off conditions. Maintaining defensive posture until trend improves."
        ${SEP}$C${row}="Avoid"${SEP}"Price is below SMA200 indicating risk-off conditions. Avoiding new positions until trend improves above key moving average."
        ${SEP}$C${row}="Add in Dip"${SEP}"Price is above support with Stochastic showing oversold conditions. Adding to position on this dip opportunity."
        ${SEP}$C${row}="Trade Long"${SEP}"Breakout signal confirmed with strong fundamentals. Initiating long position with favorable risk/reward setup."
        ${SEP}$C${row}="Accumulate"${SEP}"Trend continuation signal with VALUE fundamentals. Adding to existing position as uptrend remains intact above SMA200."
        ${SEP}$C${row}="Hold"${SEP}"Current market conditions suggest maintaining existing position. Monitoring for clearer directional signals before making changes."
        ${SEP}TRUE${SEP}"Decision framework suggests maintaining current stance until market conditions become clearer."
      )&

      IF(
        AND(
          OR($B${row}="Breakout (High Volume)"${SEP}$B${row}="Trend Continuation")${SEP}
          OR($D${row}="ZOMBIE"${SEP}$D${row}="PRICED FOR PERFECTION"${SEP}$D${row}="EXPENSIVE")
        )
        ${SEP}CHAR(10)&CHAR(10)&"⚠️ RISK WARNING: Strong technical momentum is conflicting with weak or fragile fundamentals. This creates higher risk of sharp reversals if momentum fails."
        ${SEP}IF(
          AND(
            OR($B${row}="Mean Reversion (Oversold)"${SEP}$B${row}="Stop-Out")${SEP}
            $D${row}="VALUE"
          )
          ${SEP}CHAR(10)&CHAR(10)&"💡 OPPORTUNITY NOTE: Attractive valuation is present, but technical structure needs to improve before becoming more aggressive."
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

function generateDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = ss.getSheetByName("INPUT");
  if (!input) return;

  const dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD");

  const tickers = getCleanTickers(input);
  const DATA_START_ROW = 4;
  const DATA_ROWS = Math.max(50, Math.min(500, tickers.length + 40)); // spill cushion
  const SENTINEL = "DASHBOARD_LAYOUT_V1_BLOOMBERG";

  const isInitialized = (dashboard.getRange("A1").getNote() || "").indexOf(SENTINEL) !== -1;

  // ==========================
  // ONE-TIME LAYOUT (NO ROW3+ formatting during refresh)
  // ==========================
  if (!isInitialized) {
    dashboard.clear().clearFormats();

    // --- Row 1 controls ---
    dashboard.getRange("A1")
      .setValue("UPDATE CAL")
      .setBackground("#212121").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    dashboard.getRange("B1")
      .insertCheckboxes()
      .setBackground("#212121")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    dashboard.getRange("C1")
      .setValue("UPDATE")
      .setBackground("#212121").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    dashboard.getRange("D1")
      .insertCheckboxes()
      .setBackground("#212121")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    dashboard.getRange("E1:G1").merge()
      .setBackground("#000000").setFontColor("#00FF00").setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    // --- Row 2 group headers ---
    const styleGroup = (a1, label, bg) => {
      dashboard.getRange(a1).merge()
        .setValue(label)
        .setBackground(bg).setFontColor("white").setFontWeight("bold")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
    };

    dashboard.getRange("A2:AA2").clearContent();
    styleGroup("A2:A2",   "IDENTITY",        "#263238");
    styleGroup("B2:D2",   "SIGNALING",       "#0D47A1");
    styleGroup("E2:G2",   "PRICE / VOLUME",  "#1B5E20");
    styleGroup("H2:J2",   "PERFORMANCE",     "#004D40");
    styleGroup("K2:O2",   "TREND",           "#2E7D32");
    styleGroup("P2:T2",   "MOMENTUM",        "#33691E");
    styleGroup("U2:Y2",   "LEVELS / RISK",   "#B71C1C");
    styleGroup("Z2:AA2",  "NOTES",           "#212121");
    dashboard.getRange("A2:AA2").setWrap(true);

    // --- Row 3 column headers ---
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
      .setBackground("#111111").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setWrap(true);

    // Freeze panes
    dashboard.setFrozenRows(3);
    dashboard.setFrozenColumns(1);

    // White border for top header rows (1..3)
    dashboard.getRange("A1:AA3")
      .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

    // Sentinel note
    dashboard.getRange("A1").setNote(SENTINEL);
  }
  // ==========================
  // FAST REFRESH (DATA ONLY)
  // ==========================
  dashboard.getRange(DATA_START_ROW, 1, 1000, 27).clearContent();

  // Timestamp refresh
  dashboard.getRange("E1:G1")
    .setValue(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM dd, yyyy | HH:mm:ss"));

  // Filter formula (always re-written)
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

  SpreadsheetApp.flush();

  // Apply Bloomberg formatting + heatmap ONCE
  applyDashboardBloombergFormatting_(dashboard, DATA_START_ROW);
  applyDashboardGroupMapAndColors_(dashboard);
}

function applyDashboardBloombergFormatting_(sh, DATA_START_ROW) {
  if (!sh) return;

  // ---------------------------
  // Theme (strict 3 colors)
  // ---------------------------
  const C_WHITE = "#FFFFFF";
  const C_GREEN = "#C6EFCE";
  const C_RED   = "#FFC7CE";
  const C_GREY  = "#E7E6E6";

  const HEADER_DARK = "#1F1F1F"; // header grey
  const HEADER_MID  = "#E7E6E6"; // header grey (light)

  // Data columns in DASHBOARD (A..AA = 27); notes are Z(26) and AA(27)
  const TOTAL_COLS = 27;

  // Column index map (Dashboard layout you use)
  // A Ticker
  // B SIGNAL
  // C FUNDAMENTAL
  // D DECISION
  // E Price
  // F Change %
  // G Vol Trend
  // H ATH
  // I ATH Diff %
  // J R:R
  // K Trend Score
  // L Trend State
  // M SMA20
  // N SMA50
  // O SMA200
  // P RSI
  // Q MACD Hist
  // R Divergence
  // S ADX
  // T StochK
  // U Support
  // V Resistance
  // W Target
  // X ATR
  // Y %B
  // Z TECH NOTES (hidden)
  // AA FUND NOTES (hidden)

  // ---------------------------
  // Helpers
  // ---------------------------
  const clamp = (n, lo, hi) => Math.max(lo, Math.min(hi, n));

  function findLastDataRow_() {
    // We determine the actual spill length by scanning column A from DATA_START_ROW down.
    // This avoids "retained colors" below when list shrinks.
    const maxScan = 2000; // bounded for performance
    const lastRowPossible = Math.max(sh.getLastRow(), DATA_START_ROW);
    const scanRows = clamp(lastRowPossible - DATA_START_ROW + 1, 1, maxScan);

    const vals = sh.getRange(DATA_START_ROW, 1, scanRows, 1).getDisplayValues().flat();
    let lastNonEmptyOffset = -1;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i] || "").trim() !== "") lastNonEmptyOffset = i;
    }
    if (lastNonEmptyOffset === -1) return DATA_START_ROW; // no data
    return DATA_START_ROW + lastNonEmptyOffset;
  }

  function safeHideNotes_() {
    // Z = 26, AA = 27
    try {
      sh.hideColumns(26);
      sh.hideColumns(27);
    } catch (_) {
      // ignore if already hidden / protected
    }
  }

  function clearTailFormats_(lastDataRow) {
    const maxRows = sh.getMaxRows();
    const tailStart = lastDataRow + 1;
    if (tailStart <= maxRows) {
      const tailRows = maxRows - tailStart + 1;
      if (tailRows > 0) {
        sh.getRange(tailStart, 1, tailRows, TOTAL_COLS).clearFormat().clearContent();
        // NOTE: We clearContent too, to avoid ghosts in formulas spill collisions.
        // If you do NOT want tail content cleared, remove .clearContent()
      }
    }
  }

  // ---------------------------
  // Compute active window
  // ---------------------------
  safeHideNotes_();

  const lastDataRow = findLastDataRow_();
  const numRows = Math.max(1, lastDataRow - DATA_START_ROW + 1);

  const dataRange = sh.getRange(DATA_START_ROW, 1, numRows, 25); // A..Y only (visible)
  const borderRange = sh.getRange(DATA_START_ROW, 1, numRows, 25);

  // ---------------------------
  // Header styling (rows 1–3)
  // ---------------------------
  // Row 1: control bar (A1:AA1)
  sh.getRange(1, 1, 1, TOTAL_COLS)
    .setBackground(HEADER_DARK)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");

  // Keep your existing checkbox cells readable (B1, D1) — do not override content/checkbox
  sh.getRange("A1:D1").setHorizontalAlignment("center");
  sh.getRange("E1:G1").setHorizontalAlignment("center");

  // Row 2: group headers bar (A2:AA2)
  sh.getRange(2, 1, 1, TOTAL_COLS)
    .setBackground(HEADER_MID)
    .setFontColor("#000000")
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");

  // Row 3: column headers (A3:AA3)
  sh.getRange(3, 1, 1, TOTAL_COLS)
    .setBackground(HEADER_DARK)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrap(true);

  // ---------------------------
  // Global layout (Bloomberg dense)
  // ---------------------------
  // Column width ~ 10 chars ≈ 85 px
  for (let c = 1; c <= 25; c++) sh.setColumnWidth(c, 85);
  // Notes are hidden but keep sane widths if unhidden later
  sh.setColumnWidth(26, 420);
  sh.setColumnWidth(27, 420);

  // Row heights:
  // - headers: compact
  sh.setRowHeight(1, 22);
  sh.setRowHeight(2, 18);
  sh.setRowHeight(3, 22);

  // - data rows: ~ 3 lines (18 * 3 = 54)
  sh.setRowHeights(DATA_START_ROW, numRows, 54);

  // Alignment / wrap
  dataRange
    .setBackground(C_WHITE)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setWrap(true);

  // Clip long text in visible data area (A..Y) to keep terminal dense
  dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Borders: black for data rows
  borderRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Keep header borders clean/white (Bloomberg top bar look)
  sh.getRange(1, 1, 3, TOTAL_COLS)
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // ---------------------------
  // Number formats (A..Y)
  // ---------------------------
  // Price
  sh.getRange(DATA_START_ROW, 5, numRows, 1).setNumberFormat("#,##0.00");  // E
  // Change%
  sh.getRange(DATA_START_ROW, 6, numRows, 1).setNumberFormat("0.00%");     // F
  // RVOL
  sh.getRange(DATA_START_ROW, 7, numRows, 1).setNumberFormat("0.00");      // G
  // ATH
  sh.getRange(DATA_START_ROW, 8, numRows, 1).setNumberFormat("#,##0.00");  // H
  // ATH Diff%
  sh.getRange(DATA_START_ROW, 9, numRows, 1).setNumberFormat("0.00%");     // I
  // R:R
  sh.getRange(DATA_START_ROW,10, numRows, 1).setNumberFormat("0.00");      // J
  // SMAs
  sh.getRange(DATA_START_ROW,13, numRows, 3).setNumberFormat("#,##0.00");  // M:N:O
  // RSI, ADX
  sh.getRange(DATA_START_ROW,16, numRows, 1).setNumberFormat("0.0");       // P
  sh.getRange(DATA_START_ROW,19, numRows, 1).setNumberFormat("0.0");       // S
  // MACD
  sh.getRange(DATA_START_ROW,17, numRows, 1).setNumberFormat("0.000");     // Q
  // Stoch (0..1)
  sh.getRange(DATA_START_ROW,20, numRows, 1).setNumberFormat("0.00%");     // T
  // Support/Res/Target/ATR
  sh.getRange(DATA_START_ROW,21, numRows, 4).setNumberFormat("#,##0.00");  // U..X
  // %B
  sh.getRange(DATA_START_ROW,25, numRows, 1).setNumberFormat("0.00");      // Y

  // ---------------------------
  // Clear any previous conditional formatting then apply new rules
  // ---------------------------
  const rules = [];
  const r0 = DATA_START_ROW;

  const rngCol = (col) => sh.getRange(r0, col, numRows, 1);

  const add = (formula, color, col) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(color)
        .setRanges([rngCol(col)])
        .build()
    );
  };

  // ---- SIGNAL (B) — replicate your hierarchy using existing computed SIGNAL text ----
  // Green = breakout / trend continuation / mean reversion
  // Red   = stop-out / risk-off
  // Grey  = squeeze / range / hold
  add(`=REGEXMATCH($B${r0},"Breakout|Trend Continuation|Mean Reversion")`, C_GREEN, 2);
  add(`=REGEXMATCH($B${r0},"Stop-Out|Risk-Off")`,                           C_RED,   2);
  add(`=REGEXMATCH($B${r0},"Volatility Squeeze|Range-Bound|Hold")`,         C_GREY,  2);

  // ---- FUNDAMENTAL (C) ----
  // Green = VALUE
  // Grey  = FAIR
  // Red   = EXPENSIVE / PRICED FOR PERFECTION / ZOMBIE
  add(`=$C${r0}="VALUE"`,                                                   C_GREEN, 3);
  add(`=$C${r0}="FAIR"`,                                                    C_GREY,  3);
  add(`=REGEXMATCH($C${r0},"EXPENSIVE|PRICED FOR PERFECTION|ZOMBIE")`,      C_RED,   3);

  // ---- DECISION (D) ----
  // Green = Trade Long / Accumulate / Add in Dip
  // Red   = Stop-Out / Avoid / Reduce / Take Profit
  // Grey  = Hold / Monitor / LOADING
  add(`=REGEXMATCH($D${r0},"Trade Long|Accumulate|Add in Dip")`,            C_GREEN, 4);
  add(`=REGEXMATCH($D${r0},"Stop-Out|Avoid|Reduce|Take Profit")`,          C_RED,   4);
  add(`=REGEXMATCH($D${r0},"Hold|Monitor|LOADING")`,                       C_GREY,  4);

  // ---- PRICE (E) and Change% (F) ----
  add(`=$F${r0}>0`,                                                         C_GREEN, 5);
  add(`=$F${r0}<0`,                                                         C_RED,   5);
  add(`=OR($F${r0}=0,$F${r0}="")`,                                          C_GREY,  5);

  add(`=$F${r0}>0`,                                                         C_GREEN, 6);
  add(`=$F${r0}<0`,                                                         C_RED,   6);
  add(`=OR($F${r0}=0,$F${r0}="")`,                                          C_GREY,  6);

  // ---- Vol Trend RVOL (G) ----
  add(`=$G${r0}>=1.5`,                                                      C_GREEN, 7);
  add(`=$G${r0}<=0.85`,                                                     C_RED,   7);
  add(`=AND($G${r0}>0.85,$G${r0}<1.5)`,                                     C_GREY,  7);

  // ---- ATH (H) / ATH Diff % (I) ----
  // Near ATH: green; deep below ATH: red; else grey
  add(`=AND($H${r0}>0,$E${r0}>=$H${r0}*0.995)`,                             C_GREEN, 8);
  add(`=AND($H${r0}>0,$E${r0}<=$H${r0}*0.80)`,                              C_RED,   8);
  add(`=AND($H${r0}>0,$E${r0}>$H${r0}*0.80,$E${r0}<$H${r0}*0.995)`,         C_GREY,  8);

  add(`=$I${r0}>=-0.05`,                                                    C_GREEN, 9);
  add(`=$I${r0}<=-0.20`,                                                    C_RED,   9);
  add(`=AND($I${r0}>-0.20,$I${r0}<-0.05)`,                                  C_GREY,  9);

  // ---- R:R (J) ----
  add(`=$J${r0}>=3`,                                                        C_GREEN,10);
  add(`=$J${r0}<1.5`,                                                       C_RED,  10);
  add(`=AND($J${r0}>=1.5,$J${r0}<3)`,                                       C_GREY, 10);

  // ---- Trend Score (K) — star count ----
  add(`=LEN($K${r0})>=3`,                                                   C_GREEN,11);
  add(`=LEN($K${r0})<=1`,                                                   C_RED,  11);
  add(`=LEN($K${r0})=2`,                                                    C_GREY, 11);

  // ---- Trend State (L) ----
  add(`=$L${r0}="BULL"`,                                                    C_GREEN,12);
  add(`=$L${r0}="BEAR"`,                                                    C_RED,  12);
  add(`=AND($L${r0}<>"BULL",$L${r0}<>"BEAR")`,                              C_GREY, 12);

  // ---- SMA20/50/200 (M/N/O) vs Price ----
  add(`=AND($M${r0}>0,$E${r0}>=$M${r0})`,                                   C_GREEN,13);
  add(`=AND($M${r0}>0,$E${r0}<$M${r0})`,                                    C_RED,  13);

  add(`=AND($N${r0}>0,$E${r0}>=$N${r0})`,                                   C_GREEN,14);
  add(`=AND($N${r0}>0,$E${r0}<$N${r0})`,                                    C_RED,  14);

  add(`=AND($O${r0}>0,$E${r0}>=$O${r0})`,                                   C_GREEN,15);
  add(`=AND($O${r0}>0,$E${r0}<$O${r0})`,                                    C_RED,  15);

  // ---- RSI (P) ----
  add(`=$P${r0}<=30`,                                                       C_GREEN,16);
  add(`=$P${r0}>=70`,                                                       C_RED,  16);
  add(`=AND($P${r0}>30,$P${r0}<70)`,                                        C_GREY, 16);

  // ---- MACD Hist (Q) ----
  add(`=$Q${r0}>0`,                                                         C_GREEN,17);
  add(`=$Q${r0}<0`,                                                         C_RED,  17);
  add(`=OR($Q${r0}=0,$Q${r0}="")`,                                          C_GREY, 17);

  // ---- Divergence (R) ----
  add(`=REGEXMATCH($R${r0},"BULL")`,                                        C_GREEN,18);
  add(`=REGEXMATCH($R${r0},"BEAR")`,                                        C_RED,  18);
  add(`=OR($R${r0}="—",$R${r0}="",NOT(REGEXMATCH($R${r0},"BULL|BEAR")))`,   C_GREY, 18);

  // ---- ADX (S) ----
  // Strong trend (>=25) green; low trend (<15) grey; mid grey
  add(`=$S${r0}>=25`,                                                       C_GREEN,19);
  add(`=$S${r0}<15`,                                                        C_GREY, 19);
  add(`=AND($S${r0}>=15,$S${r0}<25)`,                                       C_GREY, 19);

  // ---- Stoch %K (T) ----
  add(`=$T${r0}<=0.2`,                                                      C_GREEN,20);
  add(`=$T${r0}>=0.8`,                                                      C_RED,  20);
  add(`=AND($T${r0}>0.2,$T${r0}<0.8)`,                                      C_GREY, 20);

  // ---- Support (U) ----
  // Below support = red; within +1% above support = green; else grey
  add(`=AND($U${r0}>0,$E${r0}<$U${r0})`,                                    C_RED,  21);
  add(`=AND($U${r0}>0,$E${r0}>=$U${r0},$E${r0}<=$U${r0}*1.01)`,             C_GREEN,21);
  add(`=AND($U${r0}>0,$E${r0}>$U${r0}*1.01)`,                               C_GREY, 21);

  // ---- Resistance (V) ----
  // Near/at resistance = red; far below resistance = green; else grey
  add(`=AND($V${r0}>0,$E${r0}>=$V${r0}*0.995)`,                             C_RED,  22);
  add(`=AND($V${r0}>0,$E${r0}<=$V${r0}*0.90)`,                              C_GREEN,22);
  add(`=AND($V${r0}>0,$E${r0}>$V${r0}*0.90,$E${r0}<$V${r0}*0.995)`,         C_GREY, 22);

  // ---- Target (W) ----
  // Target meaningfully above price = green; too close = red; else grey
  add(`=AND($W${r0}>0,$W${r0}>=$E${r0}*1.05)`,                              C_GREEN,23);
  add(`=AND($W${r0}>0,$W${r0}<=$E${r0}*1.01)`,                              C_RED,  23);
  add(`=AND($W${r0}>0,$W${r0}>$E${r0}*1.01,$W${r0}<$E${r0}*1.05)`,          C_GREY, 23);

  // ---- ATR (X) as % of price ----
  // Low volatility <=2% = green; high volatility >=5% = red; else grey
  add(`=IFERROR($X${r0}/$E${r0},0)<=0.02`,                                  C_GREEN,24);
  add(`=IFERROR($X${r0}/$E${r0},0)>=0.05`,                                  C_RED,  24);
  add(`=AND(IFERROR($X${r0}/$E${r0},0)>0.02,IFERROR($X${r0}/$E${r0},0)<0.05)`,C_GREY,24);

  // ---- Bollinger %B (Y) ----
  add(`=$Y${r0}<=0.2`,                                                      C_GREEN,25);
  add(`=$Y${r0}>=0.8`,                                                      C_RED,  25);
  add(`=AND($Y${r0}>0.2,$Y${r0}<0.8)`,                                      C_GREY, 25);

  // Apply rules
  sh.setConditionalFormatRules(rules);

  // ---------------------------
  // Hard-hide notes columns (Z, AA)
  // ---------------------------
  safeHideNotes_();

  // ---------------------------
  // Cleanup below actual data end (prevents retained row colors when shrink)
  // ---------------------------
  clearTailFormats_(lastDataRow);
}

/**
 * This colors Row 2 group bars and Row 3 headers using the same group color blocks.
 */
function applyDashboardGroupMapAndColors_(sh) {
  if (!sh) return;

  // ===== GROUP COLOR PALETTE (header-only) =====
  const COLORS = {
    SIGNAL:  "#1F4FD8", // blue
    PRICE:   "#0F766E", // teal
    PERF:    "#374151", // slate
    TREND:   "#14532D", // green
    MOM:     "#7C2D12", // brown
    LEVELS:  "#4C1D95", // purple
    NOTES:   "#111827"  // dark
  };

  const FG = "#FFFFFF"; // white text

  // ===== GROUP → COLUMN MAP (1-indexed) =====
  const GROUPS = [
    { name: "SIGNALING",        c1: 2,  c2: 4,  color: COLORS.SIGNAL }, // B..D
    { name: "PRICE / VOLUME",   c1: 5,  c2: 7,  color: COLORS.PRICE  }, // E..G
    { name: "PERFORMANCE",      c1: 8,  c2: 10, color: COLORS.PERF   }, // H..J
    { name: "TREND",            c1: 11, c2: 15, color: COLORS.TREND  }, // K..O
    { name: "MOMENTUM",         c1: 16, c2: 20, color: COLORS.MOM    }, // P..T
    { name: "LEVELS / RISK",    c1: 21, c2: 25, color: COLORS.LEVELS }, // U..Y
    { name: "NOTES",            c1: 26, c2: 27, color: COLORS.NOTES  }  // Z..AA
  ];

  // ===== COMMON HEADER STYLE =====
  const style = (row, c1, c2, bg) => {
    sh.getRange(row, c1, 1, c2 - c1 + 1)
      .setBackground(bg)
      .setFontColor(FG)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
  };

  // ===== APPLY COLORS =====
  GROUPS.forEach(g => {
    // Row 2: group header bar
    style(2, g.c1, g.c2, g.color);

    // Merge + label row 2
    const r2 = sh.getRange(2, g.c1, 1, g.c2 - g.c1 + 1);
    try { r2.breakApart(); } catch (e) {}
    if (g.c1 !== g.c2) r2.merge();
    r2.setValue(g.name);

    // Row 3: column headers (same group color)
    style(3, g.c1, g.c2, g.color);
  });
}



/**
 * Call helper — keep generateDashboardSheet clean.
 * Call this AFTER you set A4 formula and flush.
 */
function applyDashboardBloombergFormattingAfterRefresh_(dashboardSheet) {
  SpreadsheetApp.flush(); // ensure FILTER() spill exists
  applyDashboardBloombergFormatting_(dashboardSheet, 4); // data starts at row 4
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

  
function myFunction() {
  generateMasterMobileReport();
}

