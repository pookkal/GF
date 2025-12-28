/**
* ==============================================================================
* BASELINE LABEL: STABLE_MASTER_DEC25_BASE_v3_6 ADX formula fix
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
function generateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;

  const tickers = getCleanTickers(inputSheet);
  let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");
  dataSheet.clear().clearFormats();

  dataSheet.getRange("A1")
    .setValue("Last Update: " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm"))
    .setFontWeight("bold")
    .setFontColor("blue");

  tickers.forEach((ticker, i) => {
    const colStart = (i * 7) + 1;

    dataSheet.getRange(2, colStart)
      .setNumberFormat("@")
      .setValue(ticker)
      .setFontWeight("bold");

    dataSheet.getRange(3, colStart).setValue("ATH:");
    dataSheet.getRange(3, colStart + 1)
      .setFormula(`=MAX(QUERY(GOOGLEFINANCE("${ticker}", "high", "1/1/2000", TODAY()), "SELECT Col2 LABEL Col2 ''"))`);

    dataSheet.getRange(4, colStart)
      .setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "all", TODAY()-800, TODAY()), "No Data")`);

    dataSheet.getRange(5, colStart, 1000, 1).setNumberFormat("yyyy-mm-dd");
    dataSheet.getRange(5, colStart + 1, 1000, 5).setNumberFormat("#,##0.00");
  });
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

  // Groups across A..AA (AB reserved for timestamp / state header)
  styleGroup("A1:A1",   "IDENTITY",        "#263238"); // A
  styleGroup("B1:D1",   "SIGNALING",       "#0D47A1"); // B-D
  styleGroup("E1:G1",   "PRICE / VOLUME",  "#1B5E20"); // E-G
  styleGroup("H1:J1",   "PERFORMANCE",     "#004D40"); // H-J
  styleGroup("K1:O1",   "TREND",           "#2E7D32"); // K-O
  styleGroup("P1:T1",   "MOMENTUM",        "#33691E"); // P-T
  styleGroup("U1:Y1",   "LEVELS / RISK",   "#B71C1C"); // U-Y
  styleGroup("Z1:AA1",  "NOTES",           "#212121"); // Z-AA

  // AB1 timestamp (not merged)
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

  tickers.forEach((ticker, i) => {
    const row = i + 3;
    const t = ticker.toString().trim().toUpperCase();
    restoredStates.push([stateMap[t] || ""]);

    // DATA block start (each ticker is 7 cols in DATA)
    const tDS = (i * 7) + 1;
    const highCol  = columnToLetter(tDS + 2);
    const lowCol   = columnToLetter(tDS + 3);
    const closeCol = columnToLetter(tDS + 4);
    const volCol   = columnToLetter(tDS + 5);
    const lastRow = `COUNTA(DATA!$${closeCol}:$${closeCol})`;

    // SIGNAL (B)
    const fSignal =
      `=IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}"LOADING"${SEP}` +
      `IFS(` +
      `$E${row}<$U${row}${SEP}"Stop-Out"${SEP}` +
      `$E${row}<$O${row}${SEP}"Risk-Off (Below SMA200)"${SEP}` +
      `$S${row}<15${SEP}"Range-Bound (Low ADX)"${SEP}` +
      `AND($G${row}>=1.5${SEP}$E${row}>=$V${row}*0.995${SEP}$Q${row}>0${SEP}$S${row}>=18)${SEP}"Breakout (High Volume)"${SEP}` +
      `AND($T${row}<=0.20${SEP}$E${row}>$U${row}${SEP}$S${row}>=18)${SEP}"Mean Reversion (Oversold)"${SEP}` +
      `AND($T${row}>=0.80${SEP}$E${row}>=$V${row}*0.97)${SEP}"Mean Reversion (Overbought)"${SEP}` +
      `AND($E${row}>$O${row}${SEP}$Q${row}>0${SEP}$S${row}>=18)${SEP}"Trend Continuation"${SEP}` +
      `TRUE${SEP}"Hold / Monitor"` +
      `))`;

    // DECISION (C) — includes SELL-side states
    const fDecision =
      `=IF($B${row}="LOADING"${SEP}"LOADING"${SEP}` +
      `IFS(` +
      `$B${row}="Stop-Out"${SEP}"Stop-Out"${SEP}` +
      `OR($D${row}="ZOMBIE"${SEP}$D${row}="BUBBLE")${SEP}"Avoid"${SEP}` +
      `$B${row}="Risk-Off (Below SMA200)"${SEP}"Avoid"${SEP}` +
      `OR(REGEXMATCH($B${row}${SEP}"Mean Reversion \\(Overbought\\)")${SEP}AND($P${row}>=70${SEP}$E${row}>=$V${row}*0.97))${SEP}"Take Profit"${SEP}` +
      `AND($Q${row}<0${SEP}$E${row}<$N${row})${SEP}"Reduce (Momentum Weak)"${SEP}` +
      `AND($B${row}="Breakout (High Volume)"${SEP}$J${row}>=1.5${SEP}$S${row}>=20)${SEP}"Trade Long"${SEP}` +
      `AND($B${row}="Trend Continuation"${SEP}$J${row}>=1.3${SEP}$S${row}>=18)${SEP}"Accumulate"${SEP}` +
      `AND(REGEXMATCH($B${row}${SEP}"Mean Reversion")${SEP}$J${row}>=1.2${SEP}$S${row}>=18)${SEP}"Trade Long"${SEP}` +
      `AND($X${row}>0${SEP}$E${row}>$M${row}+(2*$X${row}))${SEP}"Reduce (Overextended)"${SEP}` +
      `$B${row}="Range-Bound (Low ADX)"${SEP}"Hold / Monitor"${SEP}` +
      `TRUE${SEP}"Hold / Monitor"` +
      `))`;

    // FUNDAMENTAL (D)
    const fFund =
      `=IFERROR(LET(eps${SEP}GOOGLEFINANCE($A${row}${SEP}"eps")${SEP}` +
      `pe${SEP}GOOGLEFINANCE($A${row}${SEP}"pe")${SEP}` +
      `IFS(` +
      `eps<0${SEP}"ZOMBIE"${SEP}` +
      `AND(pe>0${SEP}pe>50)${SEP}"PRICED FOR PERFECTION"${SEP}` +
      `AND(pe>0${SEP}pe<25${SEP}eps>0)${SEP}"VALUE"${SEP}` +
      `AND(pe>30${SEP}eps<0.1)${SEP}"BUBBLE"${SEP}` +
      `TRUE${SEP}"FAIR"` +
      `))${SEP}"FAIR")`;

    // E..Y
    const fPrice  = `=ROUND(IFERROR(GOOGLEFINANCE("${t}"${SEP}"price")${SEP}0)${SEP}2)`;
    const fChg    = `=IFERROR(GOOGLEFINANCE("${t}"${SEP}"changepct")/100${SEP}0)`;
    const fRVOL   = `=ROUND(IFERROR(OFFSET(DATA!$${volCol}$4${SEP}${lastRow}-1${SEP}0) / AVERAGE(OFFSET(DATA!$${volCol}$4${SEP}${lastRow}-21${SEP}0${SEP}20))${SEP}1)${SEP}2)`;
    const fATH    = `=IFERROR(DATA!${columnToLetter(tDS + 1)}3${SEP}0)`;
    const fATHPct = `=IFERROR(($E${row}-$H${row})/MAX(0.01${SEP}$H${row})${SEP}0)`;
    const fRR     = `=IFERROR(ROUND(($V${row}-$E${row})/MAX(0.01${SEP}$E${row}-$U${row})${SEP}2)${SEP}0)`;
    const fStars  = `=REPT("★"${SEP} ($E${row}>$M${row}) + ($E${row}>$N${row}) + ($E${row}>$O${row}))`;
    const fTrend  = `=IF($E${row}>$O${row}${SEP}"BULL"${SEP}"BEAR")`;
    const fSMA20  = `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4${SEP}${lastRow}-20${SEP}0${SEP}20))${SEP}0)${SEP}2)`;
    const fSMA50  = `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4${SEP}${lastRow}-50${SEP}0${SEP}50))${SEP}0)${SEP}2)`;
    const fSMA200 = `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4${SEP}${lastRow}-200${SEP}0${SEP}200))${SEP}0)${SEP}2)`;
    const fRSI    = `=LIVERSI(DATA!$${closeCol}$4:$${closeCol}${SEP}$E${row})`;
    const fMACD   = `=LIVEMACD(DATA!$${closeCol}$4:$${closeCol}${SEP}$E${row})`;
    const fDiv =
      `=IFERROR(IFS(` +
      `AND($E${row}<INDEX(DATA!$${closeCol}:$${closeCol}${SEP}${lastRow}-14)${SEP}$P${row}>50)${SEP}"BULL DIV"${SEP}` +
      `AND($E${row}>INDEX(DATA!$${closeCol}:$${closeCol}${SEP}${lastRow}-14)${SEP}$P${row}<50)${SEP}"BEAR DIV"${SEP}` +
      `TRUE${SEP}"—")${SEP}"—")`;
    const fADX    = `=IFERROR(LIVEADX(DATA!$${highCol}$4:$${highCol}, DATA!$${lowCol}$4:$${lowCol}, DATA!$${closeCol}$4:$${closeCol}, $E${row}), 0)`;
    const fStoch  = `=LIVESTOCHK(DATA!$${highCol}$4:$${highCol}${SEP}DATA!$${lowCol}$4:$${lowCol}${SEP}DATA!$${closeCol}$4:$${closeCol}${SEP}$E${row})`;
    const fSup    = `=ROUND(IFERROR(MIN(OFFSET(DATA!$${lowCol}$4${SEP}${lastRow}-21${SEP}0${SEP}20))${SEP}$E${row}*0.9)${SEP}2)`;
    const fRes    = `=ROUND(IFERROR(MAX(OFFSET(DATA!$${highCol}$4${SEP}${lastRow}-51${SEP}0${SEP}50))${SEP}$E${row}*1.1)${SEP}2)`;
    const fTgt    = `=ROUND($E${row} + (($E${row}-$U${row}) * 3)${SEP}2)`;
    const fATR =
      `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(` +
      `OFFSET(DATA!$${highCol}$4${SEP}${lastRow}-14${SEP}0${SEP}14) - OFFSET(DATA!$${lowCol}$4${SEP}${lastRow}-14${SEP}0${SEP}14)` +
      `))${SEP}0)${SEP}2)`;
    const fBBP    = `=ROUND(IFERROR((($E${row}-$M${row}) / (4*STDEV(OFFSET(DATA!$${closeCol}$4${SEP}${lastRow}-20${SEP}0${SEP}20)))) + 0.5${SEP}0.5)${SEP}2)`;

    // Z TECH NOTES — your original narrative + safe rationale line (IFS)
    const fTechNotes =
      `=IF($B${row}="LOADING"${SEP}"LOADING"${SEP}` +
      `"VOLUME: RVOL "&TEXT($G${row}${SEP}"0.00")&"x — "&IF($G${row}>=1.5${SEP}"above-average participation (conviction)."${SEP}"sub-average participation (weak sponsorship).")&CHAR(10)&` +
      `"TREND REGIME: Price "&TEXT($E${row}${SEP}"0.00")&" vs SMA200 "&TEXT($O${row}${SEP}"0.00")&" — "&IF($E${row}>=$O${row}${SEP}"long-term bullish structure intact."${SEP}"risk-off regime below SMA200 (avoid chasing).")&CHAR(10)&` +
      `"VOLATILITY / STRETCH: ATR(14) "&TEXT($X${row}${SEP}"0.00")&"; SMA20 "&TEXT($M${row}${SEP}"0.00")&"; Stretch="&TEXT(($E${row}-$M${row})/MAX(0.01${SEP}$X${row})${SEP}"0.0")&"x ATR — "&IF($E${row}>$M${row}+2*$X${row}${SEP}"overextended (>+2x ATR)."${SEP}"within normal range (≤±2x ATR).")&CHAR(10)&` +
      `"MOMENTUM: RSI(14) "&TEXT($P${row}${SEP}"0.0")&" — "&IF($P${row}>=70${SEP}"overbought."${SEP}IF($P${row}<=30${SEP}"oversold."${SEP}IF($P${row}>=50${SEP}"positive bias."${SEP}"negative bias.")))&` +
      `" | MACD Hist "&TEXT($Q${row}${SEP}"0.000")&" — "&IF($Q${row}>0${SEP}"positive momentum."${SEP}"negative momentum.")&CHAR(10)&` +
      `"TREND STRENGTH: ADX(14) "&TEXT($S${row}${SEP}"0.0")&" — "&IF($S${row}<15${SEP}"no trend."${SEP}IF($S${row}<25${SEP}"weak trend."${SEP}IF($S${row}<40${SEP}"strong trend."${SEP}"very strong trend.")))&` +
      `" | Stoch %K "&TEXT($T${row}${SEP}"0.0%")&" — "&IF($T${row}>=0.8${SEP}"overbought."${SEP}IF($T${row}<=0.2${SEP}"oversold."${SEP}"neutral."))&CHAR(10)&` +
      `"RISK/REWARD: "&TEXT($J${row}${SEP}"0.00")&"x — "&IF($J${row}>=3${SEP}"institutional-grade asymmetry (≥3x)."${SEP}IF($J${row}>=2${SEP}"acceptable tactical edge (≥2x)."${SEP}"sub-optimal payout (<2x)."))&CHAR(10)&` +
      `"LEVELS: Support "&TEXT($U${row}${SEP}"0.00")&" | Resistance "&TEXT($V${row}${SEP}"0.00")&" | Target "&TEXT($W${row}${SEP}"0.00")&"."&CHAR(10)&` +
      `"DECISION RATIONALE: "&IFS(` +
        `$C${row}="Take Profit"${SEP}"Overbought / near Resistance; lock gains."${SEP}` +
        `$C${row}="Reduce (Momentum Weak)"${SEP}"MACD<0 and Price<SMA50; reduce exposure."${SEP}` +
        `$C${row}="Reduce (Overextended)"${SEP}"Stretch >2x ATR above SMA20; trim/avoid chase."${SEP}` +
        `$C${row}="Stop-Out"${SEP}"Broke below Support; structure invalid."${SEP}` +
        `$C${row}="Avoid"${SEP}"Blocked by regime/fundamental risk."${SEP}` +
        `$C${row}="Trade Long"${SEP}"Gates passed; use Support stop; target Resistance/3:1."${SEP}` +
        `$C${row}="Accumulate"${SEP}"Trend intact; add on pullbacks."${SEP}` +
        `TRUE${SEP}"Hold/Monitor: edge insufficient."` +
      `)` +
      `)`;

    // AA FUND NOTES (kept simple, safe)
    const fFundNotes =
      `=IF($B${row}="LOADING"${SEP}"LOADING"${SEP}` +
      `"VALUATION: "&$D${row}&CHAR(10)&` +
      `"REGIME: "&IF($E${row}>=$O${row}${SEP}"Above SMA200 (Risk-On)."${SEP}"Below SMA200 (Risk-Off).")&CHAR(10)&` +
      `"VERDICT: "&$C${row}` +
      `)`;

    formulas.push([
      fSignal,     // B
      fDecision,   // C
      fFund,       // D
      fPrice,      // E
      fChg,        // F
      fRVOL,       // G
      fATH,        // H
      fATHPct,     // I
      fRR,         // J
      fStars,      // K
      fTrend,      // L
      fSMA20,      // M
      fSMA50,      // N
      fSMA200,     // O
      fRSI,        // P
      fMACD,       // Q
      fDiv,        // R
      fADX,        // S
      fStoch,      // T
      fSup,        // U
      fRes,        // V
      fTgt,        // W
      fATR,        // X
      fBBP,        // Y
      fTechNotes,  // Z
      fFundNotes   // AA
    ]);
  });

  if (tickers.length > 0) {
    // B..AA (26 cols)
    calc.getRange(3, 2, formulas.length, 26).setFormulas(formulas);
    // AB LAST_STATE restore
    calc.getRange(3, 28, restoredStates.length, 1).setValues(restoredStates);
  }

  // ------------------------------------------------------------------
  // FORMATTING (FIXED HEIGHT, NOT DRIVEN BY Z)
  // ------------------------------------------------------------------
  const lr = Math.max(calc.getLastRow(), 3);
  calc.setFrozenRows(2);

  if (lr > 2) {
    const dataRows = lr - 2;

    // Fixed row height (approx 4 lines), independent of Z content
    calc.setRowHeights(3, dataRows, 72);

    // Left align + wrap for all data cells
    calc.getRange(3, 1, dataRows, 28)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle")
      .setWrap(true);
  }

  // Column widths (dense + notes wider)
  for (let c = 1; c <= 25; c++) calc.setColumnWidth(c, 90);
  calc.setColumnWidth(26, 420); // Z TECH NOTES
  calc.setColumnWidth(27, 420); // AA FUND NOTES
  calc.setColumnWidth(28, 140); // AB LAST_STATE

  // Number formats
  calc.getRange("F3:F").setNumberFormat("0.00%");
  calc.getRange("I3:I").setNumberFormat("0.00%");
  calc.getRange("T3:T").setNumberFormat("0.00%");
  calc.getRange("Y3:Y").setNumberFormat("0.00%");

  // Borders:
  // - Black grid for the whole table
  // - White border band for row 1 and row 2
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
function generateDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = ss.getSheetByName("INPUT");
  if (!input) return;

  const dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD");
  dashboard.clear().clearFormats();

  // ============================================================
  // ROW 1 — Controls (A1..G1) + D1 checkbox
  // ============================================================
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
    .setVerticalAlignment("middle")
    .setValue(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM dd, yyyy | HH:mm:ss"));

  // White border rows 1–3
  dashboard.getRange("A1:AA3")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // ============================================================
  // ROW 2 — Group headers (merged blocks)
  // ============================================================
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

  // Allow wrapping for group header row
  dashboard.getRange("A2:AA2").setWrap(true);

  // ============================================================
  // ROW 3 — Column headers (Dashboard order; C/D swapped)
  // ============================================================
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

  // ============================================================
  // ROW 4 — Hardened filter formula (INPUT-driven)
  // ONLY CHANGE: SORT index = 6 (Change % desc)
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
  SpreadsheetApp.flush();

  // ============================================================
  // GOVERNANCE FORMATTING
  // IMPORTANT: spilled output => format a deterministic window
  // ============================================================
  dashboard.setFrozenRows(3);
  dashboard.setFrozenColumns(1);

  const DATA_START_ROW = 4;
  const DATA_ROWS = 500;
  const DATA_END_ROW = DATA_START_ROW + DATA_ROWS - 1;

  // Column widths
  for (let c = 1; c <= 25; c++) dashboard.setColumnWidth(c, 90);
  dashboard.setColumnWidth(26, 420); // Z
  dashboard.setColumnWidth(27, 420); // AA

  // --- WRAP BEHAVIOR YOU REQUESTED ---
  // A..Y: WRAP ON (so text wraps inside fixed height)
  dashboard.getRange(DATA_START_ROW, 1, DATA_ROWS, 25)
    .setWrap(true);

  // Z..AA: CLIP (no wrap, prevents row height expansion)
  dashboard.getRange(DATA_START_ROW, 26, DATA_ROWS, 2)
    .setWrap(false)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Left alignment for entire formatted region (data + headers)
  dashboard.getRange(1, 1, DATA_END_ROW, 27)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");

  // Keep row-1 controls centered (optional—remove if you truly want all-left)
  dashboard.getRange("A1:D1").setHorizontalAlignment("center");
  dashboard.getRange("E1:G1").setHorizontalAlignment("center");

  // Borders
  dashboard.getRange(1, 1, DATA_END_ROW, 27)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // White band borders row 1–2
  dashboard.getRange("A1:AA2")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // Number formats
  dashboard.getRange(`F${DATA_START_ROW}:F${DATA_END_ROW}`).setNumberFormat("0.00%");
  dashboard.getRange(`I${DATA_START_ROW}:I${DATA_END_ROW}`).setNumberFormat("0.00%");
  dashboard.getRange(`T${DATA_START_ROW}:T${DATA_END_ROW}`).setNumberFormat("0.00%");
  dashboard.getRange(`Y${DATA_START_ROW}:Y${DATA_END_ROW}`).setNumberFormat("0.00%");

  // ============================================================
  // CONDITIONAL FORMATTING (same rules, deterministic ranges)
  // ============================================================
  const rules = [];

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#B71C1C")
      .setBold(true)
      .setRanges([
        dashboard.getRange(`F${DATA_START_ROW}:F${DATA_END_ROW}`),
        dashboard.getRange(`I${DATA_START_ROW}:I${DATA_END_ROW}`),
        dashboard.getRange(`Q${DATA_START_ROW}:Q${DATA_END_ROW}`)
      ])
      .build()
  );

  // RSI (P)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$P4>=70")
    .setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange(`P${DATA_START_ROW}:P${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$P4<=30")
    .setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange(`P${DATA_START_ROW}:P${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($P4>=50,$P4<70)")
    .setFontColor("#1B5E20")
    .setRanges([dashboard.getRange(`P${DATA_START_ROW}:P${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($P4>30,$P4<50)")
    .setFontColor("#E65100")
    .setRanges([dashboard.getRange(`P${DATA_START_ROW}:P${DATA_END_ROW}`)]).build());

  // ADX (S)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$S4>=25")
    .setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange(`S${DATA_START_ROW}:S${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$S4<15")
    .setFontColor("#616161")
    .setRanges([dashboard.getRange(`S${DATA_START_ROW}:S${DATA_END_ROW}`)]).build());

  // Stoch (T)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$T4>=0.8")
    .setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange(`T${DATA_START_ROW}:T${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$T4<=0.2")
    .setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange(`T${DATA_START_ROW}:T${DATA_END_ROW}`)]).build());

  // %B (Y)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$Y4>=0.8")
    .setFontColor("#B71C1C")
    .setRanges([dashboard.getRange(`Y${DATA_START_ROW}:Y${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$Y4<=0.2")
    .setFontColor("#1B5E20")
    .setRanges([dashboard.getRange(`Y${DATA_START_ROW}:Y${DATA_END_ROW}`)]).build());

  // SIGNAL (B)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Breakout|Trend Continuation|RVOL")')
    .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange(`B${DATA_START_ROW}:B${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Mean Reversion|Bounce|Oversold|Overbought")')
    .setBackground("#FFF8E1").setFontColor("#E65100").setBold(true)
    .setRanges([dashboard.getRange(`B${DATA_START_ROW}:B${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Range|Chop|Hold")')
    .setBackground("#F5F5F5").setFontColor("#616161")
    .setRanges([dashboard.getRange(`B${DATA_START_ROW}:B${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Risk-Off|Stop")')
    .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange(`B${DATA_START_ROW}:B${DATA_END_ROW}`)]).build());

  // DECISION (D swapped)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($D4,"Trade|Accumulate|Buy")')
    .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange(`D${DATA_START_ROW}:D${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($D4,"Reduce|Trim|Take Profit")')
    .setBackground("#FFF8E1").setFontColor("#E65100").setBold(true)
    .setRanges([dashboard.getRange(`D${DATA_START_ROW}:D${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($D4,"Hold|Monitor|Wait")')
    .setBackground("#F5F5F5").setFontColor("#616161")
    .setRanges([dashboard.getRange(`D${DATA_START_ROW}:D${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($D4,"Avoid|Stop")')
    .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange(`D${DATA_START_ROW}:D${DATA_END_ROW}`)]).build());

  // FUNDAMENTAL (C swapped)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($C4,"VALUE|GEM|FAIR")')
    .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange(`C${DATA_START_ROW}:C${DATA_END_ROW}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($C4,"ZOMBIE|BUBBLE")')
    .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange(`C${DATA_START_ROW}:C${DATA_END_ROW}`)]).build());

  dashboard.setConditionalFormatRules(rules);

    // Fix row height AFTER wrap settings (so it stays fixed)
    SpreadsheetApp.flush();
    dashboard.getRange(DATA_START_ROW, 26, DATA_ROWS, 2).setWrap(false); // Z:AA
    dashboard.setRowHeights(DATA_START_ROW, DATA_ROWS, 12); // ~2 lines
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
  sh.getRange("B36").setNumberFormat("0.00%");        // %B

  SpreadsheetApp.flush();

  updateDynamicChart(); // ensure chart & lines appear
}


/// ------------------------------------------------------------
// updateDynamicChart() — timestamp REMOVED, row 7 left empty
// ------------------------------------------------------------
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;

  SpreadsheetApp.flush();

  // Ticker in A1 (merged A1:B1)
  const ticker = String(sheet.getRange("A1").getValue() || "").trim();
  if (!ticker) return;

  // Interval B6 + StartDate from B5 (source of truth)
  const interval = String(sheet.getRange("B6").getValue() || "DAILY").toUpperCase();
  const isWeekly = interval === "WEEKLY";

  let startDate = sheet.getRange("B5").getValue();
  if (!(startDate instanceof Date)) {
    const now = new Date();
    startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 14);
  }

  // ------------------------------------------------------------
  // Sidebar values: robust label match (works with SUPPORT/SUP and RESISTANCE/RES)
  // ------------------------------------------------------------
  const sidebar = getSidebarValuesByLabels_(sheet, ["PRICE", "SUPPORT", "RESISTANCE", "SUP", "RES"]);

  const livePrice = Number(sidebar["PRICE"]) || 0;
  const supportVal = Number(sidebar["SUPPORT"]) || Number(sidebar["SUP"]) || 0;
  const resistanceVal = Number(sidebar["RESISTANCE"]) || Number(sidebar["RES"]) || 0;

  // Find ticker column in DATA (row 2 has ticker headers)
  const headers = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIdx = headers.indexOf(ticker);
  if (colIdx === -1) return;

  // Pull 6 cols: date, open, high, low, close, volume
  const raw = dataSheet.getRange(1, colIdx + 1, dataSheet.getLastRow(), 6).getValues();

  let master = [];
  let vols = [];
  let prices = [];

  for (let i = 4; i < raw.length; i++) {
    const d = raw[i][0];
    const close = Number(raw[i][4]);
    const vol = Number(raw[i][5]);
    if (!d || !(d instanceof Date) || !isFinite(close) || close < 0.01) continue;
    if (d < startDate) continue;
    if (isWeekly && d.getDay() !== 5) continue; // Fridays

    const slice = raw
      .slice(Math.max(4, i - 200), i + 1)
      .map(r => Number(r[4]))
      .filter(n => isFinite(n) && n > 0);

    const s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
    const s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
    const s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;

    const prevClose = (i > 4) ? Number(raw[i - 1][4]) : close;

    master.push([
      d,
      close,
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

  // Write region Z..AH (col 26..34)
  sheet.getRange(3, 26, 2000, 9).clearContent();
  if (master.length === 0) return;

  if (supportVal > 0) prices.push(supportVal);
  if (resistanceVal > 0) prices.push(resistanceVal);

  const cleanPrices = prices.filter(p => typeof p === "number" && isFinite(p) && p > 0);
  if (!cleanPrices.length) return;

  const minP = Math.min(...cleanPrices) * 0.98;
  const maxP = Math.max(...cleanPrices) * 1.02;

  const cleanVols = vols.filter(v => typeof v === "number" && isFinite(v) && v >= 0);
  const maxVol = Math.max(...cleanVols, 1);

  // Headers
  sheet.getRange(2, 26, 1, 9)
    .setValues([["Date", "Price", "Bull Vol", "Bear Vol", "SMA 20", "SMA 50", "SMA 200", "Resistance", "Support"]])
    .setFontWeight("bold")
    .setFontColor("white");

  // Data + Date format
  sheet.getRange(3, 26, master.length, 9).setValues(master);
  sheet.getRange(3, 26, master.length, 1).setNumberFormat("dd/MM/yy");

  SpreadsheetApp.flush();

  // Rebuild chart
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
    .setOption("legend", { position: "top", textStyle: { fontSize: 10 } })
    // ✅ Chart at C7
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
/**
* ------------------------------------------------------------------
* START MONITOR (UPDATED TEXT: DECISION changes + includes SELL/REDUCE)
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
  rows.push(["1) DASHBOARD COLUMN DEFINITIONS (TECHNICAL)", "", "", ""]);
  rows.push(["COLUMN", "WHAT IT IS", "HOW IT IS USED", "USER ACTION"]);

  const cols = [
    ["Ticker", "Symbol (key)", "Join key across DATA/CALCULATIONS/CHART", "Select for chart / review notes."],
    ["SIGNAL", "Technical setup label (rules engine)", "Describes setup type (breakout / trend / mean-rev / risk-off / stop-out)", "Use as setup classification; DECISION is what you act on."],
    ["FUNDAMENTAL", "EPS + P/E risk bucket", "Blocks trades in weak quality/extreme valuation regimes", "Prefer VALUE/FAIR; avoid ZOMBIE/BUBBLE."],
    ["DECISION", "Action label (gated by regime + R:R + momentum)", "Final instruction (trade/accumulate/avoid/stop/trim/profit)", "Primary action field."],
    ["Price", "Live last price (GOOGLEFINANCE)", "Used for regime tests, distance-to-levels, ATR stretch", "Confirm price vs SMA200 & levels."],
    ["Change %", "Daily % change", "Context (tape), not a signal alone", "Do not chase moves without a setup."],
    ["Vol Trend", "Relative volume proxy (RVOL)", "Conviction filter for breakouts", "Prefer >=1.5x for breakout validation."],
    ["ATH (TRUE)", "All-time high reference", "Context for price discovery / overhead supply", "Avoid chasing into ceilings without RVOL."],
    ["ATH Diff %", "Distance from ATH", "Pullback vs near-ATH classification", "Use with regime + levels."],
    ["R:R Quality", "Reward/Risk ratio proxy", "Trade quality gate", ">=3 excellent; 2–3 acceptable; <2 poor."],
    ["Trend Score", "★ count (Price above SMAs)", "Quick structure strength read", "3★ strongest; <2★ caution."],
    ["Trend State", "Bull/Bear via SMA200", "Defines risk-on vs risk-off playbook", "Below SMA200 = risk-off bias."],
    ["SMA 20", "Short-term mean", "Stretch anchor; mean reversion reference", "Avoid buying when >2x ATR above SMA20."],
    ["SMA 50", "Medium trend line", "Momentum/structure check (used in Reduce Momentum Weak)", "If lost with MACD<0, reduce risk."],
    ["SMA 200", "Long-term regime line", "Primary risk-on/risk-off filter", "Below: avoid trend-chasing."],
    ["RSI", "Momentum oscillator (0–100)", "Overbought/oversold + bias filter", "<30 oversold; >70 overbought; 50 bias."],
    ["MACD Hist", "Impulse (positive/negative)", "Momentum confirmation / deterioration", "Negative impulse with SMA50 loss = reduce."],
    ["Divergence", "Price vs RSI divergence heuristic", "Early reversal warning", "Bull div supports bounce; bear div warns."],
    ["ADX (14)", "Trend strength", "Chop vs trend filter", "<15 range; 15–25 weak; >=25 trend."],
    ["Stoch %K (14)", "Fast oscillator (0–1)", "Timing within regimes", "<0.2 oversold; >0.8 overbought."],
    ["Support", "20-day min low proxy", "Risk line / invalidation reference", "Break below = Stop-Out."],
    ["Resistance", "50-day max high proxy", "Ceiling / target reference", "Near resistance + overbought = Take Profit."],
    ["Target (3:1)", "Tactical take-profit projection", "Planning exits; not a forecast", "Use for trade planning."],
    ["ATR (14)", "Volatility proxy", "Sizing/stops + stretch detection", "Higher ATR = wider stops / smaller size."],
    ["Bollinger %B", "Band position proxy", "Compression/range heuristic", "Low %B + low ADX = chop."],
    ["TECH NOTES", "Narrative (indicator values + rationale)", "Explains what is driving the setup and action", "Read before acting."],
    ["FUND NOTES", "Narrative (fund + regime + verdict)", "Explains why decision is allowed/blocked", "Respect blockers (ZOMBIE/BUBBLE, risk-off)."]
  ];
  cols.forEach(r => rows.push(r));

  // SIGNAL vocabulary (aligned to your SIGNAL formula outputs)
  rows.push(["", "", "", ""]);
  rows.push(["2) SIGNAL — FULL VOCABULARY (WHAT IT MEANS + WHAT TO DO)", "", "", ""]);
  rows.push(["SIGNAL VALUE", "TECHNICAL DEFINITION", "WHEN IT TRIGGERS", "EXPECTED USER ACTION"]);

  const signal = [
    ["Stop-Out", "Price < Support", "Breakdown through support floor", "Exit / do not average down. Wait for base."],
    ["Risk-Off (Below SMA200)", "Price < SMA200", "Long-term risk-off regime", "Avoid chasing; only tactical trades with strict risk."],
    ["Range-Bound (Low ADX)", "ADX < 15", "No trend / chop regime", "Range tactics only; smaller size; tighter targets."],
    ["Breakout (High Volume)", "RVOL high + price near/above resistance + MACD>0 + ADX>=18", "Breakout attempt with sponsorship", "Only actionable when DECISION says Trade Long (gates pass)."],
    ["Mean Reversion (Oversold)", "StochK<=0.20 + price above support + ADX>=18", "Oversold timing in tradable structure", "Tactical long only if DECISION says Trade Long."],
    ["Mean Reversion (Overbought)", "StochK>=0.80 near resistance", "Overbought timing into ceiling", "Take profits / avoid new longs; DECISION should often be Take Profit."],
    ["Trend Continuation", "Above SMA200 with MACD>0 and ADX>=18", "Uptrend continuation regime", "Accumulate if DECISION says Accumulate; avoid chasing stretch."]
  ];
  signal.forEach(r => rows.push(r));

  // FUNDAMENTAL vocabulary (aligned to your FUND formula outputs)
  rows.push(["", "", "", ""]);
  rows.push(["3) FUNDAMENTAL — FULL VOCABULARY (FILTER + RISK)", "", "", ""]);
  rows.push(["FUNDAMENTAL VALUE", "WHAT IT MEANS (IN THIS MODEL)", "RISK PROFILE", "EXPECTED USER ACTION"]);

  const fund = [
    ["VALUE", "EPS>0 and P/E<25 (lower valuation-risk bucket)", "Lower valuation risk vs others", "Prefer for breakouts/trend setups when tech confirms."],
    ["FAIR", "Neither cheap nor extreme (fallback)", "Neutral valuation risk", "Trade only when technical gates pass (R:R, ADX)."],
    ["PRICED FOR PERFECTION", "EPS positive but P/E elevated (pe>50)", "Multiple compression risk", "Only take best setups; be selective."],
    ["BUBBLE", "High P/E with weak earnings profile", "High downside risk", "Avoid longs; only tactical with strict risk (not preferred)."],
    ["ZOMBIE", "EPS negative / fragile quality", "High blow-up risk", "Avoid."]
  ];
  fund.forEach(r => rows.push(r));

  // DECISION vocabulary (aligned to updated DECISION formula outputs)
  rows.push(["", "", "", ""]);
  rows.push(["4) DECISION — FULL VOCABULARY (WHAT TO DO)", "", "", ""]);
  rows.push(["DECISION VALUE", "WHY IT HAPPENS (ENGINE RULE)", "RISK GATES", "EXPECTED USER ACTION"]);

  const decision = [
    ["Stop-Out", "SIGNAL Stop-Out (Price < Support)", "Structure invalidated", "Exit / stand aside."],
    ["Avoid", "Fundamental blocker (ZOMBIE/BUBBLE) OR Risk-Off (<SMA200)", "Hard block", "No trade; remove from active list."],
    ["Take Profit", "Overbought / near Resistance (or RSI>=70 near resistance)", "Sell-side timing state", "Take profit / trim; do not chase new longs."],
    ["Reduce (Momentum Weak)", "MACD Hist < 0 AND Price < SMA50", "Deterioration gate", "Reduce exposure to avoid drawdown; tighten risk."],
    ["Trade Long", "Breakout or Oversold mean-reversion with gates satisfied", "ADX/R:R gates", "Tactical entry; stop at Support; target Resistance/3:1."],
    ["Accumulate", "Trend Continuation with acceptable R:R and ADX", "Trend gate", "Scale in on pullbacks; avoid chasing."],
    ["Reduce (Overextended)", "Price > SMA20 + 2×ATR", "Stretch gate", "Trim or avoid new entries; wait for mean reversion."],
    ["Hold / Monitor", "No edge or gates not met", "Neutral", "Do nothing; monitor levels and signals."],
    ["LOADING", "Data not ready", "N/A", "Wait for refresh; do not act."]
  ];
  decision.forEach(r => rows.push(r));

  // Quick playbook (updated to include sell states)
  rows.push(["", "", "", ""]);
  rows.push(["5) QUICK PLAYBOOK (HOW TO USE THE TERMINAL)", "", "", ""]);
  rows.push(["RULE", "WHY", "WHAT TO LOOK FOR", "WHAT TO AVOID"]);
  rows.push(["Trend trades", "Best expectancy in strong regimes", "Risk-On (>=SMA200), ADX>=25, MACD>0, RVOL>=1.5", "Buying in Risk-Off or with ADX<15."]);
  rows.push(["Range trades", "Chop markets are mean-reverting", "ADX<15 and price near Support/Resistance", "Chasing mid-range; poor R:R."]);
  rows.push(["Profit-taking", "Avoid giving back gains", "Take Profit (overbought near Resistance), Reduce (Overextended)", "Adding new longs when stretched/overbought."]);
  rows.push(["Loss avoidance", "Stops define survival", "Stop-Out, Reduce (Momentum Weak)", "Averaging down below Support."]);
  rows.push(["R:R gating", "Prevents low-quality trades", "R:R>=2 tactical; >=3 preferred", "R:R<2 unless exceptional setup."]);

  // Write
  sh.getRange(1, 1, rows.length, 4).setValues(rows);

  // Styling (keep your existing style)
  sh.setColumnWidth(1, 210);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 320);
  sh.setColumnWidth(4, 260);

  sh.setRowHeights(1, Math.min(rows.length, 800), 18);
  sh.setFrozenRows(3);

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
    if (/^\d\)/.test(v)) {
      sh.getRange(r, 1, 1, 4).merge()
        .setBackground("#212121").setFontColor("white")
        .setFontWeight("bold").setFontSize(10)
        .setHorizontalAlignment("left");
    }
  }

  // Table header rows
  for (let r = 1; r <= rows.length; r++) {
    const a = String(sh.getRange(r, 1).getValue() || "").trim();
    if (["COLUMN", "SIGNAL VALUE", "FUNDAMENTAL VALUE", "DECISION VALUE", "RULE"].includes(a)) {
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

  ss.toast("REFERENCE_GUIDE updated (SELL states covered: Take Profit, Reduce Momentum Weak).", "✅ DONE", 3);
}

