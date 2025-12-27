/**
* ==============================================================================
* BASELINE LABEL: STABLE_MASTER_DEC25_BASE_v2_2 
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
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const a1 = range.getA1Notation();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) DASHBOARD manual update button (DASHBOARD!B1)
  if (sheet.getName() === "DASHBOARD" && a1 === "B1" && e.value === "TRUE") {
    ss.toast("Recalculating Signals...", "⚙️ SYSTEM", 3);
    try {
      generateCalculationsSheet();
      generateDashboardSheet();
      ss.toast("Terminal Synchronized.", "✅ DONE", 2);
    } catch (err) {
      sheet.getRange("B1").setValue(false);
      ss.toast("Error: " + err.toString(), "⚠️ FAIL", 5);
    }
    return;
  }

 //2) DASHBOARD refresh only (DASHBOARD!D1)
 if (sheet.getName() === "DASHBOARD" && a1 === "D1" && e.value === "TRUE") {
  ss.toast("Refreshing Dashboard...", "⚙️ SYSTEM", 2);
  try {
    generateDashboardSheet();
    sheet.getRange("D1").setValue(false);
    ss.toast("Dashboard Refreshed.", "✅ DONE", 2);
  } catch (err) {
    sheet.getRange("D1").setValue(false);
    ss.toast("Error: " + err.toString(), "⚠️ FAIL", 5);
  }
  return;
 }


  // 3) INPUT filters (INPUT!B1 / INPUT!C1) -> refresh dashboard
  if (sheet.getName() === "INPUT" && (a1 === "B1" || a1 === "C1")) {
    try {
      generateDashboardSheet();
      SpreadsheetApp.flush();
    } catch (err) {
      ss.toast("Dashboard filter refresh error: " + err.toString(), "⚠️ FAIL", 5);
    }
    return;
  }

  // 4) CHART controls -> update dynamic chart
  if (sheet.getName() === "CHART") {
    const watchList = ["B1", "D2", "A3", "B3", "C3"];
    if (watchList.includes(a1) || (range.getRow() === 1 && range.getColumn() <= 4)) {
      try {
        updateDynamicChart();
      } catch (err) {
        ss.toast("Chart refresh error: " + err.toString(), "⚠️ FAIL", 5);
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
    if (!highHist || !lowHist || !closeHist || !currentPrice) return 0;

    const H = highHist.flat().filter(n => typeof n === 'number' && n > 0);
    const L = lowHist.flat().filter(n => typeof n === 'number' && n > 0);
    const C = closeHist.flat().filter(n => typeof n === 'number' && n > 0);

    const n = Math.min(H.length, L.length, C.length);
    if (n < 40) return 0;

    const take = Math.min(n, 90);
    const h = H.slice(n - take);
    const l = L.slice(n - take);
    const c = C.slice(n - take);

    const liveClose = Number(currentPrice);
    c[c.length - 1] = liveClose;

    const period = 14;

    let tr = [];
    let pdm = [];
    let ndm = [];
    for (let i = 1; i < c.length; i++) {
      const upMove = h[i] - h[i - 1];
      const downMove = l[i - 1] - l[i];

      const plusDM = (upMove > downMove && upMove > 0) ? upMove : 0;
      const minusDM = (downMove > upMove && downMove > 0) ? downMove : 0;

      const trueRange = Math.max(
        h[i] - l[i],
        Math.abs(h[i] - c[i - 1]),
        Math.abs(l[i] - c[i - 1])
      );

      tr.push(trueRange);
      pdm.push(plusDM);
      ndm.push(minusDM);
    }

    if (tr.length < period * 2) return 0;

    let atr = tr.slice(0, period).reduce((a, b) => a + b, 0);
    let pDM14 = pdm.slice(0, period).reduce((a, b) => a + b, 0);
    let nDM14 = ndm.slice(0, period).reduce((a, b) => a + b, 0);

    const pDI0 = (atr === 0) ? 0 : (100 * (pDM14 / atr));
    const nDI0 = (atr === 0) ? 0 : (100 * (nDM14 / atr));
    let dxArr = [];
    dxArr.push((pDI0 + nDI0 === 0) ? 0 : (100 * Math.abs(pDI0 - nDI0) / (pDI0 + nDI0)));

    for (let i = period; i < tr.length; i++) {
      atr = atr - (atr / period) + tr[i];
      pDM14 = pDM14 - (pDM14 / period) + pdm[i];
      nDM14 = nDM14 - (nDM14 / period) + ndm[i];

      const pDI = (atr === 0) ? 0 : (100 * (pDM14 / atr));
      const nDI = (atr === 0) ? 0 : (100 * (nDM14 / atr));
      const dx = (pDI + nDI === 0) ? 0 : (100 * Math.abs(pDI - nDI) / (pDI + nDI));
      dxArr.push(dx);
    }

    if (dxArr.length < period) return 0;
    let adx = dxArr.slice(0, period).reduce((a, b) => a + b, 0) / period;
    for (let i = period; i < dxArr.length; i++) {
      adx = ((adx * (period - 1)) + dxArr[i]) / period;
    }
    return Number(adx.toFixed(2));
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
* 4. CALCULATION ENGINE (PERFORMANCE AFTER PRICE & VOLUME)
* ------------------------------------------------------------------
* A  Ticker
* B  Price
* C  Change %
* D  Volume (latest)
* E  Vol Trend (RVOL proxy)
* F  ATH (TRUE)
* G  ATH Diff %
* H  R:R Quality
* I  Divergence
* J  SMA 20
* K  SMA 50
* L  SMA 200
* M  Trend State
* N  RSI
* O  MACD Hist
* P  ADX (14)
* Q  Stoch %K (14)
* R  ATR (14)
* S  Bollinger %B
* T  Support
* U  Resistance
* V  Target (3:1)
* W  FUNDAMENTAL
* X  SIGNAL
* Y  DECISION
* Z  TECH_REASON
* AA FUND_REASON
* AB LAST_STATE
* ------------------------------------------------------------------
*/
function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet || !inputSheet) return;

  const tickers = getCleanTickers(inputSheet);
  let calcSheet = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");

  /**
   * =========================================================
   * CALCULATIONS (Industry Standard) FINAL MAP — A..AB (28 cols)
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
   * T  Stoch %K (14)  (0..1 shown as %)
   * U  Support
   * V  Resistance
   * W  Target (3:1)
   * X  ATR (14)
   * Y  Bollinger %B
   * Z  TECH NOTES
   * AA FUND NOTES
   * AB LAST_STATE
   * =========================================================
   */

  // ------------------------------------------------------------
  // Persist LAST_STATE (AB)
  // ------------------------------------------------------------
  const stateMap = {};
  if (calcSheet.getLastRow() >= 3) {
    const existing = calcSheet.getRange(3, 1, calcSheet.getLastRow() - 2, 28).getValues(); // A..AB
    existing.forEach(r => {
      const t = (r[0] || "").toString().trim().toUpperCase();
      if (t) stateMap[t] = r[27]; // AB
    });
  }

  calcSheet.clear().clearFormats();

  // Timestamp
  const syncTime = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  calcSheet.getRange("A1")
    .setValue(syncTime)
    .setFontSize(8)
    .setFontColor("#757575")
    .setFontStyle("italic");

  // ------------------------------------------------------------
  // Group headers (Row 1) — aligned to the new column order
  // ------------------------------------------------------------
  calcSheet.getRange("B1:C1").merge().setValue("[ ACTION ]")
    .setBackground("#263238").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  calcSheet.getRange("D1:F1").merge().setValue("[ CORE ]")
    .setBackground("#0D47A1").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  calcSheet.getRange("G1:J1").merge().setValue("[ PERFORMANCE ]")
    .setBackground("#1B5E20").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  calcSheet.getRange("K1:T1").merge().setValue("[ TREND & MOMENTUM ]")
    .setBackground("#424242").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  calcSheet.getRange("U1:Y1").merge().setValue("[ LEVELS & RISK ]")
    .setBackground("#B71C1C").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  calcSheet.getRange("Z1:AA1").merge().setValue("[ NOTES ]")
    .setBackground("#212121").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  calcSheet.getRange("AB1").setValue("[ STATE ]")
    .setBackground("#000000").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  // ------------------------------------------------------------
  // Row 2 headers (A..AB)
  // ------------------------------------------------------------
  const headers = [[
    "Ticker",
    "SIGNAL", "DECISION", "FUNDAMENTAL",
    "Price", "Change %",
    "Vol Trend", "ATH (TRUE)", "ATH Diff %", "R:R Quality",
    "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200",
    "RSI", "MACD Hist", "Divergence", "ADX (14)", "Stoch %K (14)",
    "Support", "Resistance", "Target (3:1)", "ATR (14)", "Bollinger %B",
    "TECH NOTES", "FUND NOTES",
    "LAST_STATE"
  ]];

  calcSheet.getRange(2, 1, 1, 28)
    .setValues(headers)
    .setBackground("#212121")
    .setFontColor("white")
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrap(true);

  // ------------------------------------------------------------
  // Build formulas per ticker (B..AA = 26 cols)
  // ------------------------------------------------------------
  const formulas = [];
  const restoredStates = [];

  tickers.forEach((ticker, i) => {
    const rowNum = i + 3;
    const t = ticker.toString().trim().toUpperCase();
    restoredStates.push([stateMap[t] || ""]);

    const tDS = (i * 7) + 1;

    // DATA block columns: date, open, high, low, close, volume
    const highCol  = columnToLetter(tDS + 2);
    const lowCol   = columnToLetter(tDS + 3);
    const closeCol = columnToLetter(tDS + 4);
    const volCol   = columnToLetter(tDS + 5);

    const lastRow = `COUNTA(DATA!$${closeCol}:$${closeCol})`;

    formulas.push([
      // B SIGNAL (full tech stack; references updated to new column letters)
      `=IF(OR(ISBLANK(E${rowNum}), E${rowNum}=0), "🔄 LOADING...",
        IFERROR(IFS(
          E${rowNum}<U${rowNum}, "STOP LOSS",
          E${rowNum}<O${rowNum}, "BEAR REGIME",
          E${rowNum}>=V${rowNum}*0.99, "RESISTANCE TEST",

          AND(G${rowNum}>1.5, E${rowNum}>M${rowNum}, Q${rowNum}>0, S${rowNum}>=18), "RVOL BREAKOUT",

          AND(T${rowNum}<0.2, E${rowNum}>U${rowNum}, S${rowNum}>=18), "STOCH OVERSOLD BOUNCE",
          AND(T${rowNum}>0.8, E${rowNum}>=V${rowNum}*0.97), "STOCH OVERBOUGHT FADE",

          AND(P${rowNum}<35, E${rowNum}>U${rowNum}), "RSI SUPPORT BOUNCE",

          AND(Y${rowNum}<0.2, Q${rowNum}>0, S${rowNum}<18), "VOL SQUEEZE (CHOP)",

          S${rowNum}<18, "CHOP (LOW ADX)",
          TRUE, "CHOP"
        ), "CHOP")
      )`,

      // C DECISION (D + B confluence + risk gates; updated refs)
      `=IF(B${rowNum}="🔄 LOADING...", "🔄 LOADING...",
        IFS(
          REGEXMATCH(B${rowNum}, "STOP"), "🛑 STOP OUT",

          D${rowNum}="💀 ZOMBIE", "💤 AVOID",
          REGEXMATCH(D${rowNum}, "BUBBLE"), "💤 AVOID",

          AND(B${rowNum}="RVOL BREAKOUT", D${rowNum}="💎 GEM (Value)", J${rowNum}>=1.5, S${rowNum}>=20), "💎 PRIME BUY",
          AND(B${rowNum}="RVOL BREAKOUT", J${rowNum}<1.1), "⚠️ POOR R:R (AVOID)",
          AND(B${rowNum}="RVOL BREAKOUT", G${rowNum}<1.2), "🎣 FAKE-OUT (NO VOL)",

          AND(E${rowNum}>M${rowNum}+(2*X${rowNum})), "⏳ ATR OVEREXTENDED",

          AND(B${rowNum}="STOCH OVERSOLD BOUNCE", E${rowNum}>O${rowNum}, S${rowNum}>=18), "🚀 TRADE (MEAN REV)",
          AND(B${rowNum}="RSI SUPPORT BOUNCE", E${rowNum}>O${rowNum}, S${rowNum}>=18), "🚀 TRADE",

          B${rowNum}="BEAR REGIME", "💤 AVOID",
          TRUE, "⏳ WAIT"
        )
      )`,

      // D FUNDAMENTAL
      `=IFERROR(LET(eps, GOOGLEFINANCE(A${rowNum}, "eps"), pe, GOOGLEFINANCE(A${rowNum}, "pe"),
        IFS(
          AND(eps>0, pe>0, pe<25), "💎 GEM (Value)",
          AND(eps>0, pe>50), "⚠️ PRICED PERF.",
          eps<0, "💀 ZOMBIE",
          AND(pe>30, eps<0.1), "🔴 BUBBLE",
          TRUE, "⚖️ FAIR"
        )
      ), "-")`,

      // E Price
      `=ROUND(IFERROR(GOOGLEFINANCE("${t}", "price")), 2)`,

      // F Change %
      `=IFERROR(GOOGLEFINANCE("${t}", "changepct")/100, 0)`,

      // G Vol Trend (RVOL proxy)
      `=ROUND(IFERROR(
        OFFSET(DATA!$${volCol}$4, ${lastRow}-1, 0) /
        AVERAGE(OFFSET(DATA!$${volCol}$4, ${lastRow}-21, 0, 20)),
      1), 2)`,

      // H ATH (TRUE)
      `=IFERROR(DATA!${columnToLetter(tDS + 1)}3, "-")`,

      // I ATH Diff %
      `=IFERROR((E${rowNum}-H${rowNum})/H${rowNum}, 0)`,

      // J R:R Quality
      `=IFERROR(ROUND((V${rowNum}-E${rowNum})/MAX(0.01, E${rowNum}-U${rowNum}), 2), 0)`,

      // K Trend Score
      `=REPT("★", (E${rowNum}>M${rowNum}) + (E${rowNum}>N${rowNum}) + (E${rowNum}>O${rowNum}))`,

      // L Trend State
      `=IF(E${rowNum}>O${rowNum}, "BULL REGIME", "BEAR REGIME")`,

      // M SMA 20
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)), 0), 2)`,

      // N SMA 50
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-50, 0, 50)), 0), 2)`,

      // O SMA 200
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-200, 0, 200)), 0), 2)`,

      // P RSI
      `=LIVERSI(DATA!$${closeCol}$4:$${closeCol}, E${rowNum})`,

      // Q MACD Hist
      `=LIVEMACD(DATA!$${closeCol}$4:$${closeCol}, E${rowNum})`,

      // R Divergence
      `=IFERROR(IFS(
        AND(E${rowNum}<INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), P${rowNum}>50), "BULLISH DIV",
        AND(E${rowNum}>INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), P${rowNum}<50), "BEARISH DIV",
        TRUE, "-"
      ), "-")`,

      // S ADX(14)
      `=LIVEADX(DATA!$${highCol}$4:$${highCol}, DATA!$${lowCol}$4:$${lowCol}, DATA!$${closeCol}$4:$${closeCol}, E${rowNum})`,

      // T Stoch %K(14) (0..1)
      `=LIVESTOCHK(DATA!$${highCol}$4:$${highCol}, DATA!$${lowCol}$4:$${lowCol}, DATA!$${closeCol}$4:$${closeCol}, E${rowNum})`,

      // U Support (20-day min low)
      `=ROUND(IFERROR(MIN(OFFSET(DATA!$${lowCol}$4, ${lastRow}-21, 0, 20)), E${rowNum}*0.9), 2)`,

      // V Resistance (50-day max high)
      `=ROUND(IFERROR(MAX(OFFSET(DATA!$${highCol}$4, ${lastRow}-51, 0, 50)), E${rowNum}*1.1), 2)`,

      // W Target (3:1)
      `=ROUND(E${rowNum} + ((E${rowNum}-U${rowNum}) * 3), 2)`,

      // X ATR (14) high-low proxy
      `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(
        OFFSET(DATA!$${highCol}$4, ${lastRow}-14, 0, 14) - OFFSET(DATA!$${lowCol}$4, ${lastRow}-14, 0, 14)
      )), 0), 2)`,

      // Y Bollinger %B proxy
      `=ROUND(IFERROR(((E${rowNum}-M${rowNum}) / (4*STDEV(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)))) + 0.5, 0.5), 2)`,

      // Z TECH NOTES (institutional narrative; updated refs)
      `=
"1) Volume Confirmation (Vol Trend / RVOL): " &
  TEXT(G${rowNum},"0.00") & "x — " &
  IF(G${rowNum}>=1.50,"High participation (institutional interest).",
    IF(G${rowNum}>=1.20,"Above-average participation (supportive).",
      IF(G${rowNum}>=0.80,"Normal participation (neutral).","Low participation (moves less reliable).")
    )
  ) & CHAR(10) &
"2) Regime (Price vs SMA200): Price " &
  TEXT(E${rowNum},"0.00") & " vs SMA200 " & TEXT(O${rowNum},"0.00") & " — " &
  IF(E${rowNum}>=O${rowNum},"Bull regime (above long-term trend).","Bear regime (below long-term trend).") & CHAR(10) &
"3) Volatility Stretch (ATR Envelope): SMA20 " & TEXT(M${rowNum},"0.00") &
  " | ATR(14) " & TEXT(X${rowNum},"0.00") &
  " | Band=[" & TEXT(M${rowNum}-2*X${rowNum},"0.00") & "…"
            & TEXT(M${rowNum}+2*X${rowNum},"0.00") & "]" &
  " | Price " & TEXT(E${rowNum},"0.00") & " — " &
  IF(E${rowNum} > M${rowNum} + 2*X${rowNum},"Overextended above +2×ATR (mean reversion risk).",
    IF(E${rowNum} < M${rowNum} - 2*X${rowNum},"Oversold below −2×ATR (capitulation / bounce risk).","Within ±2×ATR (normal volatility range).")
  ) & CHAR(10) &
"4) Momentum (RSI & MACD Histogram): RSI(14) " & TEXT(P${rowNum},"0.00") & " — " &
  IF(P${rowNum}>=70,"Overbought.",
    IF(P${rowNum}>=55,"Bullish.",
      IF(P${rowNum}>=45,"Neutral.",
        IF(P${rowNum}>=30,"Bearish.","Oversold.")
      )
    )
  ) &
  " | MACD Hist " & TEXT(Q${rowNum},"0.000") & " — " &
  IF(Q${rowNum}>0,"Above 0 (bullish impulse).","Below 0 (bearish impulse).") & CHAR(10) &
"5) Trend Quality (ADX & Stoch): ADX(14) " & TEXT(S${rowNum},"0.00") & " — " &
  IF(S${rowNum}>=25,"Strong trend.",
    IF(S${rowNum}>=18,"Emerging trend.","Range-bound / low trend.")
  ) &
  " | StochK " & TEXT(T${rowNum},"0.0000") & " — " &
  IF(T${rowNum}>=0.80,"Overbought.",
    IF(T${rowNum}<=0.20,"Oversold.","Neutral.")
  ) & CHAR(10) &
"6) Risk–Reward (Resistance/Support): R:R " & TEXT(J${rowNum},"0.00") & "x — " &
  IF(J${rowNum}>=3.00,"Institutional-grade (≥3:1).",
    IF(J${rowNum}>=2.00,"Tradable (≥2:1).",
      IF(J${rowNum}>=1.50,"Marginal.","Poor.")
    )
  ) &
  " | Support " & TEXT(U${rowNum},"0.00") &
  " | Resistance " & TEXT(V${rowNum},"0.00")
`,

      // AA FUND NOTES
      `="1) Valuation: " & D${rowNum} & ". P/E " & IFERROR(GOOGLEFINANCE(A${rowNum}, "pe"),"N/A") & " | EPS " & IFERROR(GOOGLEFINANCE(A${rowNum}, "eps"),"N/A") & "." & CHAR(10) &
        "2) Regime: " & IF(E${rowNum}>O${rowNum}, "Above SMA200 (long-term bullish).", "Below SMA200 (long-term bearish).") & CHAR(10) &
        "3) Trend Quality (ADX): " & TEXT(S${rowNum},"0.00") & " — " & IF(S${rowNum}>=25,"Strong trend supports continuation.","Low ADX implies range risk.") & CHAR(10) &
        "4) Verdict: " & C${rowNum} & ". Confluence of fundamentals + technicals."`
    ]);
  });

  // Write tickers
  calcSheet.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));

  // Write formulas (B..AA = 26 columns)
  if (tickers.length > 0) {
    calcSheet.getRange(3, 2, formulas.length, 26).setFormulas(formulas);
  }

  // Restore LAST_STATE to AB
  if (tickers.length > 0) {
    calcSheet.getRange(3, 28, restoredStates.length, 1).setValues(restoredStates);
  }

  // ------------------------------------------------------------
  // Number formats (aligned to new columns)
  // ------------------------------------------------------------
  calcSheet.getRange("F3:F").setNumberFormat("0.00%"); // Change %
  calcSheet.getRange("I3:I").setNumberFormat("0.00%"); // ATH Diff %
  calcSheet.getRange("T3:T").setNumberFormat("0.00%"); // Stoch %K (0..1)
  calcSheet.getRange("Y3:Y").setNumberFormat("0.00%"); // Boll %B (proxy 0..1)

  // ------------------------------------------------------------
  // Conditional formatting (aligned to new columns)
  // ------------------------------------------------------------
  const lastRowVal = Math.max(calcSheet.getLastRow(), 3);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor("#C62828").setBold(true)
      .setRanges([
        calcSheet.getRange("F3:F" + lastRowVal), // Change %
        calcSheet.getRange("I3:I" + lastRowVal), // ATH diff %
        calcSheet.getRange("Q3:Q" + lastRowVal)  // MACD hist
      ])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($P3>70, $P3<30)').setFontColor("#C62828").setBold(true)
      .setRanges([calcSheet.getRange("P3:P" + lastRowVal)])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$S3>=25').setFontColor("#2E7D32").setBold(true)
      .setRanges([calcSheet.getRange("S3:S" + lastRowVal)])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$T3<0.2').setFontColor("#2E7D32").setBold(true)
      .setRanges([calcSheet.getRange("T3:T" + lastRowVal)])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$T3>0.8').setFontColor("#C62828").setBold(true)
      .setRanges([calcSheet.getRange("T3:T" + lastRowVal)])
      .build()
  ];
  calcSheet.setConditionalFormatRules(rules);

  // ------------------------------------------------------------
  // Bloomberg-style density formatting (requested)
  // - compact rows
  // - CLIP everywhere including notes columns (per your request)
  // ------------------------------------------------------------
  const lastDataRow = Math.max(calcSheet.getLastRow(), 3);

  // Compact row height for data rows
  if (lastDataRow > 2) calcSheet.setRowHeights(3, lastDataRow - 2, 18);

  // Header heights
  calcSheet.setRowHeight(1, 22);
  calcSheet.setRowHeight(2, 26);

  // CLIP all data rows to enforce fixed height (including notes)
  calcSheet.getRange(`A3:AA${lastDataRow}`)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment("middle")
    .setFontSize(9);

  // Keep notes visually aligned (top looks better even when clipped)
  calcSheet.getRange(`Z3:AA${lastDataRow}`)
    .setVerticalAlignment("top")
    .setFontSize(9);

  // Optional: center tickers and action columns for a terminal look
  calcSheet.getRange(`A3:C${lastDataRow}`).setHorizontalAlignment("left");
  calcSheet.getRange(`D3:Y${lastDataRow}`).setHorizontalAlignment("left");

  SpreadsheetApp.flush();
  calcSheet.setFrozenRows(2);
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
  const calc  = ss.getSheetByName("CALCULATIONS");
  if (!input || !calc) return;

  const dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD");
  dashboard.clear().clearFormats();

  const TZ = ss.getSpreadsheetTimeZone();
  const norm = s => String(s || "").trim().toUpperCase();
  const splitTokens = s => String(s || "")
    .split(",")
    .map(x => x.trim().toUpperCase())
    .filter(Boolean);

  // Blank only these metrics when 0 / "-" / blank (per your earlier instruction)
  const blankIfZeroOrDash = (v) => {
    if (v === null || v === undefined) return "";
    const s = String(v).trim();
    if (s === "" || s === "-") return "";
    if (typeof v === "number") {
      if (!isFinite(v) || v === 0) return "";
      return v;
    }
    if (s === "0") return "";
    return v;
  };

  /* ============================================================
   * ROW 1 — CONTROL BAR (as requested)
   * ============================================================ */
  dashboard.getRange("A1").setValue("UPDATE CAL")
    .setBackground("#212121").setFontColor("white")
    .setFontWeight("bold").setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  dashboard.getRange("B1").insertCheckboxes()
    .setBackground("#212121").setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  dashboard.getRange("C1").setValue("UPDATE")
    .setBackground("#212121").setFontColor("white")
    .setFontWeight("bold").setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  dashboard.getRange("D1").insertCheckboxes()
    .setBackground("#212121").setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  dashboard.getRange("E1:G1").merge()
    .setBackground("#000000")
    .setFontColor("#00FF00")
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setValue(Utilities.formatDate(new Date(), TZ, "MMM dd, yyyy | HH:mm:ss"));

  /* ============================================================
   * ROW 2 — GROUP HEADERS (wrap enabled)
   * ============================================================ */
  dashboard.setRowHeight(2, 28);

  const groups = [
    ["A2:A2",  "IDENTITY",        "#263238"],
    ["B2:D2",  "SIGNALS",         "#4A148C"],
    ["E2:G2",  "PRICE & VOLUME",  "#0D47A1"],
    ["H2:J2",  "PERFORMANCE",     "#1B5E20"],
    ["K2:O2",  "TREND",           "#004D40"],
    ["P2:T2",  "MOMENTUM",        "#33691E"],
    ["U2:Y2",  "LEVELS",          "#B71C1C"],
    ["Z2:AA2", "ANALYSIS",        "#212121"]
  ];

  groups.forEach(([rng, label, bg]) => {
    dashboard.getRange(rng).merge()
      .setValue(label)
      .setBackground(bg)
      .setFontColor("white")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  });

  dashboard.getRange("A2:AA2")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment("middle");

  /* ============================================================
   * CALCULATIONS HEADER MAP + FULL ALIAS RESOLVER
   * ============================================================ */
  const calcLastCol = calc.getLastColumn();
  const calcHeaderRow = calc.getRange(2, 1, 1, calcLastCol).getValues()[0];
  const calcHeaders = calcHeaderRow.map(norm);

  const idx = {};
  calcHeaders.forEach((h, i) => { if (h) idx[h] = i; });

  // Aliases for industry-standard drift + your older variants
  const aliases = {
    "TICKER": ["TICKER"],

    "SIGNAL": ["SIGNAL", "SIGNAL (TECH ENGINE)", "SIGNAL (RAW)"],
    "DECISION": ["DECISION", "DECISION (FINAL)"],
    "FUNDAMENTAL": ["FUNDAMENTAL", "FUNDAMENTALS"],

    "PRICE": ["PRICE", "LIVE PRICE"],
    "CHANGE %": ["CHANGE %", "CHANGE%"],

    "VOL TREND": ["VOL TREND", "RVOL", "RELATIVE VOLUME", "VOLUME TREND"],
    "ATH (TRUE)": ["ATH (TRUE)", "ATH"],
    "ATH DIFF %": ["ATH DIFF %", "DIFF FROM ATH %", "% FROM ATH"],
    "R:R QUALITY": ["R:R QUALITY", "RR QUALITY", "RR", "R:R"],

    "TREND STATE": ["TREND STATE", "REGIME", "MARKET REGIME"],

    "SMA 20": ["SMA 20", "SMA20"],
    "SMA 50": ["SMA 50", "SMA50"],
    "SMA 200": ["SMA 200", "SMA200"],

    "RSI": ["RSI", "RSI (14)"],
    "MACD HIST": ["MACD HIST", "MACD HISTOGRAM", "MACD HIST."],
    "DIVERGENCE": ["DIVERGENCE"],

    "ADX (14)": ["ADX (14)", "ADX", "ADX14"],
    "STOCH %K (14)": ["STOCH %K (14)", "STOCH %K", "STOCH K", "STOCHK"],

    "SUPPORT": ["SUPPORT", "SUPPORT FLOOR"],
    "RESISTANCE": ["RESISTANCE", "RESISTANCE CEILING"],
    "TARGET (3:1)": ["TARGET (3:1)", "TARGET (3R)", "TARGET (3:1 R:R)"],

    "ATR (14)": ["ATR (14)", "ATR", "ATR14"],
    "BOLLINGER %B": ["BOLLINGER %B", "BB %B", "%B"],

    "TECH_REASON": ["TECH_REASON", "TECH ANALYSIS", "TECH NOTES"],
    "FUND_REASON": ["FUND_REASON", "FUND ANALYSIS", "FUND NOTES"]
  };

  const resolve = (want) => {
    const key = norm(want);
    if (idx[key] !== undefined) return key;
    const list = aliases[key] || [];
    for (const cand of list) {
      const k = norm(cand);
      if (idx[k] !== undefined) return k;
    }
    return null;
  };

  const get = (row, want) => {
    const k = resolve(want);
    return (k && idx[k] !== undefined) ? row[idx[k]] : "";
  };

  // Dashboard column order (logical names)
  const dashFields = [
    "Ticker","SIGNAL","DECISION","FUNDAMENTAL","Price","Change %",
    "Vol Trend","ATH (TRUE)","ATH Diff %","R:R Quality",
    "Trend Score","Trend State",
    "SMA 20","SMA 50","SMA 200",
    "RSI","MACD Hist","Divergence",
    "ADX (14)","Stoch %K (14)",
    "Support","Resistance","Target (3:1)",
    "ATR (14)","Bollinger %B",
    "TECH_REASON","FUND_REASON"
  ];

  // Validate required fields
  const required = dashFields.filter(f => f !== "Trend Score"); // Trend Score is computed
  const missing = required.filter(f => !resolve(f));
  if (missing.length) {
    dashboard.getRange("A4")
      .setValue("CALCULATIONS header mismatch. Missing: " + missing.join(", "))
      .setFontColor("#C62828")
      .setFontWeight("bold");
    return;
  }

  /* ============================================================
   * ROW 3 — COLUMN HEADERS (NOW MATCH CALCULATIONS ACTUAL NAMES)
   * ============================================================ */
  const headerDisplay = dashFields.map(f => {
    if (f === "Trend Score") return "Trend Score"; // computed column label
    const resolvedKey = resolve(f);                // normalized calc header key
    const pos = idx[resolvedKey];
    return calcHeaderRow[pos] || f;                // use exact CALCULATIONS header text
  });

  dashboard.getRange(3, 1, 1, headerDisplay.length).setValues([headerDisplay])
    .setBackground("#212121")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // FIX 1: Row 3 wrap (your request)
  dashboard.setRowHeight(3, 34);
  dashboard.getRange("A3:AA3")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment("middle");

  /* ============================================================
   * WHITE BORDERS — ROWS 1, 2, 3
   * ============================================================ */
  ["A1:AA1", "A2:AA2", "A3:AA3"].forEach(rng => {
    dashboard.getRange(rng).setBorder(
      true, true, true, true, true, true,
      "#FFFFFF",
      SpreadsheetApp.BorderStyle.SOLID
    );
  });

  /* ============================================================
   * INPUT FILTERS (B1/C1 token match; ALL disables)
   * ============================================================ */
  const wantB = splitTokens(input.getRange("B1").getValue());
  const wantC = splitTokens(input.getRange("C1").getValue());
  const bAll = wantB.length === 0 || wantB.includes("ALL");
  const cAll = wantC.length === 0 || wantC.includes("ALL");

  const inputRows = input.getLastRow() >= 3
    ? input.getRange(3, 1, input.getLastRow() - 2, 3).getValues()
    : [];

  const allowed = new Set();
  inputRows.forEach(r => {
    const t = norm(r[0]);
    if (!t) return;

    const sectorTokens = splitTokens(r[1]);
    const industryTokens = splitTokens(r[2]);

    const okB = bAll || sectorTokens.some(x => wantB.includes(x));
    const okC = cAll || industryTokens.some(x => wantC.includes(x));

    if (okB && okC) allowed.add(t);
  });

  /* ============================================================
   * BUILD DASHBOARD DATA
   * ============================================================ */
  const calcRows = calc.getLastRow() >= 3
    ? calc.getRange(3, 1, calc.getLastRow() - 2, calcLastCol).getValues()
    : [];

  const out = [];

  calcRows.forEach(r => {
    const t = norm(get(r, "Ticker"));
    if (!t || !allowed.has(t)) return;

    const price = Number(get(r, "Price")) || 0;
    const s20 = Number(get(r, "SMA 20")) || 0;
    const s50 = Number(get(r, "SMA 50")) || 0;
    const s200 = Number(get(r, "SMA 200")) || 0;
    const trendScore = "★".repeat((price > s20) + (price > s50) + (price > s200));

    out.push([
      get(r,"Ticker"),
      get(r,"SIGNAL"),
      get(r,"DECISION"),
      get(r,"FUNDAMENTAL"),
      get(r,"Price"),
      get(r,"Change %"),

      // blanking only for the metrics you listed earlier
      blankIfZeroOrDash(get(r,"Vol Trend")),
      blankIfZeroOrDash(get(r,"ATH (TRUE)")),
      blankIfZeroOrDash(get(r,"ATH Diff %")),
      blankIfZeroOrDash(get(r,"R:R Quality")),

      trendScore,
      get(r,"Trend State"),

      get(r,"SMA 20"),
      get(r,"SMA 50"),
      get(r,"SMA 200"),

      blankIfZeroOrDash(get(r,"RSI")),
      blankIfZeroOrDash(get(r,"MACD Hist")),

      get(r,"Divergence"),
      get(r,"ADX (14)"),
      get(r,"Stoch %K (14)"),

      get(r,"Support"),
      get(r,"Resistance"),
      get(r,"Target (3:1)"),

      get(r,"ATR (14)"),
      get(r,"Bollinger %B"),

      get(r,"TECH_REASON"),
      get(r,"FUND_REASON")
    ]);
  });

  // Sort by Change % desc (index 5)
  out.sort((a, b) => (Number(b[5]) || 0) - (Number(a[5]) || 0));

  if (out.length) {
    dashboard.getRange(4, 1, out.length, dashFields.length).setValues(out);
  } else {
    dashboard.getRange("A4")
      .setValue("No Matches Found")
      .setFontColor("#9E9E9E");
  }

  /* ============================================================
   * FORMATTING GOVERNANCE (restored)
   * ============================================================ */
  dashboard.setFrozenRows(3);
  dashboard.setFrozenColumns(1);

  for (let col = 1; col <= 25; col++) dashboard.setColumnWidth(col, 75);
  dashboard.setColumnWidth(26, 350);
  dashboard.setColumnWidth(27, 350);

  const lastRow = Math.max(dashboard.getLastRow(), 4);
  const rows = Math.max(0, lastRow - 3);

  if (rows > 0) {
    dashboard.setRowHeights(4, rows, 28);

    dashboard.getRange(4, 1, rows, 25).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    dashboard.getRange(4, 26, rows, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    const dataRange = dashboard.getRange(4, 1, rows, dashFields.length);
    dataRange.setHorizontalAlignment("left").setVerticalAlignment("middle");
    dataRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

    // Percent formats based on *position* (stable)
    // Change % = col F, ATH Diff % = col I, Stoch %K = col T, Bollinger %B = col Y
    dashboard.getRange("F4:F" + lastRow).setNumberFormat("0.00%");
    dashboard.getRange("I4:I" + lastRow).setNumberFormat("0.00%");
    dashboard.getRange("T4:T" + lastRow).setNumberFormat("0.00%");
    dashboard.getRange("Y4:Y" + lastRow).setNumberFormat("0.00%");
  }

  /* ============================================================
   * CONDITIONAL FORMATTING (restored)
   * ============================================================ */
  const rules = [];

  // Negative: Change % (F), ATH Diff % (I), MACD Hist (Q)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([
        dashboard.getRange("F4:F" + lastRow),
        dashboard.getRange("I4:I" + lastRow),
        dashboard.getRange("Q4:Q" + lastRow)
      ])
      .build()
  );

  // RSI extremes (P)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($P4>70,$P4<30)')
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([dashboard.getRange("P4:P" + lastRow)])
      .build()
  );

  // ADX strong trend (S)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$S4>=25')
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("S4:S" + lastRow)])
      .build()
  );

  // Stoch zones (T)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$T4<0.2')
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("T4:T" + lastRow)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$T4>0.8')
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([dashboard.getRange("T4:T" + lastRow)])
      .build()
  );

  // Signal+Decision heat (B:C)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($B4&" "&$C4, "(?i)PRIME|TRADE|BREAKOUT|BOUNCE")')
      .setBackground("#E8F5E9")
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("B4:C" + lastRow)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($B4&" "&$C4, "(?i)FAKE-OUT|OVEREXTENDED")')
      .setBackground("#FFF3E0")
      .setFontColor("#E65100")
      .setBold(true)
      .setRanges([dashboard.getRange("B4:C" + lastRow)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($B4&" "&$C4, "(?i)STOP|AVOID|BEAR")')
      .setBackground("#FFEBEE")
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([dashboard.getRange("B4:C" + lastRow)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($B4&" "&$C4, "(?i)CHOP|WAIT|LOADING")')
      .setBackground("#F5F5F5")
      .setFontColor("#9E9E9E")
      .setRanges([dashboard.getRange("B4:C" + lastRow)])
      .build()
  );

  // Fundamental GEM highlight (D)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("GEM")
      .setBackground("#E8F5E9")
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("D4:D" + lastRow)])
      .build()
  );

  // Trend State highlight (L)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("BULL")
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("L4:L" + lastRow)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("BEAR")
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([dashboard.getRange("L4:L" + lastRow)])
      .build()
  );

  dashboard.setConditionalFormatRules(rules);
  SpreadsheetApp.flush();
}


/**
* ------------------------------------------------------------------
* 6. SETUP CHART SHEET (indices updated to latest CALCULATIONS map)
* ------------------------------------------------------------------
*/
function setupChartSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  const calcSheet = ss.getSheetByName("CALCULATIONS");
  if (!inputSheet || !calcSheet) return;

  const tickers = getCleanTickers(inputSheet);
  const chartSheet = ss.getSheetByName("CHART") || ss.insertSheet("CHART");

  chartSheet.clear().clearFormats();
  forceExpandSheet(chartSheet, 60);

  // ------------------------------------------------------------
  // Layout
  // ------------------------------------------------------------
  chartSheet.setColumnWidth(1, 180); // A
  chartSheet.setColumnWidth(2, 120); // B
  chartSheet.setColumnWidth(3, 120); // C
  chartSheet.setColumnWidth(4, 120); // D
  chartSheet.setColumnWidth(5, 125); // E
  chartSheet.setColumnWidth(6, 125); // F
  chartSheet.setColumnWidth(7, 125); // G
  chartSheet.setColumnWidth(8, 125); // H

  const headerRange = chartSheet.getRange("A1:H4");
  headerRange
    .setBackground("#000000")
    .setFontColor("#FFFF00")
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  chartSheet.getRange("A1").setValue("TICKER:").setFontWeight("bold");

  chartSheet.getRange("B1:D1")
    .merge()
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(tickers.length ? tickers : [""])
        .build()
    )
    .setValue(tickers.length ? tickers[0] : "")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(12)
    .setFontColor("#FF80AB");

  // ------------------------------------------------------------
  // Helper: Alias-aware MATCH() chain inside a Sheets formula
  // ------------------------------------------------------------
  const matchAny = (names) => {
    // Produces: IFERROR(MATCH("A",CALCULATIONS!$A$2:$ZZ$2,0),IFERROR(MATCH("B",...),MATCH("C",...)))
    const quoted = names.map(n => `"${n}"`);
    if (quoted.length === 1) return `MATCH(${quoted[0]}, CALCULATIONS!$A$2:$ZZ$2, 0)`;
    let expr = `MATCH(${quoted[quoted.length - 1]}, CALCULATIONS!$A$2:$ZZ$2, 0)`;
    for (let i = quoted.length - 2; i >= 0; i--) {
      expr = `IFERROR(MATCH(${quoted[i]}, CALCULATIONS!$A$2:$ZZ$2, 0), ${expr})`;
    }
    return expr;
  };

  const V = (aliases) => {
    // Alias-aware header lookup: VLOOKUP(ticker, A3:ZZ, matchAny([...]), 0)
    return `=IFERROR(VLOOKUP($B$1, CALCULATIONS!$A$3:$ZZ, ${matchAny(aliases)}, 0), "—")`;
  };

  const Vnum = (aliases) => {
    return `=IFERROR(VLOOKUP($B$1, CALCULATIONS!$A$3:$ZZ, ${matchAny(aliases)}, 0), 0)`;
  };

  // ------------------------------------------------------------
  // Reasoning boxes (TECH / FUND) – now alias-safe + width-safe
  // ------------------------------------------------------------
  chartSheet.getRange("E1:F4").merge()
    .setWrap(true).setVerticalAlignment("top").setFontSize(10)
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  chartSheet.getRange("G1:H4").merge()
    .setWrap(true).setVerticalAlignment("top").setFontSize(10)
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  chartSheet.getRange("E1").setFormula(
    V(["TECH_REASON", "Tech Notes", "TECH ANALYSIS", "TECH_ANALYSIS"])
  );
  chartSheet.getRange("G1").setFormula(
    V(["FUND_REASON", "Fund Notes", "FUND ANALYSIS", "FUND_ANALYSIS"])
  );

  // ------------------------------------------------------------
  // Date controls
  // ------------------------------------------------------------
  chartSheet.getRange("A2:C2").setValues([["YEAR", "MONTH", "DAY"]])
    .setFontWeight("bold").setHorizontalAlignment("center");

  const numRule = (max) =>
    SpreadsheetApp.newDataValidation()
      .requireValueInList(Array.from({ length: max + 1 }, (_, i) => i))
      .build();

  chartSheet.getRange("A3").setDataValidation(numRule(5)).setValue(1)
    .setHorizontalAlignment("center").setFontColor("#FF80AB");
  chartSheet.getRange("B3").setDataValidation(numRule(12)).setValue(0)
    .setHorizontalAlignment("center").setFontColor("#FF80AB");
  chartSheet.getRange("C3").setDataValidation(numRule(31)).setValue(0)
    .setHorizontalAlignment("center").setFontColor("#FF80AB");

  chartSheet.getRange("D2")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"]).build())
    .setValue("DAILY")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontColor("#FF80AB");

  chartSheet.getRange("A4").setValue("DATE").setFontWeight("bold");
  chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)")
    .setNumberFormat("yyyy-mm-dd");

  // ------------------------------------------------------------
  // Sidebar (A5:B...) — RENAMED labels + alias-safe pulls
  // ------------------------------------------------------------
  const t = "$B$1";

  const data = [
    ["SIGNAL", V(["SIGNAL", "SIGNAL (TECH ENGINE)", "SIGNAL (RAW)"])],
    ["DECISION", V(["DECISION", "DECISION (FINAL)"])],
    ["FUNDAMENTAL", V(["FUNDAMENTAL", "FUNDAMENTALS"])],
    ["PRICE", `=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`],
    ["CHANGE ($)", `=IFERROR(B8 - GOOGLEFINANCE(${t}, "closeyest"), 0)`],
    ["CHANGE (%)", `=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`],
    ["", ""],

    ["[ PERFORMANCE ]", ""],
    ["VOL TREND", Vnum(["VOL TREND", "RVOL", "RELATIVE VOLUME", "VOLUME TREND"])],
    ["ATH (TRUE)", Vnum(["ATH (TRUE)", "ATH"])],
    ["ATH DIFF %", Vnum(["ATH DIFF %", "DIFF FROM ATH %", "% FROM ATH"])],
    ["R:R QUALITY", Vnum(["R:R QUALITY", "RR QUALITY", "RR", "R:R"])],
    ["", ""],

    ["[ TREND ]", ""],
    ["TREND STATE", V(["TREND STATE", "Trend State", "REGIME", "MARKET REGIME"])],
    ["SMA 20", Vnum(["SMA 20", "SMA20"])],
    ["SMA 50", Vnum(["SMA 50", "SMA50"])],
    ["SMA 200", Vnum(["SMA 200", "SMA200"])],
    ["", ""],

    ["[ MOMENTUM ]", ""],
    ["RSI", Vnum(["RSI", "RSI (14)"])],
    ["MACD HIST", Vnum(["MACD HIST", "MACD Hist", "MACD HISTOGRAM", "MACD HIST."])],
    ["DIVERGENCE", V(["DIVERGENCE"])],
    ["ADX (14)", Vnum(["ADX (14)", "ADX", "ADX14"])],
    ["STOCH %K (14)", Vnum(["STOCH %K (14)", "STOCH %K", "STOCH K", "STOCHK"])],
    ["", ""],

    ["[ LEVELS ]", ""],
    ["SUPPORT", Vnum(["SUPPORT", "SUPPORT FLOOR"])],
    ["RESISTANCE", Vnum(["RESISTANCE", "RESISTANCE CEILING"])],
    ["TARGET (3:1)", Vnum(["TARGET (3:1)", "TARGET (3R)", "TARGET (3:1 R:R)", "TARGET (3:1 R:R)"])],
    ["ATR (14)", Vnum(["ATR (14)", "ATR", "ATR14"])],
    ["BOLLINGER %B", Vnum(["BOLLINGER %B", "BB %B", "%B"])]
  ];

  const startRow = 5;

  // Write sidebar labels + formulas
  chartSheet.getRange(startRow, 1, data.length, 1)
    .setValues(data.map(r => [r[0]]))
    .setFontWeight("bold");

  chartSheet.getRange(startRow, 2, data.length, 1)
    .setFormulas(data.map(r => [r[1]]));

  // Style section headers
  data.forEach((r, i) => {
    const label = String(r[0] || "");
    if (label.startsWith("[")) {
      chartSheet.getRange(startRow + i, 1, 1, 2)
        .setBackground("#444")
        .setFontColor("white")
        .setHorizontalAlignment("center")
        .setFontWeight("bold");
    }
  });

  SpreadsheetApp.flush();

  // Align values column
  chartSheet.getRange(`B${startRow}:B${startRow + data.length - 1}`).setHorizontalAlignment("left");

  // Formats (safe; does not break if blank)
  chartSheet.getRangeList(["B8", "B9", "B15", "B17:B21", "B29:B33"]).setNumberFormat("#,##0.00");
  chartSheet.getRangeList(["B10", "B11", "B6"]).setNumberFormat("0.00%"); // Change %, ATH diff %, Change %

  // Conditional formatting (lightweight)
  const rules = [];

  // Negative change red (CHANGE $ and CHANGE %)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#D32F2F")
      .setRanges([chartSheet.getRange("B9:B10")])
      .build()
  );

  // RSI extremes (apply to full sidebar value column)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($B5>70,$B5<30)')
      .setFontColor("#D32F2F")
      .setRanges([chartSheet.getRange(`B${startRow}:B${startRow + data.length - 1}`)])
      .build()
  );

  chartSheet.setConditionalFormatRules(rules);

  // Build initial chart (assumes you keep the single corrected updateDynamicChart())
  updateDynamicChart();
}



/**
* ------------------------------------------------------------------
* Update chart using ONLY the CHART sidebar values (label-based)
* ------------------------------------------------------------------
*/
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;

  SpreadsheetApp.flush();

  // Timestamp
  sheet.getRange("E5")
    .setValue("Updated: " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "HH:mm:ss"))
    .setFontColor("gray")
    .setFontSize(8)
    .setHorizontalAlignment("right");

  const ticker = String(sheet.getRange("B1").getValue() || "").trim();
  if (!ticker) return;

  // Controls
  const isWeekly = sheet.getRange("D2").getValue() === "WEEKLY";
  const years = Number(sheet.getRange("A3").getValue()) || 0;
  const months = Number(sheet.getRange("B3").getValue()) || 0;
  const days = Number(sheet.getRange("C3").getValue()) || 0;

  const now = new Date();
  let startDate = new Date(now.getFullYear() - years, now.getMonth() - months, now.getDate() - days);
  if ((now - startDate) < (7 * 24 * 60 * 60 * 1000)) {
    startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 14);
  }

  // --- Sidebar values (label-based, alias-safe) ---
  // PRICE label could be "PRICE" or "LIVE PRICE" depending on your sidebar
  const priceMap = getSidebarValuesByLabels_(sheet, ["PRICE", "LIVE PRICE"]);
  let livePrice = Number(priceMap["PRICE"] || priceMap["LIVE PRICE"]) || 0;

  const levels = getSidebarLevels_(sheet);
  const supportVal = Number(levels.support) || 0;
  const resistanceVal = Number(levels.resistance) || 0;

  // --- Find ticker block in DATA (row 2 contains ticker names per block) ---
  const rawHeaders = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIdx = rawHeaders.indexOf(ticker);
  if (colIdx === -1) {
    sheet.getRange("E1").setValue("⚠️ Ticker Not Found");
    return;
  }

  // Pull 6 cols: date, open, high, low, close, volume
  const rawData = dataSheet.getRange(1, colIdx + 1, dataSheet.getLastRow(), 6).getValues();

  let masterData = [];
  let viewVols = [];
  let prices = [];

  for (let i = 4; i < rawData.length; i++) {
    const row = rawData[i];
    const d = row[0];
    const close = Number(row[4]);
    const vol = Number(row[5]);

    if (!d || !(d instanceof Date) || !isFinite(close) || close < 0.01) continue;
    if (d < startDate) continue;
    if (isWeekly && d.getDay() !== 5) continue; // Fridays

    const slice = rawData
      .slice(Math.max(4, i - 200), i + 1)
      .map(r => Number(r[4]))
      .filter(n => isFinite(n) && n > 0);

    const s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
    const s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
    const s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;

    const prevClose = (i > 4) ? Number(rawData[i - 1][4]) : close;

    // IMPORTANT: Z-section order is:
    // Date, Price, BullVol, BearVol, SMA20, SMA50, SMA200, Resistance, Support
    masterData.push([
      Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "MMM dd"),
      close,
      (close >= prevClose) ? vol : null,
      (close < prevClose) ? vol : null,
      s20, s50, s200,
      resistanceVal || null,
      supportVal || null
    ]);

    viewVols.push(vol);
    prices.push(close);
    if (s20) prices.push(s20);
    if (s50) prices.push(s50);
    if (s200) prices.push(s200);
  }

  // Live price fallback
  let candleLabel = "🔴 LIVE";
  if ((!livePrice || !isFinite(livePrice)) && prices.length > 0) {
    livePrice = prices[prices.length - 1];
    candleLabel = "⏳ SYNCING";
  }

  // Add live candle + live SMAs
  if (livePrice && isFinite(livePrice) && livePrice > 0) {
    const allCloses = rawData.slice(4).map(r => Number(r[4])).filter(n => isFinite(n) && n > 0);

    const sma20Arr = allCloses.slice(-19).concat([livePrice]);
    const sma50Arr = allCloses.slice(-49).concat([livePrice]);
    const sma200Arr = allCloses.slice(-199).concat([livePrice]);

    const liveS20 = sma20Arr.length >= 20 ? sma20Arr.reduce((a, b) => a + b, 0) / 20 : null;
    const liveS50 = sma50Arr.length >= 50 ? sma50Arr.reduce((a, b) => a + b, 0) / 50 : null;
    const liveS200 = sma200Arr.length >= 200 ? sma200Arr.reduce((a, b) => a + b, 0) / 200 : null;

    masterData.push([candleLabel, livePrice, null, null, liveS20, liveS50, liveS200, resistanceVal || null, supportVal || null]);
    prices.push(livePrice);
  }

  // Output area (Z..AH, col 26..34)
  sheet.getRange(3, 26, 2000, 9).clearContent();
  if (masterData.length === 0) return;

  if (supportVal > 0) prices.push(supportVal);
  if (resistanceVal > 0) prices.push(resistanceVal);

  const cleanPrices = prices.filter(p => typeof p === "number" && isFinite(p) && p > 0);
  if (cleanPrices.length === 0) return;

  const minP = Math.min(...cleanPrices) * 0.98;
  const maxP = Math.max(...cleanPrices) * 1.02;

  const cleanVols = viewVols.filter(v => typeof v === "number" && isFinite(v) && v >= 0);
  const maxVol = Math.max(...cleanVols, 1);

  sheet.getRange(2, 26, 1, 9)
    .setValues([["Date", "Price", "Bull Vol", "Bear Vol", "SMA 20", "SMA 50", "SMA 200", "Resistance", "Support"]])
    .setFontWeight("bold")
    .setFontColor("white");

  sheet.getRange(3, 26, masterData.length, 9).setValues(masterData);
  SpreadsheetApp.flush();

  // Rebuild chart
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(2, 26, masterData.length + 1, 9))
    .setOption("useFirstRowAsHeaders", true)
    .setOption("series", {
      0: { type: "line", lineWidth: 3, labelInLegend: "Price" },
      1: { type: "bars", targetAxisIndex: 1, labelInLegend: "Bull Vol" },
      2: { type: "bars", targetAxisIndex: 1, labelInLegend: "Bear Vol" },
      3: { type: "line", lineWidth: 1.5, labelInLegend: "SMA 20" },
      4: { type: "line", lineWidth: 1.5, labelInLegend: "SMA 50" },
      5: { type: "line", lineWidth: 2, labelInLegend: "SMA 200" },
      6: { type: "line", lineDashStyle: [4, 4], labelInLegend: "Resistance" },
      7: { type: "line", lineDashStyle: [4, 4], labelInLegend: "Support" }
    })
    .setOption("vAxes", {
      0: { viewWindow: { min: minP, max: maxP } },
      1: { viewWindow: { min: 0, max: maxVol * 4 }, format: "short" }
    })
    .setOption("legend", { position: "top", textStyle: { fontSize: 10 } })
    .setPosition(5, 3, 0, 0)
    .setOption("width", 1150)
    .setOption("height", 650)
    .build();

  sheet.insertChart(chart);
}

/**
* Reads CHART sidebar values by labels from A5:B120.
*/
function getSidebarValuesByLabels_(chartSheet, labels) {
  const keys = chartSheet.getRange("A5:A120").getValues().flat()
    .map(v => String(v || "").trim().toUpperCase());
  const vals = chartSheet.getRange("B5:B120").getValues().flat();

  const out = {};
  labels.forEach(lbl => out[lbl] = 0);

  for (let i = 0; i < keys.length; i++) {
    labels.forEach(lbl => {
      const want = String(lbl).trim().toUpperCase();
      if (keys[i] === want) out[lbl] = vals[i];
    });
  }
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
  SpreadsheetApp.getUi().alert('🔔 MONITOR ACTIVE', 'Checking signals every 30 mins. You will only be emailed when a signal CHANGES.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function stopMarketMonitor() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'checkSignalsAndSendAlerts') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getUi().alert('🔕 MONITOR STOPPED', 'Automated checks disabled.', SpreadsheetApp.getUi().ButtonSet.OK);
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

  let newAlerts = [];

  data.forEach((row, idx) => {
    const ticker = row[0];
    const currentDecision = row[24]; // Y -> index 24
    const lastNotified = row[27];    // AB -> index 27

    if (!ticker || !currentDecision) return;

    if (currentDecision !== lastNotified) {
      const isActionable = /(PRIME|TRADE|STOP|BOUNCE|BREAKOUT)/i.test(currentDecision);
      if (isActionable) {
        newAlerts.push(`📍 TICKER: ${ticker}\n   NEW SIGNAL: ${currentDecision}\n   PREVIOUS: ${lastNotified || 'Initial Scan'}`);
      }
      calcSheet.getRange(idx + 3, 28).setValue(currentDecision);
    }
  });

  if (newAlerts.length > 0) {
    const email = Session.getActiveUser().getEmail();
    const subject = `📈 Terminal Alert: ${newAlerts.length} New Signal Change(s)`;
    const body =
      "The Institutional Terminal has detected signal turns in the CALCULATIONS engine:\n\n" +
      newAlerts.join("\n\n") +
      "\n\nView Terminal: " + ss.getUrl();

    MailApp.sendEmail(email, subject, body);
  }
}


/**
* Creates a conceptual REFERENCE_GUIDE for the Terminal.
*/
function generateReferenceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("CONCEPT_GUIDE") || ss.insertSheet("CONCEPT_GUIDE");
  sheet.clear().clearFormats();

  const data = [
    ["INSTITUTIONAL TERMINAL: CONCEPT & PHILOSOPHY", "", ""],
    ["Pillar 1: FUNDAMENTALS", "Pillar 2: TECHNICALS", "Pillar 3: DECISION"],
    ["Focus: Business Quality", "Focus: Price Timing", "Focus: Final Action"],
    ["Classifies value/quality via EPS + P/E heuristics.", "Momentum + trend filters (RSI/MACD/ADX/Stoch).", "Confluence of Quality + Timing + Risk Gates."],
    ["", "", ""],
    ["CORE INDICATOR GLOSSARY", "", ""],
    ["TERM", "MEANING", "ACTION TRIGGER"],
    ["LIVERSI", "Measures exhaustion/mean-reversion pressure.", ">70 Overbought risk; <30 Oversold opportunity"],
    ["LIVEMACD", "Trend pulse / momentum bias.", "Positive & rising supports continuation"],
    ["ADX (14)", "Trend strength filter.", "<25 = choppy; >=25 = trending"],
    ["ATR (14)", "Noise/volatility proxy.", "Used to avoid overextension + set risk bands"],
    ["R:R Quality", "Reward vs risk math.", "Prefer >= 2.0; institutional edge >= 3.0"]
  ];

  sheet.getRange(1, 1, data.length, 3).setValues(data);

  sheet.getRange("A1:C1").merge().setBackground("#0b5394").setFontColor("white").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:C3").setBackground("#cfe2f3").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A6:C6").merge().setBackground("#212121").setFontColor("white").setFontWeight("bold");
  sheet.getRange("A7:C7").setBackground("#f3f3f3").setFontWeight("bold");
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 400);
  sheet.getRange("A1:C" + data.length).setWrap(true).setVerticalAlignment("middle");

  ss.toast("Conceptual Guide Created!", "✅ DONE");
}
