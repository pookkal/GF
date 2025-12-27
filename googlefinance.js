/**
* ==============================================================================
* BASELINE LABEL: STABLE_MASTER_DEC25_BASE_v2_0_lastGF
* DATE: 27 DEC 2025
* FIXES:
* 1) Dashboard filter restored + hardened (token-exact matching, supports USA/INDIA and P0/M7).
*    - ALL anywhere in INPUT!B1 or INPUT!C1 disables that dimension.
* 2) Auto-refresh Dashboard when INPUT!B1 or INPUT!C1 changes.
* 3) Calculations: Added ADX(14) + Stoch %K(14); SIGNAL(E) uses full technical stack.
* 4) Decision(F) = D + E confluence with risk gates.
* 5) TECH_REASON and FUND_REASON moved AFTER new columns; Chart + VLOOKUP indices updated.
* 6) Alerts LAST_STATE moved to AB; alert engine updated accordingly.
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
    .addSeparator()
    .addItem('1. Fetch Data Only', 'generateDataSheet')
    .addItem('2. Build Calculations', 'generateCalculationsSheet')
    .addSeparator()
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

  // ------------------------------------------------------------------
  // 1) DASHBOARD manual update button (DASHBOARD!B1)
  // ------------------------------------------------------------------
  if (sheet.getName() === "DASHBOARD" && a1 === "B1" && e.value === "TRUE") {
    ss.toast("Recalculating Signals...", "⚙️ SYSTEM", 3);
    try {
      generateCalculationsSheet();
      generateDashboardSheet(); // rebuilds UI + applies filter formula + restores formatting
      ss.toast("Terminal Synchronized.", "✅ DONE", 2);
    } catch (err) {
      sheet.getRange("B1").setValue(false); // checkbox cleanup
      ss.toast("Error: " + err.toString(), "⚠️ FAIL", 5);
    }
    return;
  }

  // ------------------------------------------------------------------
  // 2) INPUT filters (INPUT!B1 / INPUT!C1) -> refresh dashboard
  // ------------------------------------------------------------------
  if (sheet.getName() === "INPUT" && (a1 === "B1" || a1 === "C1")) {
    try {
      generateDashboardSheet();
      SpreadsheetApp.flush();
    } catch (err) {
      ss.toast("Dashboard filter refresh error: " + err.toString(), "⚠️ FAIL", 5);
    }
    return;
  }

  // ------------------------------------------------------------------
  // 3) CHART controls -> update dynamic chart
  // ------------------------------------------------------------------
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

// ADX(14) (Wilder). Returns ~0..60+ typical.
function LIVEADX(highHist, lowHist, closeHist, currentPrice) {
  try {
    if (!highHist || !lowHist || !closeHist || !currentPrice) return 0;

    const H = highHist.flat().filter(n => typeof n === 'number' && n > 0);
    const L = lowHist.flat().filter(n => typeof n === 'number' && n > 0);
    const C = closeHist.flat().filter(n => typeof n === 'number' && n > 0);

    const n = Math.min(H.length, L.length, C.length);
    if (n < 40) return 0;

    // Keep last ~80 bars
    const take = Math.min(n, 90);
    const h = H.slice(n - take);
    const l = L.slice(n - take);
    const c = C.slice(n - take);

    // Append "live" close as last close; keep high/low of last bar as last known (best available)
    const liveClose = Number(currentPrice);
    c[c.length - 1] = liveClose;

    const period = 14;

    // Compute TR, +DM, -DM
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

    // Wilder smoothing initial sums
    let atr = tr.slice(0, period).reduce((a, b) => a + b, 0);
    let pDM14 = pdm.slice(0, period).reduce((a, b) => a + b, 0);
    let nDM14 = ndm.slice(0, period).reduce((a, b) => a + b, 0);

    // First DX
    const pDI0 = (atr === 0) ? 0 : (100 * (pDM14 / atr));
    const nDI0 = (atr === 0) ? 0 : (100 * (nDM14 / atr));
    let dxArr = [];
    dxArr.push((pDI0 + nDI0 === 0) ? 0 : (100 * Math.abs(pDI0 - nDI0) / (pDI0 + nDI0)));

    // Continue smoothing and DX
    for (let i = period; i < tr.length; i++) {
      atr = atr - (atr / period) + tr[i];
      pDM14 = pDM14 - (pDM14 / period) + pdm[i];
      nDM14 = nDM14 - (nDM14 / period) + ndm[i];

      const pDI = (atr === 0) ? 0 : (100 * (pDM14 / atr));
      const nDI = (atr === 0) ? 0 : (100 * (nDM14 / atr));
      const dx = (pDI + nDI === 0) ? 0 : (100 * Math.abs(pDI - nDI) / (pDI + nDI));
      dxArr.push(dx);
    }

    // ADX = Wilder-smoothed DX
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

    // Full table: date, open, high, low, close, volume
    dataSheet.getRange(4, colStart)
      .setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "all", TODAY()-800, TODAY()), "No Data")`);

    // Formats
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
* 4. CALCULATION ENGINE (ADX + STOCH inserted; TECH/FUND moved after)
* ------------------------------------------------------------------
* FINAL COLUMN MAP (CALCULATIONS):
* A  Ticker
* B  Price
* C  Change %
* D  FUNDAMENTAL
* E  SIGNAL (TECH ENGINE)
* F  DECISION (D+E confluence)
* G  ATH (TRUE)
* H  ATH Diff %
* I  R:R Quality
* J  Trend Score
* K  Trend State
* L  SMA 20
* M  SMA 50
* N  SMA 200
* O  Vol Trend (RVOL proxy)
* P  RSI
* Q  MACD Hist
* R  Divergence
* S  Support
* T  Target (3:1)
* U  Resistance
* V  ATR (14)
* W  Bollinger %B (proxy)
* X  ADX (14)
* Y  Stoch %K (14) in 0..1
* Z  TECH_REASON
* AA FUND_REASON
* AB LAST_STATE (alert memory)
* ------------------------------------------------------------------
*/
function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet || !inputSheet) return;

  const tickers = getCleanTickers(inputSheet);
  let calcSheet = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");

  // Persist LAST_STATE (AB)
  const stateMap = {};
  if (calcSheet.getLastRow() >= 3) {
    const existing = calcSheet.getRange(3, 1, calcSheet.getLastRow() - 2, 28).getValues();
    existing.forEach(r => {
      const t = (r[0] || "").toString().trim().toUpperCase();
      if (t) stateMap[t] = r[27];
    });
  }

  calcSheet.clear().clearFormats();

  // Timestamp
  const syncTime = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  calcSheet.getRange("A1").setValue(syncTime).setFontSize(8).setFontColor("#757575").setFontStyle("italic");

  // Group headers
  calcSheet.getRange("B1:F1").merge().setValue("[ CORE IDENT ]").setBackground("#263238").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("G1:I1").merge().setValue("[ PERFORMANCE ]").setBackground("#0D47A1").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("J1:R1").merge().setValue("[ MOMENTUM ]").setBackground("#1B5E20").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("S1:W1").merge().setValue("[ RISK LEVELS ]").setBackground("#B71C1C").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("X1:Y1").merge().setValue("[ NEW INDICATORS ]").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("Z1:AA1").merge().setValue("[ AUDIT ]").setBackground("#212121").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("AB1").setValue("[ STATE ]").setBackground("#000000").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  const headers = [[
    "Ticker", "Price", "Change %", "FUNDAMENTAL", "SIGNAL", "DECISION",
    "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State",
    "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "MACD Hist", "Divergence",
    "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B",
    "ADX (14)", "Stoch %K (14)",
    "TECH_REASON", "FUND_REASON",
    "LAST_STATE"
  ]];
  calcSheet.getRange(2, 1, 1, 28).setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold");

  const formulas = [];
  const restoredStates = [];

  tickers.forEach((ticker, i) => {
    const rowNum = i + 3;
    const t = ticker.toString().trim().toUpperCase();
    restoredStates.push([stateMap[t] || ""]);

    const tDS = (i * 7) + 1;

    // DATA block columns
    const dateCol = columnToLetter(tDS + 0);
    const openCol = columnToLetter(tDS + 1);
    const highCol = columnToLetter(tDS + 2);
    const lowCol  = columnToLetter(tDS + 3);
    const closeCol = columnToLetter(tDS + 4);
    const volCol = columnToLetter(tDS + 5);

    const lastRow = `COUNTA(DATA!$${closeCol}:$${closeCol})`;

    formulas.push([
      // B Price
      `=ROUND(IFERROR(GOOGLEFINANCE("${t}", "price")), 2)`,

      // C Change %
      `=IFERROR(GOOGLEFINANCE("${t}", "changepct")/100, 0)`,

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

      // E SIGNAL (full technical stack)
      `=IF(OR(ISBLANK(B${rowNum}), B${rowNum}=0), "🔄 LOADING...",
        IFERROR(IFS(
          B${rowNum}<S${rowNum}, "STOP LOSS",
          B${rowNum}<N${rowNum}, "BEAR REGIME",
          B${rowNum}>=U${rowNum}*0.99, "RESISTANCE TEST",

          AND(O${rowNum}>1.5, B${rowNum}>L${rowNum}, Q${rowNum}>0, X${rowNum}>=18), "RVOL BREAKOUT",

          AND(Y${rowNum}<0.2, B${rowNum}>S${rowNum}, X${rowNum}>=18), "STOCH OVERSOLD BOUNCE",
          AND(Y${rowNum}>0.8, B${rowNum}>=U${rowNum}*0.97), "STOCH OVERBOUGHT FADE",

          AND(P${rowNum}<35, B${rowNum}>S${rowNum}), "RSI SUPPORT BOUNCE",

          AND(W${rowNum}<0.2, Q${rowNum}>0, X${rowNum}<18), "VOL SQUEEZE (CHOP)",

          X${rowNum}<18, "CHOP (LOW ADX)",
          TRUE, "CHOP"
        ), "CHOP")
      )`,

      // F DECISION = D + E confluence (risk gates)
      `=IF(E${rowNum}="🔄 LOADING...", "🔄 LOADING...",
        IFS(
          REGEXMATCH(E${rowNum}, "STOP"), "🛑 STOP OUT",

          D${rowNum}="💀 ZOMBIE", "💤 AVOID",
          REGEXMATCH(D${rowNum}, "BUBBLE"), "💤 AVOID",

          AND(E${rowNum}="RVOL BREAKOUT", D${rowNum}="💎 GEM (Value)", I${rowNum}>=1.5, X${rowNum}>=20), "💎 PRIME BUY",
          AND(E${rowNum}="RVOL BREAKOUT", I${rowNum}<1.1), "⚠️ POOR R:R (AVOID)",
          AND(E${rowNum}="RVOL BREAKOUT", O${rowNum}<1.2), "🎣 FAKE-OUT (NO VOL)",

          AND(B${rowNum}>L${rowNum}+(2*V${rowNum})), "⏳ ATR OVEREXTENDED",

          AND(E${rowNum}="STOCH OVERSOLD BOUNCE", B${rowNum}>N${rowNum}, X${rowNum}>=18), "🚀 TRADE (MEAN REV)",
          AND(E${rowNum}="RSI SUPPORT BOUNCE", B${rowNum}>N${rowNum}, X${rowNum}>=18), "🚀 TRADE",

          E${rowNum}="BEAR REGIME", "💤 AVOID",
          TRUE, "⏳ WAIT"
        )
      )`,

      // G ATH (TRUE)
      `=IFERROR(DATA!${columnToLetter(tDS + 1)}3, "-")`,

      // H ATH Diff %
      `=IFERROR((B${rowNum}-G${rowNum})/G${rowNum}, 0)`,

      // I R:R Quality
      `=IFERROR(ROUND((U${rowNum}-B${rowNum})/MAX(0.01, B${rowNum}-S${rowNum}), 2), 0)`,

      // J Trend Score
      `=REPT("★", (B${rowNum}>L${rowNum}) + (B${rowNum}>M${rowNum}) + (B${rowNum}>N${rowNum}))`,

      // K Trend State
      `=IF(B${rowNum}>N${rowNum}, "BULL REGIME", "BEAR REGIME")`,

      // L SMA 20
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)), 0), 2)`,

      // M SMA 50
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-50, 0, 50)), 0), 2)`,

      // N SMA 200
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-200, 0, 200)), 0), 2)`,

      // O Vol Trend (RVOL proxy)
      `=ROUND(IFERROR(OFFSET(DATA!$${volCol}$4, ${lastRow}-1, 0) / AVERAGE(OFFSET(DATA!$${volCol}$4, ${lastRow}-21, 0, 20)), 1), 2)`,

      // P RSI
      `=LIVERSI(DATA!$${closeCol}$4:$${closeCol}, B${rowNum})`,

      // Q MACD Hist
      `=LIVEMACD(DATA!$${closeCol}$4:$${closeCol}, B${rowNum})`,

      // R Divergence
      `=IFERROR(IFS(
        AND(B${rowNum}<INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), P${rowNum}>50), "BULLISH DIV",
        AND(B${rowNum}>INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), P${rowNum}<50), "BEARISH DIV",
        TRUE, "-"
      ), "-")`,

      // S Support (20-day min low)
      `=ROUND(IFERROR(MIN(OFFSET(DATA!$${lowCol}$4, ${lastRow}-21, 0, 20)), B${rowNum}*0.9), 2)`,

      // T Target (3:1)
      `=ROUND(B${rowNum} + ((B${rowNum}-S${rowNum}) * 3), 2)`,

      // U Resistance (50-day max high)
      `=ROUND(IFERROR(MAX(OFFSET(DATA!$${highCol}$4, ${lastRow}-51, 0, 50)), B${rowNum}*1.1), 2)`,

      // V ATR (14) (high-low proxy)
      `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(
        OFFSET(DATA!$${highCol}$4, ${lastRow}-14, 0, 14) - OFFSET(DATA!$${lowCol}$4, ${lastRow}-14, 0, 14)
      )), 0), 2)`,

      // W Bollinger %B proxy (keep original proxy logic)
      `=ROUND(IFERROR(((B${rowNum}-L${rowNum}) / (4*STDEV(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)))) + 0.5, 0.5), 2)`,

      // X ADX(14)
      `=LIVEADX(DATA!$${highCol}$4:$${highCol}, DATA!$${lowCol}$4:$${lowCol}, DATA!$${closeCol}$4:$${closeCol}, B${rowNum})`,

      // Y Stoch %K(14) (0..1)
      `=LIVESTOCHK(DATA!$${highCol}$4:$${highCol}, DATA!$${lowCol}$4:$${lowCol}, DATA!$${closeCol}$4:$${closeCol}, B${rowNum})`,

      // Z TECH_REASON (explicit indicator naming)
      `="1. VOL CONVICTION (RVOL): " & IF(O${rowNum}>1.5, "Strong Institutional RVOL (" & O${rowNum} & "x).", "Low RVOL (" & O${rowNum} & "x). Negative Drag.") & CHAR(10) &
        "2. STRUCTURE (SMA200): " & IF(B${rowNum}>N${rowNum}, "Above SMA 200 (Bull Regime).", "Below SMA 200 (Bear Regime).") & CHAR(10) &
        "3. STRETCH (ATR/SMA20): " & IF(B${rowNum} > L${rowNum} + (2*V${rowNum}), "Overextended > 2x ATR. Mean Reversion Risk.", "Safe Zone. Price within 2x ATR of SMA 20.") & CHAR(10) &
        "4. MOMENTUM: RSI (" & P${rowNum} & ") is " & IF(P${rowNum}>70,"Overbought",IF(P${rowNum}<30,"Oversold","Stable")) & ". MACD is " & IF(Q${rowNum}>0,"Bullish.","Weak.") & CHAR(10) &
        "5. ADX (14): " & X${rowNum} & ". " & IF(X${rowNum}>=25,"Trend is strong.","Trend is weak/ranging.") & CHAR(10) &
        "6. STOCH %K (14): " & Y${rowNum} & ". " & IF(Y${rowNum}<0.2,"Oversold zone.",IF(Y${rowNum}>0.8,"Overbought zone.","Neutral zone.")) & CHAR(10) &
        "7. RR SCORE: " & I${rowNum} & "x. " & IF(I${rowNum}>=3, "Institutional 3:1+ edge.", "Sub-optimal payout vs risk.")`,

      // AA FUND_REASON
      `="1. VALUATION: " & D${rowNum} & ". P/E of " & IFERROR(GOOGLEFINANCE(A${rowNum}, "pe"),"N/A") & " and EPS " & IFERROR(GOOGLEFINANCE(A${rowNum}, "eps"),"N/A") & "." & CHAR(10) &
        "2. REGIME (SMA200): " & IF(B${rowNum}>N${rowNum}, "Long-term Bullish above SMA 200.", "Long-term Bearish below SMA 200.") & CHAR(10) &
        "3. TREND QUALITY (ADX): " & X${rowNum} & ". " & IF(X${rowNum}>=25,"Trend quality supports continuation trades.","Low ADX suggests range risk.") & CHAR(10) &
        "4. VERDICT: " & F${rowNum} & ". Confluence of Fundamentals + Technicals."`
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

  // Formats
  calcSheet.getRange("C3:C").setNumberFormat("0.00%");
  calcSheet.getRange("H3:H").setNumberFormat("0.00%");
  calcSheet.getRange("W3:W").setNumberFormat("0.00%");
  calcSheet.getRange("Y3:Y").setNumberFormat("0.00%");

  // Conditional formatting
  const lastRowVal = Math.max(calcSheet.getLastRow(), 3);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor("#C62828").setBold(true)
      .setRanges([calcSheet.getRange("C3:C" + lastRowVal), calcSheet.getRange("H3:H" + lastRowVal), calcSheet.getRange("Q3:Q" + lastRowVal)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR(P3>70, P3<30)').setFontColor("#C62828").setBold(true)
      .setRanges([calcSheet.getRange("P3:P" + lastRowVal)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=X3>=25').setFontColor("#2E7D32").setBold(true)
      .setRanges([calcSheet.getRange("X3:X" + lastRowVal)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=Y3<0.2').setFontColor("#2E7D32").setBold(true)
      .setRanges([calcSheet.getRange("Y3:Y" + lastRowVal)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=Y3>0.8').setFontColor("#C62828").setBold(true)
      .setRanges([calcSheet.getRange("Y3:Y" + lastRowVal)])
      .build()
  ];
  calcSheet.setConditionalFormatRules(rules);

  SpreadsheetApp.flush();
  calcSheet.setFrozenRows(2);
}


/**
* ------------------------------------------------------------------
* 5. DASHBOARD ENGINE (LIVE FILTER FORMULA - TOKEN EXACT + ALL OVERRIDE)
* ------------------------------------------------------------------
* Sector filter: INPUT!B1 tokens match INPUT!B3:B (exact token)
* Industry filter: INPUT!C1 tokens match INPUT!C3:C (exact token)
* Combined logic: (B match OR ALL) AND (C match OR ALL)
* ------------------------------------------------------------------
*/
function generateDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;
  let dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD");
  dashboard.clear().clearFormats();

  // HEADER UI & HEADERS (Rows 1-2)
  dashboard.getRange("A1").setValue("UPDATE").setBackground("#212121").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  dashboard.getRange("B1").insertCheckboxes().setBackground("#212121").setHorizontalAlignment("center");
  dashboard.getRange("C1:E1").merge().setBackground("#000000").setFontColor("#00FF00").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle").setValue(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM dd, yyyy | HH:mm:ss"));

  const headers = [["Ticker", "Price", "Change %", "FUNDAMENTAL", "SIGNAL", "DECISION", "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "MACD Hist", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "ADX (14)", "Stoch %K (14)", "TECH ANALYSIS", "FUND ANALYSIS"]];
  dashboard.getRange(2, 1, 1, 27).setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

  // NEW HARDENED FILTER FORMULA 
  const formula = '=IFERROR(SORT(FILTER(CALCULATIONS!$A$3:$AA, ISNUMBER(MATCH(CALCULATIONS!$A$3:$A, FILTER(INPUT!$A$3:$A, INPUT!$A$3:$A<>"", (IF(OR(INPUT!$B$1="", REGEXMATCH(UPPER(INPUT!$B$1), "(^|,\\s*)ALL(\\s*|,|$)")), TRUE, REGEXMATCH(","&UPPER(TRIM(INPUT!$B$3:$B))&",", ",\\s*(" & REGEXREPLACE(UPPER(TRIM(INPUT!$B$1)), "\\s*,\\s*", "|") & ")\\s*,"))) * (IF(OR(INPUT!$C$1="", REGEXMATCH(UPPER(INPUT!$C$1), "(^|,\\s*)ALL(\\s*|,|$)")), TRUE, REGEXMATCH(","&REGEXREPLACE(UPPER(INPUT!$C$3:$C), "\\s+", "")&",", ",(" & REGEXREPLACE(UPPER(TRIM(INPUT!$C$1)), "\\s*,\\s*", "|") & "),")))), 0))), 3, FALSE), "No Matches Found")';
  
  dashboard.getRange("A3").setFormula(formula);
  SpreadsheetApp.flush();

  // ---------------------------
  // FORMATTING GOVERNANCE (re-apply)
  // ---------------------------
  const lastRow = dashboard.getLastRow();

  dashboard.setFrozenRows(2);
  dashboard.setFrozenColumns(1);

  // Column widths
  for (let col = 1; col <= 25; col++) dashboard.setColumnWidth(col, 75);
  dashboard.setColumnWidth(26, 350);
  dashboard.setColumnWidth(27, 350);

  // If no rows, stop after headers
  if (lastRow < 3) return;

  const rows = lastRow - 2;

  // Row heights + wrapping
  dashboard.setRowHeights(3, rows, 28);

  // Main numeric/text area wrap, keep audits clipped (prevents giant rows)
  dashboard.getRange(3, 1, rows, 25).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  dashboard.getRange(3, 26, rows, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Alignment: keep left alignment (your ask)
  const dataRange = dashboard.getRange(3, 1, rows, 27);
  dataRange.setHorizontalAlignment("left").setVerticalAlignment("middle");

  // Borders (grid)
  dataRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Number formats
  dashboard.getRange("C3:C" + lastRow).setNumberFormat("0.00%"); // Change %
  dashboard.getRange("H3:H" + lastRow).setNumberFormat("0.00%"); // ATH Diff %
  dashboard.getRange("W3:W" + lastRow).setNumberFormat("0.00%"); // Boll %B
  dashboard.getRange("Y3:Y" + lastRow).setNumberFormat("0.00%"); // Stoch %K (0..1)

  // ---------------------------
  // CONDITIONAL FORMATTING (restore color codings)
  // ---------------------------
  const rules = [];

  // Negative Drag (Font only)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([
        dashboard.getRange("C3:C" + lastRow), // Change %
        dashboard.getRange("H3:H" + lastRow), // ATH diff %
        dashboard.getRange("Q3:Q" + lastRow)  // MACD hist
      ])
      .build()
  );

  // RSI extremes
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($P3>70, $P3<30)')
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([dashboard.getRange("P3:P" + lastRow)])
      .build()
  );

  // ADX strong trend
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$X3>=25')
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("X3:X" + lastRow)])
      .build()
  );

  // Stoch zones (0..1)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$Y3<0.2')
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("Y3:Y" + lastRow)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$Y3>0.8')
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([dashboard.getRange("Y3:Y" + lastRow)])
      .build()
  );

  // Decision States (Background for E:F)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($E3&" "&$F3, "(?i)PRIME|TRADE|BREAKOUT")')
      .setBackground("#E8F5E9")
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("E3:F" + lastRow)])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($E3&" "&$F3, "(?i)FAKE-OUT|OVEREXTENDED")')
      .setBackground("#FFF3E0")
      .setFontColor("#E65100")
      .setBold(true)
      .setRanges([dashboard.getRange("E3:F" + lastRow)])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($E3&" "&$F3, "(?i)STOP|AVOID|BEAR")')
      .setBackground("#FFEBEE")
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([dashboard.getRange("E3:F" + lastRow)])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($E3&" "&$F3, "(?i)CHOP|WAIT|LOADING")')
      .setBackground("#F5F5F5")
      .setFontColor("#9E9E9E")
      .setRanges([dashboard.getRange("E3:F" + lastRow)])
      .build()
  );

  // Fundamental GEM highlight (D)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("GEM")
      .setBackground("#E8F5E9")
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("D3:D" + lastRow)])
      .build()
  );

  // Optional: Trend State highlight (K)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("BULL")
      .setFontColor("#2E7D32")
      .setBold(true)
      .setRanges([dashboard.getRange("K3:K" + lastRow)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("BEAR")
      .setFontColor("#C62828")
      .setBold(true)
      .setRanges([dashboard.getRange("K3:K" + lastRow)])
      .build()
  );

  dashboard.setConditionalFormatRules(rules);
}

/**
 * ------------------------------------------------------------------
 * 6. SETUP CHART SHEET (Stable sidebar sections + formulas)
 * - Fixes the "SMA 20 appears as title" bug by NOT using hard-coded row numbers.
 * - Section headers are auto-detected by label starting with "[".
 * ------------------------------------------------------------------
 */
function setupChartSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  const tickers = getCleanTickers(inputSheet);

  let chartSheet = ss.getSheetByName("CHART") || ss.insertSheet("CHART");
  chartSheet.clear().clearFormats();
  forceExpandSheet(chartSheet, 60);

  // Layout
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

  // Dual reasoning boxes (keep positions)
  chartSheet.getRange("E1:F4").merge()
    .setWrap(true).setVerticalAlignment("top").setFontSize(10)
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);
  chartSheet.getRange("E1").setFormula('=IFERROR(VLOOKUP(B1, CALCULATIONS!$A$3:$AA, 26, 0), "—")'); // TECH_REASON

  chartSheet.getRange("G1:H4").merge()
    .setWrap(true).setVerticalAlignment("top").setFontSize(10)
    .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);
  chartSheet.getRange("G1").setFormula('=IFERROR(VLOOKUP(B1, CALCULATIONS!$A$3:$AA, 27, 0), "—")'); // FUND_REASON

  // Date controls
  chartSheet.getRange("A2:C2").setValues([["YEAR", "MONTH", "DAY"]]).setFontWeight("bold").setHorizontalAlignment("center");
  const numRule = (max) => SpreadsheetApp.newDataValidation().requireValueInList(Array.from({ length: max + 1 }, (_, i) => i)).build();
  chartSheet.getRange("A3").setDataValidation(numRule(5)).setValue(1).setHorizontalAlignment("center").setFontColor("#FF80AB");
  chartSheet.getRange("B3").setDataValidation(numRule(12)).setValue(0).setHorizontalAlignment("center").setFontColor("#FF80AB");
  chartSheet.getRange("C3").setDataValidation(numRule(31)).setValue(0).setHorizontalAlignment("center").setFontColor("#FF80AB");

  chartSheet.getRange("D2")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"]).build())
    .setValue("DAILY")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontColor("#FF80AB");

  chartSheet.getRange("A4").setValue("DATE").setFontWeight("bold");
  chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)").setNumberFormat("yyyy-mm-dd");

  // ------------------------------------------------------------------
  // Sidebar (A5:B...) – dynamically styled section headers
  // Uses CALCULATIONS column mapping you confirmed:
  //  5 SIGNAL, 4 FUND, 6 DECISION
  //  7 ATH, 8 ATH Diff, 12/13/14 SMAs, 16 RSI, 17 MACD, 11 Trend State, 18 Divergence, 15 Vol Trend
  //  19 Support, 21 Resistance, 20 Target, 22 ATR, 23 Boll %B, 24 ADX, 25 Stoch %K
  // ------------------------------------------------------------------
  const t = "B1";
  const data = [
    ["SIGNAL (RAW)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 5, 0), "Wait")`],
    ["FUNDAMENTAL", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 4, 0), "-")`],
    ["DECISION (FINAL)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 6, 0), "-")`],
    ["LIVE PRICE", `=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`],
    ["CHANGE ($)", `=IFERROR(B8 - GOOGLEFINANCE(${t}, "closeyest"), 0)`],
    ["CHANGE (%)", `=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`],
    ["RANGE DIFF %", `=IFERROR((B8 - INDEX(GOOGLEFINANCE(${t}, "close", B4), 2, 2)) / INDEX(GOOGLEFINANCE(${t}, "close", B4), 2, 2), 0)`],
    ["", ""],

    ["[ VALUATION METRICS ]", ""],
    ["ATH (TRUE)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 7, 0), 0)`],
    ["DIFF FROM ATH %", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 8, 0), 0)`],
    ["P/E RATIO", `=IFERROR(GOOGLEFINANCE(${t}, "pe"), 0)`],
    ["EPS", `=IFERROR(GOOGLEFINANCE(${t}, "eps"), 0)`],
    ["52W HIGH", `=IFERROR(GOOGLEFINANCE(${t}, "high52"), 0)`],
    ["52W LOW", `=IFERROR(GOOGLEFINANCE(${t}, "low52"), 0)`],
    ["", ""],

    ["[ MOMENTUM & TREND ]", ""],
    ["SMA 20", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 12, 0), 0)`],
    ["SMA 50", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 13, 0), 0)`],
    ["SMA 200", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 14, 0), 0)`],
    ["RSI (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 16, 0), 50)`],
    ["MACD HIST", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 17, 0), 0)`],
    ["ADX (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 24, 0), 0)`],
    ["STOCH %K (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 25, 0), 0)`],
    ["TREND STATE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 11, 0), "—")`],
    ["DIVERGENCE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 18, 0), "Neutral")`],
    ["RELATIVE VOLUME", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 15, 0), 1)`],
    ["", ""],

    ["[ TECHNICAL LEVELS ]", ""],
    ["SUPPORT FLOOR", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 19, 0), 0)`],
    ["RESISTANCE CEILING", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 21, 0), 0)`],
    ["TARGET (3:1 R:R)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 20, 0), 0)`],
    ["ATR (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 22, 0), 0)`],
    ["BOLLINGER %B", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$AA, 23, 0), 0)`]
  ];

  // Write sidebar labels + formulas
  const startRow = 5;
  chartSheet.getRange(startRow, 1, data.length, 1).setValues(data.map(r => [r[0]])).setFontWeight("bold");
  chartSheet.getRange(startRow, 2, data.length, 1).setFormulas(data.map(r => [r[1]]));

  // Auto-style section headers (rows whose label starts with "[")
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

  // Alignments
  chartSheet.getRange(`B${startRow}:B${startRow + data.length - 1}`).setHorizontalAlignment("left");

  // Currency formatting (keep close to your original intent)
  chartSheet.getRangeList(["B8", "B9", "B14", "B16:B19", "B23:B26", "B34:B37"]).setNumberFormat("#,##0.00");
  // Percent formatting
  chartSheet.getRangeList(["B10", "B11", "B12", "B36"]).setNumberFormat("0.00%"); // some may be blank; harmless

  // Conditional formatting rules (same as your prior approach, but without row hardcoding dependencies)
  const rules = [];

  // Negative change red
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#D32F2F")
      .setRanges([chartSheet.getRange("B9:B10")])
      .build()
  );

  // RSI extreme (by label-based read is better, but keep simple: find RSI row dynamically)
  // We will just apply to entire value column to preserve your visual cues
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(70)
      .setFontColor("#D32F2F")
      .setRanges([chartSheet.getRange(`B${startRow}:B${startRow + data.length - 1}`)])
      .build()
  );

  // MACD negative red (apply to entire column; safe)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#D32F2F")
      .setRanges([chartSheet.getRange(`B${startRow}:B${startRow + data.length - 1}`)])
      .build()
  );

  // P/E expensive red (apply to entire column; safe)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(50)
      .setFontColor("#D32F2F")
      .setRanges([chartSheet.getRange(`B${startRow}:B${startRow + data.length - 1}`)])
      .build()
  );

  chartSheet.setConditionalFormatRules(rules);

  // Build initial chart
  updateDynamicChart();
}


/**
 * ------------------------------------------------------------------
 * Update chart using ONLY the CHART sidebar values:
 * - Live Price  -> read from sidebar label "LIVE PRICE" (robust)
 * - Support     -> read from sidebar label "SUPPORT FLOOR" (robust)
 * - Resistance  -> read from sidebar label "RESISTANCE CEILING" (robust)
 *
 * This avoids any dependency on CALCULATIONS for levels and avoids hard-coded row numbers.
 * ------------------------------------------------------------------
 */
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;

  SpreadsheetApp.flush();

  sheet.getRange("E5")
    .setValue("Updated: " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "HH:mm:ss"))
    .setFontColor("gray")
    .setFontSize(8)
    .setHorizontalAlignment("right");

  const ticker = String(sheet.getRange("B1").getValue() || "").trim();
  if (!ticker) return;

  const isWeekly = sheet.getRange("D2").getValue() === "WEEKLY";
  const years = Number(sheet.getRange("A3").getValue()) || 0;
  const months = Number(sheet.getRange("B3").getValue()) || 0;
  const days = Number(sheet.getRange("C3").getValue()) || 0;

  const now = new Date();
  let startDate = new Date(now.getFullYear() - years, now.getMonth() - months, now.getDate() - days);
  if ((now - startDate) < (7 * 24 * 60 * 60 * 1000)) {
    startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 14);
  }

  // Sidebar values (label-based, stable)
  const sidebar = getSidebarValuesByLabels_(sheet, ["LIVE PRICE", "SUPPORT FLOOR", "RESISTANCE CEILING"]);
  let livePrice = Number(sidebar["LIVE PRICE"]) || 0;
  const supportVal = Number(sidebar["SUPPORT FLOOR"]) || 0;
  const resistanceVal = Number(sidebar["RESISTANCE CEILING"]) || 0;

  // Find ticker block in DATA (row 2)
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
    if (isWeekly && d.getDay() !== 5) continue;

    const slice = rawData
      .slice(Math.max(4, i - 200), i + 1)
      .map(r => Number(r[4]))
      .filter(n => isFinite(n) && n > 0);

    const s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
    const s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
    const s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;

    const prevClose = (i > 4) ? Number(rawData[i - 1][4]) : close;

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

  let candleLabel = "🔴 LIVE";
  if ((!livePrice || !isFinite(livePrice)) && prices.length > 0) {
    livePrice = prices[prices.length - 1];
    candleLabel = "⏳ SYNCING";
  }

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
      1: { viewWindow: { min: 0, max: maxVol * 4 }, textStyle: { color: "#666" }, format: "short" }
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
 * Returns a map {label: value}.
 * Matching is case-insensitive and trims whitespace.
 */
function getSidebarValuesByLabels_(chartSheet, labels) {
  const want = new Set(labels.map(l => String(l).trim().toUpperCase()));
  const keys = chartSheet.getRange("A5:A120").getValues().flat().map(v => String(v || "").trim().toUpperCase());
  const vals = chartSheet.getRange("B5:B120").getValues().flat();

  const out = {};
  for (let i = 0; i < keys.length; i++) {
    if (want.has(keys[i])) {
      out[labels.find(l => String(l).trim().toUpperCase() === keys[i])] = vals[i];
    }
  }
  // Ensure missing labels return 0
  labels.forEach(l => { if (out[l] === undefined) out[l] = 0; });
  return out;
}


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

  // -------------------------------------------------------------------
  // Sidebar reads (SOURCE OF TRUTH per your instruction)
  // -------------------------------------------------------------------

  // 1) Live price from sidebar (your existing layout uses B8)
  let livePrice = Number(sheet.getRange("B8").getValue()) || 0;

  // 2) Support/Resistance from sidebar by label (robust to row drift)
  const levels = getSidebarLevels_(sheet);
  const supportVal = levels.support;
  const resistanceVal = levels.resistance;

  // -------------------------------------------------------------------
  // Locate ticker data block in DATA sheet
  // DATA row 2 contains the ticker label per block
  // Pull 6 columns: Date, Open, High, Low, Close, Volume
  // -------------------------------------------------------------------
  const rawHeaders = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIdx = rawHeaders.indexOf(ticker);
  if (colIdx === -1) {
    sheet.getRange("E1").setValue("⚠️ Ticker Not Found");
    return;
  }

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
    if (isWeekly && d.getDay() !== 5) continue; // your convention: Fridays

    const slice = rawData
      .slice(Math.max(4, i - 200), i + 1)
      .map(r => Number(r[4]))
      .filter(n => isFinite(n) && n > 0);

    const s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a, b) => a + b, 0) / 20).toFixed(2)) : null;
    const s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a, b) => a + b, 0) / 50).toFixed(2)) : null;
    const s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a, b) => a + b, 0) / 200).toFixed(2)) : null;

    const prevClose = (i > 4) ? Number(rawData[i - 1][4]) : close;

    masterData.push([
      Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "MMM dd"),
      close,
      (close >= prevClose) ? vol : null, // Bull Vol
      (close < prevClose) ? vol : null,  // Bear Vol
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

  // Sidebar live price fallback if empty
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

    masterData.push([
      candleLabel,
      livePrice,
      null,
      null,
      liveS20,
      liveS50,
      liveS200,
      resistanceVal || null,
      supportVal || null
    ]);

    prices.push(livePrice);
  }

  // Write chart data region (Z..AH area; col 26..34)
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
      1: { viewWindow: { min: 0, max: maxVol * 4 }, textStyle: { color: "#666" }, format: "short" }
    })
    .setOption("legend", { position: "top", textStyle: { fontSize: 10 } })
    .setPosition(5, 3, 0, 0)
    .setOption("width", 1150)
    .setOption("height", 650)
    .build();

  sheet.insertChart(chart);
}

/**
 * Reads sidebar support/resistance by label from A5:B80.
 * Labels must match exactly: "SUPPORT FLOOR" and "RESISTANCE CEILING"
 */
function getSidebarLevels_(chartSheet) {
  const labelRange = chartSheet.getRange("A5:A80").getValues().flat();
  const valueRange = chartSheet.getRange("B5:B80").getValues().flat();

  const findValue = (label) => {
    const idx = labelRange.findIndex(v => String(v || "").trim().toUpperCase() === label);
    if (idx === -1) return 0;
    return Number(valueRange[idx]) || 0;
  };

  return {
    support: findValue("SUPPORT FLOOR"),
    resistance: findValue("RESISTANCE CEILING")
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
    const currentDecision = row[5];   // F
    const lastNotified = row[27];     // AB

    if (!ticker || !currentDecision) return;

    if (currentDecision !== lastNotified) {
      const isActionable = /(PRIME|TRADE|STOP|BOUNCE|BREAKOUT)/i.test(currentDecision);
      if (isActionable) {
        newAlerts.push(`📍 TICKER: ${ticker}\n   NEW SIGNAL: ${currentDecision}\n   PREVIOUS: ${lastNotified || 'Initial Scan'}`);
      }
      // Update AB (col 28)
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
 * Creates a conceptual, non-complex REFERENCE_GUIDE for the Terminal.
 * Focuses on 'Trading Philosophy' rather than code logic.
 */
function generateReferenceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("CONCEPT_GUIDE") || ss.insertSheet("CONCEPT_GUIDE");
  sheet.clear().clearFormats();

  const data = [
    ["INSTITUTIONAL TERMINAL: CONCEPT & PHILOSOPHY", "", ""],
    ["Pillar 1: FUNDAMENTALS (Column D)", "Pillar 2: TECHNICALS (Column E)", "Pillar 3: DECISION (Column F)"],
    ["Focus: Business Quality", "Focus: Price Timing", "Focus: Final Action"],
    ["Scores 0-10 based on Balance Sheet health.", "Scores 0-10 based on Price Momentum.", "The 'Confluence' of Quality + Timing."],
    ["", "", ""],
    ["CORE INDICATOR GLOSSARY", "", ""],
    ["TERM", "MEANING", "ACTION TRIGGER"],
    ["LIVERSI", "Measures if a stock is 'exhausted'.", ">70 = Danger (Overbought); <30 = Opportunity (Oversold)"],
    ["LIVEMACD", "Measures the 'Pulse' of the trend.", "Positive & Growing = Trend is Healthy."],
    ["ADX (14)", "Trend Strength Filter.", "Below 25 = 'Choppy' market. Stay patient."],
    ["ATR (14)", "The 'Noise' Level.", "Used to set Stop Losses away from daily price wiggles."],
    ["R:R Quality", "Reward vs. Risk Math.", "Must be > 2.0 to take a trade. Don't risk $1 to make $0.50."]
  ];

  sheet.getRange(1, 1, data.length, 3).setValues(data);

  // Styling for clarity
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

