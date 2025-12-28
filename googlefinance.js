/**
* ==============================================================================
* BASELINE LABEL: STABLE_MASTER_DEC25_BASE_v3_5
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
  let calc = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");

  // Persist LAST_STATE (AB)
  const stateMap = {};
  if (calc.getLastRow() >= 3) {
    const existing = calc.getRange(3, 1, calc.getLastRow() - 2, 28).getValues();
    existing.forEach(r => {
      const t = (r[0] || "").toString().trim().toUpperCase();
      if (t) stateMap[t] = r[27];
    });
  }

  calc.clear().clearFormats();

  // Timestamp (small)
  const syncTime = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  calc.getRange("A1").setValue(syncTime).setFontSize(8).setFontColor("#757575").setFontStyle("italic");

  // Headers (Row 2) – Bloomberg / industry-standard naming
  const headers = [[
    "Ticker","SIGNAL","DECISION","FUNDAMENTAL","Price","Change %","Vol Trend","ATH (TRUE)","ATH Diff %","R:R Quality",
    "Trend Score","Trend State","SMA 20","SMA 50","SMA 200","RSI","MACD Hist","Divergence","ADX (14)","Stoch %K (14)",
    "Support","Resistance","Target (3:1)","ATR (14)","Bollinger %B","TECH NOTES","FUND NOTES","LAST_STATE"
  ]];
  calc.getRange(2, 1, 1, 28)
    .setValues(headers)
    .setBackground("#212121")
    .setFontColor("white")
    .setFontWeight("bold");

  // Write tickers in A3:A
  if (tickers.length > 0) {
    calc.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  }

  // Build formulas B..AA (27 columns), AB restore separately
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

    // Helper cell references by final calc layout
    // E Price, F Change%, G Vol, H ATH, I ATH%, J RR, M SMA20, O SMA200, P RSI, Q MACD, S ADX, T STO, U SUP, V RES, W TGT, X ATR, Y %B
    formulas.push([
      // B SIGNAL (normalized vocabulary)
      `=IF(OR(ISBLANK($E${row}),$E${row}=0),"LOADING",
        IFS(
          $E${row}<$U${row},"Stop-Out",
          $E${row}<$O${row},"Risk-Off (Below SMA200)",
          $S${row}<15,"Range-Bound (Low ADX)",
          AND($G${row}>=1.5,$E${row}>=$V${row}*0.995,$Q${row}>0,$S${row}>=18),"Breakout (High Volume)",
          AND($T${row}<=0.20,$E${row}>$U${row},$S${row}>=18),"Mean Reversion (Oversold)",
          AND($T${row}>=0.80,$E${row}>=$V${row}*0.97),"Mean Reversion (Overbought)",
          AND($E${row}>$O${row},$Q${row}>0,$S${row}>=18),"Trend Continuation",
          TRUE,"Hold / Monitor"
        )
      )`,

      // C DECISION (normalized action layer)
      `=IF($B${row}="LOADING","LOADING",
        IFS(
          $B${row}="Stop-Out","Stop-Out",
          OR($D${row}="ZOMBIE",$D${row}="BUBBLE"),"Avoid",
          $B${row}="Risk-Off (Below SMA200)","Avoid",
          AND($B${row}="Breakout (High Volume)",$J${row}>=1.5,$S${row}>=20),"Trade Long",
          AND($B${row}="Trend Continuation",$J${row}>=1.3,$S${row}>=18),"Accumulate",
          AND(REGEXMATCH($B${row},"Mean Reversion"),$J${row}>=1.2,$S${row}>=18),"Trade Long",
          AND($X${row}>0,$E${row}>$M${row}+(2*$X${row})),"Reduce (Overextended)",
          $B${row}="Range-Bound (Low ADX)","Hold / Monitor",
          TRUE,"Hold / Monitor"
        )
      )`,

      // D FUNDAMENTAL (keep your existing logic or replace with your latest)
      `=IFERROR(LET(eps, GOOGLEFINANCE($A${row}, "eps"), pe, GOOGLEFINANCE($A${row}, "pe"),
        IFS(
          eps<0, "ZOMBIE",
          AND(pe>0, pe>50), "PRICED FOR PERFECTION",
          AND(pe>0, pe<25, eps>0), "VALUE",
          AND(pe>30, eps<0.1), "BUBBLE",
          TRUE, "FAIR"
        )
      ), "FAIR")`,

      // E Price
      `=ROUND(IFERROR(GOOGLEFINANCE("${t}", "price"), 0), 2)`,

      // F Change %
      `=IFERROR(GOOGLEFINANCE("${t}", "changepct")/100, 0)`,

      // G Vol Trend (RVOL proxy)
      `=ROUND(IFERROR(OFFSET(DATA!$${volCol}$4, ${lastRow}-1, 0) / AVERAGE(OFFSET(DATA!$${volCol}$4, ${lastRow}-21, 0, 20)), 1), 2)`,

      // H ATH (TRUE)
      `=IFERROR(DATA!${columnToLetter(tDS + 1)}3, 0)`,

      // I ATH Diff %
      `=IFERROR(($E${row}-$H${row})/MAX(0.01,$H${row}), 0)`,

      // J R:R Quality
      `=IFERROR(ROUND(($V${row}-$E${row})/MAX(0.01, $E${row}-$U${row}), 2), 0)`,

      // K Trend Score (stars as text is okay; Bloomberg often uses 1..3 score – keep as you had)
      `=REPT("★", ($E${row}>$M${row}) + ($E${row}>$N${row}) + ($E${row}>$O${row}))`,

      // L Trend State
      `=IF($E${row}>$O${row}, "BULL", "BEAR")`,

      // M SMA 20
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)), 0), 2)`,

      // N SMA 50
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-50, 0, 50)), 0), 2)`,

      // O SMA 200
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-200, 0, 200)), 0), 2)`,

      // P RSI (kept as your custom function)
      `=LIVERSI(DATA!$${closeCol}$4:$${closeCol}, $E${row})`,

      // Q MACD Hist (kept as your custom function)
      `=LIVEMACD(DATA!$${closeCol}$4:$${closeCol}, $E${row})`,

      // R Divergence (keep your divergence heuristic)
      `=IFERROR(IFS(
        AND($E${row}<INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), $P${row}>50), "BULL DIV",
        AND($E${row}>INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), $P${row}<50), "BEAR DIV",
        TRUE, "—"
      ), "—")`,

      // S ADX (14)
      `=LIVEADX(DATA!$${highCol}$4:$${highCol}, DATA!$${lowCol}$4:$${lowCol}, DATA!$${closeCol}$4:$${closeCol}, $E${row})`,

      // T Stoch %K (14)
      `=LIVESTOCHK(DATA!$${highCol}$4:$${highCol}, DATA!$${lowCol}$4:$${lowCol}, DATA!$${closeCol}$4:$${closeCol}, $E${row})`,

      // U Support (20d min low)
      `=ROUND(IFERROR(MIN(OFFSET(DATA!$${lowCol}$4, ${lastRow}-21, 0, 20)), $E${row}*0.9), 2)`,

      // V Resistance (50d max high)
      `=ROUND(IFERROR(MAX(OFFSET(DATA!$${highCol}$4, ${lastRow}-51, 0, 50)), $E${row}*1.1), 2)`,

      // W Target (3:1)
      `=ROUND($E${row} + (($E${row}-$U${row}) * 3), 2)`,

      // X ATR (14) proxy (high-low avg)
      `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(
        OFFSET(DATA!$${highCol}$4, ${lastRow}-14, 0, 14) - OFFSET(DATA!$${lowCol}$4, ${lastRow}-14, 0, 14)
      )), 0), 2)`,

      // Y Bollinger %B proxy
      `=ROUND(IFERROR((($E${row}-$M${row}) / (4*STDEV(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)))) + 0.5, 0.5), 2)`,

      // Z TECH NOTES (Bloomberg narrative, pure formula)
      `=IF($B${row}="LOADING","LOADING",
        "VOLUME: RVOL "&TEXT($G${row},"0.00")&"x — "&IF($G${row}>=1.5,"above-average participation (conviction).","sub-average participation (weak sponsorship).")&CHAR(10)&
        "TREND REGIME: Price "&TEXT($E${row},"0.00")&" vs SMA200 "&TEXT($O${row},"0.00")&" — "&IF($E${row}>=$O${row},"long-term bullish structure intact.","risk-off regime below SMA200 (avoid chasing).")&CHAR(10)&
        "VOLATILITY / STRETCH: ATR(14) "&TEXT($X${row},"0.00")&"; SMA20 "&TEXT($M${row},"0.00")&"; Stretch="&TEXT(($E${row}-$M${row})/MAX(0.01,$X${row}),"0.0")&"x ATR — "&IF($E${row}>$M${row}+2*$X${row},"overextended (>+2x ATR).","within normal range (≤±2x ATR).")&CHAR(10)&
        "MOMENTUM: RSI(14) "&TEXT($P${row},"0.0")&" — "&IF($P${row}>=70,"overbought.",IF($P${row}<=30,"oversold.",IF($P${row}>=50,"positive bias.","negative bias.")))&
        " | MACD Hist "&TEXT($Q${row},"0.000")&" — "&IF($Q${row}>0,"positive momentum.","negative momentum.")&CHAR(10)&
        "TREND STRENGTH: ADX(14) "&TEXT($S${row},"0.0")&" — "&IF($S${row}<15,"no trend.",IF($S${row}<25,"weak trend.",IF($S${row}<40,"strong trend.","very strong trend.")))&
        " | Stoch %K "&TEXT($T${row},"0.0%")&" — "&IF($T${row}>=0.8,"overbought.",IF($T${row}<=0.2,"oversold.","neutral."))&CHAR(10)&
        "RISK/REWARD: "&TEXT($J${row},"0.00")&"x — "&IF($J${row}>=3,"institutional-grade asymmetry (≥3x).",IF($J${row}>=2,"acceptable tactical edge (≥2x).","sub-optimal payout (<2x)."))&CHAR(10)&
        "LEVELS: Support "&TEXT($U${row},"0.00")&" | Resistance "&TEXT($V${row},"0.00")&" | Target "&TEXT($W${row},"0.00")&"."
      )`,

      // AA FUND NOTES (keep yours, or placeholder)
      `=IF($B${row}="LOADING","LOADING",
        "FUND: "&$D${row}&" | REGIME: "&IF($E${row}>=$O${row},"Risk-On","Risk-Off")&CHAR(10)&
        "GATES: ADX "&TEXT($S${row},"0.0")&", R:R "&TEXT($J${row},"0.00")&"x, Stretch "&TEXT(($E${row}-$M${row})/MAX(0.01,$X${row}),"0.0")&"x ATR"&CHAR(10)&
        "DECISION WHY: "&IFS(
        $C${row}="Avoid",IF(OR($D${row}="ZOMBIE",$D${row}="BUBBLE"),"Blocked by fundamental risk ("&$D${row}&").","Blocked by Risk-Off regime (<SMA200)."),
        $C${row}="Trade Long","Allowed: no blockers + strength/edge gates pass (ADX/R:R).",
        $C${row}="Accumulate","Allowed: trend continuation in Risk-On regime with acceptable R:R.",
        $C${row}="Reduce (Overextended)","Timing fail: price stretched >2x ATR above SMA20.",
        $C${row}="Stop-Out","Invalidation: price below Support.",
        TRUE,"Neutral: not enough edge to justify risk."))`

      // AB restored separately
    ]);
  });

  if (tickers.length > 0) {
    // Write formulas to B..AA (27 cols)
    calc.getRange(3, 2, formulas.length, 26).setFormulas(formulas);

    // Restore LAST_STATE to AB (col 28)
    calc.getRange(3, 28, restoredStates.length, 1).setValues(restoredStates);
  }

  // Formats
  const lr = Math.max(calc.getLastRow(), 3);
  calc.setFrozenRows(2);

  // Bloomberg-style density: fixed row height + clip
  if (lr > 2) {
    calc.setRowHeights(3, lr - 2, 18);
    calc.getRange(3, 1, lr - 2, 28).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    // Notes columns clipped explicitly
    calc.getRange("Z3:AA" + lr).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  calc.getRange("F3:F").setNumberFormat("0.00%"); // Change %
  calc.getRange("I3:I").setNumberFormat("0.00%"); // ATH %
  calc.getRange("T3:T").setNumberFormat("0.00%"); // Stoch
  calc.getRange("Y3:Y").setNumberFormat("0.00%"); // %B

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

  // ✅ D1 checkbox to refresh dashboard
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

  // White border rows 1–3 (your request)
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

  // A..AA (27 cols)
  // A: Identity
  // B-D: Signaling (SIGNAL, FUND, DECISION)
  // E-G: Price/Volume
  // H-J: Performance
  // K-O: Trend
  // P-T: Momentum
  // U-Y: Levels/Risk
  // Z-AA: Notes
  styleGroup("A2:A2",   "IDENTITY",        "#263238");
  styleGroup("B2:D2",   "SIGNALING",       "#0D47A1");
  styleGroup("E2:G2",   "PRICE / VOLUME",  "#1B5E20");
  styleGroup("H2:J2",   "PERFORMANCE",     "#004D40");
  styleGroup("K2:O2",   "TREND",           "#2E7D32");
  styleGroup("P2:T2",   "MOMENTUM",        "#33691E");
  styleGroup("U2:Y2",   "LEVELS / RISK",   "#B71C1C");
  styleGroup("Z2:AA2",  "NOTES",           "#212121");

  // Row 2 wrap (your request)
  dashboard.getRange("A2:AA2").setWrap(true);

  // ============================================================
  // ROW 3 — Column headers (Dashboard order; C/D swapped)
  // ============================================================
  const headers = [[
    "Ticker",           // A
    "SIGNAL",           // B (CALC B)
    "FUNDAMENTAL",      // C (CALC D)  <-- swapped
    "DECISION",         // D (CALC C)  <-- swapped
    "Price",            // E (CALC E)
    "Change %",         // F (CALC F)
    "Vol Trend",        // G (CALC G)
    "ATH (TRUE)",       // H (CALC H)
    "ATH Diff %",       // I (CALC I)
    "R:R Quality",      // J (CALC J)
    "Trend Score",      // K (CALC K)
    "Trend State",      // L (CALC L)
    "SMA 20",           // M (CALC M)
    "SMA 50",           // N (CALC N)
    "SMA 200",          // O (CALC O)
    "RSI",              // P (CALC P)
    "MACD Hist",        // Q (CALC Q)
    "Divergence",       // R (CALC R)
    "ADX (14)",         // S (CALC S)
    "Stoch %K (14)",    // T (CALC T)
    "Support",          // U (CALC U)
    "Resistance",       // V (CALC V)
    "Target (3:1)",     // W (CALC W)
    "ATR (14)",         // X (CALC X)
    "Bollinger %B",     // Y (CALC Y)
    "TECH NOTES",       // Z (CALC Z)
    "FUND NOTES"        // AA (CALC AA)
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
  // ROW 4 — Hardened filter formula (INPUT-driven; token-exact + ALL override)
  // NOTE: Uses INPUT!A as the "master ticker list to include"
  // and INPUT!B3:B + INPUT!C3:C as token columns for sector/industry (as per your baseline).
  // ============================================================
  const filterFormula =
    '=IFERROR(' +
      'SORT(' +
        'FILTER({' +
          'CALCULATIONS!$A$3:$A,' +    // Ticker
          'CALCULATIONS!$B$3:$B,' +    // SIGNAL
          'CALCULATIONS!$D$3:$D,' +    // FUNDAMENTAL (swapped)
          'CALCULATIONS!$C$3:$C,' +    // DECISION (swapped)
          'CALCULATIONS!$E$3:$E,' +    // Price
          'CALCULATIONS!$F$3:$F,' +    // Change %
          'CALCULATIONS!$G$3:$G,' +    // Vol Trend
          'CALCULATIONS!$H$3:$H,' +    // ATH
          'CALCULATIONS!$I$3:$I,' +    // ATH Diff %
          'CALCULATIONS!$J$3:$J,' +    // R:R
          'CALCULATIONS!$K$3:$K,' +    // Trend Score
          'CALCULATIONS!$L$3:$L,' +    // Trend State
          'CALCULATIONS!$M$3:$M,' +    // SMA20
          'CALCULATIONS!$N$3:$N,' +    // SMA50
          'CALCULATIONS!$O$3:$O,' +    // SMA200
          'CALCULATIONS!$P$3:$P,' +    // RSI
          'CALCULATIONS!$Q$3:$Q,' +    // MACD
          'CALCULATIONS!$R$3:$R,' +    // Divergence
          'CALCULATIONS!$S$3:$S,' +    // ADX
          'CALCULATIONS!$T$3:$T,' +    // Stoch
          'CALCULATIONS!$U$3:$U,' +    // Support
          'CALCULATIONS!$V$3:$V,' +    // Resistance
          'CALCULATIONS!$W$3:$W,' +    // Target
          'CALCULATIONS!$X$3:$X,' +    // ATR
          'CALCULATIONS!$Y$3:$Y,' +    // %B
          'CALCULATIONS!$Z$3:$Z,' +    // TECH NOTES
          'CALCULATIONS!$AA$3:$AA' +   // FUND NOTES
        '},' +
        'ISNUMBER(MATCH(' +
          'CALCULATIONS!$A$3:$A,' +
          'FILTER(INPUT!$A$3:$A,' +
            'INPUT!$A$3:$A<>"",' +

            // Sector token filter (INPUT!B1 tokens vs INPUT!B3:B)
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

            // Industry token filter (INPUT!C1 tokens vs INPUT!C3:C)
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
        ',5,FALSE' +   // sort by Price desc
      '),' +
    '"No Matches Found")';

  dashboard.getRange("A4").setFormula(filterFormula);
  SpreadsheetApp.flush();

  // ============================================================
  // FORMATTING GOVERNANCE (Bloomberg dense)
  // ============================================================
  dashboard.setFrozenRows(3);
  dashboard.setFrozenColumns(1);

  // Uniform widths for A..Y (1..25): B/C/D same as all other data columns
  for (let c = 1; c <= 25; c++) dashboard.setColumnWidth(c, 90);

  // Notes wider
  dashboard.setColumnWidth(26, 420); // Z Tech Notes
  dashboard.setColumnWidth(27, 420); // AA Fund Notes

  // Row heights + CLIP
  const lastRow = Math.max(dashboard.getLastRow(), 4);
  if (lastRow >= 4) {
    dashboard.setRowHeights(4, lastRow - 3, 18);
    dashboard.getRange(4, 1, lastRow - 3, 25).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    dashboard.getRange(4, 26, lastRow - 3, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Row 3 wrap (explicit)
  dashboard.getRange("A3:AA3").setWrap(true);

  // Alignment
  dashboard.getRange(4, 1, lastRow - 3, 27)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");

  // Grid borders (black)
  dashboard.getRange(1, 1, lastRow, 27)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Number formats
  dashboard.getRange("F4:F" + lastRow).setNumberFormat("0.00%"); // Change %
  dashboard.getRange("I4:I" + lastRow).setNumberFormat("0.00%"); // ATH Diff %
  dashboard.getRange("T4:T" + lastRow).setNumberFormat("0.00%"); // Stoch (0..1)
  dashboard.getRange("Y4:Y" + lastRow).setNumberFormat("0.00%"); // %B (0..1)

  // ============================================================
  // CONDITIONAL FORMATTING (Bloomberg palette; includes C/D swap)
  // ============================================================
  const rules = [];

  // Negative drag: Change%, ATH%, MACD => red + bold
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#B71C1C")
      .setBold(true)
      .setRanges([
        dashboard.getRange("F4:F" + lastRow), // Change %
        dashboard.getRange("I4:I" + lastRow), // ATH %
        dashboard.getRange("Q4:Q" + lastRow)  // MACD
      ])
      .build()
  );

  // RSI bands (P)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$P4>=70")
    .setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange("P4:P" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$P4<=30")
    .setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("P4:P" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($P4>=50,$P4<70)")
    .setFontColor("#1B5E20")
    .setRanges([dashboard.getRange("P4:P" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($P4>30,$P4<50)")
    .setFontColor("#E65100")
    .setRanges([dashboard.getRange("P4:P" + lastRow)]).build());

  // ADX strength (S)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$S4>=25")
    .setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("S4:S" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$S4<15")
    .setFontColor("#616161")
    .setRanges([dashboard.getRange("S4:S" + lastRow)]).build());

  // Stoch extremes (T)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$T4>=0.8")
    .setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange("T4:T" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$T4<=0.2")
    .setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("T4:T" + lastRow)]).build());

  // %B extremes (Y)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$Y4>=0.8")
    .setFontColor("#B71C1C")
    .setRanges([dashboard.getRange("Y4:Y" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$Y4<=0.2")
    .setFontColor("#1B5E20")
    .setRanges([dashboard.getRange("Y4:Y" + lastRow)]).build());

  // SIGNAL (B) backgrounds
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Breakout|Trend Continuation|RVOL")')
    .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("B4:B" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Mean Reversion|Bounce|Oversold|Overbought")')
    .setBackground("#FFF8E1").setFontColor("#E65100").setBold(true)
    .setRanges([dashboard.getRange("B4:B" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Range|Chop|Hold")')
    .setBackground("#F5F5F5").setFontColor("#616161")
    .setRanges([dashboard.getRange("B4:B" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Risk-Off|Stop")')
    .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange("B4:B" + lastRow)]).build());

  // DECISION is column D (swapped)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($D4,"Trade|Accumulate|Buy")')
    .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("D4:D" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($D4,"Reduce|Trim")')
    .setBackground("#FFF8E1").setFontColor("#E65100").setBold(true)
    .setRanges([dashboard.getRange("D4:D" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($D4,"Hold|Monitor|Wait")')
    .setBackground("#F5F5F5").setFontColor("#616161")
    .setRanges([dashboard.getRange("D4:D" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($D4,"Avoid|Stop")')
    .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange("D4:D" + lastRow)]).build());

  // FUNDAMENTAL is column C (swapped)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($C4,"VALUE|GEM")')
    .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("C4:C" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($C4,"ZOMBIE|BUBBLE")')
    .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange("C4:C" + lastRow)]).build());

  dashboard.setConditionalFormatRules(rules);
}

function applyBloombergFormattingDashboard_(dashboard) {
  const lastRow = dashboard.getLastRow();
  if (lastRow < 4) return;

  const rules = [];

  // Negative % + MACD negative in red
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#B71C1C")
      .setBold(true)
      .setRanges([
        dashboard.getRange("F4:F" + lastRow), // Change %
        dashboard.getRange("I4:I" + lastRow), // ATH %
        dashboard.getRange("Q4:Q" + lastRow)  // MACD Hist
      ])
      .build()
  );

  // RSI bands (P)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$P4>=70")
    .setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange("P4:P" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$P4<=30")
    .setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("P4:P" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($P4>=50,$P4<70)")
    .setFontColor("#1B5E20")
    .setRanges([dashboard.getRange("P4:P" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($P4>30,$P4<50)")
    .setFontColor("#E65100")
    .setRanges([dashboard.getRange("P4:P" + lastRow)]).build());

  // ADX (S): >=25 tradeable, <15 range
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$S4>=25")
    .setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("S4:S" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$S4<15")
    .setFontColor("#616161")
    .setRanges([dashboard.getRange("S4:S" + lastRow)]).build());

  // Stoch (T) extremes
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$T4>=0.8")
    .setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange("T4:T" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$T4<=0.2")
    .setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("T4:T" + lastRow)]).build());

  // %B (Y) extremes
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$Y4>=0.8")
    .setFontColor("#B71C1C")
    .setRanges([dashboard.getRange("Y4:Y" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$Y4<=0.2")
    .setFontColor("#1B5E20")
    .setRanges([dashboard.getRange("Y4:Y" + lastRow)]).build());

  // SIGNAL (B) terminal backgrounds
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Breakout|Trend Continuation")')
    .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("B4:B" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Mean Reversion")')
    .setBackground("#FFF8E1").setFontColor("#E65100").setBold(true)
    .setRanges([dashboard.getRange("B4:B" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Range-Bound")')
    .setBackground("#F5F5F5").setFontColor("#616161")
    .setRanges([dashboard.getRange("B4:B" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($B4,"Risk-Off|Stop-Out")')
    .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange("B4:B" + lastRow)]).build());

  // DECISION (C)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($C4,"Accumulate|Trade Long")')
    .setBackground("#E8F5E9").setFontColor("#1B5E20").setBold(true)
    .setRanges([dashboard.getRange("C4:C" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($C4,"Reduce")')
    .setBackground("#FFF8E1").setFontColor("#E65100").setBold(true)
    .setRanges([dashboard.getRange("C4:C" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($C4,"Hold / Monitor")')
    .setBackground("#F5F5F5").setFontColor("#616161")
    .setRanges([dashboard.getRange("C4:C" + lastRow)]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($C4,"Avoid|Stop-Out")')
    .setBackground("#FFEBEE").setFontColor("#B71C1C").setBold(true)
    .setRanges([dashboard.getRange("C4:C" + lastRow)]).build());

  dashboard.setConditionalFormatRules(rules);
  applyBloombergFormattingDashboard_(dashboard);
}


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
  sh.setColumnWidth(1, 85);    // A ~10 chars
  sh.setColumnWidth(2, 125);   // B ✅ ~15 chars
  sh.setColumnWidth(3, 520);   // C Tech Notes
  sh.setColumnWidth(4, 520);   // D Fund Notes
  sh.setColumnWidth(5, 18);    // spacer

  // ✅ B must be left-aligned + wrapped
  sh.getRange("B:B").setHorizontalAlignment("left").setWrap(true);
  sh.getRange("A:A").setHorizontalAlignment("left");

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

  const listValidation = (arr) =>
    SpreadsheetApp.newDataValidation().requireValueInList(arr, true).build();

  // B2/B3/B4 start at 0; default values: 1,0,0
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

  // Row 7 reserved empty
  sh.getRange("A7:E7").clearContent();
  sh.setRowHeight(7, 18);

  // ------------------------------------------------------------
  // Sidebar (starts row 8) + add R:R
  // ------------------------------------------------------------
  const t = "$A$1";
  const startRow = 8;

  const IDX = (colLetter, fallback) =>
    `=IFERROR(INDEX(CALCULATIONS!$${colLetter}$3:$${colLetter}, MATCH(${t}, CALCULATIONS!$A$3:$A, 0)), ${fallback})`;

  const rows = [
    ["SIGNAL",   IDX("B", '"Wait"')],
    ["FUND",     IDX("D", '"-"')],          // swapped earlier
    ["DECISION", IDX("C", '"-"')],
    ["PRICE",    `=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`],
    ["CHG%",     `=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`],
    ["R:R",      IDX("J", "0")],            // ✅ R:R added
    ["", ""],

    ["[ PERFORMANCE ]", ""],
    ["VOL TREND", IDX("G", "0")],
    ["ATH",       IDX("H", "0")],
    ["ATH %",     IDX("I", "0")],
    ["", ""],

    ["[ TREND ]", ""],
    ["SMA 20", IDX("M", "0")],
    ["SMA 50", IDX("N", "0")],
    ["SMA 200",IDX("O", "0")],
    ["RSI",    IDX("P", "50")],
    ["MACD",   IDX("Q", "0")],
    ["DIV",    IDX("R", '"-"')],
    ["ADX",    IDX("S", "0")],
    ["STO",    IDX("T", "0")],
    ["", ""],

    ["[ LEVELS ]", ""],
    // Use labels that chart reader understands (SUPPORT/RESISTANCE)
    ["SUPPORT",    IDX("U", "0")],
    ["RESISTANCE", IDX("V", "0")],
    ["TARGET",     IDX("W", "0")],
    ["ATR",        IDX("X", "0")],
    ["%B",         IDX("Y", "0")]
  ];

  sh.getRange(startRow, 1, rows.length, 1).setValues(rows.map(r => [r[0]])).setFontWeight("bold");
  sh.getRange(startRow, 2, rows.length, 1).setFormulas(rows.map(r => [r[1]]));

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
  // Number formats (robust by row numbers given this fixed sidebar layout)
  // ------------------------------------------------------------
  // PRICE row = startRow+3
  sh.getRange(`B${startRow + 3}`).setNumberFormat("#,##0.00"); // PRICE
  sh.getRange(`B${startRow + 4}`).setNumberFormat("0.00%");    // CHG%
  sh.getRange(`B${startRow + 5}`).setNumberFormat("0.00");     // R:R

  // PERFORMANCE
  sh.getRange(`B${startRow + 8}`).setNumberFormat("0.00");     // VOL TREND
  sh.getRange(`B${startRow + 9}`).setNumberFormat("#,##0.00"); // ATH
  sh.getRange(`B${startRow + 10}`).setNumberFormat("0.00%");   // ATH %

  // TREND
  // SMA rows
  sh.getRange(`B${startRow + 13}:B${startRow + 15}`).setNumberFormat("#,##0.00");
  sh.getRange(`B${startRow + 16}`).setNumberFormat("0.00");    // RSI
  sh.getRange(`B${startRow + 17}`).setNumberFormat("0.000");   // MACD
  sh.getRange(`B${startRow + 19}`).setNumberFormat("0.00");    // ADX
  sh.getRange(`B${startRow + 20}`).setNumberFormat("0.00%");   // STO (0..1)

  // LEVELS
  sh.getRange(`B${startRow + 23}:B${startRow + 26}`).setNumberFormat("#,##0.00"); // SUPPORT/RES/TARGET/ATR
  sh.getRange(`B${startRow + 27}`).setNumberFormat("0.00%"); // %B (0..1)

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


function generateReferenceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "REFERENCE_GUIDE";
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear().clearFormats();

  // -----------------------------
  // CONTENT (tight, document-style)
  // -----------------------------
  const rows = [];

  // Title
  rows.push(["INSTITUTIONAL TERMINAL — REFERENCE GUIDE", "", "", ""]);
  rows.push(["Dashboard/Chart vocabulary, column definitions, indicator glossary, and action playbook.", "", "", ""]);

  // Section: Column Map
  rows.push(["", "", "", ""]);
  rows.push(["1) DASHBOARD COLUMN DEFINITIONS (TECHNICAL)", "", "", ""]);
  rows.push(["COLUMN", "WHAT IT IS", "HOW IT IS USED", "USER ACTION"]);

  const cols = [
    ["Ticker", "Symbol (key)", "Join key across DATA/CALCULATIONS/CHART; drives lookups + chart feed", "Select ticker for chart; review SIGNAL + DECISION first, then NOTES."],
    ["SIGNAL", "Technical state label (rules engine output)", "Primary technical setup classifier (regime + levels + momentum + trend strength)", "Treat as “setup type” (breakout / trend / mean-rev / range / risk-off / stop-out)."],
    ["FUNDAMENTAL", "Valuation/quality proxy (EPS + P/E heuristic)", "Risk filter that can block/override technical signals (e.g., Avoid on BUBBLE/ZOMBIE)", "Avoid ZOMBIE/BUBBLE; prefer VALUE/FAIR for technical longs."],
    ["DECISION", "Confluence action label (SIGNAL + FUNDAMENTAL + R:R + ADX + stretch gates)", "Final action layer used for execution; the only label meant to be acted on", "Act on DECISION; validate in TECH NOTES / FUND NOTES."],
    ["Price", "Live last price (GOOGLEFINANCE)", "Used in regime tests, distance-to-levels, ATR stretch, R:R evaluation", "Confirm price vs SMA200 and vs Support/Resistance."],
    ["Change %", "Daily % change", "Tape/context only (not a signal by itself)", "Use for context; do not chase without a setup."],
    ["Vol Trend", "Relative volume proxy (current vol vs avg vol)", "Conviction filter (especially for breakouts)", "Prefer ≥1.5x for breakout entries; ignore low RVOL breakouts."],
    ["ATH (TRUE)", "All-time high reference", "Context: near ATH implies supply zone / price discovery", "Avoid chasing into resistance without RVOL + trend strength."],
    ["ATH Diff %", "Distance from ATH", "Classifies “near ATH breakout” vs “deep pullback” states", "Use with regime + levels; not a signal alone."],
    ["R:R Quality", "Reward/Risk ratio = (Resistance − Price) / (Price − Support)", "Quality gate for entries; used in DECISION gating", "≥3 strong; 1.5–3 acceptable; <1.5 poor (avoid)."],
    ["Trend Score", "★ count (Price above key SMAs)", "Quick structure read (strength of trend stack)", "3★ = best structure; 0★ = weak / avoid trend chasing."],
    ["Trend State", "Regime state derived from SMA200", "Defines risk-on vs risk-off posture", "If below SMA200: treat as Risk-Off; avoid chasing."],
    ["SMA 20", "Short-term mean", "Mean-reversion anchor + stretch baseline", "Watch distance in ATR terms; avoid buying if >2×ATR above SMA20."],
    ["SMA 50", "Medium-term trend reference", "Trend confirmation / continuation filter", "Prefer price above for momentum longs."],
    ["SMA 200", "Long-term regime line", "Primary risk-on/risk-off filter", "If below: reduce exposure; only tactical trades."],
    ["RSI", "Momentum oscillator (0–100)", "Bias + extreme condition detector", "<30 oversold; >70 overbought; ~50 bias line."],
    ["MACD Hist", "Impulse gauge (MACD line − signal line)", "Confirms bullish/bearish momentum impulse", "Positive supports longs; negative supports risk-off / fades."],
    ["Divergence", "Price vs RSI divergence heuristic", "Early reversal warning (not a trigger alone)", "Bull div supports bounce; bear div warns on rallies."],
    ["ADX (14)", "Trend strength metric (0–50+)", "Separates trend vs range (chop) regime", "<15 range/no trend; 15–25 forming; ≥25 trend."],
    ["Stoch %K (14)", "Fast oscillator (0–1)", "Timing layer for mean reversion inside regimes", "≤0.20 oversold; ≥0.80 overbought."],
    ["Support", "20-day min low proxy", "Risk line / invalidation zone", "If Price < Support: stop-out / avoid."],
    ["Resistance", "50-day max high proxy", "Ceiling / target reference", "Break above + RVOL supports breakout entry."],
    ["Target (3:1)", "Price + 3×(Price−Support)", "Tactical profit planning reference", "Use as planning marker, not a forecast."],
    ["ATR (14)", "Volatility proxy (average range)", "Defines stretch, stop width context, sizing sensitivity", "High ATR = bigger stops + smaller size; avoid overstretch entries."],
    ["Bollinger %B", "Band position proxy (0–1)", "Compression/extension context (with ADX)", "Low %B + low ADX = range/chop; high %B = extension risk."],
    ["TECH NOTES", "Narrative diagnostics with indicator values", "Explains why SIGNAL/DECISION fired", "Read before acting; it is your justification layer."],
    ["FUND NOTES", "Narrative valuation + regime alignment", "Explains FUNDAMENTAL classification and risk posture", "Avoid weak fundamentals in risk-off regimes."]
  ];
  cols.forEach(r => rows.push(r));

  // ------------------------------------------------------------
  // NEW: Indicator Glossary (requested)
  // ------------------------------------------------------------
  rows.push(["", "", "", ""]);
  rows.push(["1A) INDICATOR GLOSSARY (TERMS USED IN NOTES / SIGNAL LOGIC)", "", "", ""]);
  rows.push(["TERM", "DEFINITION", "HOW TO INTERPRET", "ENGINE USAGE"]);
  const glossary = [
    ["RVOL (Vol Trend)", "Relative volume proxy = latest volume / avg(20)", "≥1.5 suggests institutional participation", "Breakout gating + conviction language in TECH NOTES."],
    ["ADX (14)", "Average Directional Index: trend strength (not direction)", "<15 chop; 15–25 forming; ≥25 trend", "Range-Bound signal if very low; trend confidence if higher."],
    ["ATR (14)", "Average True Range proxy: typical daily range", "Higher ATR = higher volatility / wider noise band", "Overextension gate: Price > SMA20 + 2×ATR → Reduce (Overextended)."],
    ["RSI (14)", "Relative Strength Index: momentum oscillator", "<30 oversold; >70 overbought; 50 bias line", "Used to label bias and mean-rev context in TECH NOTES."],
    ["MACD Hist", "MACD line minus signal line (impulse)", ">0 bullish impulse; <0 bearish impulse", "Breakout/trend continuation gates require MACD > 0."],
    ["Stoch %K", "Fast stochastic (0–1) vs recent high/low window", "≤0.20 oversold; ≥0.80 overbought", "Mean Reversion (Oversold/Overbought) classification."],
    ["%B", "Band position proxy (0–1)", "Near 0 = low band; near 1 = high band", "Context only; helps explain compression/extension."],
    ["SMA stack", "SMA20/SMA50/SMA200 trend structure", "More SMAs below price = stronger trend stack", "Trend Score (★) + Trend State (SMA200 regime)."],
    ["R:R Quality", "(Resistance−Price)/(Price−Support)", "Higher is better payoff vs risk", "DECISION gates: Breakout requires R:R ≥ 1.5; Trend cont. requires R:R ≥ 1.3."],
    ["Support/Resistance", "Statistical proxies (20d low, 50d high)", "Define risk line and target ceiling", "Stop-Out if Price < Support; Breakout context near Resistance."]
  ];
  glossary.forEach(r => rows.push(r));

  // ------------------------------------------------------------
  // SECTION: SIGNAL vocabulary — MUST MATCH FORMULA OUTPUTS
  // SIGNAL formula outputs:
  // Stop-Out
  // Risk-Off (Below SMA200)
  // Range-Bound (Low ADX)
  // Breakout (High Volume)
  // Mean Reversion (Oversold)
  // Mean Reversion (Overbought)
  // Trend Continuation
  // Hold / Monitor
  // LOADING
  // ------------------------------------------------------------
  rows.push(["", "", "", ""]);
  rows.push(["2) SIGNAL — FULL VOCABULARY (ALIGNED TO ENGINE OUTPUT)", "", "", ""]);
  rows.push(["SIGNAL VALUE", "TECHNICAL DEFINITION", "WHEN IT TRIGGERS (ENGINE)", "EXPECTED USER ACTION"]);

  const signal = [
    ["LOADING", "Data not ready / missing", "Price cell blank/0", "Refresh; do not act."],
    ["Stop-Out", "Broken structure; invalidation event", "Price < Support", "Exit / stand aside. Do not average down."],
    ["Risk-Off (Below SMA200)", "Bear regime filter", "Price < SMA200", "Avoid trend chasing; only tactical trades with strict risk if at all."],
    ["Range-Bound (Low ADX)", "No trend / chop regime", "ADX < 15", "Range tactics only; smaller size; tighter targets; avoid breakout chasing."],
    ["Breakout (High Volume)", "Breakout attempt with sponsorship + impulse", "RVOL ≥ 1.5 AND Price ≥ Resistance*0.995 AND MACD>0 AND ADX ≥ 18", "Actionable only with acceptable R:R and non-toxic fundamentals."],
    ["Mean Reversion (Oversold)", "Oversold timing inside tradable conditions", "Stoch ≤ 0.20 AND Price > Support AND ADX ≥ 18", "Tactical long bounce; stop below support; target resistance/3:1."],
    ["Mean Reversion (Overbought)", "Overbought timing near ceiling", "Stoch ≥ 0.80 AND Price ≥ Resistance*0.97", "Do not initiate new longs; take profits / wait for pullback."],
    ["Trend Continuation", "Bull structure + impulse + trend strength", "Price > SMA200 AND MACD>0 AND ADX ≥ 18", "Prefer accumulate/add on pullbacks; manage to levels."],
    ["Hold / Monitor", "Neutral / mixed; no defined edge", "Fallback state when none of the above triggers", "Wait; do not force trades."]
  ];
  signal.forEach(r => rows.push(r));

  // ------------------------------------------------------------
  // SECTION: FUNDAMENTAL vocabulary — MUST MATCH FORMULA OUTPUTS
  // FUNDAMENTAL formula outputs:
  // ZOMBIE
  // PRICED FOR PERFECTION
  // VALUE
  // BUBBLE
  // FAIR
  // ------------------------------------------------------------
  rows.push(["", "", "", ""]);
  rows.push(["3) FUNDAMENTAL — FULL VOCABULARY (ALIGNED TO ENGINE OUTPUT)", "", "", ""]);
  rows.push(["FUNDAMENTAL VALUE", "WHAT IT MEANS (ENGINE)", "RISK PROFILE", "EXPECTED USER ACTION"]);

  const fund = [
    ["VALUE", "EPS>0 and P/E < 25 (heuristic)", "Lower valuation risk", "Prefer for Breakout/Trend Continuation if technical gates pass."],
    ["FAIR", "Default / neutral valuation state", "Neutral risk", "Trade based on technicals; require R:R gate."],
    ["PRICED FOR PERFECTION", "EPS>0 and P/E > 50", "Multiple-compression risk", "Only take best technical setups; avoid in Risk-Off."],
    ["BUBBLE", "P/E > 30 and EPS < 0.1 (heuristic fragility)", "High downside on sentiment shift", "Avoid longs; only tactical trades with tight risk (if any)."],
    ["ZOMBIE", "EPS < 0", "High blow-up / quality risk", "Avoid; engine will block to Avoid in DECISION."]
  ];
  fund.forEach(r => rows.push(r));

  // ------------------------------------------------------------
  // SECTION: DECISION vocabulary — MUST MATCH FORMULA OUTPUTS
  // DECISION formula outputs:
  // Stop-Out
  // Avoid
  // Trade Long
  // Accumulate
  // Reduce (Overextended)
  // Hold / Monitor
  // LOADING
  // ------------------------------------------------------------
  rows.push(["", "", "", ""]);
  rows.push(["4) DECISION — FULL VOCABULARY (ALIGNED TO ENGINE OUTPUT)", "", "", ""]);
  rows.push(["DECISION VALUE", "WHY IT HAPPENS (ENGINE)", "RISK GATES (ENGINE)", "EXPECTED USER ACTION"]);

  const decision = [
    ["LOADING", "Data not ready", "Price blank/0", "Refresh; do not act."],
    ["Stop-Out", "SIGNAL=Stop-Out", "Price < Support", "Exit / stand aside; wait for base/reclaim."],
    ["Avoid", "Fundamental block or regime block", "FUNDAMENTAL in {ZOMBIE, BUBBLE} OR SIGNAL=Risk-Off (Below SMA200)", "Remove from active trade list; reassess only after regime improves."],
    ["Trade Long", "High-conviction breakout or mean reversion", "Breakout + R:R ≥ 1.5 + ADX ≥ 20 OR Mean Reversion + R:R ≥ 1.2 + ADX ≥ 18", "Tactical entry with stop under support; plan exit to resistance/target."],
    ["Accumulate", "Trend continuation with acceptable edge", "SIGNAL=Trend Continuation AND R:R ≥ 1.3 AND ADX ≥ 18", "Scale in on pullbacks; avoid chasing overextended price."],
    ["Reduce (Overextended)", "Stretch condition detected", "ATR>0 AND Price > SMA20 + 2×ATR", "Trim/avoid new entries; wait for mean reversion."],
    ["Hold / Monitor", "No actionable edge or range regime", "Range-Bound OR fallback state", "Wait; monitor levels; size down in chop."]
  ];
  decision.forEach(r => rows.push(r));

  // ------------------------------------------------------------
  // Section: Practical Playbook (kept, but aligned wording)
  // ------------------------------------------------------------
  rows.push(["", "", "", ""]);
  rows.push(["5) QUICK PLAYBOOK (HOW TO USE THE TERMINAL)", "", "", ""]);
  rows.push(["RULE", "WHY", "WHAT TO LOOK FOR", "WHAT TO AVOID"]);
  rows.push(["Trend continuation plays", "Highest expectancy when trend strength exists", "Trend Continuation + ADX≥25 + MACD>0 + Price>SMA200", "Chasing when Price>SMA20+2×ATR (overextended)."]);
  rows.push(["Breakout plays", "Requires sponsorship/impulse", "Breakout (High Volume) + RVOL≥1.5 + R:R acceptable", "Low RVOL breakouts; breakouts in Risk-Off regime."]);
  rows.push(["Mean reversion plays", "Works best at extremes with defined risk", "Mean Reversion (Oversold) near Support + ADX≥18", "Buying without a stop; averaging down after Stop-Out."]);
  rows.push(["Range regime discipline", "Low ADX means no trend", "Range-Bound (Low ADX): trade levels only", "Mid-range entries with poor R:R."]);
  rows.push(["R:R gating", "Prevents low-quality setups", "R:R ≥ 1.5 (breakouts) or ≥ 1.2 (mean-rev)", "Forcing trades when R:R is poor."]);

  // Write content
  sh.getRange(1, 1, rows.length, 4).setValues(rows);

  // -----------------------------
  // STYLING (Bloomberg-like, dense) — preserved
  // -----------------------------
  sh.setColumnWidth(1, 210);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 320);
  sh.setColumnWidth(4, 260);

  sh.setRowHeights(1, Math.min(rows.length, 500), 18);
  sh.setFrozenRows(3);

  // Title styling
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
    if (/^\d\)/.test(v) || /^1A\)/.test(v)) {
      sh.getRange(r, 1, 1, 4).merge()
        .setBackground("#212121").setFontColor("white")
        .setFontWeight("bold").setFontSize(10)
        .setHorizontalAlignment("left");
    }
  }

  // Table header rows
  for (let r = 1; r <= rows.length; r++) {
    const a = String(sh.getRange(r, 1).getValue() || "").trim();
    if (["COLUMN", "SIGNAL VALUE", "FUNDAMENTAL VALUE", "DECISION VALUE", "RULE", "TERM"].includes(a)) {
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

  ss.toast("REFERENCE_GUIDE updated (vocabulary aligned to live SIGNAL/FUNDAMENTAL/DECISION engine + glossary).", "✅ DONE", 3);
}