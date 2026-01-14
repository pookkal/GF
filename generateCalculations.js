/**
 * generateCalculations.js
 * Optimized version of generateCalculationsSheet with progressive loading and error handling
 * Preserves ALL 35 columns (A-AI) and exact formula references from Code.js
 */

function generateCalculationsSheet() {
  const startTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Validate required sheets
    const dataSheet = ss.getSheetByName("DATA");
    const inputSheet = ss.getSheetByName("INPUT");
    
    if (!dataSheet || !inputSheet) {
      ss.toast('Required sheets (DATA or INPUT) not found.', '❌ Error', 3);
      return;
    }

    const tickers = getCleanTickers(inputSheet);
    if (!tickers || tickers.length === 0) {
      ss.toast('No tickers found in INPUT sheet.', '⚠️ Warning', 3);
      return;
    }

    let calc = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");

    // Locale separator: US typically ","; EU typically ";"
    const locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
    const SEP = (/^(en|en_)/.test(locale)) ? "," : ";";

    // Clear existing content
    calc.clear().clearFormats();

    // Ensure sheet has enough columns (35 total: A-AI)
    const maxCols = calc.getMaxColumns();
    if (maxCols < 35) {
      calc.insertColumnsAfter(maxCols, 35 - maxCols);
    }

    // PHASE 1: Setup headers
    setupHeaders(calc, ss, SEP);
    
    // PHASE 2: Write tickers (progressive)
    writeTickers(calc, tickers);
    
    SpreadsheetApp.flush();
    
    // PHASE 3: Write formulas in batches
    writeFormulas(calc, tickers, SEP);
    
    const elapsed = ((new Date() - startTime) / 1000).toFixed(2);
    ss.toast(`✓ CALCULATIONS sheet generated in ${elapsed}s`, 'Success', 3);
    
  } catch (error) {
    ss.toast(`Failed to generate CALCULATIONS: ${error.message}`, '❌ Error', 5);
    Logger.log(`Error in generateCalculationsSheet: ${error.stack}`);
  }
}

function setupHeaders(calc, ss, SEP) {
  // ROW 1: GROUP HEADERS (MERGED) + timestamp
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

  styleGroup("A1:A1", "IDENTITY", "#263238");
  styleGroup("B1:D1", "SIGNALING", "#0D47A1");
  styleGroup("E1:G1", "PRICE / VOLUME", "#1B5E20");
  styleGroup("H1:J1", "PERFORMANCE", "#004D40");
  styleGroup("K1:O1", "TREND", "#2E7D32");
  styleGroup("P1:T1", "MOMENTUM", "#33691E");
  styleGroup("U1:Y1", "LEVELS / RISK", "#B71C1C");
  styleGroup("Z1:Z1", "INSTITUTIONAL", "#4A148C");
  styleGroup("AA1:AB1", "NOTES", "#212121");
  styleGroup("AC1:AH1", "ENHANCED PATTERNS", "#6A1B9A");

  calc.getRange("AE1")
    .setValue(syncTime)
    .setBackground("#000000")
    .setFontColor("#00FF00")
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // ROW 2: COLUMN HEADERS (35 columns total)
  const headers = [[
    "Ticker", "SIGNAL", "FUNDAMENTAL", "DECISION", "Price", "Change %", "Vol Trend", "ATH (TRUE)", "ATH Diff %", "R:R Quality",
    "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "RSI", "MACD Hist", "Divergence", "ADX (14)", "Stoch %K (14)",
    "Support", "Resistance", "Target (3:1)", "ATR (14)", "Bollinger %B", "POSITION SIZE", "TECH NOTES", "FUND NOTES",
    "VOL REGIME", "ATH ZONE", "BBP SIGNAL", "PATTERNS", "ATR STOP", "ATR TARGET", "LAST STATE"
  ]];

  calc.getRange(2, 1, 1, 35)
    .setValues(headers)
    .setBackground("#111111")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
}

function writeTickers(calc, tickers) {
  if (tickers.length > 0) {
    calc.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  }
  SpreadsheetApp.flush();
}

function writeFormulas(calc, tickers, SEP) {
  const BLOCK = 7; // DATA block width (must match generateDataSheet)
  const BATCH_SIZE = 10; // Process 10 tickers at a time
  
  // Check if long-term signal mode is enabled
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('INPUT');
  const useLongTermSignal = inputSheet.getRange('E2').getValue() === true;
  
  // Process tickers in batches for better performance
  for (let batchStart = 0; batchStart < tickers.length; batchStart += BATCH_SIZE) {
    const batchEnd = Math.min(batchStart + BATCH_SIZE, tickers.length);
    const batchFormulas = [];
    
    for (let i = batchStart; i < batchEnd; i++) {
      const ticker = tickers[i];
      const row = i + 3;
      const formulas = generateTickerFormulas(ticker, row, i, BLOCK, SEP, useLongTermSignal);
      batchFormulas.push(formulas);
    }
    
    // Write batch (columns B-AI, 34 columns)
    if (batchFormulas.length > 0) {
      calc.getRange(batchStart + 3, 2, batchFormulas.length, 34).setFormulas(batchFormulas);
      SpreadsheetApp.flush();
    }
  }
}

function generateTickerFormulas(ticker, row, index, BLOCK, SEP, useLongTermSignal) {
  const t = String(ticker || "").trim().toUpperCase();
  
  // DATA block start (each ticker is 7 cols in DATA)
  const tDS = (index * BLOCK) + 1;
  const dateCol = columnToLetter(tDS + 0);
  const openCol = columnToLetter(tDS + 1);
  const highCol = columnToLetter(tDS + 2);
  const lowCol = columnToLetter(tDS + 3);
  const closeCol = columnToLetter(tDS + 4);
  const volCol = columnToLetter(tDS + 5);
  
  // Cached fundamentals in DATA row 3
  const athCell = `DATA!${columnToLetter(tDS + 1)}3`;
  const peCell = `DATA!${columnToLetter(tDS + 3)}3`;
  const epsCell = `DATA!${columnToLetter(tDS + 5)}3`;
  
  // Rolling window anchors
  const lastRowCount = `COUNTA(DATA!${closeCol}$5:${closeCol})`;
  const lastAbsRow = `(4+${lastRowCount})`;
  
  // Build all formulas for this ticker (34 formulas: B-AI)
  return [
    buildSignalFormula(row, SEP, useLongTermSignal),
    buildFundamentalFormula(row, peCell, epsCell, SEP),
    buildDecisionFormula(row, SEP, useLongTermSignal),
    `=ROUND(IFERROR(GOOGLEFINANCE("${t}"${SEP}"price")${SEP}0)${SEP}2)`, // Price
    `=IFERROR(GOOGLEFINANCE("${t}"${SEP}"changepct")/100${SEP}0)`, // Change %
    buildRVOLFormula(row, volCol, lastRowCount, SEP), // Vol Trend
    `=IFERROR(${athCell}${SEP}0)`, // ATH
    `=IFERROR(($E${row}-$H${row})/MAX(0.01${SEP}$H${row})${SEP}0)`, // ATH Diff %
    buildRRFormula(row, SEP), // R:R Quality
    `=REPT("★"${SEP} ($E${row}>$M${row}) + ($E${row}>$N${row}) + ($E${row}>$O${row}))`, // Trend Score
    `=IF($E${row}>$O${row}${SEP}"BULL"${SEP}"BEAR")`, // Trend State
    buildSMAFormula(closeCol, lastRowCount, 20, SEP), // SMA 20
    buildSMAFormula(closeCol, lastRowCount, 50, SEP), // SMA 50
    buildSMAFormula(closeCol, lastRowCount, 200, SEP), // SMA 200
    `=LIVERSI(DATA!${closeCol}$5:${closeCol}${SEP}$E${row})`, // RSI
    `=LIVEMACD(DATA!${closeCol}$5:${closeCol}${SEP}$E${row})`, // MACD Hist
    buildDivergenceFormula(row, closeCol, lastAbsRow, SEP), // Divergence
    `=IFERROR(LIVEADX(DATA!${highCol}$5:${highCol}${SEP}DATA!${lowCol}$5:${lowCol}${SEP}DATA!${closeCol}$5:${closeCol}${SEP}$E${row})${SEP}0)`, // ADX
    `=LIVESTOCHK(DATA!${highCol}$5:${highCol}${SEP}DATA!${lowCol}$5:${lowCol}${SEP}DATA!${closeCol}$5:${closeCol}${SEP}$E${row})`, // Stoch %K
    buildSupportFormula(row, lowCol, lastRowCount, SEP), // Support
    buildResistanceFormula(row, highCol, lastRowCount, SEP), // Resistance
    `=ROUND(MAX($V${row}${SEP}$E${row}+(($E${row}-$U${row})*3))${SEP}2)`, // Target
    `=IFERROR(LIVEATR(DATA!${highCol}$5:${highCol}${SEP}DATA!${lowCol}$5:${lowCol}${SEP}DATA!${closeCol}$5:${closeCol}${SEP}$E${row})${SEP}0)`, // ATR
    buildBBPFormula(row, closeCol, lastRowCount, SEP), // Bollinger %B
    buildPositionSizeFormula(row, SEP), // POSITION SIZE
    buildTechNotesFormula(row, SEP), // TECH NOTES
    buildFundNotesFormula(row, SEP, useLongTermSignal), // FUND NOTES
    buildVolRegimeFormula(row, SEP), // VOL REGIME
    buildATHZoneFormula(row, SEP), // ATH ZONE
    buildBBPSignalFormula(row, SEP), // BBP SIGNAL
    buildPatternsFormula(row, SEP), // PATTERNS
    `=ROUND(MAX($U${row}${SEP}$E${row}-($X${row}*2))${SEP}2)`, // ATR STOP
    `=ROUND($E${row}+($X${row}*3)${SEP}2)`, // ATR TARGET
    `=IF($A${row}=""${SEP}""${SEP}$D${row})` // LAST STATE
  ];
}

// Helper formula builders
function buildSignalFormula(row, SEP, useLongTermSignal) {
  if (useLongTermSignal) {
    return `=IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}"LOADING"${SEP}IFS($E${row}<$U${row}${SEP}"STOP OUT"${SEP}$E${row}<$O${row}${SEP}"RISK OFF"${SEP}AND($I${row}>=-0.01${SEP}$G${row}>=1.5${SEP}$S${row}>=20${SEP}$E${row}>$O${row})${SEP}"ATH BREAKOUT"${SEP}AND($X${row}>IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*1.5${SEP}$G${row}>=2.0${SEP}$E${row}>$V${row})${SEP}"VOLATILITY BREAKOUT"${SEP}AND($Y${row}<=0.1${SEP}$P${row}<=25${SEP}$T${row}<=0.20${SEP}$E${row}>$O${row})${SEP}"EXTREME OVERSOLD BUY"${SEP}AND($E${row}>$O${row}${SEP}$N${row}>$O${row}${SEP}$P${row}<=30${SEP}$Q${row}>0${SEP}$S${row}>=20${SEP}$G${row}>=1.5)${SEP}"STRONG BUY"${SEP}AND($E${row}>$O${row}${SEP}$N${row}>$O${row}${SEP}$P${row}<=40${SEP}$Q${row}>0${SEP}$S${row}>=15)${SEP}"BUY"${SEP}AND($E${row}>$O${row}${SEP}$P${row}<=35${SEP}$E${row}>=$N${row}*0.95)${SEP}"ACCUMULATE"${SEP}$P${row}<=20${SEP}"OVERSOLD"${SEP}OR($P${row}>=80${SEP}$Y${row}>=0.9)${SEP}"OVERBOUGHT"${SEP}AND($E${row}>$O${row}${SEP}$P${row}>40${SEP}$P${row}<70)${SEP}"HOLD"${SEP}TRUE${SEP}"NEUTRAL"))`;
  } else {
    return `=IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}"LOADING"${SEP}IFS($E${row}<$U${row}${SEP}"STOP OUT"${SEP}$E${row}<$O${row}${SEP}"RISK OFF"${SEP}AND($X${row}>IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*1.5${SEP}$G${row}>=2.0${SEP}$E${row}>$V${row})${SEP}"VOLATILITY BREAKOUT"${SEP}AND($I${row}>=-0.01${SEP}$G${row}>=1.5${SEP}$S${row}>=20)${SEP}"ATH BREAKOUT"${SEP}AND($G${row}>=1.5${SEP}$E${row}>=$V${row}*0.995)${SEP}"BREAKOUT"${SEP}AND($E${row}>$O${row}${SEP}$Q${row}>0${SEP}$S${row}>=20)${SEP}"MOMENTUM"${SEP}AND($E${row}>$O${row}${SEP}$N${row}>$O${row}${SEP}$S${row}>=15)${SEP}"UPTREND"${SEP}AND($E${row}>$N${row}${SEP}$E${row}>$M${row})${SEP}"BULLISH"${SEP}AND(OR($T${row}<=0.20${SEP}$Y${row}<=0.2)${SEP}$E${row}>$U${row})${SEP}"OVERSOLD"${SEP}OR($P${row}>=80${SEP}$Y${row}>=0.9)${SEP}"OVERBOUGHT"${SEP}AND($X${row}<IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*0.7${SEP}$S${row}<15${SEP}ABS($Y${row}-0.5)<0.2)${SEP}"VOLATILITY SQUEEZE"${SEP}$S${row}<15${SEP}"RANGE"${SEP}TRUE${SEP}"NEUTRAL"))`;
  }
}

function buildFundamentalFormula(row, peCell, epsCell, SEP) {
  return `=IFERROR(LET(peRaw${SEP}${peCell}${SEP}epsRaw${SEP}${epsCell}${SEP}athDiffRaw${SEP}$I${row}${SEP}pe${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(peRaw)${SEP}"[^0-9\\.\\-]"${SEP}""))${SEP}"")${SEP}eps${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(epsRaw)${SEP}"[^0-9\\.\\-]"${SEP}""))${SEP}"")${SEP}athDiff${SEP}IFERROR(VALUE(REGEXREPLACE(TO_TEXT(athDiffRaw)${SEP}"[^0-9\\.\\-]"${SEP}""))/100${SEP}"")${SEP}IFS(OR(pe=""${SEP}eps="")${SEP}"FAIR"${SEP}eps<=0${SEP}"ZOMBIE"${SEP}AND(pe>=60${SEP}athDiff<>""${SEP}athDiff>=-0.08)${SEP}"PRICED FOR PERFECTION"${SEP}pe>=35${SEP}"EXPENSIVE"${SEP}AND(pe>0${SEP}pe<=25${SEP}eps>=0.5)${SEP}"VALUE"${SEP}AND(pe>25${SEP}pe<35${SEP}eps>=0.5)${SEP}"FAIR"${SEP}TRUE${SEP}"FAIR"))${SEP}"FAIR")`;
}

function buildDecisionFormula(row, SEP, useLongTermSignal) {
  const tagExpr = `UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}""))`;
  const purchasedExpr = `ISNUMBER(SEARCH("PURCHASED"${SEP}${tagExpr}))`;
  
  if (useLongTermSignal) {
    return `=IF($A${row}=""${SEP}""${SEP}IF($B${row}="LOADING"${SEP}"LOADING"${SEP}IF(${purchasedExpr}${SEP}IFS(OR($B${row}="STOP OUT"${SEP}$B${row}="RISK OFF")${SEP}"🔴 EXIT"${SEP}AND(OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}OR(ISNUMBER(SEARCH("VALUE"${SEP}UPPER($C${row})))${SEP}ISNUMBER(SEARCH("FAIR"${SEP}UPPER($C${row})))))${SEP}"🟢 ADD"${SEP}AND(OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}ISNUMBER(SEARCH("EXPENSIVE"${SEP}UPPER($C${row}))))${SEP}"🟡 HOLD / ADD SMALL"${SEP}AND(OR($B${row}="STRONG BUY"${SEP}$B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}ISNUMBER(SEARCH("PERFECTION"${SEP}UPPER($C${row}))))${SEP}"🟡 HOLD (NO ADD)"${SEP}AND($B${row}="OVERBOUGHT"${SEP}OR(ISNUMBER(SEARCH("EXPENSIVE"${SEP}UPPER($C${row})))${SEP}ISNUMBER(SEARCH("PERFECTION"${SEP}UPPER($C${row})))))${SEP}"🟠 TRIM"${SEP}$B${row}="HOLD"${SEP}"⚖️ HOLD"${SEP}TRUE${SEP}"⚖️ HOLD")${SEP}IFS(OR($B${row}="STOP OUT"${SEP}$B${row}="RISK OFF")${SEP}"🔴 AVOID"${SEP}$B${row}="STRONG BUY"${SEP}"🟢 STRONG BUY"${SEP}OR($B${row}="BUY"${SEP}$B${row}="ACCUMULATE")${SEP}"🟢 BUY"${SEP}$B${row}="OVERSOLD"${SEP}"🟡 WATCH (OVERSOLD)"${SEP}$B${row}="OVERBOUGHT"${SEP}"⏳ WAIT (OVERBOUGHT)"${SEP}$B${row}="HOLD"${SEP}"⚖️ WATCH"${SEP}TRUE${SEP}"⚪ NEUTRAL"))))`;
  } else {
    return `=IF($A${row}=""${SEP}""${SEP}LET(tag${SEP}UPPER(IFERROR(INDEX(INPUT!$C$3:$C${SEP}MATCH($A${row}${SEP}INPUT!$A$3:$A${SEP}0))${SEP}""))${SEP}purchased${SEP}REGEXMATCH(tag${SEP}"(^|,|\\\\s)PURCHASED(\\\\s|,|$)")${SEP}IFS(AND(IFERROR(VALUE($E${row})${SEP}0)>0${SEP}IFERROR(VALUE($U${row})${SEP}0)>0${SEP}IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($U${row})${SEP}0))${SEP}"Stop-Out"${SEP}AND(purchased${SEP}OR($B${row}="OVERBOUGHT"${SEP}IFERROR(VALUE($E${row})${SEP}0)>=IFERROR(VALUE($W${row})${SEP}0)))${SEP}"Take Profit"${SEP}AND(purchased${SEP}$B${row}="RISK OFF")${SEP}"Risk-Off"${SEP}AND(NOT(purchased)${SEP}$B${row}="RISK OFF")${SEP}"Avoid"${SEP}AND(NOT(purchased)${SEP}OR($B${row}="VOLATILITY BREAKOUT"${SEP}$B${row}="ATH BREAKOUT")${SEP}OR($C${row}="VALUE"${SEP}$C${row}="FAIR"))${SEP}"Strong Trade Long"${SEP}AND(NOT(purchased)${SEP}$B${row}="BREAKOUT"${SEP}OR($C${row}="VALUE"${SEP}$C${row}="FAIR"))${SEP}"Trade Long"${SEP}AND(NOT(purchased)${SEP}$B${row}="MOMENTUM"${SEP}$C${row}="VALUE")${SEP}"Accumulate"${SEP}AND(NOT(purchased)${SEP}$B${row}="OVERSOLD")${SEP}"Add in Dip"${SEP}AND(NOT(purchased)${SEP}$B${row}="VOLATILITY SQUEEZE")${SEP}"Wait for Breakout"${SEP}$B${row}="MOMENTUM"${SEP}"Hold"${SEP}$B${row}="UPTREND"${SEP}"Hold"${SEP}$B${row}="BULLISH"${SEP}"Hold"${SEP}TRUE${SEP}"Hold")))`;
  }
}

function buildRVOLFormula(row, volCol, lastRowCount, SEP) {
  return `=ROUND(IFERROR(OFFSET(DATA!${volCol}$5${SEP}${lastRowCount}-1${SEP}0)/AVERAGE(OFFSET(DATA!${volCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))${SEP}1)${SEP}2)`;
}

function buildRRFormula(row, SEP) {
  return `=IF(OR($E${row}<=$U${row}${SEP}$E${row}=0)${SEP}0${SEP}ROUND(MAX(0${SEP}$V${row}-$E${row})/MAX($X${row}*0.5${SEP}$E${row}-$U${row})${SEP}2))`;
}

function buildSMAFormula(closeCol, lastRowCount, period, SEP) {
  return `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!${closeCol}$5${SEP}${lastRowCount}-${period}${SEP}0${SEP}${period}))${SEP}0)${SEP}2)`;
}

function buildDivergenceFormula(row, closeCol, lastAbsRow, SEP) {
  return `=IFERROR(IFS(AND($E${row}<INDEX(DATA!${closeCol}:${closeCol}${SEP}${lastAbsRow}-14)${SEP}$P${row}>50)${SEP}"BULL DIV"${SEP}AND($E${row}>INDEX(DATA!${closeCol}:${closeCol}${SEP}${lastAbsRow}-14)${SEP}$P${row}<50)${SEP}"BEAR DIV"${SEP}TRUE${SEP}"—")${SEP}"—")`;
}

function buildSupportFormula(row, lowCol, lastRowCount, SEP) {
  return `=ROUND(IFERROR(LET(win${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!${lowCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!${lowCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MIN(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.15))${SEP}out)${SEP}0)${SEP}2)`;
}

function buildResistanceFormula(row, highCol, lastRowCount, SEP) {
  return `=ROUND(IFERROR(LET(win${SEP}IFS($S${row}<20${SEP}10${SEP}$S${row}<35${SEP}22${SEP}TRUE${SEP}40)${SEP}n${SEP}${lastRowCount}${SEP}start${SEP}MAX(0${SEP}n-win)${SEP}len${SEP}MIN(win${SEP}n)${SEP}rng${SEP}IF(len<=0${SEP}OFFSET(DATA!${highCol}$5${SEP}0${SEP}0)${SEP}OFFSET(DATA!${highCol}$5${SEP}start${SEP}0${SEP}len))${SEP}out${SEP}IF(COUNTA(rng)<3${SEP}IFERROR(MAX(rng)${SEP}0)${SEP}PERCENTILE.INC(rng${SEP}0.85))${SEP}out)${SEP}0)${SEP}2)`;
}

function buildBBPFormula(row, closeCol, lastRowCount, SEP) {
  return `=ROUND(IFERROR((($E${row}-$M${row})/(4*STDEV(OFFSET(DATA!${closeCol}$5${SEP}${lastRowCount}-20${SEP}0${SEP}20))))+0.5${SEP}0.5)${SEP}2)`;
}

function buildPositionSizeFormula(row, SEP) {
  return `=IF($A${row}=""${SEP}""${SEP}LET(riskReward${SEP}$J${row}${SEP}atrRisk${SEP}$X${row}/$E${row}${SEP}athRisk${SEP}IF($I${row}>=-0.05${SEP}0.8${SEP}1.0)${SEP}volRegimeRisk${SEP}IFS(atrRisk<=0.02${SEP}1.2${SEP}atrRisk<=0.05${SEP}1.0${SEP}atrRisk<=0.08${SEP}0.7${SEP}TRUE${SEP}0.5)${SEP}baseSize${SEP}0.02${SEP}rrMultiplier${SEP}IF(riskReward>=3${SEP}1.5${SEP}IF(riskReward>=2${SEP}1.0${SEP}0.5))${SEP}finalSize${SEP}MIN(0.08${SEP}baseSize*rrMultiplier*volRegimeRisk*athRisk)${SEP}TEXT(finalSize${SEP}"0.0%")&" (Vol: "&IFS(atrRisk<=0.02${SEP}"LOW"${SEP}atrRisk<=0.05${SEP}"NORM"${SEP}atrRisk<=0.08${SEP}"HIGH"${SEP}TRUE${SEP}"EXTR")&")"))`;
}

function buildTechNotesFormula(row, SEP) {
  return `=IF($A${row}=""${SEP}""${SEP}"VOL: RVOL "&TEXT(IFERROR(VALUE($G${row})${SEP}0)${SEP}"0.00")&"x; "&IF(IFERROR(VALUE($G${row})${SEP}0)<1${SEP}"sub-average (weak sponsorship)."${SEP}"healthy participation.")&CHAR(10)&"REGIME: Price "&TEXT(IFERROR(VALUE($E${row})${SEP}0)${SEP}"0.00")&" vs SMA200 "&TEXT(IFERROR(VALUE($O${row})${SEP}0)${SEP}"0.00")&"; "&IF(IFERROR(VALUE($E${row})${SEP}0)<IFERROR(VALUE($O${row})${SEP}0)${SEP}"risk-off below SMA200."${SEP}"risk-on above SMA200.")&CHAR(10)&"VOL/STRETCH: ATR(14) "&TEXT(IFERROR(VALUE($X${row})${SEP}0)${SEP}"0.00")&"; stretch "&IF(OR(IFERROR(VALUE($X${row})${SEP}0)=0${SEP}IFERROR(VALUE($M${row})${SEP}0)=0)${SEP}"—"${SEP}TEXT((IFERROR(VALUE($E${row})${SEP}0)-IFERROR(VALUE($M${row})${SEP}0))/IFERROR(VALUE($X${row})${SEP}1)${SEP}"0.0")&"x ATR")&" (<= +/-2x)."&CHAR(10)&"MOMENTUM: RSI(14) "&TEXT(IFERROR(VALUE($P${row})${SEP}0)${SEP}"0.0")&"; "&IF(IFERROR(VALUE($P${row})${SEP}0)<40${SEP}"negative bias."${SEP}"constructive.")&" MACD hist "&TEXT(IFERROR(VALUE($Q${row})${SEP}0)${SEP}"0.000")&"; "&IF(IFERROR(VALUE($Q${row})${SEP}0)>0${SEP}"improving."${SEP}"weak.")&CHAR(10)&"TREND: ADX(14) "&TEXT(IFERROR(VALUE($S${row})${SEP}0)${SEP}"0.0")&"; "&IF(IFERROR(VALUE($S${row})${SEP}0)>=25${SEP}"strong."${SEP}"weak.")&" Stoch %K "&TEXT(IFERROR(VALUE($T${row})${SEP}0)${SEP}"0.0%")&" — "&IF(IFERROR(VALUE($T${row})${SEP}0)<=0.2${SEP}"oversold zone (mean-reversion potential)."${SEP}IF(IFERROR(VALUE($T${row})${SEP}0)>=0.8${SEP}"overbought zone (pullback risk)."${SEP}"neutral range (no timing edge)."))&CHAR(10)&"R:R: "&TEXT(IFERROR(VALUE($J${row})${SEP}0)${SEP}"0.00")&"x; "&IF(IFERROR(VALUE($J${row})${SEP}0)>=3${SEP}"favorable."${SEP}"limited"))`;
}

function buildFundNotesFormula(row, SEP, useLongTermSignal) {
  // Long-term investment mode: detailed reasoning for SIGNAL, FUNDAMENTAL, and DECISION
  const fFundNotesLong = `=IF($A${row}=""${SEP}""${SEP}` +
    `IF(OR(ISBLANK($E${row})${SEP}$E${row}=0)${SEP}` +
    `"LOADING"${SEP}` +
    `TEXTJOIN(CHAR(10)${SEP}TRUE${SEP}` +
    // 1) SIGNAL PATH (why SIGNAL fired)
    `"SIGNAL PATH: "&` +
    `IFS(` +
    `$B${row}="REDUCE / EXIT"${SEP}` +
    `IF($E${row}<$U${row}${SEP}` +
    `"STRUCTURAL BREAK: Price breached key support → de-risk / exit posture."${SEP}` +
    `"REGIME FLIP: Price below SMA200 → long-term trend broken (risk-off)."` +
    `)${SEP}` +
    `$B${row}="ACCUMULATE"${SEP}` +
    `"ACCUMULATION SETUP: Above SMA200 with bullish momentum; pullback into institutional buy-zone (near SMA50/SMA20) with trend strength (ADX) confirmation."${SEP}` +
    `$B${row}="HOLD"${SEP}` +
    `"EXTENSION / HOLD: Uptrend intact and momentum positive, but price is extended vs SMA20 → avoid chasing; hold core exposure."${SEP}` +
    `$B${row}="WAIT (PULLBACK)"${SEP}` +
    `"WAIT STATE: Trend constructive, but entry is not in the accumulation zone → wait for pullback toward SMA50/SMA20."${SEP}` +
    `TRUE${SEP}` +
    `"NO-TREND: Structure/momentum not in a bullish regime → watchlist only."` +
    `)${SEP}` +
    // 2) TECH CONTEXT (inputs that drove it)
    `"TECH CONTEXT: "` +
    `&"Px "&TEXT($E${row}${SEP}"0.00")` +
    `&" | SMA20 "&TEXT($M${row}${SEP}"0.00")` +
    `&" / SMA50 "&TEXT($N${row}${SEP}"0.00")` +
    `&" / SMA200 "&TEXT($O${row}${SEP}"0.00")` +
    `&" | RSI "&TEXT($P${row}${SEP}"0.0")` +
    `&" | MACD(H) "&TEXT($Q${row}${SEP}"0.000")` +
    `&" | ADX "&TEXT($S${row}${SEP}"0.0")` +
    `&IF($U${row}>0${SEP}" | Support "&TEXT($U${row}${SEP}"0.00")${SEP}"")${SEP}` +
    // 3) VALUATION REGIME (why FUNDAMENTAL matters)
    `"VALUATION REGIME: "&$C${row}&" — "&` +
    `IFS(` +
    `ISNUMBER(SEARCH("PERFECTION"${SEP}UPPER(TRIM($C${row}))))${SEP}` +
    `"priced-for-perfection; require margin-of-safety (smaller adds / only on pullbacks)."${SEP}` +
    `ISNUMBER(SEARCH("EXPENSIVE"${SEP}UPPER(TRIM($C${row}))))${SEP}` +
    `"valuation headwind; stage entries and avoid oversized adds at extensions."${SEP}` +
    `ISNUMBER(SEARCH("FAIR"${SEP}UPPER(TRIM($C${row}))))${SEP}` +
    `"neutral valuation; sizing can be driven mainly by technical regime and risk limits."${SEP}` +
    `OR(` +
    `ISNUMBER(SEARCH("CHEAP"${SEP}UPPER(TRIM($C${row}))))${SEP}` +
    `ISNUMBER(SEARCH("UNDER"${SEP}UPPER(TRIM($C${row}))))` +
    `)${SEP}` +
    `"valuation support; improves long-term reward-to-risk when trend is intact."${SEP}` +
    `TRUE${SEP}` +
    `"valuation unclear; treat as neutral and defer to structure + risk controls."` +
    `)${SEP}` +
    // 4) DECISION PATH (why DECISION fired)
    `"DECISION PATH: "&$D${row}&" — "&` +
    `IFS(` +
    `ISNUMBER(SEARCH("EXIT"${SEP}$D${row}))${SEP}` +
    `"capital preservation; structural rules override valuation."${SEP}` +
    `ISNUMBER(SEARCH("BUY / ADD"${SEP}$D${row}))${SEP}` +
    `"trend confirmed + valuation supportive; accumulate in tranches."${SEP}` +
    `ISNUMBER(SEARCH("ADD (SMALL)"${SEP}$D${row}))${SEP}` +
    `"trend ok but valuation elevated; add marginally and only on deeper pullbacks."${SEP}` +
    `ISNUMBER(SEARCH("HOLD / ADD CAUTIOUS"${SEP}$D${row}))${SEP}` +
    `"high valuation regime; hold core; add only on high-quality pullbacks."${SEP}` +
    `OR(` +
    `ISNUMBER(SEARCH("HOLD / REDUCE"${SEP}$D${row}))${SEP}` +
    `ISNUMBER(SEARCH("REDUCE"${SEP}$D${row}))` +
    `)${SEP}` +
    `"valuation risk management; consider trimming into strength."${SEP}` +
    `ISNUMBER(SEARCH("WAIT"${SEP}$D${row}))${SEP}` +
    `"no edge at current level; wait for mean reversion toward structure."${SEP}` +
    `TRUE${SEP}` +
    `"monitor; re-evaluate when regime inputs change."` +
    `)` +
    `)` + // TEXTJOIN
    `)` +
    `)`;

  // Trade mode: fundamental analysis with technical signal and decision reasoning
  const fFundNotesTrade = `=IF($A${row}=""${SEP}""${SEP}"FUNDAMENTAL ANALYSIS: "&IFS($C${row}="VALUE"${SEP}"This stock is attractively priced with strong earnings and reasonable valuation (PE ≤ 25). The fundamentals provide a supportive tailwind for any position."${SEP}$C${row}="FAIR"${SEP}"This stock has decent fundamentals but nothing exceptional. The valuation is neither cheap nor expensive, so fundamentals are neutral to the trade."${SEP}$C${row}="EXPENSIVE"${SEP}"This stock is trading at a premium valuation (PE 35-59). While not prohibitive, there's less margin for error and fundamentals create a headwind."${SEP}$C${row}="PRICED FOR PERFECTION"${SEP}"This stock has extremely high expectations built into the price (PE ≥ 60). Any disappointment could cause significant downside. Fundamentals are fragile."${SEP}$C${row}="ZOMBIE"${SEP}"This company is losing money or has very weak earnings quality (EPS ≤ 0). High risk of permanent capital loss. Fundamentals are severely negative."${SEP}TRUE${SEP}"Fundamental analysis is inconclusive due to missing data.")&CHAR(10)&CHAR(10)&"TECHNICAL SIGNAL: "&$B${row}&CHAR(10)&"Why this signal: "&IFS($B${row}="Stop-Out"${SEP}"Price has broken below the key support level, invalidating the bullish thesis. This is a defensive exit signal to preserve capital."${SEP}$B${row}="Breakout (High Volume)"${SEP}"Price is breaking above resistance with strong volume confirmation, suggesting institutional participation and potential for continued upside momentum."${SEP}$B${row}="Trend Continuation"${SEP}"Price is above the 200-day moving average with positive momentum indicators, suggesting the existing uptrend has room to continue higher."${SEP}$B${row}="Mean Reversion (Oversold)"${SEP}"Price is oversold on short-term indicators but holding above key support, creating a potential bounce opportunity back toward fair value."${SEP}$B${row}="Volatility Squeeze (Coiling)"${SEP}"Price volatility has compressed to extremely low levels, often preceding significant directional moves. Waiting for the breakout direction."${SEP}$B${row}="Range-Bound (Low ADX)"${SEP}"Trend strength is weak with price moving sideways. This environment favors range trading rather than directional bets."${SEP}TRUE${SEP}"Market conditions don't clearly favor any specific technical setup. Monitoring for clearer signals.")&CHAR(10)&CHAR(10)&"INVESTMENT DECISION: "&$D${row}&CHAR(10)&"Why this decision: "&IFS($D${row}="Stop-Out"${SEP}"Price has broken below support level. Exiting to prevent further losses and preserve capital."${SEP}AND($D${row}="Take Profit",IFERROR(VALUE(INDEX(CALCULATIONS!E:E,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),0)>=IFERROR(VALUE(INDEX(CALCULATIONS!W:W,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),0),IFERROR(VALUE(INDEX(CALCULATIONS!W:W,MATCH(UPPER(TRIM($A$1)),ARRAYFORMULA(UPPER(TRIM(CALCULATIONS!A:A))),0))),0)>0)${SEP}"Price has reached target level. Taking profits while conditions are favorable."${SEP}$D${row}="Take Profit"${SEP}"Price is overbought near resistance. Taking profits to avoid pullback risk from elevated RSI and Stochastic levels."${SEP}$D${row}="Reduce (Momentum Weak)"${SEP}"MACD histogram has turned negative and price is below SMA50. Reducing position size to manage deteriorating momentum."${SEP}$D${row}="Reduce (Overextended)"${SEP}"Price has extended too far above SMA20 relative to average volatility. Taking partial profits to reduce pullback risk."${SEP}$D${row}="Risk-Off (Below SMA200)"${SEP}"Price is below the 200-day moving average indicating risk-off conditions. Maintaining defensive posture until trend improves."${SEP}$D${row}="Avoid"${SEP}"Price is below SMA200 indicating risk-off conditions. Avoiding new positions until trend improves above key moving average."${SEP}$D${row}="Add in Dip"${SEP}"Price is above support with Stochastic showing oversold conditions. Adding to position on this dip opportunity."${SEP}$D${row}="Trade Long"${SEP}"Breakout signal confirmed with strong fundamentals. Initiating long position with favorable risk/reward setup."${SEP}$D${row}="Accumulate"${SEP}"Trend continuation signal with VALUE fundamentals. Adding to existing position as uptrend remains intact above SMA200."${SEP}$D${row}="Hold"${SEP}"Current market conditions suggest maintaining existing position. Monitoring for clearer directional signals before making changes."${SEP}TRUE${SEP}"Decision framework suggests maintaining current stance until market conditions become clearer.")&IF(AND(OR($B${row}="Breakout (High Volume)"${SEP}$B${row}="Trend Continuation")${SEP}OR($C${row}="ZOMBIE"${SEP}$C${row}="PRICED FOR PERFECTION"${SEP}$C${row}="EXPENSIVE"))${SEP}CHAR(10)&CHAR(10)&"⚠️ RISK WARNING: Strong technical momentum is conflicting with weak or fragile fundamentals. This creates higher risk of sharp reversals if momentum fails."${SEP}IF(AND(OR($B${row}="Mean Reversion (Oversold)"${SEP}$B${row}="Stop-Out")${SEP}$C${row}="VALUE")${SEP}CHAR(10)&CHAR(10)&"💡 OPPORTUNITY NOTE: Attractive valuation is present, but technical structure needs to improve before becoming more aggressive."${SEP}"")))`;

  return useLongTermSignal ? fFundNotesLong : fFundNotesTrade;
}

function buildVolRegimeFormula(row, SEP) {
  return `=IFS($X${row}/$E${row}<=0.02${SEP}"LOW VOL"${SEP}$X${row}/$E${row}<=0.05${SEP}"NORMAL VOL"${SEP}$X${row}/$E${row}<=0.08${SEP}"HIGH VOL"${SEP}TRUE${SEP}"EXTREME VOL")`;
}

function buildATHZoneFormula(row, SEP) {
  return `=IFS($I${row}>=-0.02${SEP}"AT ATH"${SEP}$I${row}>=-0.05${SEP}"NEAR ATH"${SEP}$I${row}>=-0.15${SEP}"RESISTANCE ZONE"${SEP}$I${row}>=-0.30${SEP}"PULLBACK ZONE"${SEP}$I${row}>=-0.50${SEP}"CORRECTION ZONE"${SEP}TRUE${SEP}"DEEP VALUE ZONE")`;
}

function buildBBPSignalFormula(row, SEP) {
  return `=IFS(AND($Y${row}>=0.9${SEP}$P${row}>=70)${SEP}"EXTREME OVERBOUGHT"${SEP}AND($Y${row}<=0.1${SEP}$P${row}<=30)${SEP}"EXTREME OVERSOLD"${SEP}AND($Y${row}>=0.8${SEP}$E${row}>$O${row})${SEP}"MOMENTUM STRONG"${SEP}AND($Y${row}<=0.2${SEP}$E${row}>$U${row})${SEP}"MEAN REVERSION"${SEP}TRUE${SEP}"NEUTRAL")`;
}

function buildPatternsFormula(row, SEP) {
  return `=TEXTJOIN(" | "${SEP}TRUE${SEP}IF(AND($X${row}>IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*1.5${SEP}$G${row}>=2.0${SEP}$E${row}>$V${row})${SEP}"VOL BREAKOUT"${SEP}"")${SEP}IF(AND($I${row}>=-0.01${SEP}$G${row}>=1.5${SEP}$S${row}>=20)${SEP}"ATH BREAKOUT"${SEP}"")${SEP}IF(AND($Y${row}<=0.15${SEP}$P${row}<=25${SEP}$T${row}<=0.20${SEP}$E${row}>$O${row})${SEP}"MEAN REVERSION SETUP"${SEP}"")${SEP}IF(AND($X${row}<IFERROR(AVERAGE(OFFSET($X${row}${SEP}-MIN(20${SEP}ROW($X${row})-1)${SEP}0${SEP}MIN(20${SEP}ROW($X${row})-1)))${SEP}$X${row})*0.7${SEP}$S${row}<15${SEP}ABS($Y${row}-0.5)<0.2)${SEP}"VOLATILITY SQUEEZE"${SEP}""))`;
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