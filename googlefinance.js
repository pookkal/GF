/**
 * ==============================================================================
 * BASELINE LABEL: STABLE_MASTER_V68_CHART_FIX
 * DATE: 23 DEC 2025
 * FIX: Chart Date calculated in script to eliminate formula lag. (Instant Updates)
 * ==============================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Institutional Terminal')
    .addItem('🚀 1-CLICK REBUILD ALL', 'FlushAllSheetsAndBuild')
    .addSeparator()
    .addItem('1. Fetch Data Only', 'generateDataSheet')
    .addItem('2. Build Calculations', 'generateCalculationsSheet')
    .addItem('3. Refresh Dashboard Only', 'generateDashboardSheet')
    .addItem('4. Setup Chart Only', 'setupChartSheet')
    .addSeparator()
    .addItem('🔔 Start Market Monitor', 'startMarketMonitor')
    .addItem('🔕 Stop Monitor', 'stopMarketMonitor')
    .addItem('📩 Test Alert Now', 'checkSignalsAndSendAlerts')
    .addToUi();
}

/**
 * ------------------------------------------------------------------
 * 5. ALERT ENGINE
 * ------------------------------------------------------------------
 */

function startMarketMonitor() {
  stopMarketMonitor(); 
  ScriptApp.newTrigger('checkSignalsAndSendAlerts')
    .timeBased()
    .everyMinutes(30)
    .create();
  SpreadsheetApp.getUi().alert('✅ Monitor Started. Alerts every 30 mins.');
}

function stopMarketMonitor() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkSignalsAndSendAlerts') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  SpreadsheetApp.getUi().alert('🔕 Monitor Stopped.');
}

function checkSignalsAndSendAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName("CALCULATIONS");
  if (!calcSheet) return;

  generateCalculationsSheet(); 
  SpreadsheetApp.flush();

  const data = calcSheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  const props = PropertiesService.getScriptProperties();
  let history = JSON.parse(props.getProperty('ALERT_HISTORY') || '{}');

  if (history.date !== today) {
    history = { date: today, sent: [] };
  }

  let newAlerts = [];

  for (let i = 2; i < data.length; i++) {
    const ticker = data[i][0];
    const decision = data[i][4]; // Column E
    const reason = data[i][21];  // Column V
    const price = data[i][1];

    if (decision.includes("EXECUTE") && !history.sent.includes(ticker)) {
      newAlerts.push({ ticker, price, reason });
      history.sent.push(ticker);
    }
  }

  if (newAlerts.length > 0) {
    const emailRecipient = Session.getActiveUser().getEmail();
    const subject = `🚀 MARKET ALERT: ${newAlerts.length} New Opportunities`;
    
    let body = "Institutional Terminal - New Buy Signals:\n\n";
    newAlerts.forEach(a => {
      body += `--------------------------------------------------\n`;
      body += `Ticker:   ${a.ticker}  @ $${a.price}\n`;
      body += `Signal:   ★ EXECUTE\n`;
      body += `Analyst:  ${a.reason}\n`;
    });
    body += `\n--------------------------------------------------\n`;
    body += `Check Dashboard: ${ss.getUrl()}`;

    MailApp.sendEmail(emailRecipient, subject, body);
    props.setProperty('ALERT_HISTORY', JSON.stringify(history));
  }
}

/**
 * ------------------------------------------------------------------
 * CORE FUNCTIONS
 * ------------------------------------------------------------------
 */

function FlushAllSheetsAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["DATA", "CALCULATIONS", "CHART", "DASHBOARD"];
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('🚨 Full Rebuild', 'Refresh All Systems?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Syncing Global Data..."), "Status");
  generateDataSheet();
  SpreadsheetApp.flush();
  Utilities.sleep(3000); 

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>2/4:</b> Calculating Signals...", "Status"));
  generateCalculationsSheet();
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>3/4:</b> Updating Dashboard..."), "Status");
  generateDashboardSheet(); 
  SpreadsheetApp.flush();

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>4/4:</b> Finalizing Report..."), "Status");
  setupChartSheet();
  ui.alert('✅ Rebuild Complete', 'Systems Online.', ui.ButtonSet.OK);
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() === "INPUT" && (e.range.getA1Notation() === "B1" || e.range.getA1Notation() === "C1")) {
    generateDashboardSheet();
  }
  if (sheet.getName() !== "CHART") return;
  const watchList = ["B1", "D2", "A3", "B3", "C3"];
  if (watchList.includes(e.range.getA1Notation())) {
    updateDynamicChart();
  }
}

function generateDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;
  let dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD", inputSheet.getIndex() + 1);
  dashboard.clear().clearFormats();
  const headers = [["Ticker", "Price", "Change %", "SIGNAL", "DECISION", "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "REASONING"]];
  dashboard.getRange(2, 1, 1, 22).setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold");

  const formula = '=IFERROR(SORT(FILTER(CALCULATIONS!$A$3:$V, ISNUMBER(MATCH(CALCULATIONS!$A$3:$A, FILTER(INPUT!$A$3:$A, ' +
    '(IF(OR(INPUT!$B$1="", INPUT!$B$1="ALL"), 1, REGEXMATCH(INPUT!$B$3:$B, "(?i)\\b(" & SUBSTITUTE(SUBSTITUTE(INPUT!$B$1, ", ", "|"), ",", "|") & ")\\b"))) * ' +
    '(IF(OR(INPUT!$C$1="", INPUT!$C$1="ALL"), 1, REGEXMATCH(INPUT!$C$3:$C, "(?i)\\b(" & SUBSTITUTE(SUBSTITUTE(INPUT!$C$1, ", ", "|"), ",", "|") & ")\\b")))' +
    '), 0))), 3, FALSE), "No Matches Found")';
  
  dashboard.getRange("A3").setFormula(formula);
  dashboard.getRange("A:V").setHorizontalAlignment("left");
  
  const range = dashboard.getRange("E3:E");
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("EXECUTE")
    .setBackground("#00C853")
    .setFontColor("white")
    .setBold(true)
    .setRanges([range])
    .build();
  dashboard.setConditionalFormatRules([rule]);

  SpreadsheetApp.flush();
  dashboard.getRangeList(['C3:C', 'G3:G', 'U3:U']).setNumberFormat("0.00%"); 
  dashboard.setFrozenRows(2);
}

function setupChartSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  const tickers = getCleanTickers(inputSheet);
  let chartSheet = ss.getSheetByName("CHART") || ss.insertSheet("CHART");
  
  chartSheet.clear().clearFormats();
  forceExpandSheet(chartSheet, 45);
  
  chartSheet.setColumnWidth(1, 180); 
  chartSheet.setColumnWidth(2, 120); 
  chartSheet.setColumnWidth(5, 500); 

  chartSheet.getRange("A1").setValue("TICKER:").setFontWeight("bold");
  chartSheet.getRange("B1").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tickers).build()).setValue(tickers[0]);
  chartSheet.getRange("D1").setValue("VIEW:").setFontWeight("bold");
  chartSheet.getRange("D2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"]).build()).setValue("DAILY");
  
  chartSheet.getRange("E1").setFormula('=IFERROR(VLOOKUP(B1, CALCULATIONS!$A$3:$V, 22, 0), "—")').setWrap(true).setVerticalAlignment("top");

  chartSheet.getRange("A2:C2").setValues([["YEAR", "MONTH", "DAY"]]).setFontWeight("bold").setHorizontalAlignment("center");
  
  const numRule = (max) => SpreadsheetApp.newDataValidation().requireValueInList(Array.from({length: max + 1}, (_, i) => i)).build();
  chartSheet.getRange("A3").setDataValidation(numRule(5)).setValue(0);
  chartSheet.getRange("B3").setDataValidation(numRule(12)).setValue(3);
  chartSheet.getRange("C3").setDataValidation(numRule(31)).setValue(0);

  chartSheet.getRange("A4").setValue("DATE").setFontWeight("bold");
  chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)").setNumberFormat("yyyy-mm-dd");

  chartSheet.getRange("A1:F4").setBackground("#000000").setFontColor("#FFFF00");
  chartSheet.getRangeList(["A3", "B3", "C3", "D2"]).setBackground("#FF80AB").setFontColor("#000000").setHorizontalAlignment("center").setFontWeight("bold");
  chartSheet.getRange("B1").setBackground("#FFEB3B").setFontColor("black").setFontWeight("bold");

  const t = "B1";
  
  const data = [
    ["SIGNAL (RAW)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 4, 0), "Wait")`], 
    ["DECISION (FINAL)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 5, 0), "-")`], 
    ["LIVE PRICE", `=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`], 
    ["CHANGE ($)", `=IFERROR(B7 - GOOGLEFINANCE(${t}, "closeyest"), 0)`], 
    ["CHANGE (%)", `=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`], 
    ["", ""], 
    ["[ VALUATION METRICS ]", ""], 
    ["ATH (TRUE)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 6, 0), 0)`], 
    ["DIFF FROM ATH %", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 7, 0), 0)`], 
    ["P/E RATIO", `=IFERROR(GOOGLEFINANCE(${t}, "pe"), 0)`],
    ["EPS", `=IFERROR(GOOGLEFINANCE(${t}, "eps"), 0)`],
    ["52W HIGH", `=IFERROR(GOOGLEFINANCE(${t}, "high52"), 0)`],
    ["52W LOW", `=IFERROR(GOOGLEFINANCE(${t}, "low52"), 0)`],
    ["", ""], 
    ["[ MOMENTUM & TREND ]", ""], 
    ["SMA 20", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 11, 0), 0)`], 
    ["SMA 50", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 12, 0), 0)`], 
    ["SMA 200", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 13, 0), 0)`], 
    ["RSI (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 15, 0), 50)`], 
    ["TREND STATE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 10, 0), "—")`], 
    ["DIVERGENCE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 16, 0), "Neutral")`], 
    ["RELATIVE VOLUME", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 14, 0), 1)`], 
    ["", ""], 
    ["[ TECHNICAL LEVELS ]", ""], 
    ["SUPPORT FLOOR", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 17, 0), 0)`], 
    ["RESISTANCE CEILING", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 19, 0), 0)`], 
    ["TARGET (3:1 R:R)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 18, 0), 0)`], 
    ["ATR (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 20, 0), 0)`], 
    ["BOLLINGER %B", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$V, 21, 0), 0)`] 
  ];

  chartSheet.getRange(5, 1, data.length, 1).setValues(data.map(r => [r[0]])).setFontWeight("bold");
  chartSheet.getRange(5, 2, data.length, 1).setFormulas(data.map(r => [r[1]]));
  
  [11, 19, 28].forEach(r => chartSheet.getRange(r, 1, 1, 2).setBackground("#444").setFontColor("white").setHorizontalAlignment("center"));
  
  SpreadsheetApp.flush();
  chartSheet.getRange("B5:B33").setHorizontalAlignment("left");
  
  chartSheet.getRangeList(["B7", "B8", "B12", "B14:B17", "B20:B23", "B26", "B29:B32"]).setNumberFormat("#,##0.00");
  chartSheet.getRangeList(["B9", "B13", "B33"]).setNumberFormat("0.00%");

  SpreadsheetApp.flush();
  updateDynamicChart();
}

/**
 * 3. CHART ENGINE (DIRECT CALCULATION FIX)
 */
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;

  const ticker = sheet.getRange("B1").getValue();
  const isWeekly = sheet.getRange("D2").getValue() === "WEEKLY";

  // FIX: CALCULATE DATE IN JS TO AVOID SHEET FORMULA LAG
  const years = sheet.getRange("A3").getValue() || 0;
  const months = sheet.getRange("B3").getValue() || 0;
  const days = sheet.getRange("C3").getValue() || 0;
  
  const now = new Date();
  // Subtract time from current date
  const startDate = new Date(now.getFullYear() - years, now.getMonth() - months, now.getDate() - days);

  SpreadsheetApp.flush();
  
  const supportVal = Number(sheet.getRange("B29").getValue()) || 0; 
  const resistanceVal = Number(sheet.getRange("B30").getValue()) || 0; 

  const rawHeaders = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIdx = rawHeaders.indexOf(ticker);
  if (colIdx === -1) return;

  const rawData = dataSheet.getRange(1, colIdx + 1, dataSheet.getLastRow(), 6).getValues();
  let masterData = [], viewVols = [], prices = [];

  for (let i = 4; i < rawData.length; i++) {
    let row = rawData[i], d = row[0], close = Number(row[4]), vol = Number(row[5]);
    if (!d || !(d instanceof Date) || isNaN(close) || close < 0.01) continue;
    if (d < startDate) continue; // Using JS Calculated Date
    if (isWeekly && d.getDay() !== 5) continue;

    let slice = rawData.slice(Math.max(4, i-200), i+1).map(r => r[4]);
    let s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a,b)=>a+b,0)/20).toFixed(2)) : null;
    let s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a,b)=>a+b,0)/50).toFixed(2)) : null;
    let s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a,b)=>a+b,0)/200).toFixed(2)) : null;

    masterData.push([Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "MMM dd"), close, (close >= (i>4?rawData[i-1][4]:close))?vol:null, (close < (i>4?rawData[i-1][4]:close))?vol:null, s20, s50, s200, resistanceVal, supportVal]);
    viewVols.push(vol); prices.push(close);
    if(s20) prices.push(s20); if(s50) prices.push(s50); if(s200) prices.push(s200);
  }

  if (masterData.length === 0) return;
  if (supportVal > 0) prices.push(supportVal); if (resistanceVal > 0) prices.push(resistanceVal);

  const minP = Math.min(...prices.filter(p => p > 0)) * 0.98;
  const maxP = Math.max(...prices.filter(p => p > 0)) * 1.02;
  const maxVol = Math.max(...viewVols);

  const chartLabels = [["Date", "Price", "Bull Vol", "Bear Vol", "SMA 20", "SMA 50", "SMA 200", "Resistance", "Support"]];
  sheet.getRange(2, 26, 1, 9).setValues(chartLabels).setFontWeight("bold").setFontColor("white");
  sheet.getRange(3, 26, 1500, 9).clearContent();
  sheet.getRange(3, 26, masterData.length, 9).setValues(masterData);
  SpreadsheetApp.flush();

  sheet.getCharts().forEach(c => sheet.removeChart(c));
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(2, 26, masterData.length + 1, 9))
    .setOption('useFirstRowAsHeaders', true)
    .setOption('series', {0:{type:'line',color:'#1A73E8',lineWidth:3,labelInLegend:'Price'},1:{type:'bars',color:'#2E7D32',targetAxisIndex:1,labelInLegend:'Bull Vol'},2:{type:'bars',color:'#C62828',targetAxisIndex:1,labelInLegend:'Bear Vol'},3:{type:'line',color:'#FBBC04',lineWidth:1.5,labelInLegend:'SMA 20'},4:{type:'line',color:'#9C27B0',lineWidth:1.5,labelInLegend:'SMA 50'},5:{type:'line',color:'#FF9800',lineWidth:2,labelInLegend:'SMA 200'},6:{type:'line',color:'#B71C1C',lineDashStyle:[4,4],labelInLegend:'Resistance'},7:{type:'line',color:'#0D47A1',lineDashStyle:[4,4],labelInLegend:'Support'}})
    .setOption('vAxes', {0:{viewWindow:{min:minP,max:maxP}},1:{viewWindow:{min:0,max:maxVol * 8},textStyle:{color:'none'}}})
    .setOption('legend', {position: 'top', textStyle: {fontSize: 10}})
    .setPosition(4, 3, 10, 10).setOption('width', 1150).setOption('height', 650).build();
  sheet.insertChart(chart);
}

function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet) return;
  const tickers = getCleanTickers(inputSheet);
  let calcSheet = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");
  calcSheet.clear().clearFormats();

  calcSheet.getRange("A1:E1").merge().setValue("[ CORE IDENT ]").setBackground("#263238").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("F1:H1").merge().setValue("[ PERFORMANCE ]").setBackground("#0D47A1").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("I1:P1").merge().setValue("[ MOMENTUM ]").setBackground("#1B5E20").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("Q1:U1").merge().setValue("[ RISK LEVELS ]").setBackground("#B71C1C").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
  calcSheet.getRange("V1").setValue("[ ANALYST ]").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

  const headers = [["Ticker", "Price", "Change %", "SIGNAL", "DECISION", "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "REASONING"]];
  calcSheet.getRange(2, 1, 1, 22).setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold");
  
  const formulas = [];
  tickers.forEach((ticker, i) => {
    const rowNum = i + 3, tickerDataStart = (i * 7) + 1, closeCol = columnToLetter(tickerDataStart + 4), lastRow = `COUNTA(DATA!$${closeCol}:$${closeCol})`;
    formulas.push([
      `=ROUND(IFERROR(GOOGLEFINANCE("${ticker}", "price")), 2)`,
      `=IFERROR(GOOGLEFINANCE("${ticker}", "changepct")/100, 0)`,
      `=IFERROR(IFS(
          B${rowNum} < Q${rowNum}, "STOP LOSS",
          OR(B${rowNum} >= R${rowNum}, U${rowNum} > 1.0, O${rowNum} > 75), "TAKE PROFIT",
          B${rowNum} < M${rowNum}, "AVOID (BEAR)", 
          AND(B${rowNum} > M${rowNum}, N${rowNum} > 1.2, B${rowNum} > S${rowNum}), "STRONG BREAKOUT",
          AND(B${rowNum} > M${rowNum}, U${rowNum} < 0.15, O${rowNum} < 35), "STRONG BUY (OVERSOLD)",
          AND(B${rowNum} > M${rowNum}, P${rowNum} = "BULLISH DIV"), "ACCUMULATE (DIV)",
          AND(B${rowNum} > M${rowNum}, B${rowNum} > K${rowNum}, B${rowNum} < S${rowNum}), "RIDE TREND",
          TRUE, "WAIT (CHOP)"
      ), "Wait")`,
      `=IF(AND(G${rowNum}<=-0.10, REGEXMATCH(D${rowNum}, "STRONG|ACCUMULATE|RIDE")), "★ EXECUTE", "-")`,
      `=IFERROR(DATA!${columnToLetter(tickerDataStart + 1)}3, "-")`, 
      `=IFERROR((B${rowNum}-F${rowNum})/F${rowNum}, 0)`, 
      `=IFERROR(IF((R${rowNum}-B${rowNum})/MAX(0.01, B${rowNum}-Q${rowNum}) >= 3, "PRIME", "RISKY"), "—")`,
      `=REPT("★", (B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20))) + (B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-50, 0, 50))))`,
      `=IF(B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-200, 0, 200)), "BULL REGIME", "BEAR REGIME")`,
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-50, 0, 50)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-200, 0, 200)), 0), 2)`,
      `=ROUND(IFERROR(OFFSET(DATA!$${columnToLetter(tickerDataStart+5)}$4, ${lastRow}-1, 0) / AVERAGE(OFFSET(DATA!$${columnToLetter(tickerDataStart+5)}$4, ${lastRow}-21, 0, 20)), 1), 2)`,
      `=ROUND(IFERROR(100-(100/(1+(AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$4, ${lastRow}-15, 0, 15)-OFFSET(DATA!$${closeCol}$4, ${lastRow}-16, 0, 15)),">0")/ABS(AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$4, ${lastRow}-15, 0, 15)-OFFSET(DATA!$${closeCol}$4, ${lastRow}-16, 0, 15)),"<0"))))), 50), 2)`,
      `=IFERROR(IFS(AND(B${rowNum} < INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), O${rowNum} > 50), "BULLISH DIV", AND(B${rowNum} > INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), O${rowNum} < 50), "BEARISH DIV", TRUE, "CONVERGENT"), "—")`,
      `=ROUND(IFERROR(MIN(OFFSET(DATA!$${columnToLetter(tickerDataStart+3)}$4, ${lastRow}-21, 0, 20)), 0), 2)`,
      `=ROUND(B${rowNum} + ((B${rowNum}-Q${rowNum}) * 3), 2)`,
      `=ROUND(IFERROR(MAX(OFFSET(DATA!$${columnToLetter(tickerDataStart+2)}$4, ${lastRow}-51, 0, 50)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(OFFSET(DATA!$${columnToLetter(tickerDataStart+2)}$4, ${lastRow}-14, 0, 14)-OFFSET(DATA!$${columnToLetter(tickerDataStart+3)}$4, ${lastRow}-14, 0, 14))), 0), 2)`,
      `=ROUND(IFERROR(((B${rowNum}-AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20))) / (4*STDEV(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)))) + 0.5, 0.5), 2)`,
      `=IFS(D${rowNum}="STOP LOSS", "🛑 STOP: Price $"&B${rowNum}&" broke Floor $"&Q${rowNum}&".", D${rowNum}="TAKE PROFIT", "💰 EXIT: Target $"&R${rowNum}&" or RSI "&O${rowNum}&" overheated.", D${rowNum}="AVOID (BEAR)", "⚠️ BEAR: Price < SMA200 ($"&M${rowNum}&"). No Longs.", D${rowNum}="STRONG BUY (OVERSOLD)", "🟢 SNIPE: Bull Trend > SMA200 ($"&M${rowNum}&") + Deep Value (RSI "&O${rowNum}&").", D${rowNum}="STRONG BREAKOUT", "🚀 BREAK: Cleared $"&S${rowNum}&" on "&TEXT(N${rowNum}, "0.0")&"x Vol.", D${rowNum}="ACCUMULATE (DIV)", "🔋 LOAD: Bullish Divergence detected in Uptrend.", TRUE, "⏳ WAIT: Consolidation between $"&Q${rowNum}&" and $"&S${rowNum}&".")`
    ]);
  });
  calcSheet.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  calcSheet.getRange(3, 2, formulas.length, 21).setFormulas(formulas);
  
  SpreadsheetApp.flush();
  const range = calcSheet.getRange(2, 1, tickers.length + 1, 22);
  range.setHorizontalAlignment("left");
  range.setBorder(true, true, true, true, true, true, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID);
  
  calcSheet.getRange("C3:C").setNumberFormat("0.00%");
  calcSheet.getRange("G3:G").setNumberFormat("0.00%");
  calcSheet.getRange("U3:U").setNumberFormat("0.00%");
  
  calcSheet.getRange(3, 1, tickers.length, 22).sort({column: 3, ascending: false});
}

function generateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;
  const tickers = getCleanTickers(inputSheet);
  let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");
  dataSheet.clear().clearFormats();
  
  dataSheet.getRange("G1").setValue("Last Update: " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm")).setFontWeight("bold").setFontColor("blue");

  tickers.forEach((ticker, i) => {
    const colStart = (i * 7) + 1;
    dataSheet.getRange(1, colStart).setNumberFormat("@").setValue(ticker).setFontWeight("bold");
    dataSheet.getRange(2, colStart).setValue("ATH:");
    dataSheet.getRange(2, colStart + 1).setFormula(`=MAX(QUERY(GOOGLEFINANCE("${ticker}", "high", "1/1/2000", TODAY()), "SELECT Col2 LABEL Col2 ''"))`);
    dataSheet.getRange(3, colStart).setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "all", TODAY()-800, TODAY()), "No Data")`);
    dataSheet.getRange(4, colStart, 1000, 1).setNumberFormat("yyyy-mm-dd");
    dataSheet.getRange(4, colStart + 1, 1000, 5).setNumberFormat("#,##0.00");
  });
}

function getCleanTickers(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange(3, 1, lastRow-2, 1).getValues().flat().filter(t => t && t.toString().trim() !== "").map(t => t.toString().toUpperCase());
}
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) { temp = (column - 1) % 26; letter = String.fromCharCode(temp + 65) + letter; column = (column - temp - 1) / 26; }
  return letter;
}
function forceExpandSheet(sheet, targetCols) {
  if (sheet.getMaxColumns() < targetCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), targetCols - sheet.getMaxColumns());
}