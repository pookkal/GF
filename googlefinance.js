/**
 * ==============================================================================
 * BASELINE LABEL: GF_WORKING_CHART_REASON_V14_FIXED
 * NAME: _21DEC_FINAL_DEEP_REASON_EDITION
 * DATE: 21 DEC 2025
 * STATUS: GOLDEN BASELINE - STABLE
 * ==============================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Institutional Terminal')
    .addItem('🚀 1-CLICK REBUILD ALL', 'FlushAllSheetsAndBuild')
    .addSeparator()
    .addItem('1. Fetch Data Only', 'generateDataSheet')
    .addItem('2. Build Dashboard Only', 'generateCalculationsSheet')
    .addItem('3. Setup Chart Only', 'setupChartSheet')
    .addToUi();
}

/**
 * MASTER ORCHESTRATOR (LOGIC SHEET REMOVED)
 */
function FlushAllSheetsAndBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ["DATA", "CALCULATIONS", "CHART"];
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert('🚨 Full Rebuild', 'Delete and rebuild all terminal sheets? (INPUT sheet will be preserved)', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  sheetsToDelete.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>Step 1/3:</b> Data Acquisition..."), "Terminal Status");
  generateDataSheet();
  SpreadsheetApp.flush();
  Utilities.sleep(5000); 

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>Step 2/3:</b> Building Indicator Matrix..."), "Terminal Status");
  generateCalculationsSheet();
  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  ui.showModelessDialog(HtmlService.createHtmlOutput("<b>Step 3/3:</b> Rendering Terminal View..."), "Terminal Status");
  setupChartSheet();
  
  ui.alert('✅ Terminal Online', 'System, Indicators, and Deep Reasoning Synchronized.', ui.ButtonSet.OK);
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "CHART") return;
  const watchList = ["B1", "D2", "A3", "B3", "C3"];
  if (watchList.includes(e.range.getA1Notation())) {
    updateDynamicChart();
  }
}

/**
 * 1. DATA ENGINE
 */
function generateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) return;
  const tickers = getCleanTickers(inputSheet);
  let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");
  dataSheet.clear().clearFormats();
  tickers.forEach((ticker, i) => {
    const colStart = (i * 7) + 1;
    dataSheet.getRange(1, colStart).setNumberFormat("@").setValue(ticker).setFontWeight("bold");
    dataSheet.getRange(2, colStart).setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "all", TODAY()-800, TODAY()), "No Data")`);
    dataSheet.getRange(3, colStart, 1000, 1).setNumberFormat("yyyy-mm-dd");
    dataSheet.getRange(3, colStart + 1, 1000, 5).setNumberFormat("#,##0.00");
  });
}

/**
 * 2. CALCULATIONS ENGINE (DEEP DIVE REASONING)
 */
function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet) return;
  const tickers = getCleanTickers(inputSheet);
  if (tickers.length === 0) return;

  let calcSheet = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");
  calcSheet.clear().clearFormats();
  calcSheet.setFrozenRows(2);

  const headers = [["Ticker", "Price", "Change %", "DECISION", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "REASONING"]];
  calcSheet.getRange(2, 1, 1, 19).setValues(headers).setFontWeight("bold").setBackground("#212121").setFontColor("white");

  const formulas = [];
  const tickerNames = [];
  tickers.forEach((ticker, i) => {
    const rowNum = i + 3;
    const tickerDataStart = (i * 7) + 1;
    const highCol = columnToLetter(tickerDataStart + 2);
    const lowCol = columnToLetter(tickerDataStart + 3);
    const closeCol = columnToLetter(tickerDataStart + 4);
    const volCol = columnToLetter(tickerDataStart + 5);
    const lastRow = `COUNTA(DATA!$${closeCol}:$${closeCol})`;
    
    const s20 = `IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-20, 0, 20)), 0)`;
    const s50 = `IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-50, 0, 50)), 0)`;
    const s200 = `IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$1, MAX(1, ${lastRow}-200), 0, MIN(200, ${lastRow}))), 0)`;
    const rsiRange = `OFFSET(DATA!$${closeCol}$1, ${lastRow}-15, 0, 15)`;
    const rsiFormula = `100-(100/(1+(MAX(0,AVERAGEIF(ARRAYFORMULA(${rsiRange}-OFFSET(${rsiRange},-1,0)),">0"))/MAX(0.0001, ABS(AVERAGEIF(ARRAYFORMULA(${rsiRange}-OFFSET(${rsiRange},-1,0)),"<0"))))))`;
    const stdev = `IFERROR(STDEV(OFFSET(DATA!$${closeCol}$1, ${lastRow}-20, 0, 20)), 1)`;

    tickerNames.push([ticker]);
    formulas.push([
      `=ROUND(IFERROR(GOOGLEFINANCE("${ticker}", "price")), 2)`,
      `=IFERROR(GOOGLEFINANCE("${ticker}", "changepct")/100, 0)`,
      `=IFERROR(IFS(B${rowNum} < N${rowNum}, "🚨 EXIT", B${rowNum} >= O${rowNum}, "💰 PROFIT", AND(R${rowNum}<0.18, L${rowNum}<45), "🎯 BUY DIP", AND(B${rowNum}>I${rowNum}, K${rowNum}>1.2, Q${rowNum}<(B${rowNum}*0.06)), "🚀 STRONG BUY", TRUE, "⚖️ HOLD"), "Wait")`,
      `=IFERROR(IF((O${rowNum}-B${rowNum})/MAX(0.01, B${rowNum}-N${rowNum}) >= 3, "💎 HIGH", "⚖️ MED"), "—")`,
      `=IFERROR(REPT("★", (B${rowNum}>${s20}) + (B${rowNum}>${s50}) + (B${rowNum}>${s200})), "—")`,
      `=IFERROR(IF(B${rowNum}>${s200}, "🚀 BULLISH", "📉 BEARISH"), "—")`,
      `=ROUND(${s20}, 2)`, `=ROUND(${s50}, 2)`, `=ROUND(${s200}, 2)`,
      `=ROUND(IFERROR(OFFSET(DATA!$${volCol}$1, ${lastRow}-1, 0) / MAX(0.01, AVERAGE(OFFSET(DATA!$${volCol}$1, ${lastRow}-21, 0, 20))), 1), 2)`,
      `=ROUND(IFERROR(${rsiFormula}, 50), 2)`,
      `=IF(AND(B${rowNum} < OFFSET(DATA!$${closeCol}$1, ${lastRow}-5, 0), L${rowNum} > OFFSET(DATA!$${closeCol}$1, ${lastRow}-10, 0)), "🐂 BULL DIV", "-")`,
      `=ROUND(IFERROR(MIN(OFFSET(DATA!$${lowCol}$1, ${lastRow}-21, 0, 20)), 0), 2)`,
      `=ROUND(IFERROR(B${rowNum} + ((B${rowNum}-N${rowNum}) * 3), 0), 2)`,
      `=ROUND(IFERROR(MAX(OFFSET(DATA!$${highCol}$1, ${lastRow}-51, 0, 50)), 0), 2)`,
      `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(OFFSET(DATA!$${highCol}$1, ${lastRow}-14, 0, 14)-OFFSET(DATA!$${lowCol}$1, ${lastRow}-14, 0, 14))), 0), 2)`, // ATR
      `=ROUND(IFERROR((B${rowNum}-(${s20}-2*${stdev})) / (MAX(0.01, 4*${stdev})), 0.5), 2)`, // Bollinger %B
      `="TREND: "&G${rowNum}&" | PRICE ($"&B${rowNum}&") vs 50-SMA ($"&I${rowNum}&"). MOMENTUM: RSI is "&IF(L${rowNum}>70,"Overbought",IF(L${rowNum}<45,"Oversold","Neutral"))&" ("&L${rowNum}&"). VOLATILITY: ATR is "&Q${rowNum}&" ("&TEXT(Q${rowNum}/B${rowNum},"0.0%")&" of Price). %B shows price is at "&TEXT(R${rowNum},"0%")&" of Volatility Band. VOLUME: Relative Trend is "&K${rowNum}&"x average."`
    ]);
  });

  calcSheet.getRange(3, 1, tickerNames.length, 1).setValues(tickerNames);
  calcSheet.getRange(3, 2, formulas.length, 18).setFormulas(formulas);
  calcSheet.getRange(3, 1, tickerNames.length, 19).setHorizontalAlignment("left").setVerticalAlignment("middle");
  calcSheet.getRange(3, 3, tickerNames.length, 1).setNumberFormat("0.00%");
  calcSheet.getRange(3, 8, formulas.length, 11).setNumberFormat("0.00");

  try { 
    calcSheet.getRange("H1:J1").merge().setValue("TREND (SMAs)").setBackground("#E3F2FD").setFontWeight("bold").setHorizontalAlignment("center");
    calcSheet.getRange("K1:M1").merge().setValue("MOMENTUM").setBackground("#F3E5F5").setFontWeight("bold").setHorizontalAlignment("center");
    calcSheet.getRange("N1:P1").merge().setValue("S/R LEVELS").setBackground("#E8F5E9").setFontWeight("bold").setHorizontalAlignment("center");
    calcSheet.getRange("Q1:R1").merge().setValue("VOLATILITY").setBackground("#FFF3E0").setFontWeight("bold").setHorizontalAlignment("center");
    calcSheet.getRange("H:J").shiftColumnGroupDepth(1); 
    calcSheet.getRange("K:M").shiftColumnGroupDepth(1); 
    calcSheet.getRange("N:P").shiftColumnGroupDepth(1); 
    calcSheet.getRange("Q:R").shiftColumnGroupDepth(1); 
  } catch(e) {}
}

/**
 * 3. CHART SHEET SETUP
 */
function setupChartSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  const tickers = getCleanTickers(inputSheet);
  let chartSheet = ss.getSheetByName("CHART") || ss.insertSheet("CHART");
  forceExpandSheet(chartSheet, 50);
  chartSheet.clear().clearFormats();

  chartSheet.getRange("A1").setValue("TICKER:").setFontWeight("bold");
  chartSheet.getRange("B1").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tickers).build()).setValue(tickers[0]).setBackground("#FFF9C4");
  chartSheet.getRange("D1").setValue("VIEW:").setFontWeight("bold");
  chartSheet.getRange("D2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"]).build()).setValue("DAILY").setBackground("#E1F5FE");

  const monthRule = SpreadsheetApp.newDataValidation().requireValueInList([0,1,2,3,4,5,6,7,8,9,10,11,12]).build();
  chartSheet.getRange("A2:C2").setValues([["Years", "Months", "Days"]]).setBackground("#222").setFontColor("#FFF").setHorizontalAlignment("center");
  chartSheet.getRange("A3:B3").setDataValidation(monthRule).setHorizontalAlignment("center");
  chartSheet.getRange("C3").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from({length:31},(_,i)=>i)).build()).setHorizontalAlignment("center");
  chartSheet.getRange("A3:C3").setValues([[0, 3, 0]]);
  chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)").setNumberFormat("yyyy-mm-dd");

  const t = "B1";
  const labels = [["SIGNAL"], ["REASONING"], ["ATR (VOL)"], ["BOLLINGER %B"], ["PRICE"], ["CHANGE %"], ["TREND STATE"], ["R:R QUALITY"], ["RSI"], ["52W HIGH"], ["52W LOW"], ["PE RATIO"], ["EPS"], ["BETA"], ["YIELD"], ["SMA 20"], ["SMA 50"], ["SMA 200"], ["SUPPORT"], ["TARGET"], ["RESISTANCE"], ["REL VOL"], ["PREV CLOSE"], ["DIFF"], ["DIFF %"], ["TREND SCORE"], ["DIVERGENCE"], ["MARKET CAP"], ["DIVIDEND"]];
  const formulas = [
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 4, 0), "Wait")`],
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 19, 0), "—")`],
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 17, 0), 0)`],
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 18, 0), 0.5)`],
    [`=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`], [`=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 7, 0), "—")`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 5, 0), "—")`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 12, 0), 50)`], 
    [`=GOOGLEFINANCE(${t}, "high52")`], [`=GOOGLEFINANCE(${t}, "low52")`], [`=GOOGLEFINANCE(${t}, "pe")`], [`=GOOGLEFINANCE(${t}, "eps")`], [`=GOOGLEFINANCE(${t}, "beta")`], [`=IFERROR(GOOGLEFINANCE(${t}, "yield")/100, 0)`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 8, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 9, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 10, 0), 0)`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 14, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 15, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 16, 0), 0)`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 11, 0), 1)`], [`=GOOGLEFINANCE(${t}, "closeyest")`], [`=B10-B28`], [`=IFERROR(B29/B28, 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 6, 0), "—")`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$S, 13, 0), "—")`], [`=GOOGLEFINANCE(${t}, "marketcap")`], [`=GOOGLEFINANCE(${t}, "dividend")`]
  ];
  
  chartSheet.getRange(6, 1, labels.length, 1).setValues(labels).setFontWeight("bold").setBackground("#EEE");
  chartSheet.getRange(6, 2, formulas.length, 1).setFormulas(formulas);
  chartSheet.getRange("B6:B34").setHorizontalAlignment("left"); 
  chartSheet.getRange("B10:B34").setNumberFormat("0.00");
  chartSheet.getRange("B7").setWrap(true);
  updateDynamicChart();
}

/**
 * 4. CHART ENGINE
 */
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;
  const ticker = sheet.getRange("B1").getValue();
  const startDate = sheet.getRange("B4").getValue();
  const isWeekly = sheet.getRange("D2").getValue() === "WEEKLY";
  const supportVal = sheet.getRange("B24").getValue();
  const resistanceVal = sheet.getRange("B26").getValue();
  const rawHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIdx = rawHeaders.indexOf(ticker);
  if (colIdx === -1) return;
  const rawData = dataSheet.getRange(1, colIdx + 1, dataSheet.getLastRow(), 6).getValues();
  let masterData = [];
  let pricesAll = [];
  let viewPrices = [supportVal, resistanceVal];
  let viewVols = [];
  for (let i = 2; i < rawData.length; i++) {
    let p = Number(rawData[i][4]);
    if (!isNaN(p) && p > 0.01) pricesAll.push(p);
  }
  let dataCounter = 0;
  for (let i = 2; i < rawData.length; i++) {
    let row = rawData[i];
    let d = row[0];
    let close = Number(row[4]);
    let vol = Number(row[5]);
    if (!d || !(d instanceof Date) || isNaN(close) || close < 0.01) continue;
    dataCounter++;
    if (d < startDate) continue;
    if (isWeekly && d.getDay() !== 5) continue;
    let s20 = dataCounter >= 20 ? Number((pricesAll.slice(dataCounter-20, dataCounter).reduce((a,b)=>a+b,0)/20).toFixed(2)) : null;
    let s50 = dataCounter >= 50 ? Number((pricesAll.slice(dataCounter-50, dataCounter).reduce((a,b)=>a+b,0)/50).toFixed(2)) : null;
    let s200 = dataCounter >= 200 ? Number((pricesAll.slice(dataCounter-200, dataCounter).reduce((a,b)=>a+b,0)/200).toFixed(2)) : null;
    masterData.push([Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "MMM dd"), close, (close >= (i>2?rawData[i-1][4]:close))?vol:null, (close < (i>2?rawData[i-1][4]:close))?vol:null, s20, s50, s200, resistanceVal, supportVal]);
    viewPrices.push(close);
    if(s20) viewPrices.push(s20); if(s50) viewPrices.push(s50); if(s200) viewPrices.push(s200);
    viewVols.push(vol);
  }
  if (masterData.length === 0) return;
  const minPrice = Math.min(...viewPrices.filter(v => v > 0)) * 0.98;
  const maxPrice = Math.max(...viewPrices.filter(v => v > 0)) * 1.02;
  sheet.getRange(3, 26, 1500, 9).clearContent();
  sheet.getRange(3, 26, masterData.length, 9).setValues(masterData);
  SpreadsheetApp.flush();
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  const chart = sheet.newChart().setChartType(Charts.ChartType.COMBO).addRange(sheet.getRange(3, 26, masterData.length, 9))
    .setOption('series', {0: {type: 'line', color: '#1A73E8', lineWidth: 3, labelInLegend: 'Price'}, 1: {type: 'bars', color: '#2E7D32', targetAxisIndex: 1, labelInLegend: 'Bull Vol'}, 2: {type: 'bars', color: '#C62828', targetAxisIndex: 1, labelInLegend: 'Bear Vol'}, 3: {type: 'line', color: '#FBBC04', lineWidth: 1.5, labelInLegend: 'SMA 20'}, 4: {type: 'line', color: '#9C27B0', lineWidth: 1.5, labelInLegend: 'SMA 50'}, 5: {type: 'line', color: '#FF9800', lineWidth: 2, labelInLegend: 'SMA 200'}, 6: {type: 'line', color: '#B71C1C', lineDashStyle: [4, 4], labelInLegend: 'Resistance'}, 7: {type: 'line', color: '#0D47A1', lineDashStyle: [4, 4], labelInLegend: 'Support'}})
    .setOption('vAxes', {0: {viewWindow: {min: minPrice, max: maxPrice}, gridlines: {color: '#f0f0f0'}}, 1: {viewWindow: {min: 0, max: Math.max(...viewVols)*5}, textStyle: {color: 'none'}, gridlines: {count: 0}}})
    .setOption('legend', {position: 'top'}).setPosition(4, 3, 10, 10).setOption('width', 1150).setOption('height', 650).build();
  sheet.insertChart(chart);
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