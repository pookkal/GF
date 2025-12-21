/**
 * ==============================================================================
 * BASELINE LABEL: STABLE_MASTER_V12_CLEAN
 * DATE: 21 DEC 2025
 * FIX: Restores 12-Month/30-Day dropdowns, Removes Target from Chart, 
 * Restores original row alignment.
 * ==============================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Institutional Terminal')
    .addItem('1. Reset & Fetch Data', 'generateDataSheet')
    .addItem('2. Build Dashboard', 'generateCalculationsSheet')
    .addItem('3. Setup Chart View', 'setupChartSheet')
    .addToUi();
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
 * 2. CALCULATIONS ENGINE
 */
function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet) return;
  const tickers = getCleanTickers(inputSheet);
  let calcSheet = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");
  calcSheet.clear().clearFormats();
  calcSheet.setFrozenRows(2);

  const headers = [["Ticker", "Price", "Change %", "DECISION", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance"]];
  calcSheet.getRange(2, 1, 1, 16).setValues(headers).setFontWeight("bold").setBackground("#212121").setFontColor("white");

  const formulas = [];
  const tickerNames = [];
  tickers.forEach((ticker, i) => {
    const rowNum = i + 3;
    const tickerDataStart = (i * 7) + 1;
    const closeCol = columnToLetter(tickerDataStart + 4);
    const lowCol = columnToLetter(tickerDataStart + 3);
    const highCol = columnToLetter(tickerDataStart + 2);
    const volCol = columnToLetter(tickerDataStart + 5);
    const lastRow = `COUNTA(DATA!$${closeCol}:$${closeCol})`;
    
    const s20 = `AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-20, 0, 20))`;
    const s50 = `AVERAGE(OFFSET(DATA!$${closeCol}$1, ${lastRow}-50, 0, 50))`;
    const s200 = `AVERAGE(OFFSET(DATA!$${closeCol}$1, MAX(1, ${lastRow}-200), 0, 200))`;
    const rsiRange = `OFFSET(DATA!$${closeCol}$1, ${lastRow}-15, 0, 15)`;
    const rsiFormula = `100-(100/(1+(MAX(0,AVERAGEIF(ARRAYFORMULA(${rsiRange}-OFFSET(${rsiRange},-1,0)),">0"))/MAX(0.0001, ABS(AVERAGEIF(ARRAYFORMULA(${rsiRange}-OFFSET(${rsiRange},-1,0)),"<0"))))))`;

    tickerNames.push([ticker]);
    formulas.push([
      `=ROUND(IFERROR(GOOGLEFINANCE("${ticker}", "price")), 2)`,
      `=IFERROR(GOOGLEFINANCE("${ticker}", "changepct")/100, 0)`,
      `=IFS(B${rowNum} < N${rowNum}, "🚨 EXIT (STOP)", B${rowNum} >= O${rowNum}, "💰 TAKE PROFIT", TRUE, "Wait")`,
      `=IF((O${rowNum}-B${rowNum})/MAX(0.01, B${rowNum}-N${rowNum}) >= 3, "💎 HIGH", "⚖️ MED")`,
      `=REPT("★", (B${rowNum}>${s20}) + (B${rowNum}>${s50}) + (B${rowNum}>${s200}))`,
      `=IF(B${rowNum}>${s200}, "🚀 BULLISH", "📉 BEARISH")`,
      `=ROUND(IFERROR(${s20}, 0), 2)`,
      `=ROUND(IFERROR(${s50}, 0), 2)`,
      `=ROUND(IFERROR(${s200}, 0), 2)`,
      `=ROUND(OFFSET(DATA!$${volCol}$1, ${lastRow}-1, 0) / MAX(0.01, AVERAGE(OFFSET(DATA!$${volCol}$1, ${lastRow}-21, 0, 20))), 2)`,
      `=ROUND(IFERROR(${rsiFormula}, 50), 2)`,
      `=IF(AND(B${rowNum} < OFFSET(DATA!$${closeCol}$1, ${lastRow}-5, 0), L${rowNum} > OFFSET(DATA!$${closeCol}$1, ${lastRow}-10, 0)), "🐂 BULL DIV", "-")`,
      `=ROUND(IFERROR(MIN(OFFSET(DATA!$${lowCol}$1, ${lastRow}-21, 0, 20)), 0), 2)`,
      `=ROUND(B${rowNum} + ((B${rowNum}-N${rowNum}) * 3), 2)`,
      `=ROUND(IFERROR(MAX(OFFSET(DATA!$${highCol}$1, ${lastRow}-51, 0, 50)), 0), 2)`
    ]);
  });
  calcSheet.getRange(3, 1, tickerNames.length, 1).setValues(tickerNames);
  calcSheet.getRange(3, 2, formulas.length, 15).setFormulas(formulas);
  calcSheet.getRange(3, 3, tickerNames.length, 1).setNumberFormat("0.00%");
  calcSheet.getRange(3, 8, formulas.length, 9).setNumberFormat("0.00");
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

  // DROPDOWN FIX: Months (0-12) | Days (0-30)
  const monthRule = SpreadsheetApp.newDataValidation().requireValueInList([0,1,2,3,4,5,6,7,8,9,10,11,12]).build();
  const daysRule = SpreadsheetApp.newDataValidation().requireValueInList(Array.from({length: 31}, (_, i) => i)).build();
  
  chartSheet.getRange("A2:C2").setValues([["Years", "Months", "Days"]]).setBackground("#222").setFontColor("#FFF").setHorizontalAlignment("center");
  chartSheet.getRange("A3:B3").setDataValidation(monthRule).setHorizontalAlignment("center");
  chartSheet.getRange("C3").setDataValidation(daysRule).setHorizontalAlignment("center");
  chartSheet.getRange("A3:C3").setValues([[0, 3, 0]]);
  
  chartSheet.getRange("A4").setValue("START:").setFontWeight("bold");
  chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)").setNumberFormat("yyyy-mm-dd");

  const t = "B1";
  const labels = [["DECISION"], ["PRICE"], ["CHANGE %"], ["TREND STATE"], ["R:R QUALITY"], ["RSI"], ["52W HIGH"], ["52W LOW"], ["PE RATIO"], ["EPS"], ["BETA"], ["YIELD"], ["SMA 20"], ["SMA 50"], ["SMA 200"], ["SUPPORT"], ["TARGET"], ["RESISTANCE"], ["REL VOL"], ["PREV CLOSE"], ["DIFF"], ["DIFF %"], ["TREND SCORE"], ["DIVERGENCE"], ["MARKET CAP"], ["DIVIDEND"]];
  const formulas = [[`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 4, 0), "Wait")`], [`=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`], [`=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 7, 0), "—")`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 5, 0), "—")`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 12, 0), 50)`], [`=GOOGLEFINANCE(${t}, "high52")`], [`=GOOGLEFINANCE(${t}, "low52")`], [`=GOOGLEFINANCE(${t}, "pe")`], [`=GOOGLEFINANCE(${t}, "eps")`], [`=GOOGLEFINANCE(${t}, "beta")`], [`=IFERROR(GOOGLEFINANCE(${t}, "yield")/100, 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 8, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 9, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 10, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 14, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 15, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 16, 0), 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 11, 0), 1)`], [`=GOOGLEFINANCE(${t}, "closeyest")`], [`=B7-B25`], [`=IFERROR(B26/B25, 0)`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 6, 0), "—")`], [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 13, 0), "—")`], [`=GOOGLEFINANCE(${t}, "marketcap")`], [`=GOOGLEFINANCE(${t}, "dividend")`]];
  
  chartSheet.getRange(6, 1, labels.length, 1).setValues(labels).setFontWeight("bold").setBackground("#EEE");
  chartSheet.getRange(6, 2, formulas.length, 1).setFormulas(formulas);
  chartSheet.getRange("B7:B31").setNumberFormat("0.00").setHorizontalAlignment("left");
  
  updateDynamicChart();
}

/**
 * 4. CHART ENGINE (NO TARGET ON CHART)
 */
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;

  const ticker = sheet.getRange("B1").getValue();
  const startDate = sheet.getRange("B4").getValue();
  const isWeekly = sheet.getRange("D2").getValue() === "WEEKLY";

  const supportVal = sheet.getRange("B21").getValue();
  const resistanceVal = sheet.getRange("B23").getValue();

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

    // Col Structure: Date, Price, BullVol, BearVol, SMA20, SMA50, SMA200, Resistance, Support
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
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(3, 26, masterData.length, 9))
    .setOption('series', {
      0: {type: 'line', color: '#1A73E8', lineWidth: 3, labelInLegend: 'Price'},
      1: {type: 'bars', color: '#2E7D32', targetAxisIndex: 1, labelInLegend: 'Bull Vol'},
      2: {type: 'bars', color: '#C62828', targetAxisIndex: 1, labelInLegend: 'Bear Vol'},
      3: {type: 'line', color: '#FBBC04', lineWidth: 1.5, labelInLegend: 'SMA 20'},
      4: {type: 'line', color: '#9C27B0', lineWidth: 1.5, labelInLegend: 'SMA 50'},
      5: {type: 'line', color: '#FF9800', lineWidth: 2, labelInLegend: 'SMA 200'},
      6: {type: 'line', color: '#B71C1C', lineDashStyle: [4, 4], labelInLegend: 'Resistance'},
      7: {type: 'line', color: '#0D47A1', lineDashStyle: [4, 4], labelInLegend: 'Support'}
    })
    .setOption('vAxes', { 
      0: {viewWindow: {min: minPrice, max: maxPrice}, gridlines: {color: '#f0f0f0'}}, 
      1: {viewWindow: {min: 0, max: Math.max(...viewVols)*5}, textStyle: {color: 'none'}, gridlines: {count: 0}} 
    })
    .setOption('legend', {position: 'top'})
    .setOption('chartArea', {left: '10%', top: '10%', width: '80%', height: '75%'})
    .setPosition(4, 3, 10, 10).setOption('width', 1150).setOption('height', 650).build();
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