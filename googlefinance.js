/**
 * 1. MASTER MENU SETUP
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Institutional Analysis')
    .addItem('1. Reset & Fetch Data', 'generateDataSheet')
    .addItem('2. Build Technical Dashboard', 'generateCalculationsSheet')
    .addItem('3. Setup CHART Deep-Dive', 'setupChartSheet')
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
 * 2. DATA MODULE
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
  });
}

/**
 * 3. DASHBOARD MODULE
 */
function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet || dataSheet.getLastRow() < 10) return;
  const tickers = getCleanTickers(inputSheet);
  let calcSheet = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");
  
  calcSheet.clear().clearFormats();
  calcSheet.setFrozenRows(2);
  calcSheet.setFrozenColumns(1);

  calcSheet.getRange("A1").setValue("ASSET").setFontWeight("bold");
  calcSheet.getRange("B1:E1").merge().setValue("PORTFOLIO STRATEGY").setBackground("#E1F5FE").setFontColor("#01579B").setFontWeight("bold").setHorizontalAlignment("center");

  const headers = [["Ticker", "DECISION", "Price", "Change %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance"]];
  calcSheet.getRange(2, 1, 1, 16).setValues(headers).setFontWeight("bold").setBackground("#212121").setFontColor("white");

  const formulas = [];
  tickers.forEach((ticker, i) => {
    const row = i + 3;
    const colIdx = (i * 7) + 1;
    const closeCol = columnToLetter(colIdx + 4); 
    const count = `COUNTA(DATA!$${closeCol}:$${closeCol})`;
    
    formulas.push([
      ticker, 
      `=IF(C${row}<N${row}, "⚠️ EXIT", IF(C${row}>=O${row}, "💰 TAKE PROFIT", IF(AND(L${row}<45, M${row}="🐂 BULL DIV"), "🎯 BUY DIP", "Wait")))`, 
      `=ROUND(IFERROR(GOOGLEFINANCE("${ticker}", "price"), OFFSET(DATA!$${closeCol}$1, ${count}-1, 0)), 2)`, 
      `=IFERROR(GOOGLEFINANCE("${ticker}", "changepct")/100, 0)`,
      `=IF((O${row}-C${row})/(C${row}-N${row}) >= 3, "💎 HIGH", "⚖️ MED")`,
      `=REPT("★", (C${row}>AVERAGE(OFFSET(DATA!$${closeCol}$1, ${count}-20, 0, 20))) + (C${row}>AVERAGE(OFFSET(DATA!$${closeCol}$1, ${count}-50, 0, 50))) + (C${row}>AVERAGE(OFFSET(DATA!$${closeCol}$1, MAX(1, ${count}-200), 0, 200))))`,
      `=IF(LEN(F${row})=3, "🚀 BULLISH", IF(LEN(F${row})=0, "📉 BEARISH", "⚖️ NEUTRAL"))`,
      `=ROUND(AVERAGE(OFFSET(DATA!$${closeCol}$1, ${count}-20, 0, 20)), 2)`,
      `=ROUND(AVERAGE(OFFSET(DATA!$${closeCol}$1, ${count}-50, 0, 50)), 2)`,
      `=ROUND(AVERAGE(OFFSET(DATA!$${closeCol}$1, MAX(1, ${count}-200), 0, 200)), 2)`,
      `=ROUND(OFFSET(DATA!$${columnToLetter(colIdx+5)}$1, ${count}-1, 0)/AVERAGE(OFFSET(DATA!$${columnToLetter(colIdx+5)}$1, ${count}-21, 0, 20)), 2)`,
      `=ROUND(IFERROR(100-(100/(1+(MAX(0,AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$1, ${count}-14, 0, 14)-OFFSET(DATA!$${closeCol}$1, ${count}-15, 0, 14)),">0"))/MAX(0.0001, ABS(AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$1, ${count}-14, 0, 14)-OFFSET(DATA!$${closeCol}$1, ${count}-15, 0, 14)),"<0")))))), 50), 2)`,
      `=IF(AND(C${row} < OFFSET(DATA!$${closeCol}$1, ${count}-6, 0)*1.01, L${row} > OFFSET(DATA!$${closeCol}$1, ${count}-7, 0)), "🐂 BULL DIV", "-")`,
      `=ROUND(MIN(OFFSET(DATA!$${columnToLetter(colIdx+3)}$1, ${count}-20, 0, 20)), 2)`,
      `=ROUND(C${row} + ((C${row}-N${row})*3), 2)`,
      `=ROUND(MAX(OFFSET(DATA!$${columnToLetter(colIdx+2)}$1, ${count}-50, 0, 50)), 2)`
    ]);
  });
  
  calcSheet.getRange(3, 1, formulas.length, 16).setValues(formulas).setHorizontalAlignment("left");
}

/**
 * ==========================================
 * 4. CHART SHEET SETUP (UI + Dropdowns)
 * ==========================================
 */
function setupChartSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  const tickers = getCleanTickers(inputSheet);
  let chartSheet = ss.getSheetByName("CHART") || ss.insertSheet("CHART");
  
  chartSheet.clear(); 

  // --- CONTROLS ---
  chartSheet.getRange("A1").setValue("TICKER:").setFontWeight("bold");
  chartSheet.getRange("B1").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tickers).build()).setValue(tickers[0]).setBackground("#FFF9C4");
  
  chartSheet.getRange("D1").setValue("VIEW:").setFontWeight("bold");
  chartSheet.getRange("D2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"]).build()).setValue("DAILY").setBackground("#E1F5FE");

  // --- DATE DROPDOWNS (Fixing the "No Dropdown" Issue) ---
  const numberList = [0,1,2,3,4,5,6,7,8,9,10,11,12];
  const dropdownRule = SpreadsheetApp.newDataValidation().requireValueInList(numberList).build();

  chartSheet.getRange("A2:C2").setValues([["Years", "Months", "Days"]]).setBackground("#222").setFontColor("#FFF").setHorizontalAlignment("center");
  
  chartSheet.getRange("A3").setDataValidation(dropdownRule).setValue(0); // Years
  chartSheet.getRange("B3").setDataValidation(dropdownRule).setValue(3); // Months
  chartSheet.getRange("C3").setDataValidation(dropdownRule).setValue(0); // Days
  
  chartSheet.getRange("A4").setValue("START:").setFontWeight("bold");
  chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)").setNumberFormat("yyyy-mm-dd");

  // --- INFO TABLE (Col A data fill fix) ---
  chartSheet.getRange("A6:A7").setValues([["52W HIGH"],["52W LOW"]]).setFontWeight("bold").setBackground("#EEE");
  chartSheet.getRange("B6:B7").setFormulas([
    [`=IFERROR(GOOGLEFINANCE(B1, "high52"), 0)`],
    [`=IFERROR(GOOGLEFINANCE(B1, "low52"), 0)`]
  ]);
  chartSheet.getRange("A1:B7").setBorder(true, true, true, true, true, true, "#999", SpreadsheetApp.BorderStyle.SOLID);
  chartSheet.getRange("B6:B7").setNumberFormat("0.00");

  updateDynamicChart();
}

/**
 * ==========================================
 * 5. CHART ENGINE (Auto-Scale + Clean 52W Range)
 * ==========================================
 */
function updateDynamicChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CHART");
  const dataSheet = ss.getSheetByName("DATA");
  if (!sheet || !dataSheet) return;

  const ticker = sheet.getRange("B1").getValue();
  const startDate = sheet.getRange("B4").getValue();
  const isWeekly = sheet.getRange("D2").getValue() === "WEEKLY";

  // 1. GET 52W LEVELS (Zero Safe)
  let high52 = sheet.getRange("B6").getValue();
  let low52 = sheet.getRange("B7").getValue();
  if (!high52 || high52 === 0) high52 = null;
  if (!low52 || low52 === 0) low52 = null;

  // 2. READ RAW DATA
  const lastCol = dataSheet.getLastColumn();
  const lastRow = dataSheet.getLastRow();
  const rawData = dataSheet.getRange(1, 1, Math.min(2000, lastRow), lastCol).getValues();

  const headers = rawData[0];
  const colIndex = headers.indexOf(ticker);
  
  if (colIndex === -1) {
    sheet.getRange("D6").setValue("Ticker not found.");
    return;
  }

  // 3. BUILD DATA ARRAY
  let masterData = [];
  let minVal = 1000000;
  let maxVal = 0;

  for (let i = 2; i < rawData.length; i++) {
    let row = rawData[i];
    let d = row[colIndex];
    
    if (!d || !(d instanceof Date)) continue;
    if (d < startDate) continue;
    if (isWeekly && d.getDay() !== 5) continue; 

    let close = Number(row[colIndex + 4]);
    if (isNaN(close)) continue;

    // Track for Scaling
    if (close < minVal) minVal = close;
    if (close > maxVal) maxVal = close;

    // Detailed X-Axis Date Format
    let dateStr = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

    masterData.push([dateStr, close, high52, low52]);
  }

  // 4. SMART SCALING
  if (low52 !== null) minVal = Math.min(minVal, low52);
  if (high52 !== null) maxVal = Math.max(maxVal, high52);

  const yMin = minVal * 0.98; 
  const yMax = maxVal * 1.02;

  // 5. PASTE TO HIDDEN AREA (Z3)
  sheet.getRange("Z3:AC").clearContent();

  if (masterData.length === 0) {
    sheet.getRange("D6").setValue("No data found.");
    return;
  }

  sheet.getRange(3, 26, masterData.length, 4).setValues(masterData);

  // 6. DRAW CHART
  const charts = sheet.getCharts();
  charts.forEach(c => sheet.removeChart(c));

  const chartRange = sheet.getRange(3, 26, masterData.length, 4); 

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartRange)
    .setOption('series', {
      0: {labelInLegend: 'Price', color: '#1976D2', lineWidth: 3}, 
      1: {labelInLegend: '52W High', color: '#F57C00', lineWidth: 1.5}, // Orange
      2: {labelInLegend: '52W Low', color: '#7B1FA2', lineWidth: 1.5}   // Purple
    })
    .setOption('curveType', 'function')
    .setOption('useFirstColumnAsDomain', true)
    
    .setOption('vAxis', {
      title: 'Price ($)',
      gridlines: {count: 5, color: '#e0e0e0'},
      viewWindowMode: 'explicit',
      viewWindow: {
        min: yMin,
        max: yMax
      }
    })
    
    .setOption('hAxis', {
      type: 'category', 
      slantedText: true, 
      textStyle: {fontSize: 10}, // Slightly larger detailed font
      maxAlternation: 1
    })
    .setOption('chartArea', {left: '8%', top: '10%', width: '85%', height: '70%'}) // Added breathing room at bottom for dates
    .setOption('title', ticker + ' Institutional Analysis')
    .setOption('legend', {position: 'top'})
    .setPosition(4, 3, 0, 0) // Anchor C4
    .setOption('width', 1100)
    .setOption('height', 550)
    .build();

  sheet.insertChart(chart);
}

// Helpers
function getCleanTickers(sheet) {
  const last = sheet.getLastRow();
  if (last < 3) return [];
  return sheet.getRange(3, 1, last-2, 1).getValues().flat()
    .filter(t => t && t.toString().trim() !== "").map(t => t.toString().toUpperCase());
}
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26; letter = String.fromCharCode(temp + 65) + letter; column = (column - temp - 1) / 26;
  }
  return letter;
}