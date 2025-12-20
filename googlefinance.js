/**
 * ==========================================
 * 1. MASTER MENU & TRIGGERS
 * ==========================================
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📈 Institutional Terminal')
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
 * ==========================================
 * 2. DATA ENGINE (Rounded, Left Aligned, Borders)
 * ==========================================
 */
function generateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  if (!inputSheet) { SpreadsheetApp.getUi().alert("Missing 'INPUT' sheet."); return; }
  
  const tickers = getCleanTickers(inputSheet);
  let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");
  dataSheet.clear();

  tickers.forEach((ticker, i) => {
    const colStart = (i * 7) + 1;
    
    // Header
    dataSheet.getRange(1, colStart).setValue(ticker).setFontWeight("bold").setHorizontalAlignment("left");
    
    // Fetch Data
    dataSheet.getRange(2, colStart).setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "all", TODAY()-800, TODAY()), "No Data")`);
    
    // FORMATTING: Range for 1500 rows, 6 columns (Date + 5 Data Cols)
    const dataRange = dataSheet.getRange(3, colStart, 1500, 6);
    
    // 1. Borders
    dataRange.setBorder(true, true, true, true, true, true, "#D3D3D3", SpreadsheetApp.BorderStyle.SOLID);
    // 2. Left Align
    dataRange.setHorizontalAlignment("left");
    // 3. Date Format (Col 1)
    dataSheet.getRange(3, colStart, 1500, 1).setNumberFormat("yyyy-mm-dd");
    // 4. Number Format 0.00 (Cols 2-6)
    dataSheet.getRange(3, colStart + 1, 1500, 5).setNumberFormat("0.00");
  });
}

/**
 * ==========================================
 * 3. CALCULATIONS ENGINE (Styled Table)
 * ==========================================
 */
function generateCalculationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("DATA");
  const inputSheet = ss.getSheetByName("INPUT");
  if (!dataSheet) return;

  const tickers = getCleanTickers(inputSheet);
  let calcSheet = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");
  
  calcSheet.clear();
  calcSheet.setFrozenRows(2);
  calcSheet.setFrozenColumns(1);

  // Headers
  calcSheet.getRange("A1").setValue("ASSET").setFontWeight("bold").setHorizontalAlignment("left");
  calcSheet.getRange("B1:E1").merge().setValue("STRATEGY DASHBOARD").setBackground("#E1F5FE").setFontWeight("bold").setHorizontalAlignment("left");

  const headers = [["Ticker", "DECISION", "Price", "Change %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance"]];
  calcSheet.getRange(2, 1, 1, 16).setValues(headers).setFontWeight("bold").setBackground("#222").setFontColor("#FFF").setHorizontalAlignment("left");

  const formulas = [];
  tickers.forEach((ticker, i) => {
    const row = i + 3;
    const colIdx = (i * 7) + 1;
    const closeCol = columnToLetter(colIdx + 4); 
    const highCol = columnToLetter(colIdx + 2); 
    const lowCol = columnToLetter(colIdx + 3); 
    const volCol = columnToLetter(colIdx + 5);
    const count = `COUNTA(DATA!$${closeCol}:$${closeCol})`;
    
    formulas.push([
      ticker, 
      `=IF(C${row}<N${row}, "⚠️ EXIT", IF(C${row}>=O${row}, "💰 TAKE PROFIT", IF(AND(L${row}<45, M${row}="🐂 BULL DIV"), "🎯 BUY DIP", "Wait")))`, 
      `=IFERROR(OFFSET(DATA!$${closeCol}$1, ${count}-1, 0), 0)`, 
      `=IFERROR(GOOGLEFINANCE("${ticker}", "changepct")/100, 0)`,
      `=IF((O${row}-C${row})/(C${row}-N${row}) >= 3, "💎 HIGH", "⚖️ MED")`,
      `=REPT("★", (C${row}>AVERAGE(OFFSET(DATA!$${closeCol}$1, ${count}-20, 0, 20))) + (C${row}>AVERAGE(OFFSET(DATA!$${closeCol}$1, ${count}-50, 0, 50))) + (C${row}>AVERAGE(OFFSET(DATA!$${closeCol}$1, MAX(1, ${count}-200), 0, 200))))`,
      `=IF(LEN(F${row})=3, "🚀 BULLISH", IF(LEN(F${row})=0, "📉 BEARISH", "⚖️ NEUTRAL"))`,
      `=AVERAGE(OFFSET(DATA!$${closeCol}$1, ${count}-20, 0, 20))`,
      `=AVERAGE(OFFSET(DATA!$${closeCol}$1, ${count}-50, 0, 50))`,
      `=AVERAGE(OFFSET(DATA!$${closeCol}$1, MAX(1, ${count}-200), 0, 200))`,
      `=OFFSET(DATA!$${volCol}$1, ${count}-1, 0)/AVERAGE(OFFSET(DATA!$${volCol}$1, ${count}-21, 0, 20))`,
      `=IFERROR(100-(100/(1+(MAX(0,AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$1, ${count}-14, 0, 14)-OFFSET(DATA!$${closeCol}$1, ${count}-15, 0, 14)),">0"))/MAX(0.0001, ABS(AVERAGEIF(ARRAYFORMULA(OFFSET(DATA!$${closeCol}$1, ${count}-14, 0, 14)-OFFSET(DATA!$${closeCol}$1, ${count}-15, 0, 14)),"<0")))))), 50)`,
      `=IF(AND(C${row} < OFFSET(DATA!$${closeCol}$1, ${count}-6, 0)*1.01, L${row} > OFFSET(DATA!$${closeCol}$1, ${count}-7, 0)), "🐂 BULL DIV", "-")`,
      `=MIN(OFFSET(DATA!$${lowCol}$1, ${count}-20, 0, 20))`,
      `=C${row} + ((C${row}-N${row})*3)`,
      `=MAX(OFFSET(DATA!$${highCol}$1, ${count}-50, 0, 50))`
    ]);
  });
  
  const numRows = formulas.length;
  // Paste Formulas
  calcSheet.getRange(3, 1, numRows, 16).setValues(formulas);

  // --- STYLING BLOCK ---
  const fullTable = calcSheet.getRange(2, 1, numRows + 1, 16);
  
  // 1. Borders
  fullTable.setBorder(true, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID);
  // 2. Left Align All
  fullTable.setHorizontalAlignment("left");
  
  // 3. Number Format 0.00 (Prices, SMAs, Targets)
  // Cols C (Price) to P (Resistance) = Index 3 to 16.
  // We apply 0.00 to almost everything except the Decision Text columns
  calcSheet.getRange(3, 3, numRows, 14).setNumberFormat("0.00");
  
  // 4. Percentage Override (Column D)
  calcSheet.getRange(3, 4, numRows, 1).setNumberFormat("0.00%");
}

/**
 * ==========================================
 * 4. CHART SHEET SETUP (Styled Table)
 * ==========================================
 */
function setupChartSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  const tickers = getCleanTickers(inputSheet);
  let chartSheet = ss.getSheetByName("CHART") || ss.insertSheet("CHART");
  
  chartSheet.clear(); 

  // --- CONTROLS ---
  chartSheet.getRange("A1").setValue("TICKER:").setFontWeight("bold").setHorizontalAlignment("left");
  chartSheet.getRange("B1").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tickers).build()).setValue(tickers[0]).setBackground("#FFF9C4").setHorizontalAlignment("left");
  
  chartSheet.getRange("D1").setValue("VIEW:").setFontWeight("bold").setHorizontalAlignment("left");
  chartSheet.getRange("D2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"]).build()).setValue("DAILY").setBackground("#E1F5FE").setHorizontalAlignment("left");

  // Date Dropdowns
  const numberList = [0,1,2,3,4,5,6,7,8,9,10,11,12];
  const dropdownRule = SpreadsheetApp.newDataValidation().requireValueInList(numberList).build();

  chartSheet.getRange("A2:C2").setValues([["Years", "Months", "Days"]]).setBackground("#222").setFontColor("#FFF").setHorizontalAlignment("left");
  chartSheet.getRange("A3:C3").setDataValidation(dropdownRule).setHorizontalAlignment("left");
  chartSheet.getRange("A3").setValue(0);
  chartSheet.getRange("B3").setValue(3);
  chartSheet.getRange("C3").setValue(0);
  
  chartSheet.getRange("A4").setValue("START:").setFontWeight("bold").setHorizontalAlignment("left");
  chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)").setNumberFormat("yyyy-mm-dd").setHorizontalAlignment("left");

  // --- INFO TABLE STYLING ---
  const infoTable = chartSheet.getRange("A6:B7");
  
  chartSheet.getRange("A6:A7").setValues([["52W HIGH"],["52W LOW"]]).setFontWeight("bold").setBackground("#EEE");
  chartSheet.getRange("B6:B7").setFormulas([
    [`=IFERROR(GOOGLEFINANCE(B1, "high52"), 0)`],
    [`=IFERROR(GOOGLEFINANCE(B1, "low52"), 0)`]
  ]);
  
  // 1. Borders
  infoTable.setBorder(true, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID);
  // 2. Left Align
  infoTable.setHorizontalAlignment("left");
  // 3. Number Format
  chartSheet.getRange("B6:B7").setNumberFormat("0.00");

  updateDynamicChart();
}

/**
 * ==========================================
 * 5. CHART ENGINE (Auto-Scale + Detailed Date)
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

  let high52 = sheet.getRange("B6").getValue();
  let low52 = sheet.getRange("B7").getValue();
  if (!high52 || high52 === 0) high52 = null;
  if (!low52 || low52 === 0) low52 = null;

  const lastCol = dataSheet.getLastColumn();
  const lastRow = dataSheet.getLastRow();
  const rawData = dataSheet.getRange(1, 1, Math.min(2000, lastRow), lastCol).getValues();

  const headers = rawData[0];
  const colIndex = headers.indexOf(ticker);
  
  if (colIndex === -1) {
    sheet.getRange("D6").setValue("Ticker not found.");
    return;
  }

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

    if (close < minVal) minVal = close;
    if (close > maxVal) maxVal = close;

    let dateStr = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    masterData.push([dateStr, close, high52, low52]);
  }

  if (low52 !== null) minVal = Math.min(minVal, low52);
  if (high52 !== null) maxVal = Math.max(maxVal, high52);
  const yMin = minVal * 0.98; 
  const yMax = maxVal * 1.02;

  sheet.getRange("Z3:AC").clearContent();

  if (masterData.length === 0) {
    sheet.getRange("D6").setValue("No data found.");
    return;
  }

  sheet.getRange(3, 26, masterData.length, 4).setValues(masterData);

  const charts = sheet.getCharts();
  charts.forEach(c => sheet.removeChart(c));
  const chartRange = sheet.getRange(3, 26, masterData.length, 4); 

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartRange)
    .setOption('series', {
      0: {labelInLegend: 'Price', color: '#1976D2', lineWidth: 3}, 
      1: {labelInLegend: '52W High', color: '#F57C00', lineWidth: 1.5},
      2: {labelInLegend: '52W Low', color: '#7B1FA2', lineWidth: 1.5}
    })
    .setOption('curveType', 'function')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('vAxis', {
      title: 'Price ($)',
      gridlines: {count: 5, color: '#e0e0e0'},
      viewWindowMode: 'explicit',
      viewWindow: { min: yMin, max: yMax }
    })
    .setOption('hAxis', {
      type: 'category', 
      slantedText: true, 
      textStyle: {fontSize: 10},
      maxAlternation: 1
    })
    .setOption('chartArea', {left: '8%', top: '10%', width: '85%', height: '70%'})
    .setOption('title', ticker + ' Institutional Analysis')
    .setOption('legend', {position: 'top'})
    .setPosition(4, 3, 0, 0)
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