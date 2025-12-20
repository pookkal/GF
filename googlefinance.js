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
 * ==========================================
 * 3. CALCULATIONS ENGINE (RESTORED GROUPINGS)
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

  // --- ROW 1: GROUPED HEADERS ---
  calcSheet.getRange("A1").setValue("ASSET").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#EEE");
  
  // Group 1: Strategy (B-E)
  calcSheet.getRange("B1:E1").merge().setValue("STRATEGY DASHBOARD")
    .setBackground("#E1F5FE").setFontColor("#01579B").setFontWeight("bold").setHorizontalAlignment("center");

  // Group 2: Trend (F-J)
  calcSheet.getRange("F1:J1").merge().setValue("TREND ANALYSIS")
    .setBackground("#E8F5E9").setFontColor("#1B5E20").setFontWeight("bold").setHorizontalAlignment("center");

  // Group 3: Technicals (K-P)
  calcSheet.getRange("K1:P1").merge().setValue("TECHNICAL LEVELS")
    .setBackground("#FFF3E0").setFontColor("#E65100").setFontWeight("bold").setHorizontalAlignment("center");

  // --- ROW 2: COLUMN HEADERS ---
  const headers = [["Ticker", "DECISION", "Price", "Change %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "Divergence", "Support", "Target (3:1)", "Resistance"]];
  calcSheet.getRange(2, 1, 1, 16).setValues(headers)
    .setFontWeight("bold").setBackground("#212121").setFontColor("#FFF").setHorizontalAlignment("left");

  // --- ROW 3+: DATA FORMULAS ---
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
      `=IFERROR(IF((P${row}-C${row}) / MAX(0.01, C${row}-N${row}) >= 3, "💎 HIGH", "⚖️ MED"), "-")`,
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
  calcSheet.getRange(3, 1, numRows, 16).setValues(formulas);
  
  // --- STRICT FORMATTING ---
  // Text Columns
  calcSheet.getRange(3, 1, numRows, 1).setNumberFormat("@"); 
  calcSheet.getRange(3, 2, numRows, 1).setNumberFormat("@");
  calcSheet.getRange(3, 5, numRows, 3).setNumberFormat("@");
  calcSheet.getRange(3, 13, numRows, 1).setNumberFormat("@");
  
  // Number Columns (0.00)
  calcSheet.getRange(3, 3, numRows, 1).setNumberFormat("0.00");
  calcSheet.getRange(3, 8, numRows, 5).setNumberFormat("0.00"); 
  calcSheet.getRange(3, 14, numRows, 3).setNumberFormat("0.00");
  
  // Percentage (0.00%)
  calcSheet.getRange(3, 4, numRows, 1).setNumberFormat("0.00%");

  // Borders & Align
  calcSheet.getRange(1, 1, numRows + 2, 16)
    .setBorder(true, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment("left");
    
  // Re-center Row 1 Headers
  calcSheet.getRange(1, 1, 1, 16).setHorizontalAlignment("center");
}

/**
 * ==========================================
 * 4. CHART SHEET SETUP (With Merged Metrics)
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

  // --- EXTENDED INFO TABLE (Merged Code) ---
  const t = "B1"; // Ticker Cell
  const d = "WORKDAY(TODAY(),-1)"; // Date for Prev Close
  
  const labels = [
    ["DECISION"], ["PRICE"], ["CHANGE"], ["CHANGE %"], ["PREV CLOSE"], 
    ["DIFF"], ["DIFF %"], ["52W HIGH"], ["52W LOW"], ["PE RATIO"], 
    ["EPS"], ["BETA"], ["YIELD"], ["R:R QUALITY"], ["TREND SCORE"], 
    ["TREND STATE"], ["SMA 20"], ["SMA 50"], ["SMA 200"], ["VOL TREND"], 
    ["RSI"], ["DIVERGENCE"], ["SUPPORT"], ["TARGET"], ["RESISTANCE"], ["REL VOL"]
  ];

  const formulas = [
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 2, FALSE), "Wait")`], // Decision
    [`=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`], 
    [`=GOOGLEFINANCE(${t}, "change")`], 
    [`=GOOGLEFINANCE(${t}, "changepct")/100`], 
    [`=INDEX(IFERROR(GOOGLEFINANCE(${t}, "price", ${d}), {0,B7}), 2, 2)`], // Prev Close
    [`=B7 - B10`], // Diff (Price - Prev)
    [`=IFERROR(B11 / B10, 0)`], // Diff %
    [`=GOOGLEFINANCE(${t}, "high52")`], 
    [`=GOOGLEFINANCE(${t}, "low52")`], 
    [`=GOOGLEFINANCE(${t}, "pe")`], 
    [`=GOOGLEFINANCE(${t}, "eps")`], 
    [`=GOOGLEFINANCE(${t}, "beta")`], 
    [`=IFERROR(GOOGLEFINANCE(${t}, "yield")/100, 0)`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 5, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 6, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 7, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 8, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 9, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 10, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 11, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 12, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 13, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 14, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 15, FALSE), "—")`], 
    [`=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$P, 16, FALSE), "—")`], 
    [`=IFERROR(GOOGLEFINANCE(${t}, "volume") / AVERAGE(GOOGLEFINANCE(${t}, "volume", TODAY()-30, TODAY())), 0)`]
  ];

  // Set Labels
  chartSheet.getRange(6, 1, labels.length, 1).setValues(labels)
    .setFontWeight("bold").setBackground("#EEE").setHorizontalAlignment("left");
  
  // Set Formulas
  chartSheet.getRange(6, 2, formulas.length, 1).setFormulas(formulas)
    .setHorizontalAlignment("left");

  // Borders & Styling
  chartSheet.getRange(1, 1, 6 + labels.length - 1, 2)
    .setBorder(true, true, true, true, true, true, "#999", SpreadsheetApp.BorderStyle.SOLID);

  // Number Formats
  chartSheet.getRange("B7:B8").setNumberFormat("0.00"); // Price, Change
  chartSheet.getRange("B9").setNumberFormat("0.00%");   // Change %
  chartSheet.getRange("B10:B11").setNumberFormat("0.00"); // Prev, Diff
  chartSheet.getRange("B12").setNumberFormat("0.00%");    // Diff %
  chartSheet.getRange("B13:B17").setNumberFormat("0.00"); // 52W, PE, EPS, Beta
  chartSheet.getRange("B18").setNumberFormat("0.00%");    // Yield
  chartSheet.getRange("B22:B24").setNumberFormat("0.00"); // SMAs
  chartSheet.getRange("B26:B27").setNumberFormat("0.00"); // RSI, Div
  chartSheet.getRange("B28:B30").setNumberFormat("0.00"); // S/T/R Levels
  chartSheet.getRange("B31").setNumberFormat("0.00");     // Vol Rel

  updateDynamicChart();
}

/**
 * ==========================================
 * 5. CHART ENGINE
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

  // Pull 52W from new table positions (Rows 13 and 14 -> Indexes 13, 14 -> B13, B14)
  let high52 = sheet.getRange("B13").getValue();
  let low52 = sheet.getRange("B14").getValue();
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
    .setOption('chartArea', {left: '8%', top: '5%', width: '85%', height: '80%'})
    .setOption('title', ticker + ' Institutional Analysis')
    .setOption('legend', {position: 'top'})
    .setPosition(4, 3, 0, 0)
    .setOption('width', 1100)
    .setOption('height', 600) // Increased height for the longer list
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