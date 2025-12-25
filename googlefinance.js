/**
* ==============================================================================
* BASELINE LABEL: STABLE_MASTER_DEC25_BASE_v1
* DATE: 25 DEC 2025
* FIX: Re-integrated Indicator Names (BB, ATR) into the reasoning text.
* Logic: "Indicator (Value): Meaning".
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
   .addItem('🔔 Start Market Monitor', 'startMarketMonitor') // New
   .addItem('🔕 Stop Monitor', 'stopMarketMonitor')       // New
   .addItem('📩 Test Alert Now', 'checkSignalsAndSendAlerts')
   .addToUi();
}


/**
* ------------------------------------------------------------------
* 1. CUSTOM MATH FUNCTIONS
* ------------------------------------------------------------------
*/


function LIVERSI(history, currentPrice) {
 if (!history || !currentPrice) return 50;
 let raw = history.flat();
 let closes = [];
 for (let i = raw.length - 1; i >= 0; i--) {
   if (typeof raw[i] === 'number' && raw[i] > 0) closes.unshift(raw[i]);
   if (closes.length >= 30) break;
 }
 closes.push(currentPrice);
 if (closes.length < 15) return 50;


 let gains = 0, losses = 0;
 for (let i = 1; i <= 14; i++) {
   let change = closes[i] - closes[i - 1];
   if (change > 0) gains += change; else losses += Math.abs(change);
 }
 let avgGain = gains / 14, avgLoss = losses / 14;
 for (let i = 15; i < closes.length; i++) {
   let change = closes[i] - closes[i - 1];
   let gain = (change > 0) ? change : 0;
   let loss = (change < 0) ? Math.abs(change) : 0;
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
   if (closes.length >= 150) break;
 }
 closes.push(currentPrice);
 if (closes.length < 26) return 0;


 function calculateEMA(data, period) {
   let k = 2 / (period + 1);
   let ema = data[0];
   let emaArray = [ema];
   for (let i = 1; i < data.length; i++) {
     ema = data[i] * k + ema * (1 - k);
     emaArray.push(ema);
   }
   return emaArray;
 }


 const ema12 = calculateEMA(closes, 12);
 const ema26 = calculateEMA(closes, 26);
 let macdLine = [];
 for (let i = 0; i < closes.length; i++) {
   macdLine.push(ema12[i] - ema26[i]);
 }
 const signalLineArr = calculateEMA(macdLine, 9);
 return Number((macdLine[macdLine.length - 1] - signalLineArr[signalLineArr.length - 1]).toFixed(3));
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
 if (ui.alert('🚨 Full Rebuild', 'Activate V111 Hybrid Text?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;


 sheetsToDelete.forEach(name => {
   let sheet = ss.getSheetByName(name);
   if (sheet) ss.deleteSheet(sheet);
 });


 ui.showModelessDialog(HtmlService.createHtmlOutput("<b>1/4:</b> Syncing Global Data..."), "Status");
 generateDataSheet();
 SpreadsheetApp.flush();
 Utilities.sleep(2000);


 ui.showModelessDialog(HtmlService.createHtmlOutput("<b>2/4:</b> Integrating Indicators..."), "Status");
 generateCalculationsSheet();
 SpreadsheetApp.flush();


 ui.showModelessDialog(HtmlService.createHtmlOutput("<b>3/4:</b> Building Dashboard..."), "Status");
 generateDashboardSheet();
 SpreadsheetApp.flush();


 ui.showModelessDialog(HtmlService.createHtmlOutput("<b>4/4:</b> Constructing Chart..."), "Status");
 setupChartSheet();
 ui.alert('✅ Rebuild Complete', 'V111 Online. Data links restored.', ui.ButtonSet.OK);
}


function onEdit(e) {
 const sheet = e.range.getSheet();
 if (sheet.getName() === "INPUT" && (e.range.getA1Notation() === "B1" || e.range.getA1Notation() === "C1")) {
   generateDashboardSheet();
 }
 if (sheet.getName() === "CHART") {
   const watchList = ["B1", "D2", "A3", "B3", "C3"];
   if (watchList.includes(e.range.getA1Notation()) || (e.range.getRow() === 1 && e.range.getColumn() <= 4)) {
     updateDynamicChart();
   }
 }
}


/**
* ------------------------------------------------------------------
* 3. CHART ENGINE
* ------------------------------------------------------------------
*/


function updateDynamicChart() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheet = ss.getSheetByName("CHART");
 const dataSheet = ss.getSheetByName("DATA");
 if (!sheet || !dataSheet) return;


 SpreadsheetApp.flush();


 sheet.getRange("E5").setValue("Updated: " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "HH:mm:ss")).setFontColor("gray").setFontSize(8).setHorizontalAlignment("right");


 const ticker = sheet.getRange("B1").getValue();
 const isWeekly = sheet.getRange("D2").getValue() === "WEEKLY";
  const years = sheet.getRange("A3").getValue() || 0;
 const months = sheet.getRange("B3").getValue() || 0;
 const days = sheet.getRange("C3").getValue() || 0;
 const now = new Date();
 let startDate = new Date(now.getFullYear() - years, now.getMonth() - months, now.getDate() - days);
  if ((now - startDate) < (7 * 24 * 60 * 60 * 1000)) {
   startDate = new Date();
   startDate.setDate(now.getDate() - 14);
 }


 const supportVal = Number(sheet.getRange("B32").getValue()) || 0;
 const resistanceVal = Number(sheet.getRange("B33").getValue()) || 0;
 let livePrice = Number(sheet.getRange("B8").getValue()) || 0;


 const rawHeaders = dataSheet.getRange(2, 1, 1, dataSheet.getLastColumn()).getValues()[0];
 const colIdx = rawHeaders.indexOf(ticker);
  if (colIdx === -1) {
   sheet.getRange("E1").setValue("⚠️ Ticker Not Found");
   return;
 }


 const rawData = dataSheet.getRange(1, colIdx + 1, dataSheet.getLastRow(), 6).getValues();
 let masterData = [], viewVols = [], prices = [];


 for (let i = 4; i < rawData.length; i++) {
   let row = rawData[i], d = row[0], close = Number(row[4]), vol = Number(row[5]);
   if (!d || !(d instanceof Date) || isNaN(close) || close < 0.01) continue;
   if (d < startDate) continue;
   if (isWeekly && d.getDay() !== 5) continue;


   let slice = rawData.slice(Math.max(4, i-200), i+1).map(r => r[4]);
   let s20 = slice.length >= 20 ? Number((slice.slice(-20).reduce((a,b)=>a+b,0)/20).toFixed(2)) : null;
   let s50 = slice.length >= 50 ? Number((slice.slice(-50).reduce((a,b)=>a+b,0)/50).toFixed(2)) : null;
   let s200 = slice.length >= 200 ? Number((slice.slice(-200).reduce((a,b)=>a+b,0)/200).toFixed(2)) : null;


   masterData.push([
     Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "MMM dd"),
     close,
     (close >= (i>4?rawData[i-1][4]:close)) ? vol : null,
     (close < (i>4?rawData[i-1][4]:close)) ? vol : null, 
     s20, s50, s200, resistanceVal, supportVal
   ]);


   viewVols.push(vol); prices.push(close);
   if(s20) prices.push(s20); if(s50) prices.push(s50); if(s200) prices.push(s200);
 }


 let candleLabel = "🔴 LIVE";
 if (livePrice === 0 && prices.length > 0) {
   livePrice = prices[prices.length - 1];
   candleLabel = "⏳ SYNCING";
 }


 if (livePrice > 0) {
   let allCloses = rawData.slice(4).map(r => Number(r[4])).filter(n => n > 0);
   let sma20Arr = allCloses.slice(-19); sma20Arr.push(livePrice);
   let sma50Arr = allCloses.slice(-49); sma50Arr.push(livePrice);
   let sma200Arr = allCloses.slice(-199); sma200Arr.push(livePrice);
   let liveS20 = sma20Arr.length >= 20 ? sma20Arr.reduce((a,b)=>a+b,0)/20 : null;
   let liveS50 = sma50Arr.length >= 50 ? sma50Arr.reduce((a,b)=>a+b,0)/50 : null;
   let liveS200 = sma200Arr.length >= 200 ? sma200Arr.reduce((a,b)=>a+b,0)/200 : null;
   masterData.push([candleLabel, livePrice, null, null, liveS20, liveS50, liveS200, resistanceVal, supportVal]);
   prices.push(livePrice);
 }


 sheet.getRange(3, 26, 2000, 9).clearContent();
 if (masterData.length === 0) return;
 if (supportVal > 0) prices.push(supportVal);
 if (resistanceVal > 0) prices.push(resistanceVal);


 const minP = Math.min(...prices.filter(p => p > 0)) * 0.98;
 const maxP = Math.max(...prices.filter(p => p > 0)) * 1.02;
 const maxVol = Math.max(...viewVols);


 sheet.getRange(2, 26, 1, 9).setValues([["Date", "Price", "Bull Vol", "Bear Vol", "SMA 20", "SMA 50", "SMA 200", "Resistance", "Support"]]).setFontWeight("bold").setFontColor("white");
 sheet.getRange(3, 26, masterData.length, 9).setValues(masterData);
  SpreadsheetApp.flush();
 sheet.getCharts().forEach(c => sheet.removeChart(c));
  const chart = sheet.newChart().setChartType(Charts.ChartType.COMBO).addRange(sheet.getRange(2, 26, masterData.length + 1, 9)).setOption('useFirstRowAsHeaders', true).setOption('series', {0:{type:'line',color:'#1A73E8',lineWidth:3,labelInLegend:'Price'},1:{type:'bars',color:'#2E7D32',targetAxisIndex:1,labelInLegend:'Bull Vol'},2:{type:'bars',color:'#C62828',targetAxisIndex:1,labelInLegend:'Bear Vol'},3:{type:'line',color:'#FBBC04',lineWidth:1.5,labelInLegend:'SMA 20'},4:{type:'line',color:'#9C27B0',lineWidth:1.5,labelInLegend:'SMA 50'},5:{type:'line',color:'#FF9800',lineWidth:2,labelInLegend:'SMA 200'},6:{type:'line',color:'#B71C1C',lineDashStyle:[4,4],labelInLegend:'Resistance'},7:{type:'line',color:'#0D47A1',lineDashStyle:[4,4],labelInLegend:'Support'}}).setOption('vAxes', {0:{viewWindow:{min:minP,max:maxP}},1:{viewWindow:{min:0,max:maxVol * 4},textStyle:{color:'#666'},format:'short'}}).setOption('legend', {position: 'top', textStyle: {fontSize: 10}}).setPosition(5, 3, 0, 0).setOption('width', 1150).setOption('height', 650).build();
 sheet.insertChart(chart);
}


/**
* ------------------------------------------------------------------
* 4. DATA ENGINE
* ------------------------------------------------------------------
*/


function generateDataSheet() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const inputSheet = ss.getSheetByName("INPUT");
 if (!inputSheet) return;
 const tickers = getCleanTickers(inputSheet);
 let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");
 dataSheet.clear().clearFormats();
 dataSheet.getRange("A1").setValue("Last Update: " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm")).setFontWeight("bold").setFontColor("blue");
 tickers.forEach((ticker, i) => {
   const colStart = (i * 7) + 1;
   dataSheet.getRange(2, colStart).setNumberFormat("@").setValue(ticker).setFontWeight("bold");
   dataSheet.getRange(3, colStart).setValue("ATH:");
   dataSheet.getRange(3, colStart + 1).setFormula(`=MAX(QUERY(GOOGLEFINANCE("${ticker}", "high", "1/1/2000", TODAY()), "SELECT Col2 LABEL Col2 ''"))`);
   dataSheet.getRange(4, colStart).setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "all", TODAY()-800, TODAY()), "No Data")`);
   dataSheet.getRange(5, colStart, 1000, 1).setNumberFormat("yyyy-mm-dd");
   dataSheet.getRange(5, colStart + 1, 1000, 5).setNumberFormat("#,##0.00");
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


/**
* ------------------------------------------------------------------
* 5. CALCULATION ENGINE (REVISED HEADER & FREEZE)
* ------------------------------------------------------------------
*/

function generateCalculationsSheet() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const dataSheet = ss.getSheetByName("DATA");
 const inputSheet = ss.getSheetByName("INPUT");
 if (!dataSheet) return;
 const tickers = getCleanTickers(inputSheet);
 let calcSheet = ss.getSheetByName("CALCULATIONS") || ss.insertSheet("CALCULATIONS");
 calcSheet.clear().clearFormats();

 // ID Header (Independent A) and Core Ident (B-F)
 calcSheet.getRange("A1").setValue("ID").setBackground("#212121").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
 calcSheet.getRange("B1:F1").merge().setValue("[ CORE IDENT ]").setBackground("#263238").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
 calcSheet.getRange("G1:I1").merge().setValue("[ PERFORMANCE ]").setBackground("#0D47A1").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
 calcSheet.getRange("J1:R1").merge().setValue("[ MOMENTUM ]").setBackground("#1B5E20").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
 calcSheet.getRange("S1:W1").merge().setValue("[ RISK LEVELS ]").setBackground("#B71C1C").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
 calcSheet.getRange("X1:Y1").merge().setValue("[ SPLIT ANALYST ]").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");
 calcSheet.getRange("Z1").setValue("MONITOR").setBackground("#212121").setFontColor("white").setHorizontalAlignment("center").setFontWeight("bold");

 const headers = [["Ticker", "Price", "Change %", "FUNDAMENTAL", "SIGNAL", "DECISION", "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "MACD Hist", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "TECH_REASON", "FUND_REASON", "LAST_STATE"]];
 const headerRange = calcSheet.getRange(2, 1, 1, 26);
 headerRange.setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold");
  
 const formulas = [];
 tickers.forEach((ticker, i) => {
   const rowNum = i + 3, tickerDataStart = (i * 7) + 1, closeCol = columnToLetter(tickerDataStart + 4), lastRow = `COUNTA(DATA!$${closeCol}:$${closeCol})`;
   formulas.push([
     `=ROUND(IFERROR(GOOGLEFINANCE("${ticker}", "price")), 2)`,
     `=IFERROR(GOOGLEFINANCE("${ticker}", "changepct")/100, 0)`,
     `=IFERROR(LET(eps, GOOGLEFINANCE(A${rowNum}, "eps"), pe, GOOGLEFINANCE(A${rowNum}, "pe"), IFS(AND(eps>0, pe>0, pe<25), "💎 GEM (Value)", AND(eps>0, pe>50), "⚠️ PRICED PERF.", eps<0, "💀 ZOMBIE", AND(pe>30, eps<0.1), "🔴 BUBBLE", TRUE, "⚖️ FAIR")), "-")`,
     `=IFERROR(IFS(B${rowNum} < S${rowNum}, "STOP LOSS", B${rowNum} >= U${rowNum}, "RESISTANCE REJECT", AND(B${rowNum} <= S${rowNum}*1.02, P${rowNum} < 35), "BOUNCE PLAY", AND(W${rowNum} > 1.0, O${rowNum} > 1.5), "VOLATILITY SQUEEZE", B${rowNum} < N${rowNum}, "AVOID (BEAR)", AND(B${rowNum} > N${rowNum}, O${rowNum} > 1.2, B${rowNum} > T${rowNum}, Q${rowNum} > 0), "STRONG BREAKOUT", AND(B${rowNum} > N${rowNum}, W${rowNum} < 0.15, P${rowNum} < 35, Q${rowNum} > 0), "STRONG BUY (OVERSOLD)", AND(B${rowNum} > N${rowNum}, R${rowNum} = "BULLISH DIV", Q${rowNum} > -0.5), "ACCUMULATE (DIV)", AND(B${rowNum} > N${rowNum}, B${rowNum} > L${rowNum}, B${rowNum} < T${rowNum}, Q${rowNum} > 0), "RIDE TREND", TRUE, "WAIT (CHOP)"), "Wait")`,
     `=IFS(REGEXMATCH(E${rowNum}, "STOP"), "🛑 STOP OUT", REGEXMATCH(E${rowNum}, "REJECT|TAKE"), "💰 CASH OUT", AND(REGEXMATCH(E${rowNum}, "STRONG|ACCUMULATE|BREAKOUT|BOUNCE|SQUEEZE"), D${rowNum}="💎 GEM (Value)"), "💎 PRIME BUY", AND(REGEXMATCH(E${rowNum}, "STRONG|ACCUMULATE|BREAKOUT|BOUNCE|SQUEEZE"), D${rowNum}="💀 ZOMBIE"), "⚠️ GAMBLE", REGEXMATCH(E${rowNum}, "STRONG|ACCUMULATE|BREAKOUT|BOUNCE|SQUEEZE"), "🚀 TRADE", E${rowNum}="RIDE TREND", "🌊 HOLD (RIDE)", E${rowNum}="AVOID (BEAR)", "💤 AVOID", TRUE, "⏳ WAIT")`,
     `=IFERROR(DATA!${columnToLetter(tickerDataStart + 1)}3, "-")`,
     `=IFERROR((B${rowNum}-G${rowNum})/G${rowNum}, 0)`,
     `=IFERROR(ROUND((U${rowNum}-B${rowNum})/MAX(0.01, B${rowNum}-S${rowNum}), 2), 0)`,
     `=REPT("★", (B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20))) + (B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-50, 0, 50))))`,
     `=IF(B${rowNum}>AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-200, 0, 200)), "BULL REGIME", "BEAR REGIME")`,
     `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)), 0), 2)`,
     `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-50, 0, 50)), 0), 2)`,
     `=ROUND(IFERROR(AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-200, 0, 200)), 0), 2)`,
     `=ROUND(IFERROR(OFFSET(DATA!$${columnToLetter(tickerDataStart+5)}$4, ${lastRow}-1, 0) / AVERAGE(OFFSET(DATA!$${columnToLetter(tickerDataStart+5)}$4, ${lastRow}-21, 0, 20)), 1), 2)`,
     `=LIVERSI(DATA!$${closeCol}$4:$${closeCol}, B${rowNum})`,
     `=LIVEMACD(DATA!$${closeCol}$4:$${closeCol}, B${rowNum})`,
     `=IFERROR(IFS(AND(B${rowNum} < INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), P${rowNum} > 50), "BULLISH DIV", AND(B${rowNum} > INDEX(DATA!$${closeCol}:$${closeCol}, ${lastRow}-14), P${rowNum} < 50), "BEARISH DIV", TRUE, "-"), "-")`,
     `=ROUND(IFERROR(MIN(OFFSET(DATA!$${columnToLetter(tickerDataStart+3)}$4, ${lastRow}-21, 0, 20)), B${rowNum}*0.9), 2)`,
     `=ROUND(B${rowNum} + ((B${rowNum}-S${rowNum}) * 3), 2)`,
     `=ROUND(IFERROR(MAX(OFFSET(DATA!$${columnToLetter(tickerDataStart+2)}$4, ${lastRow}-51, 0, 50)), B${rowNum}*1.1), 2)`,
     `=ROUND(IFERROR(AVERAGE(ARRAYFORMULA(OFFSET(DATA!$${columnToLetter(tickerDataStart+2)}$4, ${lastRow}-14, 0, 14)-OFFSET(DATA!$${columnToLetter(tickerDataStart+3)}$4, ${lastRow}-14, 0, 14))), 0), 2)`,
     `=ROUND(IFERROR(((B${rowNum}-AVERAGE(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20))) / (4*STDEV(OFFSET(DATA!$${closeCol}$4, ${lastRow}-20, 0, 20)))) + 0.5, 0.5), 2)`,
    
     // X: TECH REASONING
     `="1. TREND: " & K${rowNum} & ". Price ($" & B${rowNum} & ") is " & IF(B${rowNum}>N${rowNum}, "stable above", "weak below") & " 200SMA anchor." & CHAR(10) &
       "2. MOMENTUM: RSI " & P${rowNum} & " (" & IF(AND(P${rowNum}>40,P${rowNum}<60),"Stable","Action") & ") + MACD " & Q${rowNum} & "." & CHAR(10) &
       "3. STRUCTURE: " & IFS(AND(W${rowNum}>0.2, W${rowNum}<0.8), "Mid-Channel (No Edge).", W${rowNum}<=0.2, "Support Floor.", TRUE, "Breakout/Resistance.")`,
    
     // Y: FUND REASONING
     `="1. VALUATION: " & D${rowNum} & ". P/E " & IFERROR(GOOGLEFINANCE(A${rowNum}, "pe"),"-") & "." & CHAR(10) &
       "2. VERDICT: " & F${rowNum} & ". Signal Turn Monitor is " & IF(Z${rowNum}="","Pending","Watching " & Z${rowNum})`
   ]);
 });
 calcSheet.getRange(3, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
 calcSheet.getRange(3, 2, formulas.length, 24).setFormulas(formulas);
 SpreadsheetApp.flush();

 // UI & BORDERS
 calcSheet.setFrozenRows(2);
 calcSheet.setFrozenColumns(1);
 for (let col = 1; col <= 23; col++) calcSheet.setColumnWidth(col, 70);
 calcSheet.setColumnWidth(24, 500); calcSheet.setColumnWidth(25, 500); calcSheet.setColumnWidth(26, 100);
 calcSheet.getRange("A:W").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
 calcSheet.getRange("X:Y").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
 
 const lastRowVal = calcSheet.getLastRow();
 const fullRange = calcSheet.getRange(1, 1, lastRowVal, 26);
 fullRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
 headerRange.setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}


/**
* ------------------------------------------------------------------
* 6. DASHBOARD ENGINE (TOTAL COLOR-CODING + FUNDAMENTAL LOGIC)
* ------------------------------------------------------------------
*/

function generateDashboardSheet() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const inputSheet = ss.getSheetByName("INPUT");
 if (!inputSheet) return;
 let dashboard = ss.getSheetByName("DASHBOARD") || ss.insertSheet("DASHBOARD", inputSheet.getIndex() + 1);
 dashboard.clear().clearFormats();
 
 const headers = [["Ticker", "Price", "Change %", "FUNDAMENTAL", "SIGNAL", "DECISION", "ATH (TRUE)", "ATH Diff %", "R:R Quality", "Trend Score", "Trend State", "SMA 20", "SMA 50", "SMA 200", "Vol Trend", "RSI", "MACD Hist", "Divergence", "Support", "Target (3:1)", "Resistance", "ATR (14)", "Bollinger %B", "TECH ANALYSIS", "FUND ANALYSIS"]];
 const headerRange = dashboard.getRange(2, 1, 1, 25);
 headerRange.setValues(headers).setBackground("#212121").setFontColor("white").setFontWeight("bold");

 const formula = '=IFERROR(SORT(FILTER(CALCULATIONS!$A$3:$Y, ISNUMBER(MATCH(CALCULATIONS!$A$3:$A, FILTER(INPUT!$A$3:$A, (IF(OR(INPUT!$B$1="", INPUT!$B$1="ALL"), 1, REGEXMATCH(INPUT!$B$3:$B, "(?i)\\b(" & SUBSTITUTE(SUBSTITUTE(INPUT!$B$1, ", ", "|"), ",", "|") & ")\\b"))) * (IF(OR(INPUT!$C$1="", INPUT!$C$1="ALL"), 1, REGEXMATCH(INPUT!$C$3:$C, "(?i)\\b(" & SUBSTITUTE(SUBSTITUTE(INPUT!$C$1, ", ", "|"), ",", "|") & ")\\b")))), 0))), 3, FALSE), "No Matches Found")';
 dashboard.getRange("A3").setFormula(formula);

 SpreadsheetApp.flush(); 
 const lastRow = Math.max(dashboard.getLastRow(), 3);

 // LAYOUT
 dashboard.setFrozenRows(2);
 dashboard.setFrozenColumns(1);
 for (let col = 1; col <= 23; col++) dashboard.setColumnWidth(col, 70);
 dashboard.setColumnWidth(24, 500); dashboard.setColumnWidth(25, 500);
 dashboard.getRange("A:W").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
 dashboard.getRange("X:Y").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

 // COLOR RULES (Contribution Based)
 const rules = [];
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B3<$N3').setFontColor("#C62828").setBold(true).setRanges([dashboard.getRange("B3:B"), dashboard.getRange("N3:N")]).build()); // Price vs SMA 200
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$M3<$N3').setFontColor("#C62828").setRanges([dashboard.getRange("M3:M")]).build()); // Death Cross
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(70).setFontColor("#C62828").setRanges([dashboard.getRange("P3:P")]).build()); // RSI Overbought
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor("#2E7D32").setBold(true).setRanges([dashboard.getRange("H3:H")]).build()); // ATH
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=REGEXMATCH($D3, "(?i)ZOMBIE|BUBBLE|PRICED")').setFontColor("#C62828").setBold(true).setRanges([dashboard.getRange("D3:D")]).build()); // Fund Weakness
 
 // Background Heatmap
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=REGEXMATCH($E3&$F3, "(?i)PRIME|TRADE|BUY")').setBackground("#E8F5E9").setFontColor("#2E7D32").setBold(true).setRanges([dashboard.getRange("E3:F")]).build());
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=REGEXMATCH($E3&$F3, "(?i)STOP|REJECT|AVOID")').setBackground("#FFEBEE").setFontColor("#C62828").setBold(true).setRanges([dashboard.getRange("E3:F")]).build());
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=REGEXMATCH($E3&$F3, "(?i)WAIT|CHOP")').setBackground("#F5F5F5").setFontColor("#616161").setRanges([dashboard.getRange("E3:F")]).build());

 dashboard.setConditionalFormatRules(rules);
 dashboard.getRangeList(['C3:C', 'H3:H', 'W3:W']).setNumberFormat("0.00%");
}

/**
* ------------------------------------------------------------------
* 7. SETUP CHART SHEET (UI DUAL BOX SYMMETRY + COLORS)
* ------------------------------------------------------------------
*/


function setupChartSheet() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const inputSheet = ss.getSheetByName("INPUT");
 const tickers = getCleanTickers(inputSheet);
 let chartSheet = ss.getSheetByName("CHART") || ss.insertSheet("CHART");
 chartSheet.clear().clearFormats();
 forceExpandSheet(chartSheet, 45);
  chartSheet.setColumnWidth(1, 180); chartSheet.setColumnWidth(2, 120);
 chartSheet.setColumnWidth(5, 125); chartSheet.setColumnWidth(6, 125);
 chartSheet.setColumnWidth(7, 125); chartSheet.setColumnWidth(8, 125);
  const headerRange = chartSheet.getRange("A1:H4");
 headerRange.setBackground("#000000").setFontColor("#FFFF00")
            .setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);
  chartSheet.getRange("A1").setValue("TICKER:").setFontWeight("bold");
 chartSheet.getRange("B1:D1").merge().setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tickers).build()).setValue(tickers[0]).setFontWeight("bold").setHorizontalAlignment("center").setFontSize(12).setFontColor("#FF80AB");


 // DUAL BOX SPLIT (SYMMETRICAL)
 chartSheet.getRange("E1:F4").merge().setWrap(true).setVerticalAlignment("top").setFontSize(10).setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);
 chartSheet.getRange("E1").setFormula('=IFERROR(VLOOKUP(B1, CALCULATIONS!$A$3:$Y, 24, 0), "—")');


 chartSheet.getRange("G1:H4").merge().setWrap(true).setVerticalAlignment("top").setFontSize(10).setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);
 chartSheet.getRange("G1").setFormula('=IFERROR(VLOOKUP(B1, CALCULATIONS!$A$3:$Y, 25, 0), "—")');


 chartSheet.getRange("A2:C2").setValues([["YEAR", "MONTH", "DAY"]]).setFontWeight("bold").setHorizontalAlignment("center");
 const numRule = (max) => SpreadsheetApp.newDataValidation().requireValueInList(Array.from({length: max + 1}, (_, i) => i)).build();
 chartSheet.getRange("A3").setDataValidation(numRule(5)).setValue(1).setHorizontalAlignment("center").setFontColor("#FF80AB");
 chartSheet.getRange("B3").setDataValidation(numRule(12)).setValue(0).setHorizontalAlignment("center").setFontColor("#FF80AB");
 chartSheet.getRange("C3").setDataValidation(numRule(31)).setValue(0).setHorizontalAlignment("center").setFontColor("#FF80AB");
  chartSheet.getRange("D2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["DAILY", "WEEKLY"]).build()).setValue("DAILY").setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#FF80AB");


 chartSheet.getRange("A4").setValue("DATE").setFontWeight("bold");
 chartSheet.getRange("B4").setFormula("=DATE(YEAR(TODAY())-A3, MONTH(TODAY())-B3, DAY(TODAY())-C3)").setNumberFormat("yyyy-mm-dd");


 const t = "B1";
 const data = [
   ["SIGNAL (RAW)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 5, 0), "Wait")`],
   ["FUNDAMENTAL", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 4, 0), "-")`],
   ["DECISION (FINAL)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 6, 0), "-")`],
   ["LIVE PRICE", `=IFERROR(GOOGLEFINANCE(${t}, "price"), 0)`],
   ["CHANGE ($)", `=IFERROR(B8 - GOOGLEFINANCE(${t}, "closeyest"), 0)`],
   ["CHANGE (%)", `=IFERROR(GOOGLEFINANCE(${t}, "changepct")/100, 0)`],
   ["RANGE DIFF %", `=IFERROR((B8 - INDEX(GOOGLEFINANCE(${t}, "close", B4), 2, 2)) / INDEX(GOOGLEFINANCE(${t}, "close", B4), 2, 2), 0)`],
   ["", ""],
   ["[ VALUATION METRICS ]", ""],
   ["ATH (TRUE)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 7, 0), 0)`],
   ["DIFF FROM ATH %", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 8, 0), 0)`],
   ["P/E RATIO", `=IFERROR(GOOGLEFINANCE(${t}, "pe"), 0)`],
   ["EPS", `=IFERROR(GOOGLEFINANCE(${t}, "eps"), 0)`],
   ["52W HIGH", `=IFERROR(GOOGLEFINANCE(${t}, "high52"), 0)`],
   ["52W LOW", `=IFERROR(GOOGLEFINANCE(${t}, "low52"), 0)`], ["", ""],
   ["[ MOMENTUM & TREND ]", ""],
   ["SMA 20", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 12, 0), 0)`],
   ["SMA 50", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 13, 0), 0)`],
   ["SMA 200", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 14, 0), 0)`],
   ["RSI (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 16, 0), 50)`],
   ["MACD HIST", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 17, 0), 0)`],
   ["TREND STATE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 11, 0), "—")`],
   ["DIVERGENCE", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 18, 0), "Neutral")`],
   ["RELATIVE VOLUME", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 15, 0), 1)`], ["", ""],
   ["[ TECHNICAL LEVELS ]", ""],
   ["SUPPORT FLOOR", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 19, 0), 0)`],
   ["RESISTANCE CEILING", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 21, 0), 0)`],
   ["TARGET (3:1 R:R)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 20, 0), 0)`],
   ["ATR (14)", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 22, 0), 0)`],
   ["BOLLINGER %B", `=IFERROR(VLOOKUP(${t}, CALCULATIONS!$A$3:$Y, 23, 0), 0)`]
 ];


 chartSheet.getRange(5, 1, data.length, 1).setValues(data.map(r => [r[0]])).setFontWeight("bold");
 chartSheet.getRange(5, 2, data.length, 1).setFormulas(data.map(r => [r[1]]));
 [13, 22, 31].forEach(r => chartSheet.getRange(r, 1, 1, 2).setBackground("#444").setFontColor("white").setHorizontalAlignment("center"));
 SpreadsheetApp.flush();
 chartSheet.getRange("B5:B36").setHorizontalAlignment("left");
  // Format Currencies
 chartSheet.getRangeList(["B8", "B9", "B14", "B16:B19", "B23:B26", "B29", "B32:B35"]).setNumberFormat("#,##0.00");
  // Format Percentages
 chartSheet.getRangeList(["B10", "B11", "B15", "B36"]).setNumberFormat("0.00%");
  // CHART SIDE PANEL COLOR RULES
 const ruleNegChange = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor("#D32F2F").setRanges([chartSheet.getRange("B9:B10")]).build();
  const ruleBearish = SpreadsheetApp.newConditionalFormatRule().whenTextContains("Negative").setFontColor("#D32F2F").setRanges([chartSheet.getRange("B5:B36")]).build();
  // RSI Colors
 const rsiRed = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(70).setFontColor("#D32F2F").setRanges([chartSheet.getRange("B26")]).build(); // RSI Row
 const rsiGreen = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(30).setFontColor("#388E3C").setRanges([chartSheet.getRange("B26")]).build();


 // MACD Colors
 const macdRed = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor("#D32F2F").setRanges([chartSheet.getRange("B27")]).build(); // MACD Row
 const macdGreen = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor("#388E3C").setRanges([chartSheet.getRange("B27")]).build();


 // P/E Ratio (Expensive)
 const peRed = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(50).setFontColor("#D32F2F").setRanges([chartSheet.getRange("B16")]).build();


 const ruleSMA200 = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=B25>B8').setFontColor("#D32F2F").setRanges([chartSheet.getRange("B25")]).build();
 const ruleSMA50 = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=B24>B8').setFontColor("#D32F2F").setRanges([chartSheet.getRange("B24")]).build();


 chartSheet.setConditionalFormatRules([ruleNegChange, ruleBearish, rsiRed, rsiGreen, macdRed, macdGreen, peRed, ruleSMA200, ruleSMA50]);


 updateDynamicChart();
}

/**
* ------------------------------------------------------------------
* 8. AUTOMATED ALERT & MONITOR SYSTEM
* ------------------------------------------------------------------
*/

/** Starts the background timer to check signals every 30 minutes */
function startMarketMonitor() {
  stopMarketMonitor(); // Clear duplicates
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

  // Range: Ticker (A), Decision (F), Last Alert (Z - Col 26)
  const range = calcSheet.getRange(3, 1, lastRow - 2, 26);
  const data = range.getValues();
  let newAlerts = [];
  let updateRows = [];

  data.forEach((row, index) => {
    const ticker = row[0];
    const currentDecision = row[5];
    const lastNotifiedDecision = row[25]; // Column Z

    // Alert if the decision has CHANGED and is actionable
    if (currentDecision !== lastNotifiedDecision) {
      if (/(PRIME|TRADE|STOP|CASH|BOUNCE|BREAKOUT)/i.test(currentDecision)) {
        newAlerts.push(`🔄 SIGNAL TURN: ${ticker} is now ${currentDecision} (was ${lastNotifiedDecision || 'Unknown'})`);
      }
      // Log the new state in Column Z to prevent duplicate alerts
      calcSheet.getRange(index + 3, 26).setValue(currentDecision);
    }
  });

  if (newAlerts.length > 0) {
    const email = Session.getActiveUser().getEmail();
    MailApp.sendEmail(email, "📈 Terminal Alert: New Signal Turn", 
      "The system detected the following status changes:\n\n" + newAlerts.join("\n") + "\n\nLink: " + ss.getUrl());
  }
}


