/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v3.1_KIRO_OPTIMIZED
* ==============================================================================
*/

function generateDataSheet() {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const inputSheet = ss.getSheetByName("INPUT");
    if (!inputSheet) {
      ss.toast("INPUT sheet not found", "❌ Error", 3);
      return;
    }

    ss.toast("Fetching market data...", "⏳ Loading", 2);
    
    const tickers = getCleanTickers(inputSheet);
    if (tickers.length === 0) {
      ss.toast("No tickers found in INPUT sheet", "⚠️ Warning", 3);
      return;
    }

    let dataSheet = ss.getSheetByName("DATA") || ss.insertSheet("DATA");
    
    // Clear existing data
    dataSheet.clear({ contentsOnly: true });
    dataSheet.clearFormats();

    // Timestamp
    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm");
    dataSheet.getRange("A1")
      .setValue("Last Update: " + timestamp)
      .setFontWeight("bold")
      .setFontColor("blue");

    const colsPer = 7;
    const totalCols = tickers.length * colsPer;
    const regimeStartCol = totalCols + 1;
    const regimeColsNeeded = 6;
    const finalTotalCols = totalCols + regimeColsNeeded;

    // Ensure enough columns
    if (dataSheet.getMaxColumns() < finalTotalCols) {
      dataSheet.insertColumnsAfter(dataSheet.getMaxColumns(), finalTotalCols - dataSheet.getMaxColumns());
    }

    // Get locale separator
    const SEP = (/^(en|en_)/.test(ss.getSpreadsheetLocale())) ? "," : ";";

    // ========================================================================
    // ROW 2: Ticker headers + 52WH/52WL + Market Regime headers
    // ========================================================================
    const row2 = new Array(finalTotalCols).fill("");
    for (let i = 0; i < tickers.length; i++) {
      const b = i * colsPer;
      row2[b] = tickers[i];
      row2[b + 1] = "52WH";
      row2[b + 3] = "52WL";
    }
    row2[regimeStartCol - 1] = "USA_REGIME";
    row2[regimeStartCol] = "USA_RATIO"; 
    row2[regimeStartCol + 1] = "USA_VIX";
    row2[regimeStartCol + 2] = "INDIA_REGIME";
    row2[regimeStartCol + 3] = "INDIA_RATIO";
    row2[regimeStartCol + 4] = "INDIA_VIX";
    
    dataSheet.getRange(2, 1, 1, finalTotalCols)
      .setValues([row2])
      .setFontWeight("bold");
    
    // Add formulas for 52WH and 52WL values
    for (let i = 0; i < tickers.length; i++) {
      const t = tickers[i];
      const c = (i * colsPer) + 1;
      dataSheet.getRange(2, c + 2).setFormula(`=IFERROR(MAX(QUERY(GOOGLEFINANCE("${t}","high",TODAY()-365,TODAY()),"SELECT Col2 LABEL Col2 ''")),0)`);
      dataSheet.getRange(2, c + 4).setFormula(`=IFERROR(MIN(QUERY(GOOGLEFINANCE("${t}","low",TODAY()-365,TODAY()),"SELECT Col2 LABEL Col2 ''")),0)`);
    }

    ss.toast("Building formulas...", "⏳ Processing", 2);

    // ========================================================================
    // ROW 3: ATH / P-E / EPS + Market Regime
    // ========================================================================
    const row3Formulas = new Array(finalTotalCols).fill("");
    for (let i = 0; i < tickers.length; i++) {
      const t = tickers[i];
      const b = i * colsPer;
      row3Formulas[b + 1] = `=MAX(QUERY(GOOGLEFINANCE("${t}","high","1/1/2000",TODAY()),"SELECT Col2 LABEL Col2 ''"))`;
      row3Formulas[b + 3] = `=IFERROR(GOOGLEFINANCE("${t}","pe"),"")`;
      row3Formulas[b + 5] = `=IFERROR(GOOGLEFINANCE("${t}","eps"),"")`;
    }
    
    // USA Market Regime
    row3Formulas[regimeStartCol - 1] = 
      `=LET(spyPrice${SEP}IFERROR(GOOGLEFINANCE("SPY"${SEP}"price")${SEP}0)${SEP}` +
      `spySMA200${SEP}IFERROR(AVERAGE(QUERY(GOOGLEFINANCE("SPY"${SEP}"close"${SEP}TODAY()-250${SEP}TODAY())${SEP}"SELECT Col2 ORDER BY Col1 DESC LIMIT 200"))${SEP}spyPrice)${SEP}` +
      `regimeRatio${SEP}IF(spySMA200>0${SEP}spyPrice/spySMA200${SEP}1)${SEP}` +
      `vixLevel${SEP}IFERROR(GOOGLEFINANCE("INDEXCBOE:VIX"${SEP}"price")${SEP}20)${SEP}` +
      `IFS(AND(regimeRatio>=1.05${SEP}vixLevel<=18)${SEP}"STRONG BULL"${SEP}` +
      `AND(regimeRatio>=1.02${SEP}vixLevel<=25)${SEP}"BULL"${SEP}` +
      `AND(regimeRatio>=0.98${SEP}vixLevel<=30)${SEP}"NEUTRAL"${SEP}` +
      `AND(regimeRatio>=0.95${SEP}vixLevel<=35)${SEP}"BEAR"${SEP}` +
      `TRUE${SEP}"STRONG BEAR"))`;
    
    row3Formulas[regimeStartCol] = 
      `=IFERROR(GOOGLEFINANCE("SPY"${SEP}"price")/AVERAGE(QUERY(GOOGLEFINANCE("SPY"${SEP}"close"${SEP}TODAY()-250${SEP}TODAY())${SEP}"SELECT Col2 ORDER BY Col1 DESC LIMIT 200"))${SEP}1)`;
    
    row3Formulas[regimeStartCol + 1] = `=IFERROR(GOOGLEFINANCE("INDEXCBOE:VIX"${SEP}"price")${SEP}20)`;
    
    // India Market Regime
    row3Formulas[regimeStartCol + 2] = 
      `=LET(niftyPrice${SEP}IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"price")${SEP}0)${SEP}` +
      `niftySMA200${SEP}IFERROR(AVERAGE(QUERY(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"close"${SEP}TODAY()-250${SEP}TODAY())${SEP}"SELECT Col2 ORDER BY Col1 DESC LIMIT 200"))${SEP}niftyPrice)${SEP}` +
      `regimeRatio${SEP}IF(niftySMA200>0${SEP}niftyPrice/niftySMA200${SEP}1)${SEP}` +
      `vixLevel${SEP}IFERROR(GOOGLEFINANCE("INDEXNSE:INDIAVIX"${SEP}"price")${SEP}20)${SEP}` +
      `IFS(AND(regimeRatio>=1.05${SEP}vixLevel<=18)${SEP}"STRONG BULL"${SEP}` +
      `AND(regimeRatio>=1.02${SEP}vixLevel<=25)${SEP}"BULL"${SEP}` +
      `AND(regimeRatio>=0.98${SEP}vixLevel<=30)${SEP}"NEUTRAL"${SEP}` +
      `AND(regimeRatio>=0.95${SEP}vixLevel<=35)${SEP}"BEAR"${SEP}` +
      `TRUE${SEP}"STRONG BEAR"))`;
    
    row3Formulas[regimeStartCol + 3] = 
      `=IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"price")/AVERAGE(QUERY(GOOGLEFINANCE("INDEXNSE:NIFTY_50"${SEP}"close"${SEP}TODAY()-250${SEP}TODAY())${SEP}"SELECT Col2 ORDER BY Col1 DESC LIMIT 200"))${SEP}1)`;
    
    row3Formulas[regimeStartCol + 4] = `=IFERROR(GOOGLEFINANCE("INDEXNSE:INDIAVIX"${SEP}"price")${SEP}20)`;
    
    dataSheet.getRange(3, 1, 1, finalTotalCols).setFormulas([row3Formulas]);

    // Write labels for row 3
    for (let i = 0; i < tickers.length; i++) {
      const c = (i * colsPer) + 1;
      dataSheet.getRange(3, c).setValue("ATH:");
      dataSheet.getRange(3, c + 2).setValue("P/E:");
      dataSheet.getRange(3, c + 4).setValue("EPS:");
    }

    ss.toast("Fetching historical data...", "⏳ Loading", 2);

    // ========================================================================
    // ROW 4: GOOGLEFINANCE(all) - Historical data
    // ========================================================================
    const row4Formulas = new Array(finalTotalCols).fill("");
    for (let i = 0; i < tickers.length; i++) {
      const t = tickers[i];
      row4Formulas[i * colsPer] = `=IFERROR(GOOGLEFINANCE("${t}","all",TODAY()-800,TODAY()),"No Data")`;
    }
    dataSheet.getRange(4, 1, 1, finalTotalCols).setFormulas([row4Formulas]);

    // ========================================================================
    // NUMBER FORMATTING
    // ========================================================================
    for (let i = 0; i < tickers.length; i++) {
      const c = (i * colsPer) + 1;
      // Row 2: 52WH and 52WL formatting
      dataSheet.getRange(2, c + 2).setNumberFormat("#,##0.00");
      dataSheet.getRange(2, c + 4).setNumberFormat("#,##0.00");
      // Row 3: ATH, P/E, EPS formatting
      dataSheet.getRange(3, c + 1).setNumberFormat("#,##0.00");
      dataSheet.getRange(3, c + 3).setNumberFormat("0.00");
      dataSheet.getRange(3, c + 5).setNumberFormat("0.00");
    }
    
    // Market regime formatting
    dataSheet.getRange(3, regimeStartCol, 1, 1).setNumberFormat("@");
    dataSheet.getRange(3, regimeStartCol + 1, 1, 1).setNumberFormat("0.000");
    dataSheet.getRange(3, regimeStartCol + 2, 1, 1).setNumberFormat("0.0");
    dataSheet.getRange(3, regimeStartCol + 3, 1, 1).setNumberFormat("@");
    dataSheet.getRange(3, regimeStartCol + 4, 1, 1).setNumberFormat("0.000");
    dataSheet.getRange(3, regimeStartCol + 5, 1, 1).setNumberFormat("0.0");

    // Historical data formatting
    for (let i = 0; i < tickers.length; i++) {
      const colStart = (i * colsPer) + 1;
      dataSheet.getRange(5, colStart, 1000, 1).setNumberFormat("yyyy-mm-dd");
      dataSheet.getRange(5, colStart + 1, 1000, 5).setNumberFormat("#,##0.00");
    }

    // ========================================================================
    // STYLING
    // ========================================================================
    const LABEL_BG = "#1F2937";
    const LABEL_FG = "#F9FAFB";
    
    // Row 2: 52WH/52WL label styling
    const row2LabelA1s = [];
    for (let i = 0; i < tickers.length; i++) {
      const c = (i * colsPer) + 1;
      row2LabelA1s.push(dataSheet.getRange(2, c + 1).getA1Notation());
      row2LabelA1s.push(dataSheet.getRange(2, c + 3).getA1Notation());
    }
    dataSheet.getRangeList(row2LabelA1s)
      .setBackground(LABEL_BG)
      .setFontColor(LABEL_FG)
      .setFontWeight("bold")
      .setHorizontalAlignment("left");
    
    // Row 3: ATH/P/E/EPS label styling
    const row3LabelA1s = [];
    for (let i = 0; i < tickers.length; i++) {
      const c = (i * colsPer) + 1;
      row3LabelA1s.push(dataSheet.getRange(3, c).getA1Notation());
      row3LabelA1s.push(dataSheet.getRange(3, c + 2).getA1Notation());
      row3LabelA1s.push(dataSheet.getRange(3, c + 4).getA1Notation());
    }
    dataSheet.getRangeList(row3LabelA1s)
      .setBackground(LABEL_BG)
      .setFontColor(LABEL_FG)
      .setFontWeight("bold")
      .setHorizontalAlignment("left");
    
    // Market regime styling
    dataSheet.getRange(2, regimeStartCol, 1, regimeColsNeeded)
      .setBackground("#4A148C")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    dataSheet.getRange(3, regimeStartCol, 1, regimeColsNeeded)
      .setBackground("#E1BEE7")
      .setFontColor("#4A148C")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    SpreadsheetApp.flush();
    
    const elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    ss.toast(`✅ Data loaded for ${tickers.length} tickers in ${elapsed}s`, "Success", 3);
    
  } catch (error) {
    console.error("generateDataSheet error:", error);
    ss.toast(`Error: ${error.message}`, "❌ Failed", 5);
    throw error;
  }
}
