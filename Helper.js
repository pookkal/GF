/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v3.1_KIRO_OPTIMIZED
* ==============================================================================
*/

function getCleanTickers(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange(3, 1, lastRow - 2, 1)
    .getValues()
    .flat()
    .filter(t => t && t.toString().trim() !== "")
    .map(t => t.toString().toUpperCase().trim());
}

/**
 * Gets a safe temporary cell for calculations
 * Uses a dedicated hidden sheet to prevent interference with existing formulas
 * @returns {Range} A cell in a hidden sheet dedicated to temporary operations
 */
function getSafeTempCell_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const TEMP_SHEET_NAME = "TEMP_CALCULATIONS";
  
  // Get or create the hidden temporary sheet
  let tempSheet = ss.getSheetByName(TEMP_SHEET_NAME);
  if (!tempSheet) {
    tempSheet = ss.insertSheet(TEMP_SHEET_NAME);
    tempSheet.hideSheet();
  }
  
  // Return cell A1 from the hidden sheet
  return tempSheet.getRange('A1');
}

/**
 * Cleans up temporary cell after use
 * Clears both formula and content to ensure no residual data
 * @param {Range} tempCell - The temporary cell to clean
 */
function cleanupTempCell_(tempCell) {
  if (tempCell) {
    try {
      tempCell.clearContent();
      tempCell.clearFormat();
    } catch (e) {
      // Log error but don't throw - cleanup should be best-effort
      console.log(`Warning: Failed to cleanup temp cell: ${e.toString()}`);
    }
  }
}

/**
 * Fetches live price for a ticker with fallback to CALCULATIONS sheet
 * Uses safe temporary cell management and retry logic with exponential backoff
 * @param {string} ticker - The ticker symbol
 * @param {number} fallbackPrice - Price from CALCULATIONS sheet to use if API fails
 * @returns {number} Live price or fallback price
 */
function getLivePriceSafely(ticker, fallbackPrice) {
  const MAX_RETRIES = 2;
  const RETRY_DELAYS = [500, 1000]; // milliseconds
  
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    let tempCell = null;
    
    try {
      // Get safe temporary cell from hidden sheet
      tempCell = getSafeTempCell_();
      
      // Write GOOGLEFINANCE formula to temporary cell
      const formula = `=GOOGLEFINANCE("${ticker}", "price")`;
      tempCell.setFormula(formula);
      
      // Flush and wait for API response
      SpreadsheetApp.flush();
      Utilities.sleep(500);
      
      // Get the response value
      const response = tempCell.getValue();
      
      // Validate response is a positive finite number
      // Note: We accept any positive number > 0, not just >= 0.01
      if (typeof response === 'number' && 
          isFinite(response) && 
          response > 0) {
        // Success - return validated price
        return response;
      }
      
      // Invalid response - log and retry or fallback
      console.log(`Attempt ${attempt + 1}: Invalid GOOGLEFINANCE response for ${ticker}: ${response}`);
      
      // If we have retries left, wait and try again
      if (attempt < MAX_RETRIES) {
        console.log(`Retrying in ${RETRY_DELAYS[attempt]}ms...`);
        Utilities.sleep(RETRY_DELAYS[attempt]);
      }
      
    } catch (error) {
      // API error - log and retry or fallback
      console.log(`Attempt ${attempt + 1}: GOOGLEFINANCE API error for ${ticker}: ${error.toString()}`);
      
      // If we have retries left, wait and try again
      if (attempt < MAX_RETRIES) {
        console.log(`Retrying in ${RETRY_DELAYS[attempt]}ms...`);
        Utilities.sleep(RETRY_DELAYS[attempt]);
      }
      
    } finally {
      // Always clean up temporary cell
      cleanupTempCell_(tempCell);
    }
  }
  
  // All retries exhausted - use fallback price
  console.log(`All retries exhausted for ${ticker}. Using fallback price: ${fallbackPrice}`);
  return fallbackPrice;
}

// Export for testing (Node.js environment)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getSafeTempCell_,
    cleanupTempCell_,
    getLivePriceSafely
  };
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

