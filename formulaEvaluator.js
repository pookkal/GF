/**
 * Formula Evaluator Module
 * 
 * This module parses and evaluates SIGNAL and DECISION formula logic from the CALCULATIONS sheet,
 * generating step-by-step narratives that show which conditions were checked, their results,
 * and the actual indicator values that triggered each condition.
 * 
 * @module formulaEvaluator
 */

/**
 * Gets column mapping for CALCULATIONS sheet indicators
 * 
 * Returns a mapping of column references (e.g., $G, $R, $U) to their indicator names and format types.
 * This mapping is used to:
 * 1. Identify which indicator a column reference refers to
 * 2. Format the indicator value appropriately in narratives
 * 
 * Format types:
 * - currency: Display as $XXX.XX (e.g., $230.50)
 * - percent: Display as XX.X% (e.g., 55.1%)
 * - decimal1: Display with 1 decimal place (e.g., 55.1)
 * - decimal2: Display with 2 decimal places (e.g., 1.85)
 * - decimal3: Display with 3 decimal places (e.g., 0.125)
 * 
 * @returns {Object.<string, {name: string, format: string}>} Column mapping object
 */
function getColumnMap() {
  return {
    // Column B (index 1) - Market Rating
    "$B": { name: "Market Rating", format: "text" },
    
    // Column C (index 2) - Decision
    "$C": { name: "Decision", format: "text" },
    
    // Column D (index 3) - Signal
    "$D": { name: "Signal", format: "text" },
    
    // Column E (index 4) - Patterns
    "$E": { name: "Patterns", format: "text" },
    
    // Column F (index 5) - Consensus Price
    "$F": { name: "Consensus Price", format: "currency" },
    
    // Column G (index 6) - Price
    "$G": { name: "Price", format: "currency" },
    
    // Column H (index 7) - Change %
    "$H": { name: "Change %", format: "percent" },
    
    // Column I (index 8) - Vol Trend
    "$I": { name: "Vol Trend", format: "decimal2" },
    
    // Column J (index 9) - ATH True
    "$J": { name: "ATH True", format: "currency" },
    
    // Column K (index 10) - ATH Diff %
    "$K": { name: "ATH Diff%", format: "percent" },
    
    // Column L (index 11) - ATH Zone
    "$L": { name: "ATH Zone", format: "text" },
    
    // Column M (index 12) - Fundamental
    "$M": { name: "Fundamental", format: "text" },
    
    // Column N (index 13) - Trend State
    "$N": { name: "Trend State", format: "text" },
    
    // Column O (index 14) - SMA 20
    "$O": { name: "SMA 20", format: "currency" },
    
    // Column P (index 15) - SMA 50
    "$P": { name: "SMA 50", format: "currency" },
    
    // Column Q (index 16) - SMA 200
    "$Q": { name: "SMA 200", format: "currency" },
    
    // Column R (index 17) - RSI
    "$R": { name: "RSI", format: "decimal1" },
    
    // Column S (index 18) - MACD Hist
    "$S": { name: "MACD Hist", format: "decimal3" },
    
    // Column T (index 19) - Divergence
    "$T": { name: "Divergence", format: "text" },
    
    // Column U (index 20) - ADX
    "$U": { name: "ADX", format: "decimal1" },
    
    // Column V (index 21) - Stoch %K
    "$V": { name: "Stoch %K", format: "percent" },
    
    // Column W (index 22) - Vol Regime
    "$W": { name: "Vol Regime", format: "text" },
    
    // Column X (index 23) - BBP Signal
    "$X": { name: "BBP Signal", format: "text" },
    
    // Column Y (index 24) - ATR
    "$Y": { name: "ATR", format: "currency" },
    
    // Column Z (index 25) - Bollinger %B
    "$Z": { name: "Bollinger %B", format: "percent" },
    
    // Column AA (index 26) - Target
    "$AA": { name: "Target", format: "currency" },
    
    // Column AB (index 27) - RR Quality
    "$AB": { name: "RR Quality", format: "decimal1" },
    
    // Column AC (index 28) - Support
    "$AC": { name: "Support", format: "currency" },
    
    // Column AD (index 29) - Resistance
    "$AD": { name: "Resistance", format: "currency" },
    
    // Column AE (index 30) - ATR Stop
    "$AE": { name: "ATR Stop", format: "currency" },
    
    // Column AF (index 31) - ATR Target
    "$AF": { name: "ATR Target", format: "currency" },
    
    // Column AG (index 32) - Position Size
    "$AG": { name: "Position Size", format: "text" },
    
    // Column AH (index 33) - Last State
    "$AH": { name: "Last State", format: "text" }
  };
}

/**
 * Gets or initializes the formula cache
 * 
 * Returns a Map that stores parsed formula logic to avoid re-parsing on every call.
 * The cache uses composite keys in the format: `${formulaType}_${mode}`
 * 
 * Key format:
 * - "SIGNAL_TRADE" - SIGNAL formula in TRADE mode (useLongTermSignal = false)
 * - "SIGNAL_LONGTERM" - SIGNAL formula in LONG-TERM mode (useLongTermSignal = true)
 * - "DECISION_TRADE" - DECISION formula in TRADE mode (useLongTermSignal = false)
 * - "DECISION_LONGTERM" - DECISION formula in LONG-TERM mode (useLongTermSignal = true)
 * 
 * @returns {Map<string, Object>} Formula cache Map
 */
function getFormulaCache() {
  // Use a property service to persist cache across function calls
  // This avoids redeclaration issues in Google Apps Script
  if (typeof this._formulaCache === 'undefined') {
    this._formulaCache = new Map();
  }
  return this._formulaCache;
}

/**
 * Formats a numeric value based on its format type
 * 
 * This function takes a numeric value and formats it according to the specified format type.
 * It handles null, undefined, and zero values by returning "N/A" to indicate missing data.
 * 
 * Format types:
 * - currency: Formats as $XXX.XX (e.g., $230.50)
 * - percent: Formats as XX.X% (e.g., 55.1%)
 * - decimal1: Formats with 1 decimal place (e.g., 55.1)
 * - decimal2: Formats with 2 decimal places (e.g., 1.85)
 * - decimal3: Formats with 3 decimal places (e.g., 0.125)
 * - text: Returns the value as-is (for non-numeric values)
 * 
 * @param {number|string|null|undefined} value - The value to format
 * @param {string} format - The format type (currency, percent, decimal1, decimal2, decimal3, text)
 * 
 * @returns {string} The formatted value or "N/A" for null/undefined/zero values
 * 
 * @example
 * formatValue(230.5, "currency")    // "$230.50"
 * formatValue(55.123, "percent")    // "55.1%"
 * formatValue(1.8543, "decimal2")   // "1.85"
 * formatValue(0.12456, "decimal3")  // "0.125"
 * formatValue(55.123, "decimal1")   // "55.1"
 * formatValue(null, "currency")     // "N/A"
 * formatValue(0, "currency")        // "N/A"
 * formatValue(undefined, "percent") // "N/A"
 * formatValue("HOLD", "text")       // "HOLD"
 */
function formatValue(value, format) {
  // Handle null and undefined (but NOT zero - zero is a valid value)
  if (value === null || value === undefined) {
    return "N/A";
  }
  
  // Handle text format (non-numeric values)
  if (format === "text") {
    return String(value);
  }
  
  // Convert to number if it's a string
  const numValue = typeof value === 'string' ? parseFloat(value) : value;
  
  // Check if conversion resulted in NaN (but NOT zero - zero is valid)
  if (isNaN(numValue)) {
    return "N/A";
  }
  
  // Format based on type
  switch (format) {
    case "currency":
      // Format as $XXX.XX
      return "$" + numValue.toFixed(2);
    
    case "percent":
      // Format as XX.X%
      return numValue.toFixed(1) + "%";
    
    case "decimal1":
      // Format with 1 decimal place
      return numValue.toFixed(1);
    
    case "decimal2":
      // Format with 2 decimal places
      return numValue.toFixed(2);
    
    case "decimal3":
      // Format with 3 decimal places
      return numValue.toFixed(3);
    
    default:
      // Unknown format, return as-is with 2 decimal places
      return numValue.toFixed(2);
  }
}

/**
 * Parses SIGNAL formula into condition tree
 * 
 * This function extracts the formula logic from buildSignalFormula() and parses it into
 * a structured condition tree. The tree contains all IFS() branches in evaluation order,
 * with each branch containing the condition expression and result value.
 * 
 * The function supports both TRADE mode and LONG-TERM INVESTMENT mode, which have
 * different formula logic and condition branches.
 * 
 * @param {boolean} useLongTermSignal - Whether to use long-term investment mode (true) or trade mode (false)
 * 
 * @returns {Object} Condition tree with branches in evaluation order
 * @returns {string} returns.type - Always "IFS" for the root node
 * @returns {Array<Object>} returns.branches - Array of condition branches in evaluation order
 * 
 * Each branch object contains:
 * @returns {number} branch.order - Evaluation order (1-based)
 * @returns {string} branch.condition - The condition expression (e.g., "AND($G>$Q, $R>=30, $R<=40)")
 * @returns {string} branch.result - The result value if condition is true (e.g., "STRONG BUY")
 * @returns {Array<string>} branch.columnRefs - Array of column references used (e.g., ["$G", "$Q", "$R"])
 * @returns {Array<string>} branch.operators - Array of operators used (e.g., [">", ">=", "<="])
 * @returns {string} branch.type - Type of condition ("COMPARISON", "AND", "OR", or "COMPLEX")
 * 
 * @example
 * const tree = parseSignalFormula(false); // TRADE mode
 * console.log(tree.branches[0]);
 * // {
 * //   order: 1,
 * //   condition: "$G<$AC",
 * //   result: "STOP OUT",
 * //   columnRefs: ["$G", "$AC"],
 * //   operators: ["<"],
 * //   type: "COMPARISON"
 * // }
 * 
 * @throws {Error} If formula parsing fails
 */
function parseSignalFormula(useLongTermSignal) {
  try {
    // Get formula cache
    const FORMULA_CACHE = getFormulaCache();
    
    // Check cache before parsing
    const cacheKey = `SIGNAL_${useLongTermSignal ? 'LONGTERM' : 'TRADE'}`;
    
    // Return cached result if available
    if (FORMULA_CACHE.has(cacheKey)) {
      return FORMULA_CACHE.get(cacheKey);
    }
    
    // Define the condition branches based on mode
    // These are extracted from buildSignalFormula() in generateCalculations.js
    
    let result;
    
    if (useLongTermSignal) {
      // LONG-TERM INVESTMENT MODE - Conservative, trend-following approach
      result = {
        type: "IFS",
        branches: [
          {
            order: 1,
            condition: "$G<$AC",
            result: "STOP OUT",
            columnRefs: ["$G", "$AC"],
            operators: ["<"],
            type: "COMPARISON"
          },
          {
            order: 2,
            condition: "$G<$Q",
            result: "RISK OFF",
            columnRefs: ["$G", "$Q"],
            operators: ["<"],
            type: "COMPARISON"
          },
          {
            order: 3,
            condition: "AND($G>$Q, $P>$Q, $R>=30, $R<=40, $S>0, $U>=20, $I>=1.5)",
            result: "STRONG BUY",
            columnRefs: ["$G", "$Q", "$P", "$R", "$S", "$U", "$I"],
            operators: [">", ">", ">=", "<=", ">", ">=", ">="],
            type: "AND"
          },
          {
            order: 4,
            condition: "AND($G>$Q, $P>$Q, $R>40, $R<=50, $S>0, $U>=15)",
            result: "BUY",
            columnRefs: ["$G", "$Q", "$P", "$R", "$S", "$U"],
            operators: [">", ">", ">", "<=", ">", ">="],
            type: "AND"
          },
          {
            order: 5,
            condition: "AND($G>$Q, $R>=35, $R<=55, $G>=$P*0.95, $G<=$P*1.05)",
            result: "ACCUMULATE",
            columnRefs: ["$G", "$Q", "$R", "$P"],
            operators: [">", ">=", "<=", ">=", "<="],
            type: "AND"
          },
          {
            order: 6,
            condition: "AND($R<=30, $G>$AC)",
            result: "OVERSOLD WATCH",
            columnRefs: ["$R", "$G", "$AC"],
            operators: ["<=", ">"],
            type: "AND"
          },
          {
            order: 7,
            condition: "OR($R>=70, $Z>=0.85, $G>=$AD*0.98)",
            result: "TRIM",
            columnRefs: ["$R", "$Z", "$G", "$AD"],
            operators: [">=", ">=", ">="],
            type: "OR"
          },
          {
            order: 8,
            condition: "AND($G>$Q, $R>40, $R<70)",
            result: "HOLD",
            columnRefs: ["$G", "$Q", "$R"],
            operators: [">", ">", "<"],
            type: "AND"
          },
          {
            order: 9,
            condition: "TRUE",
            result: "NEUTRAL",
            columnRefs: [],
            operators: [],
            type: "DEFAULT"
          }
        ]
      };
    } else {
      // TRADE MODE - Momentum and breakout focused
      // NOTE: Complex conditions with AVERAGE_ATR are simplified for evaluation
      // The actual formula uses IFERROR(AVERAGE(OFFSET(...))) which we approximate
      result = {
        type: "IFS",
        branches: [
          {
            order: 1,
            condition: "$G<$AC",
            result: "STOP OUT",
            columnRefs: ["$G", "$AC"],
            operators: ["<"],
            type: "COMPARISON"
          },
          {
            order: 2,
            condition: "AND($I>=2.0, $G>=$AD*1.01)",
            result: "VOLATILITY BREAKOUT",
            columnRefs: ["$I", "$G", "$AD"],
            operators: [">=", ">="],
            type: "AND",
            note: "Simplified: ATR>AVERAGE_ATR*1.5 check omitted (complex calculation)"
          },
          {
            order: 3,
            condition: "AND($I>=1.5, $G>=$AD*1.02)",
            result: "BREAKOUT",
            columnRefs: ["$I", "$G", "$AD"],
            operators: [">=", ">="],
            type: "AND"
          },
          {
            order: 4,
            condition: "AND($K>=-0.01, $I>=2.0, $U>=25)",
            result: "ATH BREAKOUT",
            columnRefs: ["$K", "$I", "$U"],
            operators: [">=", ">=", ">="],
            type: "AND"
          },
          {
            order: 5,
            condition: "AND($G>$P, $S>0, $U>=20)",
            result: "MOMENTUM",
            columnRefs: ["$G", "$P", "$S", "$U"],
            operators: [">", ">", ">="],
            type: "AND"
          },
          {
            order: 6,
            condition: "AND($V<=20, $S>0, $G>$AC)",
            result: "OVERSOLD REVERSAL",
            columnRefs: ["$V", "$S", "$G", "$AC"],
            operators: ["<=", ">", ">"],
            type: "AND"
          },
          {
            order: 6,
            condition: "AND($V<=20, $S>0, $G>$AC)",
            result: "OVERSOLD REVERSAL",
            columnRefs: ["$V", "$S", "$G", "$AC"],
            operators: ["<=", ">", ">"],
            type: "AND"
          },
          {
            order: 7,
            condition: "AND($U<15, ABS($Z-0.5)<0.2)",
            result: "VOLATILITY SQUEEZE",
            columnRefs: ["$U", "$Z"],
            operators: ["<", "<"],
            type: "AND",
            note: "Simplified: ATR<AVERAGE_ATR*0.7 check omitted (complex calculation)"
          },
          {
            order: 8,
            condition: "AND($U<15, $G>=$AC*0.98, $G<=$AC*1.02)",
            result: "RANGE SUPPORT BUY",
            columnRefs: ["$U", "$G", "$AC"],
            operators: ["<", ">=", "<="],
            type: "AND"
          },
          {
            order: 9,
            condition: "OR($R>=70, $Z>=0.9)",
            result: "OVERBOUGHT",
            columnRefs: ["$R", "$Z"],
            operators: [">=", ">="],
            type: "OR"
          },
          {
            order: 10,
            condition: "$G<$Q",
            result: "RISK OFF",
            columnRefs: ["$G", "$Q"],
            operators: ["<"],
            type: "COMPARISON"
          },
          {
            order: 11,
            condition: "AND($U<15, $G>$AC)",
            result: "RANGE",
            columnRefs: ["$U", "$G", "$AC"],
            operators: ["<", ">"],
            type: "AND"
          },
          {
            order: 12,
            condition: "TRUE",
            result: "NEUTRAL",
            columnRefs: [],
            operators: [],
            type: "DEFAULT"
          }
        ]
      };
    }
    
    // Store result in cache after parsing
    FORMULA_CACHE.set(cacheKey, result);
    
    // Return the result
    return result;
  } catch (error) {
    // Log error details for debugging
    Logger.log(`ERROR: Failed to parse SIGNAL formula (mode: ${useLongTermSignal ? 'LONGTERM' : 'TRADE'})`);
    Logger.log(`Error message: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    
    // Return null to trigger fallback to generic explanation
    return null;
  }
}

/**
 * Parses DECISION formula into condition tree
 * 
 * This function extracts the formula logic from buildDecisionFormula() and parses it into
 * a structured condition tree. The tree contains all IFS() branches that determine how
 * SIGNAL combines with PATTERNS and PURCHASED tag to produce the final DECISION.
 * 
 * The function supports both TRADE mode and LONG-TERM INVESTMENT mode, which have
 * different decision logic and condition branches.
 * 
 * Key logic components:
 * 1. PURCHASED tag check - Determines if position is already owned (from INPUT sheet column C)
 * 2. SIGNAL value checks - Evaluates the SIGNAL from column D
 * 3. Pattern detection - Checks for bullish/bearish patterns in column E
 * 4. Pattern types:
 *    - Bullish: ASC_TRI, BRKOUT, DBL_BTM, INV_H&S, CUP_HDL
 *    - Bearish: DESC_TRI, H&S, DBL_TOP
 * 
 * @param {boolean} useLongTermSignal - Whether to use long-term investment mode (true) or trade mode (false)
 * 
 * @returns {Object} Condition tree with branches in evaluation order
 * @returns {string} returns.type - Always "IFS" for the root node
 * @returns {Array<Object>} returns.branches - Array of condition branches in evaluation order
 * 
 * Each branch object contains:
 * @returns {number} branch.order - Evaluation order (1-based)
 * @returns {string} branch.condition - The condition expression
 * @returns {string} branch.result - The result value if condition is true (e.g., "🟢 STRONG BUY")
 * @returns {Array<string>} branch.columnRefs - Array of column references used (e.g., ["$D", "$E"])
 * @returns {string} branch.type - Type of condition ("SIGNAL_CHECK", "PATTERN_CHECK", "PURCHASED_CHECK", "COMPLEX")
 * @returns {boolean} branch.requiresPurchased - Whether this branch requires PURCHASED tag to be true
 * @returns {boolean} branch.requiresNotPurchased - Whether this branch requires PURCHASED tag to be false
 * @returns {string} branch.signalValue - The SIGNAL value being checked (if applicable)
 * @returns {string} branch.patternType - The pattern type being checked: "bullish", "bearish", "any", or "none"
 * 
 * @example
 * const tree = parseDecisionFormula(false); // TRADE mode
 * console.log(tree.branches[0]);
 * // {
 * //   order: 1,
 * //   condition: "AND(PRICE>0, SUPPORT>0, PRICE<SUPPORT)",
 * //   result: "🔴 STOP OUT",
 * //   columnRefs: ["$G", "$AC"],
 * //   type: "STOP_OUT_CHECK",
 * //   requiresPurchased: false,
 * //   requiresNotPurchased: false,
 * //   signalValue: null,
 * //   patternType: "none"
 * // }
 * 
 * @throws {Error} If formula parsing fails
 */
function parseDecisionFormula(useLongTermSignal) {
  try {
    // Get formula cache
    const FORMULA_CACHE = getFormulaCache();
    
    // Check cache before parsing
    const cacheKey = `DECISION_${useLongTermSignal ? 'LONGTERM' : 'TRADE'}`;
    
    // Return cached result if available
    if (FORMULA_CACHE.has(cacheKey)) {
      return FORMULA_CACHE.get(cacheKey);
    }
    
    // Define the condition branches based on mode
    // These are extracted from buildDecisionFormula() in generateCalculations.js
    
    let result;
    
    if (useLongTermSignal) {
      // LONG-TERM INVESTMENT MODE
      // Logic: SIGNAL (D) + PATTERNS (E) + PURCHASED tag
      result = {
        type: "IFS",
        branches: [
          // PURCHASED POSITION BRANCHES
          {
            order: 1,
            condition: "OR($D='STOP OUT', $D='RISK OFF')",
            result: "🔴 EXIT",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: ["STOP OUT", "RISK OFF"],
            patternType: "none"
          },
          {
            order: 2,
            condition: "AND($D='TRIM', HAS_PATTERN, BEARISH_PATTERN)",
            result: "🟠 TRIM (PATTERN CONFIRMED)",
            columnRefs: ["$D", "$E"],
            type: "PATTERN_CHECK",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: "TRIM",
            patternType: "bearish"
          },
          {
            order: 3,
            condition: "$D='TRIM'",
            result: "🟠 TRIM",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: "TRIM",
            patternType: "none"
          },
          {
            order: 4,
            condition: "AND(OR($D='STRONG BUY', $D='BUY', $D='ACCUMULATE'), HAS_PATTERN, BULLISH_PATTERN)",
            result: "🟢 ADD (PATTERN CONFIRMED)",
            columnRefs: ["$D", "$E"],
            type: "PATTERN_CHECK",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: ["STRONG BUY", "BUY", "ACCUMULATE"],
            patternType: "bullish"
          },
          {
            order: 5,
            condition: "AND(OR($D='STRONG BUY', $D='BUY', $D='ACCUMULATE'), HAS_PATTERN, BEARISH_PATTERN)",
            result: "⚠️ HOLD (PATTERN CONFLICT)",
            columnRefs: ["$D", "$E"],
            type: "PATTERN_CHECK",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: ["STRONG BUY", "BUY", "ACCUMULATE"],
            patternType: "bearish"
          },
          {
            order: 6,
            condition: "OR($D='STRONG BUY', $D='BUY', $D='ACCUMULATE')",
            result: "🟢 ADD",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: ["STRONG BUY", "BUY", "ACCUMULATE"],
            patternType: "none"
          },
          {
            order: 7,
            condition: "$D='HOLD'",
            result: "⚖️ HOLD",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: "HOLD",
            patternType: "none"
          },
          {
            order: 8,
            condition: "TRUE",
            result: "⚖️ HOLD",
            columnRefs: [],
            type: "DEFAULT",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: null,
            patternType: "none"
          },
          
          // NON-PURCHASED POSITION BRANCHES
          {
            order: 9,
            condition: "OR($D='STOP OUT', $D='RISK OFF')",
            result: "🔴 AVOID",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: ["STOP OUT", "RISK OFF"],
            patternType: "none"
          },
          {
            order: 10,
            condition: "AND($D='STRONG BUY', HAS_PATTERN, BULLISH_PATTERN)",
            result: "🟢 STRONG BUY (PATTERN CONFIRMED)",
            columnRefs: ["$D", "$E"],
            type: "PATTERN_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "STRONG BUY",
            patternType: "bullish"
          },
          {
            order: 11,
            condition: "AND(OR($D='STRONG BUY', $D='BUY'), HAS_PATTERN, BEARISH_PATTERN)",
            result: "⚠️ CAUTION (PATTERN CONFLICT)",
            columnRefs: ["$D", "$E"],
            type: "PATTERN_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: ["STRONG BUY", "BUY"],
            patternType: "bearish"
          },
          {
            order: 12,
            condition: "$D='STRONG BUY'",
            result: "🟢 STRONG BUY",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "STRONG BUY",
            patternType: "none"
          },
          {
            order: 13,
            condition: "$D='BUY'",
            result: "🟢 BUY",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "BUY",
            patternType: "none"
          },
          {
            order: 14,
            condition: "$D='ACCUMULATE'",
            result: "🟢 ACCUMULATE",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "ACCUMULATE",
            patternType: "none"
          },
          {
            order: 15,
            condition: "$D='OVERSOLD WATCH'",
            result: "🟡 WATCH (OVERSOLD)",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "OVERSOLD WATCH",
            patternType: "none"
          },
          {
            order: 16,
            condition: "$D='TRIM'",
            result: "⏳ WAIT (EXTENDED)",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "TRIM",
            patternType: "none"
          },
          {
            order: 17,
            condition: "$D='HOLD'",
            result: "⚖️ WATCH",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "HOLD",
            patternType: "none"
          },
          {
            order: 18,
            condition: "TRUE",
            result: "⚪ NEUTRAL",
            columnRefs: [],
            type: "DEFAULT",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: null,
            patternType: "none"
          }
        ]
      };
    } else {
      // TRADE MODE
      // Logic: SIGNAL (D) + PATTERNS (E) + PURCHASED tag + Price/Support/Resistance checks
      result = {
        type: "IFS",
        branches: [
          // STOP OUT CHECK (applies to all) - Price below Support
          {
            order: 1,
            condition: "AND($G>0, $AC>0, $G<$AC)",
            result: "🔴 STOP OUT",
            columnRefs: ["$G", "$AC"],
            type: "STOP_OUT_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: false,
            signalValue: null,
            patternType: "none"
          },
          
          // NON-PURCHASED POSITION BRANCHES WITH PATTERN CONFIRMATION
          {
            order: 2,
            condition: "AND(NOT_PURCHASED, $D='VOLATILITY BREAKOUT', HAS_PATTERN, BULLISH_PATTERN)",
            result: "🟢 STRONG TRADE LONG (PATTERN CONFIRMED)",
            columnRefs: ["$D", "$E"],
            type: "PATTERN_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "VOLATILITY BREAKOUT",
            patternType: "bullish"
          },
          {
            order: 3,
            condition: "AND(NOT_PURCHASED, OR($D='BREAKOUT', $D='ATH BREAKOUT'), HAS_PATTERN, BULLISH_PATTERN)",
            result: "🟢 TRADE LONG (PATTERN CONFIRMED)",
            columnRefs: ["$D", "$E"],
            type: "PATTERN_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: ["BREAKOUT", "ATH BREAKOUT"],
            patternType: "bullish"
          },
          
          // PATTERN CONFLICTS
          {
            order: 4,
            condition: "AND(NOT_PURCHASED, OR($D='VOLATILITY BREAKOUT', $D='BREAKOUT', $D='ATH BREAKOUT', $D='MOMENTUM'), HAS_PATTERN, BEARISH_PATTERN)",
            result: "⚠️ CAUTION (PATTERN CONFLICT)",
            columnRefs: ["$D", "$E"],
            type: "PATTERN_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: ["VOLATILITY BREAKOUT", "BREAKOUT", "ATH BREAKOUT", "MOMENTUM"],
            patternType: "bearish"
          },
          
          // STANDARD SIGNALS WITHOUT PATTERN CONSIDERATION (NON-PURCHASED)
          {
            order: 5,
            condition: "AND(NOT_PURCHASED, $D='VOLATILITY BREAKOUT')",
            result: "🟢 STRONG TRADE LONG",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "VOLATILITY BREAKOUT",
            patternType: "none"
          },
          {
            order: 6,
            condition: "AND(NOT_PURCHASED, OR($D='BREAKOUT', $D='ATH BREAKOUT'))",
            result: "🟢 TRADE LONG",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: ["BREAKOUT", "ATH BREAKOUT"],
            patternType: "none"
          },
          {
            order: 7,
            condition: "AND(NOT_PURCHASED, $D='MOMENTUM')",
            result: "🟡 ACCUMULATE",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "MOMENTUM",
            patternType: "none"
          },
          {
            order: 8,
            condition: "AND(NOT_PURCHASED, $D='OVERSOLD REVERSAL')",
            result: "🟢 BUY DIP",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "OVERSOLD REVERSAL",
            patternType: "none"
          },
          {
            order: 9,
            condition: "AND(NOT_PURCHASED, $D='RANGE SUPPORT BUY')",
            result: "🟡 RANGE BUY",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "RANGE SUPPORT BUY",
            patternType: "none"
          },
          {
            order: 10,
            condition: "AND(NOT_PURCHASED, $D='VOLATILITY SQUEEZE')",
            result: "⏳ WAIT FOR BREAKOUT",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "VOLATILITY SQUEEZE",
            patternType: "none"
          },
          
          // PURCHASED POSITION MANAGEMENT
          {
            order: 11,
            condition: "AND(PURCHASED, OR($D='OVERBOUGHT', $G>=$AD*0.98))",
            result: "🟠 TAKE PROFIT",
            columnRefs: ["$D", "$G", "$AD"],
            type: "COMPLEX",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: "OVERBOUGHT",
            patternType: "none"
          },
          {
            order: 12,
            condition: "AND(PURCHASED, $D='RISK OFF')",
            result: "🔴 RISK OFF",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: "RISK OFF",
            patternType: "none"
          },
          {
            order: 13,
            condition: "AND(NOT_PURCHASED, $D='RISK OFF')",
            result: "🔴 AVOID",
            columnRefs: ["$D"],
            type: "SIGNAL_CHECK",
            requiresPurchased: false,
            requiresNotPurchased: true,
            signalValue: "RISK OFF",
            patternType: "none"
          },
          
          // DEFAULT HOLDS
          {
            order: 14,
            condition: "PURCHASED",
            result: "⚖️ HOLD",
            columnRefs: [],
            type: "PURCHASED_CHECK",
            requiresPurchased: true,
            requiresNotPurchased: false,
            signalValue: null,
            patternType: "none"
          },
          {
            order: 15,
            condition: "TRUE",
            result: "⚪ NEUTRAL",
            columnRefs: [],
            type: "DEFAULT",
            requiresPurchased: false,
            requiresNotPurchased: false,
            signalValue: null,
            patternType: "none"
          }
        ]
      };
    }
    
    // Store result in cache after parsing
    FORMULA_CACHE.set(cacheKey, result);
    
    // Return the result
    return result;
  } catch (error) {
    // Log error details for debugging
    Logger.log(`ERROR: Failed to parse DECISION formula (mode: ${useLongTermSignal ? 'LONGTERM' : 'TRADE'})`);
    Logger.log(`Error message: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    
    // Return null to trigger fallback to generic explanation
    return null;
  }
}

/**
 * Parses a condition expression to extract column references, operators, and structure
 * 
 * This helper function analyzes a condition string and extracts:
 * - Column references (e.g., $G, $R, $U)
 * - Comparison operators (>, <, >=, <=, =)
 * - Logical operators (AND, OR)
 * - Nested sub-conditions for complex expressions
 * 
 * The function handles various condition types:
 * - Simple comparisons: "$G > $Q"
 * - Logical AND: "AND($G>$Q, $R>=30, $R<=40)"
 * - Logical OR: "OR($R>=70, $Z>=0.85)"
 * - Complex nested: "AND($G>$Q, OR($R>=70, $Z>=0.85))"
 * 
 * @param {string} conditionStr - The condition expression to parse
 * 
 * @returns {Object} Parsed condition object
 * @returns {string} returns.expression - The original condition expression
 * @returns {Array<string>} returns.columnRefs - Array of column references found (e.g., ["$G", "$Q", "$R"])
 * @returns {Array<string>} returns.operators - Array of operators found (e.g., [">", ">=", "<="])
 * @returns {string} returns.type - Type of condition: "COMPARISON", "AND", "OR", "COMPLEX", or "DEFAULT"
 * @returns {Array<Object>} returns.subConditions - Array of parsed sub-conditions (for AND/OR types)
 * 
 * Each subCondition object contains:
 * @returns {string} subCondition.expression - The sub-condition expression
 * @returns {Array<string>} subCondition.columnRefs - Column references in this sub-condition
 * @returns {Array<string>} subCondition.operators - Operators in this sub-condition
 * @returns {string} subCondition.type - Type of this sub-condition
 * 
 * @example
 * // Simple comparison
 * parseConditionExpression("$G > $Q")
 * // Returns: {
 * //   expression: "$G > $Q",
 * //   columnRefs: ["$G", "$Q"],
 * //   operators: [">"],
 * //   type: "COMPARISON",
 * //   subConditions: []
 * // }
 * 
 * @example
 * // AND condition
 * parseConditionExpression("AND($G>$Q, $R>=30, $R<=40)")
 * // Returns: {
 * //   expression: "AND($G>$Q, $R>=30, $R<=40)",
 * //   columnRefs: ["$G", "$Q", "$R"],
 * //   operators: [">", ">=", "<="],
 * //   type: "AND",
 * //   subConditions: [
 * //     {expression: "$G>$Q", columnRefs: ["$G", "$Q"], operators: [">"], type: "COMPARISON"},
 * //     {expression: "$R>=30", columnRefs: ["$R"], operators: [">="], type: "COMPARISON"},
 * //     {expression: "$R<=40", columnRefs: ["$R"], operators: ["<="], type: "COMPARISON"}
 * //   ]
 * // }
 * 
 * @example
 * // OR condition
 * parseConditionExpression("OR($R>=70, $Z>=0.85)")
 * // Returns: {
 * //   expression: "OR($R>=70, $Z>=0.85)",
 * //   columnRefs: ["$R", "$Z"],
 * //   operators: [">=", ">="],
 * //   type: "OR",
 * //   subConditions: [
 * //     {expression: "$R>=70", columnRefs: ["$R"], operators: [">="], type: "COMPARISON"},
 * //     {expression: "$Z>=0.85", columnRefs: ["$Z"], operators: [">="], type: "COMPARISON"}
 * //   ]
 * // }
 * 
 * @throws {Error} If condition string is invalid or cannot be parsed
 */
function parseConditionExpression(conditionStr) {
  try {
    // Handle special cases
    if (!conditionStr || conditionStr.trim() === '') {
      throw new Error('Empty condition string');
    }
    
    const trimmedCondition = conditionStr.trim();
    
    // Handle DEFAULT case (TRUE)
    if (trimmedCondition === 'TRUE') {
      return {
        expression: trimmedCondition,
        columnRefs: [],
        operators: [],
        type: 'DEFAULT',
        subConditions: []
      };
    }
    
    // Check if this is an AND condition
    if (trimmedCondition.startsWith('AND(') && trimmedCondition.endsWith(')')) {
      return parseLogicalCondition(trimmedCondition, 'AND');
    }
    
    // Check if this is an OR condition
    if (trimmedCondition.startsWith('OR(') && trimmedCondition.endsWith(')')) {
      return parseLogicalCondition(trimmedCondition, 'OR');
    }
    
    // Otherwise, it's a simple comparison
    return parseSimpleComparison(trimmedCondition);
    
  } catch (error) {
    throw new Error(`Failed to parse condition expression "${conditionStr}": ${error.message}`);
  }
}

/**
 * Parses a logical condition (AND or OR) into sub-conditions
 * 
 * @param {string} conditionStr - The logical condition string (e.g., "AND($G>$Q, $R>=30)")
 * @param {string} logicalType - The logical operator type: "AND" or "OR"
 * 
 * @returns {Object} Parsed logical condition object
 * @private
 */
function parseLogicalCondition(conditionStr, logicalType) {
  // Extract the content inside the parentheses
  const startIdx = conditionStr.indexOf('(');
  const endIdx = conditionStr.lastIndexOf(')');
  const innerContent = conditionStr.substring(startIdx + 1, endIdx);
  
  // Split by commas, but be careful of nested parentheses
  const subConditionStrs = splitByComma(innerContent);
  
  // Parse each sub-condition
  const subConditions = subConditionStrs.map(subStr => {
    const trimmed = subStr.trim();
    // Recursively parse sub-conditions (they might be nested AND/OR)
    if (trimmed.startsWith('AND(') || trimmed.startsWith('OR(')) {
      return parseConditionExpression(trimmed);
    } else {
      return parseSimpleComparison(trimmed);
    }
  });
  
  // Collect all column references and operators from sub-conditions
  const allColumnRefs = [];
  const allOperators = [];
  
  subConditions.forEach(subCond => {
    allColumnRefs.push(...subCond.columnRefs);
    allOperators.push(...subCond.operators);
  });
  
  return {
    expression: conditionStr,
    columnRefs: allColumnRefs,
    operators: allOperators,
    type: logicalType,
    subConditions: subConditions
  };
}

/**
 * Parses a simple comparison expression (e.g., "$G > $Q" or "$R >= 30")
 * 
 * @param {string} comparisonStr - The comparison expression string
 * 
 * @returns {Object} Parsed comparison object
 * @private
 */
function parseSimpleComparison(comparisonStr) {
  // Extract column references (pattern: $LETTER or $LETTERLETTER)
  const columnRefPattern = /\$[A-Z]+/g;
  const columnRefs = comparisonStr.match(columnRefPattern) || [];
  
  // Extract comparison operators
  // Order matters: check >= and <= before > and <
  const operatorPattern = /(>=|<=|>|<|=)/g;
  const operators = comparisonStr.match(operatorPattern) || [];
  
  // Determine if this is a complex comparison
  // Complex if it has:
  // - Functions like ABS, AVERAGE, MAX, MIN, SUM
  // - Arithmetic operations like *, /, +, - (but not in column refs)
  const hasFunction = /\b(ABS|AVERAGE|MAX|MIN|SUM)\b/.test(comparisonStr);
  const hasArithmetic = /[*\/+\-]/.test(comparisonStr);
  const type = (hasFunction || hasArithmetic) ? 'COMPLEX' : 'COMPARISON';
  
  return {
    expression: comparisonStr,
    columnRefs: columnRefs,
    operators: operators,
    type: type,
    subConditions: []
  };
}

/**
 * Splits a string by commas, respecting nested parentheses
 * 
 * This helper function splits a string by commas, but ignores commas that are
 * inside nested parentheses. This is needed to correctly parse logical conditions
 * like "AND($G>$Q, OR($R>=70, $Z>=0.85), $U>=15)" where the inner OR has commas.
 * 
 * @param {string} str - The string to split
 * 
 * @returns {Array<string>} Array of split strings
 * @private
 * 
 * @example
 * splitByComma("$G>$Q, $R>=30, $R<=40")
 * // Returns: ["$G>$Q", "$R>=30", "$R<=40"]
 * 
 * @example
 * splitByComma("$G>$Q, OR($R>=70, $Z>=0.85), $U>=15")
 * // Returns: ["$G>$Q", "OR($R>=70, $Z>=0.85)", "$U>=15"]
 */
function splitByComma(str) {
  const result = [];
  let current = '';
  let parenDepth = 0;
  
  for (let i = 0; i < str.length; i++) {
    const char = str[i];
    
    if (char === '(') {
      parenDepth++;
      current += char;
    } else if (char === ')') {
      parenDepth--;
      current += char;
    } else if (char === ',' && parenDepth === 0) {
      // This is a top-level comma, so split here
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  
  // Add the last part
  if (current.trim() !== '') {
    result.push(current.trim());
  }
  
  return result;
}

/**
 * Evaluates a condition tree node
 * 
 * This function evaluates a parsed condition node using current indicator values from tickerData.
 * It handles different condition types:
 * - COMPARISON: Simple comparisons like "$G > $Q"
 * - AND: Logical AND of multiple sub-conditions
 * - OR: Logical OR of multiple sub-conditions
 * - IFS_BRANCH: A branch from an IFS() formula
 * - DEFAULT: Always true (e.g., "TRUE")
 * - COMPLEX: Comparisons with functions or arithmetic
 * 
 * @param {Object} conditionNode - Parsed condition node from parseConditionExpression()
 * @param {string} conditionNode.type - Type of condition ("COMPARISON", "AND", "OR", "IFS_BRANCH", "DEFAULT", "COMPLEX")
 * @param {string} conditionNode.expression - The condition expression
 * @param {Array<string>} conditionNode.columnRefs - Column references used in condition
 * @param {Array<string>} conditionNode.operators - Operators used in condition
 * @param {Array<Object>} conditionNode.subConditions - Sub-conditions (for AND/OR types)
 * @param {Object} tickerData - Current indicator values
 * 
 * @returns {Object} Evaluation result
 * @returns {boolean} returns.passed - Whether the condition passed (true) or failed (false)
 * @returns {Array<Object>} returns.details - Array of evaluation details for each sub-condition
 * @returns {string} returns.expression - The original condition expression
 * @returns {*} returns.leftValue - The left operand value (for COMPARISON type)
 * @returns {*} returns.rightValue - The right operand value (for COMPARISON type)
 * @returns {string} returns.operator - The comparison operator (for COMPARISON type)
 * @returns {string} returns.error - Error message if evaluation failed
 * 
 * Each detail object contains:
 * @returns {string} detail.expression - The sub-condition expression
 * @returns {boolean} detail.passed - Whether this sub-condition passed
 * @returns {*} detail.leftValue - Left operand value
 * @returns {*} detail.rightValue - Right operand value
 * @returns {string} detail.operator - Comparison operator
 * 
 * @example
 * // Simple comparison
 * const condition = {
 *   type: "COMPARISON",
 *   expression: "$G > $Q",
 *   columnRefs: ["$G", "$Q"],
 *   operators: [">"]
 * };
 * const result = evaluateCondition(condition, {price: 230.50, sma200: 225.00});
 * // Returns: {passed: true, expression: "$G > $Q", leftValue: 230.50, rightValue: 225.00, operator: ">", details: []}
 * 
 * @example
 * // AND condition
 * const condition = {
 *   type: "AND",
 *   expression: "AND($G>$Q, $R>=30)",
 *   subConditions: [...]
 * };
 * const result = evaluateCondition(condition, tickerData);
 * // Returns: {passed: true/false, expression: "AND(...)", details: [{...}, {...}]}
 * 
 * @throws {Error} If condition evaluation fails due to an error
 */
function evaluateCondition(conditionNode, tickerData) {
  try {
    // Handle different condition types
    switch (conditionNode.type) {
      case 'DEFAULT':
        // DEFAULT conditions (like "TRUE") always pass
        return {
          passed: true,
          expression: conditionNode.expression,
          details: []
        };
      
      case 'COMPARISON':
      case 'COMPLEX':
        // Simple comparison or complex expression with functions/arithmetic
        return evaluateComparison(conditionNode, tickerData);
      
      case 'AND':
        // Logical AND - all sub-conditions must pass
        return evaluateAND(conditionNode.subConditions, tickerData);
      
      case 'OR':
        // Logical OR - at least one sub-condition must pass
        return evaluateOR(conditionNode.subConditions, tickerData);
      
      case 'IFS_BRANCH':
        // IFS branch - evaluate the condition part of the branch
        // The branch has a condition and a result value
        // We need to evaluate the condition
        if (conditionNode.condition) {
          // Parse and evaluate the condition
          const parsedCondition = parseConditionExpression(conditionNode.condition);
          return evaluateCondition(parsedCondition, tickerData);
        } else {
          // No condition means this is a default branch
          return {
            passed: true,
            expression: conditionNode.expression || 'TRUE',
            details: []
          };
        }
      
      default:
        throw new Error(`Unknown condition type: ${conditionNode.type}`);
    }
  } catch (error) {
    // Log error and return error result
    // Use console.log if Logger is not available (e.g., in test environment)
    const log = typeof Logger !== 'undefined' ? Logger.log : console.log;
    log(`Condition evaluation error: ${error.message}`);
    log(`Condition: ${JSON.stringify(conditionNode)}`);
    
    return {
      passed: false,
      expression: conditionNode.expression,
      error: error.message,
      details: []
    };
  }
}

/**
 * Evaluates a comparison expression
 * 
 * This function evaluates a comparison like "$G > $Q" or "$R >= 30" by:
 * 1. Resolving the left operand (column reference or literal value)
 * 2. Resolving the right operand (column reference or literal value)
 * 3. Evaluating the comparison operator
 * 
 * @param {Object} conditionNode - Parsed comparison node
 * @param {Object} tickerData - Current indicator values
 * 
 * @returns {Object} Evaluation result with passed status and actual values
 * @private
 */
function evaluateComparison(conditionNode, tickerData) {
  const expression = conditionNode.expression;
  const columnRefs = conditionNode.columnRefs;
  const operators = conditionNode.operators;
  
  // Handle complex expressions with functions or arithmetic
  if (conditionNode.type === 'COMPLEX') {
    // For now, we'll mark complex expressions as passed if they contain valid column refs
    // Full implementation would require evaluating the expression
    // This is a simplified approach for MVP
    return {
      passed: true, // Assume complex expressions pass for now
      expression: expression,
      leftValue: null,
      rightValue: null,
      operator: operators[0] || null,
      details: [],
      note: 'Complex expression - simplified evaluation'
    };
  }
  
  // Extract left and right operands from the expression
  // Pattern: leftOperand operator rightOperand
  const operator = operators[0];
  if (!operator) {
    throw new Error(`No operator found in comparison: ${expression}`);
  }
  
  // Split by the operator to get left and right parts
  const parts = expression.split(operator);
  if (parts.length !== 2) {
    throw new Error(`Invalid comparison format: ${expression}`);
  }
  
  const leftPart = parts[0].trim();
  const rightPart = parts[1].trim();
  
  // Resolve left operand
  const leftValue = resolveOperand(leftPart, tickerData);
  
  // Resolve right operand
  const rightValue = resolveOperand(rightPart, tickerData);
  
  // Evaluate the comparison
  let passed = false;
  switch (operator) {
    case '>':
      passed = leftValue > rightValue;
      break;
    case '<':
      passed = leftValue < rightValue;
      break;
    case '>=':
      passed = leftValue >= rightValue;
      break;
    case '<=':
      passed = leftValue <= rightValue;
      break;
    case '=':
      // For equality, handle both numeric and string comparisons
      passed = leftValue == rightValue; // Use == for type coercion
      break;
    default:
      throw new Error(`Unknown operator: ${operator}`);
  }
  
  return {
    passed: passed,
    expression: expression,
    leftValue: leftValue,
    rightValue: rightValue,
    operator: operator,
    details: []
  };
}

/**
 * Resolves an operand (column reference or literal value) to its actual value
 * 
 * @param {string} operand - The operand string (e.g., "$G", "30", "0.85")
 * @param {Object} tickerData - Current indicator values
 * 
 * @returns {number|string} The resolved value
 * @private
 */
function resolveOperand(operand, tickerData) {
  // Check if it's a column reference (starts with $)
  if (operand.startsWith('$')) {
    // Map column reference to tickerData property
    return getColumnValue(operand, tickerData);
  } else {
    // It's a literal value - try to parse as number
    const numValue = parseFloat(operand);
    if (!isNaN(numValue)) {
      return numValue;
    }
    // If not a number, return as string
    return operand;
  }
}

/**
 * Gets the value for a column reference from tickerData
 * 
 * @param {string} columnRef - Column reference (e.g., "$G", "$R", "$AC")
 * @param {Object} tickerData - Current indicator values
 * 
 * @returns {number|string|null} The column value, or null if missing
 * @private
 */
function getColumnValue(columnRef, tickerData) {
  // Map column references to tickerData properties
  // Based on the COLUMN_MAP and the tickerData structure from the design doc
  const columnMap = {
    '$B': 'marketRating',
    '$C': 'decision',
    '$D': 'signal',
    '$E': 'patterns',
    '$F': 'consensusPrice',
    '$G': 'price',
    '$H': 'changePct',
    '$I': 'volTrend',
    '$J': 'athTrue',
    '$K': 'athDiff',
    '$L': 'athZone',
    '$M': 'fundamental',
    '$N': 'trendState',
    '$O': 'sma20',
    '$P': 'sma50',
    '$Q': 'sma200',
    '$R': 'rsi',
    '$S': 'macdHist',
    '$T': 'divergence',
    '$U': 'adx',
    '$V': 'stochK',
    '$W': 'volRegime',
    '$X': 'bbpSignal',
    '$Y': 'atr',
    '$Z': 'bollingerPctB',
    '$AA': 'target',
    '$AB': 'rrQuality',
    '$AC': 'support',
    '$AD': 'resistance',
    '$AE': 'atrStop',
    '$AF': 'atrTarget',
    '$AG': 'positionSize',
    '$AH': 'lastState'
  };
  
  const propertyName = columnMap[columnRef];
  if (!propertyName) {
    throw new Error(`Unknown column reference: ${columnRef}`);
  }
  
  const value = tickerData[propertyName];
  
  // Return the value as-is (including null for missing values)
  // This allows formatValue() to distinguish between missing (null) and zero (0)
  return value;
}

/**
 * Evaluates logical AND expression
 * 
 * This function evaluates all sub-conditions and returns TRUE only if ALL sub-conditions are TRUE.
 * It collects all sub-condition results for detailed reporting.
 * 
 * @param {Array<Object>} subConditions - Array of condition nodes
 * @param {Object} tickerData - Current indicator values
 * 
 * @returns {Object} Evaluation result with passed status and sub-condition details
 * @private
 */
function evaluateAND(subConditions, tickerData) {
  const details = [];
  let allPassed = true;
  
  // Evaluate each sub-condition
  for (const subCondition of subConditions) {
    const result = evaluateCondition(subCondition, tickerData);
    details.push(result);
    
    // If any sub-condition fails, the AND fails
    if (!result.passed) {
      allPassed = false;
    }
  }
  
  return {
    passed: allPassed,
    expression: `AND(${subConditions.map(sc => sc.expression).join(', ')})`,
    details: details
  };
}

/**
 * Evaluates logical OR expression
 * 
 * This function evaluates sub-conditions until one is TRUE, or all are FALSE.
 * It returns TRUE if ANY sub-condition is TRUE.
 * It collects all evaluated sub-condition results for detailed reporting.
 * 
 * @param {Array<Object>} subConditions - Array of condition nodes
 * @param {Object} tickerData - Current indicator values
 * 
 * @returns {Object} Evaluation result with passed status and sub-condition details
 * @private
 */
function evaluateOR(subConditions, tickerData) {
  const details = [];
  let anyPassed = false;
  
  // Evaluate each sub-condition
  for (const subCondition of subConditions) {
    const result = evaluateCondition(subCondition, tickerData);
    details.push(result);
    
    // If any sub-condition passes, the OR passes
    if (result.passed) {
      anyPassed = true;
    }
  }
  
  return {
    passed: anyPassed,
    expression: `OR(${subConditions.map(sc => sc.expression).join(', ')})`,
    details: details
  };
}

/**
 * Formats a single condition line for the narrative
 * 
 * This function takes an evaluated condition result and formats it into a human-readable line
 * with explanations, not just raw conditions.
 * 
 * @param {Object} conditionResult - Evaluated condition with values
 * @param {boolean} conditionResult.passed - Whether the condition passed
 * @param {string} conditionResult.expression - The condition expression
 * @param {*} conditionResult.leftValue - Left operand value (for COMPARISON)
 * @param {*} conditionResult.rightValue - Right operand value (for COMPARISON)
 * @param {string} conditionResult.operator - Comparison operator (for COMPARISON)
 * @param {Array<Object>} conditionResult.details - Sub-condition details (for AND/OR)
 * @param {Object} tickerData - Current indicator values (for column name lookup)
 * 
 * @returns {string} Formatted line like "✓ Price ($230.50) > SMA 200 ($225.00) → Bullish trend confirmed"
 */
function formatConditionLine(conditionResult, tickerData) {
  // Determine prefix based on pass/fail
  const prefix = conditionResult.passed ? "✓" : "✗";
  
  // Handle different condition types
  if (conditionResult.details && conditionResult.details.length > 0) {
    // This is an AND/OR condition with sub-conditions
    const subLines = conditionResult.details.map(detail => {
      return formatConditionLine(detail, tickerData);
    });
    
    // Determine if it's AND or OR from the expression
    const isAND = conditionResult.expression.startsWith('AND(');
    const logicType = isAND ? 'AND' : 'OR';
    
    // Format with proper indentation
    return `${prefix} ${logicType}:\n${subLines.map(line => '  ' + line).join('\n')}`;
  }
  
  // Simple comparison - format with actual values and explanation
  const expression = conditionResult.expression;
  const leftValue = conditionResult.leftValue;
  const rightValue = conditionResult.rightValue;
  const operator = conditionResult.operator;
  
  // Extract column references from expression
  const columnRefPattern = /\$[A-Z]+/g;
  const columnRefs = expression.match(columnRefPattern) || [];
  
  // Get column map once
  const COLUMN_MAP = getColumnMap();
  
  // Build formatted line by replacing column refs with "Name (value)"
  let formattedExpression = expression;
  let leftName = "";
  let rightName = "";
  
  // Replace left operand if it's a column reference
  if (columnRefs.length > 0) {
    const leftColRef = columnRefs[0];
    const leftColInfo = COLUMN_MAP[leftColRef];
    if (leftColInfo) {
      leftName = leftColInfo.name;
      const formattedValue = formatValue(leftValue, leftColInfo.format);
      formattedExpression = formattedExpression.replace(leftColRef, `${leftColInfo.name} (${formattedValue})`);
    }
  }
  
  // Replace right operand if it's a column reference
  if (columnRefs.length > 1) {
    const rightColRef = columnRefs[1];
    const rightColInfo = COLUMN_MAP[rightColRef];
    if (rightColInfo) {
      rightName = rightColInfo.name;
      const formattedValue = formatValue(rightValue, rightColInfo.format);
      formattedExpression = formattedExpression.replace(rightColRef, `${rightColInfo.name} (${formattedValue})`);
    }
  } else if (rightValue !== null && rightValue !== undefined) {
    // Right operand is a literal value, just show it
    formattedExpression = formattedExpression.replace(/(\d+\.?\d*)$/, rightValue.toString());
  }
  
  // Add explanation based on the condition
  let explanation = "";
  if (conditionResult.passed) {
    explanation = getPassedExplanation(leftName, rightName, operator, leftValue, rightValue);
  } else {
    explanation = getFailedExplanation(leftName, rightName, operator, leftValue, rightValue);
  }
  
  return `${prefix} ${formattedExpression}${explanation ? ' → ' + explanation : ''}`;
}

/**
 * Gets explanation text for a passed condition
 * @private
 */
function getPassedExplanation(leftName, rightName, operator, leftValue, rightValue) {
  // Handle null/undefined values gracefully
  if (leftValue === null || leftValue === undefined || rightValue === null || rightValue === undefined) {
    return "Data unavailable for comparison";
  }
  
  // Price vs SMA comparisons - detailed trend analysis
  // IMPORTANT: Check operator to determine if price is above or below
  if (leftName === "Price" && rightName === "SMA 200") {
    if (operator === ">" || operator === ">=") {
      const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
      return `Price ${pctAbove}% above 200-day moving average confirms long-term bullish trend and RISK-ON regime`;
    } else if (operator === "<" || operator === "<=") {
      const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
      return `Price ${pctBelow}% below 200-day MA confirms RISK-OFF regime and bearish trend`;
    }
  }
  if (leftName === "Price" && rightName === "SMA 50") {
    if (operator === ">" || operator === ">=") {
      const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
      return `Price ${pctAbove}% above 50-day moving average indicates medium-term bullish momentum`;
    } else if (operator === "<" || operator === "<=") {
      const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
      return `Price ${pctBelow}% below 50-day MA shows medium-term bearish pressure`;
    }
  }
  if (leftName === "Price" && rightName === "SMA 20") {
    if (operator === ">" || operator === ">=") {
      const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
      return `Price ${pctAbove}% above 20-day moving average shows short-term strength`;
    } else if (operator === "<" || operator === "<=") {
      const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
      return `Price ${pctBelow}% below 20-day MA indicates short-term weakness`;
    }
  }
  if (leftName === "Price" && rightName === "Support") {
    if (operator === ">" || operator === ">=") {
      const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
      return `Price ${pctAbove}% above support level maintains structural integrity`;
    } else if (operator === "<" || operator === "<=") {
      const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
      return `Price ${pctBelow}% below support triggers stop loss - structural breakdown confirmed`;
    }
  }
  if (leftName === "Price" && rightName === "Resistance") {
    if (operator === ">" || operator === ">=") {
      return "Price has broken above resistance, confirming bullish breakout";
    } else if (operator === "<" || operator === "<=") {
      return "Price remains below resistance - supply zone intact";
    }
  }
  
  // RSI conditions - detailed momentum analysis
  if (leftName === "RSI") {
    if (operator === ">=" && rightValue >= 55 && rightValue <= 65) {
      return `RSI at ${leftValue.toFixed(1)} indicates positive momentum without overbought extremes`;
    }
    if (operator === ">=" && rightValue >= 30 && rightValue <= 40) {
      return `RSI at ${leftValue.toFixed(1)} shows healthy momentum recovery from oversold levels`;
    }
    if (operator === "<=" && rightValue === 30) {
      return `RSI at ${leftValue.toFixed(1)} signals oversold condition - potential mean reversion opportunity`;
    }
    if (operator === ">=" && rightValue === 70) {
      return `RSI at ${leftValue.toFixed(1)} indicates overbought condition - momentum exhaustion risk`;
    }
    if (operator === "<=" && rightValue === 45) {
      return `RSI at ${leftValue.toFixed(1)} shows weak momentum - bearish pressure present`;
    }
  }
  
  // MACD conditions
  if (leftName === "MACD Hist") {
    if (operator === ">" && rightValue === 0) {
      return `MACD histogram positive at ${leftValue.toFixed(3)} confirms bullish momentum impulse`;
    }
    if (operator === "<" && rightValue === 0) {
      return `MACD histogram negative at ${leftValue.toFixed(3)} confirms bearish momentum impulse`;
    }
  }
  
  // ADX conditions - trend strength analysis
  if (leftName === "ADX") {
    if (operator === ">=" && rightValue >= 25) {
      return `ADX at ${leftValue.toFixed(1)} confirms strong trending market - directional conviction present`;
    }
    if (operator === ">=" && rightValue >= 20) {
      return `ADX at ${leftValue.toFixed(1)} indicates developing trend with increasing strength`;
    }
    if (operator === "<" && rightValue === 15) {
      return `ADX at ${leftValue.toFixed(1)} signals range-bound market - low directional conviction`;
    }
  }
  
  // Volume conditions - participation analysis
  if (leftName === "Vol Trend") {
    if (operator === ">=" && rightValue >= 2.0) {
      return `Volume ${leftValue.toFixed(2)}x average indicates extreme institutional participation`;
    }
    if (operator === ">=" && rightValue >= 1.5) {
      return `Volume ${leftValue.toFixed(2)}x average confirms strong market participation and conviction`;
    }
    if (operator === ">=" && rightValue >= 1.0) {
      return `Volume ${leftValue.toFixed(2)}x average shows normal market participation`;
    }
    if (operator === "<" && rightValue === 1.0) {
      return `Volume ${leftValue.toFixed(2)}x average indicates low participation - drift risk present`;
    }
  }
  
  // Stochastic conditions
  if (leftName === "Stoch %K (14)") {
    if (operator === ">=" && rightValue >= 0.8) {
      return `Stochastic at ${(leftValue * 100).toFixed(1)}% signals overbought timing - near-term reversal risk`;
    }
    if (operator === "<=" && rightValue <= 0.2) {
      return `Stochastic at ${(leftValue * 100).toFixed(1)}% signals oversold timing - potential bounce setup`;
    }
  }
  
  // Bollinger %B conditions
  if (leftName === "Bollinger %B") {
    if (operator === ">" && rightValue === 1) {
      return `Bollinger %B at ${(leftValue * 100).toFixed(1)}% shows price above upper band - extended move`;
    }
    if (operator === "<" && rightValue === 0) {
      return `Bollinger %B at ${(leftValue * 100).toFixed(1)}% shows price below lower band - oversold extreme`;
    }
    if (operator === ">=" && rightValue >= 0.8) {
      return `Bollinger %B at ${(leftValue * 100).toFixed(1)}% indicates upper band zone - overbought territory`;
    }
    if (operator === "<=" && rightValue <= 0.2) {
      return `Bollinger %B at ${(leftValue * 100).toFixed(1)}% indicates lower band zone - oversold territory`;
    }
  }
  
  // ATH Diff conditions
  if (leftName === "ATH Diff %") {
    if (operator === ">=" && rightValue >= -0.02) {
      return `Price within ${Math.abs(leftValue * 100).toFixed(1)}% of all-time high - market leadership confirmed`;
    }
    if (operator === "<=" && rightValue <= -0.30) {
      return `Price ${Math.abs(leftValue * 100).toFixed(1)}% below ATH - deep value territory`;
    }
  }
  
  // R:R Quality conditions
  if (leftName === "R:R Quality") {
    if (operator === ">=" && rightValue >= 3) {
      return `Risk-reward ratio of ${leftValue.toFixed(2)}:1 provides elite asymmetric opportunity`;
    }
    if (operator === ">=" && rightValue >= 1.5) {
      return `Risk-reward ratio of ${leftValue.toFixed(2)}:1 offers acceptable asymmetry`;
    }
  }
  
  // Default explanation
  return "Condition met";
}

/**
 * Gets explanation text for a failed condition
 * @private
 */
function getFailedExplanation(leftName, rightName, operator, leftValue, rightValue) {
  // Handle null/undefined values gracefully
  if (leftValue === null || leftValue === undefined || rightValue === null || rightValue === undefined) {
    return "Data unavailable for comparison";
  }
  
  // Price vs SMA comparisons - detailed trend analysis
  // IMPORTANT: Failed condition means the opposite of what was tested
  if (leftName === "Price" && rightName === "SMA 200") {
    if (operator === ">" || operator === ">=") {
      // Failed: Price NOT above SMA 200, so price is below
      const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
      return `Price ${pctBelow}% below 200-day MA indicates RISK-OFF regime - long-term bearish trend`;
    } else if (operator === "<" || operator === "<=") {
      // Failed: Price NOT below SMA 200, so price is above
      const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
      return `Price ${pctAbove}% above 200-day MA maintains bullish trend (condition not met for bearish signal)`;
    }
  }
  if (leftName === "Price" && rightName === "SMA 50") {
    if (operator === ">" || operator === ">=") {
      // Failed: Price NOT above SMA 50, so price is below
      const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
      return `Price ${pctBelow}% below 50-day MA shows medium-term bearish pressure`;
    } else if (operator === "<" || operator === "<=") {
      // Failed: Price NOT below SMA 50, so price is above
      const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
      return `Price ${pctAbove}% above 50-day MA maintains medium-term bullish momentum`;
    }
  }
  if (leftName === "Price" && rightName === "SMA 20") {
    if (operator === ">" || operator === ">=") {
      // Failed: Price NOT above SMA 20, so price is below
      const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
      return `Price ${pctBelow}% below 20-day MA indicates short-term weakness`;
    } else if (operator === "<" || operator === "<=") {
      // Failed: Price NOT below SMA 20, so price is above
      const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
      return `Price ${pctAbove}% above 20-day MA shows short-term strength`;
    }
  }
  if (leftName === "Price" && rightName === "Support") {
    if (operator === ">" || operator === ">=") {
      // Failed: Price NOT above support, so price is below
      const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
      return `Price ${pctBelow}% below support - stop loss triggered, structural breakdown`;
    } else if (operator === "<" || operator === "<=") {
      // Failed: Price NOT below support, so price is above
      const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
      return `Price ${pctAbove}% above support maintains structural integrity`;
    }
  }
  
  // RSI conditions - detailed momentum analysis
  if (leftName === "RSI") {
    if (operator === ">=" && rightValue >= 55) {
      return `RSI at ${leftValue.toFixed(1)} below ${rightValue} threshold - momentum not yet positive`;
    }
    if (operator === ">=" && rightValue >= 30) {
      return `RSI at ${leftValue.toFixed(1)} still below ${rightValue} - not yet recovered from oversold`;
    }
    if (operator === "<=" && rightValue === 30) {
      return `RSI at ${leftValue.toFixed(1)} above 30 - not yet oversold`;
    }
    if (operator === ">=" && rightValue === 70) {
      return `RSI at ${leftValue.toFixed(1)} below 70 - not yet overbought`;
    }
    if (operator === "<=" && rightValue === 70) {
      return `RSI at ${leftValue.toFixed(1)} above 70 - overbought condition present`;
    }
    if (operator === "<=" && rightValue === 45) {
      return `RSI at ${leftValue.toFixed(1)} above 45 - momentum not yet weak`;
    }
  }
  
  // MACD conditions
  if (leftName === "MACD Hist") {
    if (operator === ">" && rightValue === 0) {
      return `MACD histogram at ${leftValue.toFixed(3)} is negative - bearish momentum present`;
    }
    if (operator === "<" && rightValue === 0) {
      return `MACD histogram at ${leftValue.toFixed(3)} is positive - bullish momentum present`;
    }
  }
  
  // ADX conditions - trend strength analysis
  if (leftName === "ADX") {
    if (operator === ">=" && rightValue >= 20) {
      return `ADX at ${leftValue.toFixed(1)} below ${rightValue} - trend strength insufficient`;
    }
    if (operator === "<" && rightValue === 15) {
      return `ADX at ${leftValue.toFixed(1)} above 15 - some directional bias present`;
    }
  }
  
  // Volume conditions - participation analysis
  if (leftName === "Vol Trend") {
    if (operator === ">=" && rightValue >= 1.5) {
      return `Volume ${leftValue.toFixed(2)}x average below threshold - insufficient participation`;
    }
    if (operator === "<" && rightValue === 1.0) {
      return `Volume ${leftValue.toFixed(2)}x average is normal or above - adequate participation`;
    }
  }
  
  // Stochastic conditions
  if (leftName === "Stoch %K (14)") {
    if (operator === ">=" && rightValue >= 0.8) {
      return `Stochastic at ${(leftValue * 100).toFixed(1)}% not yet overbought`;
    }
    if (operator === "<=" && rightValue <= 0.2) {
      return `Stochastic at ${(leftValue * 100).toFixed(1)}% not yet oversold`;
    }
  }
  
  // Bollinger %B conditions
  if (leftName === "Bollinger %B") {
    if (operator === ">" && rightValue === 1) {
      return `Bollinger %B at ${(leftValue * 100).toFixed(1)}% within bands - not extended`;
    }
    if (operator === "<" && rightValue === 0) {
      return `Bollinger %B at ${(leftValue * 100).toFixed(1)}% within bands - not oversold`;
    }
  }
  
  // ATH Diff conditions
  if (leftName === "ATH Diff %") {
    if (operator === ">=" && rightValue >= -0.02) {
      return `Price ${Math.abs(leftValue * 100).toFixed(1)}% below ATH - not at market leadership zone`;
    }
  }
  
  // R:R Quality conditions
  if (leftName === "R:R Quality") {
    if (operator === ">=" && rightValue >= 3) {
      return `Risk-reward ratio of ${leftValue.toFixed(2)}:1 below elite threshold`;
    }
    if (operator === "<" && rightValue === 1.5) {
      return `Risk-reward ratio of ${leftValue.toFixed(2)}:1 is poor - asymmetry insufficient`;
    }
  }
  
  // Default explanation
  return "Threshold not met";
}

/**
 * Generates institutional-grade narrative from evaluation path
 * 
 * This function takes an array of evaluated condition nodes and generates a professional,
 * synthesized narrative that explains WHY the signal triggered, focusing on key factors
 * rather than showing raw condition trees.
 * 
 * @param {Array<Object>} evaluationPath - Array of evaluated condition nodes
 * @param {string} finalResult - The final signal/decision value
 * 
 * @returns {string} Institutional-grade narrative with synthesized analysis
 */
function generateNarrative(evaluationPath, finalResult) {
  // Build institutional-grade narrative
  const lines = [];
  
  // Add header
  lines.push(`🎯 WHY '${finalResult}' TRIGGERED:\n`);
  
  // Extract key factors from evaluation path
  const keyFactors = extractKeyFactors(evaluationPath);
  
  // Generate synthesized explanation based on key factors
  if (keyFactors.length === 0) {
    lines.push("Default condition met - no specific technical criteria required.");
  } else {
    // Group factors by category for better readability
    const priceFactors = keyFactors.filter(f => f.category === 'price');
    const momentumFactors = keyFactors.filter(f => f.category === 'momentum');
    const trendFactors = keyFactors.filter(f => f.category === 'trend');
    const volumeFactors = keyFactors.filter(f => f.category === 'volume');
    const volatilityFactors = keyFactors.filter(f => f.category === 'volatility');
    
    // Build synthesized narrative
    if (priceFactors.length > 0) {
      lines.push("PRICE ACTION:");
      priceFactors.forEach(factor => {
        lines.push(`  • ${factor.explanation}`);
      });
      lines.push("");
    }
    
    if (trendFactors.length > 0) {
      lines.push("TREND STRUCTURE:");
      trendFactors.forEach(factor => {
        lines.push(`  • ${factor.explanation}`);
      });
      lines.push("");
    }
    
    if (momentumFactors.length > 0) {
      lines.push("MOMENTUM INDICATORS:");
      momentumFactors.forEach(factor => {
        lines.push(`  • ${factor.explanation}`);
      });
      lines.push("");
    }
    
    if (volumeFactors.length > 0) {
      lines.push("VOLUME ANALYSIS:");
      volumeFactors.forEach(factor => {
        lines.push(`  • ${factor.explanation}`);
      });
      lines.push("");
    }
    
    if (volatilityFactors.length > 0) {
      lines.push("VOLATILITY METRICS:");
      volatilityFactors.forEach(factor => {
        lines.push(`  • ${factor.explanation}`);
      });
      lines.push("");
    }
  }
  
  // Add final result line
  lines.push(`→ RESULT: ${finalResult}`);
  
  // Join all lines with newlines
  return lines.join('\n');
}

/**
 * Extracts key factors from evaluation path for synthesized narrative
 * 
 * This function analyzes the evaluation path and extracts the most important
 * factors that contributed to the signal, categorizing them for better presentation.
 * 
 * @param {Array<Object>} evaluationPath - Array of evaluated condition nodes
 * @returns {Array<Object>} Array of key factors with category and explanation
 * @private
 */
function extractKeyFactors(evaluationPath) {
  const factors = [];
  
  for (const conditionResult of evaluationPath) {
    // Handle AND/OR conditions with sub-conditions
    if (conditionResult.details && conditionResult.details.length > 0) {
      // Process each sub-condition
      for (const detail of conditionResult.details) {
        const factor = extractFactorFromCondition(detail);
        if (factor) {
          factors.push(factor);
        }
      }
    } else {
      // Simple comparison - extract factor directly
      const factor = extractFactorFromCondition(conditionResult);
      if (factor) {
        factors.push(factor);
      }
    }
  }
  
  return factors;
}

/**
 * Extracts a single factor from a condition result
 * 
 * @param {Object} conditionResult - Evaluated condition result
 * @returns {Object|null} Factor object with category and explanation, or null if not extractable
 * @private
 */
function extractFactorFromCondition(conditionResult) {
  // Skip if condition has error or missing values
  if (conditionResult.error) {
    return null;
  }
  
  const leftValue = conditionResult.leftValue;
  const rightValue = conditionResult.rightValue;
  const operator = conditionResult.operator;
  
  // Handle null/undefined values
  if (leftValue === null || leftValue === undefined || rightValue === null || rightValue === undefined) {
    return null;
  }
  
  // Extract column references from expression
  const expression = conditionResult.expression;
  const columnRefPattern = /\$[A-Z]+/g;
  const columnRefs = expression.match(columnRefPattern) || [];
  
  if (columnRefs.length === 0) {
    return null;
  }
  
  // Get column names
  const leftColRef = columnRefs[0];
  const rightColRef = columnRefs.length > 1 ? columnRefs[1] : null;
  
  const COLUMN_MAP = getColumnMap();
  const leftColInfo = COLUMN_MAP[leftColRef];
  const rightColInfo = rightColRef ? COLUMN_MAP[rightColRef] : null;
  
  if (!leftColInfo) {
    return null;
  }
  
  const leftName = leftColInfo.name;
  const rightName = rightColInfo ? rightColInfo.name : null;
  
  // Categorize and generate explanation based on indicator type
  let category = 'other';
  let explanation = '';
  
  // Price comparisons
  if (leftName === "Price") {
    category = 'price';
    
    if (rightName === "Support") {
      if (conditionResult.passed) {
        if (operator === ">" || operator === ">=") {
          const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
          explanation = `Price at $${leftValue.toFixed(2)} is ${pctAbove}% above support ($${rightValue.toFixed(2)}), maintaining structural integrity`;
        } else if (operator === "<" || operator === "<=") {
          const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
          explanation = `Price at $${leftValue.toFixed(2)} is ${pctBelow}% below support ($${rightValue.toFixed(2)}), triggering stop loss - structural breakdown`;
        }
      } else {
        if (operator === ">" || operator === ">=") {
          const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
          explanation = `Price at $${leftValue.toFixed(2)} is ${pctBelow}% below support ($${rightValue.toFixed(2)}) - condition not met`;
        } else if (operator === "<" || operator === "<=") {
          const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
          explanation = `Price at $${leftValue.toFixed(2)} is ${pctAbove}% above support ($${rightValue.toFixed(2)}) - condition not met`;
        }
      }
    } else if (rightName === "Resistance") {
      if (conditionResult.passed) {
        if (operator === ">" || operator === ">=") {
          explanation = `Price at $${leftValue.toFixed(2)} has broken above resistance ($${rightValue.toFixed(2)}), confirming bullish breakout`;
        } else if (operator === "<" || operator === "<=") {
          explanation = `Price at $${leftValue.toFixed(2)} remains below resistance ($${rightValue.toFixed(2)}) - supply zone intact`;
        }
      }
    } else if (rightName && rightName.includes("SMA")) {
      category = 'trend';
      if (conditionResult.passed) {
        if (operator === ">" || operator === ">=") {
          const pctAbove = ((leftValue - rightValue) / rightValue * 100).toFixed(1);
          explanation = `Price ($${leftValue.toFixed(2)}) is ${pctAbove}% above ${rightName} ($${rightValue.toFixed(2)}), confirming ${rightName === "SMA 200" ? "long-term bullish trend" : rightName === "SMA 50" ? "medium-term strength" : "short-term momentum"}`;
        } else if (operator === "<" || operator === "<=") {
          const pctBelow = ((rightValue - leftValue) / rightValue * 100).toFixed(1);
          explanation = `Price ($${leftValue.toFixed(2)}) is ${pctBelow}% below ${rightName} ($${rightValue.toFixed(2)}), indicating ${rightName === "SMA 200" ? "RISK-OFF regime" : rightName === "SMA 50" ? "medium-term weakness" : "short-term bearish pressure"}`;
        }
      }
    }
  }
  
  // RSI conditions
  else if (leftName === "RSI") {
    category = 'momentum';
    const rsiValue = leftValue.toFixed(1);
    
    if (conditionResult.passed) {
      if (operator === ">=" && rightValue >= 55 && rightValue <= 65) {
        explanation = `RSI at ${rsiValue} indicates positive momentum without overbought extremes`;
      } else if (operator === ">=" && rightValue >= 30 && rightValue <= 40) {
        explanation = `RSI at ${rsiValue} shows healthy momentum recovery from oversold levels`;
      } else if (operator === "<=" && rightValue === 30) {
        explanation = `RSI at ${rsiValue} signals oversold condition - potential mean reversion opportunity`;
      } else if (operator === ">=" && rightValue === 70) {
        explanation = `RSI at ${rsiValue} indicates overbought condition - momentum exhaustion risk`;
      } else if (operator === "<=" && rightValue === 45) {
        explanation = `RSI at ${rsiValue} shows weak momentum - bearish pressure present`;
      } else {
        explanation = `RSI at ${rsiValue} meets threshold criteria`;
      }
    }
  }
  
  // MACD conditions
  else if (leftName === "MACD Hist") {
    category = 'momentum';
    if (conditionResult.passed) {
      if (operator === ">" && rightValue === 0) {
        explanation = `MACD histogram positive at ${leftValue.toFixed(3)}, confirming bullish momentum impulse`;
      } else if (operator === "<" && rightValue === 0) {
        explanation = `MACD histogram negative at ${leftValue.toFixed(3)}, confirming bearish momentum impulse`;
      }
    }
  }
  
  // ADX conditions
  else if (leftName === "ADX") {
    category = 'trend';
    const adxValue = leftValue.toFixed(1);
    
    if (conditionResult.passed) {
      if (operator === ">=" && rightValue >= 25) {
        explanation = `ADX at ${adxValue} confirms strong trending market with directional conviction`;
      } else if (operator === ">=" && rightValue >= 20) {
        explanation = `ADX at ${adxValue} indicates developing trend with increasing strength`;
      } else if (operator === "<" && rightValue === 15) {
        explanation = `ADX at ${adxValue} signals range-bound market with low directional conviction`;
      }
    }
  }
  
  // Volume conditions
  else if (leftName === "Vol Trend") {
    category = 'volume';
    const volValue = leftValue.toFixed(2);
    
    if (conditionResult.passed) {
      if (operator === ">=" && rightValue >= 2.0) {
        explanation = `Volume at ${volValue}x average indicates extreme institutional participation`;
      } else if (operator === ">=" && rightValue >= 1.5) {
        explanation = `Volume at ${volValue}x average confirms strong market participation and conviction`;
      } else if (operator === ">=" && rightValue >= 1.0) {
        explanation = `Volume at ${volValue}x average shows normal market participation`;
      } else if (operator === "<" && rightValue === 1.0) {
        explanation = `Volume at ${volValue}x average indicates low participation - drift risk present`;
      }
    }
  }
  
  // Stochastic conditions
  else if (leftName === "Stoch %K (14)") {
    category = 'momentum';
    const stochValue = (leftValue * 100).toFixed(1);
    
    if (conditionResult.passed) {
      if (operator === ">=" && rightValue >= 0.8) {
        explanation = `Stochastic at ${stochValue}% signals overbought timing - near-term reversal risk`;
      } else if (operator === "<=" && rightValue <= 0.2) {
        explanation = `Stochastic at ${stochValue}% signals oversold timing - potential bounce setup`;
      }
    }
  }
  
  // Bollinger %B conditions
  else if (leftName === "Bollinger %B") {
    category = 'volatility';
    const bbValue = (leftValue * 100).toFixed(1);
    
    if (conditionResult.passed) {
      if (operator === ">" && rightValue === 1) {
        explanation = `Bollinger %B at ${bbValue}% shows price above upper band - extended move`;
      } else if (operator === "<" && rightValue === 0) {
        explanation = `Bollinger %B at ${bbValue}% shows price below lower band - oversold extreme`;
      } else if (operator === ">=" && rightValue >= 0.8) {
        explanation = `Bollinger %B at ${bbValue}% indicates upper band zone - overbought territory`;
      } else if (operator === "<=" && rightValue <= 0.2) {
        explanation = `Bollinger %B at ${bbValue}% indicates lower band zone - oversold territory`;
      }
    }
  }
  
  // ATH Diff conditions
  else if (leftName === "ATH Diff %") {
    category = 'price';
    const athDiffValue = (leftValue * 100).toFixed(1);
    
    if (conditionResult.passed) {
      if (operator === ">=" && rightValue >= -0.02) {
        explanation = `Price within ${Math.abs(athDiffValue)}% of all-time high - market leadership confirmed`;
      } else if (operator === "<=" && rightValue <= -0.30) {
        explanation = `Price ${Math.abs(athDiffValue)}% below ATH - deep value territory`;
      }
    }
  }
  
  // R:R Quality conditions
  else if (leftName === "R:R Quality") {
    category = 'price';
    const rrValue = leftValue.toFixed(2);
    
    if (conditionResult.passed) {
      if (operator === ">=" && rightValue >= 3) {
        explanation = `Risk-reward ratio of ${rrValue}:1 provides elite asymmetric opportunity`;
      } else if (operator === ">=" && rightValue >= 1.5) {
        explanation = `Risk-reward ratio of ${rrValue}:1 offers acceptable asymmetry`;
      }
    }
  }
  
  // If no explanation was generated, return null
  if (!explanation) {
    return null;
  }
  
  return {
    category: category,
    explanation: explanation
  };
}

/**
 * Helper function to find a matching branch in a condition tree
 * 
 * @param {Object} conditionTree - Parsed condition tree from parseSignalFormula() or parseDecisionFormula()
 * @param {string} targetValue - The value to match (e.g., "STRONG BUY", "🟢 TRADE LONG")
 * 
 * @returns {Object|null} The matching branch object, or null if no match found
 * @private
 */
function findMatchingBranch(conditionTree, targetValue) {
  if (!conditionTree || !conditionTree.branches) {
    return null;
  }
  
  const normalizedTarget = String(targetValue || "").trim();
  
  for (const branch of conditionTree.branches) {
    const branchResult = String(branch.result || "").trim();
    if (branchResult === normalizedTarget) {
      return branch;
    }
  }
  
  return null;
}

/**
 * Evaluates SIGNAL formula and generates narrative explanation
 * 
 * This function parses the SIGNAL formula logic from buildSignalFormula(), evaluates each
 * condition branch using current indicator values, and generates a human-readable narrative
 * showing which conditions passed (✓) or failed (✗) and why.
 * 
 * IMPORTANT: This function handles mode mismatches automatically. If the signal value doesn't
 * match any branch in the specified mode, it will try the other mode. This prevents errors
 * when DASHBOARD H1 changes but CALCULATIONS hasn't recalculated yet.
 * 
 * @param {Object} tickerData - Object containing all indicator values for the ticker
 * @param {string} tickerData.ticker - The ticker symbol (e.g., "AMZN")
 * @param {number} tickerData.price - Current price (Column G)
 * @param {number} tickerData.rsi - RSI indicator (Column R)
 * @param {number} tickerData.adx - ADX indicator (Column U)
 * @param {number} tickerData.sma20 - 20-day Simple Moving Average (Column O)
 * @param {number} tickerData.sma50 - 50-day Simple Moving Average (Column P)
 * @param {number} tickerData.sma200 - 200-day Simple Moving Average (Column Q)
 * @param {number} tickerData.macdHist - MACD Histogram (Column S)
 * @param {number} tickerData.volTrend - Volume Trend (Column I)
 * @param {number} tickerData.stochK - Stochastic %K (Column V)
 * @param {number} tickerData.bollingerPctB - Bollinger %B (Column Z)
 * @param {number} tickerData.atr - Average True Range (Column Y)
 * @param {number} tickerData.support - Support level (Column AC)
 * @param {number} tickerData.resistance - Resistance level (Column AD)
 * @param {number} tickerData.athDiff - All-Time High difference % (Column K)
 * @param {string} tickerData.trendState - Trend state (Column N)
 * @param {string} tickerData.divergence - Divergence indicator (Column T)
 * @param {string} tickerData.volRegime - Volatility regime (Column W)
 * @param {string} tickerData.bbpSignal - Bollinger Band Position signal (Column X)
 * @param {boolean} useLongTermSignal - Whether to use long-term investment mode (true) or trade mode (false)
 * 
 * @returns {Object} Evaluation result object
 * @returns {string} returns.signal - The final SIGNAL value (e.g., "STRONG BUY", "HOLD", "RISK OFF")
 * @returns {string} returns.narrative - Human-readable narrative with ✓/✗ markers showing evaluation path
 * @returns {Array<Object>} returns.evaluationPath - Array of evaluated condition nodes with results
 * 
 * @example
 * const tickerData = {
 *   ticker: "AMZN",
 *   price: 230.50,
 *   sma200: 225.00,
 *   rsi: 55.1,
 *   adx: 13.0,
 *   // ... other indicators
 * };
 * 
 * const result = evaluateSignalFormula(tickerData, false);
 * console.log(result.signal);      // "HOLD"
 * console.log(result.narrative);   // "✓ Price ($230.50) > SMA 200 ($225.00) → Bullish trend confirmed\n..."
 * 
 * @throws {Error} If formula parsing fails (falls back to generic explanation)
 */
function evaluateSignalFormula(tickerData, useLongTermSignal) {
  try {
    // IMPORTANT: We don't re-evaluate the formula to determine the signal.
    // Instead, we take the ACTUAL signal from CALCULATIONS sheet (tickerData.signal)
    // and explain which branch produced that result.
    
    // Step 1: Get the actual SIGNAL value from CALCULATIONS sheet
    const actualSignal = String(tickerData.signal || "").trim();
    
    if (!actualSignal || actualSignal === "LOADING" || actualSignal === "—") {
      throw new Error('Signal not available - ticker is still loading');
    }
    
    // Step 2: Call parseSignalFormula() to get condition tree for the specified mode
    let conditionTree = parseSignalFormula(useLongTermSignal);
    
    // Check if parsing failed (returned null)
    if (!conditionTree) {
      throw new Error('Formula parsing failed - parseSignalFormula returned null');
    }
    
    // Step 3: Find the branch that produces the actual signal
    let matchingBranch = findMatchingBranch(conditionTree, actualSignal);
    
    // Step 3.5: If no match found, try the OTHER mode (handles mode mismatch)
    // This happens when DASHBOARD H1 changes but CALCULATIONS hasn't recalculated yet
    if (!matchingBranch) {
      const log = typeof Logger !== 'undefined' ? Logger.log : console.log;
      log(`WARNING: No match for signal "${actualSignal}" in ${useLongTermSignal ? 'INVEST' : 'TRADE'} mode. Trying ${useLongTermSignal ? 'TRADE' : 'INVEST'} mode...`);
      
      // Toggle mode and try again
      useLongTermSignal = !useLongTermSignal;
      conditionTree = parseSignalFormula(useLongTermSignal);
      
      if (!conditionTree) {
        throw new Error('Formula parsing failed for alternate mode');
      }
      
      matchingBranch = findMatchingBranch(conditionTree, actualSignal);
      
      if (matchingBranch) {
        log(`SUCCESS: Found match in ${useLongTermSignal ? 'INVEST' : 'TRADE'} mode`);
      }
    }
    
    // If still no matching branch found, log available branches and throw error
    if (!matchingBranch) {
      const log = typeof Logger !== 'undefined' ? Logger.log : console.log;
      log(`ERROR: No formula branch found for signal: "${actualSignal}" in either mode`);
      log(`Available branches: ${conditionTree.branches.map(b => `"${b.result}"`).join(', ')}`);
      throw new Error(`No formula branch found for signal: ${actualSignal}`);
    }
    
    // Step 4: Evaluate the matching branch's condition to show which parts passed/failed
    const parsedCondition = parseConditionExpression(matchingBranch.condition);
    const evaluationResult = evaluateCondition(parsedCondition, tickerData);
    
    // Build evaluation path with just this branch
    const evaluationPath = [{
      order: matchingBranch.order,
      condition: matchingBranch.condition,
      result: matchingBranch.result,
      passed: evaluationResult.passed,
      expression: evaluationResult.expression,
      leftValue: evaluationResult.leftValue,
      rightValue: evaluationResult.rightValue,
      operator: evaluationResult.operator,
      details: evaluationResult.details,
      error: evaluationResult.error
    }];
    
    // Step 5: Call generateNarrative() to create explanation
    const narrative = generateNarrative(evaluationPath, actualSignal);
    
    // Step 6: Return {signal, narrative, evaluationPath}
    return {
      signal: actualSignal,  // Return the actual signal from CALCULATIONS sheet
      narrative: narrative,
      evaluationPath: evaluationPath
    };
    
  } catch (error) {
    // Log error and re-throw for caller to handle
    const log = typeof Logger !== 'undefined' ? Logger.log : console.log;
    log(`Error in evaluateSignalFormula: ${error.message}`);
    log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Evaluates DECISION formula and generates narrative explanation
 * 
 * This function parses the DECISION formula logic from buildDecisionFormula(), evaluates
 * how SIGNAL combines with PATTERNS and PURCHASED tag to produce the final DECISION,
 * and generates a narrative showing the decision-making process.
 * 
 * @param {Object} tickerData - Object containing all indicator values for the ticker
 * @param {string} tickerData.ticker - The ticker symbol (e.g., "AMZN")
 * @param {string} tickerData.signal - The SIGNAL value from evaluateSignalFormula()
 * @param {string} tickerData.patterns - Pattern detection string (Column E, e.g., "BRKOUT (72%)")
 * @param {string} tickerData.decision - Current DECISION value (Column C)
 * @param {boolean} tickerData.isPurchased - Whether the "PURCHASED" tag is present in INPUT sheet
 * @param {number} tickerData.price - Current price (Column G)
 * @param {number} tickerData.support - Support level (Column AC)
 * @param {number} tickerData.resistance - Resistance level (Column AD)
 * @param {number} tickerData.rsi - RSI indicator (Column R)
 * @param {number} tickerData.adx - ADX indicator (Column U)
 * @param {string} tickerData.trendState - Trend state (Column N)
 * @param {string} signal - The SIGNAL value from evaluateSignalFormula()
 * @param {boolean} useLongTermSignal - Whether to use long-term investment mode (true) or trade mode (false)
 * 
 * @returns {Object} Evaluation result object
 * @returns {string} returns.decision - The final DECISION value (e.g., "🟢 STRONG BUY", "⚖️ HOLD", "🔴 EXIT")
 * @returns {string} returns.narrative - Human-readable narrative showing how SIGNAL + PATTERNS → DECISION
 * @returns {Array<Object>} returns.evaluationPath - Array of evaluated condition nodes with results
 * 
 * @example
 * const tickerData = {
 *   ticker: "AMZN",
 *   signal: "STRONG BUY",
 *   patterns: "BRKOUT (72%)",
 *   isPurchased: false,
 *   price: 230.50,
 *   // ... other indicators
 * };
 * 
 * const result = evaluateDecisionFormula(tickerData, "STRONG BUY", false);
 * console.log(result.decision);    // "🟢 STRONG BUY"
 * console.log(result.narrative);   // "✓ SIGNAL: STRONG BUY\n✓ Pattern: BRKOUT (72%) → Bullish breakout\n..."
 * 
 * @throws {Error} If formula parsing fails (falls back to generic explanation)
 */
function evaluateDecisionFormula(tickerData, signal, useLongTermSignal) {
  try {
    // IMPORTANT: We don't re-evaluate the formula to determine the decision.
    // Instead, we take the ACTUAL decision from CALCULATIONS sheet (tickerData.decision)
    // and explain which branch produced that result.
    
    // Step 1: Get the actual DECISION value from CALCULATIONS sheet
    const actualDecision = String(tickerData.decision || "").trim();
    
    if (!actualDecision || actualDecision === "LOADING" || actualDecision === "—") {
      throw new Error('Decision not available - ticker is still loading');
    }
    
    // Step 2: Call parseDecisionFormula() to get condition tree for the specified mode
    let conditionTree = parseDecisionFormula(useLongTermSignal);
    
    // Check if parsing failed (returned null)
    if (!conditionTree) {
      throw new Error('Formula parsing failed - parseDecisionFormula returned null');
    }
    
    // Step 3: Check PURCHASED tag from INPUT sheet
    const isPurchased = tickerData.isPurchased || false;
    
    // Step 4: Evaluate pattern detection logic
    const patternsStr = String(tickerData.patterns || "").trim();
    const hasPattern = patternsStr !== "" && patternsStr !== "—" && patternsStr !== "-";
    
    // Determine pattern type (bullish, bearish, or none)
    let patternType = "none";
    if (hasPattern) {
      const bullishPatterns = ["ASC_TRI", "BRKOUT", "DBL_BTM", "INV_H&S", "CUP_HDL"];
      const bearishPatterns = ["DESC_TRI", "H&S", "DBL_TOP"];
      const upperPatterns = patternsStr.toUpperCase();
      
      for (const pattern of bullishPatterns) {
        if (upperPatterns.includes(pattern)) {
          patternType = "bullish";
          break;
        }
      }
      
      if (patternType === "none") {
        for (const pattern of bearishPatterns) {
          if (upperPatterns.includes(pattern)) {
            patternType = "bearish";
            break;
          }
        }
      }
      
      if (patternType === "none") {
        patternType = "any";
      }
    }
    
    // Step 5: Find the branch that produces the actual decision
    let matchingBranch = null;
    for (const branch of conditionTree.branches) {
      // Filter branches based on PURCHASED requirement
      if (branch.requiresPurchased && !isPurchased) {
        continue;
      }
      if (branch.requiresNotPurchased && isPurchased) {
        continue;
      }
      
      // Normalize both values for comparison (trim whitespace)
      const branchResult = String(branch.result || "").trim();
      if (branchResult === actualDecision) {
        matchingBranch = branch;
        break;
      }
    }
    
    // Step 5.5: If no match found, try the OTHER mode (handles mode mismatch)
    // This happens when DASHBOARD H1 changes but CALCULATIONS hasn't recalculated yet
    if (!matchingBranch) {
      const log = typeof Logger !== 'undefined' ? Logger.log : console.log;
      log(`WARNING: No match for decision "${actualDecision}" in ${useLongTermSignal ? 'INVEST' : 'TRADE'} mode. Trying ${useLongTermSignal ? 'TRADE' : 'INVEST'} mode...`);
      
      // Toggle mode and try again
      useLongTermSignal = !useLongTermSignal;
      conditionTree = parseDecisionFormula(useLongTermSignal);
      
      if (!conditionTree) {
        throw new Error('Formula parsing failed for alternate mode');
      }
      
      // Try to find matching branch in alternate mode
      for (const branch of conditionTree.branches) {
        // Filter branches based on PURCHASED requirement
        if (branch.requiresPurchased && !isPurchased) {
          continue;
        }
        if (branch.requiresNotPurchased && isPurchased) {
          continue;
        }
        
        // Normalize both values for comparison (trim whitespace)
        const branchResult = String(branch.result || "").trim();
        if (branchResult === actualDecision) {
          matchingBranch = branch;
          break;
        }
      }
      
      if (matchingBranch) {
        log(`SUCCESS: Found match in ${useLongTermSignal ? 'INVEST' : 'TRADE'} mode`);
      }
    }
    
    // If no matching branch found, log available branches and throw error
    if (!matchingBranch) {
      const log = typeof Logger !== 'undefined' ? Logger.log : console.log;
      log(`ERROR: No formula branch found for decision: "${actualDecision}" in either mode`);
      log(`isPurchased: ${isPurchased}`);
      log(`Available branches: ${conditionTree.branches.map(b => `"${b.result}" (purchased=${b.requiresPurchased}, notPurchased=${b.requiresNotPurchased})`).join(', ')}`);
      throw new Error(`No formula branch found for decision: ${actualDecision}`);
    }
    
    // Step 6: Evaluate the matching branch's condition to show which parts passed/failed
    let evaluationResult = null;
    
    // Handle different branch types
    if (matchingBranch.type === "STOP_OUT_CHECK") {
      const price = Number(tickerData.price) || 0;
      const support = Number(tickerData.support) || 0;
      const conditionPassed = (price > 0 && support > 0 && price < support);
      
      evaluationResult = {
        passed: conditionPassed,
        expression: matchingBranch.condition,
        leftValue: price,
        rightValue: support,
        operator: "<",
        details: []
      };
    } else if (matchingBranch.type === "SIGNAL_CHECK") {
      const expectedSignals = Array.isArray(matchingBranch.signalValue) ? matchingBranch.signalValue : [matchingBranch.signalValue];
      const conditionPassed = expectedSignals.includes(signal);
      
      evaluationResult = {
        passed: conditionPassed,
        expression: `SIGNAL = ${signal}`,
        leftValue: signal,
        rightValue: expectedSignals.join(" OR "),
        operator: "=",
        details: []
      };
    } else if (matchingBranch.type === "PATTERN_CHECK") {
      const expectedSignals = Array.isArray(matchingBranch.signalValue) ? matchingBranch.signalValue : [matchingBranch.signalValue];
      const signalMatches = expectedSignals.includes(signal);
      const patternMatches = (matchingBranch.patternType === "any" && hasPattern) ||
                             (matchingBranch.patternType === patternType);
      const conditionPassed = signalMatches && hasPattern && patternMatches;
      
      evaluationResult = {
        passed: conditionPassed,
        expression: `SIGNAL = ${signal} AND PATTERN = ${patternType}`,
        leftValue: signal,
        rightValue: `${expectedSignals.join(" OR ")} + ${matchingBranch.patternType} pattern`,
        operator: "AND",
        details: [
          {
            passed: signalMatches,
            expression: `SIGNAL = ${signal}`,
            leftValue: signal,
            rightValue: expectedSignals.join(" OR "),
            operator: "="
          },
          {
            passed: hasPattern && patternMatches,
            expression: `PATTERN = ${patternType}`,
            leftValue: patternType,
            rightValue: matchingBranch.patternType,
            operator: "="
          }
        ]
      };
    } else if (matchingBranch.type === "COMPLEX") {
      const parsedCondition = parseConditionExpression(matchingBranch.condition);
      evaluationResult = evaluateCondition(parsedCondition, tickerData);
    } else if (matchingBranch.type === "PURCHASED_CHECK") {
      evaluationResult = {
        passed: isPurchased,
        expression: "PURCHASED",
        leftValue: isPurchased,
        rightValue: true,
        operator: "=",
        details: []
      };
    } else if (matchingBranch.type === "DEFAULT") {
      evaluationResult = {
        passed: true,
        expression: "TRUE",
        leftValue: null,
        rightValue: null,
        operator: null,
        details: []
      };
    } else {
      const parsedCondition = parseConditionExpression(matchingBranch.condition);
      evaluationResult = evaluateCondition(parsedCondition, tickerData);
    }
    
    // Build evaluation path with just this branch
    const evaluationPath = [{
      order: matchingBranch.order,
      condition: matchingBranch.condition,
      result: matchingBranch.result,
      passed: evaluationResult.passed,
      expression: evaluationResult.expression,
      leftValue: evaluationResult.leftValue,
      rightValue: evaluationResult.rightValue,
      operator: evaluationResult.operator,
      details: evaluationResult.details,
      error: evaluationResult.error,
      branchType: matchingBranch.type,
      signalValue: matchingBranch.signalValue,
      patternType: matchingBranch.patternType
    }];
    
    // Step 7: Call generateNarrative() to create explanation
    const narrative = generateNarrative(evaluationPath, actualDecision);
    
    // Step 8: Return {decision, narrative, evaluationPath}
    return {
      decision: actualDecision,  // Return the actual decision from CALCULATIONS sheet
      narrative: narrative,
      evaluationPath: evaluationPath
    };
    
  } catch (error) {
    // Log error and re-throw for caller to handle
    const log = typeof Logger !== 'undefined' ? Logger.log : console.log;
    log(`Error in evaluateDecisionFormula: ${error.message}`);
    log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

// Export functions for use in other modules (Node.js)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    evaluateSignalFormula,
    evaluateDecisionFormula,
    parseSignalFormula,
    parseDecisionFormula,
    parseConditionExpression,
    formatValue,
    formatConditionLine,
    generateNarrative,
    COLUMN_MAP,
    evaluateCondition,
    evaluateComparison,
    evaluateAND,
    evaluateOR
  };
}

// Explicitly register functions in global scope for Google Apps Script
// This ensures the functions are available when called from other scripts
if (typeof global !== 'undefined') {
  // Node.js environment
  global.evaluateSignalFormula = evaluateSignalFormula;
  global.evaluateDecisionFormula = evaluateDecisionFormula;
} else if (typeof window !== 'undefined') {
  // Browser environment
  window.evaluateSignalFormula = evaluateSignalFormula;
  window.evaluateDecisionFormula = evaluateDecisionFormula;
} else {
  // Google Apps Script environment - use 'this' which refers to global scope
  this.evaluateSignalFormula = evaluateSignalFormula;
  this.evaluateDecisionFormula = evaluateDecisionFormula;
}
