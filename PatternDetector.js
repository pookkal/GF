/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v3.1_KIRO_OPTIMIZED
* ==============================================================================
*/
/**
 * Detects patterns for a single ticker
 * @param {Array<Object>} priceData - Array of {date, open, high, low, close, volume}
 * @param {Object} options - Detection options {minBars, minConfidence}
 * @returns {Array<Pattern>} - Array of detected patterns
 */
function detectPatterns(priceData, options) {
  // Default options
  const minBars = options && options.minBars !== undefined ? options.minBars : 100;
  const minConfidence = options && options.minConfidence !== undefined ? options.minConfidence : 60;
  
  // Requirement 18.1: Validate input data (minimum 100 bars)
  if (!priceData || !Array.isArray(priceData) || priceData.length < minBars) {
    console.log(`detectPatterns: Insufficient data - ${priceData ? priceData.length : 0} bars (minimum ${minBars} required)`);
    return [];
  }
  
  // Extract high and low prices for pivot detection
  const highs = priceData.map(bar => bar.high);
  const lows = priceData.map(bar => bar.low);
  
  // Find pivot points using findPivotPoints with appropriate window sizes
  // Use window size of 5 for short-term patterns, 10 for longer-term patterns
  const pivots5 = findPivotPoints([...highs, ...lows], 5, 5);
  const pivots10 = findPivotPoints([...highs, ...lows], 10, 10);
  
  // Separate pivots by type for each window size
  const pivotHighs5 = [];
  const pivotLows5 = [];
  
  for (let i = 0; i < highs.length; i++) {
    // Check if this index is a pivot high
    let isPivotHigh = true;
    for (let j = Math.max(0, i - 5); j <= Math.min(highs.length - 1, i + 5); j++) {
      if (j !== i && highs[j] >= highs[i]) {
        isPivotHigh = false;
        break;
      }
    }
    if (isPivotHigh && i >= 5 && i < highs.length - 5) {
      pivotHighs5.push({ index: i, price: highs[i], type: 'high' });
    }
    
    // Check if this index is a pivot low
    let isPivotLow = true;
    for (let j = Math.max(0, i - 5); j <= Math.min(lows.length - 1, i + 5); j++) {
      if (j !== i && lows[j] <= lows[i]) {
        isPivotLow = false;
        break;
      }
    }
    if (isPivotLow && i >= 5 && i < lows.length - 5) {
      pivotLows5.push({ index: i, price: lows[i], type: 'low' });
    }
  }
  
  const pivots5Combined = [...pivotHighs5, ...pivotLows5].sort((a, b) => a.index - b.index);
  
  // Similar for 10-bar window
  const pivotHighs10 = [];
  const pivotLows10 = [];
  
  for (let i = 0; i < highs.length; i++) {
    let isPivotHigh = true;
    for (let j = Math.max(0, i - 10); j <= Math.min(highs.length - 1, i + 10); j++) {
      if (j !== i && highs[j] >= highs[i]) {
        isPivotHigh = false;
        break;
      }
    }
    if (isPivotHigh && i >= 10 && i < highs.length - 10) {
      pivotHighs10.push({ index: i, price: highs[i], type: 'high' });
    }
    
    let isPivotLow = true;
    for (let j = Math.max(0, i - 10); j <= Math.min(lows.length - 1, i + 10); j++) {
      if (j !== i && lows[j] <= lows[i]) {
        isPivotLow = false;
        break;
      }
    }
    if (isPivotLow && i >= 10 && i < lows.length - 10) {
      pivotLows10.push({ index: i, price: lows[i], type: 'low' });
    }
  }
  
  const pivots10Combined = [...pivotHighs10, ...pivotLows10].sort((a, b) => a.index - b.index);
  
  // Run all pattern detectors in sequence
  const detectedPatterns = [];
  
  // Reversal Patterns (use 10-bar pivots for more significant patterns)
  try {
    const doubleTop = detectDoubleTop(priceData, pivots10Combined);
    if (doubleTop) detectedPatterns.push(doubleTop);
  } catch (error) {
    console.error(`Error detecting double top: ${error.message}`);
  }
  
  try {
    const doubleBottom = detectDoubleBottom(priceData, pivots10Combined);
    if (doubleBottom) detectedPatterns.push(doubleBottom);
  } catch (error) {
    console.error(`Error detecting double bottom: ${error.message}`);
  }
  
  try {
    const headShoulders = detectHeadAndShoulders(priceData, pivots10Combined);
    if (headShoulders) detectedPatterns.push(headShoulders);
  } catch (error) {
    console.error(`Error detecting head and shoulders: ${error.message}`);
  }
  
  try {
    const inverseHeadShoulders = detectInverseHeadAndShoulders(priceData, pivots10Combined);
    if (inverseHeadShoulders) detectedPatterns.push(inverseHeadShoulders);
  } catch (error) {
    console.error(`Error detecting inverse head and shoulders: ${error.message}`);
  }
  
  try {
    const cupHandle = detectCupAndHandle(priceData, pivots10Combined);
    if (cupHandle) detectedPatterns.push(cupHandle);
  } catch (error) {
    console.error(`Error detecting cup and handle: ${error.message}`);
  }
  
  try {
    const roundingBottom = detectRoundingBottom(priceData, pivots10Combined);
    if (roundingBottom) detectedPatterns.push(roundingBottom);
  } catch (error) {
    console.error(`Error detecting rounding bottom: ${error.message}`);
  }
  
  // Triangle Patterns (use 5-bar pivots for more responsive detection)
  try {
    const ascendingTriangle = detectAscendingTriangle(priceData, pivots5Combined);
    if (ascendingTriangle) detectedPatterns.push(ascendingTriangle);
  } catch (error) {
    console.error(`Error detecting ascending triangle: ${error.message}`);
  }
  
  try {
    const descendingTriangle = detectDescendingTriangle(priceData, pivots5Combined);
    if (descendingTriangle) detectedPatterns.push(descendingTriangle);
  } catch (error) {
    console.error(`Error detecting descending triangle: ${error.message}`);
  }
  
  try {
    const symmetricalTriangle = detectSymmetricalTriangle(priceData, pivots5Combined);
    if (symmetricalTriangle) detectedPatterns.push(symmetricalTriangle);
  } catch (error) {
    console.error(`Error detecting symmetrical triangle: ${error.message}`);
  }
  
  // Continuation Patterns (use 5-bar pivots)
  try {
    const flag = detectFlag(priceData, pivots5Combined);
    if (flag) detectedPatterns.push(flag);
  } catch (error) {
    console.error(`Error detecting flag: ${error.message}`);
  }
  
  try {
    const pennant = detectPennant(priceData, pivots5Combined);
    if (pennant) detectedPatterns.push(pennant);
  } catch (error) {
    console.error(`Error detecting pennant: ${error.message}`);
  }
  
  // Wedge Patterns (use 5-bar pivots)
  try {
    const risingWedge = detectRisingWedge(priceData, pivots5Combined);
    if (risingWedge) detectedPatterns.push(risingWedge);
  } catch (error) {
    console.error(`Error detecting rising wedge: ${error.message}`);
  }
  
  try {
    const fallingWedge = detectFallingWedge(priceData, pivots5Combined);
    if (fallingWedge) detectedPatterns.push(fallingWedge);
  } catch (error) {
    console.error(`Error detecting falling wedge: ${error.message}`);
  }
  
  // Consolidation Patterns (use 5-bar pivots)
  try {
    const rectangle = detectRectangle(priceData, pivots5Combined);
    if (rectangle) detectedPatterns.push(rectangle);
  } catch (error) {
    console.error(`Error detecting rectangle: ${error.message}`);
  }
  
  // Breakout and Gap Patterns
  try {
    const breakout = detectBreakout(priceData, pivots5Combined);
    if (breakout) detectedPatterns.push(breakout);
  } catch (error) {
    console.error(`Error detecting breakout: ${error.message}`);
  }
  
  try {
    const gap = detectGap(priceData);
    if (gap) detectedPatterns.push(gap);
  } catch (error) {
    console.error(`Error detecting gap: ${error.message}`);
  }
  
  // Requirement 18.2: For each detected pattern, calculate confidence score
  for (const pattern of detectedPatterns) {
    try {
      pattern.confidence = calculateConfidence(pattern, priceData);
    } catch (error) {
      console.error(`Error calculating confidence for ${pattern.type}: ${error.message}`);
      pattern.confidence = 0;
    }
  }
  
  // Requirement 18.3: Filter out patterns below minimum confidence threshold
  const highConfidencePatterns = detectedPatterns.filter(pattern => 
    pattern && pattern.confidence >= minConfidence
  );
  
  console.log(`detectPatterns: Found ${detectedPatterns.length} patterns, ${highConfidencePatterns.length} above ${minConfidence}% confidence`);
  
  // Requirement 18.4: Call prioritizePatterns to handle overlapping patterns
  const prioritizedPatterns = prioritizePatterns(highConfidencePatterns);
  
  // Return final pattern list sorted by confidence (highest first)
  const sortedPatterns = prioritizedPatterns.sort((a, b) => b.confidence - a.confidence);
  
  return sortedPatterns;
}

/**
 * Detects patterns for all tickers and updates CALCULATIONS sheet
 * @param {Sheet} dataSheet - DATA sheet reference
 * @param {Sheet} calcSheet - CALCULATIONS sheet reference
 * @param {Array<string>} tickers - Array of ticker symbols
 * @param {Object} options - Optional detection options {minBars, minConfidence, batchSize, blockSize}
 */
function detectPatternsForAllTickers(dataSheet, calcSheet, tickers, options = {}) {
  // Default options
  const minBars = options.minBars || 100;
  const minConfidence = options.minConfidence || 60;
  const batchSize = options.batchSize || 10;
  const blockSize = options.blockSize || 7;
  
  // Validate inputs
  if (!dataSheet || !calcSheet || !tickers || tickers.length === 0) {
    console.log('Invalid inputs to detectPatternsForAllTickers');
    return;
  }
  
  console.log(`Starting pattern detection for ${tickers.length} tickers in batches of ${batchSize}`);
  
  // Process tickers in batches to avoid timeout and memory issues
  for (let batchStart = 0; batchStart < tickers.length; batchStart += batchSize) {
    const batchEnd = Math.min(batchStart + batchSize, tickers.length);
    const batchTickers = tickers.slice(batchStart, batchEnd);
    
    console.log(`Processing batch ${Math.floor(batchStart / batchSize) + 1}: tickers ${batchStart + 1}-${batchEnd}`);
    
    // Process each ticker in the batch
    const batchResults = [];
    
    for (let i = 0; i < batchTickers.length; i++) {
      const ticker = batchTickers[i];
      const tickerIndex = batchStart + i;
      
      try {
        // Get price data for this ticker
        const priceData = getPriceDataForTicker(dataSheet, ticker, tickerIndex, blockSize);
        
        // Check if we have sufficient data
        if (!priceData || priceData.length < minBars) {
          console.log(`Insufficient data for ${ticker}: ${priceData ? priceData.length : 0} bars (minimum ${minBars} required)`);
          batchResults.push('');
          continue;
        }
        
        // Detect patterns for this ticker
        const patterns = detectPatterns(priceData, {
          minBars: minBars,
          minConfidence: minConfidence
        });
        
        // Format patterns for sheet
        const patternString = formatPatternsForSheet(patterns);
        
        batchResults.push(patternString);
        
        console.log(`${ticker}: Found ${patterns.length} patterns - ${patternString || 'none'}`);
        
      } catch (error) {
        // Handle errors gracefully - log and continue with next ticker
        console.error(`Error processing ticker ${ticker}: ${error.message}`);
        console.error(error.stack);
        batchResults.push('');
      }
    }
    
    // Write batch results to CALCULATIONS sheet
    // Pattern results go in column AF (column 32)
    // Rows start at row 3 (header rows 1-2)
    try {
      const startRow = batchStart + 3;
      const numRows = batchResults.length;
      const columnAF = 32; // Column AF is the 32nd column
      
      // Write results as a 2D array (one column, multiple rows)
      const resultsArray = batchResults.map(result => [result]);
      
      if (numRows > 0) {
        calcSheet.getRange(startRow, columnAF, numRows, 1).setValues(resultsArray);
        SpreadsheetApp.flush(); // Ensure changes are written
      }
      
      console.log(`Wrote ${numRows} pattern results to CALCULATIONS sheet (rows ${startRow}-${startRow + numRows - 1})`);
      
    } catch (error) {
      console.error(`Error writing batch results to sheet: ${error.message}`);
      console.error(error.stack);
      // Continue with next batch even if write fails
    }
  }
  
  console.log('Pattern detection complete for all tickers');
}

// ============================================================================
// Pivot Point Detection
// ============================================================================

/**
 * Finds pivot points (local extrema) in price data
 * @param {Array<number>} prices - Array of prices (high or low)
 * @param {number} leftBars - Number of bars to the left for comparison
 * @param {number} rightBars - Number of bars to the right for comparison
 * @returns {Array<Object>} - Array of {index, price, type: 'high'|'low'}
 */
function findPivotPoints(prices, leftBars, rightBars) {
  if (!prices || prices.length === 0) {
    return [];
  }
  
  // Need at least leftBars + 1 + rightBars to find any pivots
  const minLength = leftBars + rightBars + 1;
  if (prices.length < minLength) {
    return [];
  }
  
  const pivots = [];
  
  // Start from leftBars and end at length - rightBars to ensure we have enough bars on both sides
  for (let i = leftBars; i < prices.length - rightBars; i++) {
    const currentPrice = prices[i];
    
    // Skip if current price is NaN or invalid
    if (!isFinite(currentPrice)) {
      continue;
    }
    
    // Check if this is a pivot high (local maximum)
    let isPivotHigh = true;
    for (let j = i - leftBars; j <= i + rightBars; j++) {
      if (j !== i) {
        // Skip comparison if the comparison price is NaN or invalid
        if (!isFinite(prices[j])) {
          isPivotHigh = false;
          break;
        }
        if (prices[j] >= currentPrice) {
          isPivotHigh = false;
          break;
        }
      }
    }
    
    if (isPivotHigh) {
      pivots.push({
        index: i,
        price: currentPrice,
        type: 'high'
      });
      continue; // A point can't be both high and low
    }
    
    // Check if this is a pivot low (local minimum)
    let isPivotLow = true;
    for (let j = i - leftBars; j <= i + rightBars; j++) {
      if (j !== i) {
        // Skip comparison if the comparison price is NaN or invalid
        if (!isFinite(prices[j])) {
          isPivotLow = false;
          break;
        }
        if (prices[j] <= currentPrice) {
          isPivotLow = false;
          break;
        }
      }
    }
    
    if (isPivotLow) {
      pivots.push({
        index: i,
        price: currentPrice,
        type: 'low'
      });
    }
  }
  
  return pivots;
}

// ============================================================================
// Pattern Validation Functions
// ============================================================================

/**
 * Validates a pattern against quality criteria
 * @param {Pattern} pattern - Pattern to validate
 * @param {Object} criteria - Validation criteria {minSpacing, minDepth}
 * @returns {boolean} - True if pattern is valid
 */
function validatePattern(pattern, criteria) {
  if (!pattern || !pattern.keyPoints || pattern.keyPoints.length === 0) {
    return false;
  }
  
  // Check minimum spacing if specified
  if (criteria.minSpacing !== undefined) {
    if (!hasMinimumSpacing(pattern, criteria.minSpacing)) {
      return false;
    }
  }
  
  // Check minimum depth if specified
  if (criteria.minDepth !== undefined) {
    if (!hasMinimumDepth(pattern, criteria.minDepth)) {
      return false;
    }
  }
  
  return true;
}

/**
 * Checks if pattern meets minimum bar spacing requirements
 * @param {Pattern} pattern - Pattern to check
 * @param {number} minBars - Minimum bars required
 * @returns {boolean}
 */
function hasMinimumSpacing(pattern, minBars) {
  if (!pattern || !pattern.keyPoints || pattern.keyPoints.length < 2) {
    return false;
  }
  
  // Check spacing between consecutive key points
  for (let i = 1; i < pattern.keyPoints.length; i++) {
    const prevPoint = pattern.keyPoints[i - 1];
    const currPoint = pattern.keyPoints[i];
    
    // Validate that both points have valid indices
    if (prevPoint.index === undefined || currPoint.index === undefined) {
      return false;
    }
    
    const spacing = currPoint.index - prevPoint.index;
    
    // Check if spacing meets minimum requirement
    if (spacing < minBars) {
      return false;
    }
  }
  
  return true;
}

/**
 * Checks if pattern meets depth/height requirements
 * @param {Pattern} pattern - Pattern to check
 * @param {number} minPercent - Minimum percentage move (e.g., 10 for 10%)
 * @returns {boolean}
 */
function hasMinimumDepth(pattern, minPercent) {
  if (!pattern || !pattern.keyPoints || pattern.keyPoints.length < 2) {
    return false;
  }
  
  // Find the highest and lowest prices in the pattern's key points
  let highestPrice = -Infinity;
  let lowestPrice = Infinity;
  
  for (const point of pattern.keyPoints) {
    if (point.price === undefined || !isFinite(point.price)) {
      return false;
    }
    
    if (point.price > highestPrice) {
      highestPrice = point.price;
    }
    if (point.price < lowestPrice) {
      lowestPrice = point.price;
    }
  }
  
  // Calculate the percentage depth/height of the pattern
  // Depth is calculated as: (high - low) / high * 100
  if (highestPrice <= 0) {
    return false;
  }
  
  const depthPercent = ((highestPrice - lowestPrice) / highestPrice) * 100;
  
  // Check if depth meets minimum requirement
  return depthPercent >= minPercent;
}

// ============================================================================
// Pattern Recognition Functions - Reversal Patterns
// ============================================================================

/**
 * Detects double top pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectDoubleTop(priceData, pivots) {
  if (!priceData || priceData.length === 0 || !pivots || pivots.length === 0) {
    return null;
  }
  
  // Filter for pivot highs only
  const pivotHighs = pivots.filter(p => p.type === 'high');
  
  // Need at least 2 pivot highs to form a double top
  if (pivotHighs.length < 2) {
    return null;
  }
  
  // Look for two peaks at similar price levels
  // Iterate through pairs of pivot highs
  for (let i = 0; i < pivotHighs.length - 1; i++) {
    const peak1 = pivotHighs[i];
    
    for (let j = i + 1; j < pivotHighs.length; j++) {
      const peak2 = pivotHighs[j];
      
      // Check if peaks are separated by at least 10 bars (Requirement 1.2)
      const barSpacing = peak2.index - peak1.index;
      if (barSpacing < 10) {
        continue;
      }
      
      // Check if peaks are at approximately the same price level
      // Allow up to 3% difference between peaks
      const priceDifference = Math.abs(peak1.price - peak2.price);
      const avgPeakPrice = (peak1.price + peak2.price) / 2;
      const percentDifference = (priceDifference / avgPeakPrice) * 100;
      
      if (percentDifference > 3) {
        continue;
      }
      
      // Find the valley (lowest point) between the two peaks
      let valleyIndex = peak1.index;
      let valleyPrice = peak1.price;
      
      for (let k = peak1.index + 1; k < peak2.index; k++) {
        if (priceData[k] && priceData[k].low < valleyPrice) {
          valleyPrice = priceData[k].low;
          valleyIndex = k;
        }
      }
      
      // Verify valley depth is at least 10% below the peak price (Requirement 1.3)
      const valleyDepthPercent = ((avgPeakPrice - valleyPrice) / avgPeakPrice) * 100;
      
      if (valleyDepthPercent < 10) {
        continue;
      }
      
      // We found a valid double top pattern
      // Calculate neckline (the valley price serves as the neckline level)
      const neckline = valleyPrice;
      
      // Check if pattern is confirmed (price has broken below neckline)
      let confirmed = false;
      for (let k = peak2.index + 1; k < priceData.length; k++) {
        if (priceData[k] && priceData[k].close < neckline) {
          confirmed = true;
          break;
        }
      }
      
      // Calculate target price (traditional measure: neckline - pattern height)
      const patternHeight = avgPeakPrice - neckline;
      const targetPrice = neckline - patternHeight;
      
      // Return the pattern object (Requirement 1.4)
      return {
        type: 'DOUBLE_TOP',
        startIndex: peak1.index,
        endIndex: peak2.index,
        keyPoints: [
          { index: peak1.index, price: peak1.price, label: 'peak1' },
          { index: valleyIndex, price: valleyPrice, label: 'valley' },
          { index: peak2.index, price: peak2.price, label: 'peak2' }
        ],
        neckline: neckline,
        confirmed: confirmed,
        confidence: 0, // Will be calculated by calculateConfidence function
        direction: 'BEARISH',
        targetPrice: targetPrice,
        metadata: {
          peakDifference: percentDifference,
          valleyDepth: valleyDepthPercent
        }
      };
    }
  }
  
  // No valid double top pattern found
  return null;
}

/**
 * Detects double bottom pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectDoubleBottom(priceData, pivots) {
  if (!priceData || priceData.length === 0 || !pivots || pivots.length === 0) {
    return null;
  }
  
  // Filter for pivot lows only
  const pivotLows = pivots.filter(p => p.type === 'low');
  
  // Need at least 2 pivot lows to form a double bottom
  if (pivotLows.length < 2) {
    return null;
  }
  
  // Look for two troughs at similar price levels
  // Iterate through pairs of pivot lows
  for (let i = 0; i < pivotLows.length - 1; i++) {
    const trough1 = pivotLows[i];
    
    for (let j = i + 1; j < pivotLows.length; j++) {
      const trough2 = pivotLows[j];
      
      // Check if troughs are separated by at least 10 bars (Requirement 2.2)
      const barSpacing = trough2.index - trough1.index;
      if (barSpacing < 10) {
        continue;
      }
      
      // Check if troughs are at approximately the same price level
      // Allow up to 3% difference between troughs
      const priceDifference = Math.abs(trough1.price - trough2.price);
      const avgTroughPrice = (trough1.price + trough2.price) / 2;
      const percentDifference = (priceDifference / avgTroughPrice) * 100;
      
      if (percentDifference > 3) {
        continue;
      }
      
      // Find the peak (highest point) between the two troughs
      let peakIndex = trough1.index;
      let peakPrice = trough1.price;
      
      for (let k = trough1.index + 1; k < trough2.index; k++) {
        if (priceData[k] && priceData[k].high > peakPrice) {
          peakPrice = priceData[k].high;
          peakIndex = k;
        }
      }
      
      // Verify peak height is at least 10% above the trough price (Requirement 2.3)
      const peakHeightPercent = ((peakPrice - avgTroughPrice) / avgTroughPrice) * 100;
      
      if (peakHeightPercent < 10) {
        continue;
      }
      
      // We found a valid double bottom pattern
      // Calculate neckline (the peak price serves as the neckline level)
      const neckline = peakPrice;
      
      // Check if pattern is confirmed (price has broken above neckline)
      let confirmed = false;
      for (let k = trough2.index + 1; k < priceData.length; k++) {
        if (priceData[k] && priceData[k].close > neckline) {
          confirmed = true;
          break;
        }
      }
      
      // Calculate target price (traditional measure: neckline + pattern height)
      const patternHeight = neckline - avgTroughPrice;
      const targetPrice = neckline + patternHeight;
      
      // Return the pattern object (Requirement 2.4)
      return {
        type: 'DOUBLE_BOTTOM',
        startIndex: trough1.index,
        endIndex: trough2.index,
        keyPoints: [
          { index: trough1.index, price: trough1.price, label: 'trough1' },
          { index: peakIndex, price: peakPrice, label: 'peak' },
          { index: trough2.index, price: trough2.price, label: 'trough2' }
        ],
        neckline: neckline,
        confirmed: confirmed,
        confidence: 0, // Will be calculated by calculateConfidence function
        direction: 'BULLISH',
        targetPrice: targetPrice,
        metadata: {
          troughDifference: percentDifference,
          peakHeight: peakHeightPercent
        }
      };
    }
  }
  
  // No valid double bottom pattern found
  return null;
}

/**
 * Detects head and shoulders pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectHeadAndShoulders(priceData, pivots) {
  if (!priceData || priceData.length === 0 || !pivots || pivots.length === 0) {
    return null;
  }
  
  // Filter for pivot highs and lows
  const pivotHighs = pivots.filter(p => p.type === 'high');
  const pivotLows = pivots.filter(p => p.type === 'low');
  
  // Need at least 3 pivot highs (left shoulder, head, right shoulder) and 2 pivot lows (troughs)
  if (pivotHighs.length < 3 || pivotLows.length < 2) {
    return null;
  }
  
  // Look for three peaks where the middle peak is highest
  // Iterate through combinations of three consecutive pivot highs
  for (let i = 0; i < pivotHighs.length - 2; i++) {
    const leftShoulder = pivotHighs[i];
    const head = pivotHighs[i + 1];
    const rightShoulder = pivotHighs[i + 2];
    
    // Requirement 3.1: Middle peak (head) must be higher than both shoulders
    if (head.price <= leftShoulder.price || head.price <= rightShoulder.price) {
      continue;
    }
    
    // Requirement 3.3: Each peak must be separated by at least 10 bars
    const spacing1 = head.index - leftShoulder.index;
    const spacing2 = rightShoulder.index - head.index;
    
    if (spacing1 < 10 || spacing2 < 10) {
      continue;
    }
    
    // Requirement 3.2: Left and right shoulders should be at approximately the same height
    // Allow up to 5% difference between shoulders
    const shoulderAvg = (leftShoulder.price + rightShoulder.price) / 2;
    const shoulderDiff = Math.abs(leftShoulder.price - rightShoulder.price);
    const shoulderDiffPercent = (shoulderDiff / shoulderAvg) * 100;
    
    if (shoulderDiffPercent > 5) {
      continue;
    }
    
    // Find the troughs (valleys) between the peaks
    // Trough 1: between left shoulder and head
    let trough1Index = leftShoulder.index;
    let trough1Price = leftShoulder.price;
    
    for (let k = leftShoulder.index + 1; k < head.index; k++) {
      if (priceData[k] && priceData[k].low < trough1Price) {
        trough1Price = priceData[k].low;
        trough1Index = k;
      }
    }
    
    // Trough 2: between head and right shoulder
    let trough2Index = head.index;
    let trough2Price = head.price;
    
    for (let k = head.index + 1; k < rightShoulder.index; k++) {
      if (priceData[k] && priceData[k].low < trough2Price) {
        trough2Price = priceData[k].low;
        trough2Index = k;
      }
    }
    
    // Requirement 3.4: Calculate neckline connecting the troughs
    // The neckline is typically drawn as a line connecting the two troughs
    // For simplicity, we'll use the average of the two trough prices
    // In a more sophisticated implementation, we could calculate the actual trendline
    const neckline = (trough1Price + trough2Price) / 2;
    
    // Verify the pattern has reasonable depth
    // The head should be significantly higher than the neckline (at least 10%)
    const patternHeight = head.price - neckline;
    const heightPercent = (patternHeight / head.price) * 100;
    
    if (heightPercent < 10) {
      continue;
    }
    
    // Check if pattern is confirmed (price has broken below neckline)
    // Requirement 3.5: Pattern is confirmed when price breaks below neckline
    let confirmed = false;
    for (let k = rightShoulder.index + 1; k < priceData.length; k++) {
      if (priceData[k] && priceData[k].close < neckline) {
        confirmed = true;
        break;
      }
    }
    
    // Calculate target price (traditional measure: neckline - pattern height)
    const targetPrice = neckline - patternHeight;
    
    // We found a valid head and shoulders pattern
    return {
      type: 'HEAD_SHOULDERS',
      startIndex: leftShoulder.index,
      endIndex: rightShoulder.index,
      keyPoints: [
        { index: leftShoulder.index, price: leftShoulder.price, label: 'leftShoulder' },
        { index: trough1Index, price: trough1Price, label: 'trough1' },
        { index: head.index, price: head.price, label: 'head' },
        { index: trough2Index, price: trough2Price, label: 'trough2' },
        { index: rightShoulder.index, price: rightShoulder.price, label: 'rightShoulder' }
      ],
      neckline: neckline,
      confirmed: confirmed,
      confidence: 0, // Will be calculated by calculateConfidence function
      direction: 'BEARISH',
      targetPrice: targetPrice,
      metadata: {
        shoulderDifference: shoulderDiffPercent,
        patternHeight: patternHeight,
        heightPercent: heightPercent,
        trough1Price: trough1Price,
        trough2Price: trough2Price
      }
    };
  }
  
  // No valid head and shoulders pattern found
  return null;
}

/**
 * Detects inverse head and shoulders pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectInverseHeadAndShoulders(priceData, pivots) {
  if (!priceData || priceData.length === 0 || !pivots || pivots.length === 0) {
    return null;
  }
  
  // Filter for pivot lows and highs
  const pivotLows = pivots.filter(p => p.type === 'low');
  const pivotHighs = pivots.filter(p => p.type === 'high');
  
  // Need at least 3 pivot lows (left shoulder, head, right shoulder) and 2 pivot highs (peaks)
  if (pivotLows.length < 3 || pivotHighs.length < 2) {
    return null;
  }
  
  // Look for three troughs where the middle trough is lowest
  // Iterate through combinations of three consecutive pivot lows
  for (let i = 0; i < pivotLows.length - 2; i++) {
    const leftShoulder = pivotLows[i];
    const head = pivotLows[i + 1];
    const rightShoulder = pivotLows[i + 2];
    
    // Requirement 4.1: Middle trough (head) must be lower than both shoulders
    if (head.price >= leftShoulder.price || head.price >= rightShoulder.price) {
      continue;
    }
    
    // Requirement 4.3: Each trough must be separated by at least 10 bars
    const spacing1 = head.index - leftShoulder.index;
    const spacing2 = rightShoulder.index - head.index;
    
    if (spacing1 < 10 || spacing2 < 10) {
      continue;
    }
    
    // Requirement 4.2: Left and right shoulders should be at approximately the same depth
    // Allow up to 5% difference between shoulders
    const shoulderAvg = (leftShoulder.price + rightShoulder.price) / 2;
    const shoulderDiff = Math.abs(leftShoulder.price - rightShoulder.price);
    const shoulderDiffPercent = (shoulderDiff / shoulderAvg) * 100;
    
    if (shoulderDiffPercent > 5) {
      continue;
    }
    
    // Find the peaks between the troughs
    // Peak 1: between left shoulder and head
    let peak1Index = leftShoulder.index;
    let peak1Price = leftShoulder.price;
    
    for (let k = leftShoulder.index + 1; k < head.index; k++) {
      if (priceData[k] && priceData[k].high > peak1Price) {
        peak1Price = priceData[k].high;
        peak1Index = k;
      }
    }
    
    // Peak 2: between head and right shoulder
    let peak2Index = head.index;
    let peak2Price = head.price;
    
    for (let k = head.index + 1; k < rightShoulder.index; k++) {
      if (priceData[k] && priceData[k].high > peak2Price) {
        peak2Price = priceData[k].high;
        peak2Index = k;
      }
    }
    
    // Requirement 4.4: Calculate neckline connecting the peaks
    // The neckline is typically drawn as a line connecting the two peaks
    // For simplicity, we'll use the average of the two peak prices
    // In a more sophisticated implementation, we could calculate the actual trendline
    const neckline = (peak1Price + peak2Price) / 2;
    
    // Verify the pattern has reasonable depth
    // The head should be significantly lower than the neckline (at least 10%)
    const patternHeight = neckline - head.price;
    const heightPercent = (patternHeight / neckline) * 100;
    
    if (heightPercent < 10) {
      continue;
    }
    
    // Check if pattern is confirmed (price has broken above neckline)
    // Requirement 4.5: Pattern is confirmed when price breaks above neckline
    let confirmed = false;
    for (let k = rightShoulder.index + 1; k < priceData.length; k++) {
      if (priceData[k] && priceData[k].close > neckline) {
        confirmed = true;
        break;
      }
    }
    
    // Calculate target price (traditional measure: neckline + pattern height)
    const targetPrice = neckline + patternHeight;
    
    // We found a valid inverse head and shoulders pattern
    return {
      type: 'INVERSE_HEAD_SHOULDERS',
      startIndex: leftShoulder.index,
      endIndex: rightShoulder.index,
      keyPoints: [
        { index: leftShoulder.index, price: leftShoulder.price, label: 'leftShoulder' },
        { index: peak1Index, price: peak1Price, label: 'peak1' },
        { index: head.index, price: head.price, label: 'head' },
        { index: peak2Index, price: peak2Price, label: 'peak2' },
        { index: rightShoulder.index, price: rightShoulder.price, label: 'rightShoulder' }
      ],
      neckline: neckline,
      confirmed: confirmed,
      confidence: 0, // Will be calculated by calculateConfidence function
      direction: 'BULLISH',
      targetPrice: targetPrice,
      metadata: {
        shoulderDifference: shoulderDiffPercent,
        patternHeight: patternHeight,
        heightPercent: heightPercent,
        peak1Price: peak1Price,
        peak2Price: peak2Price
      }
    };
  }
  
  // No valid inverse head and shoulders pattern found
  return null;
}

/**
 * Detects cup and handle pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectCupAndHandle(priceData, pivots) {
  if (!priceData || priceData.length < 40) {
    return null;
  }
  
  // Requirement 5.2: Cup must span at least 30 bars, plus we need room for the handle
  const minCupLength = 30;
  const minHandleLength = 5;
  const minTotalLength = minCupLength + minHandleLength;
  
  // Look for U-shaped cup formations
  // We'll scan through the price data looking for a starting high, a bottom, and a recovery to near the starting level
  for (let cupStartIdx = 0; cupStartIdx < priceData.length - minTotalLength; cupStartIdx++) {
    const cupStartPrice = priceData[cupStartIdx].close;
    
    // Look for potential cup endpoints (at least minCupLength bars away)
    for (let cupEndIdx = cupStartIdx + minCupLength; cupEndIdx < priceData.length - minHandleLength; cupEndIdx++) {
      const cupEndPrice = priceData[cupEndIdx].close;
      
      // Requirement 5.1: Cup should show recovery back to near starting level
      // Check if cup end price is close to cup start price (within 5%)
      const cupPriceDifference = Math.abs(cupStartPrice - cupEndPrice);
      const cupAvgPrice = (cupStartPrice + cupEndPrice) / 2;
      const cupPercentDifference = (cupPriceDifference / cupAvgPrice) * 100;
      
      if (cupPercentDifference > 5) {
        continue;
      }
      
      // Find the bottom of the cup (lowest point in the range)
      let cupBottomIdx = cupStartIdx;
      let cupBottomPrice = cupStartPrice;
      
      for (let i = cupStartIdx + 1; i < cupEndIdx; i++) {
        if (priceData[i].low < cupBottomPrice) {
          cupBottomPrice = priceData[i].low;
          cupBottomIdx = i;
        }
      }
      
      // Verify the cup has reasonable depth (at least 10%)
      const cupDepthPercent = ((cupAvgPrice - cupBottomPrice) / cupAvgPrice) * 100;
      if (cupDepthPercent < 10) {
        continue;
      }
      
      // Verify U-shape: bottom should be roughly in the middle of the cup
      // Allow the bottom to be anywhere from 30% to 70% through the cup
      const cupLength = cupEndIdx - cupStartIdx;
      const bottomPosition = (cupBottomIdx - cupStartIdx) / cupLength;
      
      if (bottomPosition < 0.3 || bottomPosition > 0.7) {
        continue;
      }
      
      // Calculate the cup's advance (from bottom to cup end)
      const cupAdvance = cupEndPrice - cupBottomPrice;
      
      // Requirement 5.3: Now look for the handle (smaller consolidation following the cup)
      // The handle should be a pullback from the cup's high
      // Look for a handle that spans between 5 and 20 bars
      const maxHandleLength = 20;
      
      for (let handleEndIdx = cupEndIdx + minHandleLength; 
           handleEndIdx <= Math.min(cupEndIdx + maxHandleLength, priceData.length - 1); 
           handleEndIdx++) {
        
        // Find the lowest point in the handle (handle bottom)
        let handleBottomIdx = cupEndIdx;
        let handleBottomPrice = cupEndPrice;
        
        for (let i = cupEndIdx + 1; i <= handleEndIdx; i++) {
          if (priceData[i].low < handleBottomPrice) {
            handleBottomPrice = priceData[i].low;
            handleBottomIdx = i;
          }
        }
        
        // Requirement 5.4: Verify handle retraces no more than 50% of the cup's advance
        const handleRetracement = cupEndPrice - handleBottomPrice;
        const retracementPercent = (handleRetracement / cupAdvance) * 100;
        
        if (retracementPercent > 50) {
          continue;
        }
        
        // Verify the handle doesn't retrace too little (at least 10% to be meaningful)
        if (retracementPercent < 10) {
          continue;
        }
        
        // Find the highest point in the handle (handle resistance)
        let handleHighIdx = cupEndIdx;
        let handleHighPrice = cupEndPrice;
        
        for (let i = cupEndIdx; i <= handleEndIdx; i++) {
          if (priceData[i].high > handleHighPrice) {
            handleHighPrice = priceData[i].high;
            handleHighIdx = i;
          }
        }
        
        // The handle resistance should be close to the cup end price
        const handleResistance = Math.max(cupEndPrice, handleHighPrice);
        
        // Requirement 5.5: Check if pattern is confirmed (price breaks above handle's resistance)
        let confirmed = false;
        
        for (let i = handleEndIdx + 1; i < priceData.length; i++) {
          if (priceData[i].close > handleResistance) {
            confirmed = true;
            break;
          }
        }
        
        // Calculate target price (traditional measure: cup depth added to breakout level)
        const cupDepth = cupAvgPrice - cupBottomPrice;
        const targetPrice = handleResistance + cupDepth;
        
        // We found a valid cup and handle pattern
        return {
          type: 'CUP_AND_HANDLE',
          startIndex: cupStartIdx,
          endIndex: handleEndIdx,
          keyPoints: [
            { index: cupStartIdx, price: cupStartPrice, label: 'cupStart' },
            { index: cupBottomIdx, price: cupBottomPrice, label: 'cupBottom' },
            { index: cupEndIdx, price: cupEndPrice, label: 'cupEnd' },
            { index: handleBottomIdx, price: handleBottomPrice, label: 'handleBottom' },
            { index: handleEndIdx, price: priceData[handleEndIdx].close, label: 'handleEnd' }
          ],
          neckline: handleResistance,
          confirmed: confirmed,
          confidence: 0, // Will be calculated by calculateConfidence function
          direction: 'BULLISH',
          targetPrice: targetPrice,
          metadata: {
            cupDepthPercent: cupDepthPercent,
            cupLength: cupLength,
            cupAdvance: cupAdvance,
            handleLength: handleEndIdx - cupEndIdx,
            handleRetracement: handleRetracement,
            retracementPercent: retracementPercent,
            handleResistance: handleResistance
          }
        };
      }
    }
  }
  
  // No valid cup and handle pattern found
  return null;
}

/**
 * Detects rounding bottom pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectRoundingBottom(priceData, pivots) {
  if (!priceData || priceData.length < 30) {
    return null;
  }
  
  // Requirement 14.2: Pattern must span at least 30 bars
  const minPatternLength = 30;
  
  // Look for U-shaped formations by scanning through the price data
  // We'll look for a starting high point, a bottom, and an ending high point
  for (let startIdx = 0; startIdx < priceData.length - minPatternLength; startIdx++) {
    const startPrice = priceData[startIdx].close;
    
    // Look for potential pattern endpoints (at least minPatternLength bars away)
    for (let endIdx = startIdx + minPatternLength; endIdx < priceData.length; endIdx++) {
      const endPrice = priceData[endIdx].close;
      
      // Check if start and end prices are at similar levels (within 5%)
      const priceDifference = Math.abs(startPrice - endPrice);
      const avgPrice = (startPrice + endPrice) / 2;
      const percentDifference = (priceDifference / avgPrice) * 100;
      
      if (percentDifference > 5) {
        continue;
      }
      
      // Find the lowest point (bottom of the U) in the range
      let bottomIdx = startIdx;
      let bottomPrice = startPrice;
      
      for (let i = startIdx + 1; i < endIdx; i++) {
        if (priceData[i].low < bottomPrice) {
          bottomPrice = priceData[i].low;
          bottomIdx = i;
        }
      }
      
      // Verify the bottom is significantly lower than start/end (at least 10%)
      const depthPercent = ((avgPrice - bottomPrice) / avgPrice) * 100;
      if (depthPercent < 10) {
        continue;
      }
      
      // Requirement 14.3: Verify symmetry of decline and recovery
      // Calculate the decline phase (start to bottom) and recovery phase (bottom to end)
      const declineLength = bottomIdx - startIdx;
      const recoveryLength = endIdx - bottomIdx;
      
      // Check if decline and recovery are approximately symmetrical (within 40% difference)
      const lengthDifference = Math.abs(declineLength - recoveryLength);
      const avgLength = (declineLength + recoveryLength) / 2;
      const symmetryPercent = (lengthDifference / avgLength) * 100;
      
      if (symmetryPercent > 40) {
        continue;
      }
      
      // Requirement 14.4: Check volume pattern (decreasing during decline, increasing during recovery)
      // Calculate average volume for decline phase
      let declineVolumeSum = 0;
      let declineVolumeCount = 0;
      for (let i = startIdx; i < bottomIdx; i++) {
        if (priceData[i].volume !== undefined && isFinite(priceData[i].volume)) {
          declineVolumeSum += priceData[i].volume;
          declineVolumeCount++;
        }
      }
      
      // Calculate average volume for recovery phase
      let recoveryVolumeSum = 0;
      let recoveryVolumeCount = 0;
      for (let i = bottomIdx; i < endIdx; i++) {
        if (priceData[i].volume !== undefined && isFinite(priceData[i].volume)) {
          recoveryVolumeSum += priceData[i].volume;
          recoveryVolumeCount++;
        }
      }
      
      // Skip if we don't have volume data
      if (declineVolumeCount === 0 || recoveryVolumeCount === 0) {
        continue;
      }
      
      const avgDeclineVolume = declineVolumeSum / declineVolumeCount;
      const avgRecoveryVolume = recoveryVolumeSum / recoveryVolumeCount;
      
      // Verify volume increases during recovery (at least 10% higher than decline)
      const volumeIncreasePercent = ((avgRecoveryVolume - avgDeclineVolume) / avgDeclineVolume) * 100;
      
      if (volumeIncreasePercent < 10) {
        continue;
      }
      
      // Check if pattern is confirmed (price has broken above starting level)
      let confirmed = false;
      for (let i = endIdx + 1; i < priceData.length; i++) {
        if (priceData[i].close > startPrice) {
          confirmed = true;
          break;
        }
      }
      
      // Calculate target price (traditional measure: starting level + pattern height)
      const patternHeight = avgPrice - bottomPrice;
      const targetPrice = startPrice + patternHeight;
      
      // We found a valid rounding bottom pattern
      return {
        type: 'ROUNDING_BOTTOM',
        startIndex: startIdx,
        endIndex: endIdx,
        keyPoints: [
          { index: startIdx, price: startPrice, label: 'start' },
          { index: bottomIdx, price: bottomPrice, label: 'bottom' },
          { index: endIdx, price: endPrice, label: 'end' }
        ],
        neckline: startPrice, // The starting level serves as the breakout level
        confirmed: confirmed,
        confidence: 0, // Will be calculated by calculateConfidence function
        direction: 'BULLISH',
        targetPrice: targetPrice,
        metadata: {
          depthPercent: depthPercent,
          symmetryPercent: symmetryPercent,
          volumeIncreasePercent: volumeIncreasePercent,
          declineLength: declineLength,
          recoveryLength: recoveryLength
        }
      };
    }
  }
  
  // No valid rounding bottom pattern found
  return null;
}

// ============================================================================
// Pattern Recognition Functions - Triangle Patterns
// ============================================================================

/**
 * Detects ascending triangle pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectAscendingTriangle(priceData, pivots) {
  if (!priceData || priceData.length < 15 || !pivots || pivots.length < 4) {
    return null;
  }
  
  // Requirement 6.3: Pattern must span at least 15 bars
  const minPatternLength = 15;
  
  // Separate pivot highs and lows
  const pivotHighs = pivots.filter(p => p.type === 'high');
  const pivotLows = pivots.filter(p => p.type === 'low');
  
  // Requirement 6.1: Need at least 2 peaks for horizontal resistance line
  // Requirement 6.2: Need at least 2 lows for upward-sloping support line
  if (pivotHighs.length < 2 || pivotLows.length < 2) {
    return null;
  }
  
  // Helper function to calculate linear regression slope
  const calculateSlope = (points) => {
    const n = points.length;
    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
    
    for (const point of points) {
      sumX += point.index;
      sumY += point.price;
      sumXY += point.index * point.price;
      sumX2 += point.index * point.index;
    }
    
    const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    return slope;
  };
  
  // Look for horizontal resistance line (at least 2 peaks at similar levels)
  for (let i = 0; i < pivotHighs.length - 1; i++) {
    const resistancePeak1 = pivotHighs[i];
    
    for (let j = i + 1; j < pivotHighs.length; j++) {
      const resistancePeak2 = pivotHighs[j];
      
      // Requirement 6.1: Check if peaks form a horizontal resistance (within 2%)
      const resistanceAvg = (resistancePeak1.price + resistancePeak2.price) / 2;
      const resistanceDiff = Math.abs(resistancePeak1.price - resistancePeak2.price);
      const resistancePercent = (resistanceDiff / resistanceAvg) * 100;
      
      // For horizontal line, allow up to 2% variation
      if (resistancePercent > 2) {
        continue;
      }
      
      // Collect all resistance points that align with this horizontal level
      const resistancePoints = [resistancePeak1, resistancePeak2];
      
      for (let k = 0; k < pivotHighs.length; k++) {
        if (k !== i && k !== j) {
          const testPeak = pivotHighs[k];
          const testDiff = Math.abs(testPeak.price - resistanceAvg);
          const testPercent = (testDiff / resistanceAvg) * 100;
          
          if (testPercent <= 2) {
            resistancePoints.push(testPeak);
          }
        }
      }
      
      // Now look for upward-sloping support line
      // Requirement 6.2: Support line should be formed by at least 2 higher lows
      for (let m = 0; m < pivotLows.length - 1; m++) {
        for (let n = m + 1; n < pivotLows.length; n++) {
          const supportPoints = [pivotLows[m], pivotLows[n]];
          
          // Check if we can add more support points
          for (let p = 0; p < pivotLows.length; p++) {
            if (p !== m && p !== n) {
              const testSlope = calculateSlope([...supportPoints, pivotLows[p]]);
              // Support line must be upward-sloping (positive slope)
              if (testSlope > 0) {
                supportPoints.push(pivotLows[p]);
              }
            }
          }
          
          // Calculate support line slope
          const supportSlope = calculateSlope(supportPoints);
          
          // Requirement 6.2: Support line must be upward-sloping (positive slope)
          if (supportSlope <= 0) {
            continue;
          }
          
          // Determine pattern boundaries
          const allKeyPoints = [...resistancePoints, ...supportPoints];
          allKeyPoints.sort((a, b) => a.index - b.index);
          
          const startIndex = allKeyPoints[0].index;
          const endIndex = allKeyPoints[allKeyPoints.length - 1].index;
          
          // Requirement 6.3: Verify pattern spans at least 15 bars
          const patternLength = endIndex - startIndex;
          if (patternLength < minPatternLength) {
            continue;
          }
          
          // Requirement 6.4: Verify support and resistance lines converge
          // Calculate support line equation: y = supportSlope * x + supportIntercept
          const supportIntercept = supportPoints[0].price - supportSlope * supportPoints[0].index;
          
          // For horizontal resistance, slope is 0, so resistance = resistanceAvg
          // Find where support line reaches resistance level
          // supportSlope * x + supportIntercept = resistanceAvg
          // x = (resistanceAvg - supportIntercept) / supportSlope
          
          if (supportSlope === 0) {
            continue; // Can't converge if support is also horizontal
          }
          
          const convergenceIndex = (resistanceAvg - supportIntercept) / supportSlope;
          
          // Verify convergence happens beyond the pattern end (lines are converging)
          if (convergenceIndex <= endIndex) {
            continue;
          }
          
          // Verify the support line is below the resistance line throughout the pattern
          // Check at start and end of pattern
          const supportAtStart = supportSlope * startIndex + supportIntercept;
          const supportAtEnd = supportSlope * endIndex + supportIntercept;
          
          if (supportAtStart >= resistanceAvg || supportAtEnd >= resistanceAvg) {
            continue;
          }
          
          // Check if pattern is confirmed (price breaks above resistance with volume)
          // Requirement 6.5: Pattern is confirmed when price breaks above resistance with increased volume
          let confirmed = false;
          let breakoutVolume = 0;
          
          // Calculate average volume during pattern
          let patternVolumeSum = 0;
          let patternVolumeCount = 0;
          
          for (let q = startIndex; q <= endIndex; q++) {
            if (priceData[q] && priceData[q].volume !== undefined && isFinite(priceData[q].volume)) {
              patternVolumeSum += priceData[q].volume;
              patternVolumeCount++;
            }
          }
          
          const avgPatternVolume = patternVolumeCount > 0 ? patternVolumeSum / patternVolumeCount : 0;
          
          // Check for breakout after pattern end
          for (let q = endIndex + 1; q < priceData.length; q++) {
            if (priceData[q]) {
              // Check if price breaks above resistance
              if (priceData[q].close > resistanceAvg) {
                // Check if volume is increased (at least 50% above average)
                if (priceData[q].volume && isFinite(priceData[q].volume)) {
                  breakoutVolume = priceData[q].volume;
                  const volumeIncreasePercent = avgPatternVolume > 0 ? 
                    ((breakoutVolume - avgPatternVolume) / avgPatternVolume) * 100 : 0;
                  
                  if (volumeIncreasePercent >= 50) {
                    confirmed = true;
                    break;
                  }
                }
              }
            }
          }
          
          // Calculate target price (traditional measure: resistance + pattern height)
          const patternHeight = resistanceAvg - supportAtEnd;
          const targetPrice = resistanceAvg + patternHeight;
          
          // We found a valid ascending triangle pattern
          return {
            type: 'ASCENDING_TRIANGLE',
            startIndex: startIndex,
            endIndex: endIndex,
            keyPoints: allKeyPoints.map((p, idx) => ({
              index: p.index,
              price: p.price,
              label: p.type === 'high' ? `resistance${idx}` : `support${idx}`
            })),
            neckline: resistanceAvg,
            confirmed: confirmed,
            confidence: 0, // Will be calculated by calculateConfidence function
            direction: 'BULLISH',
            targetPrice: targetPrice,
            metadata: {
              resistanceLevel: resistanceAvg,
              supportSlope: supportSlope,
              supportIntercept: supportIntercept,
              convergencePoint: convergenceIndex,
              convergencePrice: resistanceAvg,
              patternHeight: patternHeight,
              avgPatternVolume: avgPatternVolume,
              breakoutVolume: breakoutVolume
            }
          };
        }
      }
    }
  }
  
  // No valid ascending triangle pattern found
  return null;
}

/**
 * Detects descending triangle pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectDescendingTriangle(priceData, pivots) {
  if (!priceData || priceData.length < 15 || !pivots || pivots.length < 4) {
    return null;
  }
  
  // Requirement 7.3: Pattern must span at least 15 bars
  const minPatternLength = 15;
  
  // Separate pivot highs and lows
  const pivotHighs = pivots.filter(p => p.type === 'high');
  const pivotLows = pivots.filter(p => p.type === 'low');
  
  // Requirement 7.1: Need at least 2 troughs for horizontal support line
  // Requirement 7.2: Need at least 2 highs for downward-sloping resistance line
  if (pivotHighs.length < 2 || pivotLows.length < 2) {
    return null;
  }
  
  // Helper function to calculate linear regression slope
  const calculateSlope = (points) => {
    const n = points.length;
    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
    
    for (const point of points) {
      sumX += point.index;
      sumY += point.price;
      sumXY += point.index * point.price;
      sumX2 += point.index * point.index;
    }
    
    const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    return slope;
  };
  
  // Look for horizontal support line (at least 2 troughs at similar levels)
  for (let i = 0; i < pivotLows.length - 1; i++) {
    const supportTrough1 = pivotLows[i];
    
    for (let j = i + 1; j < pivotLows.length; j++) {
      const supportTrough2 = pivotLows[j];
      
      // Requirement 7.1: Check if troughs form a horizontal support (within 2%)
      const supportAvg = (supportTrough1.price + supportTrough2.price) / 2;
      const supportDiff = Math.abs(supportTrough1.price - supportTrough2.price);
      const supportPercent = (supportDiff / supportAvg) * 100;
      
      // For horizontal line, allow up to 2% variation
      if (supportPercent > 2) {
        continue;
      }
      
      // Collect all support points that align with this horizontal level
      const supportPoints = [supportTrough1, supportTrough2];
      
      for (let k = 0; k < pivotLows.length; k++) {
        if (k !== i && k !== j) {
          const testTrough = pivotLows[k];
          const testDiff = Math.abs(testTrough.price - supportAvg);
          const testPercent = (testDiff / supportAvg) * 100;
          
          if (testPercent <= 2) {
            supportPoints.push(testTrough);
          }
        }
      }
      
      // Now look for downward-sloping resistance line
      // Requirement 7.2: Resistance line should be formed by at least 2 lower highs
      for (let m = 0; m < pivotHighs.length - 1; m++) {
        for (let n = m + 1; n < pivotHighs.length; n++) {
          const resistancePoints = [pivotHighs[m], pivotHighs[n]];
          
          // Check if we can add more resistance points
          for (let p = 0; p < pivotHighs.length; p++) {
            if (p !== m && p !== n) {
              const testSlope = calculateSlope([...resistancePoints, pivotHighs[p]]);
              // Resistance line must be downward-sloping (negative slope)
              if (testSlope < 0) {
                resistancePoints.push(pivotHighs[p]);
              }
            }
          }
          
          // Calculate resistance line slope
          const resistanceSlope = calculateSlope(resistancePoints);
          
          // Requirement 7.2: Resistance line must be downward-sloping (negative slope)
          if (resistanceSlope >= 0) {
            continue;
          }
          
          // Determine pattern boundaries
          const allKeyPoints = [...supportPoints, ...resistancePoints];
          allKeyPoints.sort((a, b) => a.index - b.index);
          
          const startIndex = allKeyPoints[0].index;
          const endIndex = allKeyPoints[allKeyPoints.length - 1].index;
          
          // Requirement 7.3: Verify pattern spans at least 15 bars
          const patternLength = endIndex - startIndex;
          if (patternLength < minPatternLength) {
            continue;
          }
          
          // Requirement 7.4: Verify support and resistance lines converge
          // Calculate resistance line equation: y = resistanceSlope * x + resistanceIntercept
          const resistanceIntercept = resistancePoints[0].price - resistanceSlope * resistancePoints[0].index;
          
          // For horizontal support, slope is 0, so support = supportAvg
          // Find where resistance line reaches support level
          // resistanceSlope * x + resistanceIntercept = supportAvg
          // x = (supportAvg - resistanceIntercept) / resistanceSlope
          
          if (resistanceSlope === 0) {
            continue; // Can't converge if resistance is also horizontal
          }
          
          const convergenceIndex = (supportAvg - resistanceIntercept) / resistanceSlope;
          
          // Verify convergence happens beyond the pattern end (lines are converging)
          if (convergenceIndex <= endIndex) {
            continue;
          }
          
          // Verify the resistance line is above the support line throughout the pattern
          // Check at start and end of pattern
          const resistanceAtStart = resistanceSlope * startIndex + resistanceIntercept;
          const resistanceAtEnd = resistanceSlope * endIndex + resistanceIntercept;
          
          if (resistanceAtStart <= supportAvg || resistanceAtEnd <= supportAvg) {
            continue;
          }
          
          // Check if pattern is confirmed (price breaks below support with volume)
          // Requirement 7.5: Pattern is confirmed when price breaks below support with increased volume
          let confirmed = false;
          let breakoutVolume = 0;
          
          // Calculate average volume during pattern
          let patternVolumeSum = 0;
          let patternVolumeCount = 0;
          
          for (let q = startIndex; q <= endIndex; q++) {
            if (priceData[q] && priceData[q].volume !== undefined && isFinite(priceData[q].volume)) {
              patternVolumeSum += priceData[q].volume;
              patternVolumeCount++;
            }
          }
          
          const avgPatternVolume = patternVolumeCount > 0 ? patternVolumeSum / patternVolumeCount : 0;
          
          // Check for breakout after pattern end
          for (let q = endIndex + 1; q < priceData.length; q++) {
            if (priceData[q]) {
              // Check if price breaks below support
              if (priceData[q].close < supportAvg) {
                // Check if volume is increased (at least 50% above average)
                if (priceData[q].volume && isFinite(priceData[q].volume)) {
                  breakoutVolume = priceData[q].volume;
                  const volumeIncreasePercent = avgPatternVolume > 0 ? 
                    ((breakoutVolume - avgPatternVolume) / avgPatternVolume) * 100 : 0;
                  
                  if (volumeIncreasePercent >= 50) {
                    confirmed = true;
                    break;
                  }
                }
              }
            }
          }
          
          // Calculate target price (traditional measure: support - pattern height)
          const patternHeight = resistanceAtEnd - supportAvg;
          const targetPrice = supportAvg - patternHeight;
          
          // We found a valid descending triangle pattern
          return {
            type: 'DESCENDING_TRIANGLE',
            startIndex: startIndex,
            endIndex: endIndex,
            keyPoints: allKeyPoints.map((p, idx) => ({
              index: p.index,
              price: p.price,
              label: p.type === 'low' ? `support${idx}` : `resistance${idx}`
            })),
            neckline: supportAvg,
            confirmed: confirmed,
            confidence: 0, // Will be calculated by calculateConfidence function
            direction: 'BEARISH',
            targetPrice: targetPrice,
            metadata: {
              supportLevel: supportAvg,
              resistanceSlope: resistanceSlope,
              resistanceIntercept: resistanceIntercept,
              convergencePoint: convergenceIndex,
              convergencePrice: supportAvg,
              patternHeight: patternHeight,
              avgPatternVolume: avgPatternVolume,
              breakoutVolume: breakoutVolume
            }
          };
        }
      }
    }
  }
  
  // No valid descending triangle pattern found
  return null;
}

/**
 * Detects symmetrical triangle pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectSymmetricalTriangle(priceData, pivots) {
  if (!priceData || priceData.length < 15 || !pivots || pivots.length < 4) {
    return null;
  }
  
  // Requirement 8.3: Pattern must span at least 15 bars
  const minPatternLength = 15;
  
  // Separate pivot highs and lows
  const pivotHighs = pivots.filter(p => p.type === 'high');
  const pivotLows = pivots.filter(p => p.type === 'low');
  
  // Requirement 8.1: Need at least 2 lows for upward-sloping support line
  // Requirement 8.2: Need at least 2 highs for downward-sloping resistance line
  if (pivotHighs.length < 2 || pivotLows.length < 2) {
    return null;
  }
  
  // Helper function to calculate linear regression slope
  const calculateSlope = (points) => {
    const n = points.length;
    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
    
    for (const point of points) {
      sumX += point.index;
      sumY += point.price;
      sumXY += point.index * point.price;
      sumX2 += point.index * point.index;
    }
    
    const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    return slope;
  };
  
  // Try different combinations of pivot points to find symmetrical triangle
  // Look for upward-sloping support and downward-sloping resistance
  for (let i = 0; i < pivotLows.length - 1; i++) {
    for (let j = i + 1; j < pivotLows.length; j++) {
      const supportPoints = [pivotLows[i], pivotLows[j]];
      
      // Check if we can add more support points
      for (let k = j + 1; k < pivotLows.length; k++) {
        const testSlope = calculateSlope([...supportPoints, pivotLows[k]]);
        // Support line must be upward-sloping (positive slope)
        if (testSlope > 0) {
          supportPoints.push(pivotLows[k]);
        }
      }
      
      // Calculate support line slope
      const supportSlope = calculateSlope(supportPoints);
      
      // Requirement 8.1: Support line must be upward-sloping (positive slope)
      if (supportSlope <= 0) {
        continue;
      }
      
      // Now look for downward-sloping resistance line
      for (let m = 0; m < pivotHighs.length - 1; m++) {
        for (let n = m + 1; n < pivotHighs.length; n++) {
          const resistancePoints = [pivotHighs[m], pivotHighs[n]];
          
          // Check if we can add more resistance points
          for (let p = n + 1; p < pivotHighs.length; p++) {
            const testSlope = calculateSlope([...resistancePoints, pivotHighs[p]]);
            // Resistance line must be downward-sloping (negative slope)
            if (testSlope < 0) {
              resistancePoints.push(pivotHighs[p]);
            }
          }
          
          // Calculate resistance line slope
          const resistanceSlope = calculateSlope(resistancePoints);
          
          // Requirement 8.2: Resistance line must be downward-sloping (negative slope)
          if (resistanceSlope >= 0) {
            continue;
          }
          
          // Requirement 8.2: Validate both lines have opposite slopes
          // Support should be positive, resistance should be negative
          // For symmetrical triangle, the slopes should be roughly equal in magnitude
          // Allow up to 50% difference in slope magnitude for symmetry
          const slopeMagnitudeRatio = Math.abs(supportSlope / resistanceSlope);
          
          if (slopeMagnitudeRatio < 0.5 || slopeMagnitudeRatio > 2.0) {
            continue; // Slopes are too different, not symmetrical
          }
          
          // Determine pattern boundaries
          const allKeyPoints = [...supportPoints, ...resistancePoints];
          allKeyPoints.sort((a, b) => a.index - b.index);
          
          const startIndex = allKeyPoints[0].index;
          const endIndex = allKeyPoints[allKeyPoints.length - 1].index;
          
          // Requirement 8.3: Verify pattern spans at least 15 bars
          const patternLength = endIndex - startIndex;
          if (patternLength < minPatternLength) {
            continue;
          }
          
          // Requirement 8.4: Verify support and resistance lines converge
          // Calculate line equations
          const supportIntercept = supportPoints[0].price - supportSlope * supportPoints[0].index;
          const resistanceIntercept = resistancePoints[0].price - resistanceSlope * resistancePoints[0].index;
          
          // Find intersection point
          // supportSlope * x + supportIntercept = resistanceSlope * x + resistanceIntercept
          // (supportSlope - resistanceSlope) * x = resistanceIntercept - supportIntercept
          const convergenceIndex = (resistanceIntercept - supportIntercept) / (supportSlope - resistanceSlope);
          
          // Verify convergence happens beyond the pattern end (lines are converging)
          if (convergenceIndex <= endIndex) {
            continue;
          }
          
          // Calculate the price at convergence point
          const convergencePrice = supportSlope * convergenceIndex + supportIntercept;
          
          // Verify the lines don't cross within the pattern
          // Check at start and end of pattern
          const supportAtStart = supportSlope * startIndex + supportIntercept;
          const resistanceAtStart = resistanceSlope * startIndex + resistanceIntercept;
          const supportAtEnd = supportSlope * endIndex + supportIntercept;
          const resistanceAtEnd = resistanceSlope * endIndex + resistanceIntercept;
          
          if (supportAtStart >= resistanceAtStart || supportAtEnd >= resistanceAtEnd) {
            continue;
          }
          
          // Check if pattern is confirmed (price breaks above or below with volume)
          // Requirement 8.5: Pattern is confirmed when price breaks above or below with increased volume
          let confirmed = false;
          let breakoutDirection = null;
          let breakoutVolume = 0;
          
          // Calculate average volume during pattern
          let patternVolumeSum = 0;
          let patternVolumeCount = 0;
          
          for (let q = startIndex; q <= endIndex; q++) {
            if (priceData[q] && priceData[q].volume !== undefined && isFinite(priceData[q].volume)) {
              patternVolumeSum += priceData[q].volume;
              patternVolumeCount++;
            }
          }
          
          const avgPatternVolume = patternVolumeCount > 0 ? patternVolumeSum / patternVolumeCount : 0;
          
          // Check for breakout after pattern end
          for (let q = endIndex + 1; q < priceData.length; q++) {
            if (priceData[q]) {
              // Calculate expected support and resistance at this index
              const expectedSupport = supportSlope * q + supportIntercept;
              const expectedResistance = resistanceSlope * q + resistanceIntercept;
              
              // Check for upward breakout (above resistance)
              if (priceData[q].close > expectedResistance) {
                if (priceData[q].volume && isFinite(priceData[q].volume)) {
                  breakoutVolume = priceData[q].volume;
                  const volumeIncreasePercent = avgPatternVolume > 0 ? 
                    ((breakoutVolume - avgPatternVolume) / avgPatternVolume) * 100 : 0;
                  
                  if (volumeIncreasePercent >= 50) {
                    confirmed = true;
                    breakoutDirection = 'BULLISH';
                    break;
                  }
                }
              }
              // Check for downward breakout (below support)
              else if (priceData[q].close < expectedSupport) {
                if (priceData[q].volume && isFinite(priceData[q].volume)) {
                  breakoutVolume = priceData[q].volume;
                  const volumeIncreasePercent = avgPatternVolume > 0 ? 
                    ((breakoutVolume - avgPatternVolume) / avgPatternVolume) * 100 : 0;
                  
                  if (volumeIncreasePercent >= 50) {
                    confirmed = true;
                    breakoutDirection = 'BEARISH';
                    break;
                  }
                }
              }
            }
          }
          
          // Calculate target price based on breakout direction
          // Traditional measure: pattern height at widest point added/subtracted from breakout level
          const patternHeight = resistanceAtStart - supportAtStart;
          let targetPrice;
          let direction;
          let neckline;
          
          if (breakoutDirection === 'BULLISH') {
            neckline = resistanceAtEnd;
            targetPrice = resistanceAtEnd + patternHeight;
            direction = 'BULLISH';
          } else if (breakoutDirection === 'BEARISH') {
            neckline = supportAtEnd;
            targetPrice = supportAtEnd - patternHeight;
            direction = 'BEARISH';
          } else {
            // No breakout yet, direction is neutral
            neckline = convergencePrice;
            targetPrice = null;
            direction = 'NEUTRAL';
          }
          
          // We found a valid symmetrical triangle pattern
          return {
            type: 'SYMMETRICAL_TRIANGLE',
            startIndex: startIndex,
            endIndex: endIndex,
            keyPoints: allKeyPoints.map((p, idx) => ({
              index: p.index,
              price: p.price,
              label: p.type === 'low' ? `support${idx}` : `resistance${idx}`
            })),
            neckline: neckline,
            confirmed: confirmed,
            confidence: 0, // Will be calculated by calculateConfidence function
            direction: direction,
            targetPrice: targetPrice,
            metadata: {
              supportSlope: supportSlope,
              resistanceSlope: resistanceSlope,
              supportIntercept: supportIntercept,
              resistanceIntercept: resistanceIntercept,
              convergencePoint: convergenceIndex,
              convergencePrice: convergencePrice,
              patternHeight: patternHeight,
              slopeMagnitudeRatio: slopeMagnitudeRatio,
              avgPatternVolume: avgPatternVolume,
              breakoutVolume: breakoutVolume,
              breakoutDirection: breakoutDirection
            }
          };
        }
      }
    }
  }
  
  // No valid symmetrical triangle pattern found
  return null;
}

// ============================================================================
// Pattern Recognition Functions - Continuation Patterns
// ============================================================================

/**
 * Detects flag pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectFlag(priceData, pivots) {
  if (!priceData || priceData.length < 25) {
    return null;
  }
  
  // Requirement 9.3: Consolidation should span between 5 and 20 bars
  const minConsolidationLength = 5;
  const maxConsolidationLength = 20;
  
  // Requirement 9.1: Identify strong price move (the flagpole)
  // Look for a strong directional move (at least 10% in 5-15 bars)
  const minFlagpoleLength = 5;
  const maxFlagpoleLength = 15;
  const minFlagpoleMove = 10; // 10% minimum move
  
  // Scan through price data looking for flagpole + consolidation patterns
  for (let flagpoleStart = 0; flagpoleStart < priceData.length - minFlagpoleLength - minConsolidationLength; flagpoleStart++) {
    
    // Try different flagpole lengths
    for (let flagpoleEnd = flagpoleStart + minFlagpoleLength; 
         flagpoleEnd <= Math.min(flagpoleStart + maxFlagpoleLength, priceData.length - minConsolidationLength - 1); 
         flagpoleEnd++) {
      
      const flagpoleStartPrice = priceData[flagpoleStart].close;
      const flagpoleEndPrice = priceData[flagpoleEnd].close;
      
      // Skip if prices are invalid
      if (!isFinite(flagpoleStartPrice) || !isFinite(flagpoleEndPrice) || flagpoleStartPrice === 0) {
        continue;
      }
      
      // Calculate flagpole move percentage
      const flagpoleMove = ((flagpoleEndPrice - flagpoleStartPrice) / flagpoleStartPrice) * 100;
      
      // Determine flagpole direction
      let flagpoleDirection;
      
      if (flagpoleMove >= minFlagpoleMove) {
        flagpoleDirection = 'BULLISH';
      } else if (flagpoleMove <= -minFlagpoleMove) {
        flagpoleDirection = 'BEARISH';
      } else {
        // Flagpole move is not strong enough
        continue;
      }
      
      // Requirement 9.2: Detect rectangular consolidation pattern that slopes against the trend
      // Look for consolidation period after the flagpole
      for (let consolidationEnd = flagpoleEnd + minConsolidationLength; 
           consolidationEnd <= Math.min(flagpoleEnd + maxConsolidationLength, priceData.length - 1); 
           consolidationEnd++) {
        
        // Calculate consolidation boundaries
        let consolidationHigh = -Infinity;
        let consolidationLow = Infinity;
        
        for (let i = flagpoleEnd + 1; i <= consolidationEnd; i++) {
          if (priceData[i]) {
            if (isFinite(priceData[i].high) && priceData[i].high > consolidationHigh) {
              consolidationHigh = priceData[i].high;
            }
            if (isFinite(priceData[i].low) && priceData[i].low < consolidationLow) {
              consolidationLow = priceData[i].low;
            }
          }
        }
        
        // Skip if consolidation boundaries are invalid
        if (!isFinite(consolidationHigh) || !isFinite(consolidationLow)) {
          continue;
        }
        
        // Calculate consolidation range
        const consolidationRange = consolidationHigh - consolidationLow;
        const consolidationMidpoint = (consolidationHigh + consolidationLow) / 2;
        
        // Skip if consolidation midpoint is zero
        if (consolidationMidpoint === 0) {
          continue;
        }
        
        // Verify consolidation is relatively tight (rectangular)
        // Range should be less than 8% of price
        const consolidationRangePercent = (consolidationRange / consolidationMidpoint) * 100;
        
        if (consolidationRangePercent > 8) {
          continue;
        }
        
        // Requirement 9.4: Verify consolidation retraces no more than 38% of the flagpole
        const flagpoleHeight = Math.abs(flagpoleEndPrice - flagpoleStartPrice);
        
        // Calculate retracement
        let retracement;
        
        if (flagpoleDirection === 'BULLISH') {
          // For bullish flag, consolidation should drift down slightly
          // Retracement is measured from flagpole end to consolidation low
          retracement = flagpoleEndPrice - consolidationLow;
        } else {
          // For bearish flag, consolidation should drift up slightly
          // Retracement is measured from flagpole end to consolidation high
          retracement = consolidationHigh - flagpoleEndPrice;
        }
        
        const retracementPercent = (retracement / flagpoleHeight) * 100;
        
        // Verify retracement is within limits (no more than 38%)
        if (retracementPercent < 0 || retracementPercent > 38) {
          continue;
        }
        
        // Requirement 9.2: Verify consolidation slopes against the trend
        // Calculate slope of consolidation (should be slightly counter to flagpole direction)
        const consolidationStartPrice = priceData[flagpoleEnd + 1].close;
        const consolidationEndPrice = priceData[consolidationEnd].close;
        
        if (!isFinite(consolidationStartPrice) || !isFinite(consolidationEndPrice) || consolidationStartPrice === 0) {
          continue;
        }
        
        const consolidationSlope = ((consolidationEndPrice - consolidationStartPrice) / consolidationStartPrice) * 100;
        
        // For bullish flag, consolidation should slope down (negative slope)
        // For bearish flag, consolidation should slope up (positive slope)
        if (flagpoleDirection === 'BULLISH' && consolidationSlope > 0) {
          continue; // Consolidation is not sloping against the trend
        }
        if (flagpoleDirection === 'BEARISH' && consolidationSlope < 0) {
          continue; // Consolidation is not sloping against the trend
        }
        
        // Verify consolidation slope is not too steep (should be relatively flat)
        if (Math.abs(consolidationSlope) > 5) {
          continue;
        }
        
        // Requirement 9.5: Check if pattern is confirmed (breakout in direction of flagpole with volume)
        let confirmed = false;
        let breakoutVolume = 0;
        
        // Calculate average volume during consolidation
        let consolidationVolumeSum = 0;
        let consolidationVolumeCount = 0;
        
        for (let i = flagpoleEnd + 1; i <= consolidationEnd; i++) {
          if (priceData[i] && priceData[i].volume !== undefined && isFinite(priceData[i].volume)) {
            consolidationVolumeSum += priceData[i].volume;
            consolidationVolumeCount++;
          }
        }
        
        const avgConsolidationVolume = consolidationVolumeCount > 0 ? 
          consolidationVolumeSum / consolidationVolumeCount : 0;
        
        // Check for breakout after consolidation end
        for (let i = consolidationEnd + 1; i < priceData.length; i++) {
          if (priceData[i]) {
            let breakoutOccurred = false;
            
            // Check for breakout in flagpole direction
            if (flagpoleDirection === 'BULLISH' && priceData[i].close > consolidationHigh) {
              breakoutOccurred = true;
            } else if (flagpoleDirection === 'BEARISH' && priceData[i].close < consolidationLow) {
              breakoutOccurred = true;
            }
            
            if (breakoutOccurred) {
              // Check volume confirmation
              if (priceData[i].volume && isFinite(priceData[i].volume)) {
                breakoutVolume = priceData[i].volume;
                const volumeIncreasePercent = avgConsolidationVolume > 0 ? 
                  ((breakoutVolume - avgConsolidationVolume) / avgConsolidationVolume) * 100 : 0;
                
                if (volumeIncreasePercent >= 50) {
                  confirmed = true;
                  break;
                }
              }
            }
          }
        }
        
        // Calculate target price (traditional measure: flagpole height added to breakout level)
        let targetPrice;
        let neckline;
        
        if (flagpoleDirection === 'BULLISH') {
          neckline = consolidationHigh;
          targetPrice = consolidationHigh + flagpoleHeight;
        } else {
          neckline = consolidationLow;
          targetPrice = consolidationLow - flagpoleHeight;
        }
        
        // We found a valid flag pattern
        return {
          type: 'FLAG',
          startIndex: flagpoleStart,
          endIndex: consolidationEnd,
          keyPoints: [
            { index: flagpoleStart, price: flagpoleStartPrice, label: 'flagpoleStart' },
            { index: flagpoleEnd, price: flagpoleEndPrice, label: 'flagpoleEnd' },
            { index: consolidationEnd, price: consolidationEndPrice, label: 'consolidationEnd' }
          ],
          neckline: neckline,
          confirmed: confirmed,
          confidence: 0, // Will be calculated by calculateConfidence function
          direction: flagpoleDirection,
          targetPrice: targetPrice,
          metadata: {
            flagpoleLength: flagpoleEnd - flagpoleStart,
            flagpoleMove: flagpoleMove,
            flagpoleHeight: flagpoleHeight,
            consolidationLength: consolidationEnd - flagpoleEnd,
            consolidationHigh: consolidationHigh,
            consolidationLow: consolidationLow,
            consolidationRange: consolidationRange,
            consolidationRangePercent: consolidationRangePercent,
            consolidationSlope: consolidationSlope,
            retracement: retracement,
            retracementPercent: retracementPercent,
            avgConsolidationVolume: avgConsolidationVolume,
            breakoutVolume: breakoutVolume
          }
        };
      }
    }
  }
  
  // No valid flag pattern found
  return null;
}

/**
 * Detects pennant pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectPennant(priceData, pivots) {
  if (!priceData || priceData.length < 25 || !pivots || pivots.length < 4) {
    return null;
  }
  
  // Requirement 10.3: Consolidation should span between 5 and 20 bars
  const minConsolidationLength = 5;
  const maxConsolidationLength = 20;
  
  // Requirement 10.1: Identify strong price move (the flagpole)
  // Look for a strong directional move (at least 10% in 5-15 bars)
  const minFlagpoleLength = 5;
  const maxFlagpoleLength = 15;
  const minFlagpoleMove = 10; // 10% minimum move
  
  // Helper function to calculate linear regression slope
  const calculateSlope = (points) => {
    const n = points.length;
    if (n < 2) return 0;
    
    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
    
    for (const point of points) {
      sumX += point.index;
      sumY += point.price;
      sumXY += point.index * point.price;
      sumX2 += point.index * point.index;
    }
    
    const denominator = n * sumX2 - sumX * sumX;
    if (denominator === 0) return 0;
    
    const slope = (n * sumXY - sumX * sumY) / denominator;
    return slope;
  };
  
  // Scan through price data looking for flagpole + pennant patterns
  for (let flagpoleStart = 0; flagpoleStart < priceData.length - minFlagpoleLength - minConsolidationLength; flagpoleStart++) {
    
    // Try different flagpole lengths
    for (let flagpoleEnd = flagpoleStart + minFlagpoleLength; 
         flagpoleEnd <= Math.min(flagpoleStart + maxFlagpoleLength, priceData.length - minConsolidationLength - 1); 
         flagpoleEnd++) {
      
      const flagpoleStartPrice = priceData[flagpoleStart].close;
      const flagpoleEndPrice = priceData[flagpoleEnd].close;
      
      // Skip if prices are invalid
      if (!isFinite(flagpoleStartPrice) || !isFinite(flagpoleEndPrice) || flagpoleStartPrice === 0) {
        continue;
      }
      
      // Calculate flagpole move percentage
      const flagpoleMove = ((flagpoleEndPrice - flagpoleStartPrice) / flagpoleStartPrice) * 100;
      
      // Determine flagpole direction
      let flagpoleDirection;
      
      if (flagpoleMove >= minFlagpoleMove) {
        flagpoleDirection = 'BULLISH';
      } else if (flagpoleMove <= -minFlagpoleMove) {
        flagpoleDirection = 'BEARISH';
      } else {
        // Flagpole move is not strong enough
        continue;
      }
      
      // Requirement 10.2: Detect converging trendlines forming a small symmetrical triangle
      // Look for consolidation period after the flagpole
      for (let consolidationEnd = flagpoleEnd + minConsolidationLength; 
           consolidationEnd <= Math.min(flagpoleEnd + maxConsolidationLength, priceData.length - 1); 
           consolidationEnd++) {
        
        const consolidationLength = consolidationEnd - flagpoleEnd;
        
        // Get pivots within the consolidation period
        const consolidationPivots = pivots.filter(p => 
          p.index > flagpoleEnd && p.index <= consolidationEnd
        );
        
        // Need at least 2 highs and 2 lows to form converging trendlines
        const consolidationHighs = consolidationPivots.filter(p => p.type === 'high');
        const consolidationLows = consolidationPivots.filter(p => p.type === 'low');
        
        if (consolidationHighs.length < 2 || consolidationLows.length < 2) {
          continue;
        }
        
        // Calculate resistance trendline (connecting highs)
        const resistanceSlope = calculateSlope(consolidationHighs);
        
        // Calculate support trendline (connecting lows)
        const supportSlope = calculateSlope(consolidationLows);
        
        // For a pennant (symmetrical triangle), we need:
        // - Resistance line sloping down (negative slope)
        // - Support line sloping up (positive slope)
        // - Lines converging (opposite slopes)
        
        if (resistanceSlope >= 0 || supportSlope <= 0) {
          continue; // Lines are not converging properly
        }
        
        // Verify slopes are roughly symmetrical (within 50% of each other in magnitude)
        const slopeMagnitudeRatio = Math.abs(supportSlope / resistanceSlope);
        
        if (slopeMagnitudeRatio < 0.5 || slopeMagnitudeRatio > 2.0) {
          continue; // Slopes are too different, not symmetrical
        }
        
        // Calculate trendline intercepts
        const resistanceIntercept = consolidationHighs[0].price - resistanceSlope * consolidationHighs[0].index;
        const supportIntercept = consolidationLows[0].price - supportSlope * consolidationLows[0].index;
        
        // Find convergence point
        const convergenceIndex = (resistanceIntercept - supportIntercept) / (supportSlope - resistanceSlope);
        
        // Verify convergence happens beyond the consolidation end
        if (convergenceIndex <= consolidationEnd) {
          continue;
        }
        
        // Calculate consolidation boundaries at the end of the pattern
        const resistanceAtEnd = resistanceSlope * consolidationEnd + resistanceIntercept;
        const supportAtEnd = supportSlope * consolidationEnd + supportIntercept;
        
        // Verify the lines haven't crossed
        if (supportAtEnd >= resistanceAtEnd) {
          continue;
        }
        
        // Requirement 10.4: Verify consolidation retraces no more than 38% of the flagpole
        const flagpoleHeight = Math.abs(flagpoleEndPrice - flagpoleStartPrice);
        
        // Calculate retracement (maximum pullback during consolidation)
        let maxRetracement = 0;
        
        for (let i = flagpoleEnd + 1; i <= consolidationEnd; i++) {
          if (priceData[i]) {
            let retracement;
            
            if (flagpoleDirection === 'BULLISH') {
              // For bullish pennant, measure pullback from flagpole end
              retracement = flagpoleEndPrice - priceData[i].low;
            } else {
              // For bearish pennant, measure pullback from flagpole end
              retracement = priceData[i].high - flagpoleEndPrice;
            }
            
            if (isFinite(retracement) && retracement > maxRetracement) {
              maxRetracement = retracement;
            }
          }
        }
        
        const retracementPercent = (maxRetracement / flagpoleHeight) * 100;
        
        // Verify retracement is within limits (no more than 38%)
        if (retracementPercent < 0 || retracementPercent > 38) {
          continue;
        }
        
        // Requirement 10.5: Check if pattern is confirmed (breakout in direction of flagpole with volume)
        let confirmed = false;
        let breakoutVolume = 0;
        
        // Calculate average volume during consolidation
        let consolidationVolumeSum = 0;
        let consolidationVolumeCount = 0;
        
        for (let i = flagpoleEnd + 1; i <= consolidationEnd; i++) {
          if (priceData[i] && priceData[i].volume !== undefined && isFinite(priceData[i].volume)) {
            consolidationVolumeSum += priceData[i].volume;
            consolidationVolumeCount++;
          }
        }
        
        const avgConsolidationVolume = consolidationVolumeCount > 0 ? 
          consolidationVolumeSum / consolidationVolumeCount : 0;
        
        // Check for breakout after consolidation end
        for (let i = consolidationEnd + 1; i < priceData.length; i++) {
          if (priceData[i]) {
            // Calculate expected resistance and support at this index
            const expectedResistance = resistanceSlope * i + resistanceIntercept;
            const expectedSupport = supportSlope * i + supportIntercept;
            
            let breakoutOccurred = false;
            
            // Check for breakout in flagpole direction
            if (flagpoleDirection === 'BULLISH' && priceData[i].close > expectedResistance) {
              breakoutOccurred = true;
            } else if (flagpoleDirection === 'BEARISH' && priceData[i].close < expectedSupport) {
              breakoutOccurred = true;
            }
            
            if (breakoutOccurred) {
              // Check volume confirmation
              if (priceData[i].volume && isFinite(priceData[i].volume)) {
                breakoutVolume = priceData[i].volume;
                const volumeIncreasePercent = avgConsolidationVolume > 0 ? 
                  ((breakoutVolume - avgConsolidationVolume) / avgConsolidationVolume) * 100 : 0;
                
                if (volumeIncreasePercent >= 50) {
                  confirmed = true;
                  break;
                }
              }
            }
          }
        }
        
        // Calculate target price (traditional measure: flagpole height added to breakout level)
        let targetPrice;
        let neckline;
        
        if (flagpoleDirection === 'BULLISH') {
          neckline = resistanceAtEnd;
          targetPrice = resistanceAtEnd + flagpoleHeight;
        } else {
          neckline = supportAtEnd;
          targetPrice = supportAtEnd - flagpoleHeight;
        }
        
        // We found a valid pennant pattern
        return {
          type: 'PENNANT',
          startIndex: flagpoleStart,
          endIndex: consolidationEnd,
          keyPoints: [
            { index: flagpoleStart, price: flagpoleStartPrice, label: 'flagpoleStart' },
            { index: flagpoleEnd, price: flagpoleEndPrice, label: 'flagpoleEnd' },
            ...consolidationHighs.map((p, idx) => ({ 
              index: p.index, 
              price: p.price, 
              label: `resistance${idx}` 
            })),
            ...consolidationLows.map((p, idx) => ({ 
              index: p.index, 
              price: p.price, 
              label: `support${idx}` 
            }))
          ],
          neckline: neckline,
          confirmed: confirmed,
          confidence: 0, // Will be calculated by calculateConfidence function
          direction: flagpoleDirection,
          targetPrice: targetPrice,
          metadata: {
            flagpoleLength: flagpoleEnd - flagpoleStart,
            flagpoleMove: flagpoleMove,
            flagpoleHeight: flagpoleHeight,
            consolidationLength: consolidationLength,
            resistanceSlope: resistanceSlope,
            supportSlope: supportSlope,
            resistanceIntercept: resistanceIntercept,
            supportIntercept: supportIntercept,
            convergencePoint: convergenceIndex,
            maxRetracement: maxRetracement,
            retracementPercent: retracementPercent,
            avgConsolidationVolume: avgConsolidationVolume,
            breakoutVolume: breakoutVolume
          }
        };
      }
    }
  }
  
  // No valid pennant pattern found
  return null;
}

// ============================================================================
// Pattern Recognition Functions - Wedge Patterns
// ============================================================================

/**
 * Detects rising wedge pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectRisingWedge(priceData, pivots) {
  if (!priceData || priceData.length < 15 || !pivots || pivots.length < 4) {
    return null;
  }
  
  // Requirement 11.3: Pattern must span at least 15 bars
  const minPatternLength = 15;
  
  // Separate pivot highs and lows
  const pivotHighs = pivots.filter(p => p.type === 'high');
  const pivotLows = pivots.filter(p => p.type === 'low');
  
  // Requirement 11.1: Need at least 2 highs and 2 lows to form trendlines
  if (pivotHighs.length < 2 || pivotLows.length < 2) {
    return null;
  }
  
  // Helper function to calculate linear regression slope
  const calculateSlope = (points) => {
    const n = points.length;
    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
    
    for (const point of points) {
      sumX += point.index;
      sumY += point.price;
      sumXY += point.index * point.price;
      sumX2 += point.index * point.index;
    }
    
    const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    return slope;
  };
  
  // Try different combinations of pivot points to find rising wedge
  // Look for at least 2 lows and 2 highs that form upward-sloping lines
  for (let i = 0; i < pivotLows.length - 1; i++) {
    for (let j = i + 1; j < pivotLows.length; j++) {
      const supportPoints = [pivotLows[i], pivotLows[j]];
      
      // Check if we can add more support points
      for (let k = j + 1; k < pivotLows.length; k++) {
        // Check if this point roughly aligns with the support line
        const testSlope = calculateSlope([...supportPoints, pivotLows[k]]);
        if (testSlope > 0) { // Must be upward sloping
          supportPoints.push(pivotLows[k]);
        }
      }
      
      // Calculate support line slope
      const supportSlope = calculateSlope(supportPoints);
      
      // Requirement 11.1: Support line must be upward-sloping
      if (supportSlope <= 0) {
        continue;
      }
      
      // Now look for resistance line (also upward-sloping)
      for (let m = 0; m < pivotHighs.length - 1; m++) {
        for (let n = m + 1; n < pivotHighs.length; n++) {
          const resistancePoints = [pivotHighs[m], pivotHighs[n]];
          
          // Check if we can add more resistance points
          for (let p = n + 1; p < pivotHighs.length; p++) {
            const testSlope = calculateSlope([...resistancePoints, pivotHighs[p]]);
            if (testSlope > 0) { // Must be upward sloping
              resistancePoints.push(pivotHighs[p]);
            }
          }
          
          // Calculate resistance line slope
          const resistanceSlope = calculateSlope(resistancePoints);
          
          // Requirement 11.1: Resistance line must be upward-sloping
          if (resistanceSlope <= 0) {
            continue;
          }
          
          // Requirement 11.2: Support slope must be steeper than resistance slope
          if (supportSlope <= resistanceSlope) {
            continue;
          }
          
          // Determine pattern boundaries
          const allKeyPoints = [...supportPoints, ...resistancePoints];
          allKeyPoints.sort((a, b) => a.index - b.index);
          
          const startIndex = allKeyPoints[0].index;
          const endIndex = allKeyPoints[allKeyPoints.length - 1].index;
          
          // Requirement 11.3: Verify pattern spans at least 15 bars
          const patternLength = endIndex - startIndex;
          if (patternLength < minPatternLength) {
            continue;
          }
          
          // Requirement 11.4: Verify trendlines converge
          // Calculate where the lines would intersect
          // Support line: y = supportSlope * x + supportIntercept
          // Resistance line: y = resistanceSlope * x + resistanceIntercept
          
          // Calculate intercepts using first point of each line
          const supportIntercept = supportPoints[0].price - supportSlope * supportPoints[0].index;
          const resistanceIntercept = resistancePoints[0].price - resistanceSlope * resistancePoints[0].index;
          
          // Find intersection point: supportSlope * x + supportIntercept = resistanceSlope * x + resistanceIntercept
          // (supportSlope - resistanceSlope) * x = resistanceIntercept - supportIntercept
          const convergenceIndex = (resistanceIntercept - supportIntercept) / (supportSlope - resistanceSlope);
          
          // Verify convergence happens beyond the pattern end (lines are converging, not diverging)
          if (convergenceIndex <= endIndex) {
            continue;
          }
          
          // Calculate the price at convergence point
          const convergencePrice = supportSlope * convergenceIndex + supportIntercept;
          
          // Check if pattern is confirmed (price breaks below support line)
          // Requirement 11.5: Pattern is confirmed when price breaks below support
          let confirmed = false;
          
          for (let q = endIndex + 1; q < priceData.length; q++) {
            if (priceData[q]) {
              // Calculate expected support price at this index
              const expectedSupportPrice = supportSlope * q + supportIntercept;
              
              if (priceData[q].close < expectedSupportPrice) {
                confirmed = true;
                break;
              }
            }
          }
          
          // Calculate target price (traditional measure: support level at breakout - pattern height)
          const patternHeight = (resistanceSlope * endIndex + resistanceIntercept) - 
                                (supportSlope * endIndex + supportIntercept);
          const supportAtEnd = supportSlope * endIndex + supportIntercept;
          const targetPrice = supportAtEnd - patternHeight;
          
          // We found a valid rising wedge pattern
          return {
            type: 'RISING_WEDGE',
            startIndex: startIndex,
            endIndex: endIndex,
            keyPoints: allKeyPoints.map((p, idx) => ({
              index: p.index,
              price: p.price,
              label: p.type === 'low' ? `support${idx}` : `resistance${idx}`
            })),
            neckline: supportAtEnd, // Support level at pattern end
            confirmed: confirmed,
            confidence: 0, // Will be calculated by calculateConfidence function
            direction: 'BEARISH',
            targetPrice: targetPrice,
            metadata: {
              supportSlope: supportSlope,
              resistanceSlope: resistanceSlope,
              convergencePoint: convergenceIndex,
              convergencePrice: convergencePrice,
              patternHeight: patternHeight,
              supportIntercept: supportIntercept,
              resistanceIntercept: resistanceIntercept
            }
          };
        }
      }
    }
  }
  
  // No valid rising wedge pattern found
  return null;
}

/**
 * Detects falling wedge pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectFallingWedge(priceData, pivots) {
  if (!priceData || priceData.length < 15 || !pivots || pivots.length < 4) {
    return null;
  }
  
  // Requirement 12.3: Pattern must span at least 15 bars
  const minPatternLength = 15;
  
  // Separate pivot highs and lows
  const pivotHighs = pivots.filter(p => p.type === 'high');
  const pivotLows = pivots.filter(p => p.type === 'low');
  
  // Requirement 12.1: Need at least 2 highs and 2 lows to form trendlines
  if (pivotHighs.length < 2 || pivotLows.length < 2) {
    return null;
  }
  
  // Helper function to calculate linear regression slope
  const calculateSlope = (points) => {
    const n = points.length;
    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
    
    for (const point of points) {
      sumX += point.index;
      sumY += point.price;
      sumXY += point.index * point.price;
      sumX2 += point.index * point.index;
    }
    
    const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    return slope;
  };
  
  // Try different combinations of pivot points to find falling wedge
  // Look for at least 2 lows and 2 highs that form downward-sloping lines
  for (let i = 0; i < pivotLows.length - 1; i++) {
    for (let j = i + 1; j < pivotLows.length; j++) {
      const supportPoints = [pivotLows[i], pivotLows[j]];
      
      // Check if we can add more support points
      for (let k = j + 1; k < pivotLows.length; k++) {
        // Check if this point roughly aligns with the support line
        const testSlope = calculateSlope([...supportPoints, pivotLows[k]]);
        if (testSlope < 0) { // Must be downward sloping
          supportPoints.push(pivotLows[k]);
        }
      }
      
      // Calculate support line slope
      const supportSlope = calculateSlope(supportPoints);
      
      // Requirement 12.1: Support line must be downward-sloping
      if (supportSlope >= 0) {
        continue;
      }
      
      // Now look for resistance line (also downward-sloping)
      for (let m = 0; m < pivotHighs.length - 1; m++) {
        for (let n = m + 1; n < pivotHighs.length; n++) {
          const resistancePoints = [pivotHighs[m], pivotHighs[n]];
          
          // Check if we can add more resistance points
          for (let p = n + 1; p < pivotHighs.length; p++) {
            const testSlope = calculateSlope([...resistancePoints, pivotHighs[p]]);
            if (testSlope < 0) { // Must be downward sloping
              resistancePoints.push(pivotHighs[p]);
            }
          }
          
          // Calculate resistance line slope
          const resistanceSlope = calculateSlope(resistancePoints);
          
          // Requirement 12.1: Resistance line must be downward-sloping
          if (resistanceSlope >= 0) {
            continue;
          }
          
          // Requirement 12.2: Resistance slope must be steeper (more negative) than support slope
          // Since both are negative, resistance should be more negative (steeper decline)
          if (resistanceSlope >= supportSlope) {
            continue;
          }
          
          // Determine pattern boundaries
          const allKeyPoints = [...supportPoints, ...resistancePoints];
          allKeyPoints.sort((a, b) => a.index - b.index);
          
          const startIndex = allKeyPoints[0].index;
          const endIndex = allKeyPoints[allKeyPoints.length - 1].index;
          
          // Requirement 12.3: Verify pattern spans at least 15 bars
          const patternLength = endIndex - startIndex;
          if (patternLength < minPatternLength) {
            continue;
          }
          
          // Requirement 12.4: Verify trendlines converge
          // Calculate where the lines would intersect
          // Support line: y = supportSlope * x + supportIntercept
          // Resistance line: y = resistanceSlope * x + resistanceIntercept
          
          // Calculate intercepts using first point of each line
          const supportIntercept = supportPoints[0].price - supportSlope * supportPoints[0].index;
          const resistanceIntercept = resistancePoints[0].price - resistanceSlope * resistancePoints[0].index;
          
          // Find intersection point: supportSlope * x + supportIntercept = resistanceSlope * x + resistanceIntercept
          // (supportSlope - resistanceSlope) * x = resistanceIntercept - supportIntercept
          const convergenceIndex = (resistanceIntercept - supportIntercept) / (supportSlope - resistanceSlope);
          
          // Verify convergence happens beyond the pattern end (lines are converging, not diverging)
          if (convergenceIndex <= endIndex) {
            continue;
          }
          
          // Calculate the price at convergence point
          const convergencePrice = supportSlope * convergenceIndex + supportIntercept;
          
          // Check if pattern is confirmed (price breaks above resistance line)
          // Requirement 12.5: Pattern is confirmed when price breaks above resistance
          let confirmed = false;
          
          for (let q = endIndex + 1; q < priceData.length; q++) {
            if (priceData[q]) {
              // Calculate expected resistance price at this index
              const expectedResistancePrice = resistanceSlope * q + resistanceIntercept;
              
              if (priceData[q].close > expectedResistancePrice) {
                confirmed = true;
                break;
              }
            }
          }
          
          // Calculate target price (traditional measure: resistance level at breakout + pattern height)
          const patternHeight = (resistanceSlope * endIndex + resistanceIntercept) - 
                                (supportSlope * endIndex + supportIntercept);
          const resistanceAtEnd = resistanceSlope * endIndex + resistanceIntercept;
          const targetPrice = resistanceAtEnd + patternHeight;
          
          // We found a valid falling wedge pattern
          return {
            type: 'FALLING_WEDGE',
            startIndex: startIndex,
            endIndex: endIndex,
            keyPoints: allKeyPoints.map((p, idx) => ({
              index: p.index,
              price: p.price,
              label: p.type === 'low' ? `support${idx}` : `resistance${idx}`
            })),
            neckline: resistanceAtEnd, // Resistance level at pattern end
            confirmed: confirmed,
            confidence: 0, // Will be calculated by calculateConfidence function
            direction: 'BULLISH',
            targetPrice: targetPrice,
            metadata: {
              supportSlope: supportSlope,
              resistanceSlope: resistanceSlope,
              convergencePoint: convergenceIndex,
              convergencePrice: convergencePrice,
              patternHeight: patternHeight,
              supportIntercept: supportIntercept,
              resistanceIntercept: resistanceIntercept
            }
          };
        }
      }
    }
  }
  
  // No valid falling wedge pattern found
  return null;
}

// ============================================================================
// Pattern Recognition Functions - Consolidation Patterns
// ============================================================================

/**
 * Detects rectangle pattern
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectRectangle(priceData, pivots) {
  if (!priceData || priceData.length < 15 || !pivots || pivots.length < 4) {
    return null;
  }
  
  // Requirement 13.3: Pattern must span at least 15 bars
  const minPatternLength = 15;
  
  // Separate pivot highs and lows
  const pivotHighs = pivots.filter(p => p.type === 'high');
  const pivotLows = pivots.filter(p => p.type === 'low');
  
  // Requirement 13.1: Need at least 2 peaks for resistance line
  // Requirement 13.2: Need at least 2 troughs for support line
  if (pivotHighs.length < 2 || pivotLows.length < 2) {
    return null;
  }
  
  // Look for horizontal resistance and support lines
  // Try different combinations of pivot highs to find horizontal resistance
  for (let i = 0; i < pivotHighs.length - 1; i++) {
    const resistancePeak1 = pivotHighs[i];
    
    for (let j = i + 1; j < pivotHighs.length; j++) {
      const resistancePeak2 = pivotHighs[j];
      
      // Check if peaks are at approximately the same level (within 2% for horizontal line)
      const resistanceAvg = (resistancePeak1.price + resistancePeak2.price) / 2;
      const resistanceDiff = Math.abs(resistancePeak1.price - resistancePeak2.price);
      const resistancePercent = (resistanceDiff / resistanceAvg) * 100;
      
      if (resistancePercent > 2) {
        continue;
      }
      
      // Now look for horizontal support line
      for (let k = 0; k < pivotLows.length - 1; k++) {
        const supportTrough1 = pivotLows[k];
        
        for (let l = k + 1; l < pivotLows.length; l++) {
          const supportTrough2 = pivotLows[l];
          
          // Check if troughs are at approximately the same level (within 2% for horizontal line)
          const supportAvg = (supportTrough1.price + supportTrough2.price) / 2;
          const supportDiff = Math.abs(supportTrough1.price - supportTrough2.price);
          const supportPercent = (supportDiff / supportAvg) * 100;
          
          if (supportPercent > 2) {
            continue;
          }
          
          // Determine the pattern boundaries
          const allKeyPoints = [resistancePeak1, resistancePeak2, supportTrough1, supportTrough2];
          allKeyPoints.sort((a, b) => a.index - b.index);
          
          const startIndex = allKeyPoints[0].index;
          const endIndex = allKeyPoints[allKeyPoints.length - 1].index;
          
          // Requirement 13.3: Verify pattern spans at least 15 bars
          const patternLength = endIndex - startIndex;
          if (patternLength < minPatternLength) {
            continue;
          }
          
          // Requirement 13.4: Verify support and resistance lines are approximately parallel
          // For horizontal lines, they should both be horizontal (already checked above)
          // Also verify the distance between them is reasonable (at least 5% of price range)
          const rangePercent = ((resistanceAvg - supportAvg) / resistanceAvg) * 100;
          if (rangePercent < 5) {
            continue;
          }
          
          // Verify that the price stays within the rectangle boundaries
          // Check that most bars between start and end stay within the range
          let barsInRange = 0;
          let totalBars = 0;
          
          for (let m = startIndex; m <= endIndex; m++) {
            if (priceData[m]) {
              totalBars++;
              // Check if the bar's high and low are within the rectangle
              if (priceData[m].low >= supportAvg * 0.98 && priceData[m].high <= resistanceAvg * 1.02) {
                barsInRange++;
              }
            }
          }
          
          // At least 70% of bars should stay within the rectangle
          const percentInRange = (barsInRange / totalBars) * 100;
          if (percentInRange < 70) {
            continue;
          }
          
          // Check if pattern is confirmed (breakout above resistance or below support)
          let confirmed = false;
          let breakoutDirection = null;
          
          for (let m = endIndex + 1; m < priceData.length; m++) {
            if (priceData[m]) {
              // Check for breakout above resistance
              if (priceData[m].close > resistanceAvg) {
                confirmed = true;
                breakoutDirection = 'BULLISH';
                break;
              }
              // Check for breakout below support
              if (priceData[m].close < supportAvg) {
                confirmed = true;
                breakoutDirection = 'BEARISH';
                break;
              }
            }
          }
          
          // Calculate target price based on breakout direction
          const rectangleHeight = resistanceAvg - supportAvg;
          let targetPrice;
          let direction;
          
          if (breakoutDirection === 'BULLISH') {
            targetPrice = resistanceAvg + rectangleHeight;
            direction = 'BULLISH';
          } else if (breakoutDirection === 'BEARISH') {
            targetPrice = supportAvg - rectangleHeight;
            direction = 'BEARISH';
          } else {
            // No breakout yet, direction is neutral
            targetPrice = null;
            direction = 'NEUTRAL';
          }
          
          // We found a valid rectangle pattern
          return {
            type: 'RECTANGLE',
            startIndex: startIndex,
            endIndex: endIndex,
            keyPoints: [
              { index: resistancePeak1.index, price: resistancePeak1.price, label: 'resistance1' },
              { index: resistancePeak2.index, price: resistancePeak2.price, label: 'resistance2' },
              { index: supportTrough1.index, price: supportTrough1.price, label: 'support1' },
              { index: supportTrough2.index, price: supportTrough2.price, label: 'support2' }
            ],
            neckline: breakoutDirection === 'BULLISH' ? resistanceAvg : supportAvg,
            confirmed: confirmed,
            confidence: 0, // Will be calculated by calculateConfidence function
            direction: direction,
            targetPrice: targetPrice,
            metadata: {
              resistanceLevel: resistanceAvg,
              supportLevel: supportAvg,
              rectangleHeight: rectangleHeight,
              rangePercent: rangePercent,
              percentInRange: percentInRange,
              breakoutDirection: breakoutDirection
            }
          };
        }
      }
    }
  }
  
  // No valid rectangle pattern found
  return null;
}

// ============================================================================
// Pattern Recognition Functions - Breakout and Gap Detection
// ============================================================================

/**
 * Detects breakout from consolidation patterns
 * @param {Array<Object>} priceData - OHLCV data
 * @param {Array<Object>} pivots - Pivot points from PivotPointFinder
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectBreakout(priceData, pivots) {
  if (!priceData || priceData.length < 20) {
    return null;
  }
  
  // Requirement 15.1: Identify periods of price consolidation
  // We'll look for periods where price stays within a narrow range
  const minConsolidationLength = 10; // Minimum bars for consolidation
  const maxConsolidationLength = 50; // Maximum bars to look back
  const consolidationThreshold = 0.05; // 5% range for consolidation
  
  // Scan through price data looking for consolidation periods followed by breakouts
  for (let consolidationStart = 0; consolidationStart < priceData.length - minConsolidationLength - 1; consolidationStart++) {
    
    // Try different consolidation lengths
    for (let consolidationEnd = consolidationStart + minConsolidationLength; 
         consolidationEnd <= Math.min(consolidationStart + maxConsolidationLength, priceData.length - 2); 
         consolidationEnd++) {
      
      // Calculate the consolidation range
      let consolidationHigh = -Infinity;
      let consolidationLow = Infinity;
      
      for (let i = consolidationStart; i <= consolidationEnd; i++) {
        if (priceData[i]) {
          if (priceData[i].high > consolidationHigh) {
            consolidationHigh = priceData[i].high;
          }
          if (priceData[i].low < consolidationLow) {
            consolidationLow = priceData[i].low;
          }
        }
      }
      
      // Requirement 15.2: Calculate the consolidation range
      const consolidationRange = consolidationHigh - consolidationLow;
      const consolidationMidpoint = (consolidationHigh + consolidationLow) / 2;
      const rangePercent = (consolidationRange / consolidationMidpoint) * 100;
      
      // Check if this is a valid consolidation (price stays within narrow range)
      // Range should be relatively small (less than 10% of price)
      if (rangePercent > 10) {
        continue;
      }
      
      // Requirement 15.3: Detect price movement beyond the consolidation range
      // Check the bars immediately after consolidation for a breakout
      const breakoutIndex = consolidationEnd + 1;
      
      if (breakoutIndex >= priceData.length) {
        continue;
      }
      
      const breakoutBar = priceData[breakoutIndex];
      
      if (!breakoutBar) {
        continue;
      }
      
      // Determine if there's a breakout (price moves beyond consolidation range)
      let breakoutDirection = null;
      let breakoutPrice = null;
      
      // Check for upward breakout (close above consolidation high)
      if (breakoutBar.close > consolidationHigh) {
        breakoutDirection = 'BULLISH';
        breakoutPrice = breakoutBar.close;
      }
      // Check for downward breakout (close below consolidation low)
      else if (breakoutBar.close < consolidationLow) {
        breakoutDirection = 'BEARISH';
        breakoutPrice = breakoutBar.close;
      }
      else {
        // No breakout at this consolidation end point
        continue;
      }
      
      // Requirement 15.4: Validate volume increase (≥50% above average)
      // Calculate average volume during consolidation period
      let consolidationVolumeSum = 0;
      let consolidationVolumeCount = 0;
      
      for (let i = consolidationStart; i <= consolidationEnd; i++) {
        if (priceData[i] && priceData[i].volume !== undefined && isFinite(priceData[i].volume)) {
          consolidationVolumeSum += priceData[i].volume;
          consolidationVolumeCount++;
        }
      }
      
      // Skip if we don't have volume data
      if (consolidationVolumeCount === 0) {
        continue;
      }
      
      const avgConsolidationVolume = consolidationVolumeSum / consolidationVolumeCount;
      
      // Check breakout bar volume
      if (!breakoutBar.volume || !isFinite(breakoutBar.volume)) {
        continue;
      }
      
      // Verify volume is at least 50% above average
      const volumeIncreasePercent = ((breakoutBar.volume - avgConsolidationVolume) / avgConsolidationVolume) * 100;
      
      if (volumeIncreasePercent < 50) {
        continue;
      }
      
      // Calculate target price based on breakout direction
      // Traditional measure: consolidation range added/subtracted from breakout level
      let targetPrice;
      let neckline;
      
      if (breakoutDirection === 'BULLISH') {
        neckline = consolidationHigh;
        targetPrice = consolidationHigh + consolidationRange;
      } else {
        neckline = consolidationLow;
        targetPrice = consolidationLow - consolidationRange;
      }
      
      // Breakout is confirmed by definition (price has already broken out)
      const confirmed = true;
      
      // Requirement 15.5: Return the breakout with direction, price level, and volume confirmation
      return {
        type: 'BREAKOUT',
        startIndex: consolidationStart,
        endIndex: breakoutIndex,
        keyPoints: [
          { index: consolidationStart, price: priceData[consolidationStart].close, label: 'consolidationStart' },
          { index: consolidationEnd, price: priceData[consolidationEnd].close, label: 'consolidationEnd' },
          { index: breakoutIndex, price: breakoutPrice, label: 'breakout' }
        ],
        neckline: neckline,
        confirmed: confirmed,
        confidence: 0, // Will be calculated by calculateConfidence function
        direction: breakoutDirection,
        targetPrice: targetPrice,
        metadata: {
          consolidationHigh: consolidationHigh,
          consolidationLow: consolidationLow,
          consolidationRange: consolidationRange,
          rangePercent: rangePercent,
          consolidationLength: consolidationEnd - consolidationStart + 1,
          breakoutPrice: breakoutPrice,
          breakoutVolume: breakoutBar.volume,
          avgConsolidationVolume: avgConsolidationVolume,
          volumeIncreasePercent: volumeIncreasePercent
        }
      };
    }
  }
  
  // No valid breakout pattern found
  return null;
}

/**
 * Detects price gaps
 * @param {Array<Object>} priceData - OHLCV data (requires open prices)
 * @returns {Pattern|null} - Detected pattern or null
 */
function detectGap(priceData) {
  if (!priceData || priceData.length < 2) {
    return null;
  }
  
  // Scan through price data looking for gaps
  // A gap occurs when there's no price overlap between consecutive bars
  for (let i = 1; i < priceData.length; i++) {
    const prevBar = priceData[i - 1];
    const currBar = priceData[i];
    
    // Skip if either bar is missing required data or has invalid values
    if (!prevBar || !currBar || 
        !isFinite(prevBar.high) || !isFinite(prevBar.low) || !isFinite(prevBar.close) ||
        !isFinite(currBar.high) || !isFinite(currBar.low) || !isFinite(currBar.open)) {
      continue;
    }
    
    // Requirement 16.1: Identify gap up (current bar's low is above previous bar's high)
    const isGapUp = currBar.low > prevBar.high;
    
    // Requirement 16.2: Identify gap down (current bar's high is below previous bar's low)
    const isGapDown = currBar.high < prevBar.low;
    
    // If no gap exists, continue to next bar
    if (!isGapUp && !isGapDown) {
      continue;
    }
    
    // Determine gap direction and calculate gap size
    let gapDirection;
    let gapSize;
    let gapLowPrice;
    let gapHighPrice;
    
    if (isGapUp) {
      gapDirection = 'BULLISH';
      gapLowPrice = prevBar.high;
      gapHighPrice = currBar.low;
      gapSize = currBar.low - prevBar.high;
    } else {
      gapDirection = 'BEARISH';
      gapLowPrice = currBar.high;
      gapHighPrice = prevBar.low;
      gapSize = prevBar.low - currBar.high;
    }
    
    // Requirement 16.3: Calculate gap size as percentage of previous close
    const gapSizePercent = (gapSize / prevBar.close) * 100;
    
    // Requirement 16.4: Classify gap type based on context
    // We'll use a simplified classification based on gap size and position
    let gapType;
    
    // Analyze context to classify gap type
    // Look at the trend before the gap and the gap size
    
    // Calculate trend before gap (look back 5-10 bars)
    const lookbackBars = Math.min(10, i);
    let trendDirection = 'NEUTRAL';
    
    if (lookbackBars >= 5) {
      const startBar = priceData[i - lookbackBars];
      const startPrice = startBar ? startBar.close : null;
      const endPrice = prevBar.close;
      
      // Only calculate trend if we have valid data
      if (startPrice && isFinite(startPrice) && isFinite(endPrice) && startPrice !== 0) {
        const trendChange = ((endPrice - startPrice) / startPrice) * 100;
        
        if (trendChange > 5) {
          trendDirection = 'UPTREND';
        } else if (trendChange < -5) {
          trendDirection = 'DOWNTREND';
        }
      }
    }
    
    // Classify gap type:
    // - Common gap: Small gap (< 2%) in sideways market
    // - Breakaway gap: Medium to large gap (2-5%) at start of new trend
    // - Runaway gap: Medium gap (2-4%) in middle of strong trend
    // - Exhaustion gap: Large gap (> 4%) near end of trend
    
    if (gapSizePercent < 2) {
      gapType = 'COMMON';
    } else if (gapSizePercent >= 2 && gapSizePercent < 5) {
      // Determine if breakaway or runaway based on trend
      if (trendDirection === 'NEUTRAL') {
        gapType = 'BREAKAWAY';
      } else if ((gapDirection === 'BULLISH' && trendDirection === 'UPTREND') ||
                 (gapDirection === 'BEARISH' && trendDirection === 'DOWNTREND')) {
        gapType = 'RUNAWAY';
      } else {
        gapType = 'BREAKAWAY';
      }
    } else {
      // Large gap (>= 5%)
      // Could be breakaway or exhaustion
      // If gap is in same direction as existing strong trend, likely exhaustion
      // If gap is against trend or in neutral market, likely breakaway
      if ((gapDirection === 'BULLISH' && trendDirection === 'UPTREND') ||
          (gapDirection === 'BEARISH' && trendDirection === 'DOWNTREND')) {
        gapType = 'EXHAUSTION';
      } else {
        gapType = 'BREAKAWAY';
      }
    }
    
    // Check if gap has been filled (price has returned to gap range)
    let gapFilled = false;
    
    for (let j = i + 1; j < priceData.length; j++) {
      if (priceData[j]) {
        if (gapDirection === 'BULLISH') {
          // Gap is filled if price drops back into gap range
          if (isFinite(priceData[j].low) && priceData[j].low <= gapHighPrice) {
            gapFilled = true;
            break;
          }
        } else {
          // Gap is filled if price rises back into gap range
          if (isFinite(priceData[j].high) && priceData[j].high >= gapLowPrice) {
            gapFilled = true;
            break;
          }
        }
      }
    }
    
    // Gaps are typically considered "confirmed" if they remain unfilled
    const confirmed = !gapFilled;
    
    // Calculate target price based on gap type
    // For breakaway and runaway gaps, target is typically gap size added to current price
    // For exhaustion gaps, target is often back to pre-gap levels
    let targetPrice;
    
    if (gapType === 'EXHAUSTION') {
      // Exhaustion gap suggests reversal, target is back to gap fill
      targetPrice = gapDirection === 'BULLISH' ? gapLowPrice : gapHighPrice;
    } else {
      // Breakaway and runaway gaps suggest continuation
      if (gapDirection === 'BULLISH') {
        targetPrice = currBar.close + gapSize;
      } else {
        targetPrice = currBar.close - gapSize;
      }
    }
    
    // Requirement 16.5: Return gap with type, size, and price levels
    return {
      type: 'GAP',
      startIndex: i - 1,
      endIndex: i,
      keyPoints: [
        { index: i - 1, price: prevBar.close, label: 'preGap' },
        { index: i, price: currBar.open, label: 'postGap' }
      ],
      neckline: gapDirection === 'BULLISH' ? gapLowPrice : gapHighPrice,
      confirmed: confirmed,
      confidence: 0, // Will be calculated by calculateConfidence function
      direction: gapDirection,
      targetPrice: targetPrice,
      metadata: {
        gapType: gapType,
        gapSize: gapSize,
        gapSizePercent: gapSizePercent,
        gapLowPrice: gapLowPrice,
        gapHighPrice: gapHighPrice,
        gapFilled: gapFilled,
        trendDirection: trendDirection,
        prevClose: prevBar.close,
        currOpen: currBar.open
      }
    };
  }
  
  // No gap found
  return null;
}

// ============================================================================
// Confidence Scoring
// ============================================================================

/**
 * Calculates confidence score for a pattern
 * @param {Pattern} pattern - Pattern to score
 * @param {Array<Object>} priceData - Full price data for context
 * @returns {number} - Confidence score (0-100)
 */
function calculateConfidence(pattern, priceData) {
  if (!pattern || !priceData || priceData.length === 0) {
    return 0;
  }
  
  let totalScore = 0;
  let maxScore = 0;
  
  // Factor 1: Pattern Symmetry (0-25 points)
  // Applies to: Double Top/Bottom, Head and Shoulders, Inverse Head and Shoulders
  if (pattern.type === 'DOUBLE_TOP' || pattern.type === 'DOUBLE_BOTTOM') {
    maxScore += 25;
    
    // For double patterns, check peak/trough similarity
    if (pattern.metadata && pattern.metadata.peakDifference !== undefined) {
      const difference = pattern.metadata.peakDifference;
      // Perfect symmetry (0% difference) = 25 points
      // 3% difference = 0 points (linear scale)
      const symmetryScore = Math.max(0, 25 * (1 - difference / 3));
      totalScore += symmetryScore;
    } else if (pattern.metadata && pattern.metadata.troughDifference !== undefined) {
      const difference = pattern.metadata.troughDifference;
      const symmetryScore = Math.max(0, 25 * (1 - difference / 3));
      totalScore += symmetryScore;
    }
  } else if (pattern.type === 'HEAD_SHOULDERS' || pattern.type === 'INVERSE_HEAD_SHOULDERS') {
    maxScore += 25;
    
    // For head and shoulders, check shoulder symmetry
    if (pattern.metadata && pattern.metadata.shoulderDifference !== undefined) {
      const difference = pattern.metadata.shoulderDifference;
      // Perfect symmetry (0% difference) = 25 points
      // 5% difference = 0 points (linear scale)
      const symmetryScore = Math.max(0, 25 * (1 - difference / 5));
      totalScore += symmetryScore;
    }
  }
  
  // Factor 2: Volume Confirmation (0-25 points)
  // Applies to: All patterns with breakout volume data
  if (pattern.confirmed && pattern.metadata) {
    maxScore += 25;
    
    // Check for volume increase on breakout
    if (pattern.metadata.breakoutVolume !== undefined && pattern.metadata.avgPatternVolume !== undefined) {
      const avgVolume = pattern.metadata.avgPatternVolume;
      const breakoutVolume = pattern.metadata.breakoutVolume;
      
      if (avgVolume > 0) {
        const volumeIncreasePercent = ((breakoutVolume - avgVolume) / avgVolume) * 100;
        
        // 50% increase = 10 points (minimum for confirmation)
        // 150% increase or more = 25 points (excellent)
        // Linear scale between 50% and 150%
        const volumeScore = Math.min(25, Math.max(0, ((volumeIncreasePercent - 50) / 100) * 25));
        totalScore += volumeScore;
      }
    } else if (pattern.metadata.avgConsolidationVolume !== undefined && pattern.metadata.breakoutVolume !== undefined) {
      // For flag and pennant patterns
      const avgVolume = pattern.metadata.avgConsolidationVolume;
      const breakoutVolume = pattern.metadata.breakoutVolume;
      
      if (avgVolume > 0) {
        const volumeIncreasePercent = ((breakoutVolume - avgVolume) / avgVolume) * 100;
        const volumeScore = Math.min(25, Math.max(0, ((volumeIncreasePercent - 50) / 100) * 25));
        totalScore += volumeScore;
      }
    }
  } else if (!pattern.confirmed) {
    // Pattern not yet confirmed, no volume score
    maxScore += 25;
    // Add 0 points for unconfirmed patterns
  }
  
  // Factor 3: Trendline Quality (R² for triangles and wedges) (0-20 points)
  // Applies to: Triangles, Wedges, Pennants
  if (pattern.type === 'ASCENDING_TRIANGLE' || pattern.type === 'DESCENDING_TRIANGLE' || 
      pattern.type === 'SYMMETRICAL_TRIANGLE' || pattern.type === 'RISING_WEDGE' || 
      pattern.type === 'FALLING_WEDGE' || pattern.type === 'PENNANT') {
    maxScore += 20;
    
    // Calculate R² for trendlines
    // For simplicity, we'll use the number of pivot points that align with the trendlines
    // More aligned points = higher quality trendlines
    
    if (pattern.keyPoints && pattern.keyPoints.length >= 4) {
      // Good trendline quality: 4+ key points = 20 points
      // Minimum trendline quality: 2-3 key points = 10 points
      const numPoints = pattern.keyPoints.length;
      const trendlineScore = Math.min(20, Math.max(10, (numPoints - 2) * 5));
      totalScore += trendlineScore;
    } else {
      // Fewer than 4 points, lower quality
      totalScore += 10;
    }
  }
  
  // Factor 4: Pattern Duration (0-15 points)
  // Longer patterns are generally more reliable
  maxScore += 15;
  
  const patternLength = pattern.endIndex - pattern.startIndex;
  
  // 15 bars (minimum) = 5 points
  // 30 bars = 10 points
  // 50+ bars = 15 points (maximum)
  let durationScore;
  if (patternLength >= 50) {
    durationScore = 15;
  } else if (patternLength >= 30) {
    durationScore = 10 + ((patternLength - 30) / 20) * 5;
  } else if (patternLength >= 15) {
    durationScore = 5 + ((patternLength - 15) / 15) * 5;
  } else {
    durationScore = Math.max(0, (patternLength / 15) * 5);
  }
  
  totalScore += durationScore;
  
  // Factor 5: Price Action Clarity (0-15 points)
  // Measure how clean the pattern is (less noise = higher score)
  maxScore += 15;
  
  // Calculate price volatility during the pattern
  // Lower volatility (relative to pattern size) = cleaner pattern
  if (pattern.startIndex >= 0 && pattern.endIndex < priceData.length) {
    let priceSum = 0;
    let priceCount = 0;
    let priceVariance = 0;
    
    // Calculate average price during pattern
    for (let i = pattern.startIndex; i <= pattern.endIndex; i++) {
      if (priceData[i] && isFinite(priceData[i].close)) {
        priceSum += priceData[i].close;
        priceCount++;
      }
    }
    
    if (priceCount > 0) {
      const avgPrice = priceSum / priceCount;
      
      // Calculate variance
      for (let i = pattern.startIndex; i <= pattern.endIndex; i++) {
        if (priceData[i] && isFinite(priceData[i].close)) {
          const diff = priceData[i].close - avgPrice;
          priceVariance += diff * diff;
        }
      }
      
      const variance = priceVariance / priceCount;
      const stdDev = Math.sqrt(variance);
      
      // Calculate coefficient of variation (CV = stdDev / mean)
      // Lower CV = less noise = higher score
      const cv = avgPrice > 0 ? stdDev / avgPrice : 1;
      
      // CV of 0.02 (2%) or less = 15 points (very clean)
      // CV of 0.10 (10%) or more = 0 points (very noisy)
      // Linear scale between 0.02 and 0.10
      const clarityScore = Math.max(0, Math.min(15, 15 * (1 - (cv - 0.02) / 0.08)));
      totalScore += clarityScore;
    }
  }
  
  // Additional bonus points for specific pattern characteristics
  
  // Bonus 1: Confirmed patterns get extra points (0-10 points)
  if (pattern.confirmed) {
    totalScore += 10;
  }
  
  // Bonus 2: Patterns with good depth/height get extra points (0-10 points)
  if (pattern.metadata) {
    if (pattern.metadata.valleyDepth !== undefined) {
      // For double top: deeper valley = better pattern
      const depth = pattern.metadata.valleyDepth;
      // 10% depth = 5 points, 20%+ depth = 10 points
      const depthScore = Math.min(10, Math.max(0, ((depth - 10) / 10) * 10));
      totalScore += depthScore;
    } else if (pattern.metadata.peakHeight !== undefined) {
      // For double bottom: higher peak = better pattern
      const height = pattern.metadata.peakHeight;
      const heightScore = Math.min(10, Math.max(0, ((height - 10) / 10) * 10));
      totalScore += heightScore;
    } else if (pattern.metadata.patternHeight !== undefined && pattern.metadata.heightPercent !== undefined) {
      // For head and shoulders: good height = better pattern
      const heightPercent = pattern.metadata.heightPercent;
      const heightScore = Math.min(10, Math.max(0, ((heightPercent - 10) / 10) * 10));
      totalScore += heightScore;
    } else if (pattern.metadata.depthPercent !== undefined) {
      // For cup and handle, rounding bottom: good depth = better pattern
      const depth = pattern.metadata.depthPercent;
      const depthScore = Math.min(10, Math.max(0, ((depth - 10) / 10) * 10));
      totalScore += depthScore;
    }
  }
  
  // Calculate final confidence score as percentage
  // Base score from factors (out of maxScore)
  // Plus bonus points (up to 20)
  const maxPossibleScore = maxScore + 20;
  
  // Ensure score is between 0 and 100
  const confidenceScore = Math.min(100, Math.max(0, (totalScore / maxPossibleScore) * 100));
  
  // Round to nearest integer
  return Math.round(confidenceScore);
}

// ============================================================================
// Pattern Filtering and Prioritization
// ============================================================================

/**
 * Filters and prioritizes overlapping patterns
 * @param {Array<Pattern>} patterns - All detected patterns
 * @returns {Array<Pattern>} - Filtered patterns
 */
function prioritizePatterns(patterns) {
  // Handle empty or invalid input
  if (!patterns || !Array.isArray(patterns) || patterns.length === 0) {
    return [];
  }
  
  // If only one pattern, no prioritization needed
  if (patterns.length === 1) {
    return patterns;
  }
  
  // Helper function to check if two patterns overlap
  // Two patterns overlap if their bar ranges intersect
  function patternsOverlap(pattern1, pattern2) {
    // Get the bar ranges for each pattern
    const start1 = pattern1.startIndex;
    const end1 = pattern1.endIndex;
    const start2 = pattern2.startIndex;
    const end2 = pattern2.endIndex;
    
    // Check if ranges overlap
    // Patterns overlap if: start1 <= end2 AND start2 <= end1
    return start1 <= end2 && start2 <= end1;
  }
  
  // Helper function to compare patterns for prioritization
  // Returns negative if pattern1 has higher priority, positive if pattern2 has higher priority
  function comparePatternPriority(pattern1, pattern2) {
    // Priority 1: Recency (most recent first)
    // More recent patterns have higher endIndex values
    if (pattern1.endIndex !== pattern2.endIndex) {
      return pattern2.endIndex - pattern1.endIndex; // Higher endIndex = higher priority
    }
    
    // Priority 2: Confidence (highest first)
    const confidence1 = pattern1.confidence || 0;
    const confidence2 = pattern2.confidence || 0;
    if (confidence1 !== confidence2) {
      return confidence2 - confidence1; // Higher confidence = higher priority
    }
    
    // Priority 3: Confirmation status (confirmed first)
    const confirmed1 = pattern1.confirmed ? 1 : 0;
    const confirmed2 = pattern2.confirmed ? 1 : 0;
    if (confirmed1 !== confirmed2) {
      return confirmed2 - confirmed1; // Confirmed = higher priority
    }
    
    // If all else is equal, maintain original order
    return 0;
  }
  
  // Sort patterns by priority (highest priority first)
  const sortedPatterns = [...patterns].sort(comparePatternPriority);
  
  // Filter out overlapping patterns, keeping only the highest priority ones
  const filteredPatterns = [];
  
  for (const pattern of sortedPatterns) {
    // Check if this pattern overlaps with any already-selected pattern
    let hasOverlap = false;
    
    for (const selectedPattern of filteredPatterns) {
      if (patternsOverlap(pattern, selectedPattern)) {
        hasOverlap = true;
        break;
      }
    }
    
    // If no overlap, add this pattern to the filtered list
    if (!hasOverlap) {
      filteredPatterns.push(pattern);
    }
  }
  
  return filteredPatterns;
}

// ============================================================================
// Google Sheets Integration
// ============================================================================

/**
 * Retrieves price data for a specific ticker from DATA sheet
 * @param {Sheet} dataSheet - DATA sheet reference
 * @param {string} ticker - Ticker symbol
 * @param {number} tickerIndex - Index of ticker in list
 * @param {number} blockSize - Size of data blocks (default 7 columns per ticker)
 * @returns {Array<Object>} - Array of price data objects {date, open, high, low, close, volume}
 */
function getPriceDataForTicker(dataSheet, ticker, tickerIndex, blockSize) {
  // Validate inputs
  if (!dataSheet) {
    console.error('getPriceDataForTicker: dataSheet is null or undefined');
    return [];
  }
  
  if (tickerIndex < 0) {
    console.error(`getPriceDataForTicker: invalid tickerIndex ${tickerIndex}`);
    return [];
  }
  
  // Default blockSize to 7 if not provided
  const BLOCK = blockSize || 7;
  
  try {
    // Calculate column offset for this ticker
    // Each ticker occupies BLOCK columns in the DATA sheet
    // Ticker 0 starts at column 1, ticker 1 starts at column 8, etc.
    const columnOffset = (tickerIndex * BLOCK) + 1;
    
    // Column indices (relative to columnOffset):
    // 0: Date
    // 1: Open
    // 2: High
    // 3: Low
    // 4: Close
    // 5: Volume
    // 6: (unused in GOOGLEFINANCE output)
    
    // Historical data starts at row 5 (rows 1-4 are headers and metadata)
    const startRow = 5;
    
    // Get the last row with data in the date column for this ticker
    const dateColumn = columnOffset;
    const lastRow = dataSheet.getLastRow();
    
    // If no data rows exist, return empty array
    if (lastRow < startRow) {
      console.log(`getPriceDataForTicker: No data rows for ticker ${ticker} (lastRow: ${lastRow})`);
      return [];
    }
    
    // Calculate number of rows to read
    const numRows = lastRow - startRow + 1;
    
    // Read all 6 columns (date, open, high, low, close, volume) for this ticker
    // We read 6 columns starting from columnOffset
    const dataRange = dataSheet.getRange(startRow, columnOffset, numRows, 6);
    const rawData = dataRange.getValues();
    
    // Parse data into array of price objects
    const priceData = [];
    
    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      
      // Extract values from row
      const date = row[0];
      const open = row[1];
      const high = row[2];
      const low = row[3];
      const close = row[4];
      const volume = row[5];
      
      // Skip rows with missing or invalid data
      // Date must be present and valid
      if (!date || date === '' || !(date instanceof Date)) {
        continue;
      }
      
      // OHLC values must be numbers and greater than 0
      if (typeof open !== 'number' || open <= 0 ||
          typeof high !== 'number' || high <= 0 ||
          typeof low !== 'number' || low <= 0 ||
          typeof close !== 'number' || close <= 0) {
        continue;
      }
      
      // Volume must be a number (can be 0 for some data sources)
      if (typeof volume !== 'number' || volume < 0) {
        continue;
      }
      
      // Validate OHLC relationships (high >= low, etc.)
      if (high < low || close > high || close < low || open > high || open < low) {
        console.warn(`getPriceDataForTicker: Invalid OHLC data at row ${startRow + i} for ${ticker}`);
        continue;
      }
      
      // Add valid price data object
      priceData.push({
        date: date,
        open: open,
        high: high,
        low: low,
        close: close,
        volume: volume
      });
    }
    
    console.log(`getPriceDataForTicker: Retrieved ${priceData.length} valid bars for ${ticker}`);
    
    return priceData;
    
  } catch (error) {
    console.error(`getPriceDataForTicker: Error reading data for ticker ${ticker}: ${error.message}`);
    console.error(error.stack);
    return [];
  }
}

// ============================================================================
// Pattern Short Form Mapping
// ============================================================================

/**
 * Mapping of full pattern names to short forms for compact display
 * Reduces cell size while maintaining readability
 */
const PATTERN_SHORT_FORMS = {
  'DOUBLE_TOP': 'DBL_TOP',
  'DOUBLE_BOTTOM': 'DBL_BTM',
  'HEAD_SHOULDERS': 'H&S',
  'INVERSE_HEAD_SHOULDERS': 'INV_H&S',
  'CUP_HANDLE': 'CUP_HDL',
  'ROUNDING_BOTTOM': 'RND_BTM',
  'ASCENDING_TRIANGLE': 'ASC_TRI',
  'DESCENDING_TRIANGLE': 'DESC_TRI',
  'SYMMETRICAL_TRIANGLE': 'SYM_TRI',
  'FLAG': 'FLAG',
  'PENNANT': 'PENNANT',
  'RISING_WEDGE': 'RISE_WDG',
  'FALLING_WEDGE': 'FALL_WDG',
  'RECTANGLE': 'RECT',
  'BREAKOUT': 'BRKOUT',
  'GAP': 'GAP'
};

/**
 * Formats detected patterns as pipe-separated string for sheet display
 * Uses short forms to reduce cell size
 * @param {Array<Pattern>} patterns - Detected patterns
 * @returns {string} - Formatted pattern string (e.g., "DBL_TOP (78%) | BRKOUT (85%)")
 */
function formatPatternsForSheet(patterns) {
  if (!patterns || patterns.length === 0) {
    return '';
  }
  
  // Filter patterns with confidence >= 60% (as per design document)
  const validPatterns = patterns.filter(p => p && p.confidence >= 60);
  
  if (validPatterns.length === 0) {
    return '';
  }
  
  // Format each pattern as "SHORT_FORM (CONFIDENCE%)"
  const formattedPatterns = validPatterns.map(pattern => {
    // Ensure pattern type is in uppercase with underscores
    const patternName = pattern.type.toUpperCase().replace(/\s+/g, '_');
    
    // Use short form if available, otherwise use full name
    const displayName = PATTERN_SHORT_FORMS[patternName] || patternName;
    
    const confidence = Math.round(pattern.confidence);
    return `${displayName} (${confidence}%)`;
  });
  
  // Join with pipe separator
  return formattedPatterns.join(' | ');
}

// ============================================================================
// Exports for Testing
// ============================================================================

// Export functions for testing (Node.js environment)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    findPivotPoints,
    detectPatterns,
    detectPatternsForAllTickers,
    validatePattern,
    hasMinimumSpacing,
    hasMinimumDepth,
    detectDoubleTop,
    detectDoubleBottom,
    detectHeadAndShoulders,
    detectInverseHeadAndShoulders,
    detectCupAndHandle,
    detectRoundingBottom,
    detectAscendingTriangle,
    detectDescendingTriangle,
    detectSymmetricalTriangle,
    detectFlag,
    detectPennant,
    detectRisingWedge,
    detectFallingWedge,
    detectRectangle,
    detectBreakout,
    detectGap,
    calculateConfidence,
    prioritizePatterns,
    getPriceDataForTicker,
    formatPatternsForSheet
  };
}

// ============================================================================
// Pattern Caching Functions
// ============================================================================

/**
 * Caching layer for pattern detection to enable formula-based pattern display
 * that updates when prices change, without recalculating patterns on every cell change.
 * 
 * Pattern detection is expensive (1-2 seconds per ticker), so we cache results
 * and only recalculate when the DATA sheet is updated or cache expires.
 */

// Global cache configuration
const CACHE_KEY_PREFIX = 'PATTERN_CACHE_';
const CACHE_EXPIRY_MINUTES = 60; // Cache expires after 1 hour

/**
 * Custom function to get cached patterns for a ticker
 * Can be called from a formula: =GETPATTERNS(A3, E3)
 * 
 * @param {string} ticker - Ticker symbol
 * @param {number} currentPrice - Current price (used to trigger recalculation)
 * @returns {string} - Formatted pattern string or empty string
 * @customfunction
 */
function GETPATTERNS(ticker, currentPrice) {
  // Validate inputs
  if (!ticker || ticker === '') {
    console.log('GETPATTERNS: Empty ticker');
    return '';
  }
  
  console.log(`GETPATTERNS called: ticker="${ticker}", price=${currentPrice}`);
  
  // Try to get from cache first
  const cachedPattern = getCachedPattern(ticker);
  console.log(`GETPATTERNS: Cache lookup result for ${ticker}: ${cachedPattern === null ? 'NULL' : `"${cachedPattern}"`}`);
  
  if (cachedPattern !== null) {
    console.log(`GETPATTERNS: Returning cached pattern for ${ticker}: "${cachedPattern}"`);
    return cachedPattern;
  }
  
  // CRITICAL FIX: If cache is missing, recalculate patterns instead of returning empty
  console.log(`GETPATTERNS: Cache miss for ${ticker}, recalculating patterns...`);
  
  try {
    // Get spreadsheet and DATA sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName('DATA');
    const inputSheet = ss.getSheetByName('INPUT');
    
    if (!dataSheet || !inputSheet) {
      console.log(`GETPATTERNS: Required sheets not found`);
      return '';
    }
    
    // Find ticker index in INPUT sheet
    const tickers = inputSheet.getRange('A3:A').getValues().flat().filter(t => t !== '');
    const tickerIndex = tickers.findIndex(t => String(t).toUpperCase() === String(ticker).toUpperCase());
    
    if (tickerIndex === -1) {
      console.log(`GETPATTERNS: Ticker ${ticker} not found in INPUT sheet`);
      return '';
    }
    
    // Get price data for this ticker
    const BLOCK = 7; // DATA block width
    const priceData = getPriceDataForTicker(dataSheet, ticker, tickerIndex, BLOCK);
    
    if (!priceData || priceData.length === 0) {
      console.log(`GETPATTERNS: No price data for ${ticker}`);
      setCachedPattern(ticker, '');
      return '';
    }
    
    // Detect patterns
    const patterns = detectPatterns(priceData, {minBars: 100, minConfidence: 60});
    const patternString = formatPatternsForSheet(patterns);
    
    // Cache the result for future calls
    setCachedPattern(ticker, patternString);
    
    console.log(`GETPATTERNS: Recalculated and cached patterns for ${ticker}: "${patternString}"`);
    return patternString;
    
  } catch (error) {
    console.error(`GETPATTERNS: Error recalculating patterns for ${ticker}: ${error.message}`);
    // Cache empty string to avoid repeated errors
    setCachedPattern(ticker, '');
    return '';
  }
}

/**
 * Gets cached pattern for a ticker
 * @param {string} ticker - Ticker symbol
 * @returns {string|null} - Cached pattern string or null if not found/expired
 */
function getCachedPattern(ticker) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = CACHE_KEY_PREFIX + ticker.toUpperCase();
    console.log(`getCachedPattern: Looking up key="${cacheKey}" for ticker="${ticker}"`);
    const cachedValue = cache.get(cacheKey);
    
    if (cachedValue) {
      console.log(`getCachedPattern: Found value="${cachedValue}" for key="${cacheKey}"`);
      return cachedValue;
    }
    
    console.log(`getCachedPattern: No value found for key="${cacheKey}"`);
    return null;
  } catch (error) {
    console.error(`Error getting cached pattern for ${ticker}: ${error.message}`);
    return null;
  }
}

/**
 * Sets cached pattern for a ticker
 * @param {string} ticker - Ticker symbol
 * @param {string} patternString - Formatted pattern string
 */
function setCachedPattern(ticker, patternString) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = CACHE_KEY_PREFIX + ticker.toUpperCase();
    const expirySeconds = CACHE_EXPIRY_MINUTES * 60;
    
    cache.put(cacheKey, patternString, expirySeconds);
  } catch (error) {
    console.error(`Error setting cached pattern for ${ticker}: ${error.message}`);
  }
}

/**
 * Clears all cached patterns
 * Call this when you want to force recalculation of all patterns
 */
function clearPatternCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll([CACHE_KEY_PREFIX]);
    console.log('Pattern cache cleared');
  } catch (error) {
    console.error(`Error clearing pattern cache: ${error.message}`);
  }
}

/**
 * Refreshes pattern cache for all tickers
 * This should be called after DATA sheet is updated or on a schedule
 * 
 * @param {Array<string>} tickers - Array of ticker symbols (optional, gets from INPUT if not provided)
 */
function refreshPatternCache(tickers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('DATA');
  
  // Get tickers from INPUT sheet if not provided
  if (!tickers || tickers.length === 0) {
    const inputSheet = ss.getSheetByName('INPUT');
    tickers = getCleanTickersForCache(inputSheet);
  }
  
  if (!tickers || tickers.length === 0) {
    console.log('No tickers found to refresh cache');
    return;
  }
  
  console.log(`Refreshing pattern cache for ${tickers.length} tickers...`);
  
  const BLOCK = 7;
  const BATCH_SIZE = 10;
  
  // Process tickers in batches
  for (let batchStart = 0; batchStart < tickers.length; batchStart += BATCH_SIZE) {
    const batchEnd = Math.min(batchStart + BATCH_SIZE, tickers.length);
    
    for (let i = batchStart; i < batchEnd; i++) {
      const ticker = tickers[i];
      
      try {
        // Get price data for this ticker
        const priceData = getPriceDataForTicker(dataSheet, ticker, i, BLOCK);
        
        // Detect patterns with minimum 100 bars and 60% confidence
        const patterns = detectPatterns(priceData, {minBars: 100, minConfidence: 60});
        
        // Format patterns for sheet display
        const patternString = formatPatternsForSheet(patterns);
        
        // Cache the result
        setCachedPattern(ticker, patternString);
        
        console.log(`${ticker}: Cached patterns - ${patternString || 'none'}`);
        
      } catch (error) {
        console.error(`Error refreshing cache for ${ticker}: ${error.message}`);
        // Cache empty string on error
        setCachedPattern(ticker, '');
      }
    }
    
    // Flush to avoid timeout
    SpreadsheetApp.flush();
  }
  
  console.log('Pattern cache refresh complete');
}

/**
 * Helper function to get clean tickers from INPUT sheet
 * @param {Sheet} inputSheet - INPUT sheet reference
 * @returns {Array<string>} - Array of ticker symbols
 */
function getCleanTickersForCache(inputSheet) {
  if (!inputSheet) {
    return [];
  }
  
  const tickerRange = inputSheet.getRange('A3:A').getValues();
  const tickers = [];
  
  for (let i = 0; i < tickerRange.length; i++) {
    const ticker = String(tickerRange[i][0] || '').trim().toUpperCase();
    if (ticker && ticker !== '') {
      tickers.push(ticker);
    } else {
      break; // Stop at first empty cell
    }
  }
  
  return tickers;
}

/**
 * Creates a time-based trigger to refresh pattern cache periodically
 * Call this once to set up automatic cache refresh
 * 
 * @param {number} intervalHours - How often to refresh (default: 1 hour)
 */
function setupPatternCacheRefreshTrigger(intervalHours) {
  intervalHours = intervalHours || 1;
  
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'refreshPatternCache') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create new trigger
  ScriptApp.newTrigger('refreshPatternCache')
    .timeBased()
    .everyHours(intervalHours)
    .create();
  
  console.log(`Pattern cache refresh trigger created (every ${intervalHours} hour(s))`);
}

/**
 * Removes the pattern cache refresh trigger
 */
function removePatternCacheRefreshTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'refreshPatternCache') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  
  console.log(`Removed ${removed} pattern cache refresh trigger(s)`);
}
