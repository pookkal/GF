# Trading System Requirements Specification

## Project Overview

This specification documents the requirements for a comprehensive Google Sheets-based trading system that provides automated signal generation, fundamental analysis, and mobile reporting capabilities. The system integrates real-time market data with technical indicators to support institutional-grade investment decisions.

## System Architecture

### Core Components

1. **INPUT Sheet** - User configuration and ticker management
2. **DATA Sheet** - Raw market data from Google Finance API
3. **CALCULATIONS Sheet** - Signal generation and technical analysis
4. **DASHBOARD Sheet** - Filtered institutional view
5. **REPORT Sheet** - Mobile-optimized detailed analysis
6. **CHART Sheet** - Interactive charting with live data

### Data Flow

```
INPUT (Tickers) → DATA (Google Finance) → CALCULATIONS (Signals) → DASHBOARD/REPORT (Views)
```

## Current Implementation Status

### ✅ Completed Features

#### 1. Data Integration (DATA Sheet)
- **Status**: ✅ Complete
- **Description**: Automated Google Finance API integration
- **Features**:
  - Real-time price data for multiple tickers
  - Historical OHLCV data (800+ days)
  - Fundamental metrics (P/E, EPS, ATH)
  - Market regime indicators (USA/India VIX, SMA200 ratios)
  - Automatic data refresh and formatting

#### 2. Signal Generation (CALCULATIONS Sheet)
- **Status**: ✅ Complete
- **Description**: Advanced technical analysis engine
- **Features**:
  - Multi-timeframe signal generation (Long-term vs Trend-based)
  - 34-column comprehensive analysis
  - Enhanced pattern recognition (ATH breakouts, volatility breakouts)
  - Risk management integration
  - Position sizing algorithms
  - Fundamental valuation scoring

#### 3. Dashboard Interface (DASHBOARD Sheet)
- **Status**: ✅ Complete
- **Description**: Bloomberg-style institutional interface
- **Features**:
  - Real-time filtering by sector/tags
  - Professional color coding and formatting
  - Sortable multi-column display
  - Automatic refresh controls

#### 4. Mobile Report (REPORT Sheet)
- **Status**: ✅ Complete with Recent Updates
- **Description**: Mobile-optimized detailed analysis
- **Recent Enhancements**:
  - ✅ Bull/bear volume color coding in charts
  - ✅ Live data integration using Google Finance API
  - ✅ Chart formatting improvements (white labels, proper scaling)
  - ✅ Ticker dropdown from DASHBOARD sheet
  - ✅ Enhanced pattern recognition display

#### 5. Interactive Charting
- **Status**: ✅ Complete
- **Description**: Dynamic price and volume charts
- **Features**:
  - Real-time price data with SMA overlays
  - Bull/bear volume bars with proper scaling
  - Live data stitching for current day
  - Interactive controls and date selection

## Current Issue: Column Swap Implementation

### Problem Statement
User requested to swap FUNDAMENTAL and DECISION columns in the system display. The swap has been implemented in the CALCULATIONS sheet formulas, but verification is needed to ensure all references are working correctly.

### Implementation Details

#### ✅ CALCULATIONS Sheet Formula Assignment
```javascript
formulas.push([
  fSignal,      // B - SIGNAL
  fFund,        // C - FUNDAMENTAL (swapped from D)
  fDecision,    // D - DECISION (swapped from C)
  // ... other columns
]);
```

#### ✅ DASHBOARD Sheet Headers
```javascript
const headers = [[
  "Ticker", "SIGNAL", "FUNDAMENTAL", "DECISION", // Correct order
  // ... other headers
]];
```

#### ✅ DASHBOARD Sheet Data Formula
The filter formula correctly pulls:
- `CALCULATIONS!$C$3:$C` (FUNDAMENTAL)
- `CALCULATIONS!$D$3:$D` (DECISION)

#### ✅ Mobile Report References
```javascript
REPORT.getRange('B5').setFormula(lookup('C')); // FUNDAMENTAL from column C
REPORT.getRange('B6').setFormula(lookup('D')); // DECISION from column D
```

#### ✅ Helper.js Documentation
- References updated to reflect new column structure

#### ✅ Monitor.js Alert Messages
- Alert messages updated to reference correct columns

### Verification Requirements

To ensure the column swap is working correctly, the following should be verified:

1. **CALCULATIONS Sheet Data**
   - Column C should contain fundamental analysis (VALUE, FAIR, EXPENSIVE, etc.)
   - Column D should contain investment decisions (BUY, HOLD, SELL, etc.)

2. **DASHBOARD Sheet Display**
   - Column C should show fundamental ratings
   - Column D should show investment decisions
   - Headers should read "FUNDAMENTAL" and "DECISION" in correct positions

3. **Mobile Report Display**
   - FUNDAMENTAL row should show valuation analysis
   - DECISION row should show investment recommendation
   - Data should match CALCULATIONS sheet

4. **Cross-Reference Consistency**
   - All lookup functions should reference correct columns
   - No hardcoded references to old column positions
   - Alert systems should reference correct data

## User Stories

### US-1: Column Swap Verification
**As a** system user  
**I want** FUNDAMENTAL and DECISION columns to be swapped throughout the system  
**So that** the display order matches my preferred workflow  

**Acceptance Criteria:**
- [ ] CALCULATIONS sheet shows FUNDAMENTAL in column C, DECISION in column D
- [ ] DASHBOARD sheet displays swapped columns correctly
- [ ] Mobile report shows correct data in swapped positions
- [ ] All formulas reference correct columns after swap
- [ ] No data inconsistencies between sheets

### US-2: Data Consistency Validation
**As a** system user  
**I want** all sheets to show consistent data after the column swap  
**So that** I can trust the system's recommendations  

**Acceptance Criteria:**
- [ ] Same ticker shows identical FUNDAMENTAL rating across all sheets
- [ ] Same ticker shows identical DECISION recommendation across all sheets
- [ ] Lookup functions return correct values from swapped columns
- [ ] No caching issues affecting data display

## Technical Requirements

### TR-1: Formula Integrity
- All lookup formulas must reference correct column letters (C for FUNDAMENTAL, D for DECISION)
- No hardcoded cell references to old column positions
- Conditional formatting rules must apply to correct columns

### TR-2: Performance Requirements
- Sheet refresh time must remain under 10 seconds
- Formula calculations must complete without errors
- No circular reference issues from column changes

### TR-3: Data Validation
- FUNDAMENTAL values must be from valid set: VALUE, FAIR, EXPENSIVE, PRICED FOR PERFECTION, ZOMBIE
- DECISION values must be from valid set: BUY, SELL, HOLD, AVOID, etc.
- No blank or error values in swapped columns

## Testing Scenarios

### Test Case 1: Basic Column Swap Verification
1. Select a ticker in CALCULATIONS sheet
2. Verify column C shows fundamental rating (e.g., "VALUE", "EXPENSIVE")
3. Verify column D shows investment decision (e.g., "BUY", "HOLD")
4. Check same ticker in DASHBOARD sheet shows identical values in same columns
5. Check mobile report displays correct values for FUNDAMENTAL and DECISION rows

### Test Case 2: Cross-Sheet Consistency
1. Change ticker selection in mobile report
2. Verify FUNDAMENTAL and DECISION values update correctly
3. Compare values with CALCULATIONS and DASHBOARD sheets
4. Ensure no data mismatches or delays

### Test Case 3: Formula Reference Validation
1. Manually check lookup formulas in mobile report
2. Verify `lookup('C')` returns FUNDAMENTAL data
3. Verify `lookup('D')` returns DECISION data
4. Test with multiple tickers to ensure consistency

## Success Criteria

The column swap implementation will be considered successful when:

1. **Functional Requirements Met**
   - All sheets display FUNDAMENTAL in column C, DECISION in column D
   - Data consistency maintained across all views
   - No formula errors or broken references

2. **User Experience Preserved**
   - System performance remains unchanged
   - All existing functionality continues to work
   - Mobile report displays correct information

3. **Data Integrity Maintained**
   - No data loss during column swap
   - All historical functionality preserved
   - Lookup functions return correct values

## Implementation Notes

### Files Modified for Column Swap
- `Code.js` - Formula array assignment in generateCalculationsSheet()
- `mobilereport-formulas.js` - Lookup references for FUNDAMENTAL/DECISION
- `Helper.js` - Documentation references
- `Monitor.js` - Alert message references

### Key Technical Considerations
- Google Sheets formula references are case-sensitive
- Column letters must be exact (C vs D)
- Merged cells in mobile report require careful handling
- Conditional formatting rules may need column updates

## Future Enhancements

### Potential Improvements
1. **Enhanced Validation**
   - Add data validation rules to prevent invalid entries
   - Implement cross-sheet consistency checks
   - Add automated testing for column references

2. **User Interface**
   - Add visual indicators for column swap status
   - Implement user preferences for column ordering
   - Add tooltips explaining FUNDAMENTAL vs DECISION

3. **Performance Optimization**
   - Cache frequently accessed lookup values
   - Optimize formula calculations for large datasets
   - Implement incremental refresh for changed data only

## Conclusion

The column swap implementation appears to be technically complete based on code analysis. The primary requirement is verification that the changes are working correctly in the live system and that all data displays consistently across sheets. This specification provides the framework for validating the implementation and ensuring user requirements are met.