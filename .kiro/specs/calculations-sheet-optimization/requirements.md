# Requirements Document

## Introduction

This specification addresses the performance optimization of the Calculations Sheet in the Institutional Terminal Google Sheets application. The current implementation experiences slow loading times that impact user experience and operational efficiency.

## Glossary

- **Calculations_Sheet**: The primary analytical sheet that computes technical indicators, signals, and decisions for multiple tickers
- **DATA_Sheet**: The source sheet containing raw market data from GOOGLEFINANCE functions
- **Terminal**: The complete Google Sheets-based institutional trading application
- **GOOGLEFINANCE**: Google Sheets built-in function for fetching real-time market data
- **Indicator**: Technical analysis calculation (RSI, MACD, ADX, ATR, Stochastic, etc.)
- **Ticker**: Stock symbol identifier (e.g., AAPL, GOOGL)
- **Formula_Overhead**: Computational cost of complex nested formulas in spreadsheet cells
- **Batch_Operation**: Processing multiple items together rather than individually

## Requirements

### Requirement 1: Reduce Calculations Sheet Load Time

**User Story:** As a trader, I want the Calculations sheet to load quickly, so that I can make timely trading decisions without waiting for data to populate.

#### Acceptance Criteria

1. WHEN the generateCalculationsSheet function executes, THE System SHALL complete within 15 seconds for up to 50 tickers
2. WHEN formulas are written to the sheet, THE System SHALL use batch operations instead of individual cell writes
3. WHEN complex calculations are performed, THE System SHALL minimize nested formula depth to reduce computation overhead
4. WHEN the sheet refreshes, THE System SHALL preserve existing data during updates to avoid full recalculation

### Requirement 2: Optimize Formula Complexity

**User Story:** As a system administrator, I want formulas to be computationally efficient, so that the spreadsheet remains responsive during market hours.

#### Acceptance Criteria

1. WHEN formulas reference DATA sheet, THE System SHALL use direct cell references instead of MATCH/INDEX lookups where possible
2. WHEN formulas perform calculations, THE System SHALL pre-compute static values in Apps Script rather than in formulas
3. WHEN formulas use OFFSET functions, THE System SHALL limit the range size to minimum required rows
4. WHEN formulas use array functions, THE System SHALL specify explicit ranges instead of entire columns
5. WHEN formulas are duplicated across rows, THE System SHALL use array formulas where appropriate to reduce formula count

### Requirement 3: Minimize GOOGLEFINANCE API Calls

**User Story:** As a developer, I want to reduce redundant API calls, so that the application stays within Google's quota limits and performs faster.

#### Acceptance Criteria

1. WHEN market data is needed, THE System SHALL fetch data once in DATA sheet and reference it in CALCULATIONS
2. WHEN PE and EPS values are needed, THE System SHALL read cached values from DATA sheet row 3 instead of calling GOOGLEFINANCE again
3. WHEN ATH values are needed, THE System SHALL read cached values from DATA sheet row 3 instead of recalculating
4. WHEN multiple indicators need the same price data, THE System SHALL reference a single source cell
5. WHEN the sheet updates, THE System SHALL avoid triggering unnecessary GOOGLEFINANCE recalculations
6. If any functions call from file IndicatorFuncs.js does not need live data , move them to DATA

### Requirement 4: Implement Efficient Data Structures

**User Story:** As a developer, I want the code to use efficient data structures, so that processing time is minimized.

#### Acceptance Criteria

1. WHEN writing formulas to multiple cells, THE System SHALL use setFormulas() with 2D arrays instead of individual setFormula() calls
2. WHEN reading data from sheets, THE System SHALL use getValues() to read ranges in bulk instead of individual getValue() calls
3. WHEN applying formatting, THE System SHALL batch format operations by range instead of cell-by-cell
4. WHEN building formula strings, THE System SHALL use string concatenation efficiently or template literals
5. WHEN processing tickers, THE System SHALL minimize loops and use array operations where possible

### Requirement 5: Reduce Formula Recalculation Triggers

**User Story:** As a trader, I want the sheet to avoid unnecessary recalculations, so that I can work with the data without performance degradation.

#### Acceptance Criteria

1. WHEN formulas reference volatile functions, THE System SHALL minimize their use or cache results
2. WHEN formulas use TODAY() or NOW(), THE System SHALL limit usage to essential cells only
3. WHEN formulas use OFFSET with dynamic ranges, THE System SHALL use fixed ranges where the data size is predictable
4. WHEN circular dependencies could occur, THE System SHALL structure formulas to avoid them
5. WHEN the sheet is idle, THE System SHALL not trigger automatic recalculations unnecessarily

### Requirement 6: Optimize Indicator Calculations

**User Story:** As a trader, I want technical indicators to calculate efficiently, so that I can analyze multiple stocks simultaneously.

#### Acceptance Criteria

1. WHEN RSI is calculated, THE System SHALL use the LIVERSI custom function with optimized array processing
2. WHEN MACD is calculated, THE System SHALL use the LIVEMACD custom function with efficient EMA computation
3. WHEN ADX is calculated, THE System SHALL use the LIVEADX custom function with bounded data windows
4. WHEN ATR is calculated, THE System SHALL use the LIVEATR custom function with limited historical data
5. WHEN Stochastic is calculated, THE System SHALL use the LIVESTOCHK custom function with efficient min/max operations

### Requirement 7: Implement Progressive Loading

**User Story:** As a trader, I want to see partial results quickly, so that I can start analyzing data while the rest loads.

#### Acceptance Criteria

1. WHEN the sheet generates, THE System SHALL write ticker symbols and basic price data first
2. WHEN the sheet generates, THE System SHALL write simple indicators (SMA, Price, Change%) before complex ones
3. WHEN the sheet generates, THE System SHALL write complex indicators (RSI, MACD, ADX) in a second pass
4. WHEN the sheet generates, THE System SHALL write narrative text formulas last
5. WHEN the sheet generates, THE System SHALL call SpreadsheetApp.flush() after each major section to show progress

### Requirement 8: Monitor and Log Performance

**User Story:** As a developer, I want to measure execution time, so that I can identify bottlenecks and validate optimizations.

#### Acceptance Criteria

1. WHEN the generateCalculationsSheet function starts, THE System SHALL record the start timestamp
2. WHEN major operations complete, THE System SHALL log elapsed time for each section
3. WHEN the function completes, THE System SHALL log total execution time
4. WHEN performance issues occur, THE System SHALL log the number of tickers processed
5. WHEN debugging is needed, THE System SHALL provide detailed timing breakdowns for each operation type
