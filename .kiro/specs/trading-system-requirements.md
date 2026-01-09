# Trading System Requirements Specification

## Executive Summary

This specification documents a comprehensive Google Sheets-based institutional trading terminal that provides advanced technical analysis, risk management, and decision-making capabilities for equity trading. The system integrates multiple sheets with sophisticated formulas, custom JavaScript functions, and automated monitoring to deliver professional-grade trading intelligence.

## System Architecture

### Core Components

The trading system consists of six interconnected Google Sheets:

1. **INPUT Sheet** - Ticker management and configuration
2. **DATA Sheet** - Market data aggregation and storage  
3. **CALCULATIONS Sheet** - Technical indicator computations and decision engine
4. **DASHBOARD Sheet** - Multi-ticker overview with Bloomberg-style formatting
5. **CHART Sheet** - Interactive charting with dynamic controls
6. **REPORT Sheet** - Mobile-optimized single-ticker analysis

### Technology Stack

- **Platform**: Google Sheets with Google Apps Script
- **Data Sources**: Google Finance API for real-time market data
- **Programming Languages**: JavaScript (Google Apps Script)
- **UI Framework**: Custom HTML/CSS for popups and reports
- **Automation**: Time-based triggers for monitoring and alerts

## Functional Requirements

### FR-001: Market Data Management

**Description**: The system shall aggregate and process real-time market data for multiple securities.

**Acceptance Criteria**:
- Fetch real-time price, volume, and fundamental data via Google Finance API
- Support 50+ concurrent tickers with automatic data refresh
- Calculate all-time highs (ATH) with historical data analysis
- Store P/E ratios, EPS, and other fundamental metrics
- Implement market regime detection (USA and India markets)
- Handle data errors gracefully with fallback mechanisms

**Implementation**: DATA sheet with automated GOOGLEFINANCE formulas and error handling

### FR-002: Technical Indicator Engine

**Description**: The system shall compute comprehensive technical indicators using custom algorithms.

**Technical Indicators Required**:
- **Moving Averages**: SMA 20, 50, 200 with trend alignment scoring
- **Momentum Oscillators**: RSI (14), MACD histogram, Stochastic %K
- **Trend Strength**: ADX (14) with directional movement analysis
- **Volatility Measures**: ATR (14), Bollinger Bands %B
- **Volume Analysis**: Relative volume (RVOL) vs 20-day average
- **Support/Resistance**: Dynamic levels using percentile-based calculations

**Acceptance Criteria**:
- All indicators update in real-time with price changes
- Custom JavaScript functions (LIVERSI, LIVEMACD, LIVEADX, LIVEATR) provide accurate calculations
- Indicators handle edge cases (insufficient data, market holidays)
- Performance optimized for 50+ tickers simultaneously

**Implementation**: CALCULATIONS sheet with custom Google Apps Script functions

### FR-003: Dual-Mode Signal Engine

**Description**: The system shall provide two distinct signal generation modes for different trading strategies.

**Mode 1: INVEST (Long-term)**
- Focus on trend-following and institutional accumulation patterns
- Enhanced pattern recognition (ATH breakouts, volatility breakouts, extreme oversold)
- Position-aware logic differentiating owned vs unowned securities
- Risk management with SMA200 regime filtering

**Mode 2: TRADE (Tactical)**  
- Momentum-based signals for active trading
- Breakout detection with volume confirmation
- Mean reversion setups with multiple oscillator confirmation
- Volatility squeeze identification for coiling patterns

**Acceptance Criteria**:
- Mode selection via INPUT!E2 toggle (TRUE = INVEST, FALSE = TRADE)
- Signal accuracy >65% for INVEST mode, >55% for TRADE mode
- Clear signal hierarchy with 15+ distinct signal types
- Real-time signal updates with price movements

**Implementation**: Complex IFS formulas in CALCULATIONS sheet with mode-dependent logic

### FR-004: Position-Aware Decision Engine

**Description**: The system shall generate different recommendations based on current position status.

**Position Detection**:
- "PURCHASED" tag in INPUT sheet column C triggers position-aware logic
- Different decision trees for owned vs unowned securities
- Risk management overrides for stop-loss and profit-taking

**Decision Categories**:
- **Entry Signals**: STRONG BUY, BUY, ACCUMULATE, OVERSOLD BUY
- **Hold Signals**: HOLD, WATCH, MONITOR
- **Exit Signals**: STOP OUT, RISK OFF, TAKE PROFIT, TRIM
- **Avoidance**: AVOID, WAIT

**Acceptance Criteria**:
- Position status automatically detected from INPUT sheet tags
- Decision logic prevents overtrading and optimizes position management
- Clear action recommendations with color-coded formatting
- Integration with fundamental analysis (VALUE/FAIR/EXPENSIVE ratings)

### FR-005: Advanced Pattern Recognition

**Description**: The system shall identify institutional-grade trading patterns and setups.

**Pattern Types**:
- **ATH Breakout**: New highs with volume and momentum confirmation
- **Volatility Breakout**: ATR expansion >50% with volume surge
- **Extreme Oversold**: Multi-indicator oversold alignment in uptrends
- **Volatility Squeeze**: Low ATR with range compression
- **Mean Reversion**: BBP extremes with RSI confirmation

**Enhanced Features**:
- **Volatility Regime Classification**: LOW/NORMAL/HIGH/EXTREME based on ATR/Price ratio
- **ATH Psychological Zones**: 6 zones from AT ATH to DEEP VALUE
- **BBP Mean Reversion Signals**: Enhanced Bollinger Band position analysis
- **ATR-Based Risk Management**: Dynamic stops at 2x ATR, targets at 3x ATR

**Acceptance Criteria**:
- Pattern detection accuracy >70% for major setups
- Real-time pattern identification with clear descriptions
- Integration with volume and momentum confirmations
- Risk-adjusted position sizing based on pattern type

### FR-006: Dynamic Risk Management

**Description**: The system shall provide sophisticated risk management with volatility-adjusted position sizing.

**Position Sizing Algorithm**:
- Base allocation: 2% of portfolio per position
- ATR-based volatility adjustment (0.5x to 1.5x multiplier)
- ATH proximity risk reduction (0.8x near all-time highs)
- Risk/reward ratio optimization (1.5x to 3x multiplier)
- Maximum position size cap: 8% per security

**Risk Controls**:
- Automatic stop-loss at support level breakdown
- SMA200 regime filtering (no new longs in RISK-OFF)
- Overextension warnings (>2x ATR from SMA20)
- Fundamental risk flags (EXPENSIVE, PRICED FOR PERFECTION)

**Acceptance Criteria**:
- Position sizes automatically calculated and displayed
- Risk metrics update in real-time with price changes
- Clear risk warnings and invalidation levels
- Maximum drawdown <15% for INVEST mode, <20% for TRADE mode

### FR-007: Interactive Dashboard

**Description**: The system shall provide a Bloomberg-style multi-ticker dashboard with professional formatting.

**Dashboard Features**:
- 50+ ticker overview with real-time updates
- Color-coded cells based on signal strength and risk levels
- Sortable columns with conditional formatting
- Hidden technical and fundamental notes columns
- Compact layout optimized for institutional use

**Visual Design**:
- Dark theme with professional color palette
- Green/red/yellow color coding for signals
- Alternating row colors for readability
- Bold headers with grouped sections
- Responsive layout for different screen sizes

**Acceptance Criteria**:
- Dashboard loads <5 seconds with 50 tickers
- Real-time updates without manual refresh
- Professional appearance matching institutional standards
- All data synchronized with CALCULATIONS sheet

### FR-008: Advanced Charting System

**Description**: The system shall provide interactive charting with multiple technical overlays.

**Chart Features**:
- **Price Series**: OHLC data with real-time updates
- **Moving Averages**: SMA 20, 50, 200 with color coding
- **Volume Bars**: Bull/bear volume separation
- **Support/Resistance Lines**: Dynamic levels from calculations
- **ATR-Based Levels**: Stop and target lines

**Interactive Controls**:
- Ticker selection dropdown
- Time period controls (years, months, days)
- Interval selection (daily/weekly)
- Indicator toggles (checkboxes for each overlay)
- Date range picker with historical data

**Acceptance Criteria**:
- Charts update automatically with control changes
- Smooth performance with 200+ data points
- Professional appearance with proper scaling
- Integration with CALCULATIONS sheet for levels

### FR-009: Mobile Report Generation

**Description**: The system shall generate comprehensive single-ticker reports optimized for mobile viewing.

**Report Sections**:
- **Market Snapshot**: Price, change, P/E, EPS, ATH analysis
- **Trend Analysis**: Moving averages, ADX, trend scoring
- **Momentum Oscillators**: MACD, Stochastic with interpretations
- **Volatility & Volume**: ATR, RVOL, Bollinger Bands analysis
- **Support & Resistance**: Key levels with risk/reward ratios
- **Enhanced Patterns**: Advanced pattern recognition results
- **Risk Management**: ATR-based stops and targets

**Report Features**:
- Formula-based implementation for real-time updates
- Narrative explanations for each indicator
- Color-coded conditional formatting
- Mobile-responsive layout
- Chart integration with toggle controls

**Acceptance Criteria**:
- Reports generate in <3 seconds
- All data synchronized with CALCULATIONS sheet
- Clear explanations suitable for mobile reading
- Professional formatting with consistent styling

### FR-010: Automated Monitoring & Alerts

**Description**: The system shall provide automated monitoring with email alerts for significant changes.

**Monitoring Features**:
- Decision change detection (CALCULATIONS column C)
- 30-minute monitoring intervals
- Email alerts for actionable signals
- Alert filtering to prevent spam

**Alert Types**:
- **Entry Signals**: STRONG BUY, BUY, ACCUMULATE
- **Exit Signals**: STOP OUT, TAKE PROFIT, RISK OFF
- **Risk Warnings**: Support breakdown, overextension

**Acceptance Criteria**:
- Reliable alert delivery within 30 minutes
- No false positives or duplicate alerts
- Clear alert messages with ticker and reasoning
- Easy start/stop controls for monitoring

## Non-Functional Requirements

### NFR-001: Performance

- Dashboard refresh time: <5 seconds for 50 tickers
- Individual ticker calculations: <1 second
- Chart rendering: <3 seconds with full indicators
- Memory usage: <100MB for entire workbook

### NFR-002: Reliability

- 99.5% uptime during market hours
- Graceful handling of Google Finance API failures
- Automatic error recovery and retry mechanisms
- Data consistency across all sheets

### NFR-003: Scalability

- Support up to 100 tickers simultaneously
- Handle 1000+ historical data points per ticker
- Efficient formula calculations to prevent timeouts
- Optimized data structures for large datasets

### NFR-004: Usability

- Intuitive interface requiring minimal training
- Clear visual indicators for all signals
- Consistent color coding and formatting
- Mobile-friendly report layouts

### NFR-005: Maintainability

- Modular code structure with clear separation of concerns
- Comprehensive inline documentation
- Version control for all script files
- Easy configuration updates via INPUT sheet

## User Stories

### Epic 1: Market Analysis

**US-001**: As a trader, I want to view real-time market data for multiple tickers so that I can monitor my watchlist efficiently.

**US-002**: As an analyst, I want to see comprehensive technical indicators so that I can make informed trading decisions.

**US-003**: As a portfolio manager, I want to understand market regime (RISK-ON/RISK-OFF) so that I can adjust my strategy accordingly.

### Epic 2: Signal Generation

**US-004**: As a long-term investor, I want INVEST mode signals so that I can identify quality accumulation opportunities.

**US-005**: As an active trader, I want TRADE mode signals so that I can capture short-term momentum moves.

**US-006**: As a risk manager, I want position-aware recommendations so that I can optimize my existing holdings.

### Epic 3: Risk Management

**US-007**: As a trader, I want automatic position sizing so that I can maintain consistent risk across positions.

**US-008**: As a portfolio manager, I want stop-loss levels so that I can protect against significant losses.

**US-009**: As an analyst, I want risk/reward ratios so that I can prioritize the best opportunities.

### Epic 4: Reporting & Visualization

**US-010**: As a trader, I want interactive charts so that I can visualize price action and technical levels.

**US-011**: As a mobile user, I want comprehensive reports so that I can analyze stocks on my phone.

**US-012**: As a portfolio manager, I want a professional dashboard so that I can present to clients and stakeholders.

### Epic 5: Automation & Monitoring

**US-013**: As a busy trader, I want automated alerts so that I don't miss important signal changes.

**US-014**: As a systematic trader, I want consistent monitoring so that I can maintain discipline in my approach.

**US-015**: As a portfolio manager, I want historical tracking so that I can measure system performance over time.

## Technical Architecture

### System Components

The trading system consists of 6 interconnected Google Sheets with comprehensive JavaScript automation:

1. **INPUT Sheet**: Configuration and ticker management
2. **DATA Sheet**: Market data storage with regime analysis
3. **CALCULATIONS Sheet**: 34-column technical analysis engine
4. **DASHBOARD Sheet**: Main institutional trading interface
5. **CHART Sheet**: Interactive charting with dynamic updates
6. **REPORT Sheet**: Mobile-optimized reporting with floating charts

### Core JavaScript Files

- **Code.js** (2,299 lines): Main system logic, onEdit triggers, sheet generation
- **mobilereport-formulas.js** (915 lines): Complete REPORT sheet implementation
- **IndicatorFuncs.js**: Technical indicator implementations (RSI, MACD, ATR, Stochastic)
- **Helper.js**: Utility functions and data processing
- **Monitor.js**: Market monitoring and alert system

### Data Flow Architecture

```
INPUT (Tickers) → DATA (Market Data + Regime) → CALCULATIONS (34 Indicators) → 
DASHBOARD/CHART/REPORT (Real-time Updates via onEdit triggers)
```

### CALCULATIONS Sheet Structure (34 Columns)

**Identity & Signaling (A-D)**
- A: Ticker
- B: SIGNAL (Enhanced pattern recognition)
- C: DECISION (Position management logic)
- D: FUNDAMENTAL (P/E, EPS analysis)

**Price & Volume (E-G)**
- E: Current Price
- F: Change %
- G: Volume Trend (Relative Volume)

**Performance Metrics (H-J)**
- H: All-Time High (ATH)
- I: ATH Difference %
- J: Risk/Reward Quality

**Trend Analysis (K-O)**
- K: Trend Score (Moving average alignment)
- L: Trend State (BULL/BEAR/NEUTRAL)
- M: SMA 20
- N: SMA 50
- O: SMA 200

**Momentum Indicators (P-T)**
- P: RSI (Wilder's smoothed)
- Q: MACD Histogram
- R: Divergence Detection
- S: ADX (Trend Strength)
- T: Stochastic %K

**Levels & Risk Management (U-Y)**
- U: Support Level
- V: Resistance Level
- W: Price Target (3:1 R/R)
- X: ATR (Average True Range)
- Y: Bollinger %B

**Institutional Features (Z-AH)**
- Z: Position Size (ATR & ATH adjusted)
- AA: Technical Notes
- AB: Fundamental Notes
- AC: Volatility Regime
- AD: ATH Zone Classification
- AE: Bollinger Band Position Signal
- AF: Enhanced Pattern Recognition
- AG: ATR-based Stop Loss
- AH: ATR-based Target

### Enhanced Signal Engine

**Long-term Signal Mode** (useLongTermSignal = true):
- STOP OUT: Price below support
- RISK OFF: Price below SMA200
- ATH BREAKOUT: New highs with volume/momentum
- VOLATILITY BREAKOUT: ATR expansion with volume
- EXTREME OVERSOLD BUY: Multiple oversold indicators
- STRONG BUY: All bullish conditions aligned
- BUY/ACCUMULATE: Good entry conditions
- HOLD/NEUTRAL: Stable conditions

**Trend Signal Mode** (useLongTermSignal = false):
- Focus on breakout patterns and momentum
- Volatility squeeze detection
- Enhanced pattern recognition
- Range-bound identification

### REPORT Sheet Implementation

**Mobile-Optimized Layout**:
- Row 1: Ticker dropdown (A1:C1 merged) + Date display (D1)
- Row 2: Date selection dropdowns (A2:C2) + Interval (D2)
- Row 3: Calculated date display (A3:B3) + Weekly/Daily (C3)
- Rows 4-6: Decision matrix (SIGNAL, FUNDAMENTAL, DECISION)
- Row 7: Market regime status
- Rows 8+: Sectioned data with narratives

**Chart Controls (E1:M2)**:
- PRICE, SMA20, SMA50, SMA200, VOLUME
- SUPPORT, RESISTANCE, ATR STOP, ATR TARGET
- Interactive checkboxes for series selection

**Floating Chart (E3:N22)**:
- Dynamic data from DATA sheet
- Real-time price and indicator overlays
- Volume bars on secondary axis
- Professional dark theme styling

### Key Algorithms Implementation

**RSI (LIVERSI Function)**:
- Wilder's smoothed RSI with 14-period default
- Defensive parsing and validation
- Chronological data handling
- Returns 50 for insufficient data

**MACD (LIVEMACD Function)**:
- 12/26/9 EMA configuration
- SMA seed for EMA initialization
- Histogram calculation (MACD - Signal)
- Industry-standard implementation

**ATR (LIVEATR Function)**:
- True Range calculation with Wilder smoothing
- Handles high/low/close arrays
- 14-period default
- Returns absolute price values (not percentages)

**Stochastic (LIVESTOCHK Function)**:
- 14-period %K calculation
- Optional smoothing parameter
- Robust data validation
- Returns 0.5 for insufficient data

### Market Regime Analysis

**USA Market Regime**:
- SPY price vs SMA200 ratio
- VIX level integration
- Classifications: STRONG BULL, BULL, NEUTRAL, BEAR, STRONG BEAR

**India Market Regime**:
- NIFTY_50 price vs SMA200 ratio
- India VIX integration
- Same classification system as USA

### Position Sizing Algorithm

**ATR & ATH Risk Adjusted**:
- Base size: 2% of portfolio
- ATR volatility adjustment (0.5x to 1.2x)
- ATH proximity risk reduction
- Risk/reward ratio multiplier
- Maximum position: 8%

### Integration Points

- **Google Finance API**: Real-time market data, fundamentals
- **Custom JavaScript Functions**: All technical indicators
- **Automated Chart Generation**: Dynamic series configuration
- **onEdit Triggers**: Real-time sheet updates
- **Market Monitoring**: Automated alert system
- **Email Notifications**: Signal and alert delivery

## Quality Assurance

### Testing Strategy

1. **Unit Testing**: Individual indicator calculations
2. **Integration Testing**: Cross-sheet data consistency
3. **Performance Testing**: Load testing with maximum tickers
4. **User Acceptance Testing**: Real-world trading scenarios
5. **Regression Testing**: Ensure updates don't break existing functionality

### Validation Criteria

- Signal accuracy measured against historical performance
- Risk management effectiveness in drawdown scenarios
- System reliability during high-volatility periods
- User satisfaction with interface and functionality

## Implementation Roadmap

### Phase 1: Core Infrastructure (Completed)
- ✅ Basic sheet structure and data connections
- ✅ Technical indicator calculations
- ✅ Signal generation engine
- ✅ Dashboard formatting and layout

### Phase 2: Advanced Features (Completed)
- ✅ Dual-mode signal engine (INVEST/TRADE)
- ✅ Position-aware decision logic
- ✅ Enhanced pattern recognition
- ✅ Dynamic risk management

### Phase 3: User Experience (Completed)
- ✅ Interactive charting system
- ✅ Mobile report generation
- ✅ Professional dashboard styling
- ✅ Automated monitoring and alerts

### Phase 4: Optimization & Enhancement (Ongoing)
- 🔄 Performance optimization for larger datasets
- 🔄 Additional technical indicators and patterns
- 🔄 Enhanced fundamental analysis integration
- 🔄 Advanced portfolio management features

## Success Metrics

### Performance Targets

- **Signal Accuracy**: >65% for INVEST mode, >55% for TRADE mode
- **Risk-Adjusted Returns**: Sharpe Ratio >1.5 (INVEST), >1.2 (TRADE)
- **Maximum Drawdown**: <15% (INVEST), <20% (TRADE)
- **System Uptime**: >99.5% during market hours
- **Response Time**: <5 seconds for dashboard refresh

### User Satisfaction

- Ease of use rating: >4.5/5
- Feature completeness: >90% of requirements met
- Performance satisfaction: >4.0/5
- Reliability rating: >4.5/5

## Risk Assessment

### Technical Risks

- **Google Finance API Limitations**: Mitigated by error handling and fallback mechanisms
- **Google Sheets Performance**: Addressed through optimized formulas and data structures
- **Script Execution Timeouts**: Managed via modular design and efficient algorithms

### Market Risks

- **Signal Accuracy Degradation**: Monitored through continuous performance tracking
- **Market Regime Changes**: Addressed by dual-mode engine and regime detection
- **Volatility Spikes**: Handled by dynamic risk management and position sizing

### Operational Risks

- **User Error**: Minimized through intuitive design and clear documentation
- **Data Corruption**: Prevented by version control and backup procedures
- **System Downtime**: Reduced through robust error handling and monitoring

## Conclusion

This trading system represents a comprehensive solution for institutional-grade equity analysis and trading. The combination of advanced technical analysis, sophisticated risk management, and professional presentation creates a powerful tool for traders, analysts, and portfolio managers.

The system's dual-mode architecture allows it to serve both long-term investors and active traders, while the position-aware decision engine ensures optimal portfolio management. The integration of real-time data, automated monitoring, and mobile-optimized reporting provides a complete trading workflow solution.

With its proven track record of reliable performance and continuous enhancement, this trading system establishes a new standard for Google Sheets-based financial analysis tools.