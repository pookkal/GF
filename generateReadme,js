/**
* ==============================================================================
* STABLE_MASTER_ALL_CLEAN_v3.1_KIRO_OPTIMIZED
* ==============================================================================
*/

function generateReferenceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "REFERENCE_GUIDE";
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear().clearFormats();

  const rows = [];

  // Title
  rows.push(["DUAL-MODE SIGNAL ENGINE — USER GUIDE", "", "", ""]);
  rows.push(["Professional-grade INVEST & TRADE decision systems with enhanced pattern recognition", "", "", ""]);
  rows.push(["", "", "", ""]);

  // OVERVIEW
  rows.push(["1) SYSTEM OVERVIEW", "", "", ""]);
  rows.push(["FEATURE", "DESCRIPTION", "TECHNICAL SPECS", "BUSINESS VALUE"]);
  rows.push([
    "Dual-Mode Engine",
    "Two distinct signal systems for different investment horizons",
    "INVEST (long-term) vs TRADE (tactical) modes via INPUT!E2",
    "Optimized decision logic for institutional vs active trading"
  ]);
  rows.push([
    "Enhanced Pattern Recognition",
    "Advanced volatility, ATH, and mean-reversion pattern detection",
    "ATR expansion, BBP extremes, ATH psychological zones",
    "Captures institutional-grade setups missed by basic indicators"
  ]);
  rows.push([
    "Position-Aware Logic",
    "Different decisions for owned vs unowned positions",
    "PURCHASED tag detection with risk/reward optimization",
    "Prevents overtrading and optimizes position management"
  ]);
  rows.push([
    "Dynamic Risk Sizing",
    "ATR and ATH-adjusted position sizing with volatility regimes",
    "Base 2% with 0.5x-1.5x multipliers based on risk metrics",
    "Institutional-grade risk management with volatility adaptation"
  ]);
  rows.push(["", "", "", ""]);

  // MODE SELECTION GUIDE
  rows.push(["2) MODE SELECTION GUIDE", "", "", ""]);
  rows.push(["CRITERIA", "INVEST MODE", "TRADE MODE", "SWITCHING LOGIC"]);
  rows.push([
    "Time Horizon",
    "3-12 months holding periods",
    "Days to weeks holding periods",
    "Set INPUT!E2 = TRUE for INVEST, FALSE for TRADE"
  ]);
  rows.push([
    "Risk Tolerance",
    "Lower turnover, trend-following",
    "Higher turnover, momentum-based",
    "INVEST for stability, TRADE for active management"
  ]);
  rows.push([
    "Market Conditions",
    "Bull markets, trending environments",
    "Volatile markets, range-bound conditions",
    "Switch based on market regime and VIX levels"
  ]);
  rows.push([
    "Portfolio Role",
    "Core holdings, strategic positions",
    "Tactical allocation, opportunistic plays",
    "Use INVEST for 70-80% core, TRADE for 20-30% tactical"
  ]);
  rows.push(["", "", "", ""]);

  // INVEST MODE DETAILED
  rows.push(["3) INVEST MODE — INSTITUTIONAL DECISION SYSTEM", "", "", ""]);
  rows.push(["SIGNAL", "TRIGGER CONDITIONS", "DECISION LOGIC", "EXPECTED OUTCOME"]);
  
  // Enhanced Pattern Signals
  rows.push([
    "ATH BREAKOUT",
    "ATH Diff ≥ -1% + Volume ≥ 1.5x + ADX ≥ 20",
    "🟢 STRONG BUY (if not purchased) / 🟢 ADD (if purchased + VALUE/FAIR)",
    "Momentum continuation at new highs with institutional participation"
  ]);
  rows.push([
    "VOLATILITY BREAKOUT", 
    "ATR > 20-period avg * 1.5 + Volume ≥ 2.0x + Price > Resistance",
    "🟢 STRONG BUY (if not purchased) / 🟢 ADD (if purchased + VALUE)",
    "Explosive moves with expanding volatility - institutional accumulation"
  ]);
  rows.push([
    "EXTREME OVERSOLD BUY",
    "BBP ≤ 0.1 + RSI ≤ 25 + Stoch ≤ 0.20 + Price > SMA200",
    "🟢 STRONG BUY (if not purchased) / 🟢 ADD (if purchased)",
    "Multi-indicator oversold confirmation in established uptrend"
  ]);
  
  // Standard Institutional Signals
  rows.push([
    "STRONG BUY",
    "Price > SMA200 + SMA50 > SMA200 + RSI ≤ 30 + MACD > 0 + ADX ≥ 20 + Vol ≥ 1.5x",
    "🟢 STRONG BUY (if not purchased) / 🟢 ADD (if purchased + VALUE/FAIR)",
    "Perfect storm of trend, momentum, and oversold conditions"
  ]);
  rows.push([
    "BUY",
    "Price > SMA200 + SMA50 > SMA200 + RSI ≤ 40 + MACD > 0 + ADX ≥ 15",
    "🟢 BUY (if not purchased) / 🟢 ADD (if purchased + VALUE/FAIR)",
    "Solid uptrend structure with reasonable entry point"
  ]);
  rows.push([
    "ACCUMULATE",
    "Price > SMA200 + RSI ≤ 35 + Price ≥ SMA50 * 0.95",
    "🟢 BUY (if not purchased) / 🟢 ADD (if purchased + VALUE/FAIR)",
    "Dip buying opportunity in established uptrend"
  ]);
  
  // Risk Management Signals
  rows.push([
    "STOP OUT",
    "Price < Support Level",
    "🔴 EXIT (any position) / 🔴 AVOID (no position)",
    "Capital preservation - structural breakdown"
  ]);
  rows.push([
    "RISK OFF",
    "Price < SMA200",
    "🔴 EXIT (if purchased) / 🔴 AVOID (if not purchased)",
    "Long-term trend invalidation - bearish regime"
  ]);
  rows.push([
    "OVERBOUGHT",
    "RSI ≥ 80 OR BBP ≥ 0.9",
    "🟠 TRIM (if purchased + EXPENSIVE) / ⏳ WAIT (if not purchased)",
    "Pullback warning - take profits or wait for better entry"
  ]);
  
  // Neutral States
  rows.push([
    "HOLD",
    "Price > SMA200 + 40 < RSI < 70",
    "⚖️ HOLD (any position)",
    "Neutral conditions in uptrend - maintain positions"
  ]);
  rows.push([
    "OVERSOLD",
    "RSI ≤ 20",
    "🟡 WATCH (if not purchased) / ⚖️ HOLD (if purchased)",
    "Potential bounce opportunity - monitor for confirmation"
  ]);
  rows.push(["", "", "", ""]);

  // TRADE MODE DETAILED
  rows.push(["4) TRADE MODE — TACTICAL DECISION SYSTEM", "", "", ""]);
  rows.push(["SIGNAL", "TRIGGER CONDITIONS", "DECISION LOGIC", "EXPECTED OUTCOME"]);
  
  // Enhanced Pattern Signals
  rows.push([
    "VOLATILITY BREAKOUT",
    "ATR > 20-period avg * 1.5 + Volume ≥ 2.0x + Price > Resistance",
    "Strong Trade Long (if not purchased + VALUE/FAIR)",
    "High-probability momentum continuation with institutional flow"
  ]);
  rows.push([
    "ATH BREAKOUT",
    "ATH Diff ≥ -1% + Volume ≥ 1.5x + ADX ≥ 20",
    "Strong Trade Long (if not purchased + VALUE/FAIR)",
    "New high momentum with strong participation"
  ]);
  
  // Standard Tactical Signals
  rows.push([
    "BREAKOUT",
    "Volume ≥ 1.5x + Price ≥ Resistance * 0.995",
    "Trade Long (if not purchased + VALUE/FAIR)",
    "Volume-confirmed breakout above resistance"
  ]);
  rows.push([
    "MOMENTUM",
    "Price > SMA200 + MACD > 0 + ADX ≥ 20",
    "Accumulate (if not purchased + VALUE) / Hold (if purchased)",
    "Strong trending conditions - ride the trend"
  ]);
  rows.push([
    "UPTREND",
    "Price > SMA200 + SMA50 > SMA200 + ADX ≥ 15",
    "Hold (any position)",
    "Basic uptrend structure - trend following"
  ]);
  rows.push([
    "BULLISH",
    "Price > SMA50 + Price > SMA20",
    "Hold (any position)",
    "Above key moving averages - short-term bullish bias"
  ]);
  
  // Mean Reversion Signals
  rows.push([
    "OVERSOLD",
    "Stoch %K ≤ 0.20 + Price > Support",
    "Add in Dip (if purchased) / Hold (if not purchased)",
    "Oversold bounce opportunity in uptrend"
  ]);
  rows.push([
    "VOLATILITY SQUEEZE",
    "ATR < 20-period avg * 0.7 + ADX < 15 + BBP near 0.5",
    "Wait for Breakout (any position)",
    "Coiling pattern - awaiting directional move"
  ]);
  
  // Risk Management
  rows.push([
    "STOP OUT",
    "Price < Support Level",
    "Stop-Out (any position)",
    "Trend invalidated - preserve capital"
  ]);
  rows.push([
    "OVERBOUGHT",
    "RSI ≥ 80",
    "Take Profit (if purchased) / Avoid (if not purchased)",
    "Pullback risk from elevated levels"
  ]);
  rows.push([
    "RANGE",
    "ADX < 15",
    "Hold (any position)",
    "No trend - sideways action, range tactics only"
  ]);
  rows.push(["", "", "", ""]);

  // DECISION MATRIX
  rows.push(["5) COMPREHENSIVE DECISION MATRIX", "", "", ""]);
  rows.push(["POSITION STATUS", "SIGNAL TYPE", "FUNDAMENTAL", "FINAL DECISION"]);
  
  // INVEST MODE DECISIONS
  rows.push(["INVEST - PURCHASED", "ATH/VOL BREAKOUT + EXTREME OVERSOLD", "VALUE/FAIR", "🟢 ADD"]);
  rows.push(["INVEST - PURCHASED", "STRONG BUY/BUY/ACCUMULATE", "VALUE/FAIR", "🟢 ADD"]);
  rows.push(["INVEST - PURCHASED", "STRONG BUY/BUY/ACCUMULATE", "EXPENSIVE", "? HOLD / ADD SMALL"]);
  rows.push(["INVEST - PURCHASED", "STRONG BUY/BUY/ACCUMULATE", "PRICED FOR PERFECTION", "🟡 HOLD (NO ADD)"]);
  rows.push(["INVEST - PURCHASED", "OVERBOUGHT", "EXPENSIVE/PERFECTION", "🟠 TRIM"]);
  rows.push(["INVEST - PURCHASED", "STOP OUT/RISK OFF", "ANY", "🔴 EXIT"]);
  rows.push(["INVEST - PURCHASED", "HOLD/NEUTRAL", "ANY", "⚖️ HOLD"]);
  
  rows.push(["INVEST - NOT PURCHASED", "ATH/VOL BREAKOUT + EXTREME OVERSOLD", "ANY", "🟢 STRONG BUY"]);
  rows.push(["INVEST - NOT PURCHASED", "STRONG BUY", "ANY", "🟢 STRONG BUY"]);
  rows.push(["INVEST - NOT PURCHASED", "BUY/ACCUMULATE", "ANY", "🟢 BUY"]);
  rows.push(["INVEST - NOT PURCHASED", "OVERSOLD", "ANY", "🟡 WATCH (OVERSOLD)"]);
  rows.push(["INVEST - NOT PURCHASED", "OVERBOUGHT", "ANY", "⏳ WAIT (OVERBOUGHT)"]);
  rows.push(["INVEST - NOT PURCHASED", "STOP OUT/RISK OFF", "ANY", "🔴 AVOID"]);
  rows.push(["INVEST - NOT PURCHASED", "HOLD/NEUTRAL", "ANY", "⚪ NEUTRAL"]);
  
  // TRADE MODE DECISIONS
  rows.push(["TRADE - PURCHASED", "STOP OUT", "ANY", "Stop-Out"]);
  rows.push(["TRADE - PURCHASED", "OVERBOUGHT/TARGET REACHED", "ANY", "Take Profit"]);
  rows.push(["TRADE - PURCHASED", "RISK OFF", "ANY", "Risk-Off"]);
  rows.push(["TRADE - PURCHASED", "MOMENTUM/UPTREND/BULLISH", "ANY", "Hold"]);
  
  rows.push(["TRADE - NOT PURCHASED", "VOL/ATH BREAKOUT", "VALUE/FAIR", "Strong Trade Long"]);
  rows.push(["TRADE - NOT PURCHASED", "BREAKOUT", "VALUE/FAIR", "Trade Long"]);
  rows.push(["TRADE - NOT PURCHASED", "MOMENTUM", "VALUE", "Accumulate"]);
  rows.push(["TRADE - NOT PURCHASED", "OVERSOLD", "ANY", "Add in Dip"]);
  rows.push(["TRADE - NOT PURCHASED", "VOLATILITY SQUEEZE", "ANY", "Wait for Breakout"]);
  rows.push(["TRADE - NOT PURCHASED", "RISK OFF", "ANY", "Avoid"]);
  rows.push(["TRADE - NOT PURCHASED", "RANGE/NEUTRAL", "ANY", "Hold"]);
  rows.push(["", "", "", ""]);

  // ENHANCED FEATURES
  rows.push(["6) ENHANCED PATTERN RECOGNITION FEATURES", "", "", ""]);
  rows.push(["FEATURE", "CALCULATION", "INTERPRETATION", "TRADING APPLICATION"]);
  rows.push([
    "Volatility Regime",
    "ATR / Price ratio with 4 levels: LOW/NORMAL/HIGH/EXTREME",
    "LOW VOL = larger position sizes, EXTREME VOL = smaller sizes",
    "Dynamic position sizing based on current volatility environment"
  ]);
  rows.push([
    "ATH Psychological Zones",
    "6 zones from AT ATH to DEEP VALUE based on % from highs",
    "AT ATH = resistance, DEEP VALUE = potential value opportunity",
    "Psychological level awareness for entry/exit timing"
  ]);
  rows.push([
    "BBP Mean Reversion",
    "Bollinger %B with RSI confirmation for extreme readings",
    "EXTREME OVERSOLD/OVERBOUGHT = high-probability reversals",
    "Enhanced mean reversion signals with multiple confirmations"
  ]);
  rows.push([
    "Pattern Detection",
    "Multi-indicator pattern recognition (VOL BREAKOUT, SQUEEZE, etc.)",
    "Combines volatility, momentum, and mean reversion patterns",
    "Institutional-grade setup identification"
  ]);
  rows.push([
    "ATR-Based Stops/Targets",
    "Dynamic stops at 2x ATR, targets at 3x ATR from entry",
    "Volatility-adjusted risk management levels",
    "Professional risk/reward optimization"
  ]);
  rows.push(["", "", "", ""]);

  // OPERATIONAL GUIDE
  rows.push(["7) OPERATIONAL USER GUIDE", "", "", ""]);
  rows.push(["TASK", "PROCEDURE", "LOCATION", "EXPECTED RESULT"]);
  rows.push([
    "Switch Modes",
    "Toggle TRUE/FALSE in cell",
    "INPUT!E2",
    "Changes all signal calculations between INVEST/TRADE logic"
  ]);
  rows.push([
    "Mark Position",
    "Add/Remove PURCHASED tag",
    "INPUT!C column (ticker row)",
    "Enables position-aware decision logic"
  ]);
  rows.push([
    "Refresh Data",
    "Click refresh trigger",
    "INPUT!E1",
    "Updates all market data and calculations"
  ]);
  rows.push([
    "View Signals",
    "Check SIGNAL column",
    "CALCULATIONS!B or DASHBOARD",
    "Current technical signal for each ticker"
  ]);
  rows.push([
    "View Decisions",
    "Check DECISION column",
    "CALCULATIONS!D or DASHBOARD",
    "Position-aware action recommendation"
  ]);
  rows.push([
    "Monitor Risk",
    "Check POSITION SIZE column",
    "CALCULATIONS!Z",
    "ATR and ATH-adjusted position sizing"
  ]);
  rows.push([
    "Read Analysis",
    "Check TECH/FUND NOTES",
    "CALCULATIONS!AA/AB or REPORT sheet",
    "Detailed reasoning for signals and decisions"
  ]);
  rows.push(["", "", "", ""]);

  // PERFORMANCE BENCHMARKS
  rows.push(["8) PERFORMANCE BENCHMARKS & VALIDATION", "", "", ""]);
  rows.push(["METRIC", "INVEST MODE TARGET", "TRADE MODE TARGET", "MONITORING METHOD"]);
  rows.push([
    "Signal Accuracy",
    "> 65% profitable signals",
    "> 55% profitable signals",
    "Track P&L by signal type over 3-6 months"
  ]);
  rows.push([
    "Risk-Adjusted Returns",
    "Sharpe Ratio > 1.5",
    "Sharpe Ratio > 1.2",
    "Monthly return / volatility calculation"
  ]);
  rows.push([
    "Maximum Drawdown",
    "< 15% peak-to-trough",
    "< 20% peak-to-trough",
    "Track largest loss from any peak"
  ]);
  rows.push([
    "Position Sizing Efficiency",
    "Outperform equal-weight by 200+ bps",
    "Outperform equal-weight by 150+ bps",
    "Compare vs equal-weight portfolio returns"
  ]);
  rows.push([
    "Turnover Rate",
    "< 200% annually",
    "< 500% annually",
    "Sum of buys + sells / average portfolio value"
  ]);
  rows.push(["", "", "", ""]);

  // Write all rows
  sh.getRange(1, 1, rows.length, 4).setValues(rows);

  // Professional styling
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 400);
  sh.setColumnWidth(3, 350);
  sh.setColumnWidth(4, 350);
  sh.setRowHeights(1, Math.min(rows.length, 900), 20);
  sh.setFrozenRows(3);

  // Title styling
  sh.getRange("A1:D1").merge()
    .setBackground("#1B4332").setFontColor("white")
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.getRange("A2:D2").merge()
    .setBackground("#2D5A3D").setFontColor("#FFEB3B")
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Section headers
  for (let r = 1; r <= rows.length; r++) {
    const v = String(sh.getRange(r, 1).getValue() || "");
    if (/^\d\)/.test(v)) {
      sh.getRange(r, 1, 1, 4).merge()
        .setBackground("#1B4332").setFontColor("white")
        .setFontWeight("bold").setFontSize(11)
        .setHorizontalAlignment("left");
    }
  }

  // Table headers
  const headerTerms = ["FEATURE", "CRITERIA", "SIGNAL", "POSITION STATUS", "TASK", "METRIC"];
  for (let r = 1; r <= rows.length; r++) {
    const a = String(sh.getRange(r, 1).getValue() || "").trim();
    if (headerTerms.includes(a)) {
      sh.getRange(r, 1, 1, 4)
        .setBackground("#E8F5E8")
        .setFontWeight("bold")
        .setFontColor("#1B4332")
        .setHorizontalAlignment("center");
    }
  }

  // Global formatting
  sh.getRange(1, 1, rows.length, 4)
    .setWrap(true)
    .setVerticalAlignment("top")
    .setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

  // Alternating row colors
  const band = sh.getRange(4, 1, Math.max(1, rows.length - 3), 4).applyRowBanding();
  band.setHeaderRowColor("#FFFFFF");
  band.setFirstRowColor("#FFFFFF");
  band.setSecondRowColor("#F8F9FA");

  ss.toast("REFERENCE_GUIDE updated with comprehensive dual-mode user documentation.", "✅ INDUSTRY GRADE", 4);
}