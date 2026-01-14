# Dashboard Column Structure - Final Layout

## Overview
The Dashboard sheet has been restructured with 32 columns (A-AF) organized into 7 groups.

## Column Layout

### Group 1: IDENTITY (Column A)
- **A**: Ticker

### Group 2: SIGNALING (Columns B-D)
- **B**: SIGNAL
- **C**: FUNDAMENTAL
- **D**: DECISION

### Group 3: PRICE (Columns E-F)
- **E**: Price
- **F**: Change %

### Group 4: PERFORMANCE (Columns G-J)
- **G**: ATH (TRUE)
- **H**: ATH Diff %
- **I**: ATH ZONE ← **NEW** (from CALC AD)
- **J**: R:R Quality

### Group 5: TREND (Columns K-O)
- **K**: Trend Score
- **L**: Trend State
- **M**: SMA 20
- **N**: SMA 50
- **O**: SMA 200

### Group 6: MOMENTUM (Columns P-T)
- **P**: RSI
- **Q**: MACD Hist
- **R**: Divergence
- **S**: ADX (14)
- **T**: Stoch %K (14)

### Group 7: VOLUME / VOLATILITY (Columns U-Z)
- **U**: Vol Trend ← **MOVED** (from CALC G)
- **V**: ATR (14)
- **W**: Bollinger %B
- **X**: VOL REGIME
- **Y**: POSITION SIZE
- **Z**: EVENT ← **RENAMED** (was PATTERNS, from CALC AF)

### Group 8: TARGET (Columns AA-AF)
- **AA**: BBP SIGNAL ← **NEW** (from CALC AE)
- **AB**: Support
- **AC**: Resistance
- **AD**: Target (3:1) ← **MOVED** (from CALC W)
- **AE**: ATR STOP
- **AF**: ATR TARGET

## Mapping from CALCULATIONS Sheet

| Dashboard | CALCULATIONS | Description |
|-----------|--------------|-------------|
| A | A | Ticker |
| B | B | SIGNAL |
| C | C | FUNDAMENTAL |
| D | D | DECISION |
| E | E | Price |
| F | F | Change % |
| G | H | ATH (TRUE) |
| H | I | ATH Diff % |
| I | AD | ATH ZONE |
| J | J | R:R Quality |
| K | K | Trend Score |
| L | L | Trend State |
| M | M | SMA 20 |
| N | N | SMA 50 |
| O | O | SMA 200 |
| P | P | RSI |
| Q | Q | MACD Hist |
| R | R | Divergence |
| S | S | ADX (14) |
| T | T | Stoch %K (14) |
| U | G | Vol Trend |
| V | X | ATR (14) |
| W | Y | Bollinger %B |
| X | AC | VOL REGIME |
| Y | Z | POSITION SIZE |
| Z | AF | EVENT (was PATTERNS) |
| AA | AE | BBP SIGNAL |
| AB | U | Support |
| AC | V | Resistance |
| AD | W | Target (3:1) |
| AE | AG | ATR STOP |
| AF | AH | ATR TARGET |

## Key Changes Made

1. ✅ **PRICE / VOLUME renamed to PRICE** - now only contains Price and Change %
2. ✅ **Vol Trend moved to VOLUME / VOLATILITY** - from column G to column U
3. ✅ **ATH ZONE added to PERFORMANCE** - new column I after ATH Diff %
4. ✅ **PATTERNS renamed to EVENT** - column Z in VOLUME / VOLATILITY group
5. ✅ **BBP SIGNAL added to TARGET** - new column AA (first column in TARGET group)
6. ✅ **Target (3:1) moved to TARGET** - from standalone to column AD in TARGET group
7. ✅ **PATTERNS group renamed to TARGET** - now contains BBP SIGNAL, Support, Resistance, Target, ATR STOP, ATR TARGET
8. ✅ **INSTITUTIONAL and NOTES groups removed** - consolidated into TARGET group
9. ✅ **All conditional formatting updated** - rules adjusted for new column positions
10. ✅ **All number formats updated** - formats applied to correct columns

## Color Scheme

- **IDENTITY**: #263238 (Dark Gray)
- **SIGNALING**: #0D47A1 (Blue)
- **PRICE**: #1B5E20 (Green)
- **PERFORMANCE**: #004D40 (Teal)
- **TREND**: #2E7D32 (Dark Green)
- **MOMENTUM**: #33691E (Olive)
- **VOLUME / VOLATILITY**: #B71C1C (Red)
- **TARGET**: #6A1B9A (Purple)

## Notes

- Total columns: 32 (A-AF)
- All columns are visible (no hidden columns)
- TECH NOTES removed from Dashboard (still available in CALCULATIONS sheet column AA)
- Market indices (NIFTY 50, S&P 500) remain in row 1 with conditional formatting
- Checkboxes in B1 and D1 reset to false after each update
