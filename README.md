# Stock Portfolio Analyzer

Fetches live market data for a list of stock tickers and generates a formatted Excel workbook with portfolio analysis, EUR-normalised P&L tracking, charts, and an AI-powered insights sheet powered by Claude.

## Quick start

```bash
pip install -r requirements.txt
python3 stock_analyzer.py --open
```

See [SETUP.md](SETUP.md) for API key configuration and full instructions.

## Features

- **Live market data** via [yfinance](https://github.com/ranaroussi/yfinance) — no API key required for market data
- **P&L tracking** — provide buy date and shares in `tickers.txt`; historical buy price is fetched automatically
- **EUR normalisation** — live FX rates fetched for all currencies; Total Value and P&L shown in EUR alongside native currency
- **tqdm progress bar** — visual fetch progress; degrades gracefully if tqdm is not installed
- **Five Excel sheets:**

| Sheet | Contents |
|-------|----------|
| **Portfolio** | Current prices · returns 1W→5Y · P&L in native currency and EUR · FX rate · fundamentals (P/E, beta, 52W high/low, dividend yield) |
| **Distribution** | Country and sector breakdown with EUR-normalised totals — auto-updates via SUMIF formulas |
| **Charts** | Sector distribution pie chart · Unrealised P&L % bar chart |
| **Errors** | Failed tickers with error messages (only shown when fetches fail) |
| **AI Analysis** | Structured portfolio insights from Claude: diversification, performance, risk, concentration warnings |

- **Conditional formatting** — green/red fills on all return, P&L, and EUR P&L columns
- **Freeze panes** — ticker/name/price stay visible when scrolling right or down
- **Editable shares column** — highlighted yellow; all dependent cells update automatically

## Usage

```bash
python3 stock_analyzer.py                          # uses tickers.txt, auto-timestamped output
python3 stock_analyzer.py my_list.txt              # custom tickers file
python3 stock_analyzer.py --output report.xlsx     # fixed output filename
python3 stock_analyzer.py --no-ai                  # skip AI sheet even if key is set
python3 stock_analyzer.py --open                   # open file after generation
python3 stock_analyzer.py --watch                  # re-run on every tickers file save
python3 stock_analyzer.py --model claude-opus-4-6  # use a different Claude model
python3 stock_analyzer.py --add "TSLA,10,2024-03-01"  # append ticker and exit
```

## Tickers file format

```
# TICKER[,SHARES[,BUY_DATE]]  # inline comments are supported
AAPL,50,2022-06-15       # full entry: fetches buy price, shows P&L
MSFT,20                  # shares only: no P&L
NESN.SW                  # ticker only: price and fundamentals only
```

Supports any Yahoo Finance symbol including international exchanges:
`ASML.AS` (Amsterdam), `SAP.DE` (Frankfurt), `NESN.SW` (Zurich), `NOVO-B.CO` (Copenhagen).

## Excel column layout

| Columns | Content |
|---------|---------|
| A – D | Ticker · Company Name · Current Price · Currency |
| E – K | Buy Date · Buy Price · **# Shares** (yellow) · Total Value · Cost Basis · P&L $ · P&L % |
| L – U | Returns: 1W · 1M · 3M · 6M · YTD · 1Y · 2Y · 3Y · 4Y · 5Y |
| V – AE | Country · Sector · Industry · Market Cap · P/E · 52W High · 52W Low · % from 52W High · Div Yield · Beta |
| AF – AH | **FX Rate (→EUR)** · **Total Value (EUR)** · **P&L (EUR)** |

## AI Analysis sheet

Calls `claude-sonnet-4-6` with the full portfolio data and produces a structured Markdown analysis covering:

- Portfolio Overview
- Geographic & Sector Diversification
- Performance Highlights (with P&L context where available)
- Risk Assessment (beta, valuation, drawdowns)
- Sector Concentration Warnings
- Key Takeaways

Requires `ANTHROPIC_API_KEY` or `ANTHROPIC_AUTH_TOKEN` in `.env`. Skipped gracefully if absent.

## Requirements

- Python 3.8+
- Dependencies: `yfinance`, `openpyxl`, `pandas`, `anthropic`, `python-dotenv`, `tqdm`
