# Stock Portfolio Analyzer

Fetches live market data for a list of stock tickers and generates a formatted Excel workbook with portfolio analysis, P&L tracking, and an AI-powered insights sheet powered by Claude.

## Quick start

```bash
pip install -r requirements.txt
python3 stock_analyzer.py
```

See [SETUP.md](SETUP.md) for API key configuration and full instructions.

## Features

- **Live market data** via [yfinance](https://github.com/ranaroussi/yfinance) — no API key required for market data
- **P&L tracking** — provide buy date and shares in `tickers.txt`; historical buy price is fetched automatically
- **Multi-currency** — USD, EUR, DKK, CHF, and any Yahoo Finance currency
- **Three Excel sheets:**

| Sheet | Contents |
|-------|----------|
| **Portfolio** | Current prices · returns 1W→5Y · P&L · fundamentals (P/E, beta, 52W high/low, dividend yield) |
| **Distribution** | Country and sector breakdown — auto-updates via SUMIF formulas as you fill in share counts |
| **AI Analysis** | Structured portfolio insights from Claude: diversification, performance, risk, concentration warnings |

- **Conditional formatting** — green/red fills on all return and P&L columns
- **Freeze panes** — ticker and name stay visible when scrolling right
- **Editable shares column** — highlighted yellow; all dependent cells update automatically

## Tickers file format

```
# TICKER[,SHARES[,BUY_DATE]]
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
| E – K | Buy Date · Buy Price · **# Shares** · Total Value · Cost Basis · P&L $ · P&L % |
| L – U | Returns: 1W · 1M · 3M · 6M · YTD · 1Y · 2Y · 3Y · 4Y · 5Y |
| V – AE | Country · Sector · Industry · Market Cap · P/E · 52W High · 52W Low · % from 52W High · Div Yield · Beta |

## AI Analysis sheet

Calls `claude-sonnet-4-6` with the full portfolio data and produces a structured analysis covering:

- Portfolio Overview
- Geographic & Sector Diversification
- Performance Highlights (with P&L context where available)
- Risk Assessment (beta, valuation, drawdowns)
- Sector Concentration Warnings
- Key Takeaways

Requires `ANTHROPIC_API_KEY` in `.env` with API credits. Skipped gracefully if absent.

## Requirements

- Python 3.8+
- Dependencies: `yfinance`, `openpyxl`, `pandas`, `anthropic`, `python-dotenv`
