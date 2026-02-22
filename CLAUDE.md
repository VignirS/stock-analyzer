# CLAUDE.md — Project Context for Claude Code

## What this project does
Reads stock tickers from `tickers.txt`, fetches live market data via yfinance (no API key needed),
and generates a multi-sheet Excel workbook with portfolio analysis and Claude AI insights.

## Key files
- `stock_analyzer.py` — single-file main script (~690 lines)
- `tickers.txt` — input file, format: `TICKER[,SHARES[,BUY_DATE]]`
- `.env` — contains `ANTHROPIC_API_KEY` (never commit this file)
- `.env.example` — template to copy from
- `requirements.txt` — 5 dependencies

## How to run
```bash
python3 stock_analyzer.py              # uses tickers.txt
python3 stock_analyzer.py my_list.txt  # custom file
```
Output: `portfolio_YYYYMMDD_HHMM.xlsx` — excluded from git via `.gitignore`.

## Python version
System Python 3.9.6 on macOS. Do NOT use `list[str]` or `dict[str, int]` type annotation
subscripts (not supported in 3.9) — use bare `list` or `typing.List` / `typing.Dict` instead.

## Excel column layout
| Range | Content |
|-------|---------|
| A–D   | Identity: Ticker, Company Name, Current Price, Currency |
| E–K   | Position/P&L: Buy Date, Buy Price, # Shares (yellow), Total Value, Cost Basis, P&L $, P&L % |
| L–U   | Returns: 1W%, 1M%, 3M%, 6M%, YTD%, 1Y%, 2Y%, 3Y%, 4Y%, 5Y% |
| V–AE  | Fundamentals: Country, Sector, Industry, Market Cap, P/E, 52W High/Low, % from 52W High, Div Yield, Beta |

- `COLUMNS` list drives everything — order, headers, formats, widths
- `PCT_COLS = COLUMNS[11:21]` — the returns block (L–U)
- `LAST_COL = "AE"` — used for title merge and autofilter range
- Freeze panes at `E4` — keeps A–D visible when scrolling right
- Distribution sheet SUMIF/COUNTIF formulas reference: Country=V, Sector=W, Total Value=H

## Style constants
All colours defined at the top of `stock_analyzer.py`:
`DARK_BLUE`, `MED_BLUE`, `LIGHT_BLUE`, `ALT_ROW`, `YELLOW_FILL`, `GREEN_FILL`, etc.

## AI Analysis sheet
- Called via `anthropic.Anthropic(api_key=...).messages.stream(model="claude-sonnet-4-6", ...)`
- Requires `ANTHROPIC_API_KEY` in `.env` with credits on console.anthropic.com
- Skipped gracefully if key is absent or invalid — other two sheets still generated
- Prompt sends full portfolio JSON; expects structured Markdown with 6 sections

## Common tasks
- **Add a ticker**: edit `tickers.txt`, one line per ticker
- **Add a column**: add a tuple to `COLUMNS`, handle the cell value in `create_portfolio_sheet()`
- **Change AI model**: update the `model=` string in `create_ai_sheet()`
- **Fix date display issues**: always use lowercase `yyyy-mm-dd` in `number_format` — Excel treats uppercase Y/D as literal characters

## Dependencies
```
yfinance>=0.2.50    # market data
openpyxl>=3.1.0     # Excel generation
pandas>=2.0.0       # data manipulation
anthropic>=0.40.0   # Claude API
python-dotenv>=1.0.0 # .env loading
```
