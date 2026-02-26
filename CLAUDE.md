# CLAUDE.md — Project Context for Claude Code

## What this project does
Reads stock tickers from `tickers.txt`, fetches live market data via yfinance (no API key needed),
and generates a multi-sheet Excel workbook with portfolio analysis and Claude AI insights.

## Key files
- `stock_analyzer.py` — single-file main script (~450 lines)
- `tickers.txt` — input file, format: `TICKER[,SHARES[,BUY_DATE]]  [# inline comment]`
- `.env` — contains `ANTHROPIC_API_KEY` (never commit this file)
- `.env.example` — template to copy from
- `requirements.txt` — 6 dependencies (yfinance, openpyxl, pandas, anthropic, python-dotenv, tqdm)

## How to run
```bash
python3 stock_analyzer.py                        # uses tickers.txt, auto-timestamped output
python3 stock_analyzer.py my_list.txt            # custom tickers file
python3 stock_analyzer.py --output report.xlsx   # fixed output filename
python3 stock_analyzer.py --no-ai                # skip AI sheet even if key is set
python3 stock_analyzer.py --open                 # open file after generation (macOS/Linux/Win)
python3 stock_analyzer.py --watch                # re-run on tickers file changes
python3 stock_analyzer.py --model claude-opus-4-6  # use a different Claude model
python3 stock_analyzer.py --add "TSLA,10,2024-03-01"  # append ticker and exit
```
Output: `portfolio_YYYYMMDD_HHMM.xlsx` — excluded from git via `.gitignore`.

## Python version
System Python 3.9.6 on macOS. Do NOT use `list[str]` or `dict[str, int]` type annotation
subscripts (not supported in 3.9) — use bare `list` or `typing.List` / `typing.Dict` instead.
Also avoid walrus operator (`:=`) — not well supported in 3.9 edge cases.

## Excel sheets generated
1. **Portfolio** — columns A–AE with data + TOTAL row + currency warning + portfolio stats
2. **Distribution** — country/sector breakdown with SUMIF/COUNTIF formulas
3. **Charts** — sector pie chart + P&L % bar chart (auto-generated with openpyxl.chart)
4. **Errors** — (only if any tickers failed) lists failed tickers with error messages
5. **AI Analysis** — (only if ANTHROPIC_API_KEY set and --no-ai not passed) Claude analysis

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
- Freeze panes at `E4` — keeps A–D visible when scrolling right, and rows 1–3 (title/headers) visible when scrolling down
- Distribution sheet SUMIF/COUNTIF formulas reference: Country=V, Sector=W, Total Value=H

## Portfolio Statistics section (below TOTAL row)
Computed in Python using market-cap weighting across all stocks that have `market_cap` data:
- Weighted Avg Beta
- Weighted Avg P/E (TTM)
- Weighted Avg Div Yield
Appears a couple of rows below TOTAL. If multiple currencies are present, an orange warning
banner appears between TOTAL and the stats section.

## Style constants
All colours defined at the top of `stock_analyzer.py`:
`DARK_BLUE`, `MED_BLUE`, `LIGHT_BLUE`, `ALT_ROW`, `YELLOW_FILL`, `GREEN_FILL`, `ORANGE_FILL`, etc.

## AI Analysis sheet
- Called via `anthropic.Anthropic(api_key=...).messages.stream(model=model, ...)`
- Streams tokens to terminal as they arrive (live output during the 15–30 s wait)
- Requires `ANTHROPIC_API_KEY` in `.env` with credits on console.anthropic.com
- Skipped gracefully if key is absent, invalid, or `--no-ai` flag is passed
- `--model` CLI flag controls which Claude model is used (default: `claude-sonnet-4-6`)
- Prompt sends full portfolio JSON; expects structured Markdown with 6 sections

## tqdm progress bar
- Shown during data fetching if `tqdm` is installed (in requirements.txt)
- Gracefully degrades to plain print output if tqdm is missing
- Use `log()` helper (not `print()`) inside fetch_stock for tqdm-compatible output

## Common tasks
- **Add a ticker interactively**: `python3 stock_analyzer.py --add "TICKER,SHARES,BUY_DATE"`
- **Add a ticker manually**: edit `tickers.txt`, one line per ticker; inline `# comments` supported
- **Add a column**: add a tuple to `COLUMNS`, handle the cell value in `create_portfolio_sheet()`
- **Change AI model**: use `--model` flag, or update `default=` in the argparse definition
- **Fix date display issues**: always use lowercase `yyyy-mm-dd` in `number_format`

## Dependencies
```
yfinance>=0.2.50    # market data
openpyxl>=3.1.0     # Excel generation + charts
pandas>=2.0.0       # data manipulation
anthropic>=0.40.0   # Claude API
python-dotenv>=1.0.0 # .env loading
tqdm>=4.60.0        # progress bar (optional — degrades gracefully if missing)
```
