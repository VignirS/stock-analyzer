# Setup Guide

## 1. Install dependencies

```bash
pip install -r requirements.txt
```

## 2. Configure the API key (optional)

The AI Analysis sheet requires an Anthropic API key. The script runs fine without it —
Portfolio and Distribution sheets are always generated.

```bash
cp .env.example .env
```

Edit `.env` and add your key:

```
ANTHROPIC_API_KEY=sk-ant-...
```

**Get a key:** [console.anthropic.com/settings/keys](https://console.anthropic.com/settings/keys)
**Add credits:** [console.anthropic.com/settings/billing](https://console.anthropic.com/settings/billing)

> Note: Claude Code (claude.ai subscription) and the Anthropic API use separate billing.
> You need API credits on console.anthropic.com to use the AI Analysis sheet.

## 3. Edit your tickers

Open `tickers.txt` and add your holdings:

```
# TICKER[,SHARES[,BUY_DATE]]
AAPL,50,2022-06-15
MSFT,20
NESN.SW
```

- **TICKER** — required, any Yahoo Finance symbol
- **SHARES** — optional, pre-fills column G in the Excel file
- **BUY_DATE** — optional (`YYYY-MM-DD`), enables buy price lookup and P&L columns

## 4. Run

```bash
python3 stock_analyzer.py
```

Opens a file named `portfolio_YYYYMMDD_HHMM.xlsx` in the project directory.

## 5. Use the Excel file

1. Open the file in Excel or LibreOffice Calc
2. Fill in **column G** (# Shares, highlighted yellow) for any positions without share counts
3. **Total Value**, **Cost Basis**, **P&L $**, and the **Distribution** sheet update automatically

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `No history data returned` for a ticker | Check the symbol on finance.yahoo.com |
| Buy date shows `yyyy-03-dd` | Ensure openpyxl >= 3.1 and Python >= 3.8 |
| AI Analysis sheet missing | Check `ANTHROPIC_API_KEY` in `.env` and API credits |
| `401 invalid x-api-key` | Key was revoked — generate a new one at console.anthropic.com |
| `credit balance too low` | Add credits at console.anthropic.com → Plans & Billing |
