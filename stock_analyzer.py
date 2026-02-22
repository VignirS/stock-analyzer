#!/usr/bin/env python3
"""
Stock Portfolio Analyzer
========================
Reads stock tickers from a text file, fetches market data via yfinance,
generates an Excel workbook with portfolio analysis, and optionally
calls the Claude API for AI-powered insights.

Tickers file format: TICKER[,SHARES[,BUY_DATE]]
  BUY_DATE format: YYYY-MM-DD  (enables P&L columns when provided)

Usage:
    python stock_analyzer.py [tickers_file]

Default tickers file: tickers.txt
Output: portfolio_YYYYMMDD_HHMM.xlsx in the same directory as the tickers file.
"""

import sys
import os
import json
import argparse
from datetime import datetime, date as date_type
from pathlib import Path
from typing import Optional

import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import CellIsRule
from dotenv import load_dotenv

load_dotenv()

# ─── Style constants ──────────────────────────────────────────────────────────
DARK_BLUE     = "1F3864"
MED_BLUE      = "2E75B6"
LIGHT_BLUE    = "D9E1F2"
ALT_ROW       = "F2F2F2"
YELLOW_FILL   = "FFFF00"
GREEN_FILL    = "E2EFDA"
GREEN_TOTAL   = "375623"
RED_CF        = "FFC7CE"
RED_CF_FONT   = "9C0006"
GREEN_CF      = "C6EFCE"
GREEN_CF_FONT = "276221"

# ─── Column definitions ───────────────────────────────────────────────────────
# Tuples: (column_letter, header_label, stock_data_key_or_None, number_format, column_width)
#
# Layout:
#   A–D   : Identity  (Ticker, Name, Price, Currency)
#   E–K   : Position  (Buy Date, Buy Price, # Shares, Total Value, Cost Basis, P&L $, P&L %)
#   L–U   : Returns   (1W … 5Y %)
#   V–AE  : Fundamentals (Country, Sector, Industry, Mkt Cap, P/E, 52W Hi/Lo, drawdown, Div, Beta)

COLUMNS = [
    # ── Identity ──
    ("A",  "Ticker",           "ticker",         "@",              10),
    ("B",  "Company Name",     "name",           "@",              28),
    ("C",  "Current Price",    "current_price",  "#,##0.00",       13),
    ("D",  "Currency",         "currency",       "@",               8),
    # ── Position / P&L ──
    ("E",  "Buy Date",         "buy_date",       "yyyy-mm-dd",     12),   # from tickers file
    ("F",  "Buy Price",        "buy_price",      "#,##0.00",       12),   # fetched from history
    ("G",  "# Shares",         None,             "#,##0.####",     10),   # user input (yellow)
    ("H",  "Total Value",      None,             "#,##0.00",       15),   # formula (green)
    ("I",  "Cost Basis",       None,             "#,##0.00",       14),   # formula (green)
    ("J",  "P&L $",            None,             "#,##0.00",       13),   # formula + CF
    ("K",  "P&L %",            None,             "0.00%",          10),   # formula + CF
    # ── Returns ──
    ("L",  "1W%",              "1W%",            "0.00%",           8),
    ("M",  "1M%",              "1M%",            "0.00%",           8),
    ("N",  "3M%",              "3M%",            "0.00%",           8),
    ("O",  "6M%",              "6M%",            "0.00%",           8),
    ("P",  "YTD%",             "YTD%",           "0.00%",           8),
    ("Q",  "1Y%",              "1Y%",            "0.00%",           8),
    ("R",  "2Y%",              "2Y%",            "0.00%",           8),
    ("S",  "3Y%",              "3Y%",            "0.00%",           8),
    ("T",  "4Y%",              "4Y%",            "0.00%",           8),
    ("U",  "5Y%",              "5Y%",            "0.00%",           8),
    # ── Fundamentals ──
    ("V",  "Country",          "country",        "@",              14),
    ("W",  "Sector",           "sector",         "@",              18),
    ("X",  "Industry",         "industry",       "@",              22),
    ("Y",  "Market Cap",       "market_cap",     '#,##0,,"M"',     14),
    ("Z",  "P/E (TTM)",        "pe_ratio",       "0.00",           10),
    ("AA", "52W High",         "52w_high",       "#,##0.00",       12),
    ("AB", "52W Low",          "52w_low",        "#,##0.00",       12),
    ("AC", "% from 52W High",  None,             "0.00%",          15),   # formula
    ("AD", "Div. Yield %",     "dividend_yield", "0.00%",          12),
    ("AE", "Beta",             "beta",           "0.00",            8),
]

# Returns columns L–U (indices 11–20)
PCT_COLS = COLUMNS[11:21]

# Last column letter (for title merge / autofilter)
LAST_COL = COLUMNS[-1][0]   # "AE"


# ─── Style helpers ────────────────────────────────────────────────────────────

def mk_fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def mk_font(bold: bool = False, color: str = "000000", size: int = 10,
            italic: bool = False) -> Font:
    return Font(bold=bold, italic=italic, color=color, name="Calibri", size=size)


def mk_center() -> Alignment:
    return Alignment(horizontal="center", vertical="center")


def mk_vcenter() -> Alignment:
    return Alignment(vertical="center")


def mk_wrap() -> Alignment:
    return Alignment(wrap_text=True, vertical="top")


# ─── Data fetching ────────────────────────────────────────────────────────────

def read_tickers(filepath: str) -> list:
    """Parse tickers file. Returns list of dicts: {ticker, shares, buy_date}."""
    entries = []
    with open(filepath, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in line.split(",")]
            ticker = parts[0].upper()
            if not ticker:
                continue

            shares = None
            if len(parts) > 1 and parts[1]:
                try:
                    shares = float(parts[1])
                except ValueError:
                    print(f"  Warning: invalid share count '{parts[1]}' for {ticker}, ignoring.")

            buy_date = None
            if len(parts) > 2 and parts[2]:
                try:
                    datetime.strptime(parts[2], "%Y-%m-%d")
                    buy_date = parts[2]
                except ValueError:
                    print(f"  Warning: invalid date '{parts[2]}' for {ticker}, ignoring.")

            entries.append({"ticker": ticker, "shares": shares, "buy_date": buy_date})
    return entries


def normalize_history(hist: pd.DataFrame) -> pd.DataFrame:
    """Strip timezone so that date arithmetic works simply."""
    if hist.empty:
        return hist
    if getattr(hist.index, "tz", None) is not None:
        hist = hist.copy()
        hist.index = hist.index.tz_convert("UTC").tz_localize(None)
    return hist


def pct_change(hist: pd.DataFrame, days: int) -> Optional[float]:
    """Percentage change from `days` calendar days ago to the latest close."""
    if hist.empty:
        return None
    latest_price = float(hist["Close"].iloc[-1])
    target_date  = hist.index[-1] - pd.Timedelta(days=days)
    past = hist[hist.index <= target_date]
    if past.empty:
        return None
    past_price = float(past["Close"].iloc[-1])
    if not past_price or pd.isna(past_price) or past_price == 0:
        return None
    return (latest_price - past_price) / past_price


def ytd_change(hist: pd.DataFrame) -> Optional[float]:
    """Percentage change from last trading day of the previous year."""
    if hist.empty:
        return None
    latest_price  = float(hist["Close"].iloc[-1])
    prev_year_end = pd.Timestamp(hist.index[-1].year - 1, 12, 31)
    past = hist[hist.index <= prev_year_end]
    if past.empty:
        return None
    past_price = float(past["Close"].iloc[-1])
    if not past_price or pd.isna(past_price) or past_price == 0:
        return None
    return (latest_price - past_price) / past_price


def fetch_stock(entry: dict) -> dict:
    """Fetch all data for one ticker entry. Returns dict with 'error' key on failure."""
    symbol   = entry["ticker"]
    buy_date = entry.get("buy_date")

    print(f"  Fetching {symbol}...")
    result: dict = {"ticker": symbol, "input_shares": entry.get("shares"), "buy_date": buy_date}

    try:
        ticker   = yf.Ticker(symbol)
        hist_raw = ticker.history(period="6y")
        try:
            info = ticker.info or {}
        except Exception:
            info = {}
    except Exception as exc:
        print(f"    ✗ {symbol}: {exc}")
        result["error"] = str(exc)
        return result

    hist = normalize_history(hist_raw)
    if hist.empty:
        print(f"    ✗ {symbol}: no history returned")
        result["error"] = "No history data returned"
        return result

    current_price = float(hist["Close"].iloc[-1])

    # Percentage changes
    timeframes = {
        "1W%": 7, "1M%": 30, "3M%": 91,
        "6M%": 182, "1Y%": 365, "2Y%": 730,
        "3Y%": 1095, "4Y%": 1460, "5Y%": 1825,
    }
    for key, days in timeframes.items():
        result[key] = pct_change(hist, days)
    result["YTD%"] = ytd_change(hist)

    # 52-week high/low (trailing 365 calendar days)
    cutoff  = hist.index[-1] - pd.Timedelta(days=365)
    hist_1y = hist[hist.index >= cutoff]
    result["52w_high"] = float(hist_1y["High"].max()) if not hist_1y.empty else None
    result["52w_low"]  = float(hist_1y["Low"].min())  if not hist_1y.empty else None

    # Buy price — closing price on or immediately after buy date
    result["buy_price"] = None
    if buy_date:
        buy_ts      = pd.Timestamp(buy_date)
        on_or_after = hist[hist.index >= buy_ts]
        if not on_or_after.empty:
            result["buy_price"] = float(on_or_after["Close"].iloc[0])
            actual = on_or_after.index[0].strftime("%Y-%m-%d")
            if actual != buy_date:
                print(f"    ℹ {symbol}: {buy_date} not a trading day, using {actual}")
        else:
            print(f"    ✗ {symbol}: buy date {buy_date} is beyond available history")

    # Info fields
    result.update({
        "name":           info.get("longName") or info.get("shortName") or symbol,
        "currency":       info.get("currency", "USD"),
        "country":        info.get("country")  or "",
        "sector":         info.get("sector")   or "",
        "industry":       info.get("industry") or "",
        "market_cap":     info.get("marketCap"),
        "pe_ratio":       info.get("trailingPE"),
        "beta":           info.get("beta"),
        "dividend_yield": info.get("dividendYield"),
        "current_price":  current_price,
    })

    bp_str = f"  buy@{result['buy_price']:,.2f}" if result["buy_price"] else ""
    print(f"    ✓ {symbol}: {result['name']}  {current_price:,.2f} {result['currency']}{bp_str}")
    return result


# ─── Sheet: Portfolio ─────────────────────────────────────────────────────────

def create_portfolio_sheet(wb: Workbook, stocks: list) -> None:
    ws = wb.active
    ws.title = "Portfolio"

    # ── Row 1: title bar ──
    ws.merge_cells(f"A1:{LAST_COL}1")
    c = ws["A1"]
    c.value     = f"Stock Portfolio  ·  Generated {datetime.now():%Y-%m-%d %H:%M}"
    c.font      = mk_font(bold=True, color="FFFFFF", size=13)
    c.fill      = mk_fill(DARK_BLUE)
    c.alignment = mk_center()
    ws.row_dimensions[1].height = 26

    # ── Row 2: thin spacer ──
    ws.row_dimensions[2].height = 4

    # ── Row 3: column headers ──
    for col_letter, header, *_ in COLUMNS:
        c = ws[f"{col_letter}3"]
        c.value     = header
        c.font      = mk_font(bold=True, color="FFFFFF", size=10)
        c.fill      = mk_fill(MED_BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 32

    # ── Data rows (starting row 4) ──
    n = len(stocks)
    for i, stock in enumerate(stocks):
        row       = i + 4
        row_fill  = mk_fill(ALT_ROW) if i % 2 else None
        base_font = mk_font()

        def _c(col: str):
            """Return cell with alternating fill and base font applied."""
            cell = ws[f"{col}{row}"]
            if row_fill:
                cell.fill = row_fill
            cell.font = base_font
            return cell

        # ── Identity ──
        c = _c("A");  c.value = stock["ticker"];              c.alignment = mk_center()
        c = _c("B");  c.value = stock.get("name", "");        c.alignment = mk_vcenter()
        c = _c("C");  c.value = stock.get("current_price");   c.number_format = "#,##0.00";   c.alignment = mk_center()
        c = _c("D");  c.value = stock.get("currency", "");    c.alignment = mk_center()

        # ── E: Buy Date ──
        c = ws[f"E{row}"]
        if row_fill: c.fill = row_fill
        c.font = base_font;  c.alignment = mk_center()
        if stock.get("buy_date"):
            try:
                c.value = datetime.strptime(stock["buy_date"], "%Y-%m-%d").date()
                c.number_format = "yyyy-mm-dd"
            except ValueError:
                c.value = stock["buy_date"]

        # ── F: Buy Price ──
        c = ws[f"F{row}"]
        if row_fill: c.fill = row_fill
        c.font = base_font;  c.alignment = mk_center()
        if stock.get("buy_price") is not None:
            c.value = stock["buy_price"];  c.number_format = "#,##0.00"

        # ── G: # Shares (user input, yellow; pre-filled if in tickers file) ──
        c = ws[f"G{row}"]
        c.fill          = mk_fill(YELLOW_FILL)
        c.font          = base_font
        c.alignment     = mk_center()
        c.number_format = "#,##0.####"
        if stock.get("input_shares") is not None:
            c.value = stock["input_shares"]

        # ── H: Total Value (formula, green) ──
        c = ws[f"H{row}"]
        c.value         = f'=IF(OR(G{row}="",G{row}=0),"",C{row}*G{row})'
        c.number_format = "#,##0.00"
        c.fill          = mk_fill(GREEN_FILL)
        c.font          = base_font;  c.alignment = mk_center()

        # ── I: Cost Basis (formula, green) ──
        c = ws[f"I{row}"]
        c.value         = f'=IF(OR(F{row}="",G{row}="",G{row}=0),"",F{row}*G{row})'
        c.number_format = "#,##0.00"
        c.fill          = mk_fill(GREEN_FILL)
        c.font          = base_font;  c.alignment = mk_center()

        # ── J: P&L $ (formula) ──
        c = ws[f"J{row}"]
        c.value         = f'=IF(OR(H{row}="",I{row}=""),"",H{row}-I{row})'
        c.number_format = "#,##0.00"
        if row_fill: c.fill = row_fill
        c.font = base_font;  c.alignment = mk_center()

        # ── K: P&L % (formula) ──
        c = ws[f"K{row}"]
        c.value         = f'=IF(OR(F{row}="",F{row}=0),"",(C{row}-F{row})/F{row})'
        c.number_format = "0.00%"
        if row_fill: c.fill = row_fill
        c.font = base_font;  c.alignment = mk_center()

        # ── L–U: % changes ──
        for col_letter, _, key, fmt, _ in PCT_COLS:
            c = _c(col_letter)
            c.value = stock.get(key);  c.number_format = fmt;  c.alignment = mk_center()

        # ── Fundamentals ──
        c = _c("V");  c.value = stock.get("country", "");   c.alignment = mk_vcenter()
        c = _c("W");  c.value = stock.get("sector", "");    c.alignment = mk_vcenter()
        c = _c("X");  c.value = stock.get("industry", "");  c.alignment = mk_vcenter()
        c = _c("Y");  c.value = stock.get("market_cap");    c.number_format = '#,##0,,"M"'; c.alignment = mk_center()
        c = _c("Z");  c.value = stock.get("pe_ratio");      c.number_format = "0.00";       c.alignment = mk_center()
        c = _c("AA"); c.value = stock.get("52w_high");      c.number_format = "#,##0.00";   c.alignment = mk_center()
        c = _c("AB"); c.value = stock.get("52w_low");       c.number_format = "#,##0.00";   c.alignment = mk_center()

        # ── AC: % from 52W High (formula) ──
        c = ws[f"AC{row}"]
        c.value         = f'=IF(OR(AA{row}="",AA{row}=0),"",(C{row}-AA{row})/AA{row})'
        c.number_format = "0.00%"
        if row_fill: c.fill = row_fill
        c.font = base_font;  c.alignment = mk_center()

        c = _c("AD"); c.value = stock.get("dividend_yield"); c.number_format = "0.00%"; c.alignment = mk_center()
        c = _c("AE"); c.value = stock.get("beta");           c.number_format = "0.00";  c.alignment = mk_center()

    # ── TOTAL row ──
    total_row = n + 4

    c = ws[f"A{total_row}"]
    c.value = "TOTAL";  c.font = mk_font(bold=True, color="FFFFFF", size=11)
    c.fill = mk_fill(DARK_BLUE);  c.alignment = mk_center()

    for col in ("H", "I", "J"):
        c = ws[f"{col}{total_row}"]
        c.value         = f"=SUM({col}4:{col}{total_row - 1})"
        c.number_format = "#,##0.00"
        c.font          = mk_font(bold=True, color="FFFFFF", size=11)
        c.fill          = mk_fill(GREEN_TOTAL)
        c.alignment     = mk_center()

    # ── Conditional formatting ──
    data_end = total_row - 1
    red_fill_cf   = PatternFill(start_color=RED_CF,   end_color=RED_CF,   fill_type="solid")
    green_fill_cf = PatternFill(start_color=GREEN_CF, end_color=GREEN_CF, fill_type="solid")

    for cf_range in [f"J4:K{data_end}", f"L4:U{data_end}"]:
        ws.conditional_formatting.add(cf_range, CellIsRule(
            operator="lessThan",    formula=["0"], fill=red_fill_cf,   font=Font(color=RED_CF_FONT)))
        ws.conditional_formatting.add(cf_range, CellIsRule(
            operator="greaterThan", formula=["0"], fill=green_fill_cf, font=Font(color=GREEN_CF_FONT)))

    # ── Freeze panes: keep A–D visible when scrolling right ──
    ws.freeze_panes = "E4"

    # ── AutoFilter ──
    ws.auto_filter.ref = f"A3:{LAST_COL}{data_end}"

    # ── Column widths ──
    for col_letter, _, _, _, width in COLUMNS:
        ws.column_dimensions[col_letter].width = width


# ─── Sheet: Distribution ──────────────────────────────────────────────────────

def create_distribution_sheet(wb: Workbook, stocks: list) -> None:
    ws = wb.create_sheet("Distribution")
    n = len(stocks)

    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value = "Portfolio Distribution";  c.font = mk_font(bold=True, color="FFFFFF", size=13)
    c.fill = mk_fill(DARK_BLUE);  c.alignment = mk_center()
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 4

    for col, label in [
        ("A", "Country"), ("B", "# Stocks"), ("C", "% Count"), ("D", "Total Value"),
        ("F", "Sector"),  ("G", "# Stocks"), ("H", "% Count"), ("I", "Total Value"),
    ]:
        c = ws[f"{col}3"]
        c.value = label;  c.font = mk_font(bold=True, color="FFFFFF", size=10)
        c.fill = mk_fill(MED_BLUE);  c.alignment = mk_center()
    ws.row_dimensions[3].height = 28
    ws.row_dimensions[4].height = 4

    countries = sorted({s.get("country", "") for s in stocks if s.get("country")})
    sectors   = sorted({s.get("sector",  "") for s in stocks if s.get("sector")})

    # Country is now column V; Total Value is column H
    for i, country in enumerate(countries):
        row = i + 5;  rf = mk_fill(ALT_ROW) if i % 2 else None
        c = ws[f"A{row}"];  c.value = country;  c.font = mk_font()
        if rf: c.fill = rf
        for col, formula, fmt in [
            ("B", f"=COUNTIF(Portfolio!V:V,A{row})",               "0"),
            ("C", f'=IF(B{row}=0,"",B{row}/{n})',                  "0.00%"),
            ("D", f"=SUMIF(Portfolio!V:V,A{row},Portfolio!H:H)",   "#,##0.00"),
        ]:
            c = ws[f"{col}{row}"];  c.value = formula;  c.number_format = fmt
            c.alignment = mk_center();  c.font = mk_font()
            if rf: c.fill = rf

    # Sector is now column W; Total Value is column H
    for i, sector in enumerate(sectors):
        row = i + 5;  rf = mk_fill(ALT_ROW) if i % 2 else None
        c = ws[f"F{row}"];  c.value = sector;  c.font = mk_font()
        if rf: c.fill = rf
        for col, formula, fmt in [
            ("G", f"=COUNTIF(Portfolio!W:W,F{row})",               "0"),
            ("H", f'=IF(G{row}=0,"",G{row}/{n})',                  "0.00%"),
            ("I", f"=SUMIF(Portfolio!W:W,F{row},Portfolio!H:H)",   "#,##0.00"),
        ]:
            c = ws[f"{col}{row}"];  c.value = formula;  c.number_format = fmt
            c.alignment = mk_center();  c.font = mk_font()
            if rf: c.fill = rf

    for col, width in [
        ("A", 18), ("B", 10), ("C", 10), ("D", 14),
        ("E",  3),
        ("F", 22), ("G", 10), ("H", 10), ("I", 14),
    ]:
        ws.column_dimensions[col].width = width


# ─── Sheet: AI Analysis ───────────────────────────────────────────────────────

def create_ai_sheet(wb: Workbook, stocks: list, api_key: str) -> None:
    import anthropic

    ws = wb.create_sheet("AI Analysis")
    ws.column_dimensions["A"].width = 110

    c = ws["A1"]
    c.value     = "AI Portfolio Analysis  ·  Powered by Claude"
    c.font      = mk_font(bold=True, color="FFFFFF", size=13)
    c.fill      = mk_fill(DARK_BLUE)
    c.alignment = mk_center()
    ws.row_dimensions[1].height = 26

    def safe_pct(val):
        return round(val * 100, 2) if val is not None else None

    portfolio_data = [
        {
            "ticker":        s.get("ticker"),
            "name":          s.get("name"),
            "sector":        s.get("sector"),
            "country":       s.get("country"),
            "currency":      s.get("currency"),
            "price":         round(s["current_price"], 2) if s.get("current_price") else None,
            "market_cap_M":  round(s["market_cap"] / 1e6) if s.get("market_cap") else None,
            "pe_ratio":      round(s["pe_ratio"], 2) if s.get("pe_ratio") else None,
            "beta":          round(s["beta"], 2) if s.get("beta") else None,
            "div_yield_pct": safe_pct(s.get("dividend_yield")),
            "ytd_pct":       safe_pct(s.get("YTD%")),
            "1y_pct":        safe_pct(s.get("1Y%")),
            "3y_pct":        safe_pct(s.get("3Y%")),
            "5y_pct":        safe_pct(s.get("5Y%")),
            "buy_date":      s.get("buy_date"),
            "buy_price":     round(s["buy_price"], 2) if s.get("buy_price") else None,
            "pnl_pct":       (
                safe_pct((s["current_price"] - s["buy_price"]) / s["buy_price"])
                if s.get("buy_price") and s.get("current_price") else None
            ),
            "pct_from_52w_high": (
                round((s["current_price"] - s["52w_high"]) / s["52w_high"] * 100, 2)
                if s.get("current_price") and s.get("52w_high") else None
            ),
        }
        for s in stocks
    ]

    prompt = (
        "You are a professional financial analyst. Analyse the following stock portfolio "
        "and provide a structured, data-driven report in Markdown.\n\n"
        f"Portfolio data (JSON):\n{json.dumps(portfolio_data, indent=2)}\n\n"
        "Structure your response with these exact section headers:\n\n"
        "## Portfolio Overview\n"
        "Brief summary: number of stocks, currencies, geographic spread, sector mix.\n\n"
        "## Geographic & Sector Diversification\n"
        "Assess diversification. Flag any single geography or sector exceeding ~40%.\n\n"
        "## Performance Highlights\n"
        "Notable outperformers and underperformers by YTD, 1Y, and 5Y. "
        "Where buy dates are available, include unrealised P&L context.\n\n"
        "## Risk Assessment\n"
        "Beta analysis, valuation concerns (elevated P/E), drawdown risks, "
        "and any stocks significantly below their 52-week high.\n\n"
        "## Sector Concentration Warnings\n"
        "Flag any sector representing more than 35% of identifiable holdings.\n\n"
        "## Key Takeaways\n"
        "3–5 concise bullet points of the most important insights.\n\n"
        "Keep the tone professional. Reference specific tickers and numbers. "
        "Append a disclaimer that this is not financial advice."
    )

    print("  Calling Claude API (claude-sonnet-4-6)...")
    try:
        client_ai = anthropic.Anthropic(api_key=api_key)
        with client_ai.messages.stream(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        ) as stream:
            final    = stream.get_final_message()
        analysis = final.content[0].text
        print("  ✓ Analysis complete.")
    except Exception as exc:
        analysis = f"Error generating AI analysis: {exc}\n\nCheck that ANTHROPIC_API_KEY is valid."
        print(f"  ✗ Claude API error: {exc}")

    row = 3
    for line in analysis.splitlines():
        stripped = line.strip()
        c = ws.cell(row=row, column=1)
        c.alignment = mk_wrap()
        if stripped.startswith("## "):
            c.value = stripped[3:];  c.font = mk_font(bold=True, color=DARK_BLUE, size=12)
            c.fill = mk_fill(LIGHT_BLUE);  ws.row_dimensions[row].height = 20
        elif stripped.startswith("# "):
            c.value = stripped[2:];  c.font = mk_font(bold=True, color="FFFFFF", size=13)
            c.fill = mk_fill(DARK_BLUE);   ws.row_dimensions[row].height = 22
        elif stripped == "":
            ws.row_dimensions[row].height = 8
        else:
            c.value = line;  c.font = mk_font(size=10);  ws.row_dimensions[row].height = 15
        row += 1

    row += 1
    c = ws.cell(row=row, column=1)
    c.value = (f"Generated by Claude (claude-sonnet-4-6) · {datetime.now():%Y-%m-%d %H:%M} · "
               "For informational purposes only — not financial advice.")
    c.font = mk_font(italic=True, color="888888", size=9)


# ─── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Stock Portfolio Analyzer — yfinance + openpyxl + Claude"
    )
    parser.add_argument(
        "tickers_file", nargs="?", default="tickers.txt",
        help="Path to tickers file (default: tickers.txt)",
    )
    args = parser.parse_args()

    tickers_path = Path(args.tickers_file)
    if not tickers_path.exists():
        print(f"Error: '{tickers_path}' not found.")
        sys.exit(1)

    print(f"Reading tickers from '{tickers_path}'...")
    entries = read_tickers(str(tickers_path))
    if not entries:
        print("No tickers found. Check your tickers file.")
        sys.exit(1)

    print(f"Found {len(entries)} ticker(s): {', '.join(e['ticker'] for e in entries)}\n")

    print("Fetching market data from Yahoo Finance...")
    all_data = [fetch_stock(e) for e in entries]
    stocks   = [s for s in all_data if "error" not in s]
    failed   = [s for s in all_data if "error" in s]

    if failed:
        print(f"\nWarning: failed to fetch {len(failed)} ticker(s): "
              f"{', '.join(s['ticker'] for s in failed)}")
    if not stocks:
        print("No valid stock data retrieved. Exiting.")
        sys.exit(1)

    has_pnl = sum(1 for s in stocks if s.get("buy_price"))
    print(f"\nSuccessfully fetched {len(stocks)} / {len(entries)} stock(s)."
          + (f"  {has_pnl} with P&L data." if has_pnl else ""))

    print("\nBuilding Excel workbook...")
    wb = Workbook()

    create_portfolio_sheet(wb, stocks)
    print("  ✓ Portfolio sheet")

    create_distribution_sheet(wb, stocks)
    print("  ✓ Distribution sheet")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if api_key:
        print("\nGenerating AI analysis (this may take 15–30 s)...")
        create_ai_sheet(wb, stocks, api_key)
        print("  ✓ AI Analysis sheet")
    else:
        print(
            "\n  Note: ANTHROPIC_API_KEY not set — AI Analysis sheet skipped.\n"
            "  Copy .env.example to .env, add your key, and re-run to enable it."
        )

    out_dir = tickers_path.parent.resolve()
    output  = out_dir / f"portfolio_{datetime.now():%Y%m%d_%H%M}.xlsx"
    wb.save(output)

    print(f"\n{'=' * 60}")
    print(f"  Saved: {output}")
    print(f"{'=' * 60}")
    print("\nNext steps:")
    print("  1. Open the file in Excel / LibreOffice Calc")
    print("  2. Adjust '# Shares' (column G, yellow) if needed")
    print("  3. P&L columns (J, K) and Distribution sheet update automatically")


if __name__ == "__main__":
    main()
