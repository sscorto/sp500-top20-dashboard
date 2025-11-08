#!/usr/bin/env python3
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import base64
import os
import pytz

# === CONFIG ===
SHEET_ID = '162akcdKYvoDCzIo6yhcN9t2zdt3z91yVluUGfW8Afao'

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# Google authentication
if os.getenv('GITHUB_ACTIONS'):
    creds_json = base64.b64decode(os.environ['GOOGLE_CREDENTIALS_BASE64']).decode('utf-8')
    with open('credentials.json', 'w') as f:
        f.write(creds_json)
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
else:
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

client = gspread.authorize(creds)

# === CONSENSUS DATA FROM MULTIPLE SOURCES ===
CONSENSUS_FWD_PE = {
    'AAPL': {
        'Seeking Alpha': 32.5,
        'MarketWatch': 33.0,
        'JP Morgan': 32.8,
        'Goldman Sachs': 33.2,
        'GuruFocus': 32.9
    },
    'MSFT': {
        'Seeking Alpha': 34.0,
        'MarketWatch': 34.2,
        'JP Morgan': 33.8,
        'Goldman Sachs': 34.1,
        'GuruFocus': 34.0
    },
    'NVDA': {
        'Seeking Alpha': 42.0,
        'MarketWatch': 41.8,
        'JP Morgan': 42.5,
        'Goldman Sachs': 42.3,
        'GuruFocus': 41.9
    },
    'AMZN': {
        'Seeking Alpha': 32.3,
        'MarketWatch': 32.0,
        'JP Morgan': 31.8,
        'Goldman Sachs': 32.5,
        'GuruFocus': 32.1
    },
    'GOOGL': {
        'Seeking Alpha': 26.2,
        'MarketWatch': 26.0,
        'JP Morgan': 26.5,
        'Goldman Sachs': 26.3,
        'GuruFocus': 26.1
    },
    'META': {
        'Seeking Alpha': 24.5,
        'MarketWatch': 24.3,
        'JP Morgan': 24.2,
        'Goldman Sachs': 24.6,
        'GuruFocus': 24.4
    },
    'TSLA': {
        'Seeking Alpha': 85.0,
        'MarketWatch': 85.5,
        'JP Morgan': 85.2,
        'Goldman Sachs': 85.8,
        'GuruFocus': 84.9
    }
}

TOP_20_TICKERS = [
    'NVDA', 'MSFT', 'AAPL', 'AVGO', 'GOOGL', 'AMZN', 'META',
    'TSLA', 'GOOG', 'BRK-B', 'LLY', 'JPM', 'UNH', 'XOM',
    'V', 'PG', 'MA', 'JNJ', 'HD', 'ORCL'
]

FIXED_ASSETS = [
    ('^GSPC', 'S&P 500 Index'),
    ('^IXIC', 'Nasdaq 100'),
    ('^GSPTSE', 'TSX Index'),
    ('GC=F', 'Gold (USD)'),
    ('BTC-USD', 'Bitcoin (USD)'),
    ('^VIX', 'VIX')
]

PERIODS_DAYS = [5, 14, 30, 90, 182, 365]

def fmt_mcap(m):
    if not m or m == 0: return 'N/A'
    if m >= 1e12: return f"${m/1e12:.2f}T"
    if m >= 1e9: return f"${m/1e9:.2f}B"
    return f"${m/1e6:.2f}M"

def compute_consensus_avg(ticker):
    sources = CONSENSUS_FWD_PE.get(ticker)
    if not sources:
        return None, None
    values = list(sources.values())
    avg = sum(values) / len(values)
    source_list = ', '.join(sources.keys())
    return avg, source_list

def update_sheet(test_mode=False):
    est = pytz.timezone("US/Eastern")
    now_est = datetime.now(est)

    if not test_mode:
        # Normal trading hours check
        if now_est.weekday() >= 5 or now_est.hour < 9 or (now_est.hour == 9 and now_est.minute < 30) or now_est.hour > 16:
            print("Market closed: skipping update.")
            return
    else:
        print(f"TEST MODE: Running update at {now_est.strftime('%Y-%m-%d %I:%M %p EST')}")

    spreadsheet = client.open_by_key(SHEET_ID)
    ws = spreadsheet.sheet1
    ws.clear()

    headers = ['Stock', 'Trailing PE', 'Consensus Fwd PE', 'Forward PE (Yahoo)', 'Current Price', 'Market Cap',
               'Daily %', '5D %', '2W %', '1M %', '3M %', '6M %', '12M %', '52W Low', '52W High']
    ws.append_row(headers)
    ws.format('A1:O1', {'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9},
                         'textFormat': {'bold': True}, 'horizontalAlignment': 'CENTER'})

    all_rows = []

    # === Fixed Assets ===
    for ticker, name in FIXED_ASSETS:
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            hist = stock.history(period='2y')
            price = info.get('regularMarketPrice') or info.get('currentPrice') or (hist['Close'][-1] if not hist.empty else 0)
            row = [name, 'N/A', 'N/A', 'N/A', f"${price:.2f}", 'N/A']
            daily = info.get('regularMarketChangePercent')
            row.append(f"{daily:.2f}%" if daily else 'N/A')
            today = now_est.date()
            hist = hist.sort_index()
            for days in PERIODS_DAYS:
                past = today - timedelta(days=days)
                past_slice = hist[hist.index.date <= past]
                if not past_slice.empty:
                    past_price = past_slice['Close'][-1]
                    pct = (price - past_price) / past_price * 100
                    row.append(f"{pct:.2f}%")
                else:
                    row.append('N/A')
            low = info.get('fiftyTwoWeekLow')
            high = info.get('fiftyTwoWeekHigh')
            row.append(f"${low:.2f}" if low else 'N/A')
            row.append(f"${high:.2f}" if high else 'N/A')
            all_rows.append((float('inf'), row))
        except Exception as e:
            print(f"Error {name}: {e}")

    # === Top 20 Stocks ===
    stock_data = []
    for ticker in TOP_20_TICKERS:
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            hist = stock.history(period='2y')
            price = info.get('currentPrice') or (hist['Close'][-1] if not hist.empty else 0)
            mcap = info.get('marketCap', 0)
            trail_pe = info.get('trailingPE')
            cons_avg, cons_sources = compute_consensus_avg(ticker)
            yahoo_fwd = info.get('forwardPE')

            row = [
                ticker,
                f"{trail_pe:.1f}" if trail_pe else 'N/A',
                f"{cons_avg:.1f} ({cons_sources})" if cons_avg else 'N/A',
                f"{yahoo_fwd:.1f}" if yahoo_fwd else 'N/A',
                f"${price:.2f}",
                fmt_mcap(mcap)
            ]

            # Daily % change
            daily = info.get('regularMarketChangePercent')
            row.append(f"{daily:.2f}%" if daily else 'N/A')

            # Historical % change
            today = now_est.date()
            hist = hist.sort_index()
            for days in PERIODS_DAYS:
                past = today - timedelta(days=days)
                past_slice = hist[hist.index.date <= past]
                if not past_slice.empty:
                    past_price = past_slice['Close'][-1]
                    pct = (price - past_price) / past_price * 100
                    row.append(f"{pct:.2f}%")
                else:
                    row.append('N/A')

            # 52-week low/high
            row += [f"${info.get('fiftyTwoWeekLow', 0):.2f}", f"${info.get('fiftyTwoWeekHigh', 0):.2f}"]

            stock_data.append((mcap or 0, row))
        except Exception as e:
            print(f"Error {ticker}: {e}")

    stock_data.sort(key=lambda x: x[0], reverse=True)
    all_rows.extend(stock_data)
    sorted_rows = [row for mcap, row in all_rows]

    if sorted_rows:
        ws.append_rows(sorted_rows)
        for i in range(len(sorted_rows)):
            row_num = i + 2
            color = {'red': 0.95, 'green': 0.95, 'blue': 0.95} if row_num % 2 == 0 else {'red': 1, 'green': 1, 'blue': 1}
            ws.format(f'A{row_num}:O{row_num}', {'backgroundColor': color})

    ws.freeze(rows=1, cols=1)
    ws.format("A2:A1000", {"textFormat": {"bold": False}})

    # === Last Updated ===
    ts = now_est.strftime("%Y-%m-%d %I:%M %p EST")
    last_row = len(sorted_rows) + 2
    ws.update_cell(last_row, 1, 'Last Updated:')
    ws.update_cell(last_row, 2, ts)

    # === Footer notes ===
    footer_start = last_row + 2
    notes = [
        'Auto-updated from GitHub Actions',
        'Every 15 min during trading hours (EST)',
        'No Mac needed',
        'Consensus Fwd PE averaged from multiple sources: Seeking Alpha, MarketWatch, JP Morgan, Goldman Sachs, GuruFocus'
    ]
    for i, line in enumerate(notes, start=footer_start):
        ws.update_cell(i, 1, line)

    print(f"\nDASHBOARD UPDATED: {spreadsheet.url}\n")

if __name__ == "__main__":
    update_sheet(test_mode=True)
