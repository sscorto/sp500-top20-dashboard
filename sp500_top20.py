#!/usr/bin/env python3
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import sys
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

# === CONSENSUS DATA WITH SOURCE ===
CONSENSUS_FWD_PE = {
    'AMZN': (32.3, "Yahoo Finance"),
    'GOOGL': (26.2, "Yahoo Finance"),
    'META': (24.5, "Yahoo Finance"),
    'TSLA': (85.2, "Yahoo Finance"),
    'MSFT': (34.1, "Yahoo Finance"),
    'NVDA': (42.1, "Yahoo Finance"),
    'AAPL': (33.5, "Yahoo Finance"),
}

TOP_20_TICKERS = [
    'NVDA', 'MSFT', 'AAPL', 'AVGO', 'GOOGL', 'AMZN', 'META',
    'TSLA', 'GOOG', 'BRK-B', 'LLY', 'JPM', 'UNH', 'XOM',
    'V', 'PG', 'MA', 'JNJ', 'HD', 'ORCL'
]

FIXED_ASSETS = [
    ('^GSPC',   'S&P 500 Index'),
    ('^IXIC',   'Nasdaq 100'),
    ('^GSPTSE', 'TSX Index'),
    ('GC=F',    'Gold (USD)'),
    ('BTC-USD', 'Bitcoin (USD)'),
    ('^VIX',    'VIX')
]

PERIODS_DAYS = [5, 14, 30, 90, 182, 365]

def fmt_mcap(m):
    if not m or m == 0: return 'N/A'
    if m >= 1e12: return f"${m/1e12:.2f}T"
    if m >= 1e9: return f"${m/1e9:.2f}B"
    return f"${m/1e6:.2f}M"

def update_sheet(test_mode=False):
    est = pytz.timezone("US/Eastern")
    now_est = datetime.now(est)

    if not test_mode:
        # Normal market hours check
        if now_est.weekday() >= 5 or now_est.hour < 9 or (now_est.hour == 9 and now_est.minute < 30) or now_est.hour > 16:
            print("Market closed: skipping update.")
            sys.exit()
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

    # Fixed assets
    for ticker, name in FIXED_ASSETS:
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            price = info.get('regularMarketPrice') or info.get('currentPrice') or 0
            row = [name, 'N/A', 'N/A', 'N/A', f"${price:.2f}", 'N/A']
            all_rows.append((float('inf'), row))
        except:
            continue

    # Top 20 tickers
    for ticker in TOP_20_TICKERS:
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            price = info.get('currentPrice') or 0
            mcap = info.get('marketCap', 0)
            trail_pe = info.get('trailingPE')
            cons_val, cons_source = CONSENSUS_FWD_PE.get(ticker, (info.get('forwardPE'), "Yahoo Finance"))
            yahoo_fwd = info.get('forwardPE')
            row = [
                ticker,
                f"{trail_pe:.1f}" if trail_pe else 'N/A',
                f"{cons_val:.1f} ({cons_source})" if cons_val else 'N/A',
                f"{yahoo_fwd:.1f}" if yahoo_fwd else 'N/A',
                f"${price:.2f}",
                fmt_mcap(mcap)
            ]
            all_rows.append((mcap or 0, row))
        except:
            continue

    all_rows.sort(key=lambda x: x[0], reverse=True)
    sorted_rows = [row for mcap, row in all_rows]
    if sorted_rows:
        ws.append_rows(sorted_rows)

    # Last updated EST
    ts = now_est.strftime("%Y-%m-%d %I:%M %p EST")
    last_row = len(sorted_rows) + 2
    ws.update_cell(last_row, 1, 'Last Updated:')
    ws.update_cell(last_row, 2, ts)

    # Footer notes
    footer_start = last_row + 2
    notes = [
        'Auto-updated from GitHub Actions',
        'Every 15 min during trading hours (EST)',
        'No Mac needed',
        'Consensus Fwd PE sourced from Yahoo Finance'
    ]
    for i, line in enumerate(notes, start=footer_start):
        ws.update_cell(i, 1, line)

    print(f"\nTEST DASHBOARD UPDATED: {spreadsheet.url}\n")

if __name__ == "__main__":
    update_sheet(test_mode=True)
