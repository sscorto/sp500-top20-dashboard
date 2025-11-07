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

# === CONFIG ===
SHEET_ID = '162akcdKYvoDCzIo6yhcN9t2zdt3z91yVluUGfW8Afao'
REFRESH_INTERVAL = 15  # minutes (only used locally)

# Google auth – works locally AND on GitHub Actions
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

if os.getenv('GITHUB_ACTIONS'):
    # === GITHUB ACTIONS MODE ===
    # Secret is stored as base64 string → decode it
    creds_json = base64.b64decode(os.environ['GOOGLE_CREDENTIALS_BASE64']).decode('utf-8')
    with open('credentials.json', 'w') as f:
        f.write(creds_json)
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
else:
    # === LOCAL MAC MODE ===
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

client = gspread.authorize(creds)

def update_sheet():
    spreadsheet = client.open_by_key(SHEET_ID)
    ws = spreadsheet.sheet1
    ws.clear()

    # === FIXED TOP ROWS ===
    fixed_assets = [
        ('^GSPC',   'S&P 500 Index'),
        ('^IXIC',   'Nasdaq 100'),
        ('^GSPTSE', 'TSX Index'),
        ('GC=F',    'Gold (USD)'),
        ('BTC-USD', 'Bitcoin (USD)'),
        ('^VIX',    'VIX')
    ]

    top_20_tickers = [
        'NVDA', 'MSFT', 'AAPL', 'AVGO', 'GOOGL', 'AMZN', 'META',
        'TSLA', 'GOOG', 'BRK-B', 'LLY', 'JPM', 'UNH', 'XOM',
        'V', 'PG', 'MA', 'JNJ', 'HD', 'ORCL'
    ]

    CONSENSUS_FWD_PE = {
        'AMZN': 32.3, 'GOOGL': 26.2, 'META': 24.5, 'TSLA': 85.2,
        'MSFT': 34.1, 'NVDA': 42.1, 'AAPL': 33.5,
    }

    headers = [
        'Stock', 'Trailing PE', 'Consensus Fwd PE', 'Forward PE (Yahoo)', 'Current Price', 'Market Cap',
        'Daily %', '5D %', '2W %', '1M %', '3M %', '6M %', '12M %',
        '52W Low', '52W High'
    ]
    ws.append_row(headers)
    ws.format('A1:O1', {
        'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9},
        'textFormat': {'bold': True},
        'horizontalAlignment': 'CENTER'
    })

    def fmt_mcap(m):
        if not m or m == 0: return 'N/A'
        if m >= 1e12: return f"${m/1e12:.2f}T"
        if m >= 1e9: return f"${m/1e9:.2f}B"
        return f"${m/1e6:.2f}M"

    periods_days = [5, 14, 30, 90, 182, 365]
    all_rows = []

    # === 1. FIXED ASSETS ===
    for ticker, name in fixed_assets:
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            hist = stock.history(period='2y')
            price = info.get('regularMarketPrice') or info.get('currentPrice') or (hist['Close'][-1] if not hist.empty else None)
            if not price: continue

            row = [name, 'N/A', 'N/A', 'N/A', f"${price:.2f}", 'N/A']
            daily = info.get('regularMarketChangePercent')
            row.append(f"{daily:.2f}%" if daily else 'N/A')

            today = datetime.now().date()
            hist = hist.sort_index()
            for days in periods_days:
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
            print(f"Fetched {name}")
        except Exception as e:
            print(f"Error {name}: {e}")

    # === 2. TOP 20 STOCKS ===
    stock_data = []
    for ticker in top_20_tickers:
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            hist = stock.history(period='2y')
            price = info.get('currentPrice') or (hist['Close'][-1] if not hist.empty else None)
            mcap = info.get('marketCap', 0)
            if not price: continue

            row = [ticker]
            trail_pe = info.get('trailingPE')
            row.append(f"{trail_pe:.1f}" if trail_pe else 'N/A')
            cons = CONSENSUS_FWD_PE.get(ticker, info.get('forwardPE'))
            row.append(f"{cons:.1f}" if cons else 'N/A')
            yahoo_fwd = info.get('forwardPE')
            row.append(f"{yahoo_fwd:.1f}" if yahoo_fwd else 'N/A')
            row.append(f"${price:.2f}")
            row.append(fmt_mcap(mcap))

            daily = info.get('regularMarketChangePercent')
            row.append(f"{daily:.2f}%" if daily else 'N/A')

            today = datetime.now().date()
            hist = hist.sort_index()
            for days in periods_days:
                past = today - timedelta(days=days)
                past_slice = hist[hist.index.date <= past]
                if not past_slice.empty:
                    past_price = past_slice['Close'][-1]
                    pct = (price - past_price) / past_price * 100
                    row.append(f"{pct:.2f}%")
                else:
                    row.append('N/A')

            row += [
                f"${info.get('fiftyTwoWeekLow', 'N/A'):.2f}",
                f"${info.get('fiftyTwoWeekHigh', 'N/A'):.2f}"
            ]
            stock_data.append((mcap or 0, row))
            print(f"Fetched {ticker}")
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

    # === FREEZE + BOLD-FIX ===
    ws.freeze(rows=1, cols=1)
    ws.format("A2:A1000", {"textFormat": {"bold": False}})

    # === TIMESTAMP ===
    ts = datetime.now().strftime("%Y-%m-%d %I:%M %p")
    last_row = len(sorted_rows) + 2
    ws.update_cell(last_row, 1, 'Last Updated:')
    ws.update_cell(last_row, 2, ts)

    # === FOOTER ===
    footer_start = last_row + 2
    notes = [
        'Notes (Nov 2025):',
        'Top 6 rows fixed • Live market-cap sort below',
        'Auto-updated by GitHub Actions every hour',
        'Source: yfinance + GuruFocus consensus PE'
    ]
    for i, line in enumerate(notes, start=footer_start):
        ws.update_cell(i, 1, line)

    ws.format(f'A{footer_start}:O{footer_start+len(notes)-1}', {
        'textFormat': {'italic': True, 'fontSize': 9},
        'backgroundColor': {'red': 0.98, 'green': 0.98, 'blue': 0.98}
    })

    print(f"\nDASHBOARD UPDATED: {spreadsheet.url}\n")

# === RUN ONCE (GitHub) OR LOOP (local) ===
if __name__ == "__main__":
    if os.getenv('GITHUB_ACTIONS'):
        update_sheet()  # GitHub: run once and exit
    else:
        # Local Mac: keep refreshing every 15 min
        print("Local mode – auto-refresh every 15 min (Ctrl+C to stop)")
        while True:
            try:
                update_sheet()
                print(f"Sleeping {REFRESH_INTERVAL} min...")
                time.sleep(REFRESH_INTERVAL * 60)
            except KeyboardInterrupt:
                print("\nStopped by user.")
                break
            except Exception as e:
                print(f"Error: {e}")
                time.sleep(60)