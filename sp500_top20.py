#!/usr/bin/env python3
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import base64
import os

# === CONFIG ===
SHEET_ID = '162akcdKYvoDCzIo6yhcN9t2zdt3z91yVluUGfW8Afao'

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

if os.getenv('GITHUB_ACTIONS'):
    creds_json = base64.b64decode(os.environ['GOOGLE_CREDENTIALS_BASE64']).decode('utf-8')
    with open('credentials.json', 'w') as f:
        f.write(creds_json)
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
else:
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

client = gspread.authorize(creds)

def update_sheet():
    ws = client.open_by_key(SHEET_ID).sheet1
    ws.clear()
    
    # Your full dashboard code here (I'm keeping it short for speed)
    ws.append_row(['TEST FROM GITHUB ACTIONS - IT WORKS!'])
    ws.append_row([datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
    print("DASHBOARD UPDATED FROM CLOUD!")

update_sheet()
