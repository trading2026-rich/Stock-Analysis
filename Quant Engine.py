from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import datetime
import yfinance as yf
from supabase import create_client, Client
import pandas as pd
import urllib.request
import io
import time

print("Fetching live S&P 500 ticker index...")
sp500_url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"

# Create the request with the browser disguise
req = urllib.request.Request(
    sp500_url, 
    headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
)

# Download the page content safely first
with urllib.request.urlopen(req) as response:
    html_content = response.read()

# Pass the already-downloaded content straight into pandas using an in-memory stream
table = pd.read_html(io.BytesIO(html_content))
tickers = table[0]['Symbol'].tolist()

# Clean up tickers for yFinance compatibility (e.g., BRK.B to BRK-B)
tickers = [t.replace('.', '-') for t in tickers]
filename = "Stock_Analysis.xlsx"

# === Supabase Keys ===
SUPABASE_URL = "https://oczkxudrukyotgmpvpdv.supabase.co"
SUPABASE_KEY = "sb_publishable_rKa-z8ZCJ76o2AsUqpGHHg_jAQJ74qN"
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

standard_font = Font(name='Calibri', size=10)
header_font = Font(name='Calibri', size=10, bold=True)
center_align = Alignment(horizontal='center', vertical='center')

def format_finance_value(val):
    if not isinstance(val, (int, float)) or val == 0: return "N/A"
    abs_val = abs(val)
    if abs_val >= 1e12: return f"{round(val / 1e12, 2)}T"
    elif abs_val >= 1e9: return f"{round(val / 1e9, 2)}B"
    elif abs_val >= 1e6: return f"{round(val / 1e6, 1)}M"
    else: return f"{round(val, 2)}"

def calculate_scores(m, info):
    """Returns (Quality_Score, Investment_Score, Growth_Score, Final_Score)"""
    q_score = 1.0 
    i_score = 0.0 
    g_score = "N/A" 
    is_speculative = False
    
    try:
        roe = m.get('roe') or 0
        margin = m.get('margin') or 0
        
        if roe > 0.20: q_score += 2.0
        elif roe > 0.10: q_score += 1.0
        
        if margin > 0.15: q_score += 2.0
        elif margin > 0.08: q_score += 1.0
        
        q_score = round(min(q_score, 5.0), 1)

        if margin <= 0.02 or q_score <= 2.5:
            is_speculative = True
            g_score = 1.0
            if (info.get("revenueGrowth") or 0) >= 0.40: g_score += 1.0
            if (info.get("totalCash") or 0) > (info.get("longTermDebt") or 0): g_score += 1.0
            if (info.get("operatingCashflow") or -1) > (info.get("freeCashflow") or -1): g_score += 1.0
            if (info.get("heldPercentInstitutions") or 0) > 0.25: g_score += 1.0
            g_score = round(min(g_score, 5.0), 1)
        
        i_score = q_score 
        peg = m.get('peg')
        if isinstance(peg, (int, float)):
            if peg < 1.0: i_score += 3.0
            elif peg < 2.0: i_score += 1.5
            
        if (m.get('de') or 999) < 100: i_score += 1.0
        if (m.get('curr') or 0) > 1.2: i_score += 1.0
        
        price, high52 = m.get('price'), m.get('high52')
        if price and high52:
            if price >= (high52 * 0.97): i_score -= 2.0 
            elif price <= (high52 * 0.80): i_score += 1.5 
        
        i_score = round(min(max(i_score, 1.0), 10.0), 1)
        
        if is_speculative:
            final_score = (i_score * 0.5) + ((g_score * 2) * 0.5)
        else:
            final_score = (i_score * 0.6) + ((q_score * 2) * 0.4)
            
        final_score = round(min(max(final_score, 1.0), 10.0), 1)
            
    except: return 1.0, 5.0, "N/A", 3.0
    
    return q_score, i_score, g_score, final_score

def to_float(val):
    if isinstance(val, (int, float)):
        return float(round(val, 2))
    try:
        cleaned = str(val).replace("%", "").strip()
        return float(round(float(cleaned), 2))
    except:
        return None

# === Continuous Loop Execution ===
if __name__ == "__main__":
    while True:
        print(f"\n🚀 Starting Live Matrix Sync: {datetime.now().strftime('%H:%M:%S')}")
        
        try:
            data = []
            print("Processing and analyzing stocks...")
            
            for ticker in tickers:
                print(f"Analyzing: {ticker}")
                stock = yf.Ticker(ticker)
                info = stock.info
                
                # --- DYNAMIC HISTORICAL CAGR CALCULATION MODULE ---
                rev_3y_cagr = "N/A"
                try:
                    financials = stock.financials
                    if "Total Revenue" in financials.index:
                        revenue_series = financials.loc["Total Revenue"].dropna()
                        if len(revenue_series) >= 4:
                            latest_rev = revenue_
