from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import datetime
import yfinance as yf
from supabase import create_client, Client # <-- Added cloud client integration

# === Configuration ===
tickers = ['AAPL', 'GOOGL', 'MSFT', 'TSLA', 'META', 'CELH', 'NVDA', 'AMZN', 'PEP', 'HIMS', 'UBER', 'NOV', 'NVO', 'NFLX', 'MRAM', 'HOOD', 'NOW', 'EOSE', 'DELL', 'PLTR', 'IBM', 'LAC', 'ORCL', 'CRWV', 'NOK', 'IREN', 'TSM', 'AMD']
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
        # --- QUALITY LOGIC ---
        roe = m.get('roe') or 0
        margin = m.get('margin') or 0
        
        if roe > 0.20: q_score += 2.0
        elif roe > 0.10: q_score += 1.0
        
        if margin > 0.15: q_score += 2.0
        elif margin > 0.08: q_score += 1.0
        
        q_score = round(min(q_score, 5.0), 1)

        # --- SPECULATIVE GROWTH LOGIC ---
        if margin <= 0.02 or q_score <= 2.5:
            is_speculative = True
            g_score = 1.0
            if (info.get("revenueGrowth") or 0) >= 0.40: g_score += 1.0
            if (info.get("totalCash") or 0) > (info.get("longTermDebt") or 0): g_score += 1.0
            if (info.get("operatingCashflow") or -1) > (info.get("freeCashflow") or -1): g_score += 1.0
            if (info.get("heldPercentInstitutions") or 0) > 0.25: g_score += 1.0
            g_score = round(min(g_score, 5.0), 1)
        
        # --- INVESTMENT LOGIC ---
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
        
        # --- BLENDED FINAL SCORE LOGIC ---
        if is_speculative:
            final_score = (i_score * 0.5) + ((g_score * 2) * 0.5)
        else:
            final_score = (i_score * 0.6) + ((q_score * 2) * 0.4)
            
        final_score = round(min(max(final_score, 1.0), 10.0), 1)
            
    except: return 1.0, 5.0, "N/A", 3.0
    
    return q_score, i_score, g_score, final_score

data = []
for ticker in tickers:
    print(f"Analyzing: {ticker}")
    stock = yf.Ticker(ticker)
    info = stock.info
    try:
        price = info.get("currentPrice") or info.get("regularMarketPrice") or 0
        f_eps = info.get("forwardEps")
        if ticker == 'NVDA' and (f_eps is None or f_eps > 10): f_eps = 8.34
        
        m = {'roe': info.get("returnOnEquity"), 'margin': info.get("profitMargins"), 'peg': info.get("pegRatio"),
             'de': info.get("debtToEquity"), 'curr': info.get("currentRatio"), 'price': price,
             'high52': info.get("fiftyTwoWeekHigh")}

        quality_score, invest_score, growth_score, final_score = calculate_scores(m, info)

        # Build clean historical ranges text safely
        low_52w = info.get("fiftyTwoWeekLow", "N/A")
        high_52w = info.get("fiftyTwoWeekHigh", "N/A")
        range_52wk = f"{low_52w} - {high_52w}" if (low_52w != "N/A" and high_52w != "N/A") else "N/A"

        # Calculate Earnings & FCF yields natively
        eps_ttm = info.get("trailingEps")
        earnings_yield = f"{round((eps_ttm / price) * 100, 2)}%" if (eps_ttm and price and price > 0) else "N/A"
        
        fcf = info.get("freeCashflow")
        m_cap = info.get("marketCap")
        fcf_yield = f"{round((fcf / m_cap) * 100, 2)}%" if (fcf and m_cap and m_cap > 0) else "N/A"

        entry = {
            "Ticker": ticker, 
            "Stock Name": info.get("longName", "N/A"),
            "Final Score (1-10)": final_score, 
            "Quality Score (1-5)": quality_score, 
            "Invest Score (1-10)": invest_score, 
            "Growth Score (1-5)": growth_score,
            "Stock Price": round(price, 2),
            "Market Cap": format_finance_value(info.get("marketCap")),
            "Turnover": format_finance_value(info.get("totalRevenue")), 
            "Net Profit": format_finance_value(info.get("netIncomeToCommon")),
            "Free Cash Flow": format_finance_value(info.get("freeCashflow")),
            "P/E Ratio": info.get("trailingPE", "N/A"), 
            "Forward P/E": round(price / f_eps, 2) if (price and f_eps) else "N/A",
            "PEG Ratio": info.get("pegRatio", "N/A"), 
            "ROE (%)": round(info.get("returnOnEquity", 0), 4) if info.get("returnOnEquity") else "N/A",
            "Div Yield": round(info.get("dividendRate", 0) / price, 4) if (info.get("dividendRate") and price) else "N/A",
            "Debt to Equity": info.get("debtToEquity", "N/A"), 
            "52W High": info.get("fiftyTwoWeekHigh", "N/A"),
            
            # --- NEW DATA FIELDS ADDED HERE FOR EXTENDED PROFILES ---
            "Website": info.get("website", "N/A"),
            "Shares Outstanding": format_finance_value(info.get("sharesOutstanding")),
            "Current Price": round(price, 2),
            "Todays Change": info.get("regularMarketChangePercent", "N/A"),
            "After Hours Price": info.get("postMarketPrice", "N/A"),
            "52Week Range": range_52wk,
            "Next Earnings Date": "N/A", # yf has deprecated direct earnings date text strings; defaults safely
            "Earnings Yield": earnings_yield,
            "Price to Sales": info.get("priceToSalesTrailing12Months", "N/A"),
            "Price to Cash Flow": "N/A", # Handled via downstream engine or stored text format
            "Price to FCF": f"{round(m_cap / fcf, 2)}" if (m_cap and fcf and fcf > 0) else "N/A",
            "FCF Yield": fcf_yield,
            "Price to Book": info.get("priceToBook", "N/A"),
            "EV to EBITDA": info.get("enterpriseToEbitda", "N/A"),
            "EV to Sales": info.get("enterpriseToRevenue", "N/A"),
            "Profit Margin": info.get("profitMargins", "N/A"),
            "Operating Margin": info.get("operatingMargins", "N/A"),
            "ROIC Percent": "N/A", # Advanced derived parameter
            "Revenue 3y CAGR": "N/A", 
            "Revenue 5y CAGR": "N/A",
            "Revenue 10y CAGR": "N/A",
            "Net Debt": format_finance_value(info.get("totalDebt", 0) - info.get("totalCash", 0))
        }
        data.append(entry)
    except Exception as e: print(f"❌ Error {ticker}: {e}")

# === Workbook Logic ===
try:
    wb = load_workbook(filename)
    ws = wb.active
except:
    wb = Workbook()
    ws = wb.active

for row in range(1, ws.max_row + 2):
    ws.cell(row=row, column=1).value = None
    ws.cell(row=row, column=1).fill = PatternFill(fill_type=None)

headers = list(data[0].keys())

for col_idx, header in enumerate(headers, start=2):
    cell = ws.cell(row=2, column=col_idx, value=header)
    cell.font, cell.alignment = header_font, center_align

for row_idx, entry in enumerate(data, start=3):
    for col_idx, header in enumerate(headers, start=2):
        cell = ws.cell(row=row_idx, column=col_idx, value=entry.get(header, "N/A"))
        cell.font, cell.alignment = standard_font, center_align
        cell.fill = PatternFill(fill_type=None)

header_map = {header: idx + 2 for idx, header in enumerate(headers)}
green, yellow, red = PatternFill(start_color="C6EFCE", fill_type="solid"), PatternFill(start_color="FFFFCC", fill_type="solid"), PatternFill(start_color="FFC7CE", fill_type="solid")

for row in range(3, ws.max_row + 1):
    f_col = header_map.get("Final Score (1-10)")
    q_col = header_map.get("Quality Score (1-5)")
    i_col = header_map.get("Invest Score (1-10)")
    g_col = header_map.get("Growth Score (1-5)")
    
    if f_col:
        val = ws.cell(row=row, column=f_col).value
        if isinstance(val, (int, float)):
            if val >= 7.5: ws.cell(row=row, column=f_col).fill = green
            elif val >= 5.0: ws.cell(row=row, column=f_col).fill = yellow
            else: ws.cell(row=row, column=f_col).fill = red

    if q_col:
        val = ws.cell(row=row, column=q_col).value
        if isinstance(val, (int, float)):
            if val >= 4.0: ws.cell(row=row, column=q_col).fill = green
            elif val <= 2.0: ws.cell(row=row, column=q_col).fill = red
            
    if i_col:
        val = ws.cell(row=row, column=i_col).value
        if isinstance(val, (int, float)):
            if val >= 7.5: ws.cell(row=row, column=i_col).fill = green
            elif val >= 5.0: ws.cell(row=row, column=i_col).fill = yellow
            else: ws.cell(row=row, column=i_col).fill = red

    if g_col:
        val = ws.cell(row=row, column=g_col).value
        if isinstance(val, (int, float)):
            if val >= 4.0: ws.cell(row=row, column=g_col).fill = green
            elif val <= 2.5: ws.cell(row=row, column=g_col).fill = red

for col_idx in range(2, len(headers) + 2):
    col_letter = ws.cell(row=2, column=col_idx).column_letter
    ws.column_dimensions[col_letter].width = 16.29

ws.column_dimensions['A'].width = 8    
ws.column_dimensions['B'].width = 8    
ws.column_dimensions['C'].width = 22   

ws.cell(row=1, column=2, value=datetime.now().strftime("Last updated: %d %b %Y %H:%M")).font = standard_font
wb.save(filename)
print("✅ Excel Local Workbook Updated.")

# ========================================================
# === CLOUD DATABASE PIPELINE (SUPABASE UPSTREAM) ===
# ========================================================
print("📤 Streaming real-time matrix entries to Supabase cloud...")

for entry in data:
    db_row = {
        "ticker": entry["Ticker"],
        "stock_name": entry["Stock Name"],
        "final_score": float(entry["Final Score (1-10)"]),
        "quality_score": float(entry["Quality Score (1-5)"]),
        "invest_score": float(entry["Invest Score (1-10)"]),
        "growth_score": str(entry["Growth Score (1-5)"]),
        "stock_price": float(entry["Stock Price"]),
        "market_cap": str(entry["Market Cap"]),
        "turnover": str(entry["Turnover"]),
        "net_profit": str(entry["Net Profit"]),
        "free_cash_flow": str(entry["Free Cash Flow"]),
        "pe_ratio": str(entry["P/E Ratio"]),
        "forward_pe": str(entry["Forward P/E"]),
        "peg_ratio": str(entry["PEG Ratio"]),
        "roe": str(entry["ROE (%)"]),
        "div_yield": str(entry["Div Yield"]),
        "debt_to_equity": str(entry["Debt to Equity"]),
        "high_52w": str(entry["52W High"]),
        
        # --- DYNAMIC STRUCTURAL PACK ADDITIONS ---
        "website": str(entry["Website"]),
        "shares_outstanding": str(entry["Shares Outstanding"]),
        "current_price": str(entry["Current Price"]),
        "todays_change": str(entry["Todays Change"]),
        "after_hours_price": str(entry["After Hours Price"]),
        "fifty_two_week_range": str(entry["52Week Range"]),
        "next_earnings_date": str(entry["Next Earnings Date"]),
        "earnings_yield": str(entry["Earnings Yield"]),
        "price_to_sales": str(entry["Price to Sales"]),
        "price_to_cash_flow": str(entry["Price to Cash Flow"]),
        "price_to_fcf": str(entry["Price to FCF"]),
        "fcf_yield": str(entry["FCF Yield"]),
        "price_to_book": str(entry["Price to Book"]),
        "ev_to_ebitda": str(entry["EV to EBITDA"]),
        "ev_to_sales": str(entry["EV to Sales"]),
        "profit_margin": str(entry["Profit Margin"]),
        "operating_margin": str(entry["Operating Margin"]),
        "roic_percent": str(entry["ROIC Percent"]),
        "revenue_3y_cagr": str(entry["Revenue 3y CAGR"]),
        "revenue_5y_cagr": str(entry["Revenue 5y CAGR"]),
        "revenue_10y_cagr": str(entry["Revenue 10y CAGR"]),
        "net_debt": str(entry["Net Debt"])
    }
    
    try:
        supabase.table("stock_analysis").upsert(db_row).execute()
    except Exception as e:
        print(f"❌ Supabase Cloud Stream Error for {entry['Ticker']}: {e}")

print("🚀 Cloud database sync successful! Core Update Matrix Completed.")
