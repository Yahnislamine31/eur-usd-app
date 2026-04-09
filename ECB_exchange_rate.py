import streamlit as st
import requests
import pandas as pd
from io import StringIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import tempfile
from decimal import Decimal, ROUND_HALF_UP

st.set_page_config(page_title="EUR/USD Downloader Pro", layout="centered")

st.title("📊 EUR/USD Exchange Rate Downloader")

# ── HELPERS ─────────────────────────────────────────────────
def precise_round(value, precision=4):
    """Force rounding up for .5 cases (1.04125 -> 1.0413)"""
    if pd.isna(value):
        return value
    return float(Decimal(str(value)).quantize(Decimal('1.' + '0' * precision), rounding=ROUND_HALF_UP))

def write_headers(ws, headers):
    hfill = PatternFill("solid", start_color="003366")
    hfont = Font(bold=True, color="FFFFFF")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = hfill
        cell.font = hfont
        cell.alignment = Alignment(horizontal="center")

def style_row(ws, r, ncols, even):
    if even:
        for c in range(1, ncols + 1):
            ws.cell(r, c).fill = PatternFill("solid", start_color="EEF3FA")

def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

# ── USER INPUTS ─────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start date", pd.to_datetime("2015-01-01"))
with col2:
    end_date = st.date_input("End date", pd.to_datetime("today"))

if start_date > end_date:
    st.error("Start date must be before end date")
    st.stop()

if st.button("🚀 Generate Excel File"):
    with st.spinner("Fetching data from ECB..."):
        # ECB API URLs
        # Daily data is used as the base for Weekly and Annual calculations to ensure accuracy
        url_d = "https://data-api.ecb.europa.eu/service/data/EXR/D.USD.EUR.SP00.A"
        params_d = {
            "startPeriod": start_date.strftime("%Y-%m-%d"),
            "endPeriod": end_date.strftime("%Y-%m-%d"),
            "format": "csvdata"
        }
        
        # Monthly data for the Monthly sheet
        url_m = "https://data-api.ecb.europa.eu/service/data/EXR/M.USD.EUR.SP00.A"
        params_m = {
            "startPeriod": start_date.strftime("%Y-%m"),
            "endPeriod": end_date.strftime("%Y-%m"),
            "format": "csvdata"
        }

        try:
            daily_raw = pd.read_csv(StringIO(requests.get(url_d, params=params_d).text))
            monthly_raw = pd.read_csv(StringIO(requests.get(url_m, params=params_m).text))
        except Exception as e:
            st.error(f"Error fetching data: {e}")
            st.stop()

    # ── Data Processing ─────────────────────────────────────
    # Process Daily Base
    df_d = daily_raw[["TIME_PERIOD", "OBS_VALUE"]].copy()
    df_d.columns = ["Date", "EUR_USD"]
    df_d["Date"] = pd.to_datetime(df_d["Date"])
    df_d["EUR_USD"] = pd.to_numeric(df_d["EUR_USD"], errors="coerce")
    df_d = df_d.dropna().sort_values("Date")
    
    # Apply precise rounding
    df_d["EUR_USD"] = df_d["EUR_USD"].apply(precise_round)
    df_d["USD_EUR"] = (1 / df_d["EUR_USD"]).apply(precise_round)

    # ── Excel Workbook ──────────────────────────────────────
    wb = Workbook()
    
    # 1. ANNUAL SHEET
    ws_a = wb.active
    ws_a.title = "Annual"
    df_a = df_d.groupby(df_d['Date'].dt.year)['EUR_USD'].mean().reset_index()
    df_a.columns = ['Year', 'EUR_USD']
    df_a['USD_EUR'] = (1 / df_a['EUR_USD']).apply(precise_round)
    df_a['EUR_USD'] = df_a['EUR_USD'].apply(precise_round)
    
    headers_a = ["Year", "Avg EUR/USD", "Avg USD/EUR"]
    write_headers(ws_a, headers_a)
    for i, row in df_a.iterrows():
        r = i + 2
        ws_a.cell(r, 1, row["Year"])
        ws_a.cell(r, 2, row["EUR_USD"])
        ws_a.cell(r, 3, row["USD_EUR"])
        style_row(ws_a, r, 3, i % 2 == 0)
    set_widths(ws_a, {"A": 10, "B": 15, "C": 15})

    # 2. MONTHLY SHEET
    ws_m = wb.create_sheet("Monthly")
    df_m = monthly_raw[["TIME_PERIOD", "OBS_VALUE"]].copy()
    df_m.columns = ["Month", "EUR_USD"]
    df_m["EUR_USD"] = pd.to_numeric(df_m["EUR_USD"], errors="coerce")
    df_m = df_m.dropna().sort_values("Month")
    df_m["Year"] = df_m["Month"].str[:4]
    df_m["Month Name"] = pd.to_datetime(df_m["Month"]).dt.strftime("%B")
    df_m["EUR_USD"] = df_m["EUR_USD"].apply(precise_round)
    df_m["USD_EUR"] = (1 / df_m["EUR_USD"]).apply(precise_round)

    headers_m = ["Month", "Year", "Month Name", "EUR/USD", "USD/EUR"]
    write_headers(ws_m, headers_m)
    for i, row in df_m.iterrows():
        r = i + 2
        ws_m.cell(r, 1, row["Month"])
        ws_m.cell(r, 2, row["Year"])
        ws_m.cell(r, 3, row["Month Name"])
        ws_m.cell(r, 4, row["EUR_USD"])
        ws_m.cell(r, 5, row["USD_EUR"])
        style_row(ws_m, r, 5, i % 2 == 0)
    set_widths(ws_m, {"A": 12, "B": 8, "C": 14, "D": 14, "E": 14})

    # 3. WEEKLY SHEET
    ws_w = wb.create_sheet("Weekly")
    # Group by ISO week
    df_w = df_d.copy()
    df_w['Year'] = df_w['Date'].dt.isocalendar().year
    df_w['Week'] = df_w['Date'].dt.isocalendar().week
    
    # Calculate Week Start (Monday) and Week End (Sunday)
    df_weekly = df_w.groupby(['Year', 'Week'])['EUR_USD'].mean().reset_index()
    
    # Logic to get the Monday of that specific ISO week
    df_weekly['Week Start'] = df_weekly.apply(lambda x: pd.to_datetime(f"{int(x['Year'])}-W{int(x['Week'])}-1", format='%G-W%V-%u').strftime('%Y-%m-%d'), axis=1)
    df_weekly['Week End'] = df_weekly.apply(lambda x: (pd.to_datetime(f"{int(x['Year'])}-W{int(x['Week'])}-1", format='%G-W%V-%u') + pd.Timedelta(days=6)).strftime('%Y-%m-%d'), axis=1)
    
    df_weekly['USD_EUR'] = (1 / df_weekly['EUR_USD']).apply(precise_round)
    df_weekly['EUR_USD'] = df_weekly['EUR_USD'].apply(precise_round)

    headers_w = ["Year", "Week #", "Week Start (Mon)", "Week End (Sun)", "EUR/USD", "USD/EUR"]
    write_headers(ws_w, headers_w)
    for i, row in df_weekly.iterrows():
        r = i + 2
        ws_w.cell(r, 1, row["Year"])
        ws_w.cell(r, 2, row["Week"])
        ws_w.cell(r, 3, row["Week Start"])
        ws_w.cell(r, 4, row["Week End"])
        ws_w.cell(r, 5, row["EUR_USD"])
        ws_w.cell(r, 6, row["USD_EUR"])
        style_row(ws_w, r, 6, i % 2 == 0)
    set_widths(ws_w, {"A": 8, "B": 8, "C": 18, "D": 18, "E": 14, "F": 14})

    # 4. DAILY SHEET
    ws_d_sheet = wb.create_sheet("Daily")
    headers_d = ["Date", "EUR/USD", "USD/EUR"]
    write_headers(ws_d_sheet, headers_d)
    for i, row in df_d.reset_index().iterrows():
        r = i + 2
        ws_d_sheet.cell(r, 1, row["Date"].strftime("%Y-%m-%d"))
        ws_d_sheet.cell(r, 2, row["EUR_USD"])
        ws_d_sheet.cell(r, 3, row["USD_EUR"])
        style_row(ws_d_sheet, r, 3, i % 2 == 0)
    set_widths(ws_d_sheet, {"A": 12, "B": 14, "C": 14})

    # ── Finalize ────────────────────────────────────────────
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    
    with open(tmp.name, "rb") as f:
        st.download_button(
            "📥 Download Multi-Period Excel",
            f,
            file_name=f"EUR_USD_Full_Report_{start_date}_to_{end_date}.xlsx"
        )
    st.success("✅ Analysis Complete! You have Annual, Monthly, Weekly, and Daily data in one file.")