import streamlit as st
import requests
import pandas as pd
from io import StringIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import tempfile
import time
from decimal import Decimal, ROUND_HALF_UP

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Global Data Downloader",
    page_icon="🌐",
    layout="centered",
)

# ── LIGHT THEME CSS ───────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;600&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .stApp { background-color: #f8f9fb; color: #1a1d23; }

    .app-header {
        background: linear-gradient(135deg, #003366 0%, #0055a4 100%);
        border-radius: 10px;
        padding: 1.4rem 1.8rem 1.2rem;
        margin-bottom: 1.5rem;
    }
    .app-header h1 {
        font-family: 'IBM Plex Mono', monospace;
        color: #ffffff;
        font-size: 1.6rem;
        font-weight: 600;
        margin: 0 0 0.3rem 0;
    }
    .app-header p { color: #a8c8f0; font-size: 0.88rem; margin: 0; }

    .info-banner {
        background: #eef4ff;
        border: 1px solid #b8d0f5;
        border-left: 4px solid #0055a4;
        border-radius: 6px;
        padding: 0.55rem 1rem;
        font-size: 0.84rem;
        color: #1a3a6b;
        margin-bottom: 0.9rem;
    }

    .warn-banner {
        background: #fff8e6;
        border: 1px solid #f5d580;
        border-left: 4px solid #e6a817;
        border-radius: 6px;
        padding: 0.5rem 1rem;
        font-size: 0.83rem;
        color: #7a5000;
        margin-top: 0.4rem;
    }

    .stButton > button {
        background: #003366 !important;
        color: #fff !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 7px !important;
        padding: 0.65rem 2rem !important;
        width: 100% !important;
        font-size: 0.98rem !important;
    }
    .stButton > button:hover { background: #0055a4 !important; }

    .stDownloadButton > button {
        background: #1a7f37 !important;
        color: #fff !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 7px !important;
        width: 100% !important;
    }

    .stCheckbox label { font-size: 0.94rem; color: #2d3748; }

    div[data-testid="stExpander"] {
        background: #ffffff !important;
        border: 1px solid #e2e6ed !important;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 0.8rem;
    }

    hr { border-color: #e2e6ed !important; }
</style>
""", unsafe_allow_html=True)


# ── CONSTANTS ─────────────────────────────────────────────────────────────────
WORLD_BANK_BASE = "https://api.worldbank.org/v2"
HEADER_COLOR    = "003366"
ALTERNATE_ROW   = "EEF3FA"

# 20 indicators in 6 themed groups
WB_INDICATOR_GROUPS: dict[str, dict[str, str]] = {
    "📈 Growth & Output": {
        "GDP growth (annual %)":           "NY.GDP.MKTP.KD.ZG",
        "GDP, constant prices (USD)":      "NY.GDP.MKTP.KD",
        "GDP per capita, constant prices": "NY.GDP.PCAP.KD",
        "GDP deflator (annual %)":         "NY.GDP.DEFL.KD.ZG",
    },
    "👥 Population": {
        "Population, total":               "SP.POP.TOTL",
        "Population growth (annual %)":    "SP.POP.GROW",
    },
    "💼 Labour Market": {
        "Unemployment rate (%)":              "SL.UEM.TOTL.ZS",
        "Labor force, total":                 "SL.TLF.TOTL.IN",
        "Labor force participation rate (%)": "SL.TLF.CACT.ZS",
        "GDP per person employed (USD)":      "SL.GDP.PCAP.EM.KD",
    },
    "💰 Prices & Savings": {
        "Inflation, consumer prices (%)":     "FP.CPI.TOTL.ZG",
        "Gross capital formation (% of GDP)": "NE.GDI.TOTL.ZS",
        "Gross savings (% of GDP)":           "NY.GNS.ICTR.ZS",
    },
    "🌍 Trade & External": {
        "Current account balance (% of GDP)":  "BN.CAB.XOKA.GD.ZS",
        "Exports of goods & services (% GDP)": "NE.EXP.GNFS.ZS",
        "Imports of goods & services (% GDP)": "NE.IMP.GNFS.ZS",
        "FDI, net inflows (% of GDP)":         "BX.KLT.DINV.WD.GD.ZS",
    },
    "🏛️ Fiscal": {
        "General government debt (% of GDP)":  "GC.DOD.TOTL.GD.ZS",
        "Central govt revenue (% of GDP)":     "GC.REV.XGRT.GD.ZS",
        "Central govt expenditure (% of GDP)": "GC.XPN.TOTL.GD.ZS",
    },
}

# Flat label → code lookup
WB_INDICATORS: dict[str, str] = {
    label: code
    for grp in WB_INDICATOR_GROUPS.values()
    for label, code in grp.items()
}

WB_INDICATOR_NOTES: dict[str, str] = {
    "NY.GDP.MKTP.KD.ZG":   "Constant 2015 USD. Annual % change.",
    "NY.GDP.MKTP.KD":      "Constant 2015 USD. Not inflation-adjusted.",
    "NY.GDP.PCAP.KD":      "Constant 2015 USD per capita.",
    "NY.GDP.DEFL.KD.ZG":   "Annual % change in implicit price deflator.",
    "SP.POP.TOTL":         "De facto population, mid-year estimates.",
    "SP.POP.GROW":         "Annual population growth rate (%).",
    "SL.UEM.TOTL.ZS":      "ILO modelled estimates. % of total labour force.",
    "SL.TLF.TOTL.IN":      "Total labour force (persons).",
    "SL.TLF.CACT.ZS":      "Labour force as % of population ages 15+.",
    "SL.GDP.PCAP.EM.KD":   "Constant 1990 PPP USD per employed person.",
    "FP.CPI.TOTL.ZG":      "Consumer price index, annual % change.",
    "NE.GDI.TOTL.ZS":      "Gross capital formation as % of GDP.",
    "NY.GNS.ICTR.ZS":      "Gross savings as % of GDP.",
    "BN.CAB.XOKA.GD.ZS":   "Current account balance as % of GDP.",
    "NE.EXP.GNFS.ZS":      "Exports of goods and services as % of GDP.",
    "NE.IMP.GNFS.ZS":      "Imports of goods and services as % of GDP.",
    "BX.KLT.DINV.WD.GD.ZS":"FDI net inflows as % of GDP.",
    "GC.DOD.TOTL.GD.ZS":   "Central + sub-national govt debt as % of GDP.",
    "GC.REV.XGRT.GD.ZS":   "General government revenue as % of GDP.",
    "GC.XPN.TOTL.GD.ZS":   "General government expenditure as % of GDP.",
}

ECB_SOURCE = {
    "source_name": "European Central Bank (ECB)",
    "source_url":  "https://data-api.ecb.europa.eu/service/data/EXR/",
    "notes":       "Statistical Data Warehouse — EXR series. Business days (Mon–Fri) only.",
}


# ── HELPERS ───────────────────────────────────────────────────────────────────
def precise_round(value, precision=4):
    if pd.isna(value):
        return value
    return float(Decimal(str(value)).quantize(
        Decimal("1." + "0" * precision), rounding=ROUND_HALF_UP
    ))


def write_headers(ws, headers, row=1):
    hfill = PatternFill("solid", fgColor=HEADER_COLOR)
    hfont = Font(bold=True, color="FFFFFF")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.fill = hfill
        cell.font = hfont
        cell.alignment = Alignment(horizontal="center")


def style_row(ws, r, ncols, even):
    if even:
        for c in range(1, ncols + 1):
            ws.cell(r, c).fill = PatternFill("solid", fgColor=ALTERNATE_ROW)


def set_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ── WORLD BANK FETCHER (paginated + exponential back-off retry) ───────────────
@st.cache_data(show_spinner=False, ttl=3600)
def fetch_wb_indicator(
    indicator_code: str,
    countries: tuple,
    start_year: int,
    end_year: int,
    max_retries: int = 4,
) -> pd.DataFrame:
    """
    Fetches a WB indicator with full pagination and retry on 5xx / timeout.
    countries = tuple of ISO-2 codes, or ("all",) for every individual country.
    """
    fetch_all   = list(countries) == ["all"]
    country_str = "all" if fetch_all else ";".join(countries)
    url         = f"{WORLD_BANK_BASE}/country/{country_str}/indicator/{indicator_code}"
    page, per_page = 1, 1000
    all_rows: list[dict] = []

    while True:
        params = {
            "date":     f"{start_year}:{end_year}",
            "format":   "json",
            "per_page": per_page,
            "page":     page,
        }
        last_exc = None
        for attempt in range(max_retries):
            try:
                resp = requests.get(url, params=params, timeout=60)
                if resp.status_code in (502, 503, 504):
                    raise requests.HTTPError(f"HTTP {resp.status_code}")
                resp.raise_for_status()
                last_exc = None
                break
            except (requests.RequestException, requests.HTTPError) as exc:
                last_exc = exc
                time.sleep(2 ** attempt)  # 1 s, 2 s, 4 s, 8 s

        if last_exc:
            raise last_exc

        data = resp.json()
        if not isinstance(data, list) or len(data) < 2 or not data[1]:
            break

        for item in data[1]:
            country_id = item.get("country", {}).get("id", "")
            # Skip regional / income-group aggregates (IDs are 3+ chars: EAP, LMC, WLD…)
            if fetch_all and len(country_id) != 2:
                continue
            all_rows.append({
                "Country":      item.get("country", {}).get("value", ""),
                "Country Code": item.get("countryiso3code") or country_id,
                "Year":         int(item["date"]),
                "Value":        item["value"],
            })

        total_pages = data[0].get("pages", 1)
        if page >= total_pages:
            break
        page += 1

    if not all_rows:
        return pd.DataFrame()

    df = pd.DataFrame(all_rows)
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    df = df.dropna(subset=["Value"])
    df = df.sort_values(["Country", "Year"]).reset_index(drop=True)
    return df


@st.cache_data(show_spinner=False, ttl=86400)
def get_wb_countries() -> pd.DataFrame:
    url    = f"{WORLD_BANK_BASE}/country"
    params = {"format": "json", "per_page": 500}
    resp   = requests.get(url, params=params, timeout=20)
    resp.raise_for_status()
    data   = resp.json()
    rows   = [
        {"name": c["name"], "iso2": c["id"]}
        for c in data[1]
        if c.get("region", {}).get("id") != "NA"   # drop aggregates
    ]
    return pd.DataFrame(rows).sort_values("name").reset_index(drop=True)


# ── ECB FETCHERS ──────────────────────────────────────────────────────────────
def fetch_ecb_daily(start_date, end_date) -> pd.DataFrame:
    url    = "https://data-api.ecb.europa.eu/service/data/EXR/D.USD.EUR.SP00.A"
    params = {
        "startPeriod": start_date.strftime("%Y-%m-%d"),
        "endPeriod":   end_date.strftime("%Y-%m-%d"),
        "format":      "csvdata",
    }
    raw = pd.read_csv(StringIO(requests.get(url, params=params, timeout=30).text))
    df  = raw[["TIME_PERIOD", "OBS_VALUE"]].copy()
    df.columns = ["Date", "EUR_USD"]
    df["Date"]    = pd.to_datetime(df["Date"])
    df["EUR_USD"] = pd.to_numeric(df["EUR_USD"], errors="coerce")
    df = df.dropna().sort_values("Date").reset_index(drop=True)
    df["EUR_USD"] = df["EUR_USD"].apply(precise_round)
    df["USD_EUR"] = (1 / df["EUR_USD"]).apply(precise_round)
    return df


def fetch_ecb_monthly(start_date, end_date) -> pd.DataFrame:
    url    = "https://data-api.ecb.europa.eu/service/data/EXR/M.USD.EUR.SP00.A"
    params = {
        "startPeriod": start_date.strftime("%Y-%m"),
        "endPeriod":   end_date.strftime("%Y-%m"),
        "format":      "csvdata",
    }
    raw = pd.read_csv(StringIO(requests.get(url, params=params, timeout=30).text))
    df  = raw[["TIME_PERIOD", "OBS_VALUE"]].copy()
    df.columns = ["Month", "EUR_USD"]
    df["EUR_USD"]    = pd.to_numeric(df["EUR_USD"], errors="coerce")
    df = df.dropna().sort_values("Month").reset_index(drop=True)
    df["Year"]       = df["Month"].str[:4]
    df["Month Name"] = pd.to_datetime(df["Month"]).dt.strftime("%B")
    df["EUR_USD"]    = df["EUR_USD"].apply(precise_round)
    df["USD_EUR"]    = (1 / df["EUR_USD"]).apply(precise_round)
    return df


# ── EXCEL BUILDERS ────────────────────────────────────────────────────────────
def build_fx_sheets(wb: Workbook, start_date, end_date):
    df_d = fetch_ecb_daily(start_date, end_date)

    # ── Annual ──────────────────────────────────────────────────────────────
    ws_a = wb.active
    ws_a.title = "ECB - FX - Annual"
    df_a = df_d.groupby(df_d["Date"].dt.year)["EUR_USD"].mean().reset_index()
    df_a.columns = ["Year", "EUR_USD"]
    df_a["USD_EUR"] = (1 / df_a["EUR_USD"]).apply(precise_round)
    df_a["EUR_USD"] = df_a["EUR_USD"].apply(precise_round)
    write_headers(ws_a, ["Year", "Avg EUR/USD", "Avg USD/EUR"])
    for i, row in df_a.iterrows():
        r = i + 2
        ws_a.cell(r, 1, int(row["Year"]))
        ws_a.cell(r, 2, row["EUR_USD"])
        ws_a.cell(r, 3, row["USD_EUR"])
        style_row(ws_a, r, 3, i % 2 == 0)
    set_widths(ws_a, {"A": 10, "B": 15, "C": 15})

    # ── Monthly ─────────────────────────────────────────────────────────────
    ws_m = wb.create_sheet("ECB - FX - Monthly")
    df_m = fetch_ecb_monthly(start_date, end_date)
    write_headers(ws_m, ["Month", "Year", "Month Name", "EUR/USD", "USD/EUR"])
    for i, row in df_m.iterrows():
        r = i + 2
        ws_m.cell(r, 1, row["Month"])
        ws_m.cell(r, 2, row["Year"])
        ws_m.cell(r, 3, row["Month Name"])
        ws_m.cell(r, 4, row["EUR_USD"])
        ws_m.cell(r, 5, row["USD_EUR"])
        style_row(ws_m, r, 5, i % 2 == 0)
    set_widths(ws_m, {"A": 12, "B": 8, "C": 14, "D": 14, "E": 14})

    # ── Weekly ──────────────────────────────────────────────────────────────
    # ECB publishes Mon–Fri only.
    # "Week Start" = Monday, "Week End" = Friday (Mon + 4 days).
    ws_w = wb.create_sheet("ECB - FX - Weekly")
    df_w = df_d.copy()
    df_w["ISOYear"] = df_w["Date"].dt.isocalendar().year
    df_w["ISOWeek"] = df_w["Date"].dt.isocalendar().week
    df_weekly = df_w.groupby(["ISOYear", "ISOWeek"])["EUR_USD"].mean().reset_index()

    def _monday(row):
        return pd.to_datetime(
            f"{int(row['ISOYear'])}-W{int(row['ISOWeek'])}-1", format="%G-W%V-%u"
        ).strftime("%Y-%m-%d")

    def _friday(row):
        return (
            pd.to_datetime(
                f"{int(row['ISOYear'])}-W{int(row['ISOWeek'])}-1", format="%G-W%V-%u"
            ) + pd.Timedelta(days=4)   # Monday + 4 = Friday
        ).strftime("%Y-%m-%d")

    df_weekly["Week Start (Mon)"] = df_weekly.apply(_monday,  axis=1)
    df_weekly["Week End (Fri)"]   = df_weekly.apply(_friday,  axis=1)
    df_weekly["USD_EUR"]          = (1 / df_weekly["EUR_USD"]).apply(precise_round)
    df_weekly["EUR_USD"]          = df_weekly["EUR_USD"].apply(precise_round)

    write_headers(ws_w, ["Year", "Week #", "Week Start (Mon)", "Week End (Fri)", "EUR/USD", "USD/EUR"])
    for i, row in df_weekly.iterrows():
        r = i + 2
        ws_w.cell(r, 1, int(row["ISOYear"]))
        ws_w.cell(r, 2, int(row["ISOWeek"]))
        ws_w.cell(r, 3, row["Week Start (Mon)"])
        ws_w.cell(r, 4, row["Week End (Fri)"])
        ws_w.cell(r, 5, row["EUR_USD"])
        ws_w.cell(r, 6, row["USD_EUR"])
        style_row(ws_w, r, 6, i % 2 == 0)
    set_widths(ws_w, {"A": 8, "B": 8, "C": 18, "D": 18, "E": 14, "F": 14})

    # ── Daily ────────────────────────────────────────────────────────────────
    ws_d = wb.create_sheet("ECB - FX - Daily")
    write_headers(ws_d, ["Date", "EUR/USD", "USD/EUR"])
    for i, row in df_d.iterrows():
        r = i + 2
        ws_d.cell(r, 1, row["Date"].strftime("%Y-%m-%d"))
        ws_d.cell(r, 2, row["EUR_USD"])
        ws_d.cell(r, 3, row["USD_EUR"])
        style_row(ws_d, r, 3, i % 2 == 0)
    set_widths(ws_d, {"A": 12, "B": 14, "C": 14})


def build_wb_indicator_sheet(
    wb: Workbook,
    label: str,
    indicator_code: str,
    country_codes: list,
    start_year: int,
    end_year: int,
) -> str:
    prefix     = "WB - "
    safe_title = (prefix + label)[:31]
    existing   = {s.title for s in wb.worksheets}
    if safe_title in existing:
        safe_title = safe_title[:28] + "_2"

    ws = wb.create_sheet(safe_title)
    df = fetch_wb_indicator(indicator_code, tuple(country_codes), start_year, end_year)

    if df.empty:
        ws.cell(1, 1, "No data returned by the World Bank API for these parameters.")
        return safe_title

    write_headers(ws, ["Country", "Country Code", "Year", label])
    for i, row in df.iterrows():
        r = i + 2
        ws.cell(r, 1, row["Country"])
        ws.cell(r, 2, row["Country Code"])
        ws.cell(r, 3, int(row["Year"]))
        ws.cell(r, 4, row["Value"])
        style_row(ws, r, 4, i % 2 == 0)
    set_widths(ws, {"A": 28, "B": 14, "C": 8, "D": 26})
    return safe_title


# ── SOURCES SHEET ─────────────────────────────────────────────────────────────
def build_sources_sheet(wb: Workbook, registry: list[dict]):
    ws = wb.create_sheet("Sources", 0)   # always first tab

    title_cell = ws.cell(1, 1, "Data Sources")
    title_cell.font  = Font(bold=True, size=13, color="FFFFFF")
    title_cell.fill  = PatternFill("solid", fgColor="003366")
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24
    ws.merge_cells("A1:E1")

    sub = ws.cell(2, 1, "Copy the sheet name + URL into your footnote / bibliography.")
    sub.font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("A2:E2")

    write_headers(ws, ["Sheet Name", "Dataset", "Source Organisation", "URL", "Notes"], row=3)
    set_widths(ws, {"A": 26, "B": 38, "C": 30, "D": 54, "E": 42})

    link_font = Font(color="1155CC", underline="single")
    for i, entry in enumerate(registry):
        r = i + 4
        ws.cell(r, 1, entry["sheet_name"])
        ws.cell(r, 2, entry["dataset"])
        ws.cell(r, 3, entry["source_name"])
        uc = ws.cell(r, 4, entry["source_url"])
        uc.font = link_font
        ws.cell(r, 5, entry.get("notes", ""))
        style_row(ws, r, 5, i % 2 == 0)


# ═════════════════════════════════════════════════════════════════════════════
# UI
# ═════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🌐 Global Data Downloader</h1>
  <p>Select only the datasets you need — data is fetched on demand and exported to a single Excel file.</p>
</div>
""", unsafe_allow_html=True)


# ── SECTION 1 — EUR/USD ───────────────────────────────────────────────────────
with st.expander("📈  EUR/USD Exchange Rate  (ECB)", expanded=True):
    include_fx = st.checkbox("Include EUR/USD exchange rate data", value=True)
    if include_fx:
        c1, c2 = st.columns(2)
        with c1:
            fx_start = st.date_input("Start date", pd.to_datetime("2015-01-01"), key="fx_start")
        with c2:
            fx_end = st.date_input("End date", pd.to_datetime("today"), key="fx_end")
        if fx_start > fx_end:
            st.error("Start date must be before end date.")
        st.caption("Produces 4 sheets: Annual · Monthly · Weekly (Mon–Fri) · Daily")


# ── SECTION 2 — World Bank ────────────────────────────────────────────────────
with st.expander("🏦  World Bank Indicators", expanded=False):
    include_wb = st.checkbox("Include World Bank data", value=False)

    selected_indicators: list[tuple[str, str]] = []
    selected_countries:  list[str]             = ["all"]
    wb_start_year = 2000
    wb_end_year   = 2023

    if include_wb:
        # ── Indicator picker: one multiselect per themed group ────────────
        st.markdown("**Select indicators** — one Excel sheet per indicator")
        st.markdown(
            '<div class="info-banner">20 indicators across 6 themes. '
            "Expand a group dropdown and tick what you need.</div>",
            unsafe_allow_html=True,
        )

        for group_name, group_dict in WB_INDICATOR_GROUPS.items():
            chosen = st.multiselect(
                group_name,
                options=list(group_dict.keys()),
                default=[],
                key=f"grp_{group_name}",
                placeholder="— none selected —",
            )
            for label in chosen:
                selected_indicators.append((label, group_dict[label]))

        if selected_indicators:
            n = len(selected_indicators)
            names = ", ".join(f"<em>{l}</em>" for l, _ in selected_indicators)
            st.markdown(
                f'<div class="info-banner">✔ <strong>{n} indicator{"s" if n > 1 else ""} selected:</strong> {names}</div>',
                unsafe_allow_html=True,
            )

        # ── Year range ────────────────────────────────────────────────────
        st.markdown("**Year range**")
        y1, y2 = st.columns(2)
        with y1:
            wb_start_year = st.number_input("From year", min_value=1960, max_value=2024, value=2000, step=1)
        with y2:
            wb_end_year   = st.number_input("To year",   min_value=1960, max_value=2024, value=2023, step=1)

        # ── Country picker ────────────────────────────────────────────────
        st.markdown("**Countries** — leave blank to include all countries")
        try:
            all_countries_df = get_wb_countries()
            iso2_map         = dict(zip(all_countries_df["name"], all_countries_df["iso2"]))
            chosen_names     = st.multiselect(
                "Filter by country (optional)",
                options=all_countries_df["name"].tolist(),
                default=[],
                placeholder="All countries  (slower for long date ranges)",
                key="wb_countries",
            )
            selected_countries = [iso2_map[n] for n in chosen_names] if chosen_names else ["all"]
            if not chosen_names:
                st.markdown(
                    '<div class="warn-banner">⚠ Fetching all countries can be slow for long date ranges or many indicators. '
                    "Consider selecting specific countries or a shorter period.</div>",
                    unsafe_allow_html=True,
                )
        except Exception:
            st.warning("Could not load country list — all countries will be fetched.")
            selected_countries = ["all"]


# ── GENERATE ──────────────────────────────────────────────────────────────────
nothing_selected = (not include_fx) and (not include_wb or not selected_indicators)

if nothing_selected:
    st.info("☝ Expand a section above and select at least one dataset to enable the download.")
else:
    if st.button("🚀  Generate & Download Excel"):
        all_ok = True

        if include_fx and fx_start > fx_end:
            st.error("Fix the FX date range first.")
            all_ok = False
        if include_wb and wb_start_year > wb_end_year:
            st.error("'From year' must be ≤ 'To year'.")
            all_ok = False

        if all_ok:
            wb_excel: Workbook          = Workbook()
            sheet_registry: list[dict]  = []
            first_added                 = False
            total_steps = (4 if include_fx else 0) + len(selected_indicators)
            progress    = st.progress(0, text="Starting…")
            step        = 0

            # ── FX sheets ─────────────────────────────────────────────────
            if include_fx:
                progress.progress(0, text="Fetching EUR/USD data from ECB…")
                try:
                    build_fx_sheets(wb_excel, fx_start, fx_end)
                    first_added = True
                    for period in ["Annual", "Monthly", "Weekly", "Daily"]:
                        sheet_registry.append({
                            "sheet_name":  f"ECB - FX - {period}",
                            "dataset":     f"EUR/USD Exchange Rate — {period}",
                            **ECB_SOURCE,
                        })
                    step += 4
                    progress.progress(step / total_steps, text="EUR/USD data loaded ✓")
                except Exception as e:
                    st.error(f"Error fetching FX data: {e}")
                    all_ok = False

            # ── World Bank sheets ─────────────────────────────────────────
            if include_wb and all_ok:
                if not first_added:
                    wb_excel.active.title = "_tmp"
                    first_added = True

                for label, code in selected_indicators:
                    step += 1
                    progress.progress(step / total_steps, text=f"Fetching: {label}…")
                    try:
                        sname = build_wb_indicator_sheet(
                            wb_excel, label, code,
                            selected_countries,
                            int(wb_start_year), int(wb_end_year),
                        )
                        sheet_registry.append({
                            "sheet_name":  sname,
                            "dataset":     label,
                            "source_name": "World Bank Open Data",
                            "source_url":  f"https://data.worldbank.org/indicator/{code}",
                            "notes":       WB_INDICATOR_NOTES.get(code, "World Development Indicators (WDI)."),
                        })
                    except Exception as e:
                        st.warning(f"Could not fetch '{label}': {e}")

            if "_tmp" in wb_excel.sheetnames:
                del wb_excel["_tmp"]

            # ── Sources sheet (tab #1) ────────────────────────────────────
            progress.progress(0.97, text="Writing Sources sheet…")
            build_sources_sheet(wb_excel, sheet_registry)
            progress.progress(1.0, text="Saving…")

            if all_ok:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                wb_excel.save(tmp.name)

                fname_parts = []
                if include_fx:
                    fname_parts.append(f"FX_{fx_start}_to_{fx_end}")
                if include_wb and selected_indicators:
                    fname_parts.append(f"WB_{int(wb_start_year)}-{int(wb_end_year)}")
                filename = "_".join(fname_parts) + ".xlsx"

                with open(tmp.name, "rb") as f:
                    st.download_button(
                        "📥  Download Excel",
                        f,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                progress.empty()
                n = len(wb_excel.sheetnames)
                st.success(
                    f"✅ Done! {n} sheet{'s' if n != 1 else ''} exported "
                    f"(Sources tab included)."
                )
