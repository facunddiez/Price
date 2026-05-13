import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import openpyxl
from datetime import datetime
import requests
import io

WSD_URL  = "https://docs.google.com/spreadsheets/d/1zfRWlDKS--rTqD4qcK2VJFgaQE2iQ8vu/export?format=xlsx"
IMHS_URL = "https://docs.google.com/spreadsheets/d/1T3nl9w5DpT3yOVVQOhzZpamA0qVGQA_n/export?format=xlsx"
ZCHS_URL = "https://docs.google.com/spreadsheets/d/1sqyx9DUO3Dhbe51hPP5YRTNhCutY1b6_/export?format=xlsx"

def fetch_wb(url):
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)

st.set_page_config(
    page_title="Worldsprings Pricing Dashboard",
    page_icon="♨️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #1a1a2e; }
[data-testid="stSidebar"] * { color: #e0e0e0 !important; }
.metric-card {
    background: linear-gradient(135deg, #1e3a5f, #2d5a8e);
    border-radius: 12px; padding: 20px; margin: 8px 0;
    border-left: 4px solid #4fc3f7;
}
.metric-card h4 { color: #4fc3f7; margin: 0 0 8px 0; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; }
.metric-card .price { color: #ffffff; font-size: 28px; font-weight: 700; }
.metric-card .sub { color: #90caf9; font-size: 12px; margin-top: 4px; }
.location-header { font-size: 22px; font-weight: 700; padding: 4px 0 12px 0; }
.wsd-color { color: #ff6b6b; }
.imhs-color { color: #4ecdc4; }
.zchs-color { color: #ffe66d; }
</style>
""", unsafe_allow_html=True)

# ── data loaders ──────────────────────────────────────────────────────────────
@st.cache_data
def load_wsd():
    wb = fetch_wb(WSD_URL)

    # current rates
    ws = wb["2026 rates "]
    rows = list(ws.iter_rows(values_only=True))

    current = {
        "Day Pass":     {"Mon-Thu": 85, "Fri & Sun": 95, "Saturday": 105},
        "Evening Pass": {"Mon-Thu": 55, "Fri & Sun": 65, "Saturday": 85},
    }

    spa_current = []
    for row in rows[1:]:
        if row[0] == "Spa " and row[1] and row[2] is not None:
            try:
                spa_current.append({
                    "Service": row[1],
                    "Mon-Thu": row[2] if isinstance(row[2], (int, float)) else None,
                    "Fri & Sun": row[3] if isinstance(row[3], (int, float)) else None,
                    "Saturday": row[4] if isinstance(row[4], (int, float)) else None,
                })
            except Exception:
                pass

    # historical from 2025 vs 2026 sheet
    ws2 = wb["2025 vs 2026"]
    history_rows = list(ws2.iter_rows(values_only=True))
    history = []
    periods = [
        ("Jul-Sep 2024",  "3 Hour Soak", 55, 75),
        ("Oct-Jan 2025",  "3 Hour Soak", 59, 85),
        ("Feb-May 2025",  "3 Hour Soak", 49, 69),
        ("Jun-Dec 2025",  "3 Hour Soak", 69, 79),
        ("Jan-Feb 2026",  "Day Pass",    85, 95),
    ]
    for label, prod, mth, frs in periods:
        history.append({"Period": label, "Product": prod, "Mon-Thu": mth, "Fri-Sun": frs})

    quick_dip_hist = [
        ("Jul-Sep 2024",  35, 35),
        ("Oct-Jan 2025",  39, 39),
        ("Feb-May 2025",  39, 39),
        ("Jun-Dec 2025",  49, 59),
        ("Jan-Feb 2026",  55, 65),
    ]
    for label, mth, frs in quick_dip_hist:
        history.append({"Period": label, "Product": "Quick Dip / Evening", "Mon-Thu": mth, "Fri-Sun": frs})

    df_hist = pd.DataFrame(history)
    df_spa = pd.DataFrame(spa_current)
    return current, df_hist, df_spa


@st.cache_data
def load_imhs():
    wb = fetch_wb(IMHS_URL)

    # latest pricing – use last sheet
    latest_sheet = "Pricing 3.30.26-5.31.26"
    ws = wb[latest_sheet]
    rows = list(ws.iter_rows(values_only=True))
    # row[2] = label, [3]=Select3h, [4]=SelectAll, [5]=Premier3h, [6]=PremierAll
    non_peak = {}
    peak = {}
    section = None
    for row in rows:
        label = row[2] if len(row) > 2 else None
        if label == "Non-Peak Pricing":
            section = "non_peak"
        elif label == "Holiday/Peak Pricing":
            section = "peak"
        elif label in ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun") and section:
            d = {
                "Select 3hr": row[3], "Select All-Day": row[4],
                "Premier 3hr": row[5], "Premier All-Day": row[6],
            }
            if section == "non_peak":
                non_peak[label] = d
            else:
                peak[label] = d

    # historical – parse key sheets
    sheet_ranges = {
        "5.23.22": ("5.23.2022-1.22.2023", "5.23.2022"),
        "1.23.23": ("1.23.2023-4.30.2023", "1.23.2023"),
        "5.1.23":  ("5.1.2023-9.4.2023",   "5.1.2023"),
        "9.5.23":  ("9.5.2023-10.31.2023",  "9.5.2023"),
        "11.1.23": ("11.1.2023-12.21.2023", "11.1.2023"),
        "12.22.23":("12.22.2023-5.23.2024","12.22.2023"),
        "5.24.24": ("5.24.2024-11.3.2024",  "5.24.2024"),
        "11.4.24": ("11.4.24-2.28.25",      "11.4.2024"),
        "3.1.25":  ("3.1.25-6.30.25",       "3.1.2025"),
        "7.1.25":  ("7.1.25-8.30.25",       "7.1.2025"),
        "9.1.25":  ("9.1.25-9.30.25",       "9.1.2025"),
        "10.1.25": ("10.1.25 - 2.1.26",     "10.1.2025"),
        "2.2.26":  ("Pricing 2.2.26-3.29.26","2.2.2026"),
        "3.30.26": ("Pricing 3.30.26-5.31.26","3.30.2026"),
    }

    hist_records = []
    for label, (sheet_name, date_str) in sheet_ranges.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws_h = wb[sheet_name]
        rows_h = list(ws_h.iter_rows(values_only=True))
        sec = None
        for row in rows_h:
            lbl = row[2] if len(row) > 2 else None
            if lbl == "Non-Peak Pricing":
                sec = "non_peak"
            elif lbl == "Holiday/Peak Pricing":
                sec = "peak"
            elif lbl == "Mon" and sec == "non_peak":
                try:
                    hist_records.append({
                        "Period": label,
                        "Date": date_str,
                        "Type": "Non-Peak Mon",
                        "Select 3hr": row[3],
                        "Select All-Day": row[4],
                        "Premier 3hr": row[5],
                        "Premier All-Day": row[6],
                    })
                except Exception:
                    pass
            elif lbl == "Sat" and sec == "peak":
                try:
                    hist_records.append({
                        "Period": label,
                        "Date": date_str,
                        "Type": "Peak Sat",
                        "Select 3hr": row[3],
                        "Select All-Day": row[4],
                        "Premier 3hr": row[5],
                        "Premier All-Day": row[6],
                    })
                except Exception:
                    pass

    # pre-2022 from year-over-year sheet
    ws_yoy = wb["2015 - Present DAy"]
    rows_yoy = list(ws_yoy.iter_rows(values_only=True))
    # columns: 0=label, 1=2015, 2=2016, 3=2017, 4=2018, 5=2019, 6=2020, 7=2021, 8=11/21, 9=12/21, 10=1/22
    years = [("2015","2015"), ("2016","2016"), ("2017","2017"), ("2018","2018"),
             ("2019","2019"), ("2020","2020"), ("2021","2021")]
    adult_peak_row = None
    for row in rows_yoy:
        if row[0] == "Everyday Adult" and adult_peak_row is None:
            adult_peak_row = row
            break

    if adult_peak_row:
        for i, (yr, date_s) in enumerate(years):
            val = adult_peak_row[i + 1]
            if val is not None:
                hist_records.append({
                    "Period": yr,
                    "Date": f"6.1.{yr}",
                    "Type": "Peak (Legacy Adult)",
                    "Select 3hr": val,
                    "Select All-Day": val,
                    "Premier 3hr": None,
                    "Premier All-Day": None,
                })

    df_hist = pd.DataFrame(hist_records)
    return non_peak, peak, df_hist


@st.cache_data
def load_zchs():
    wb = fetch_wb(ZCHS_URL)

    # current = Pricing 11_2025 (latest)
    ws = wb["Pricing 11_2025"]
    rows = list(ws.iter_rows(values_only=True))

    current = {}
    for row in rows:
        label = str(row[0]).strip() if row[0] else ""
        if "Quick Dip Adult / Teen" in label and "Select" in label:
            current["Quick Dip Select (13+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "Quick Dip Adult: Premier" in label:
            current["Quick Dip Premier (21+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "3-hour Adult / Teen" in label and "Select" in label and "Eat" not in label:
            if "3hr Select (13+)" not in current:
                current["3hr Select (13+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "3- Hour Soak Adult: Premier" in label:
            if "3hr Premier (21+)" not in current:
                current["3hr Premier (21+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "Full Day Soak Ages 13+" in label:
            current["Full Day Select (13+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "Full Day Adult: Premier" in label:
            current["Full Day Premier (21+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}

    # historical across pricing sheets
    pricing_sheets = {
        "Jan 2025": ("Pricing 2_3_2025", "1.22.2025"),
        "Jul 2025": ("Pricing 07_2025", "7.1.2025"),
        "Nov 2025": ("Pricing 11_2025", "11.1.2025"),
    }
    hist_records = []
    for period, (sheet_name, date_s) in pricing_sheets.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws_h = wb[sheet_name]
        rows_h = list(ws_h.iter_rows(values_only=True))
        for row in rows_h:
            lbl = str(row[0]).strip() if row[0] else ""
            if "3-hour Adult / Teen" in lbl and "Select" in lbl and "Eat" not in lbl:
                try:
                    v1 = row[1] if isinstance(row[1], (int, float)) else None
                    v2 = row[2] if isinstance(row[2], (int, float)) else None
                    if v1:
                        hist_records.append({"Period": period, "Date": date_s, "Product": "3hr Select (13+)", "Mon-Thu": v1, "Fri/Sat/Sun": v2})
                except Exception:
                    pass
            elif "3- Hour Soak Adult: Premier" in lbl:
                try:
                    v1 = row[1] if isinstance(row[1], (int, float)) else None
                    v2 = row[2] if isinstance(row[2], (int, float)) else None
                    if v1:
                        hist_records.append({"Period": period, "Date": date_s, "Product": "3hr Premier (21+)", "Mon-Thu": v1, "Fri/Sat/Sun": v2})
                except Exception:
                    pass
            elif "Full Day Soak Ages 13+" in lbl:
                try:
                    v1 = row[1] if isinstance(row[1], (int, float)) else None
                    v2 = row[2] if isinstance(row[2], (int, float)) else None
                    if v1:
                        hist_records.append({"Period": period, "Date": date_s, "Product": "Full Day Select (13+)", "Mon-Thu": v1, "Fri/Sat/Sun": v2})
                except Exception:
                    pass

    df_hist = pd.DataFrame(hist_records)
    return current, df_hist


# ── nav ───────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ♨️ Worldsprings")
    st.markdown("---")
    page = st.radio(
        "Navigate",
        ["Overview", "WSD – Dallas", "IMHS – Colorado", "ZCHS – Utah"],
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.caption("Data source: Excel pricing files")


# ── helpers ───────────────────────────────────────────────────────────────────
DAY_ORDER = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
COLORS = {"WSD": "#ff6b6b", "IMHS": "#4ecdc4", "ZCHS": "#ffe66d"}


def price_table(data: dict, title: str, color: str):
    """Render a compact pricing grid."""
    rows_html = ""
    for product, prices in data.items():
        cols = "".join(f"<td style='padding:6px 14px;text-align:center;'>${v}</td>"
                       if isinstance(v, (int, float)) else
                       f"<td style='padding:6px 14px;text-align:center;color:#888'>{v}</td>"
                       for v in prices.values())
        rows_html += f"<tr><td style='padding:6px 14px;font-weight:600'>{product}</td>{cols}</tr>"

    headers = "".join(f"<th style='padding:6px 14px;background:{color}22;color:{color}'>{h}</th>"
                      for h in next(iter(data.values())).keys())
    table = f"""
    <h4 style='color:{color};margin:16px 0 6px'>{title}</h4>
    <table style='width:100%;border-collapse:collapse;font-size:14px;'>
      <thead><tr><th style='padding:6px 14px;background:{color}22;color:{color}'>Product</th>{headers}</tr></thead>
      <tbody>{rows_html}</tbody>
    </table>"""
    st.markdown(table, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
if page == "Overview":
    wsd_current, _, _ = load_wsd()
    imhs_non_peak, imhs_peak, _ = load_imhs()
    zchs_current, _ = load_zchs()

    st.title("♨️ Worldsprings — Pricing Overview")
    st.caption(f"As of {datetime.today().strftime('%B %d, %Y')}")
    st.markdown("---")

    col1, col2, col3 = st.columns(3)

    # ── WSD ──────────────────────────────────────────────────────────────────
    with col1:
        st.markdown("<div class='location-header wsd-color'>🔴 WSD – Dallas, TX</div>", unsafe_allow_html=True)
        price_table(wsd_current, "Soak Admission (2026)", COLORS["WSD"])

        st.markdown(f"""
        <h4 style='color:{COLORS["WSD"]};margin:16px 0 6px'>Key Add-ons</h4>
        <table style='width:100%;border-collapse:collapse;font-size:14px;'>
          <tr><td style='padding:5px 10px'>Robe Rental</td><td style='padding:5px 10px;text-align:right'>$10</td></tr>
          <tr><td style='padding:5px 10px'>Cabana (Mon-Thu)</td><td style='padding:5px 10px;text-align:right'>$199</td></tr>
          <tr><td style='padding:5px 10px'>Cabana (Sat)</td><td style='padding:5px 10px;text-align:right'>$299</td></tr>
          <tr><td style='padding:5px 10px'>Day Pass E-Cert (gift)</td><td style='padding:5px 10px;text-align:right'>$95</td></tr>
        </table>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <h4 style='color:{COLORS["WSD"]};margin:16px 0 6px'>Memberships</h4>
        <table style='width:100%;border-collapse:collapse;font-size:14px;'>
          <tr><td style='padding:5px 10px'>1-Month Trial</td><td style='padding:5px 10px;text-align:right'>$125</td></tr>
          <tr><td style='padding:5px 10px'>Weekday Annual</td><td style='padding:5px 10px;text-align:right'>$948/yr</td></tr>
          <tr><td style='padding:5px 10px'>Anytime Annual</td><td style='padding:5px 10px;text-align:right'>$1,788/yr</td></tr>
        </table>
        """, unsafe_allow_html=True)

    # ── IMHS ─────────────────────────────────────────────────────────────────
    with col2:
        st.markdown("<div class='location-header imhs-color'>🟢 IMHS – Glenwood Springs, CO</div>", unsafe_allow_html=True)

        np_table = {
            day: {
                "Sel 3hr": imhs_non_peak.get(day, {}).get("Select 3hr"),
                "Sel All-Day": imhs_non_peak.get(day, {}).get("Select All-Day"),
                "Pre 3hr": imhs_non_peak.get(day, {}).get("Premier 3hr"),
                "Pre All-Day": imhs_non_peak.get(day, {}).get("Premier All-Day"),
            }
            for day in DAY_ORDER if day in imhs_non_peak
        }
        price_table(np_table, "Non-Peak Pricing (current)", COLORS["IMHS"])

        pk_table = {
            day: {
                "Sel 3hr": imhs_peak.get(day, {}).get("Select 3hr"),
                "Sel All-Day": imhs_peak.get(day, {}).get("Select All-Day"),
                "Pre 3hr": imhs_peak.get(day, {}).get("Premier 3hr"),
                "Pre All-Day": imhs_peak.get(day, {}).get("Premier All-Day"),
            }
            for day in DAY_ORDER if day in imhs_peak
        }
        price_table(pk_table, "Peak / Holiday Pricing", COLORS["IMHS"])

    # ── ZCHS ─────────────────────────────────────────────────────────────────
    with col3:
        st.markdown("<div class='location-header zchs-color'>🟡 ZCHS – Zion Canyon, UT</div>", unsafe_allow_html=True)

        if zchs_current:
            zchs_display = {}
            for product, prices in zchs_current.items():
                safe = {}
                for k, v in prices.items():
                    safe[k] = v if isinstance(v, (int, float)) else "—"
                zchs_display[product] = safe
            price_table(zchs_display, "Soak Pricing (Nov 2025 rates)", COLORS["ZCHS"])

        st.markdown(f"""
        <h4 style='color:{COLORS["ZCHS"]};margin:16px 0 6px'>Add-ons</h4>
        <table style='width:100%;border-collapse:collapse;font-size:14px;'>
          <tr><td style='padding:5px 10px'>Robe</td><td style='padding:5px 10px;text-align:right'>$10</td></tr>
          <tr><td style='padding:5px 10px'>Single Cabana (Mon-Thu)</td><td style='padding:5px 10px;text-align:right'>$149</td></tr>
          <tr><td style='padding:5px 10px'>Single Cabana (Fri-Sun)</td><td style='padding:5px 10px;text-align:right'>$249</td></tr>
          <tr><td style='padding:5px 10px'>Double Cabana (Mon-Thu)</td><td style='padding:5px 10px;text-align:right'>$199</td></tr>
          <tr><td style='padding:5px 10px'>Double Cabana (Fri-Sun)</td><td style='padding:5px 10px;text-align:right'>$299</td></tr>
        </table>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <h4 style='color:{COLORS["ZCHS"]};margin:16px 0 6px'>Memberships</h4>
        <table style='width:100%;border-collapse:collapse;font-size:14px;'>
          <tr><td style='padding:5px 10px'>Anytime Annual</td><td style='padding:5px 10px;text-align:right'>$1,499</td></tr>
          <tr><td style='padding:5px 10px'>Weekday Annual</td><td style='padding:5px 10px;text-align:right'>$920</td></tr>
          <tr><td style='padding:5px 10px'>Snowbird (3 mo)</td><td style='padding:5px 10px;text-align:right'>$499</td></tr>
        </table>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("Comparable: Standard Soak Entry — Weekday vs Weekend")

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="WSD Day Pass (Mon-Thu)",
        x=["WSD Dallas"], y=[85], marker_color=COLORS["WSD"],
        text=["$85"], textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="WSD Day Pass (Sat)",
        x=["WSD Dallas"], y=[105], marker_color=COLORS["WSD"],
        opacity=0.6, text=["$105"], textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="IMHS Select All-Day (Mon)",
        x=["IMHS Colorado"], y=[imhs_non_peak.get("Mon", {}).get("Select All-Day", 0)],
        marker_color=COLORS["IMHS"],
        text=[f"${imhs_non_peak.get('Mon', {}).get('Select All-Day', 0)}"], textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="IMHS Premier All-Day (Sat peak)",
        x=["IMHS Colorado"], y=[imhs_peak.get("Sat", {}).get("Premier All-Day", 0)],
        marker_color=COLORS["IMHS"], opacity=0.6,
        text=[f"${imhs_peak.get('Sat', {}).get('Premier All-Day', 0)}"], textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="ZCHS 3hr Select (Mon-Thu)",
        x=["ZCHS Utah"], y=[zchs_current.get("3hr Select (13+)", {}).get("Mon-Thu", 0)],
        marker_color=COLORS["ZCHS"],
        text=[f"${zchs_current.get('3hr Select (13+)', {}).get('Mon-Thu', 0)}"], textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="ZCHS 3hr Select (Fri-Sun)",
        x=["ZCHS Utah"], y=[zchs_current.get("3hr Select (13+)", {}).get("Fri/Sat/Sun", 0)],
        marker_color=COLORS["ZCHS"], opacity=0.6,
        text=[f"${zchs_current.get('3hr Select (13+)', {}).get('Fri/Sat/Sun', 0)}"], textposition="outside",
    ))

    fig.update_layout(
        barmode="group",
        plot_bgcolor="#0e1117", paper_bgcolor="#0e1117",
        font_color="white", height=380,
        legend=dict(orientation="h", y=-0.25),
        yaxis=dict(title="Price (USD)", gridcolor="#333"),
        margin=dict(t=20, b=20),
    )
    st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: WSD
# ══════════════════════════════════════════════════════════════════════════════
elif page == "WSD – Dallas":
    wsd_current, wsd_hist, wsd_spa = load_wsd()

    st.title("🔴 WSD — Worldsprings Dallas")
    st.markdown("---")

    tab1, tab2, tab3 = st.tabs(["Current Rates", "Price History", "Spa Menu"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            price_table(wsd_current, "Soak Admission 2026", COLORS["WSD"])

            st.markdown(f"""
            <h4 style='color:{COLORS["WSD"]};margin:20px 0 6px'>Memberships</h4>
            <table style='width:100%;border-collapse:collapse;font-size:14px;'>
              <tr><td style='padding:5px 10px'>1-Month Trial (4 day passes)</td><td style='padding:5px 10px;text-align:right;font-weight:700'>$125</td></tr>
              <tr><td style='padding:5px 10px'>Weekday Annual (unlimited M-F + 5 guest passes)</td><td style='padding:5px 10px;text-align:right;font-weight:700'>$948/yr</td></tr>
              <tr><td style='padding:5px 10px'>Anytime Annual (unlimited + 10 guest passes)</td><td style='padding:5px 10px;text-align:right;font-weight:700'>$1,788/yr</td></tr>
            </table>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <h4 style='color:{COLORS["WSD"]};margin:0 0 6px'>Add-ons</h4>
            <table style='width:100%;border-collapse:collapse;font-size:14px;'>
              <tr><td style='padding:5px 10px'>Robe Rental</td><td style='padding:5px 10px;text-align:right'>$10</td></tr>
              <tr><td style='padding:5px 10px'>Cabana (Mon-Thu)</td><td style='padding:5px 10px;text-align:right'>$199</td></tr>
              <tr><td style='padding:5px 10px'>Cabana (Sat/Holiday)</td><td style='padding:5px 10px;text-align:right'>$299</td></tr>
              <tr><td style='padding:5px 10px'>Day Pass E-Certificate</td><td style='padding:5px 10px;text-align:right'>$95</td></tr>
            </table>

            <h4 style='color:{COLORS["WSD"]};margin:20px 0 6px'>Complimentary Classes</h4>
            <table style='width:100%;border-collapse:collapse;font-size:14px;'>
              <tr><td style='padding:5px 10px'>Hot Springs Sound Bath (30 min)</td><td style='padding:5px 10px;text-align:right;color:#4fc3f7'>FREE</td></tr>
              <tr><td style='padding:5px 10px'>Guided Contrast Therapy (30 min)</td><td style='padding:5px 10px;text-align:right;color:#4fc3f7'>FREE</td></tr>
              <tr><td style='padding:5px 10px'>Stretch, Sauna & Breathwork (30 min)</td><td style='padding:5px 10px;text-align:right;color:#4fc3f7'>FREE</td></tr>
              <tr><td style='padding:5px 10px'>Cold Plunge 101 (10 min)</td><td style='padding:5px 10px;text-align:right;color:#4fc3f7'>FREE</td></tr>
            </table>
            """, unsafe_allow_html=True)

    with tab2:
        st.subheader("Soak Price History — WSD")
        st.caption("Key products compared across pricing periods")

        day_pass = wsd_hist[wsd_hist["Product"].isin(["3 Hour Soak", "Day Pass"])].copy()
        evening = wsd_hist[wsd_hist["Product"] == "Quick Dip / Evening"].copy()

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=day_pass["Period"], y=day_pass["Mon-Thu"],
            mode="lines+markers", name="Day Pass / 3hr (Mon-Thu)",
            line=dict(color=COLORS["WSD"], width=2),
            marker=dict(size=8),
        ))
        fig.add_trace(go.Scatter(
            x=day_pass["Period"], y=day_pass["Fri-Sun"],
            mode="lines+markers", name="Day Pass / 3hr (Fri-Sun)",
            line=dict(color=COLORS["WSD"], width=2, dash="dash"),
            marker=dict(size=8),
        ))
        fig.add_trace(go.Scatter(
            x=evening["Period"], y=evening["Mon-Thu"],
            mode="lines+markers", name="Evening / Quick Dip (Mon-Thu)",
            line=dict(color="#ffa07a", width=2),
            marker=dict(size=8),
        ))
        fig.add_trace(go.Scatter(
            x=evening["Period"], y=evening["Fri-Sun"],
            mode="lines+markers", name="Evening / Quick Dip (Fri-Sun)",
            line=dict(color="#ffa07a", width=2, dash="dash"),
            marker=dict(size=8),
        ))

        fig.update_layout(
            plot_bgcolor="#0e1117", paper_bgcolor="#0e1117",
            font_color="white", height=420,
            legend=dict(orientation="h", y=-0.25),
            yaxis=dict(title="Price (USD)", gridcolor="#333"),
            xaxis=dict(title="Pricing Period"),
            margin=dict(t=20, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)

        st.caption("Note: WSD launched in 2024. 2026 rates include restructured product names (Day Pass / Evening Pass).")
        st.dataframe(wsd_hist.set_index("Period"), use_container_width=True)

    with tab3:
        st.subheader("Spa Services — WSD (2026 Rates)")
        st.caption("All 50-min services include Day Pass & robe unless noted.")
        if not wsd_spa.empty:
            wsd_spa_disp = wsd_spa.copy()
            for col in ["Mon-Thu", "Fri & Sun", "Saturday"]:
                wsd_spa_disp[col] = wsd_spa_disp[col].apply(
                    lambda v: f"${v:,.0f}" if isinstance(v, (int, float)) else (str(v) if v else "—")
                )
            st.dataframe(wsd_spa_disp.set_index("Service"), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: IMHS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "IMHS – Colorado":
    imhs_non_peak, imhs_peak, imhs_hist = load_imhs()

    st.title("🟢 IMHS — Iron Mountain Hot Springs, CO")
    st.markdown("---")

    tab1, tab2 = st.tabs(["Current Rates", "Price History (2015–Present)"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            np_data = {
                day: {
                    "Select 3hr": f"${imhs_non_peak[day]['Select 3hr']}",
                    "Select All-Day": f"${imhs_non_peak[day]['Select All-Day']}",
                    "Premier 3hr": f"${imhs_non_peak[day]['Premier 3hr']}",
                    "Premier All-Day": f"${imhs_non_peak[day]['Premier All-Day']}",
                }
                for day in DAY_ORDER if day in imhs_non_peak
            }
            price_table(np_data, "Non-Peak Pricing", COLORS["IMHS"])

        with col2:
            pk_data = {
                day: {
                    "Select 3hr": f"${imhs_peak[day]['Select 3hr']}",
                    "Select All-Day": f"${imhs_peak[day]['Select All-Day']}",
                    "Premier 3hr": f"${imhs_peak[day]['Premier 3hr']}",
                    "Premier All-Day": f"${imhs_peak[day]['Premier All-Day']}",
                }
                for day in DAY_ORDER if day in imhs_peak
            }
            price_table(pk_data, "Peak / Holiday Pricing", COLORS["IMHS"])

        st.markdown("---")
        st.markdown(f"""
        <h4 style='color:{COLORS["IMHS"]};margin:0 0 10px'>Product Types</h4>
        <table style='width:100%;border-collapse:collapse;font-size:14px;'>
          <tr style='background:#1a3a2a'>
            <th style='padding:8px 12px;color:{COLORS["IMHS"]}'>Tier</th>
            <th style='padding:8px 12px;color:{COLORS["IMHS"]}'>Access</th>
          </tr>
          <tr>
            <td style='padding:7px 12px;font-weight:600'>Select 3hr</td>
            <td style='padding:7px 12px'>3-hour timed soak, select pools</td>
          </tr>
          <tr>
            <td style='padding:7px 12px;font-weight:600'>Select All-Day</td>
            <td style='padding:7px 12px'>Unlimited duration, select pools</td>
          </tr>
          <tr>
            <td style='padding:7px 12px;font-weight:600'>Premier 3hr</td>
            <td style='padding:7px 12px'>3-hour timed soak, all pools including premier</td>
          </tr>
          <tr>
            <td style='padding:7px 12px;font-weight:600'>Premier All-Day</td>
            <td style='padding:7px 12px'>Unlimited duration, all pools including premier</td>
          </tr>
        </table>
        """, unsafe_allow_html=True)

    with tab2:
        st.subheader("IMHS Price History — 2015 to Present")

        type_filter = st.selectbox(
            "Day type",
            options=imhs_hist["Type"].unique().tolist() if not imhs_hist.empty else [],
        )

        if not imhs_hist.empty:
            filtered = imhs_hist[imhs_hist["Type"] == type_filter].copy()

            fig = go.Figure()
            for metric, color in [
                ("Select 3hr", COLORS["IMHS"]),
                ("Select All-Day", "#26a69a"),
                ("Premier 3hr", "#80cbc4"),
                ("Premier All-Day", "#b2dfdb"),
            ]:
                df_m = filtered.dropna(subset=[metric])
                if not df_m.empty:
                    fig.add_trace(go.Scatter(
                        x=df_m["Period"], y=df_m[metric],
                        mode="lines+markers", name=metric,
                        line=dict(color=color, width=2),
                        marker=dict(size=8),
                    ))

            fig.update_layout(
                plot_bgcolor="#0e1117", paper_bgcolor="#0e1117",
                font_color="white", height=440,
                legend=dict(orientation="h", y=-0.25),
                yaxis=dict(title="Price (USD)", gridcolor="#333"),
                xaxis=dict(title="Period start"),
                margin=dict(t=20, b=20),
            )
            st.plotly_chart(fig, use_container_width=True)

            st.caption("Pre-2022 data uses the legacy 'Everyday Adult' rate (no Select/Premier split). Post-2022 reflects new tiered pricing structure.")
            st.dataframe(filtered.set_index("Period"), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: ZCHS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "ZCHS – Utah":
    zchs_current, zchs_hist = load_zchs()

    st.title("🟡 ZCHS — Zion Canyon Hot Springs, UT")
    st.markdown("---")

    tab1, tab2 = st.tabs(["Current Rates", "Price History"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            if zchs_current:
                display = {}
                for product, prices in zchs_current.items():
                    display[product] = {
                        k: (f"${v}" if isinstance(v, (int, float)) else "—")
                        for k, v in prices.items()
                    }
                price_table(display, "Soak Pricing (Nov 2025)", COLORS["ZCHS"])

        with col2:
            st.markdown(f"""
            <h4 style='color:{COLORS["ZCHS"]};margin:0 0 6px'>Add-ons</h4>
            <table style='width:100%;border-collapse:collapse;font-size:14px;'>
              <tr><td style='padding:5px 10px'>Robe</td><td style='padding:5px 10px;text-align:right'>$10</td></tr>
              <tr><td style='padding:5px 10px'>Single Cabana (Mon-Thu)</td><td style='padding:5px 10px;text-align:right'>$149</td></tr>
              <tr><td style='padding:5px 10px'>Single Cabana (Fri-Sun)</td><td style='padding:5px 10px;text-align:right'>$249</td></tr>
              <tr><td style='padding:5px 10px'>Double Cabana (Mon-Thu)</td><td style='padding:5px 10px;text-align:right'>$199</td></tr>
              <tr><td style='padding:5px 10px'>Double Cabana (Fri-Sun)</td><td style='padding:5px 10px;text-align:right'>$299</td></tr>
            </table>

            <h4 style='color:{COLORS["ZCHS"]};margin:20px 0 6px'>Memberships</h4>
            <table style='width:100%;border-collapse:collapse;font-size:14px;'>
              <tr><td style='padding:5px 10px'>Anytime Annual</td><td style='padding:5px 10px;text-align:right;font-weight:700'>$1,499</td></tr>
              <tr><td style='padding:5px 10px'>Weekday Annual (M-F)</td><td style='padding:5px 10px;text-align:right;font-weight:700'>$920</td></tr>
              <tr><td style='padding:5px 10px'>Snowbird (3 months)</td><td style='padding:5px 10px;text-align:right;font-weight:700'>$499</td></tr>
            </table>

            <h4 style='color:{COLORS["ZCHS"]};margin:20px 0 6px'>Special Rates</h4>
            <table style='width:100%;border-collapse:collapse;font-size:14px;'>
              <tr><td style='padding:5px 10px'>Washington County Residents</td><td style='padding:5px 10px;text-align:right'>20% off</td></tr>
              <tr><td style='padding:5px 10px'>Hotel Partner Rate</td><td style='padding:5px 10px;text-align:right'>10% off</td></tr>
            </table>
            """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown(f"""
        <h4 style='color:{COLORS["ZCHS"]};margin:0 0 10px'>Access Tiers</h4>
        <table style='width:100%;border-collapse:collapse;font-size:14px;'>
          <tr style='background:#2a2a1a'>
            <th style='padding:8px 12px;color:{COLORS["ZCHS"]}'>Tier</th>
            <th style='padding:8px 12px;color:{COLORS["ZCHS"]}'>Access</th>
          </tr>
          <tr><td style='padding:7px 12px;font-weight:600'>Select (13+)</td><td style='padding:7px 12px'>Standard pools, teens & adults</td></tr>
          <tr><td style='padding:7px 12px;font-weight:600'>Select Youth (3–12)</td><td style='padding:7px 12px'>Standard pools, children</td></tr>
          <tr><td style='padding:7px 12px;font-weight:600'>Premier (21+)</td><td style='padding:7px 12px'>All pools including 21+ only areas</td></tr>
        </table>
        """, unsafe_allow_html=True)

    with tab2:
        st.subheader("ZCHS Price History — Across Pricing Periods")
        st.caption("ZCHS opened January 2025. Three pricing periods captured.")

        if not zchs_hist.empty:
            product_filter = st.selectbox("Product", options=zchs_hist["Product"].unique().tolist())
            filtered = zchs_hist[zchs_hist["Product"] == product_filter]

            fig = go.Figure()
            fig.add_trace(go.Bar(
                name="Mon-Thu", x=filtered["Period"], y=filtered["Mon-Thu"],
                marker_color=COLORS["ZCHS"], text=filtered["Mon-Thu"].apply(lambda v: f"${v}" if pd.notna(v) else ""),
                textposition="outside",
            ))
            fig.add_trace(go.Bar(
                name="Fri/Sat/Sun", x=filtered["Period"], y=filtered["Fri/Sat/Sun"],
                marker_color="#c8a800", text=filtered["Fri/Sat/Sun"].apply(lambda v: f"${v}" if pd.notna(v) else ""),
                textposition="outside",
            ))

            fig.update_layout(
                barmode="group",
                plot_bgcolor="#0e1117", paper_bgcolor="#0e1117",
                font_color="white", height=420,
                legend=dict(orientation="h", y=-0.25),
                yaxis=dict(title="Price (USD)", gridcolor="#333"),
                margin=dict(t=30, b=20),
            )
            st.plotly_chart(fig, use_container_width=True)

            st.dataframe(filtered.set_index("Period"), use_container_width=True)
        else:
            st.info("No historical data parsed yet.")
