import streamlit as st
import pandas as pd
import plotly.graph_objects as go
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

# ── color palette ─────────────────────────────────────────────────────────────
C = {
    "WSD":  {"p": "#38BDF8", "light": "#BAE6FD", "dark": "#0369A1", "bg": "#051525", "glow": "#38BDF820"},
    "IMHS": {"p": "#4ADE80", "light": "#BBF7D0", "dark": "#15803D", "bg": "#051505", "glow": "#4ADE8020"},
    "ZCHS": {"p": "#F87171", "light": "#FECACA", "dark": "#B91C1C", "bg": "#1a0505", "glow": "#F8717120"},
    "base": "#0D1117",
}

DAY_ORDER = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

# ── base CSS ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

* { font-family: 'Inter', sans-serif !important; }

/* App background */
.stApp { background-color: #0D1117; }
[data-testid="stAppViewContainer"] { background-color: #0D1117; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #090D13 0%, #0D1117 100%);
    border-right: 1px solid #1E2D3D;
}
[data-testid="stSidebar"] * { color: #94A3B8 !important; }
[data-testid="stSidebar"] .stRadio label { font-size: 14px !important; }

/* Radio buttons in sidebar */
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] { display: none; }

/* Tabs */
button[data-baseweb="tab"] {
    background: transparent !important;
    color: #94A3B8 !important;
    border-bottom: 2px solid transparent !important;
    font-weight: 500 !important;
    font-size: 14px !important;
}
button[data-baseweb="tab"]:hover {
    color: #F8FAFC !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #FFFFFF !important;
    font-weight: 600 !important;
    border-bottom: 2px solid currentColor !important;
}

/* DataFrames */
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #0D1117; }
::-webkit-scrollbar-thumb { background: #1E2D3D; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)


# ── helpers ───────────────────────────────────────────────────────────────────
def inject_page_bg(loc: str):
    bg = C[loc]["bg"]
    glow = C[loc]["glow"]
    st.markdown(f"""
    <style>
    .stApp, [data-testid="stAppViewContainer"] {{
        background: radial-gradient(ellipse at top left, {glow} 0%, {bg} 40%, #0D1117 100%) !important;
    }}
    </style>
    """, unsafe_allow_html=True)


def banner(title: str, subtitle: str, loc: str, emoji: str):
    p = C[loc]["p"]
    light = C[loc]["light"]
    dark = C[loc]["dark"]
    st.markdown(f"""
    <div style='
        background: linear-gradient(135deg, {dark}55 0%, {p}22 100%);
        border: 1px solid {p}44;
        border-radius: 16px;
        padding: 28px 32px;
        margin-bottom: 28px;
        position: relative;
        overflow: hidden;
    '>
        <div style='font-size:36px;margin-bottom:8px'>{emoji}</div>
        <div style='font-size:28px;font-weight:700;color:{light};letter-spacing:-0.5px'>{title}</div>
        <div style='font-size:14px;color:{p};margin-top:6px;font-weight:500'>{subtitle}</div>
        <div style='
            position:absolute;right:-20px;top:-20px;
            width:120px;height:120px;
            background:{p}15;
            border-radius:50%;
        '></div>
    </div>
    """, unsafe_allow_html=True)


def section_card(content_html: str, loc: str):
    p = C[loc]["p"]
    st.markdown(f"""
    <div style='
        background: linear-gradient(135deg, #ffffff08, #ffffff04);
        border: 1px solid {p}25;
        border-radius: 12px;
        padding: 20px 22px;
        margin-bottom: 16px;
    '>{content_html}</div>
    """, unsafe_allow_html=True)


def price_table(data: dict, title: str, loc: str, subtitle: str = ""):
    p = C[loc]["p"]
    light = C[loc]["light"]
    dark = C[loc]["dark"]

    header_cols = "".join(
        f"<th style='padding:8px 16px;background:{dark}55;color:{light};font-weight:600;font-size:12px;letter-spacing:0.5px;text-align:center'>{h}</th>"
        for h in next(iter(data.values())).keys()
    )

    rows_html = ""
    for i, (product, prices) in enumerate(data.items()):
        bg = "#ffffff06" if i % 2 == 0 else "transparent"
        cols = ""
        for v in prices.values():
            if isinstance(v, (int, float)):
                cols += f"<td style='padding:9px 16px;text-align:center;font-weight:600;color:#F8FAFC;font-size:15px'>${v}</td>"
            else:
                cols += f"<td style='padding:9px 16px;text-align:center;color:#64748B;font-size:13px'>{v if v else '—'}</td>"
        rows_html += f"<tr style='background:{bg}'><td style='padding:9px 16px;font-weight:500;color:#CBD5E1;font-size:13px'>{product}</td>{cols}</tr>"

    sub = f"<div style='font-size:12px;color:#64748B;margin-top:2px'>{subtitle}</div>" if subtitle else ""
    table = f"""
    <div style='margin-bottom:20px'>
      <div style='font-size:13px;font-weight:700;color:{p};text-transform:uppercase;letter-spacing:1px;margin-bottom:4px'>{title}</div>
      {sub}
      <table style='width:100%;border-collapse:collapse;margin-top:10px;font-size:13px;'>
        <thead>
          <tr>
            <th style='padding:8px 16px;background:{dark}55;color:{light};font-weight:600;font-size:12px;letter-spacing:0.5px;text-align:left'>Product</th>
            {header_cols}
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>"""
    st.markdown(table, unsafe_allow_html=True)


def simple_table(rows: list, title: str, loc: str):
    """rows = list of (label, value) tuples"""
    p = C[loc]["p"]
    rows_html = "".join(
        f"<tr style='background:{'#ffffff06' if i%2==0 else 'transparent'}'>"
        f"<td style='padding:8px 14px;color:#CBD5E1;font-size:13px'>{r[0]}</td>"
        f"<td style='padding:8px 14px;text-align:right;font-weight:600;color:#F8FAFC;font-size:14px'>{r[1]}</td></tr>"
        for i, r in enumerate(rows)
    )
    st.markdown(f"""
    <div style='margin-bottom:20px'>
      <div style='font-size:13px;font-weight:700;color:{p};text-transform:uppercase;letter-spacing:1px;margin-bottom:8px'>{title}</div>
      <table style='width:100%;border-collapse:collapse;'>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
    """, unsafe_allow_html=True)


def chart_layout(fig, _loc: str, height: int = 420):
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#FFFFFF", family="Inter"),
        height=height,
        legend=dict(
            orientation="h", y=-0.22,
            bgcolor="rgba(0,0,0,0)",
            font=dict(color="#FFFFFF", size=12),
        ),
        yaxis=dict(
            title=dict(text="Price (USD)", font=dict(color="#FFFFFF")),
            gridcolor="#1E2D3D",
            tickfont=dict(color="#FFFFFF"),
            zeroline=False,
        ),
        xaxis=dict(
            tickfont=dict(color="#FFFFFF"),
        ),
        margin=dict(t=20, b=60, l=10, r=10),
    )


# ── data loaders ──────────────────────────────────────────────────────────────
@st.cache_data
def load_wsd():
    wb = fetch_wb(WSD_URL)
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

    history = []
    periods = [
        ("Jul-Sep 2024", "3 Hour Soak", 55, 75),
        ("Oct-Jan 2025", "3 Hour Soak", 59, 85),
        ("Feb-May 2025", "3 Hour Soak", 49, 69),
        ("Jun-Dec 2025", "3 Hour Soak", 69, 79),
        ("Jan-Feb 2026", "Day Pass",    85, 95),
    ]
    for label, prod, mth, frs in periods:
        history.append({"Period": label, "Product": prod, "Mon-Thu": mth, "Fri-Sun": frs})
    for label, mth, frs in [
        ("Jul-Sep 2024", 35, 35), ("Oct-Jan 2025", 39, 39),
        ("Feb-May 2025", 39, 39), ("Jun-Dec 2025", 49, 59),
        ("Jan-Feb 2026", 55, 65),
    ]:
        history.append({"Period": label, "Product": "Quick Dip / Evening", "Mon-Thu": mth, "Fri-Sun": frs})

    return current, pd.DataFrame(history), pd.DataFrame(spa_current)


@st.cache_data
def load_imhs():
    wb = fetch_wb(IMHS_URL)
    latest_sheet = "Pricing 3.30.26-5.31.26"
    ws = wb[latest_sheet]
    rows = list(ws.iter_rows(values_only=True))
    non_peak, peak, section = {}, {}, None
    for row in rows:
        label = row[2] if len(row) > 2 else None
        if label == "Non-Peak Pricing":
            section = "non_peak"
        elif label == "Holiday/Peak Pricing":
            section = "peak"
        elif label in DAY_ORDER and section:
            d = {"Select 3hr": row[3], "Select All-Day": row[4], "Premier 3hr": row[5], "Premier All-Day": row[6]}
            (non_peak if section == "non_peak" else peak)[label] = d

    sheet_ranges = {
        "5.23.22": ("5.23.2022-1.22.2023", "5.23.2022"),
        "1.23.23": ("1.23.2023-4.30.2023", "1.23.2023"),
        "5.1.23":  ("5.1.2023-9.4.2023",   "5.1.2023"),
        "9.5.23":  ("9.5.2023-10.31.2023",  "9.5.2023"),
        "11.1.23": ("11.1.2023-12.21.2023", "11.1.2023"),
        "12.22.23":("12.22.2023-5.23.2024", "12.22.2023"),
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
        sec = None
        for row in ws_h.iter_rows(values_only=True):
            lbl = row[2] if len(row) > 2 else None
            if lbl == "Non-Peak Pricing": sec = "non_peak"
            elif lbl == "Holiday/Peak Pricing": sec = "peak"
            elif lbl == "Mon" and sec == "non_peak":
                try:
                    hist_records.append({"Period": label, "Date": date_str, "Type": "Non-Peak Mon",
                        "Select 3hr": row[3], "Select All-Day": row[4], "Premier 3hr": row[5], "Premier All-Day": row[6]})
                except Exception: pass
            elif lbl == "Sat" and sec == "peak":
                try:
                    hist_records.append({"Period": label, "Date": date_str, "Type": "Peak Sat",
                        "Select 3hr": row[3], "Select All-Day": row[4], "Premier 3hr": row[5], "Premier All-Day": row[6]})
                except Exception: pass

    ws_yoy = wb["2015 - Present DAy"]
    adult_peak_row = None
    for row in ws_yoy.iter_rows(values_only=True):
        if row[0] == "Everyday Adult" and adult_peak_row is None:
            adult_peak_row = row
            break
    if adult_peak_row:
        for i, yr in enumerate(["2015","2016","2017","2018","2019","2020","2021"]):
            val = adult_peak_row[i + 1]
            if val is not None:
                hist_records.append({"Period": yr, "Date": f"6.1.{yr}", "Type": "Peak (Legacy Adult)",
                    "Select 3hr": val, "Select All-Day": val, "Premier 3hr": None, "Premier All-Day": None})

    return non_peak, peak, pd.DataFrame(hist_records)


@st.cache_data
def load_zchs():
    wb = fetch_wb(ZCHS_URL)
    ws = wb["Pricing 11_2025"]
    rows = list(ws.iter_rows(values_only=True))
    current = {}
    for row in rows:
        label = str(row[0]).strip() if row[0] else ""
        if "Quick Dip Adult / Teen" in label and "Select" in label:
            if "Quick Dip Select (13+)" not in current:
                current["Quick Dip Select (13+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "Quick Dip Adult: Premier" in label:
            if "Quick Dip Premier (21+)" not in current:
                current["Quick Dip Premier (21+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "3-hour Adult / Teen" in label and "Select" in label and "Eat" not in label:
            if "3hr Select (13+)" not in current:
                current["3hr Select (13+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "3- Hour Soak Adult: Premier" in label:
            if "3hr Premier (21+)" not in current:
                current["3hr Premier (21+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "Full Day Soak Ages 13+" in label:
            if "Full Day Select (13+)" not in current:
                current["Full Day Select (13+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}
        elif "Full Day Adult: Premier" in label:
            if "Full Day Premier (21+)" not in current:
                current["Full Day Premier (21+)"] = {"Mon-Thu": row[1], "Fri/Sat/Sun": row[2]}

    pricing_sheets = {
        "Jan 2025": ("Pricing 2_3_2025", "1.22.2025"),
        "Jul 2025": ("Pricing 07_2025",  "7.1.2025"),
        "Nov 2025": ("Pricing 11_2025",  "11.1.2025"),
    }
    hist_records = []
    for period, (sheet_name, date_s) in pricing_sheets.items():
        if sheet_name not in wb.sheetnames: continue
        for row in wb[sheet_name].iter_rows(values_only=True):
            lbl = str(row[0]).strip() if row[0] else ""
            if "3-hour Adult / Teen" in lbl and "Select" in lbl and "Eat" not in lbl:
                v1 = row[1] if isinstance(row[1], (int, float)) else None
                v2 = row[2] if isinstance(row[2], (int, float)) else None
                if v1: hist_records.append({"Period": period, "Product": "3hr Select (13+)", "Mon-Thu": v1, "Fri/Sat/Sun": v2})
            elif "3- Hour Soak Adult: Premier" in lbl:
                v1 = row[1] if isinstance(row[1], (int, float)) else None
                v2 = row[2] if isinstance(row[2], (int, float)) else None
                if v1: hist_records.append({"Period": period, "Product": "3hr Premier (21+)", "Mon-Thu": v1, "Fri/Sat/Sun": v2})
            elif "Full Day Soak Ages 13+" in lbl:
                v1 = row[1] if isinstance(row[1], (int, float)) else None
                v2 = row[2] if isinstance(row[2], (int, float)) else None
                if v1: hist_records.append({"Period": period, "Product": "Full Day Select (13+)", "Mon-Thu": v1, "Fri/Sat/Sun": v2})

    return current, pd.DataFrame(hist_records)


# ── sidebar nav ───────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding: 20px 8px 8px 8px;'>
        <div style='font-size:22px;font-weight:700;color:#F8FAFC;letter-spacing:-0.5px'>♨️ Worldsprings</div>
        <div style='font-size:12px;color:#475569;margin-top:4px'>Pricing Dashboard</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    page = st.radio(
        "nav",
        ["🌐  Overview", "🔵  WSD – Dallas", "🟢  IMHS – Colorado", "🔴  ZCHS – Utah"],
        label_visibility="collapsed",
    )

    st.markdown(f"<div style='margin-top:24px;font-size:11px;color:#334155;border-top:1px solid #1E2D3D;padding-top:10px'>Updated {datetime.today().strftime('%b %d, %Y')}</div>", unsafe_allow_html=True)


def base_chart():
    return dict(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#FFFFFF", family="Inter"),
        legend=dict(orientation="h", y=-0.22, bgcolor="rgba(0,0,0,0)", font=dict(color="#FFFFFF", size=12)),
        yaxis=dict(gridcolor="#1E2D3D", tickfont=dict(color="#FFFFFF"), zeroline=False,
                   title=dict(text="Price (USD)", font=dict(color="#FFFFFF"))),
        xaxis=dict(tickfont=dict(color="#FFFFFF")),
        margin=dict(t=20, b=60, l=10, r=10),
    )


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
if "Overview" in page:
    wsd_current, _, _ = load_wsd()
    imhs_non_peak, imhs_peak, _ = load_imhs()
    zchs_current, _ = load_zchs()

    st.markdown("""
    <div style='padding:32px 0 16px 0'>
        <div style='font-size:32px;font-weight:700;color:#F8FAFC;letter-spacing:-1px'>Pricing Overview</div>
        <div style='font-size:14px;color:#64748B;margin-top:6px'>All three locations — current rates at a glance</div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        wsd_p = C["WSD"]["p"]; wsd_d = C["WSD"]["dark"]
        st.markdown(f"""<div style='background:linear-gradient(135deg,{wsd_d}55,{wsd_p}20);
            border:1px solid {wsd_p}40;border-radius:12px;padding:14px 18px;margin-bottom:12px'>
            <span style='font-size:11px;font-weight:700;color:{wsd_p};text-transform:uppercase;letter-spacing:1.5px'>WSD</span>
            <span style='font-size:15px;font-weight:700;color:#F8FAFC;margin-left:10px'>Dallas, TX</span>
        </div>""", unsafe_allow_html=True)
        price_table(wsd_current, "Soak Admission", "WSD")
        simple_table([("1-Month Trial","$125"),("Weekday Annual","$948/yr"),("Anytime Annual","$1,788/yr")], "Memberships", "WSD")
        simple_table([("Robe Rental","$10"),("Cabana Mon-Thu","$199"),("Cabana Sat","$299")], "Add-ons", "WSD")

    with col2:
        imhs_p = C["IMHS"]["p"]; imhs_d = C["IMHS"]["dark"]
        st.markdown(f"""<div style='background:linear-gradient(135deg,{imhs_d}55,{imhs_p}20);
            border:1px solid {imhs_p}40;border-radius:12px;padding:14px 18px;margin-bottom:12px'>
            <span style='font-size:11px;font-weight:700;color:{imhs_p};text-transform:uppercase;letter-spacing:1.5px'>IMHS</span>
            <span style='font-size:15px;font-weight:700;color:#F8FAFC;margin-left:10px'>Glenwood Springs, CO</span>
        </div>""", unsafe_allow_html=True)
        np_t = {d: {"Sel 3hr": imhs_non_peak.get(d,{}).get("Select 3hr"),
                    "Sel All-Day": imhs_non_peak.get(d,{}).get("Select All-Day"),
                    "Pre 3hr": imhs_non_peak.get(d,{}).get("Premier 3hr"),
                    "Pre All-Day": imhs_non_peak.get(d,{}).get("Premier All-Day")}
                for d in DAY_ORDER if d in imhs_non_peak}
        price_table(np_t, "Non-Peak Pricing", "IMHS")
        pk_t = {d: {"Sel 3hr": imhs_peak.get(d,{}).get("Select 3hr"),
                    "Sel All-Day": imhs_peak.get(d,{}).get("Select All-Day"),
                    "Pre 3hr": imhs_peak.get(d,{}).get("Premier 3hr"),
                    "Pre All-Day": imhs_peak.get(d,{}).get("Premier All-Day")}
                for d in DAY_ORDER if d in imhs_peak}
        price_table(pk_t, "Peak / Holiday Pricing", "IMHS")

    with col3:
        zchs_p = C["ZCHS"]["p"]; zchs_d = C["ZCHS"]["dark"]
        st.markdown(f"""<div style='background:linear-gradient(135deg,{zchs_d}55,{zchs_p}20);
            border:1px solid {zchs_p}40;border-radius:12px;padding:14px 18px;margin-bottom:12px'>
            <span style='font-size:11px;font-weight:700;color:{zchs_p};text-transform:uppercase;letter-spacing:1.5px'>ZCHS</span>
            <span style='font-size:15px;font-weight:700;color:#F8FAFC;margin-left:10px'>Zion Canyon, UT</span>
        </div>""", unsafe_allow_html=True)
        if zchs_current:
            disp = {pr: {k: v if isinstance(v,(int,float)) else "—" for k,v in prc.items()} for pr,prc in zchs_current.items()}
            price_table(disp, "Soak Pricing (Nov 2025)", "ZCHS")
        simple_table([("Anytime Annual","$1,499"),("Weekday Annual","$920"),("Snowbird (3 mo)","$499")], "Memberships", "ZCHS")
        simple_table([("Single Cabana Mon-Thu","$149"),("Single Cabana Fri-Sun","$249"),
                      ("Double Cabana Mon-Thu","$199"),("Double Cabana Fri-Sun","$299")], "Add-ons", "ZCHS")

    st.markdown("<div style='margin-top:32px'></div>", unsafe_allow_html=True)
    st.markdown("""<div style='font-size:18px;font-weight:700;color:#F8FAFC;margin-bottom:4px'>Comparable Entry Price</div>
    <div style='font-size:13px;color:#64748B;margin-bottom:16px'>Standard soak — weekday vs weekend</div>""", unsafe_allow_html=True)

    wsd_p=C["WSD"]["p"]; imhs_p=C["IMHS"]["p"]; zchs_p=C["ZCHS"]["p"]
    wsd_d=C["WSD"]["dark"]; imhs_d=C["IMHS"]["dark"]; zchs_d=C["ZCHS"]["dark"]

    fig = go.Figure()
    fig.add_trace(go.Bar(name="WSD Weekday", x=["WSD Dallas"], y=[85], marker_color=wsd_p,
        text=["$85"], textposition="outside", textfont=dict(color=wsd_p)))
    fig.add_trace(go.Bar(name="WSD Weekend", x=["WSD Dallas"], y=[105], marker_color=wsd_d,
        text=["$105"], textposition="outside", textfont=dict(color=wsd_p)))
    fig.add_trace(go.Bar(name="IMHS Weekday", x=["IMHS Colorado"],
        y=[imhs_non_peak.get("Mon",{}).get("Select All-Day",0)], marker_color=imhs_p,
        text=[f"${imhs_non_peak.get('Mon',{}).get('Select All-Day',0)}"],
        textposition="outside", textfont=dict(color=imhs_p)))
    fig.add_trace(go.Bar(name="IMHS Weekend Peak", x=["IMHS Colorado"],
        y=[imhs_peak.get("Sat",{}).get("Premier All-Day",0)], marker_color=imhs_d,
        text=[f"${imhs_peak.get('Sat',{}).get('Premier All-Day',0)}"],
        textposition="outside", textfont=dict(color=imhs_p)))
    fig.add_trace(go.Bar(name="ZCHS Weekday", x=["ZCHS Utah"],
        y=[zchs_current.get("3hr Select (13+)",{}).get("Mon-Thu",0)], marker_color=zchs_p,
        text=[f"${zchs_current.get('3hr Select (13+)',{}).get('Mon-Thu',0)}"],
        textposition="outside", textfont=dict(color=zchs_p)))
    fig.add_trace(go.Bar(name="ZCHS Weekend", x=["ZCHS Utah"],
        y=[zchs_current.get("3hr Select (13+)",{}).get("Fri/Sat/Sun",0)], marker_color=zchs_d,
        text=[f"${zchs_current.get('3hr Select (13+)',{}).get('Fri/Sat/Sun',0)}"],
        textposition="outside", textfont=dict(color=zchs_p)))

    layout = base_chart()
    layout.update(barmode="group", height=360,
        xaxis=dict(tickfont=dict(color="#FFFFFF", size=13)),
        bargap=0.3, bargroupgap=0.08)
    fig.update_layout(**layout)
    st.plotly_chart(fig, use_container_width=True)

    # ── price by day filter ───────────────────────────────────────────────────
    st.markdown("""<div style='font-size:18px;font-weight:700;color:#F8FAFC;margin-bottom:4px'>Price by Day of Week</div>
    <div style='font-size:13px;color:#64748B;margin-bottom:16px'>Select a day to compare admission prices across the three locations</div>""",
    unsafe_allow_html=True)

    selected_day = st.radio("Day", DAY_ORDER, horizontal=True, label_visibility="collapsed")

    # WSD: map day → tier
    wsd_tier_map = {"Mon":"Mon-Thu","Tue":"Mon-Thu","Wed":"Mon-Thu","Thu":"Mon-Thu",
                    "Fri":"Fri & Sun","Sat":"Saturday","Sun":"Fri & Sun"}
    wsd_day_price  = wsd_current["Day Pass"][wsd_tier_map[selected_day]]
    wsd_eve_price  = wsd_current["Evening Pass"][wsd_tier_map[selected_day]]

    # IMHS: non_peak prices for that day
    imhs_sel3  = imhs_non_peak.get(selected_day, {}).get("Select 3hr", 0)
    imhs_selAD = imhs_non_peak.get(selected_day, {}).get("Select All-Day", 0)
    imhs_pre3  = imhs_non_peak.get(selected_day, {}).get("Premier 3hr", 0)
    imhs_preAD = imhs_non_peak.get(selected_day, {}).get("Premier All-Day", 0)

    # ZCHS: Mon-Thu vs Fri/Sat/Sun
    zchs_key = "Mon-Thu" if selected_day in ("Mon","Tue","Wed","Thu") else "Fri/Sat/Sun"
    zchs_qd  = zchs_current.get("Quick Dip Select (13+)", {}).get(zchs_key, 0)
    zchs_3hr = zchs_current.get("3hr Select (13+)", {}).get(zchs_key, 0)
    zchs_fd  = zchs_current.get("Full Day Select (13+)", {}).get(zchs_key, 0)

    wsd_p=C["WSD"]["p"]; imhs_p=C["IMHS"]["p"]; zchs_p=C["ZCHS"]["p"]
    wsd_d=C["WSD"]["dark"]; imhs_d=C["IMHS"]["dark"]; zchs_d=C["ZCHS"]["dark"]

    # subplots: one per location so bars are always centered
    fig2 = make_subplots(rows=1, cols=3,
        subplot_titles=["WSD — Dallas", "IMHS — Colorado", "ZCHS — Utah"],
        shared_yaxes=True)

    # WSD — Day Pass + Evening Pass
    for label, val, color in [
        ("Day Pass",    wsd_day_price, wsd_p),
        ("Evening Pass", wsd_eve_price, wsd_d),
    ]:
        fig2.add_trace(go.Bar(name=label, x=[label], y=[val], marker_color=color,
            text=[f"${val}"], textposition="outside", textfont=dict(color="#FFFFFF"),
            showlegend=False), row=1, col=1)

    # IMHS — 4 tiers
    for label, val, color in [
        ("Sel 3hr",  imhs_sel3,  imhs_p),
        ("Sel All-Day", imhs_selAD, "#22c55e"),
        ("Pre 3hr",  imhs_pre3,  "#86efac"),
        ("Pre All-Day", imhs_preAD, imhs_d),
    ]:
        fig2.add_trace(go.Bar(name=label, x=[label], y=[val], marker_color=color,
            text=[f"${val}"], textposition="outside", textfont=dict(color="#FFFFFF"),
            showlegend=False), row=1, col=2)

    # ZCHS — 3 products
    for label, val, color in [
        ("Quick Dip", zchs_qd,  zchs_p),
        ("3hr Select", zchs_3hr, "#ef4444"),
        ("Full Day",  zchs_fd,  zchs_d),
    ]:
        fig2.add_trace(go.Bar(name=label, x=[label], y=[val], marker_color=color,
            text=[f"${val}"], textposition="outside", textfont=dict(color="#FFFFFF"),
            showlegend=False), row=1, col=3)

    fig2.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#FFFFFF", family="Inter"),
        height=400, showlegend=False,
        margin=dict(t=40, b=20, l=10, r=10),
    )
    fig2.update_annotations(font=dict(color="#FFFFFF", size=13))
    for i in range(1, 4):
        fig2.update_xaxes(tickfont=dict(color="#FFFFFF", size=11), row=1, col=i)
        fig2.update_yaxes(gridcolor="#1E2D3D", tickfont=dict(color="#FFFFFF"),
                          zeroline=False, row=1, col=i)

    st.plotly_chart(fig2, use_container_width=True)
    st.caption("IMHS: non-peak pricing · ZCHS: Mon–Thu or Fri/Sat/Sun tier")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: WSD
# ══════════════════════════════════════════════════════════════════════════════
elif "WSD" in page:
    inject_page_bg("WSD")
    wsd_current, wsd_hist, wsd_spa = load_wsd()
    p = C["WSD"]["p"]; light = C["WSD"]["light"]

    banner("Worldsprings Dallas", "WSD — Dallas, TX  ·  2026 Pricing", "WSD", "💧")

    tab1, tab2, tab3 = st.tabs(["  Current Rates  ", "  Pricing by Day  ", "  Price History  "])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            price_table(wsd_current, "Soak Admission 2026", "WSD",
                        "Jan–Feb: same rates · March onward: robe included TBD")
            simple_table([
                ("1-Month Trial (4 day passes)", "$125"),
                ("Weekday Annual (unlimited M-F + 5 guests)", "$948 / yr"),
                ("Anytime Annual (unlimited + 10 guests)", "$1,788 / yr"),
            ], "Memberships", "WSD")
        with col2:
            simple_table([
                ("Robe Rental", "$10"),
                ("Cabana (Mon-Thu)", "$199"),
                ("Cabana (Sat / Holiday)", "$299"),
                ("Day Pass E-Certificate", "$95"),
            ], "Add-ons", "WSD")
            simple_table([
                ("Hot Springs Sound Bath (30 min)", "FREE"),
                ("Guided Contrast Therapy (30 min)", "FREE"),
                ("Stretch, Sauna & Breathwork (30 min)", "FREE"),
                ("Cold Plunge 101 (10 min)", "FREE"),
            ], "Complimentary Classes", "WSD")

    with tab2:
        st.markdown(f"<div style='font-size:16px;font-weight:600;color:{light};margin-bottom:16px'>Price by Day of Week — WSD 2026</div>", unsafe_allow_html=True)
        # WSD has 3 tiers: Mon-Thu, Fri&Sun, Sat → map to each day
        days_full = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        day_tier  = ["Mon-Thu","Mon-Thu","Mon-Thu","Mon-Thu","Fri & Sun","Saturday","Fri & Sun"]
        day_pass_by_day  = [wsd_current["Day Pass"][t] for t in day_tier]
        evening_by_day   = [wsd_current["Evening Pass"][t] for t in day_tier]

        colors_day = [p if t=="Mon-Thu" else (C["WSD"]["dark"] if t=="Saturday" else C["WSD"]["light"]) for t in day_tier]

        fig = go.Figure()
        fig.add_trace(go.Bar(name="Day Pass", x=days_full, y=day_pass_by_day,
            marker_color=colors_day, text=[f"${v}" for v in day_pass_by_day],
            textposition="outside", textfont=dict(color="#FFFFFF")))
        fig.add_trace(go.Bar(name="Evening Pass", x=days_full, y=evening_by_day,
            marker_color=[c + "99" for c in colors_day],
            text=[f"${v}" for v in evening_by_day],
            textposition="outside", textfont=dict(color="#FFFFFF")))

        layout = base_chart()
        layout.update(barmode="group", height=400,
            annotations=[
                dict(x=1.5, y=max(day_pass_by_day)+12, text="Mon–Thu", showarrow=False,
                     font=dict(color=p, size=11)),
                dict(x=4,   y=max(day_pass_by_day)+12, text="Fri & Sun", showarrow=False,
                     font=dict(color=C["WSD"]["light"], size=11)),
                dict(x=5,   y=max(day_pass_by_day)+12, text="Saturday", showarrow=False,
                     font=dict(color=C["WSD"]["dark"], size=11)),
            ])
        fig.update_layout(**layout)
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Color indicates pricing tier: celeste = Mon-Thu · light blue = Fri & Sun · dark blue = Saturday")

    with tab3:
        st.markdown(f"<div style='font-size:16px;font-weight:600;color:{light};margin-bottom:4px'>Soak Price History</div>", unsafe_allow_html=True)
        st.caption("WSD launched mid-2024. 2026 rates reflect restructured product names.")

        day_pass_h = wsd_hist[wsd_hist["Product"].isin(["3 Hour Soak", "Day Pass"])].copy()
        evening_h  = wsd_hist[wsd_hist["Product"] == "Quick Dip / Evening"].copy()

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=day_pass_h["Period"], y=day_pass_h["Mon-Thu"],
            mode="lines+markers", name="Day Pass (Mon-Thu)",
            line=dict(color=p, width=3), marker=dict(size=9, color=p)))
        fig.add_trace(go.Scatter(x=day_pass_h["Period"], y=day_pass_h["Fri-Sun"],
            mode="lines+markers", name="Day Pass (Fri-Sun)",
            line=dict(color=p, width=3, dash="dot"), marker=dict(size=9, color=p)))
        fig.add_trace(go.Scatter(x=evening_h["Period"], y=evening_h["Mon-Thu"],
            mode="lines+markers", name="Evening (Mon-Thu)",
            line=dict(color=light, width=2), marker=dict(size=8, color=light)))
        fig.add_trace(go.Scatter(x=evening_h["Period"], y=evening_h["Fri-Sun"],
            mode="lines+markers", name="Evening (Fri-Sun)",
            line=dict(color=light, width=2, dash="dot"), marker=dict(size=8, color=light)))

        chart_layout(fig, "WSD")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(wsd_hist.set_index("Period"), use_container_width=True)



# ══════════════════════════════════════════════════════════════════════════════
# PAGE: IMHS
# ══════════════════════════════════════════════════════════════════════════════
elif "IMHS" in page:
    inject_page_bg("IMHS")
    imhs_non_peak, imhs_peak, imhs_hist = load_imhs()
    p = C["IMHS"]["p"]; light = C["IMHS"]["light"]

    banner("Iron Mountain Hot Springs", "IMHS — Glenwood Springs, CO  ·  Current Pricing", "IMHS", "🏔️")

    tab1, tab2, tab3 = st.tabs(["  Current Rates  ", "  Pricing by Day  ", "  Price History (2015–Present)  "])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            np_data = {d: {"Select 3hr": imhs_non_peak[d]["Select 3hr"],
                           "Select All-Day": imhs_non_peak[d]["Select All-Day"],
                           "Premier 3hr": imhs_non_peak[d]["Premier 3hr"],
                           "Premier All-Day": imhs_non_peak[d]["Premier All-Day"]}
                       for d in DAY_ORDER if d in imhs_non_peak}
            price_table(np_data, "Non-Peak Pricing", "IMHS")
        with col2:
            pk_data = {d: {"Select 3hr": imhs_peak[d]["Select 3hr"],
                           "Select All-Day": imhs_peak[d]["Select All-Day"],
                           "Premier 3hr": imhs_peak[d]["Premier 3hr"],
                           "Premier All-Day": imhs_peak[d]["Premier All-Day"]}
                       for d in DAY_ORDER if d in imhs_peak}
            price_table(pk_data, "Peak / Holiday Pricing", "IMHS")
        simple_table([
            ("Select 3hr", "3-hour timed soak · select pools"),
            ("Select All-Day", "Unlimited duration · select pools"),
            ("Premier 3hr", "3-hour timed soak · all pools"),
            ("Premier All-Day", "Unlimited duration · all pools"),
        ], "Product Tiers", "IMHS")

    with tab2:
        st.markdown(f"<div style='font-size:16px;font-weight:600;color:{light};margin-bottom:16px'>Price by Day of Week — IMHS (current period)</div>", unsafe_allow_html=True)

        section_sel = st.radio("Pricing season", ["Non-Peak", "Peak / Holiday"], horizontal=True, label_visibility="collapsed")
        data_src = imhs_non_peak if section_sel == "Non-Peak" else imhs_peak

        days_avail = [d for d in DAY_ORDER if d in data_src]
        metrics_imhs = ["Select 3hr", "Select All-Day", "Premier 3hr", "Premier All-Day"]
        colors_imhs4 = [p, light, "#22c55e", "#86efac"]

        fig = go.Figure()
        for metric, color in zip(metrics_imhs, colors_imhs4):
            vals = [data_src[d][metric] for d in days_avail]
            fig.add_trace(go.Bar(name=metric, x=days_avail, y=vals,
                marker_color=color,
                text=[f"${v}" for v in vals], textposition="outside",
                textfont=dict(color=color, size=11)))

        layout = base_chart()
        layout.update(barmode="group", height=430, bargap=0.2, bargroupgap=0.06)
        fig.update_layout(**layout)
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Shows how pricing varies across days of the week within the current pricing period.")

    with tab3:
        st.markdown(f"<div style='font-size:16px;font-weight:600;color:{light};margin-bottom:12px'>Price History — 2015 to Present</div>", unsafe_allow_html=True)

        if not imhs_hist.empty:
            type_filter = st.selectbox("Select day type", options=imhs_hist["Type"].unique().tolist(), label_visibility="visible")
            filtered = imhs_hist[imhs_hist["Type"] == type_filter].copy()

            colors_imhs = [p, light, "#22c55e", "#86efac"]
            metrics = ["Select 3hr", "Select All-Day", "Premier 3hr", "Premier All-Day"]

            fig = go.Figure()
            for metric, color in zip(metrics, colors_imhs):
                df_m = filtered.dropna(subset=[metric])
                if not df_m.empty:
                    fig.add_trace(go.Scatter(x=df_m["Period"], y=df_m[metric],
                        mode="lines+markers", name=metric,
                        line=dict(color=color, width=2), marker=dict(size=8, color=color)))

            chart_layout(fig, "IMHS", height=440)
            st.plotly_chart(fig, use_container_width=True)
            st.caption("Pre-2022: legacy 'Everyday Adult' rate (no Select/Premier split).")
            st.dataframe(filtered.set_index("Period"), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: ZCHS
# ══════════════════════════════════════════════════════════════════════════════
elif "ZCHS" in page:
    inject_page_bg("ZCHS")
    zchs_current, zchs_hist = load_zchs()
    p = C["ZCHS"]["p"]; light = C["ZCHS"]["light"]

    banner("Zion Canyon Hot Springs", "ZCHS — Zion Canyon, UT  ·  Nov 2025 Pricing", "ZCHS", "🏜️")

    tab1, tab2, tab3 = st.tabs(["  Current Rates  ", "  Pricing by Day  ", "  Price History  "])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            if zchs_current:
                disp = {pr: {k: v if isinstance(v,(int,float)) else "—" for k,v in prc.items()} for pr,prc in zchs_current.items()}
                price_table(disp, "Soak Pricing (Nov 2025)", "ZCHS")
        with col2:
            simple_table([("Anytime Annual","$1,499"),("Weekday Annual (M-F)","$920"),("Snowbird (3 months)","$499")], "Memberships", "ZCHS")
            simple_table([("Robe","$10"),("Single Cabana Mon-Thu","$149"),("Single Cabana Fri-Sun","$249"),
                          ("Double Cabana Mon-Thu","$199"),("Double Cabana Fri-Sun","$299")], "Add-ons", "ZCHS")
            simple_table([("Washington County Residents","20% off"),("Hotel Partner Rate","10% off")], "Special Rates", "ZCHS")
        simple_table([("Select (13+)","Standard pools · teens & adults"),
                      ("Select Youth (3–12)","Standard pools · children"),
                      ("Premier (21+)","All pools including 21+ only areas")], "Access Tiers", "ZCHS")

    with tab2:
        st.markdown(f"<div style='font-size:16px;font-weight:600;color:{light};margin-bottom:16px'>Price by Day Type — ZCHS (Nov 2025)</div>", unsafe_allow_html=True)

        if zchs_current:
            products  = list(zchs_current.keys())
            wkday_vals = [zchs_current[pr].get("Mon-Thu",0) for pr in products]
            wkend_vals = [zchs_current[pr].get("Fri/Sat/Sun",0) for pr in products]
            pct_diff   = [round((w-d)/d*100,1) if d else 0 for d,w in zip(wkday_vals, wkend_vals)]

            fig = go.Figure()
            fig.add_trace(go.Bar(name="Mon–Thu", x=products, y=wkday_vals, marker_color=p,
                text=[f"${v}" for v in wkday_vals], textposition="outside", textfont=dict(color=p)))
            fig.add_trace(go.Bar(name="Fri / Sat / Sun", x=products, y=wkend_vals,
                marker_color=C["ZCHS"]["dark"],
                text=[f"${v}" for v in wkend_vals], textposition="outside", textfont=dict(color=light)))

            layout = base_chart()
            layout.update(barmode="group", height=430, bargap=0.25, bargroupgap=0.08,
                xaxis=dict(tickfont=dict(color="#FFFFFF", size=11), tickangle=-20))
            fig.update_layout(**layout)
            st.plotly_chart(fig, use_container_width=True)

            # weekend premium callouts
            st.markdown(f"<div style='font-size:13px;font-weight:600;color:{p};margin-top:4px;margin-bottom:8px'>Weekend premium by product</div>", unsafe_allow_html=True)
            cols = st.columns(len(products))
            for i, (pr, pct) in enumerate(zip(products, pct_diff)):
                with cols[i]:
                    st.markdown(f"""<div style='background:{C["ZCHS"]["dark"]}33;border:1px solid {p}30;border-radius:8px;
                        padding:10px;text-align:center'>
                        <div style='font-size:11px;color:#64748B;margin-bottom:4px'>{pr}</div>
                        <div style='font-size:20px;font-weight:700;color:{p}'>+{pct}%</div>
                    </div>""", unsafe_allow_html=True)

    with tab3:
        st.markdown(f"<div style='font-size:16px;font-weight:600;color:{light};margin-bottom:12px'>Price History — Since Opening</div>", unsafe_allow_html=True)
        st.caption("ZCHS opened January 2025. Three pricing periods captured.")

        if not zchs_hist.empty:
            product_filter = st.selectbox("Product", options=zchs_hist["Product"].unique().tolist(), label_visibility="visible")
            filtered = zchs_hist[zchs_hist["Product"] == product_filter]

            fig = go.Figure()
            fig.add_trace(go.Bar(name="Mon-Thu", x=filtered["Period"], y=filtered["Mon-Thu"],
                marker_color=p, text=filtered["Mon-Thu"].apply(lambda v: f"${v}" if pd.notna(v) else ""),
                textposition="outside", textfont=dict(color=p)))
            fig.add_trace(go.Bar(name="Fri/Sat/Sun", x=filtered["Period"], y=filtered["Fri/Sat/Sun"],
                marker_color=C["ZCHS"]["dark"], text=filtered["Fri/Sat/Sun"].apply(lambda v: f"${v}" if pd.notna(v) else ""),
                textposition="outside", textfont=dict(color=light)))

            layout = base_chart()
            layout.update(barmode="group", height=400, bargap=0.35)
            fig.update_layout(**layout)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(filtered.set_index("Period"), use_container_width=True)
        else:
            st.info("No historical data parsed yet.")
