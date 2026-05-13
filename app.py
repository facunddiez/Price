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

    def _try_parse_start(sname):
        """Best-effort: extract a start Timestamp from a sheet name. Returns pd.NaT if unparseable."""
        n = sname.strip()
        if n.lower().startswith("pricing "):
            n = n[8:].strip()
        start_part = n.replace(" ", "").split("-")[0]
        parts = start_part.split(".")
        if len(parts) != 3:
            return pd.NaT
        try:
            mo, da, yr = int(parts[0]), int(parts[1]), int(parts[2])
            if yr < 100:
                yr += 2000
            if not (1 <= mo <= 12 and 1 <= da <= 31 and 2015 <= yr <= 2035):
                return pd.NaT
            return pd.Timestamp(f"{yr}-{mo:02d}-{da:02d}")
        except Exception:
            return pd.NaT

    # Identify pricing sheets by name (skip known non-pricing sheets).
    # Use column-agnostic content parsing: find section markers in ANY column,
    # then read the 4 data values immediately to the right of the day label.
    NON_PRICING = {"2015 - Present DAy", "breif discritpion of pricing",
                   "brief discritpion of pricing", "brief description of pricing"}

    def _extract_pricing(ws_h):
        """Return (np_row_data, pk_row_data) as dicts, or None for each if not found."""
        sec = None
        label_col = None
        np_data = pk_data = None
        NP_MARKERS = {"non-peak pricing", "non peak pricing"}
        PK_MARKERS = {"holiday/peak pricing", "peak pricing", "holiday pricing", "peak/holiday pricing"}
        for row in ws_h.iter_rows(values_only=True):
            for ci, val in enumerate(row):
                sv = str(val).strip().lower() if val is not None else ""
                if sv in NP_MARKERS:
                    sec = "non_peak"; label_col = ci; break
                elif sv in PK_MARKERS:
                    sec = "peak"; label_col = ci; break
            else:
                # No section marker — check if this row has a day label in label_col
                if sec and label_col is not None and label_col < len(row):
                    day = row[label_col]
                    if day == "Mon" and sec == "non_peak" and np_data is None:
                        try:
                            c = label_col + 1
                            if row[c] is not None:
                                np_data = {"Select 3hr": row[c], "Select All-Day": row[c+1],
                                           "Premier 3hr": row[c+2], "Premier All-Day": row[c+3]}
                        except Exception:
                            pass
                    elif day == "Sat" and sec == "peak" and pk_data is None:
                        try:
                            c = label_col + 1
                            if row[c] is not None:
                                pk_data = {"Select 3hr": row[c], "Select All-Day": row[c+1],
                                           "Premier 3hr": row[c+2], "Premier All-Day": row[c+3]}
                        except Exception:
                            pass
            if np_data and pk_data:
                break
        return np_data, pk_data

    pricing_sheets_ordered = []
    for idx, sname in enumerate(wb.sheetnames):
        if sname.strip().lower() in {s.lower() for s in NON_PRICING}:
            continue
        ts = _try_parse_start(sname)
        if ts is pd.NaT:
            continue  # skip sheets whose name can't be parsed as a date range
        pricing_sheets_ordered.append((ts, idx, sname))
    pricing_sheets_ordered.sort(key=lambda x: x[0])

    hist_records = []
    for ts, idx, sname in pricing_sheets_ordered:
        mo, da, yr = ts.month, ts.day, ts.year
        label    = f"{mo}.{da}.{str(yr)[2:]}"
        date_str = f"{mo}.{da}.{yr}"
        np_data, pk_data = _extract_pricing(wb[sname])
        if np_data:
            hist_records.append({"Period": label, "Date": date_str, "Type": "Non-Peak Mon", **np_data})
        if pk_data:
            hist_records.append({"Period": label, "Date": date_str, "Type": "Peak Sat", **pk_data})

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

    # Build debug report
    periods_found = sorted(set(r["Period"] for r in hist_records if r["Type"] == "Non-Peak Mon"))
    debug_lines = [
        f"**Pricing sheets found:** {len(pricing_sheets_ordered)}",
        f"**Non-Peak Mon records:** {sum(1 for r in hist_records if r['Type']=='Non-Peak Mon')}",
        f"**Peak Sat records:** {sum(1 for r in hist_records if r['Type']=='Peak Sat')}",
        f"**Periods:** {periods_found}",
        "", "**All workbook sheets:**"
    ]
    for sname in wb.sheetnames:
        ts = _try_parse_start(sname)
        status = f"✅ parsed as {ts.date()}" if ts is not pd.NaT else "⚠️ skipped (name not parsed)"
        debug_lines.append(f"- `{sname}` — {status}")

    return non_peak, peak, pd.DataFrame(hist_records), "\n".join(debug_lines)


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
        "Feb 2025": ("Pricing 2_3_2025", "2.3.2025"),
        "Jul 2025": ("Pricing 07_2025",  "7.1.2025"),
        "Nov 2025": ("Pricing 11_2025",  "11.1.2025"),
    }
    hist_records = []
    seen = {}
    for period, (sheet_name, date_s) in pricing_sheets.items():
        if sheet_name not in wb.sheetnames: continue
        seen[period] = set()
        in_section = None  # tracks current section for Jan 2025 format

        for row in wb[sheet_name].iter_rows(values_only=True):
            lbl = str(row[0]).strip() if row[0] else ""
            v1 = row[1] if isinstance(row[1], (int, float)) else None
            v2 = row[2] if isinstance(row[2], (int, float)) else None

            # Detect section headers (non-numeric col B = header row)
            if v1 is None:
                if "Quick Dip" in lbl:
                    in_section = "qd"
                elif "Day Pass" in lbl or ("Full Day" in lbl and "Soak" in lbl):
                    in_section = "fd"
                continue

            def _add(prod):
                key = (prod, period)
                if key not in seen[period]:
                    seen[period].add(key)
                    hist_records.append({"Period": period, "Date": date_s,
                                         "Product": prod, "Mon-Thu": v1, "Fri/Sat/Sun": v2})

            # Jan 2025 format: bare "Select Access" row identified by section context
            # "Quick Dip (3 hours)" in Jan 2025 = same 3-hour product as "3hr Select" later
            # The short Quick Dip (~1.5hr) is a new product added from Jul 2025 onward
            if lbl == "Select Access":
                if in_section == "qd":
                    _add("3hr Select (13+)")
                elif in_section == "fd":
                    _add("Full Day Select (13+)")

            # Jul/Nov 2025 format: full descriptive labels
            if "Quick Dip Adult / Teen" in lbl and "Select" in lbl:
                _add("Quick Dip Select (13+)")
            elif "3-hour Adult / Teen" in lbl and "Select" in lbl and "Eat" not in lbl:
                _add("3hr Select (13+)")
            elif "Full Day Soak Ages 13+" in lbl:
                _add("Full Day Select (13+)")

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
    imhs_non_peak, imhs_peak, _, _dbg = load_imhs()
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
            price_table(disp, "Soak Pricing (May 2026)", "ZCHS")
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
    imhs_non_peak, imhs_peak, imhs_hist, _imhs_debug = load_imhs()
    p = C["IMHS"]["p"]; light = C["IMHS"]["light"]

    banner("Iron Mountain Hot Springs", "IMHS — Glenwood Springs, CO  ·  Current Pricing", "IMHS", "🏔️")

    with st.expander("🔍 Debug: Sheet Detection (remove after fix)"):
        st.markdown(_imhs_debug)

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

        days_np   = [d for d in DAY_ORDER if d in imhs_non_peak]
        days_pk   = [d for d in DAY_ORDER if d in imhs_peak]
        metrics_imhs = ["Select 3hr", "Select All-Day", "Premier 3hr", "Premier All-Day"]
        # 4-step green gradient: light → dark
        colors_g4 = ["#A7F3D0", "#4ADE80", "#22C55E", "#15803D"]

        def _imhs_bar(data_src, days, title):
            fig = go.Figure()
            for metric, color in zip(metrics_imhs, colors_g4):
                vals = [data_src.get(d, {}).get(metric, 0) for d in days]
                fig.add_trace(go.Bar(name=metric, x=days, y=vals,
                    marker_color=color,
                    text=[f"${v}" for v in vals], textposition="outside",
                    textfont=dict(color="#FFFFFF", size=10)))
            layout = base_chart()
            layout.update(barmode="group", height=370, bargap=0.2, bargroupgap=0.06,
                title=dict(text=title, font=dict(color="#FFFFFF", size=13), x=0.5),
                xaxis=dict(tickfont=dict(color="#FFFFFF")),
                yaxis=dict(tickprefix="$", tickfont=dict(color="#FFFFFF")),
                legend=dict(font=dict(color="#FFFFFF", size=10)),
                showlegend=True)
            fig.update_layout(**layout)
            return fig

        col_np, col_pk = st.columns(2)
        with col_np:
            st.plotly_chart(_imhs_bar(imhs_non_peak, days_np, "Non-Peak Pricing"), use_container_width=True)
        with col_pk:
            st.plotly_chart(_imhs_bar(imhs_peak, days_pk, "Peak / Holiday Pricing"), use_container_width=True)

        # ── KPIs ──────────────────────────────────────────────────────────────
        def _kpi_card(label, value_str, color, bg):
            return f"""<div style='background:{bg};border:1px solid {color}40;border-radius:8px;
                padding:10px 8px;text-align:center;min-width:0'>
                <div style='font-size:10px;color:#94A3B8;margin-bottom:4px;line-height:1.3'>{label}</div>
                <div style='font-size:18px;font-weight:700;color:{color}'>{value_str}</div>
            </div>"""

        bg_card = C["IMHS"]["dark"] + "33"

        # -- Weekend premium (Non-Peak: Sat vs Mon) --
        np_mon = imhs_non_peak.get("Mon", {})
        np_sat = imhs_non_peak.get("Sat", {})
        st.markdown(f"<div style='font-size:13px;font-weight:600;color:{p};margin:18px 0 8px'>Weekend Premium — Non-Peak (Sat vs Mon)</div>", unsafe_allow_html=True)
        cols_w = st.columns(len(metrics_imhs))
        for i, metric in enumerate(metrics_imhs):
            mon_v = np_mon.get(metric) or 0
            sat_v = np_sat.get(metric) or 0
            pct = round((sat_v - mon_v) / mon_v * 100, 1) if mon_v else 0
            with cols_w[i]:
                st.markdown(_kpi_card(metric, f"+{pct}%", p, bg_card), unsafe_allow_html=True)

        # -- Select vs Premier (Non-Peak Mon) --
        st.markdown(f"<div style='font-size:13px;font-weight:600;color:{p};margin:18px 0 8px'>Premier vs Select Premium — Non-Peak Mon</div>", unsafe_allow_html=True)
        comparisons_sp = [
            ("3hr  ·  Premier vs Select", "Premier 3hr", "Select 3hr"),
            ("All-Day  ·  Premier vs Select", "Premier All-Day", "Select All-Day"),
        ]
        cols_sp = st.columns(2)
        for i, (lbl, higher, lower) in enumerate(comparisons_sp):
            h = np_mon.get(higher) or 0
            l = np_mon.get(lower) or 0
            pct = round((h - l) / l * 100, 1) if l else 0
            with cols_sp[i]:
                st.markdown(_kpi_card(lbl, f"+{pct}%", C["IMHS"]["light"], bg_card), unsafe_allow_html=True)

        # -- Non-Peak vs Peak (Mon and Sat) --
        pk_mon = imhs_peak.get("Mon", {})
        pk_sat = imhs_peak.get("Sat", {})
        st.markdown(f"<div style='font-size:13px;font-weight:600;color:{p};margin:18px 0 8px'>Peak vs Non-Peak Premium</div>", unsafe_allow_html=True)
        ref_days = [("Mon", np_mon, pk_mon), ("Sat", np_sat, pk_sat)]
        for day_lbl, np_d, pk_d in ref_days:
            cols_pk = st.columns(len(metrics_imhs))
            st.markdown(f"<div style='font-size:11px;color:#64748B;margin:6px 0 4px'>{day_lbl}</div>", unsafe_allow_html=True)
            for i, metric in enumerate(metrics_imhs):
                np_v = np_d.get(metric) or 0
                pk_v = pk_d.get(metric) or 0
                pct = round((pk_v - np_v) / np_v * 100, 1) if np_v else 0
                sign = "+" if pct >= 0 else ""
                with cols_pk[i]:
                    st.markdown(_kpi_card(metric, f"{sign}{pct}%", C["IMHS"]["light"], bg_card), unsafe_allow_html=True)

    with tab3:
        st.markdown(f"<div style='font-size:16px;font-weight:600;color:{light};margin-bottom:12px'>Price History — 2022 to Present</div>", unsafe_allow_html=True)

        if not imhs_hist.empty:
            def _parse_dt(s):
                try:
                    p = s.strip().split(".")
                    return pd.Timestamp(f"{p[2]}-{int(p[0]):02d}-{int(p[1]):02d}")
                except Exception:
                    return pd.NaT

            metrics_imhs = ["Select 3hr", "Select All-Day", "Premier 3hr", "Premier All-Day"]
            colors_imhs4 = ["#A7F3D0", "#4ADE80", "#22C55E", "#15803D"]
            end_ts = pd.Timestamp("2026-05-31")

            # Build dated dataframes for each day type
            def _make_dated(type_name):
                df = imhs_hist[imhs_hist["Type"] == type_name].copy()
                df["Date_ts"] = df["Date"].apply(_parse_dt)
                df = df.dropna(subset=["Date_ts"]).sort_values("Date_ts").reset_index(drop=True)
                if df.empty:
                    return df
                last = df.iloc[-1].copy()
                last["Period"] = "5.31.26"
                last["Date_ts"] = end_ts
                return pd.concat([df, pd.DataFrame([last])], ignore_index=True)

            np_ext = _make_dated("Non-Peak Mon")
            pk_ext = _make_dated("Peak Sat")

            st.markdown("""<style>
            div[data-testid="stRadio"] label {color:#FFFFFF !important; font-weight:600}
            div[data-testid="stRadio"] label p {color:#FFFFFF !important}
            </style>""", unsafe_allow_html=True)

            col_left, col_right = st.columns([1, 2])
            with col_left:
                day_sel = st.radio("Day type", ["Non-Peak Mon", "Peak Sat"], horizontal=True)
            with col_right:
                prod_sel = st.multiselect("Products", metrics_imhs, default=metrics_imhs)

            df_trend = (np_ext if day_sel == "Non-Peak Mon" else pk_ext).copy()

            # ── Trendline ─────────────────────────────────────────────────────
            fig_trend = go.Figure()
            for metric, col in zip(metrics_imhs, colors_imhs4):
                if metric not in prod_sel:
                    continue
                df_m = df_trend.dropna(subset=[metric])
                if df_m.empty:
                    continue
                fig_trend.add_trace(go.Scatter(
                    x=df_m["Date_ts"], y=df_m[metric],
                    name=metric, mode="lines+markers",
                    line=dict(color=col, width=2.5),
                    marker=dict(size=8, color=col),
                    hovertemplate=f"<b>{metric}</b><br>%{{x|%b %Y}}: $%{{y}}<extra></extra>"
                ))
            layout_t = base_chart()
            layout_t.update(height=420,
                xaxis=dict(tickformat="%b %Y", tickfont=dict(color="#FFFFFF", size=11)),
                yaxis=dict(tickprefix="$", tickfont=dict(color="#FFFFFF", size=11)),
                legend=dict(font=dict(color="#FFFFFF", size=11)),
                title=dict(text=f"Price Trend — {day_sel}", font=dict(color="#FFFFFF", size=14), x=0.5),
            )
            fig_trend.update_layout(**layout_t)
            st.plotly_chart(fig_trend, use_container_width=True)

            # ── Scrollable table ──────────────────────────────────────────────
            st.markdown(f"<div style='font-size:14px;font-weight:600;color:{light};margin:20px 0 8px'>Price Table by Period</div>", unsafe_allow_html=True)
            # sort by Date_ts BEFORE dropping columns
            tbl = df_trend[df_trend["Period"] != "5.31.26"].sort_values("Date_ts")
            tbl = tbl[["Period", "Date"] + metrics_imhs].copy()
            for m in metrics_imhs:
                tbl[m] = tbl[m].apply(lambda v: f"${v:.0f}" if pd.notna(v) else "—")
            st.dataframe(tbl.set_index("Period"), use_container_width=True, height=300)

            # ── Bar chart per period ───────────────────────────────────────────
            st.markdown(f"<div style='font-size:14px;font-weight:600;color:{light};margin:20px 0 8px'>Price by Period — All Products</div>", unsafe_allow_html=True)
            df_b = df_trend[df_trend["Period"] != "5.31.26"].sort_values("Date_ts")
            fig_bar = go.Figure()
            for metric, col in zip(metrics_imhs, colors_imhs4):
                fig_bar.add_trace(go.Bar(
                    name=metric, x=df_b["Period"], y=df_b[metric],
                    marker_color=col,
                    text=df_b[metric].apply(lambda v: f"${v:.0f}" if pd.notna(v) else ""),
                    textposition="outside", textfont=dict(color="#FFFFFF", size=10)
                ))
            layout_b = base_chart()
            layout_b.update(barmode="group", height=450, bargap=0.15, bargroupgap=0.05,
                xaxis=dict(tickfont=dict(color="#FFFFFF", size=10), tickangle=-30),
                yaxis=dict(tickprefix="$", tickfont=dict(color="#FFFFFF", size=11)),
                legend=dict(font=dict(color="#FFFFFF", size=11)),
            )
            fig_bar.update_layout(**layout_b)
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("No historical data parsed yet.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: ZCHS
# ══════════════════════════════════════════════════════════════════════════════
elif "ZCHS" in page:
    inject_page_bg("ZCHS")
    zchs_current, zchs_hist = load_zchs()
    p = C["ZCHS"]["p"]; light = C["ZCHS"]["light"]

    banner("Zion Canyon Hot Springs", "ZCHS — Zion Canyon, UT  ·  May 2026 Pricing", "ZCHS", "🏜️")

    tab1, tab2, tab3 = st.tabs(["  Current Rates  ", "  Pricing by Day  ", "  Price History  "])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            if zchs_current:
                disp = {pr: {k: v if isinstance(v,(int,float)) else "—" for k,v in prc.items()} for pr,prc in zchs_current.items()}
                price_table(disp, "Soak Pricing (May 2026)", "ZCHS")
        with col2:
            simple_table([("Anytime Annual","$1,499"),("Weekday Annual (M-F)","$920"),("Snowbird (3 months)","$499")], "Memberships", "ZCHS")
            simple_table([("Robe","$10"),("Single Cabana Mon-Thu","$149"),("Single Cabana Fri-Sun","$249"),
                          ("Double Cabana Mon-Thu","$199"),("Double Cabana Fri-Sun","$299")], "Add-ons", "ZCHS")
            simple_table([("Washington County Residents","20% off"),("Hotel Partner Rate","10% off")], "Special Rates", "ZCHS")
        simple_table([("Select (13+)","Standard pools · teens & adults"),
                      ("Select Youth (3–12)","Standard pools · children"),
                      ("Premier (21+)","All pools including 21+ only areas")], "Access Tiers", "ZCHS")

    with tab2:
        st.markdown(f"<div style='font-size:16px;font-weight:600;color:{light};margin-bottom:16px'>Price by Day Type — ZCHS (May 2026)</div>", unsafe_allow_html=True)

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

        if not zchs_hist.empty:
            date_map = {
                "Feb 2025": pd.Timestamp("2025-02-03"),
                "Jul 2025": pd.Timestamp("2025-07-01"),
                "Nov 2025": pd.Timestamp("2025-11-01"),
            }
            hist_dated = zchs_hist.copy()
            hist_dated["Date"] = hist_dated["Period"].map(date_map)

            # Extend Nov 2025 prices to today (no newer sheet)
            today_ts = pd.Timestamp("2026-05-13")
            nov_rows = hist_dated[hist_dated["Period"] == "Nov 2025"].copy()
            nov_rows["Period"] = "May 2026"
            nov_rows["Date"] = today_ts
            extended = pd.concat([hist_dated, nov_rows], ignore_index=True)

            key_products = ["Quick Dip Select (13+)", "3hr Select (13+)", "Full Day Select (13+)"]
            prod_colors  = [C["ZCHS"]["light"], p, C["ZCHS"]["dark"]]
            periods_order = ["Feb 2025", "Jul 2025", "Nov 2025", "May 2026"]

            day_type = st.radio("Price type", ["Mon-Thu", "Fri/Sat/Sun"], horizontal=True)

            # ── Trendline ─────────────────────────────────────────────────────
            fig_trend = go.Figure()
            for prod, col in zip(key_products, prod_colors):
                df_p = extended[extended["Product"] == prod].sort_values("Date")
                if df_p.empty: continue
                fig_trend.add_trace(go.Scatter(
                    x=df_p["Date"], y=df_p[day_type],
                    name=prod, mode="lines+markers",
                    line=dict(color=col, width=2.5),
                    marker=dict(size=9, color=col),
                    hovertemplate=f"<b>{prod}</b><br>%{{x|%b %Y}}: $%{{y}}<extra></extra>"
                ))
            layout_t = base_chart()
            layout_t.update(height=420,
                xaxis=dict(tickformat="%b %Y", tickfont=dict(color="#FFFFFF", size=11)),
                yaxis=dict(tickprefix="$", tickfont=dict(color="#FFFFFF", size=11)),
                legend=dict(font=dict(color="#FFFFFF", size=11)),
                title=dict(text=f"Price Trend ({day_type})", font=dict(color="#FFFFFF", size=14), x=0.5),
            )
            fig_trend.update_layout(**layout_t)
            st.plotly_chart(fig_trend, use_container_width=True)

            # ── Table ─────────────────────────────────────────────────────────
            st.markdown(f"<div style='font-size:14px;font-weight:600;color:{light};margin:20px 0 8px'>Price Table by Period</div>", unsafe_allow_html=True)
            tbl = extended[extended["Product"].isin(key_products)][["Period","Product","Mon-Thu","Fri/Sat/Sun"]].copy()
            tbl["Period"] = pd.Categorical(tbl["Period"], categories=periods_order, ordered=True)
            tbl = tbl.sort_values(["Period","Product"]).rename(columns={"Period": "Month"})
            tbl["Mon-Thu"]     = tbl["Mon-Thu"].apply(lambda v: f"${v:.0f}" if pd.notna(v) else "—")
            tbl["Fri/Sat/Sun"] = tbl["Fri/Sat/Sun"].apply(lambda v: f"${v:.0f}" if pd.notna(v) else "—")
            st.dataframe(tbl.set_index("Month"), use_container_width=True)

            # ── Bar chart per period ───────────────────────────────────────────
            st.markdown(f"<div style='font-size:14px;font-weight:600;color:{light};margin:20px 0 8px'>Price by Period — Key Products</div>", unsafe_allow_html=True)
            fig_bar = go.Figure()
            for prod, col in zip(key_products, prod_colors):
                df_p = extended[extended["Product"] == prod].copy()
                df_p["Period"] = pd.Categorical(df_p["Period"], categories=periods_order, ordered=True)
                df_p = df_p.drop_duplicates("Period").set_index("Period").reindex(periods_order).reset_index()
                fig_bar.add_trace(go.Bar(
                    name=prod, x=df_p["Period"], y=df_p[day_type],
                    marker_color=col,
                    text=df_p[day_type].apply(lambda v: f"${v:.0f}" if pd.notna(v) else ""),
                    textposition="outside", textfont=dict(color="#FFFFFF", size=11)
                ))
            layout_b = base_chart()
            layout_b.update(barmode="group", height=420, bargap=0.2, bargroupgap=0.08,
                xaxis=dict(tickfont=dict(color="#FFFFFF", size=11)),
                yaxis=dict(tickprefix="$", tickfont=dict(color="#FFFFFF", size=11)),
                legend=dict(font=dict(color="#FFFFFF", size=11)),
            )
            fig_bar.update_layout(**layout_b)
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("No historical data parsed yet.")
