import io
import base64
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Shipment Control Tower", page_icon="🚢", layout="wide", initial_sidebar_state="expanded")

ROOT = Path(__file__).resolve().parent
LOCAL_FILES = [
    ROOT / "IMPORT DELIVERY SHIPMENTS (1).xlsx",
    ROOT / "IMPORT DELIVERY SHIPMENTS.xlsx",
]
NAVY = "#0b3158"
BLUE = "#155caa"
TEAL = "#15945b"
ORANGE = "#dc4545"
RED = "#c45c56"

st.markdown("""
<style>
.block-container {padding-top:1.15rem; padding-bottom:2rem; max-width:1600px}
h1,h2,h3 {color:#17324f; letter-spacing:-.025em}
div[data-testid="stMetric"] {background:#fff; border:1px solid #e3e9f0; border-radius:9px; padding:14px 16px}
div[data-testid="stMetricLabel"] {color:#5e7186}
div[data-testid="stMetricValue"] {color:#17324f}
[data-testid="stSidebar"] {background:#f5f8fb}
.stTabs [data-baseweb="tab"] {font-weight:600}
.small-note {color:#64748b; font-size:.85rem}
</style>
""", unsafe_allow_html=True)


# Corporate maritime dashboard styling: inspired by the approved visual reference.
st.markdown("""
<style>
.stApp {background:linear-gradient(180deg,#f0f7fd 0%,#f7fbff 48%,#ffffff 100%);color:#11365b}
[data-testid="stSidebar"] {background:#f5f9fd!important;border-right:1px solid #dce7f1}
[data-testid="stSidebar"] h2,[data-testid="stSidebar"] label {color:#103a63!important}
[data-testid="stSidebar"] [data-baseweb="select"]>div,
[data-testid="stSidebar"] input {border-radius:9px!important;border-color:#d9e5f0!important}
.block-container {max-width:1750px!important;padding-top:0.9rem!important}
div[data-testid="stMetric"] {background:#fff!important;border:1px solid #e7eef5!important;
  border-radius:13px!important;padding:18px 16px!important;box-shadow:0 4px 18px rgba(16,59,99,.075);
  min-height:116px;position:relative;overflow:hidden}
div[data-testid="stMetric"]:before {content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:#48a7df}
div[data-testid="stMetricLabel"] {color:#315678!important;font-weight:600}
/* Status KPI cards: Delivered green, Other / Pending red */
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child(2) div[data-testid="stMetric"] {background:#ecf9f1!important;border-color:#b8e7c9!important}
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child(2) div[data-testid="stMetric"]:before {background:#15945b!important}
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child(2) div[data-testid="stMetricValue"] {color:#127443!important}
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child(3) div[data-testid="stMetric"] {background:#fff0f0!important;border-color:#f3c5c5!important}
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child(3) div[data-testid="stMetric"]:before {background:#dc4545!important}
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child(3) div[data-testid="stMetricValue"] {color:#bb2929!important}
div[data-testid="stMetricValue"] {color:#10365e!important;font-size:clamp(22px,2vw,35px)!important;font-weight:750}
.stTabs [data-baseweb="tab-list"] {gap:10px}
.stTabs [data-baseweb="tab"] {border-radius:9px 9px 0 0;padding:11px 20px;color:#174b78;font-weight:700}
.stTabs [aria-selected="true"] {color:#0d4f8a!important}
[data-testid="stPlotlyChart"] {background:#fff;border:1px solid #e7eef5;border-radius:13px;
  box-shadow:0 4px 18px rgba(16,59,99,.06);padding:8px}
.stButton>button {border-radius:9px;border-color:#d7e5f2}
h3 {color:#10365e!important;font-weight:750!important}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
[data-testid="stSidebar"] {background:linear-gradient(180deg,#f1f7fc 0%,#ffffff 100%)!important}
[data-testid="stSidebar"] [data-testid="stSidebarContent"] {padding-top:1.1rem}
.filter-intro{background:linear-gradient(125deg,#07365f,#17618f);color:white;padding:22px 17px 20px;
  border-radius:15px;margin:3px 0 19px;box-shadow:0 6px 18px rgba(13,61,100,.16)}
.filter-eyebrow{font-size:10px;letter-spacing:1.6px;color:#b9dbef;font-weight:750;margin-bottom:8px}
.filter-heading{font-size:24px;font-weight:800;letter-spacing:-.5px;line-height:1.25}
.filter-description{font-size:12px;color:#d7eafa;line-height:1.6;margin-top:9px}
.filter-section{font-size:11px;font-weight:800;letter-spacing:1.1px;color:#1b638d;
  border-bottom:1px solid #d6e6f1;padding:11px 0 9px;margin:13px 0 13px}
.filter-footer{background:#e9f4fa;border:1px solid #d6e9f4;border-radius:11px;
  padding:12px 13px;margin-top:19px;color:#3a6481;font-size:11px;line-height:1.9}
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p{font-size:12px!important;font-weight:700!important;color:#234e71!important}
[data-testid="stSidebar"] [data-baseweb="select"]>div,
[data-testid="stSidebar"] [data-testid="stTextInput"] input{
 background:#fff!important;border:1px solid #d5e4ef!important;border-radius:10px!important;
 min-height:43px!important;box-shadow:0 2px 6px rgba(20,65,98,.035)}
[data-testid="stSidebar"] [data-baseweb="select"]>div:focus-within,
[data-testid="stSidebar"] [data-testid="stTextInput"] input:focus{border-color:#2386bd!important;box-shadow:0 0 0 2px #2386bd20!important}
[data-testid="stSidebar"] [data-testid="stCheckbox"]{padding:5px 0}
</style>
""",unsafe_allow_html=True)

@st.cache_data(show_spinner="Reading shipment workbook...")
def read_workbook(data: bytes) -> pd.DataFrame:
    excel = pd.ExcelFile(io.BytesIO(data))
    # Dashboard source: BRANCHES only. MAY-JULY is intentionally excluded.
    if "BRANCHES" not in excel.sheet_names:
        raise ValueError("BRANCHES worksheet was not found in the Excel workbook.")
    df = pd.read_excel(
        excel,
        sheet_name="BRANCHES",
        dtype={"CNTNR NO": str, "BL NO": str, "PO NO": str},
    )
    df.columns = [str(c).strip().upper() for c in df.columns]
    df = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    df["SOURCE SHEET"] = "BRANCHES"
    if "CNTNR NO" not in df.columns:
        raise ValueError("CNTNR NO column is missing.")
    df["CNTNR NO"] = df["CNTNR NO"].astype("string").str.strip()
    # Keep only actual container records; exclude blank, placeholder and header rows.
    invalid_ids = {"", "nan", "none", "nat", "<na>", "cntnr no", "container no", "-"}
    df = df[df["CNTNR NO"].notna() & ~df["CNTNR NO"].str.lower().isin(invalid_ids)].copy()
    for col in ("SUPPLIER", "STATUS", "DELIVERY LOC", "FORWARDING AGENT", "ITEMS", "REMARKS", "BL NO", "PO NO"):
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()
    # Excel contains two columns named STATUS. Pandas renames the second to STATUS.1.
    # STATUS = delivery progress; STATUS.1 = No / To Paid payment status.
    payment_col = next((c for c in df.columns if c.upper() == "STATUS.1"), None)
    if payment_col:
        df["PAYMENT STATUS"] = df[payment_col].fillna("").astype(str).str.strip()
    else:
        df["PAYMENT STATUS"] = ""
    df["PAYMENT STATUS"] = df["PAYMENT STATUS"].replace("", "Not specified")
    df["STATUS"] = df["STATUS"].str.upper().replace("", "NOT SPECIFIED")
    for col in ("SUPPLIER", "DELIVERY LOC", "FORWARDING AGENT"):
        df[col] = df[col].replace("", "Not specified")
    for col in ("VESSEL ETA", "DELIVERED", "RFC", "MT RETURN", "LAST DAY FT DEM", "LAST DAY FT DET"):
        if col not in df.columns:
            df[col] = pd.NaT
        df[col] = pd.to_datetime(df[col], errors="coerce")
    for col in ("DEM TO PAY / DAYS", "DET TO PAY / DAYS"):
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        df[col + " POSITIVE"] = df[col].clip(lower=0)
    df["ETA MONTH"] = df["VESSEL ETA"].dt.strftime("%Y-%m").fillna("No ETA")
    df["DELIVERY GROUP"] = df["STATUS"].apply(lambda s: "Delivered" if s == "DELIVERED" else "Other / Pending")
    return df.reset_index(drop=True)


def plot(fig, height=350):
    fig.update_layout(
        height=height, template="plotly_white", margin=dict(l=8, r=12, t=18, b=10),
        font=dict(family="Arial", color=NAVY), legend_title_text="",
        hoverlabel=dict(bgcolor="white"),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})


def bar_by(df, col, title, top=12):
    st.subheader(title)
    d = df.groupby(col, dropna=False).size().nlargest(top).sort_values().reset_index(name="Containers")
    if d.empty:
        st.info("No records for this selection.")
        return
    fig = px.bar(d, x="Containers", y=col, orientation="h", text="Containers", color_discrete_sequence=[BLUE])
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(yaxis_title=None, xaxis_title="Containers")
    plot(fig, 365)



logo_file = ROOT / "FFM Logo.png"

# Three animated vessels, layered waves, and FFM logo in the main header.
st.markdown("""
<style>
.ship-hero{position:relative;overflow:hidden;height:235px;border-radius:13px;
  background:linear-gradient(105deg,#04264b 0%,#0b4b80 55%,#297aa6 100%);
  margin:4px 0 20px;isolation:isolate;box-shadow:0 8px 25px rgba(11,44,73,.12)}
.ship-hero *{box-sizing:border-box}
.ship-titlebar{position:absolute;top:20px;left:28px;display:flex;align-items:center;gap:20px;z-index:8}
.ship-logo{height:76px;width:76px;object-fit:contain;filter:drop-shadow(0 1px 4px #001a2b)}
.ship-divider{height:70px;width:1px;background:rgba(255,255,255,.45)}
.ship-titlebar h2{color:#fff!important;font-size:clamp(21px,2.5vw,33px);margin:0!important;font-weight:750;line-height:1.2}
.ship-titlebar p{color:#dcecf9!important;font-size:clamp(12px,1.3vw,16px);margin:7px 0 0!important}
.ship-cloud{position:absolute;background:#b1d2e9;opacity:.16;border-radius:80px;filter:blur(2px)}
.ship-cloud.one{width:145px;height:28px;top:44px;right:21%}
.ship-cloud.two{width:100px;height:22px;top:73px;right:5%}
.ship-moon{position:absolute;width:58px;height:58px;border-radius:50%;background:#ffe3a7;right:8%;top:24px;opacity:.82;box-shadow:0 0 45px #f6c67e60}
.ship-horizon{position:absolute;bottom:0;left:0;right:0;height:91px;background:linear-gradient(#1b7fae,#0873a5);z-index:1}
.ship-wave{position:absolute;left:-50%;width:200%;height:48px;z-index:5;opacity:.85;
  background-repeat:repeat-x;background-size:240px 48px;
  background-image:url('data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" width="240" height="48" viewBox="0 0 240 48"%3E%3Cpath d="M0 20 Q30 2 60 20 T120 20 T180 20 T240 20 V48 H0Z" fill="%233b9ec1"/%3E%3C/svg%3E');
  animation:wave-drift 8s linear infinite}
.ship-wave.back{bottom:11px;opacity:.52;animation-duration:12s;z-index:2}
.ship-wave.front{bottom:-16px;opacity:.84;animation-duration:7s;animation-direction:reverse;z-index:6}
@keyframes wave-drift{to{transform:translateX(120px)}}
.ship-boat{position:absolute;z-index:4;width:225px;height:89px;bottom:25px;left:-300px;animation:sail 25s linear infinite}
.ship-boat.small{width:165px;height:66px;bottom:57px;opacity:.88;z-index:3;animation-duration:36s}
.ship-boat.mid{width:185px;height:73px;bottom:37px;animation-duration:30s}
.ship-boat.a{animation-delay:-8s}.ship-boat.b{animation-delay:-22s}.ship-boat.c{animation-delay:-2s}
.ship-boat svg{width:100%;height:100%;display:block;animation:ship-rock 3.5s ease-in-out infinite}
@keyframes sail{from{left:-300px}to{left:calc(100% + 300px)}}
@keyframes ship-rock{0%,100%{transform:translateY(0) rotate(0deg)}50%{transform:translateY(-4px) rotate(.5deg)}}
.ship-cranes{position:absolute;bottom:74px;right:2%;width:245px;height:100px;opacity:.40;z-index:2}
.ship-cranes svg{height:100%;width:100%}
.ship-birds{position:absolute;right:22%;top:38px;z-index:3;opacity:.85}
.ship-birds svg{width:95px;height:36px}
.ship-foam{position:absolute;bottom:21px;left:0;width:100%;height:16px;opacity:.25;z-index:7;
 background:repeating-linear-gradient(100deg,transparent 0 70px,#d3faff 75px 89px,transparent 95px 140px)}
@media(max-width:800px){.ship-hero{height:240px}.ship-titlebar{left:16px;gap:10px}.ship-logo{width:55px;height:55px}.ship-divider{height:54px}.ship-titlebar h2{font-size:22px}.ship-titlebar p{font-size:12px}}
@media(prefers-reduced-motion:reduce){.ship-boat,.ship-boat svg,.ship-wave{animation:none!important}}
</style>
""",unsafe_allow_html=True)

ship_svg = """<svg viewBox="0 0 250 100" xmlns="http://www.w3.org/2000/svg" aria-label="Animated cargo ship sailing to the right">
<g stroke="#15384e" stroke-width="1.3">
<rect x="38" y="19" width="40" height="20" fill="#e5a63e"/><rect x="81" y="19" width="40" height="20" fill="#dc6a4b"/><rect x="124" y="19" width="40" height="20" fill="#51a9a9"/><rect x="167" y="19" width="40" height="20" fill="#d35e63"/>
<rect x="38" y="42" width="40" height="20" fill="#52a6a5"/><rect x="81" y="42" width="40" height="20" fill="#d15d5e"/><rect x="124" y="42" width="40" height="20" fill="#e7ae49"/><rect x="167" y="42" width="40" height="20" fill="#4c98b8"/>
</g><g stroke="#1e4457" opacity=".3"><path d="M49 20v18m10-18v18m10-18v18m23-18v18m10-18v18m10-18v18m23-18v18m10-18v18m10-18v18m23-18v18m10-18v18M49 43v18m10-18v18m10-18v18m23-18v18m10-18v18m10-18v18m23-18v18m10-18v18m10-18v18m23-18v18m10-18v18"/></g>
<rect x="18" y="40" width="18" height="24" fill="#e8f2f8"/><rect x="22" y="44" width="11" height="7" fill="#8cb8d1"/><path d="M10 65 H220 L246 72 L227 88 Q119 100 30 88Z" fill="#e7eff4"/><path d="M13 76 H240 L227 88 Q119 100 30 88Z" fill="#c85356"/><path d="M10 94 Q25 87 40 94 T70 94 T100 94 T130 94 T160 94 T190 94 T220 94" stroke="#b6edff" stroke-width="2.5" fill="none"/>
</svg>"""

logo_html = ""
if logo_file.is_file():
    logo_b64 = base64.b64encode(logo_file.read_bytes()).decode("ascii")
    logo_html = f'<img class="ship-logo" src="data:image/png;base64,{logo_b64}" alt="FFM logo">'

hero_html = f"""<div class="ship-hero">
<div class="ship-cloud one"></div><div class="ship-cloud two"></div><div class="ship-moon"></div>
<div class="ship-birds"><svg viewBox="0 0 100 40" fill="none" stroke="#f0f9ff" stroke-width="2.5" stroke-linecap="round"><path d="M5 12q8 -8 16 0q8 -8 16 0M45 26q6 -6 12 0q6 -6 12 0M70 8q5 -5 10 0q5 -5 10 0"/></svg></div>
<div class="ship-cranes"><svg viewBox="0 0 260 105" fill="none" stroke="#042d54" stroke-width="4"><path d="M12 102V30H100M30 30L85 12M85 12V102M100 30V70M132 102V20H240M150 20L210 4M210 4V102M240 20V76"/><path d="M0 102H260M22 47H92M142 43H225" stroke-width="2"/></svg></div>
<div class="ship-titlebar">{logo_html}<div class="ship-divider"></div><div><h2>Import Delivery Shipments</h2><p>FFM Logistics | Container Control Tower</p></div></div>
<div class="ship-horizon"></div><div class="ship-wave back"></div>
<div class="ship-boat small a">{ship_svg}</div>
<div class="ship-boat mid b">{ship_svg}</div>
<div class="ship-boat c">{ship_svg}</div>
<div class="ship-foam"></div><div class="ship-wave front"></div>
</div>"""
st.markdown(hero_html,unsafe_allow_html=True)

# The dashboard loads the local workbook automatically; no upload panel.
local = next((p for p in LOCAL_FILES if p.is_file()), None)
if local is None:
    st.error("Excel workbook not found. Place IMPORT DELIVERY SHIPMENTS (1).xlsx in the same folder as this script.")
    st.stop()
try:
    df = read_workbook(local.read_bytes())
except Exception as exc:
    st.error(f"Unable to load workbook: {exc}")
    st.stop()

with st.sidebar:
    st.markdown("""<div class="filter-intro"><div class="filter-eyebrow">FFM LOGISTICS · CONTROL TOWER</div>
    <div class="filter-heading">Filters & Search</div>
    <div class="filter-description">Explore your shipment data using the selections below.</div></div>""", unsafe_allow_html=True)
    st.markdown('<div class="filter-section">01 &nbsp; SEARCH CONTAINER</div>', unsafe_allow_html=True)
    search = st.text_input("Container number", placeholder="🔎  Enter container number", key="container", label_visibility="collapsed")
    st.markdown('<div class="filter-section">02 &nbsp; SHIPMENT DETAILS</div>', unsafe_allow_html=True)
    def filter_multi(label, column, key):
        options = sorted(df[column].dropna().astype(str).unique().tolist())
        return st.multiselect(label, options, default=[], key=key, placeholder="All", help=f"Select one or more {label.lower()} values. Leave empty to show all.")
    months = filter_multi("ETA month", "ETA MONTH", "month")
    locations = filter_multi("Delivery location", "DELIVERY LOC", "location")
    agents = filter_multi("Forwarding agent", "FORWARDING AGENT", "agent")
    statuses = filter_multi("Delivery status", "STATUS", "status")
    payments = filter_multi("Payment Status (Excel: STATUS)", "PAYMENT STATUS", "payment_status")
    suppliers = filter_multi("Supplier", "SUPPLIER", "supplier")
    st.markdown('<div class="filter-section">03 &nbsp; DEM / DET</div>', unsafe_allow_html=True)
    only_positive = st.checkbox("Show positive DEM / DET days only", key="positive_only")
    st.markdown('<div class="filter-footer">✓ &nbsp; Filters update the dashboard automatically<br>Empty selections show all records</div>', unsafe_allow_html=True)

f = df.copy()
for col, chosen in (
    ("ETA MONTH", months), ("DELIVERY LOC", locations),
    ("FORWARDING AGENT", agents), ("STATUS", statuses), ("PAYMENT STATUS", payments), ("SUPPLIER", suppliers),
):
    if chosen:
        f = f[f[col].isin(chosen)]
if search.strip():
    f = f[f["CNTNR NO"].str.contains(search.strip(), case=False, regex=False, na=False)]
if only_positive:
    f = f[(f["DEM TO PAY / DAYS POSITIVE"] > 0) | (f["DET TO PAY / DAYS POSITIVE"] > 0)]

total = len(f)
delivered = int((f["DELIVERY GROUP"] == "Delivered").sum())
other = total - delivered
rate = delivered / total if total else 0
dem = f["DEM TO PAY / DAYS POSITIVE"].sum()
det = f["DET TO PAY / DAYS POSITIVE"].sum()
positive_count = int(((f["DEM TO PAY / DAYS POSITIVE"] > 0) | (f["DET TO PAY / DAYS POSITIVE"] > 0)).sum())

kpis = [
    ("Container records", f"{total:,}"),
    ("Delivered", f"{delivered:,}"),
    ("Other / Pending", f"{other:,}"),
    ("Delivery rate", f"{rate:.1%}"),
    ("DEM positive days", f"{dem:,.0f}"),
    ("DET positive days", f"{det:,.0f}"),
]
for c, (name, value) in zip(st.columns(6), kpis):
    c.metric(name, value)

if f.empty:
    st.warning("No records match the selected filters. Clear your selections to view all records.")
    st.stop()

# Payment status summary, based on the second Excel STATUS column (STATUS.1).
payment_counts = f["PAYMENT STATUS"].astype(str).str.strip().value_counts()
to_paid = int(payment_counts.get("To Paid", 0))
no_payment = int(payment_counts.get("No", 0))
other_payment = total - to_paid - no_payment
st.markdown("### Payment Status Overview")
pa, pb, pc = st.columns(3)
pa.metric("🔴 To Paid", f"{to_paid:,}")
pb.metric("🟢 No", f"{no_payment:,}")
pc.metric("⚪ Other / Not specified", f"{other_payment:,}")

overview, risks, records = st.tabs(["📊 Overview", "⚠️ DEM / DET analysis", "🔎 Shipment explorer"])

with overview:
    a, b = st.columns([1.65, 1])
    with a:
        st.subheader("Monthly shipment trend")
        trend = f[f["ETA MONTH"] != "No ETA"].groupby(["ETA MONTH", "DELIVERY GROUP"]).size().reset_index(name="Containers")
        if not trend.empty:
            fig = px.bar(trend, x="ETA MONTH", y="Containers", color="DELIVERY GROUP",
                         barmode="group", color_discrete_map={"Delivered": TEAL, "Other / Pending": ORANGE})
            fig.update_layout(xaxis_title=None, yaxis_title="Containers")
            st.plotly_chart(fig, use_container_width=True, config={"displaylogo":False})
        else:
            st.info("No valid ETA dates.")
    with b:
        st.subheader("Delivery status")
        status = f.groupby("DELIVERY GROUP").size().reset_index(name="Containers")
        fig = px.pie(status, names="DELIVERY GROUP", values="Containers", hole=.68,
                     color="DELIVERY GROUP", color_discrete_map={"Delivered": TEAL, "Other / Pending": ORANGE})
        fig.update_traces(textinfo="percent")
        st.plotly_chart(fig, use_container_width=True, config={"displaylogo":False})
    st.subheader("Payment status by container")
    payment_summary = f.groupby("PAYMENT STATUS", dropna=False).size().reset_index(name="Containers")
    payment_summary["PAYMENT STATUS"] = payment_summary["PAYMENT STATUS"].astype(str)
    fig = px.bar(
        payment_summary.sort_values("Containers", ascending=True),
        x="Containers", y="PAYMENT STATUS", orientation="h", color="PAYMENT STATUS",
        text="Containers",
        color_discrete_map={"To Paid": "#D83A42", "No": "#20A66B", "Not specified": "#9AA9B8"},
    )
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(showlegend=False, xaxis_title="Containers", yaxis_title=None,
                      margin=dict(l=10, r=45, t=10, b=10), height=max(230, 65 * len(payment_summary)))
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo":False})
    a, b = st.columns(2)
    for panel, col, title, key in ((a,"DELIVERY LOC","By delivery location","cross_location"),
                                   (b,"FORWARDING AGENT","By forwarding agent","cross_agent")):
        with panel:
            st.subheader(title)
            # Hide unassigned agents from this chart ONLY; retain their rows in f,
            # so all KPIs, other charts, filters, and Shipment Explorer remain accurate.
            chart_data = f
            if col == "FORWARDING AGENT":
                agent_names = f[col].astype("string").str.strip()
                chart_data = f[agent_names.notna() & ~agent_names.str.casefold().isin(
                    ["", "not specified", "nan", "none", "<na>"]
                )]
            d = chart_data.groupby(col, dropna=False).size().nlargest(12).sort_values().reset_index(name="Containers")
            if not d.empty:
                chart = px.bar(d, x="Containers", y=col, orientation="h", text="Containers", color_discrete_sequence=[BLUE])
                chart.update_traces(textposition="outside", cliponaxis=False)
                chart.update_layout(height=365, template="plotly_white", yaxis_title=None, xaxis_title="Containers")
                st.plotly_chart(chart, use_container_width=True, config={"displaylogo":False})

with risks:
    a, b, c = st.columns(3)
    a.metric("Records with positive DEM / DET", f"{positive_count:,}")
    b.metric("Positive DEM days", f"{dem:,.0f}")
    c.metric("Positive DET days", f"{det:,.0f}")
    st.caption("Positive days are recorded source values; these are NOT financial charges or verified overdue cases.")
    risk = f.groupby("FORWARDING AGENT")[["DEM TO PAY / DAYS POSITIVE", "DET TO PAY / DAYS POSITIVE"]].sum()
    risk["TOTAL"] = risk.sum(axis=1)
    risk = risk.nlargest(12, "TOTAL").sort_values("TOTAL").drop(columns="TOTAL").reset_index()
    melted = risk.melt(id_vars="FORWARDING AGENT", var_name="Type", value_name="Days")
    st.subheader("DEM / DET days by forwarding agent")
    fig = px.bar(melted, x="Days", y="FORWARDING AGENT", color="Type", orientation="h",
                 barmode="group", color_discrete_map={"DEM TO PAY / DAYS POSITIVE": "#155CAA", "DET TO PAY / DAYS POSITIVE": "#F0A12B"})
    fig.update_layout(yaxis_title=None, legend_title_text="Days type", legend=dict(orientation="h", y=1.13, x=0))
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo":False})
    st.subheader("Records with positive DEM / DET")
    risk_rows = f[(f["DEM TO PAY / DAYS POSITIVE"] > 0) | (f["DET TO PAY / DAYS POSITIVE"] > 0)]
    cols = ["CNTNR NO", "SUPPLIER", "DELIVERY LOC", "FORWARDING AGENT", "STATUS", "PAYMENT STATUS",
            "DEM TO PAY / DAYS POSITIVE", "DET TO PAY / DAYS POSITIVE", "VESSEL ETA"]
    st.dataframe(risk_rows[cols], hide_index=True, use_container_width=True, height=340)

with records:
    st.subheader("Shipment explorer")
    st.caption("Select a container to inspect its details, or browse the shipment records below.")
    show = ["SUPPLIER", "BL NO", "CNTNR NO", "PO NO", "STATUS", "PAYMENT STATUS",
            "ITEMS", "DELIVERY LOC", "VESSEL ETA", "DELIVERED", "FORWARDING AGENT",
            "DEM TO PAY / DAYS", "DET TO PAY / DAYS", "REMARKS"]
    options = sorted(f["CNTNR NO"].dropna().unique().tolist())
    chosen = st.selectbox("Inspect container", options, index=None, placeholder="Select a container...")
    if chosen:
        row = f.loc[f["CNTNR NO"] == chosen].iloc[0]
        left, right = st.columns(2)
        with left:
            st.write("**Container:**", row["CNTNR NO"])
            st.write("**Supplier:**", row["SUPPLIER"])
            st.write("**Delivery status:**", row["STATUS"])
            st.write("**Payment Status (No / To Paid):**", row["PAYMENT STATUS"])
            st.write("**Delivery location:**", row["DELIVERY LOC"])
        with right:
            st.write("**Forwarding agent:**", row["FORWARDING AGENT"])
            st.write("**Vessel ETA:**", row["VESSEL ETA"])
            st.write("**DEM days:**", row["DEM TO PAY / DAYS"])
            st.write("**DET days:**", row["DET TO PAY / DAYS"])
        st.write("**Remarks:**", row["REMARKS"] or "—")
    st.divider()
    st.subheader("Shipment records")
    st.dataframe(f[show], use_container_width=True, hide_index=True, height=440)
    st.download_button("Export filtered data (CSV)", data=f[show].to_csv(index=False).encode("utf-8-sig"),
                       file_name="shipment_filtered.csv", mime="text/csv")

st.caption("Source data may contain missing or inconsistent statuses. 'Other / Pending' includes every status other than exactly DELIVERED.")
