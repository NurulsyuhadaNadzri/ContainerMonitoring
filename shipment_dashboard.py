import io
import base64
import html
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

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

st.markdown('\n<style>\n/* Wider, readable, compact filters sidebar */\n[data-testid="stSidebar"] {\n  width: 345px !important;\n  min-width: 345px !important;\n}\n[data-testid="stSidebar"] > div:first-child {\n  width: 345px !important;\n}\n[data-testid="stSidebar"] [data-testid="stSidebarContent"] {\n  padding: 12px 17px 20px !important;\n  overflow-x: hidden !important;\n}\n[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {\n  gap: .48rem !important;\n}\n[data-testid="stSidebar"] .filter-intro {\n  padding: 17px 17px 15px !important;\n  margin: 0 0 8px !important;\n  border-radius: 11px !important;\n}\n[data-testid="stSidebar"] .filter-heading {\n  font-size: 21px !important;\n}\n[data-testid="stSidebar"] .filter-description {\n  font-size: 11px !important;\n  margin-top: 5px !important;\n}\n[data-testid="stSidebar"] .filter-section {\n  font-size: 10px !important;\n  letter-spacing: .8px !important;\n  margin: 7px 0 5px !important;\n  padding: 7px 0 !important;\n}\n[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {\n  font-size: 12px !important;\n  white-space: normal !important;\n  overflow-wrap: anywhere !important;\n  line-height: 1.35 !important;\n}\n[data-testid="stSidebar"] [data-baseweb="select"] > div,\n[data-testid="stSidebar"] [data-testid="stTextInput"] input {\n  min-height: 38px !important;\n}\n[data-testid="stSidebar"] [data-testid="stWidgetLabel"] {\n  margin-bottom: 3px !important;\n}\n@media(max-width: 760px) {\n  [data-testid="stSidebar"],[data-testid="stSidebar"] > div:first-child {\n    width: min(345px, 90vw) !important;\n    min-width: 0 !important;\n  }\n}\n</style>\n', unsafe_allow_html=True)

st.markdown("""
<style>
/* Keep sidebar open; hide only its collapse/hide button. */
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"],
[data-testid="stSidebarHeader"] button,
button[aria-label="Close sidebar"],
button[aria-label="Collapse sidebar"] {display:none!important;}
/* Page navigation lives in the sidebar, not as a crowded horizontal tab bar. */
.nav-eyebrow {font-size:10px;font-weight:800;letter-spacing:1.5px;color:#5682a1;margin:8px 0 7px}
.nav-heading {font-size:20px;font-weight:800;color:#10365e;letter-spacing:-.5px;margin-bottom:15px}
.nav-divider {height:1px;background:#dce8f2;margin:18px 0 13px}
.page-eyebrow {font-size:11px;letter-spacing:1.3px;font-weight:750;color:#52718d;margin:8px 0 13px}
[data-testid="stSidebar"] [role="radiogroup"] {gap:5px!important}
[data-testid="stSidebar"] [role="radiogroup"] label {border-radius:9px;padding:8px 10px!important;
  background:transparent;border:1px solid transparent;transition:background .15s ease}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
  background:#e3f0fa!important;border-color:#b9d9ee!important}
[data-testid="stSidebar"] [role="radiogroup"] label p {font-size:13px!important;font-weight:650!important;color:#234b6d!important}
[data-testid="stSidebar"] [data-testid="stExpander"] {background:#fff;border:1px solid #dce8f2;border-radius:11px}
[data-testid="stSidebar"] [data-testid="stExpander"] summary {font-weight:700;color:#174b78}
</style>
""", unsafe_allow_html=True)

# Subtle AI-inspired motion design: visual polish only, no artificial intelligence claims.
st.markdown("""
<style>
:root { --motion-navy:#10365e; --motion-cyan:#53bdd7; --motion-green:#238b57; }
/* Gentle staggered entrance on each selected page. */
@keyframes motion-enter {from{opacity:0;transform:translateY(9px)}to{opacity:1;transform:translateY(0)}}
@keyframes motion-sheen {from{transform:translateX(-130%) skewX(-20deg)}to{transform:translateX(280%) skewX(-20deg)}}
@keyframes motion-grow {from{transform:scaleX(0)}to{transform:scaleX(1)}}
@keyframes motion-pulse {0%,100%{box-shadow:0 0 0 0 rgba(53,163,198,.18)}50%{box-shadow:0 0 0 7px rgba(53,163,198,0)}}
@keyframes motion-float {0%,100%{transform:translateY(0)}50%{transform:translateY(-5px)}}
@keyframes motion-scan {0%{left:-35%}100%{left:120%}}
/* Make navigation feel responsive while keeping its position unchanged. */
[data-testid="stSidebar"] [role="radiogroup"] label {transition:background-color .22s ease,border-color .22s ease,transform .22s ease!important}
[data-testid="stSidebar"] [role="radiogroup"] label:hover {transform:translateX(3px);background:#eaf5fc!important;border-color:#d0e6f5!important}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {box-shadow:inset 3px 0 0 #2794c3,0 2px 9px rgba(19,81,123,.06)}
/* Small lift, never rotating/flashing or distracting. */
[data-testid="stPlotlyChart"],div[data-testid="stMetric"] {transition:box-shadow .24s ease,transform .24s ease,border-color .24s ease}
[data-testid="stPlotlyChart"]:hover,div[data-testid="stMetric"]:hover {transform:translateY(-2px);box-shadow:0 10px 24px rgba(13,60,100,.11)!important;border-color:#c8e1ed!important}
/* Corporate header: very light data-grid and scanning glow. */
.ship-hero::before {content:"";position:absolute;inset:0;z-index:1;pointer-events:none;opacity:.18;
 background-image:linear-gradient(rgba(161,225,247,.18) 1px,transparent 1px),linear-gradient(90deg,rgba(161,225,247,.18) 1px,transparent 1px);
 background-size:42px 42px;mask-image:linear-gradient(90deg,transparent 22%,black 100%)}
.ship-hero::after {content:"";position:absolute;top:0;bottom:0;left:-35%;width:30%;z-index:2;pointer-events:none;
 background:linear-gradient(90deg,transparent,rgba(170,228,249,.09),transparent);animation:motion-scan 18s linear infinite}
.ship-moon {animation:motion-pulse 5s ease-in-out infinite}
.ship-titlebar {animation:motion-enter .65s ease both}
/* Management overview entrance and animated completion track. */
.ov-header {animation:motion-enter .45s ease both}
.ov-summary {animation:motion-enter .6s .08s ease both}
.ov-stat {transition:background-color .25s ease}
.ov-stat:hover {background:#f6fbff}
.ov-progress-wrap {animation:motion-enter .65s .12s ease both;overflow:hidden}
.ov-progress-fill,.ov-progress-rest {transform-origin:left center;animation:motion-grow 1.1s .18s cubic-bezier(.2,.8,.2,1) both}
.ov-progress-track {position:relative}
.ov-progress-track::after {content:"";position:absolute;inset:0 auto 0 0;width:30%;pointer-events:none;
 background:linear-gradient(90deg,transparent,rgba(255,255,255,.32),transparent);animation:motion-sheen 4.5s 1s ease-in-out infinite}
/* Noticeable but polished motion for the entire analytics experience. */
@keyframes ocean-chart-reveal {from{clip-path:inset(0 100% 0 0);opacity:.6}to{clip-path:inset(0 0 0 0);opacity:1}}
@keyframes ocean-rise {from{opacity:0;transform:translateY(22px)}to{opacity:1;transform:translateY(0)}}
@keyframes ocean-bar-wipe {from{clip-path:inset(0 100% 0 0)}to{clip-path:inset(0 0 0 0)}}
@keyframes ocean-glow {0%,100%{filter:drop-shadow(0 0 0 rgba(36,150,180,0))}50%{filter:drop-shadow(0 0 5px rgba(36,150,180,.22))}}
@keyframes ocean-ship-bob {0%,100%{translate:0 0}50%{translate:0 -4px}}
/* Plotly itself stays fully interactive; reveal the chart without replacing Plotly. */
[data-testid="stPlotlyChart"] {animation:ocean-chart-reveal 1.15s cubic-bezier(.16,.9,.2,1) both;transform-origin:left center}
[data-testid="stPlotlyChart"]:nth-of-type(2n) {animation-delay:.12s}
[data-testid="stPlotlyChart"]:nth-of-type(3n) {animation-delay:.2s}
.ov-stat {animation:ocean-rise .7s cubic-bezier(.2,.85,.2,1) both}
.ov-stat:nth-child(1) {animation-delay:.08s}
.ov-stat:nth-child(2) {animation-delay:.18s}
.ov-stat:nth-child(3) {animation-delay:.28s}
.ov-stat:nth-child(4) {animation-delay:.38s}
.ov-stat:nth-child(5) {animation-delay:.48s}
.ov-progress-track {animation:ocean-glow 3s ease-in-out 2}
.ov-progress-fill {animation:motion-grow 1.45s .25s cubic-bezier(.15,.85,.2,1) both}
.ov-progress-rest {animation:motion-grow 1.45s .4s cubic-bezier(.15,.85,.2,1) both}
/* Keep movement smooth and avoid page layout shifts. */
[data-testid="stDataFrame"] {animation:ocean-rise .65s .12s ease-out both}
[data-testid="stMetric"] {animation:ocean-rise .6s ease-out both}
/* Reduce motion for accessibility, including the pre-existing animated ships. */
@media(prefers-reduced-motion:reduce){
 .ship-hero::after,.ship-moon,.ship-titlebar,.ov-header,.ov-summary,.ov-progress-wrap,
 .ov-progress-fill,.ov-progress-rest,.ov-progress-track::after, .ov-stat,
 [data-testid="stPlotlyChart"],[data-testid="stDataFrame"],[data-testid="stMetric"] {animation:none!important}
 [data-testid="stPlotlyChart"],div[data-testid="stMetric"],[data-testid="stSidebar"] [role="radiogroup"] label {transition:none!important;transform:none!important}
}
</style>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner="Reading shipment workbook...")
def read_workbook(data: bytes) -> pd.DataFrame:
    excel = pd.ExcelFile(io.BytesIO(data))
    # Combine both worksheets; the source label is kept internally but not displayed.
    sheets = ["MAY-JULY", "BRANCHES"]
    missing = [name for name in sheets if name not in excel.sheet_names]
    if missing:
        raise ValueError(f"Required worksheet(s) not found: {', '.join(missing)}")
    frames = []
    for sheet in sheets:
        part = pd.read_excel(
            excel, sheet_name=sheet,
            dtype={"CNTNR NO": str, "BL NO": str, "PO NO": str},
        )
        part.columns = [str(c).strip().upper() for c in part.columns]
        part = part.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
        part["SOURCE SHEET"] = sheet
        frames.append(part)
    df = pd.concat(frames, ignore_index=True, sort=False)
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
        df[col] = pd.to_numeric(df[col], errors="coerce")
        df[col + " POSITIVE"] = df[col].clip(lower=0).fillna(0)
    for col in ("DEM FREE TIME", "DET FREE TIME"):
        if col not in df.columns:
            df[col] = pd.NA
        df[col] = pd.to_numeric(df[col], errors="coerce")
    # Excel formula errors can appear as dates around 1899/1900; never treat as real deadlines.
    for col in ("VESSEL ETA", "DELIVERED", "RFC", "MT RETURN", "LAST DAY FT DEM", "LAST DAY FT DET"):
        df.loc[df[col] < pd.Timestamp("2000-01-01"), col] = pd.NaT
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

# Navigation is separate from filtering: choose a page first, then optionally
# expand the filters. This avoids seven cramped tabs across the page header.
PAGES = [
    ("📦 Container Monitoring", "monitor"),
    ("📊 Overview", "overview"),
    ("⚠️ DEM / DET Analysis", "risks"),
    ("⏱️ Lead-Time Analysis", "leadtime"),
    ("📈 Performance Comparison", "comparison"),
    ("🔍 Root-Cause Analysis", "rootcause"),
    ("🔎 Shipment Explorer", "records"),
]
with st.sidebar:
    st.markdown('<div class="nav-eyebrow">FFM · LOGISTICS INTELLIGENCE</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-heading">Dashboard Navigation</div>', unsafe_allow_html=True)
    page_label = st.radio(
        "Dashboard pages", [label for label, _ in PAGES],
        key="active_dashboard_page", label_visibility="collapsed",
    )
    active_page = dict(PAGES)[page_label]
    st.markdown('<div class="nav-divider"></div>', unsafe_allow_html=True)
    with st.expander("🔎 Filters & Search", expanded=False):
        search = st.text_input("Container number", placeholder="Search container number", key="container")
        def filter_multi(label, column, key):
            options = sorted(df[column].dropna().astype(str).unique().tolist())
            return st.multiselect(label, options, default=[], key=key, placeholder="All")
        months = filter_multi("ETA month", "ETA MONTH", "month")
        locations = filter_multi("Delivery location", "DELIVERY LOC", "location")
        agents = filter_multi("Forwarding agent", "FORWARDING AGENT", "agent")
        statuses = filter_multi("Delivery status", "STATUS", "status")
        payments = filter_multi("Payment status", "PAYMENT STATUS", "payment_status")
        suppliers = filter_multi("Supplier", "SUPPLIER", "supplier")
        only_positive = st.checkbox("Positive DEM / DET days only", key="positive_only")
    active_filter_count = sum(bool(v) for v in [search.strip(), months, locations, agents, statuses, payments, suppliers, only_positive])
    if active_filter_count:
        st.caption(f"{active_filter_count} active filter(s) · Applies to all pages")
    else:
        st.caption("All shipments · Filters apply to every page")


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
]
for c, (name, value) in zip(st.columns(4), kpis):
    c.metric(name, value)

if f.empty:
    st.warning("No records match the selected filters. Clear your selections to view all records.")
    st.stop()

# Reusable, validated stage intervals. Negative durations are data-quality issues,
# not negative process lead times.
def stage_days(frame, start, end):
    days = (frame[end] - frame[start]).dt.total_seconds() / 86400
    return days.where(days >= 0)


def category_summary(frame, field):
    d = frame.copy()
    d["Recorded chargeable days"] = d["DEM TO PAY / DAYS POSITIVE"] + d["DET TO PAY / DAYS POSITIVE"]
    d["Has recorded chargeable days"] = d["Recorded chargeable days"] > 0
    d["ETA to Delivered (days)"] = stage_days(d, "VESSEL ETA", "DELIVERED")
    d["Delivered to RFC (days)"] = stage_days(d, "DELIVERED", "RFC")
    d["RFC to Empty Return (days)"] = stage_days(d, "RFC", "MT RETURN")
    result = d.groupby(field, dropna=False).agg(
        Containers=("CNTNR NO", "size"),
        Chargeable_containers=("Has recorded chargeable days", "sum"),
        Recorded_chargeable_days=("Recorded chargeable days", "sum"),
        Avg_ETA_to_Delivery=("ETA to Delivered (days)", "mean"),
        Avg_Delivery_to_RFC=("Delivered to RFC (days)", "mean"),
        Avg_RFC_to_Return=("RFC to Empty Return (days)", "mean"),
    ).reset_index()
    result["Chargeable rate (%)"] = 100 * result["Chargeable_containers"] / result["Containers"]
    return result.sort_values("Recorded_chargeable_days", ascending=False)


st.markdown(f'<div class="page-eyebrow">FFM LOGISTICS / {html.escape(page_label.replace("📦 ", "").replace("📊 ", ""))}</div>', unsafe_allow_html=True)

# Optional vessel register: map BL numbers to actual vessel names.
# The original Excel has VESSEL ETA but does not have a vessel-name column.
VESSEL_MAP_FILE = ROOT / "vessel_mapping.csv"
if VESSEL_MAP_FILE.exists():
    try:
        vessel_map = pd.read_csv(VESSEL_MAP_FILE, dtype=str).fillna("")
        vessel_map.columns = vessel_map.columns.str.strip().str.upper()
        if not {"BL NO", "VESSEL NAME"}.issubset(vessel_map.columns):
            raise ValueError("vessel_mapping.csv requires BL NO and VESSEL NAME columns")
        vessel_map["BL NO"] = vessel_map["BL NO"].str.strip().str.upper()
        vessel_map["VESSEL NAME"] = vessel_map["VESSEL NAME"].str.strip()
        mapping = (vessel_map[vessel_map["VESSEL NAME"] != ""]
                   .drop_duplicates("BL NO").set_index("BL NO")["VESSEL NAME"].to_dict())
    except Exception as exc:
        st.error(f"Vessel mapping error: {exc}")
        mapping = {}
else:
    mapping = {}

# Vessel identity is not inferable from ETA. Use a clearly named BL-based
# provisional shipment group when no actual vessel name is mapped.
working = f.copy()
working["BL KEY"] = working["BL NO"].astype(str).str.strip().str.upper()
working["VESSEL NAME"] = working["BL KEY"].map(mapping)
working["VESSEL GROUP"] = working.apply(
    lambda r: r["VESSEL NAME"] if pd.notna(r["VESSEL NAME"]) and r["VESSEL NAME"]
    else ("BL " + r["BL KEY"] if r["BL KEY"] and r["BL KEY"].lower() not in ("nan", "none")
          else "Unassigned · " + str(r["CNTNR NO"])), axis=1)

if active_page == "monitor":
    st.subheader("Vessel Monitoring")
    choices = sorted(working["VESSEL GROUP"].unique().tolist())
    selected_vessel = st.selectbox("Vessel / BL shipment group", choices, index=None,
                                   placeholder="Search vessel or BL number...", key="vessel_monitor")
    if selected_vessel:
        subset = working[working["VESSEL GROUP"] == selected_vessel].copy()
        # Deduplicate overlapping records from MAY-JULY and BRANCHES. Keep the
        # furthest confirmed milestones, without counting a container twice.
        today = pd.Timestamp.now().normalize()
        stage_cols = ["VESSEL ETA", "DELIVERED", "RFC", "MT RETURN"]
        def stage_done(row, col):
            value = row[col]
            return pd.notna(value) and value.normalize() <= today
        by_container = subset.groupby("CNTNR NO", dropna=False)
        rows = []
        for container, group in by_container:
            dates = {}
            for col in stage_cols:
                valid = pd.to_datetime(group[col], errors="coerce").dropna()
                valid = valid[valid.dt.normalize() <= today]
                dates[col] = valid.min() if not valid.empty else pd.NaT
            # R/S are signed deadline differences: negative = early, zero = on time,
            # positive = days late. Keep missing values unknown, not zero.
            delay_values = {}
            for delay_col, output_col in [
                ("DEM TO PAY / DAYS", "DEM Delay Days"),
                ("DET TO PAY / DAYS", "DET Delay Days"),
            ]:
                vals = pd.to_numeric(group[delay_col], errors="coerce").dropna()
                # If duplicate sheets disagree, use the most conservative recorded value.
                delay_values[output_col] = vals.max() if not vals.empty else float("nan")
            rows.append({"Container":str(container), **dates, **delay_values})
        progress = pd.DataFrame(rows)
        total_containers = len(progress)
        counts = {col:int(progress[col].notna().sum()) for col in stage_cols}
        delivered_count = counts["DELIVERED"]
        rfc_count = counts["RFC"]
        returned_count = counts["MT RETURN"]
        st.markdown(f"""<div style="background:white;border:1px solid #dce7f2;border-radius:14px;
            padding:18px 22px;margin:10px 0 16px;box-shadow:0 5px 20px #173b5910">
            <div style="font-size:11px;letter-spacing:1px;color:#5e7b95;font-weight:800">VESSEL PROGRESS</div>
            <div style="font-size:22px;font-weight:800;color:#12395d;margin-top:6px">{html.escape(selected_vessel)}</div>
            <div style="font-size:12px;color:#668198;margin-top:5px">{total_containers} unique containers</div>
            </div>""", unsafe_allow_html=True)
        st.markdown("#### Container Journey")
        stages = [("ETA", "VESSEL ETA"), ("Delivered", "DELIVERED"),
                  ("RFC · Empty", "RFC"), ("Empty Return", "MT RETURN")]
        # Move the vessel along the journey only as each entire stage completes.
        # E.g. 5/12 delivered => 5/12 of the distance ETA -> Delivered.
        # Once 12/12 delivered, RFC progress moves it toward RFC, and so on.
        delivery_fraction = counts["DELIVERED"] / total_containers if total_containers else 0.0
        rfc_fraction = counts["RFC"] / total_containers if total_containers else 0.0
        return_fraction = counts["MT RETURN"] / total_containers if total_containers else 0.0
        if delivery_fraction < 1:
            journey_progress = delivery_fraction
            current_phase = "ETA → Delivered"
        elif rfc_fraction < 1:
            journey_progress = 1 + rfc_fraction
            current_phase = "Delivered → RFC"
        else:
            journey_progress = 2 + return_fraction
            current_phase = "RFC → Empty Return" if return_fraction < 1 else "All stages complete"
        journey_progress = max(0.0, min(3.0, journey_progress))
        ship_position = 12.5 + 25.0 * journey_progress
        track_fill = 100.0 * journey_progress / 3.0
        ship_svg = ("<svg viewBox='0 0 64 42' width='58' height='39' aria-label='Container vessel' "
                    "xmlns='http://www.w3.org/2000/svg'><path d='M5 26h54l-10 11H16Z' fill='#143c62'/>"
                    "<rect x='11' y='15' width='14' height='10' rx='1' fill='#e7a332'/>"
                    "<rect x='27' y='15' width='14' height='10' rx='1' fill='#66a8cb'/>"
                    "<rect x='19' y='4' width='14' height='10' rx='1' fill='#e7a332'/>"
                    "<path d='M8 39q7-4 14 0t14 0t14 0t12 0' stroke='#6db5d8' "
                    "stroke-width='2' fill='none'/></svg>")
        progress_cells = []
        for idx, (label, col) in enumerate(stages):
            completed = counts[col]
            pct = 100 * completed / total_containers if total_containers else 0
            latest_date = progress[col].max()
            date_text = latest_date.strftime("%d %b %Y") if pd.notna(latest_date) else "—"
            done_class = " is-complete" if completed == total_containers and total_containers else ""
            progress_cells.append(f"""<div class="vessel-phase{done_class}">
                <div class="vessel-dot"></div>
                <div class="vessel-label">{html.escape(label)}</div>
                <div class="vessel-count">{completed} / {total_containers}</div>
                <div class="vessel-small">{pct:.0f}% completed</div>
                <div class="vessel-small">Latest: {date_text}</div>
                </div>""")
        # Actual vessel milestone elapsed days, never a statistical average.
        # Each stage is measured from the previous stage's latest recorded date;
        # ETA is the earliest scheduled ETA across selected containers.
        milestone_dates = [progress["VESSEL ETA"].min(),
                           progress["DELIVERED"].max(),
                           progress["RFC"].max(),
                           progress["MT RETURN"].max()]
        lead_labels = []
        for idx in range(3):
            start_date, end_date = milestone_dates[idx:idx + 2]
            segment_complete = total_containers > 0 and counts[stages[idx + 1][1]] == total_containers
            if pd.notna(start_date) and pd.notna(end_date) and end_date >= start_date:
                elapsed_days = (end_date.normalize() - start_date.normalize()).days
                duration_text = f"{elapsed_days} days"
                detail_text = "" if segment_complete else "so far"
            else:
                duration_text = "Pending"
                detail_text = ""
            lead_labels.append(f"""<div class="journey-duration duration-{idx}">
                <strong>{html.escape(duration_text)}</strong>
                <span>{html.escape(detail_text)}</span>
                </div>""")
        all_returned = total_containers > 0 and counts["MT RETURN"] == total_containers
        if all_returned and pd.notna(milestone_dates[0]) and pd.notna(milestone_dates[3]) and milestone_dates[3] >= milestone_dates[0]:
            total_days = (milestone_dates[3].normalize() - milestone_dates[0].normalize()).days
            overall_text = f"{total_days} days"
            overall_status = "Journey Completed"
        else:
            overall_text = "In progress"
            overall_status = "Awaiting completion of all containers"
        components.html(f"""<!doctype html><html><head><meta charset="utf-8"><style>
          html,body{{margin:0;padding:0;background:transparent;font-family:Arial,sans-serif;}}
          @keyframes track-reveal{{from{{clip-path:inset(0 100% 0 0)}}to{{clip-path:inset(0 0 0 0)}}}}
          @keyframes ship-appear{{from{{opacity:0;transform:translate(-50%,-7px)}}to{{opacity:1;transform:translate(-50%,0)}}}}
          @media(prefers-reduced-motion:reduce){{.vessel-track-fill,.vessel-ship{{animation:none!important;transition:none!important}}}}
          .vessel-journey{{position:relative;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));
            margin:10px 0 0;padding:10px 0 0;background:white;border:1px solid #dce7f2;
            border-radius:14px;box-shadow:0 5px 20px #173b5910;overflow:hidden}}
          .vessel-track{{position:absolute;top:132px;left:12.5%;width:75%;height:3px;
            background:#c8d9e7;z-index:0}}
          .vessel-track-fill{{height:100%;width:{track_fill:.3f}%;background:#155caa;transition:width .8s ease;animation:track-reveal 1s ease-out both}}
          @keyframes vessel-line-shimmer{{0%{{opacity:.78}}50%{{opacity:1}}100%{{opacity:.78}}}}
          .vessel-track-fill{{animation:track-reveal 1.25s cubic-bezier(.16,.85,.2,1) both,vessel-line-shimmer 3.5s ease-in-out 1.25s infinite}}
          .vessel-phase{{position:relative;text-align:center;padding:160px 5px 24px;z-index:1}}
          .vessel-ship{{position:absolute;top:60px;left:{ship_position:.3f}%;
            transform:translateX(-50%);z-index:5;line-height:0;transition:left .8s ease;animation:ship-appear .7s ease-out both}}
          .vessel-dot{{position:absolute;top:124px;left:50%;transform:translateX(-50%);
            height:16px;width:16px;background:white;border:3px solid #8daecb;
            border-radius:50%;z-index:2;box-sizing:border-box}}
          .vessel-phase.is-complete .vessel-dot{{background:#155caa;border-color:#155caa;
            box-shadow:0 0 0 3px #e2effa}}
          .vessel-label{{font-size:14px;font-weight:700;color:#244968;line-height:1.5;min-height:22px}}
          .vessel-count{{font-size:20px;font-weight:800;color:#12395d;margin:6px 0}}
          .vessel-small{{font-size:11px;color:#668198;margin-top:4px}}
          .journey-duration{{position:absolute;top:104px;width:24%;text-align:center;z-index:3;
            color:#204c71;display:flex;flex-direction:column;gap:3px;pointer-events:none}}
          .journey-duration strong{{font-size:12px;line-height:1.1;white-space:nowrap}}
          .journey-duration span{{font-size:10px;color:#6a8295;line-height:1.1;white-space:nowrap}}
          .duration-0{{left:13%}}.duration-1{{left:38%}}.duration-2{{left:63%}}
          .journey-note{{font-size:11px;color:#49677e;text-align:center;padding:8px 0 3px}}
          .overall-lead{{margin:6px 0 0;padding:10px 14px;background:#f3f7fb;
            border:1px solid #dce7f2;border-radius:9px;color:#244968;font-size:12px;text-align:center}}
          .overall-lead strong{{color:#12395d;font-size:15px}}
          </style></head><body><div class="vessel-journey">
          <div class="vessel-track"><div class="vessel-track-fill"></div></div>
          {''.join(lead_labels)}
          <div class="vessel-ship">{ship_svg}</div>{''.join(progress_cells)}</div>
          <div class="overall-lead">Total Vessel Journey · <strong>{html.escape(overall_text)}</strong>
            &nbsp; · &nbsp; {html.escape(overall_status)}</div>
          <div class="journey-note">{html.escape(current_phase)} · Ship moves with milestone completion</div>
          </body></html>""", height=375, scrolling=False)

        # The separate Daily Delivery table is omitted here to avoid duplicating
        # container-journey information. The detailed container table remains below.
        st.markdown("#### Delivery & Return Balances")
        bal1,bal2,bal3=st.columns(3)
        bal1.metric("Not yet delivered",max(0,total_containers-delivered_count))
        bal2.metric("Delivered, awaiting RFC",max(0,delivered_count-rfc_count))
        bal3.metric("RFC, awaiting return",max(0,rfc_count-returned_count))
        st.markdown("#### DEM / DET Deadline Performance")
        # Completion status is only meaningful once the respective milestone is recorded.
        # DEM is provisionally tied to Delivered; DET to MT RETURN, following the workbook.
        def deadline_result(days, completed):
            if pd.isna(completed) or pd.isna(days):
                return "Not confirmed"
            if days > 0:
                return "Late"
            if days < 0:
                return "Early"
            return "On time"

        progress["DEM Status"] = progress.apply(
            lambda r: deadline_result(r["DEM Delay Days"], r["DELIVERED"]), axis=1)
        progress["DET Status"] = progress.apply(
            lambda r: deadline_result(r["DET Delay Days"], r["MT RETURN"]), axis=1)
        dem_col, det_col = st.columns(2)
        for kind, area in (("DEM", dem_col), ("DET", det_col)):
            stat_col = f"{kind} Status"
            days_col = f"{kind} Delay Days"
            early_or_ontime = int(progress[stat_col].isin(["Early", "On time"]).sum())
            late = int((progress[stat_col] == "Late").sum())
            unknown = int((progress[stat_col] == "Not confirmed").sum())
            delay_days = int(progress.loc[progress[stat_col] == "Late", days_col].sum())
            with area:
                st.markdown(f"**{kind}**")
                st.caption(f"{early_or_ontime} early / on time · {unknown} not confirmed")
                st.metric("Containers late", late)
                st.caption(f"{delay_days} total days late")

        # One container-level table replaces separate DEM/DET, daily milestone,
        # and vessel status tables. The journey already shows milestone totals.
        with st.expander("Container details · dates and DEM / DET", expanded=False):
            detail = progress[["Container", "VESSEL ETA", "DELIVERED",
                               "RFC", "MT RETURN", "DEM Delay Days", "DEM Status",
                               "DET Delay Days", "DET Status"]].copy()
            detail["Current stage"] = progress.apply(
                lambda r: "Empty returned" if pd.notna(r["MT RETURN"]) else
                          "Unloaded / RFC" if pd.notna(r["RFC"]) else
                          "Delivered" if pd.notna(r["DELIVERED"]) else
                          "Awaiting delivery", axis=1)
            for col in ("VESSEL ETA", "DELIVERED", "RFC", "MT RETURN"):
                detail[col] = pd.to_datetime(detail[col]).dt.strftime("%d %b %Y").fillna("—")
            for kind in ("DEM", "DET"):
                col = f"{kind} Delay Days"
                detail[col] = detail[col].map(
                    lambda v: "—" if pd.isna(v) else
                    f"({abs(int(v))} days early)" if v < 0 else
                    f"{int(v)} days late" if v > 0 else "0 days")
            def status_colors(v):
                if v in ("Early", "On time"):
                    return "background-color:#e3f6e9;color:#116b3a;font-weight:700"
                if v == "Late":
                    return "background-color:#fde8e8;color:#b42318;font-weight:700"
                return "color:#697c8f"
            st.dataframe(detail.style.map(status_colors, subset=["DEM Status", "DET Status"]),
                         hide_index=True, use_container_width=True)
    else:
        st.info("Select a vessel or BL shipment group to view total containers, delivery, RFC and return balances.")

if active_page == "overview":
    # Corporate management overview — compact, consistent, no donut or embedded scrollbars.
    import plotly.graph_objects as go

    ov_navy = "#123B61"
    ov_blue = "#276A9B"
    ov_teal = "#238B57"  # Delivered: green
    ov_gold = "#E5AC38"  # Pending: amber
    ov_muted = "#72869B"
    ov_grid = "#E9EFF5"
    ov_bg = "#FFFFFF"

    st.markdown("""
    <style>
    .ov-header{display:flex;align-items:flex-end;justify-content:space-between;
      border-bottom:1px solid #DCE7F0;padding:6px 0 17px;margin-bottom:19px}
    .ov-kicker{font-size:10px;letter-spacing:2px;font-weight:800;color:#397DA3;margin-bottom:7px}
    .ov-title{font-size:27px;line-height:1.18;font-weight:780;color:#123B61;letter-spacing:-.7px}
    .ov-subtitle{font-size:12px;color:#6D8399;margin-top:7px}
    .ov-summary{display:flex;gap:0;align-items:center;flex-wrap:wrap;background:#fff;
      border:1px solid #E0EAF2;border-radius:12px;padding:15px 8px;margin:0 0 19px;
      box-shadow:0 3px 12px rgba(18,59,97,.035)}
    .ov-stat{flex:1 1 130px;padding:3px 18px;border-right:1px solid #E6EEF4}
    .ov-stat:last-child{border-right:none}
    .ov-stat-label{font-size:11px;color:#6D8399;font-weight:600;margin-bottom:5px}
    .ov-stat-value{font-size:24px;color:#123B61;font-weight:800;line-height:1.2}
    .ov-stat-detail{font-size:10px;color:#8094A6;margin-top:4px}
    .ov-stat-delivered .ov-stat-value{color:#238B57}
    .ov-stat-pending .ov-stat-value{color:#B8780D}
    .ov-section{font-size:15px;font-weight:750;color:#173F63;margin:7px 0 1px}
    .ov-section-sub{font-size:11px;color:#7B8DA1;margin:0 0 8px}
    .ov-progress-wrap{background:#fff;border:1px solid #E0EAF2;border-radius:12px;
      padding:20px 22px 16px;margin-top:3px;box-shadow:0 3px 12px rgba(18,59,97,.035)}
    .ov-progress-title{font-size:12px;font-weight:750;color:#173F63}
    .ov-progress-percent{font-size:29px;font-weight:800;color:#123B61;margin-top:8px}
    .ov-progress-track{display:flex;width:100%;height:12px;background:#E9EFF5;
      border-radius:9px;overflow:hidden;margin:14px 0 12px}
    .ov-progress-fill{height:100%;background:#238B57}
    .ov-progress-rest{height:100%;background:#E5AC38}
    .ov-legend{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;
      color:#5D748B;font-size:11px;margin-top:8px}
    .ov-dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px}
    .ov-divider{border-top:1px solid #E6EDF3;margin:17px 0 12px}
    </style>
    <div class="ov-header"><div><div class="ov-kicker">FFM · LOGISTICS INTELLIGENCE</div>
    <div class="ov-title">Shipment Performance Overview</div>
    <div class="ov-subtitle">Delivery completion, shipment activity and operational distribution</div></div></div>
    """, unsafe_allow_html=True)

    total = int(f["CNTNR NO"].nunique())
    delivered_ids = int(f.loc[f["DELIVERY GROUP"] == "Delivered", "CNTNR NO"].nunique())
    pending_ids = max(0, total - delivered_ids)
    completion = 100.0 * delivered_ids / total if total else 0.0
    agents_count = int(f.loc[
        ~f["FORWARDING AGENT"].astype(str).str.strip().str.casefold().isin(
            ["", "not specified", "nan", "none", "<na>"]), "FORWARDING AGENT"].nunique())
    locations_count = int(f.loc[
        ~f["DELIVERY LOC"].astype(str).str.strip().str.casefold().isin(
            ["", "not specified", "nan", "none", "<na>"]), "DELIVERY LOC"].nunique())

    st.markdown(f"""<div class="ov-summary">
      <div class="ov-stat"><div class="ov-stat-label">Total containers</div><div class="ov-stat-value">{total:,}</div><div class="ov-stat-detail">Selected records</div></div>
      <div class="ov-stat ov-stat-delivered"><div class="ov-stat-label">Delivered</div><div class="ov-stat-value">{delivered_ids:,}</div><div class="ov-stat-detail">{completion:.1f}% completion</div></div>
      <div class="ov-stat ov-stat-pending"><div class="ov-stat-label">Pending delivery</div><div class="ov-stat-value">{pending_ids:,}</div><div class="ov-stat-detail">Outstanding containers</div></div>
      <div class="ov-stat"><div class="ov-stat-label">Forwarding agents</div><div class="ov-stat-value">{agents_count:,}</div><div class="ov-stat-detail">Identified agents</div></div>
      <div class="ov-stat"><div class="ov-stat-label">Delivery locations</div><div class="ov-stat-value">{locations_count:,}</div><div class="ov-stat-detail">Identified locations</div></div>
    </div>""", unsafe_allow_html=True)

    def ov_chart(fig, height=280):
        fig.update_layout(
            template="plotly_white", height=height,
            paper_bgcolor=ov_bg, plot_bgcolor=ov_bg,
            font=dict(family="Arial", size=11, color=ov_navy),
            margin=dict(l=8, r=16, t=12, b=14),
            bargap=.43,
            hoverlabel=dict(bgcolor="#FFFFFF", font_size=11),
        )
        fig.update_xaxes(showline=False, zeroline=False, showgrid=False,
                         tickfont=dict(size=10, color=ov_muted), automargin=True)
        fig.update_yaxes(showline=False, zeroline=False, gridcolor=ov_grid,
                         tickfont=dict(size=10, color=ov_muted), automargin=True)
        st.plotly_chart(fig, use_container_width=True,
                        config={"displayModeBar": False, "displaylogo": False, "responsive": True})

    left, right = st.columns([1.65, 1], gap="medium")
    with left:
        st.markdown('<div class="ov-section">Monthly Shipment Volume</div><div class="ov-section-sub">Containers by vessel ETA month · delivered versus outstanding</div>', unsafe_allow_html=True)
        trend = (f[f["ETA MONTH"] != "No ETA"]
                 .groupby(["ETA MONTH", "DELIVERY GROUP"])["CNTNR NO"]
                 .nunique().unstack(fill_value=0).sort_index())
        if trend.empty:
            st.caption("No valid vessel ETA dates for this selection.")
        else:
            months = pd.to_datetime(trend.index + "-01", errors="coerce").strftime("%b %Y")
            fig = go.Figure()
            for label, color in [("Delivered", ov_teal), ("Other / Pending", ov_gold)]:
                if label in trend:
                    fig.add_bar(name="Delivered" if label == "Delivered" else "Pending",
                                x=months, y=trend[label].values, marker_color=color,
                                hovertemplate="%{x}<br>%{y:,} containers<extra>" + ("Delivered" if label == "Delivered" else "Pending") + "</extra>")
            fig.update_layout(barmode="stack", showlegend=True,
                              legend=dict(orientation="h", y=1.16, x=1, xanchor="right",
                                          font=dict(size=10, color=ov_muted)),
                              margin=dict(l=5, r=8, t=32, b=12))
            fig.update_yaxes(title=None, rangemode="tozero")
            ov_chart(fig, 300)
    with right:
        st.markdown('<div class="ov-section">Delivery Completion</div><div class="ov-section-sub">Share of containers already delivered</div>', unsafe_allow_html=True)
        st.markdown(f"""<div class="ov-progress-wrap">
            <div class="ov-progress-title">Overall delivery progress</div>
            <div class="ov-progress-percent">{completion:.1f}%</div>
            <div class="ov-progress-track"><div class="ov-progress-fill" style="width:{completion:.4f}%"></div>
            <div class="ov-progress-rest" style="width:{100-completion:.4f}%"></div></div>
            <div class="ov-legend"><span><i class="ov-dot" style="background:#238B57"></i>Delivered · {delivered_ids:,}</span>
            <span><i class="ov-dot" style="background:#E5AC38"></i>Pending · {pending_ids:,}</span></div>
            <div class="ov-divider"></div>
            <div style="font-size:11px;color:#73899C;line-height:1.65">Delivery progress is based on unique container numbers in the selected filters.</div>
            </div>""", unsafe_allow_html=True)

    st.markdown('<div class="ov-divider"></div><div class="ov-section">Payment Status</div><div class="ov-section-sub">Payment status by container · amber: To Paid · blue: No · light grey: Not specified</div>', unsafe_allow_html=True)
    pay = (f.groupby("PAYMENT STATUS")["CNTNR NO"].nunique()
           .sort_values(ascending=True))
    if not pay.empty:
        colors = {"To Paid": "#E8A547", "No": "#347A9F", "Not specified": "#C3CFD9"}
        fig = go.Figure(go.Bar(
            x=pay.values, y=pay.index.tolist(), orientation="h",
            marker=dict(color=[colors.get(str(p), "#C3CFD9") for p in pay.index], line=dict(width=0)),
            text=[f"{n:,}" for n in pay.values], textposition="outside",
            cliponaxis=False,
            hovertemplate="%{y}: %{x:,} containers<extra></extra>"))
        fig.update_layout(margin=dict(l=8, r=45, t=8, b=8), bargap=0.42)
        fig.update_xaxes(visible=False, range=[0, max(pay.max()*1.2, 1)])
        fig.update_yaxes(showgrid=False)
        ov_chart(fig, max(125, 42*len(pay)+35))

    st.markdown('<div class="ov-divider"></div><div class="ov-section">Operational Distribution</div><div class="ov-section-sub">Top locations and forwarding agents by container volume</div>', unsafe_allow_html=True)
    dist_left, dist_right = st.columns(2, gap="large")
    for panel, field, label in [
        (dist_left, "DELIVERY LOC", "Delivery Location"),
        (dist_right, "FORWARDING AGENT", "Forwarding Agent")
    ]:
        with panel:
            st.markdown(f'<div class="ov-section" style="font-size:13px">{label}</div>', unsafe_allow_html=True)
            names = f[field].astype("string").str.strip()
            valid = f[names.notna() & ~names.str.casefold().isin(
                ["", "not specified", "nan", "none", "<na>"])]
            ranked = (valid.groupby(field)["CNTNR NO"].nunique()
                      .nlargest(7).sort_values())
            if ranked.empty:
                st.caption("No identified records in this selection.")
            else:
                fig = go.Figure(go.Bar(
                    x=ranked.values, y=ranked.index.tolist(), orientation="h",
                    marker=dict(color=ov_blue, line=dict(width=0)),
                    text=[f"{v:,}" for v in ranked.values], textposition="outside",
                    cliponaxis=False,
                    hovertemplate="%{y}<br>%{x:,} containers<extra></extra>"))
                fig.update_xaxes(visible=False, range=[0, max(ranked.max()*1.22, 1)])
                fig.update_yaxes(showgrid=False, tickfont=dict(size=10))
                ov_chart(fig, 290)

if active_page == "risks":
    st.subheader("Forwarding Agent Performance")

    agent_data = f.copy()
    agent_name = agent_data["FORWARDING AGENT"].astype("string").str.strip()
    agent_data = agent_data[agent_name.notna() & ~agent_name.str.casefold().isin(
        ["", "not specified", "nan", "none", "<na>"]
    )].copy()
    if agent_data.empty:
        st.info("No forwarding-agent records available for the current filters.")
    else:
        # Count unique container numbers per agent; DEM and DET categories may overlap.
        agent_data["HAS DEM"] = pd.to_numeric(
            agent_data["DEM TO PAY / DAYS POSITIVE"], errors="coerce"
        ).gt(0)
        agent_data["HAS DET"] = pd.to_numeric(
            agent_data["DET TO PAY / DAYS POSITIVE"], errors="coerce"
        ).gt(0)
        per_container = (
            agent_data.groupby(["FORWARDING AGENT", "CNTNR NO"], as_index=False)
            .agg(**{
                "HAS DEM": ("HAS DEM", "max"),
                "HAS DET": ("HAS DET", "max"),
            })
        )
        agent_summary = (
            per_container.groupby("FORWARDING AGENT", as_index=False)
            .agg(**{
                "Total Containers": ("CNTNR NO", "nunique"),
                "Containers with DEM": ("HAS DEM", "sum"),
                "Containers with DET": ("HAS DET", "sum"),
            })
        )
        agent_summary["DEM Rate (%)"] = (
            agent_summary["Containers with DEM"] / agent_summary["Total Containers"] * 100
        ).round(1)
        agent_summary["DET Rate (%)"] = (
            agent_summary["Containers with DET"] / agent_summary["Total Containers"] * 100
        ).round(1)
        agent_summary = agent_summary.sort_values(
            "Total Containers", ascending=False
        ).reset_index(drop=True)

        st.markdown("#### DEM / DET vs Total Containers by Forwarding Agent")
        graph = agent_summary.melt(
            id_vars=["FORWARDING AGENT"],
            value_vars=["Total Containers", "Containers with DEM", "Containers with DET"],
            var_name="Metric", value_name="Containers",
        )
        fig = px.bar(
            graph, x="Containers", y="FORWARDING AGENT", color="Metric",
            orientation="h", barmode="group", text="Containers",
            color_discrete_map={
                "Total Containers": "#7796B5",
                "Containers with DEM": "#155CAA",
                "Containers with DET": "#F0A12B",
            },
            category_orders={
                "FORWARDING AGENT": agent_summary["FORWARDING AGENT"].iloc[::-1].tolist(),
                "Metric": ["Total Containers", "Containers with DEM", "Containers with DET"],
            },
        )
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_layout(
            template="plotly_white",
            height=max(340, len(agent_summary) * 88 + 120),
            xaxis_title="Number of containers",
            yaxis_title=None,
            legend_title_text="",
            legend=dict(
                orientation="h",
                x=0, xanchor="left",
                y=1.02, yanchor="bottom",
                font=dict(size=12, color="#294765"),
                bgcolor="rgba(255,255,255,0)",
                traceorder="normal",
            ),
            margin=dict(l=10, r=65, t=80, b=30),
        )
        st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

        st.dataframe(
            agent_summary,
            hide_index=True,
            use_container_width=True,
            column_config={
                "DEM Rate (%)": st.column_config.NumberColumn(format="%.1f%%"),
                "DET Rate (%)": st.column_config.NumberColumn(format="%.1f%%"),
            },
        )
    st.markdown("#### Highest recorded chargeable-day exposure")
    for label, group_col in [("Suppliers", "SUPPLIER"), ("Delivery locations", "DELIVERY LOC")]:
        st.markdown(f"**{label}**")
        exposure = category_summary(f, group_col)
        st.dataframe(exposure[[group_col, "Containers", "Chargeable_containers", "Recorded_chargeable_days", "Chargeable rate (%)"]].head(10),
                     hide_index=True, use_container_width=True)

    st.subheader("Records with positive DEM / DET")
    risk_rows = f[(f["DEM TO PAY / DAYS POSITIVE"] > 0) | (f["DET TO PAY / DAYS POSITIVE"] > 0)]
    cols = ["CNTNR NO", "SUPPLIER", "DELIVERY LOC", "FORWARDING AGENT", "STATUS", "PAYMENT STATUS",
            "DEM TO PAY / DAYS POSITIVE", "DET TO PAY / DAYS POSITIVE", "VESSEL ETA"]
    st.dataframe(risk_rows[cols], hide_index=True, use_container_width=True, height=340)

if active_page == "leadtime":
    st.subheader("Lead-Time Analysis")

    intervals = [
        ("ETA → Delivered", "VESSEL ETA", "DELIVERED"),
        ("Delivered → RFC", "DELIVERED", "RFC"),
        ("RFC → Empty Return", "RFC", "MT RETURN"),
    ]
    rows = []
    for label, start, end in intervals:
        duration = stage_days(f, start, end)
        rows.append({"Process stage": label, "Average days": duration.mean(),
                     "Maximum days": duration.max(), "Completed records": int(duration.notna().sum())})
    summary = pd.DataFrame(rows)
    st.dataframe(summary, hide_index=True, use_container_width=True,
                 column_config={"Average days": st.column_config.NumberColumn(format="%.1f"),
                                "Maximum days": st.column_config.NumberColumn(format="%.1f")})
    valid = summary.dropna(subset=["Average days"])
    if not valid.empty:
        fig = px.bar(valid, x="Process stage", y="Average days", text_auto=".1f",
                     color_discrete_sequence=[BLUE])
        plot(fig, 320)
    st.markdown("#### Process lead time by container")
    detail = f[["CNTNR NO", "SUPPLIER", "DELIVERY LOC", "FORWARDING AGENT"]].copy()
    for label, start, end in intervals:
        detail[label + " (days)"] = stage_days(f, start, end)
    st.dataframe(detail, hide_index=True, use_container_width=True, height=350)

if active_page == "comparison":
    st.subheader("Performance Comparison")

    dim = st.selectbox("Compare by", ["Supplier", "Forwarding agent", "Delivery location"], key="compare_dimension")
    field = {"Supplier": "SUPPLIER", "Forwarding agent": "FORWARDING AGENT", "Delivery location": "DELIVERY LOC"}[dim]
    summary = category_summary(f, field)
    if not summary.empty:
        top = summary.head(15).sort_values("Recorded_chargeable_days")
        fig = px.bar(top, x="Recorded_chargeable_days", y=field, orientation="h",
                     text="Recorded_chargeable_days", color_discrete_sequence=[BLUE])
        fig.update_layout(xaxis_title="Recorded chargeable days", yaxis_title=None)
        plot(fig, 400)
        st.dataframe(summary.rename(columns={
            "Chargeable_containers": "Containers with recorded chargeable days",
            "Recorded_chargeable_days": "Total recorded chargeable days",
            "Avg_ETA_to_Delivery": "Avg ETA → Delivered (days)",
            "Avg_Delivery_to_RFC": "Avg Delivered → RFC (days)",
            "Avg_RFC_to_Return": "Avg RFC → Return (days)"}),
            hide_index=True, use_container_width=True, height=340)
    st.markdown("#### Monthly trend — May to July")

    months_df = f[f["VESSEL ETA"].notna()].copy()
    months_df = months_df[months_df["VESSEL ETA"].dt.month.isin([5, 6, 7])]
    if months_df.empty:
        st.info("No May–July records with a valid Vessel ETA under the current filters.")
    else:
        years = sorted(months_df["VESSEL ETA"].dt.year.unique().tolist(), reverse=True)
        selected_year = st.selectbox("Year", years, key="comparison_year")
        months_df = months_df[months_df["VESSEL ETA"].dt.year == selected_year]
        months_df["Month"] = months_df["VESSEL ETA"].dt.strftime("%b")
        months_df["Chargeable days"] = months_df["DEM TO PAY / DAYS POSITIVE"] + months_df["DET TO PAY / DAYS POSITIVE"]
        monthly = months_df.groupby("Month", as_index=False).agg(Containers=("CNTNR NO", "size"),
                                                                Recorded_chargeable_days=("Chargeable days", "sum"))
        monthly["Month"] = pd.Categorical(monthly["Month"], categories=["May", "Jun", "Jul"], ordered=True)
        monthly = monthly.sort_values("Month")
        fig = px.line(monthly, x="Month", y="Recorded_chargeable_days", markers=True,
                      color_discrete_sequence=[BLUE])
        fig.update_layout(yaxis_title="Recorded chargeable days")
        plot(fig, 320)
        st.dataframe(monthly, hide_index=True, use_container_width=True)

if active_page == "rootcause":
    st.subheader("Root-Cause Analysis from Remarks")

    import re
    rules = {
        "Customs / OGA delay": [r"customs?", r"\boga\b", r"kastam", r"permit", r"inspection", r"approval"],
        "Payment delay": [r"payment", r"credit limit", r"credit hold", r"outstanding", r"unpaid", r"invoice", r"bank release"],
        "Warehouse capacity": [r"warehouse", r"capacity", r"space", r"congestion", r"unload", r"unloading", r"storage", r"full warehouse"],
        "Documentation / HS code": [r"document", r"documentation", r"\bhs[ -]?code\b", r"incorrect code", r"bl amendment", r"bill of lading", r"manifest", r"declaration"],
    }
    tagged = []
    for _, rec in f.iterrows():
        remark = str(rec.get("REMARKS", "") or "")
        labels = [label for label, patterns in rules.items()
                  if any(re.search(pattern, remark, flags=re.IGNORECASE) for pattern in patterns)]
        if not labels:
            labels = ["Unclassified"]
        for label in labels:
            tagged.append({"CNTNR NO": rec["CNTNR NO"], "Category": label,
                           "Remarks": remark, "Supplier": rec["SUPPLIER"],
                           "Location": rec["DELIVERY LOC"], "Agent": rec["FORWARDING AGENT"]})
    causes = pd.DataFrame(tagged)
    if not causes.empty:
        counts = causes["Category"].value_counts().rename_axis("Potential cause").reset_index(name="Records")
        fig = px.bar(counts, x="Records", y="Potential cause", orientation="h",
                     color_discrete_sequence=[BLUE], text="Records")
        fig.update_layout(yaxis_title=None)
        plot(fig, 350)
        selected_cause = st.selectbox("Inspect potential cause", counts["Potential cause"].tolist(), key="root_cause_filter")
        st.dataframe(causes[causes["Category"] == selected_cause], hide_index=True,
                     use_container_width=True, height=340)

if active_page == "records":
    st.subheader("Shipment explorer")

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
