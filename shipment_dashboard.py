import io
import base64
import html
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

st.markdown('\n<style>\n/* Wider, readable, compact filters sidebar */\n[data-testid="stSidebar"] {\n  width: 345px !important;\n  min-width: 345px !important;\n}\n[data-testid="stSidebar"] > div:first-child {\n  width: 345px !important;\n}\n[data-testid="stSidebar"] [data-testid="stSidebarContent"] {\n  padding: 12px 17px 20px !important;\n  overflow-x: hidden !important;\n}\n[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {\n  gap: .48rem !important;\n}\n[data-testid="stSidebar"] .filter-intro {\n  padding: 17px 17px 15px !important;\n  margin: 0 0 8px !important;\n  border-radius: 11px !important;\n}\n[data-testid="stSidebar"] .filter-heading {\n  font-size: 21px !important;\n}\n[data-testid="stSidebar"] .filter-description {\n  font-size: 11px !important;\n  margin-top: 5px !important;\n}\n[data-testid="stSidebar"] .filter-section {\n  font-size: 10px !important;\n  letter-spacing: .8px !important;\n  margin: 7px 0 5px !important;\n  padding: 7px 0 !important;\n}\n[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {\n  font-size: 12px !important;\n  white-space: normal !important;\n  overflow-wrap: anywhere !important;\n  line-height: 1.35 !important;\n}\n[data-testid="stSidebar"] [data-baseweb="select"] > div,\n[data-testid="stSidebar"] [data-testid="stTextInput"] input {\n  min-height: 38px !important;\n}\n[data-testid="stSidebar"] [data-testid="stWidgetLabel"] {\n  margin-bottom: 3px !important;\n}\n@media(max-width: 760px) {\n  [data-testid="stSidebar"],[data-testid="stSidebar"] > div:first-child {\n    width: min(345px, 90vw) !important;\n    min-width: 0 !important;\n  }\n}\n</style>\n', unsafe_allow_html=True)

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
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        df[col + " POSITIVE"] = df[col].clip(lower=0)
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

with st.sidebar:
    st.markdown("""<div class="filter-intro"><div class="filter-eyebrow">FFM LOGISTICS · CONTROL TOWER</div>
    <div class="filter-heading">Filters & Search</div>
    </div>""", unsafe_allow_html=True)
    st.markdown('<div class="filter-section">SEARCH CONTAINER</div>', unsafe_allow_html=True)
    search = st.text_input("Container number", placeholder="🔎  Enter container number", key="container", label_visibility="collapsed")
    st.markdown('<div class="filter-section">SHIPMENT DETAILS</div>', unsafe_allow_html=True)
    def filter_multi(label, column, key):
        options = sorted(df[column].dropna().astype(str).unique().tolist())
        return st.multiselect(label, options, default=[], key=key, placeholder="All")
    months = filter_multi("ETA month", "ETA MONTH", "month")
    locations = filter_multi("Delivery location", "DELIVERY LOC", "location")
    agents = filter_multi("Forwarding agent", "FORWARDING AGENT", "agent")
    statuses = filter_multi("Delivery status", "STATUS", "status")
    payments = filter_multi("Payment status", "PAYMENT STATUS", "payment_status")
    suppliers = filter_multi("Supplier", "SUPPLIER", "supplier")
    st.markdown('<div class="filter-section">DEM / DET</div>', unsafe_allow_html=True)
    only_positive = st.checkbox("Show positive DEM / DET days only", key="positive_only")


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


monitor, overview, risks, leadtime, comparison, rootcause, records = st.tabs(["📦 Container Monitoring", "📊 Overview", "⚠️ DEM / DET analysis", "⏱️ Lead-Time Analysis", "📈 Performance Comparison", "🔍 Root-Cause Analysis", "🔎 Shipment explorer"])

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

with monitor:
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
            rows.append({"Container":str(container), **dates})
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
        a,b,c,d = st.columns(4)
        a.metric("Total containers",total_containers)
        b.metric("Delivered",f"{delivered_count} / {total_containers}")
        c.metric("RFC · Empty",f"{rfc_count} / {total_containers}")
        d.metric("Empty returned",f"{returned_count} / {total_containers}")
        a,b,c = st.columns(3)
        a.metric("Balance to deliver",max(0,total_containers-delivered_count))
        b.metric("Balance to empty (RFC)",max(0,total_containers-rfc_count))
        c.metric("Balance to return",max(0,total_containers-returned_count))
        st.markdown("#### Vessel Process Progress")
        stages = [("ETA", "VESSEL ETA"), ("Delivered", "DELIVERED"),
                  ("RFC · Empty", "RFC"), ("Empty Return", "MT RETURN")]
        # One aligned journey with four milestones and three connecting segments.
        progress_cells=[]
        for idx,(label,col) in enumerate(stages):
            completed=counts[col]
            pct=100*completed/total_containers if total_containers else 0
            first_date=progress[col].min()
            date_text=first_date.strftime("%d %b %Y") if pd.notna(first_date) else "—"
            progress_cells.append(f"""<div class="vessel-phase">
                <div class="vessel-dot"></div><div class="vessel-label">{html.escape(label)}</div>
                <div class="vessel-count">{completed} / {total_containers}</div>
                <div class="vessel-small">{pct:.0f}% completed</div>
                <div class="vessel-small">First: {date_text}</div>
                </div>""")
        st.markdown(f"""<style>
          .vessel-journey{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));
            margin:18px 0 22px;padding:10px 0 0;background:white;border:1px solid #dce7f2;
            border-radius:14px;box-shadow:0 5px 20px #173b5910}}
          .vessel-phase{{position:relative;text-align:center;padding:40px 5px 20px}}
          .vessel-phase:not(:last-child)::after{{content:"";position:absolute;top:18px;
            left:50%;width:100%;height:2px;background:#a5c0d9;z-index:0}}
          .vessel-dot{{position:absolute;top:10px;left:50%;transform:translateX(-50%);
            height:17px;width:17px;background:#155caa;border:3px solid white;
            border-radius:50%;box-shadow:0 0 0 2px #155caa;z-index:2}}
          .vessel-label{{font-size:12px;color:#4f6880}}
          .vessel-count{{font-size:20px;font-weight:800;color:#12395d;margin:6px 0}}
          .vessel-small{{font-size:11px;color:#668198;margin-top:4px}}
          </style><div class="vessel-journey">{''.join(progress_cells)}</div>""",unsafe_allow_html=True)
        st.markdown("#### Daily Delivery & Emptying Progress")
        # Milestone dates count the number of distinct containers that completed
        # the stage on that calendar day; running cumulative totals show balances.
        dates=[]
        for col in stage_cols[1:]:
            dates.extend(progress[col].dropna().dt.normalize().tolist())
        if dates:
            start_day=min(dates)
            end_day=max(dates)
            calendar=pd.DataFrame({"Date":pd.date_range(start_day,end_day,freq="D")})
            for col,name in [("DELIVERED","Delivered"),("RFC","RFC / Empty"),("MT RETURN","Empty returned")]:
                daily=progress[col].dropna().dt.normalize().value_counts()
                calendar[name+" today"]=calendar["Date"].map(daily).fillna(0).astype(int)
                calendar[name+" cumulative"]=calendar[name+" today"].cumsum()
                calendar[name+" balance"]=(total_containers-calendar[name+" cumulative"]).clip(lower=0)
            calendar["Date"]=calendar["Date"].dt.strftime("%d %b %Y")
            st.dataframe(calendar,hide_index=True,use_container_width=True,height=320)
        else:
            st.info("No completed delivery, RFC or return dates recorded for this group.")
        st.markdown("#### Container Status by Vessel")
        status=progress.rename(columns={"VESSEL ETA":"ETA", "MT RETURN":"Empty Return"}).copy()
        for col in ("ETA","DELIVERED","RFC","Empty Return"):
            status[col]=pd.to_datetime(status[col]).dt.strftime("%d %b %Y").fillna("—")
        status["Current Stage"]=progress.apply(
            lambda r: "Empty returned" if pd.notna(r["MT RETURN"]) else
                      "RFC · Empty" if pd.notna(r["RFC"]) else
                      "Delivered" if pd.notna(r["DELIVERED"]) else
                      "Awaiting delivery",axis=1)
        st.dataframe(status,hide_index=True,use_container_width=True)
    else:
        st.info("Select a vessel or BL shipment group to view total containers, delivery, RFC and return balances.")

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
            margin=dict(l=10, r=65, t=25, b=30),
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

with leadtime:
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

with comparison:
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

with rootcause:
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

with records:
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

