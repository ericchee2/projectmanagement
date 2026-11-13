import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime, date, timedelta
from io import BytesIO
from html import escape
import openpyxl
import re

st.set_page_config(page_title="Emart | Store Opening Control Centre", page_icon="🏬", layout="wide")

# ───────────────────────── Paths ─────────────────────────
ROOT = Path(__file__).parent
DATA_DIR = ROOT / "data"
PHOTO_DIR = DATA_DIR / "photos"
DATA_DIR.mkdir(exist_ok=True)
PHOTO_DIR.mkdir(exist_ok=True)
MASTER = DATA_DIR / "store_opening_master.xlsx"
_default = ROOT / "Test_Dashboard(2).xlsx"
if not _default.exists():
    _c = sorted(ROOT.glob("Test_Dashboard*.xlsx"))
    _default = _c[0] if _c else _default
DEFAULT_SOURCE = _default

# ───────────────────────── Design tokens ─────────────────────────
NAVY = "#0B2A5B"
BLUE = "#1D6FF2"
SKY = "#4DA3FF"
CYAN = "#22C7F0"
PALE = "#EAF3FF"
GREEN = "#12B76A"
AMBER = "#F79009"
RED = "#F04438"
SLATE = "#98A2B3"
STATUS_COLORS = {"Completed": GREEN, "In Progress": BLUE, "Not Started": SLATE, "On Hold": AMBER}
STATUS_LIST = ["Not Started", "In Progress", "Completed", "On Hold"]

CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], .stApp {{font-family:'Inter',-apple-system,'Segoe UI','Microsoft YaHei',sans-serif;}}
.stApp {{background:#F3F7FE;}}
.block-container {{padding-top:1.2rem;padding-bottom:2.5rem;max-width:1560px;}}
header[data-testid="stHeader"] {{background:transparent;}}
#MainMenu, footer {{visibility:hidden;}}
h1,h2,h3,h4 {{color:{NAVY};font-weight:700;letter-spacing:-.2px;}}

/* Sidebar */
[data-testid="stSidebar"] {{background:#FFFFFF;border-right:1px solid #E3ECFB;}}
[data-testid="stSidebar"] > div:first-child {{padding-top:.4rem;}}
.brand {{display:flex;align-items:center;gap:12px;padding:6px 6px 14px 6px;}}
.brand-logo {{width:42px;height:42px;border-radius:12px;background:linear-gradient(135deg,{BLUE},{CYAN});display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;font-size:20px;box-shadow:0 6px 16px rgba(29,111,242,.35);}}
.brand-t {{font-weight:800;color:{NAVY};font-size:17px;line-height:1.1;}}
.brand-s {{font-size:10.5px;color:#6B83A8;letter-spacing:1.2px;font-weight:600;margin-top:3px;}}
.nav-cap {{font-size:11px;font-weight:700;color:#8AA0C2;letter-spacing:1.4px;margin:10px 6px 6px 6px;}}
[data-testid="stSidebar"] [data-testid="stRadioGroup"] {{gap:4px;}}
[data-testid="stSidebar"] [data-testid="stRadioOption"] {{padding:10px 14px;border-radius:12px;width:100%;cursor:pointer;transition:all .15s;margin:0;}}
[data-testid="stSidebar"] [data-testid="stRadioOption"]:hover {{background:{PALE};}}
[data-testid="stSidebar"] [data-testid="stRadioOption"] > div > div:first-child {{display:none;}}
[data-testid="stSidebar"] [data-testid="stRadioOption"] p {{font-size:14px;font-weight:600;color:{NAVY};margin:0;}}
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] {{background:linear-gradient(90deg,{BLUE},#3B8CFF);box-shadow:0 6px 16px rgba(29,111,242,.30);}}
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] p {{color:#fff !important;}}
[class*="st-key-card"] {{background:#fff;border:1px solid #E3ECFB !important;border-radius:18px;padding:16px 18px 6px 18px;box-shadow:0 4px 18px rgba(29,111,242,.07);margin-bottom:16px;}}

/* Hero */
.hero {{background:linear-gradient(120deg,#0A3AA8 0%,{BLUE} 55%,#38BDF8 100%);border-radius:20px;padding:24px 30px;color:#fff;display:flex;justify-content:space-between;align-items:center;gap:24px;box-shadow:0 12px 32px rgba(29,111,242,.28);margin-bottom:18px;position:relative;overflow:hidden;}}
.hero:after {{content:"";position:absolute;right:-60px;top:-80px;width:280px;height:280px;border-radius:50%;background:rgba(255,255,255,.10);}}
.hero:before {{content:"";position:absolute;right:140px;bottom:-110px;width:220px;height:220px;border-radius:50%;background:rgba(255,255,255,.07);}}
.hero h1 {{color:#fff !important;font-size:26px;margin:0;font-weight:800;}}
.hero p {{color:rgba(255,255,255,.86);margin:5px 0 0 0;font-size:13.5px;}}
.hero-stats {{display:flex;gap:14px;z-index:1;}}
.hero-stat {{background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.28);backdrop-filter:blur(6px);border-radius:14px;padding:10px 20px;text-align:center;min-width:112px;}}
.hero-stat b {{display:block;font-size:26px;font-weight:800;line-height:1.15;}}
.hero-stat span {{font-size:11px;letter-spacing:.8px;opacity:.9;font-weight:600;}}

/* Cards */
.card {{background:#fff;border:1px solid #E3ECFB;border-radius:18px;padding:20px 22px;box-shadow:0 4px 18px rgba(29,111,242,.07);margin-bottom:16px;}}
.card-h {{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;}}
.card-t {{font-size:15.5px;font-weight:700;color:{NAVY};}}
.card-s {{font-size:12px;color:#6B83A8;}}
.kpi {{background:#fff;border:1px solid #E3ECFB;border-radius:18px;padding:16px 18px;box-shadow:0 4px 18px rgba(29,111,242,.07);display:flex;gap:14px;align-items:center;height:100%;}}
.kpi-ic {{width:46px;height:46px;border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:21px;flex-shrink:0;}}
.kpi-l {{font-size:11px;color:#6B83A8;font-weight:700;letter-spacing:.7px;}}
.kpi-v {{font-size:28px;color:{NAVY};font-weight:800;line-height:1.15;}}
.kpi-v small {{font-size:15px;font-weight:700;}}
.kpi-s {{font-size:11.5px;color:#6B83A8;margin-top:1px;}}

/* Progress bars & rings */
.pb {{margin:11px 0;}}
.pb-h {{display:flex;justify-content:space-between;font-size:13px;color:{NAVY};font-weight:600;margin-bottom:5px;}}
.pb-h span em {{font-style:normal;color:#8AA0C2;font-weight:500;margin-left:6px;font-size:11.5px;}}
.pb-t {{height:9px;background:#E6EEFB;border-radius:99px;overflow:hidden;}}
.pb-f {{height:100%;border-radius:99px;}}
.ring-card {{background:{PALE};border-radius:16px;padding:14px 16px;display:flex;align-items:center;gap:14px;}}
.ring {{width:68px;height:68px;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-shrink:0;}}
.ring-in {{width:50px;height:50px;border-radius:50%;background:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:16px;color:{NAVY};}}
.ring-in small {{font-size:10px;margin-left:1px;}}
.ring-t b {{display:block;font-size:13.5px;color:{NAVY};line-height:1.25;}}
.ring-t span {{font-size:11.5px;color:#6B83A8;}}

/* Badges, list rows */
.badge {{display:inline-block;padding:3px 10px;border-radius:99px;font-size:11px;font-weight:700;}}
.b-green {{background:#D1FADF;color:#027A48;}} .b-blue {{background:#DCEBFF;color:#1457C7;}}
.b-red {{background:#FEE4E2;color:#B42318;}} .b-amber {{background:#FEF0C7;color:#B54708;}} .b-grey {{background:#EEF2F7;color:#475467;}}
.row {{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid #EEF3FB;}}
.row:last-child {{border-bottom:none;}}
.row-t {{font-size:13.5px;font-weight:600;color:{NAVY};}}
.row-s {{font-size:11.5px;color:#6B83A8;margin-top:1px;}}
.date-chip {{background:{PALE};color:{BLUE};border-radius:10px;padding:5px 10px;font-size:12px;font-weight:700;text-align:center;min-width:58px;}}
.visit {{background:#fff;border:1px solid #E3ECFB;border-left:5px solid {BLUE};border-radius:14px;padding:14px 18px;margin-bottom:12px;box-shadow:0 3px 12px rgba(29,111,242,.05);}}
.visit h4 {{margin:0 0 4px 0;font-size:14.5px;}} .visit p {{margin:3px 0;font-size:13px;color:#344767;}}
.matrix {{display:grid;grid-template-columns:70px repeat(3,1fr);gap:6px;font-size:12px;font-weight:700;color:{NAVY};text-align:center;}}
.mx {{border-radius:12px;padding:16px 0;font-size:20px;font-weight:800;color:#fff;}}
.note {{font-size:12px;color:#6B83A8;}}

/* Streamlit widget polish */
div.stButton > button, div.stDownloadButton > button, div[data-testid="stFormSubmitButton"] > button {{border-radius:10px;font-weight:600;border:1px solid #CFE0FB;}}
div.stButton > button[kind="primary"], div[data-testid="stFormSubmitButton"] > button[kind="primary"] {{background:linear-gradient(90deg,{BLUE},#3B8CFF);border:none;color:#fff;box-shadow:0 6px 14px rgba(29,111,242,.28);}}
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {{border:1px solid #E3ECFB;border-radius:14px;overflow:hidden;background:#fff;}}
.stTabs [data-baseweb="tab-list"] {{gap:6px;background:#E6EEFB;padding:4px;border-radius:12px;width:fit-content;}}
.stTabs [data-baseweb="tab"] {{border-radius:9px;padding:6px 18px;font-weight:600;height:auto;}}
.stTabs [aria-selected="true"] {{background:#fff;color:{BLUE};box-shadow:0 2px 8px rgba(29,111,242,.15);}}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{display:none;}}
div[data-testid="stForm"] {{background:#fff;border:1px solid #E3ECFB;border-radius:18px;padding:20px;box-shadow:0 4px 18px rgba(29,111,242,.06);}}
div[data-testid="stExpander"] {{background:#fff;border-radius:14px;border:1px solid #E3ECFB;}}
/* v2.1 additions */
.kpi2 {{padding:0 0 14px 0;}}
.kpi2-row {{display:flex;justify-content:space-between;align-items:center;margin:4px 0 2px 0;}}
.kpi2 .kpi-ic {{width:34px;height:34px;border-radius:10px;font-size:17px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}}
.kpi2 .kpi-l {{white-space:nowrap;font-size:10.5px;letter-spacing:.4px;}}
.kpi2 .kpi-v {{line-height:1.1;}}
.kpi2 .kpi-s {{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:11px;}}
[class*="st-key-kpi_"], [class*="st-key-dept_"] {{background:#fff;border:1px solid #E3ECFB;border-radius:18px;padding:14px 16px 10px 16px;box-shadow:0 4px 18px rgba(29,111,242,.07);gap:.35rem;}}
[class*="st-key-kpi_"] button, [class*="st-key-dept_"] button {{background:{PALE};color:{BLUE};border:none;font-size:12px;min-height:28px;padding:0 8px;white-space:nowrap;}}
[class*="st-key-kpi_"] button:hover, [class*="st-key-dept_"] button:hover {{background:{BLUE};color:#fff;}}
.ring-card {{min-width:0;}} .ring-t {{min-width:0;}}
.ring-t b, .ring-t span {{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.dot {{display:inline-block;width:9px;height:9px;border-radius:50%;margin:0 5px 0 0;}}
.cat {{padding:10px 0;border-bottom:1px solid #EEF3FB;}} .cat:last-child {{border-bottom:none;}}
.cat-h {{display:flex;justify-content:space-between;align-items:baseline;}}
.cat-h b {{font-size:13.5px;color:{NAVY};}} .cat-p {{font-size:15px;font-weight:800;}}
.seg {{display:flex;height:10px;border-radius:99px;overflow:hidden;background:#E6EEFB;margin:7px 0 6px 0;}}
.seg i {{display:block;height:100%;}}
.cat-s {{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:6px;font-size:11.5px;color:#6B83A8;}}
.strip {{display:grid;grid-template-columns:repeat(5,1fr) 1.6fr;gap:12px;margin:6px 0 14px 0;}}
.chip {{background:#fff;border:1px solid #E3ECFB;border-radius:14px;padding:10px 14px;box-shadow:0 3px 12px rgba(29,111,242,.05);}}
.chip b {{display:block;font-size:19px;color:{NAVY};font-weight:800;white-space:nowrap;}} .chip span {{font-size:10.5px;color:#6B83A8;font-weight:700;letter-spacing:.6px;}}
[data-testid="stButtonGroup"] button {{border-radius:10px;font-weight:600;}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# ───────────────────────── Data layer ─────────────────────────
VISIT_COLS = ["Visit ID", "Visit Date", "Task ID", "Task Name", "Status", "Observation", "Follow-up Action", "PIC", "Recorded By"]
PIC_COLS = ["PIC Name", "Department", "Email", "Contact", "Notes"]
PHOTO_COLS = ["Photo ID", "Date", "Task ID", "Task Name", "Caption", "File Name", "Uploaded By"]
ISSUE_COLS = ["Issue ID", "Date Identified", "Title", "Related Task ID", "Category", "Likelihood", "Impact", "Priority", "Owner", "Status", "Due Date", "Mitigation"]
ISSUE_CATS = ["Schedule", "Budget", "Quality", "Safety", "Operations", "Compliance", "Supply Chain", "IT / Systems", "Manpower"]
LEVELS = ["Low", "Medium", "High"]
ISSUE_STATUS = ["Open", "Analyzing", "Planning", "In Progress", "Resolved"]


CAT_EN = {
    "选址与项目管理": "Site Selection & Project Management", "店面布局与动线规划": "Store Layout & Flow Planning",
    "政府执照、注册与行政手续": "Licenses, Registration & Admin", "采购与物资准备": "Procurement & Materials",
    "装修、设施与施工": "Fit-out, Facilities & Construction", "系统安装与测试": "Systems Installation & Testing",
    "人力资源与员工准备": "HR & Staff Readiness", "商品、供应商与库存准备": "Merchandise, Suppliers & Inventory",
    "租户管理（如适用）": "Tenant Management", "培训与营运测试": "Training & Operations Trial",
    "开业前最终准备": "Pre-Opening Final Preparation", "开业当天与开业后": "Opening Day & Post-Opening",
}


def clean_date(v):
    if v is None or (not isinstance(v, (datetime, date)) and (pd.isna(v) or v in ("", "NA", "N/A"))):
        return None
    try:
        return pd.to_datetime(v).date()
    except Exception:
        return None


def parse_progress(v):
    if v is None or v == "":
        return 0
    if isinstance(v, (int, float)):
        return round(float(v) * 100 if float(v) <= 1 else float(v), 1)
    s = str(v).replace("%", "").strip()
    try:
        return float(s)
    except Exception:
        return 0


def read_workbook(file_obj):
    wb = openpyxl.load_workbook(file_obj, data_only=False)
    if "Checklist" not in wb.sheetnames:
        raise ValueError("Excel file must contain a sheet named 'Checklist'.")
    ws = wb["Checklist"]
    tasks, category = [], "Uncategorised"
    for row in range(2, ws.max_row + 1):
        raw_id = ws.cell(row, 1).value
        name = ws.cell(row, 2).value
        if isinstance(raw_id, str) and raw_id.strip() and not name:
            category = CAT_EN.get(raw_id.strip(), raw_id.strip())
            continue
        if isinstance(raw_id, (int, float)) and name:
            v = [ws.cell(row, c).value for c in range(1, 14)]
            tasks.append({
                "ID": int(raw_id), "Category": category,
                "Task Name": str(v[1]).strip(), "Local Name": v[2] or "",
                "Resource Names": v[3] or "", "Duration": v[4] or "",
                "Start": clean_date(v[5]), "Finish": clean_date(v[6]),
                "Predecessors": v[7] if v[7] is not None else "", "Dependency": v[8] or "",
                "Actual Start": clean_date(v[9]), "Actual Finish": clean_date(v[10]),
                "Status": v[11] or "Not Started", "Progress": parse_progress(v[12]),
            })
    return wb, pd.DataFrame(tasks)


def load_sheet(name, columns, path=None):
    path = path or (MASTER if MASTER.exists() else DEFAULT_SOURCE)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        if name not in wb.sheetnames:
            return pd.DataFrame(columns=columns)
        vals = list(wb[name].values)
        if not vals:
            return pd.DataFrame(columns=columns)
        df = pd.DataFrame(vals[1:], columns=vals[0])
        for c in columns:
            if c not in df:
                df[c] = ""
        return df[columns]
    except Exception:
        return pd.DataFrame(columns=columns)


def load_aux(path=None):
    st.session_state.visits = load_sheet("Visit Records", VISIT_COLS, path)
    st.session_state.pics = load_sheet("PIC Management", PIC_COLS, path)
    st.session_state.photos = load_sheet("Photo Evidence", PHOTO_COLS, path)
    st.session_state.issues = load_sheet("Issues & Risks", ISSUE_COLS, path)


def get_data():
    if "tasks" not in st.session_state:
        source = MASTER if MASTER.exists() else DEFAULT_SOURCE
        if not source.exists():
            st.session_state.tasks = pd.DataFrame(columns=["ID", "Category", "Task Name", "Local Name", "Resource Names", "Duration", "Start", "Finish", "Predecessors", "Dependency", "Actual Start", "Actual Finish", "Status", "Progress"])
            st.session_state.wb = openpyxl.Workbook()
            st.session_state.wb.active.title = "Checklist"
        else:
            st.session_state.wb, st.session_state.tasks = read_workbook(source)
        load_aux()
    return st.session_state.tasks


def write_df_sheet(wb, name, df):
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.append(list(df.columns))
    for row in df.fillna("").itertuples(index=False, name=None):
        ws.append([v.isoformat() if isinstance(v, (date, datetime)) else v for v in row])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor=BLUE.lstrip("#"))


def save_all():
    tasks = st.session_state.tasks.copy()
    wb = st.session_state.wb
    ws = wb["Checklist"]
    row_by_id = {}
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, (int, float)):
            row_by_id[int(v)] = r
    mapping = {"Resource Names": 4, "Start": 6, "Finish": 7, "Predecessors": 8, "Actual Start": 10, "Actual Finish": 11, "Status": 12, "Progress": 13}
    for _, t in tasks.iterrows():
        rid = row_by_id.get(int(t["ID"]))
        if not rid:
            continue
        for key, col in mapping.items():
            value = t.get(key)
            if key in ("Start", "Finish", "Actual Start", "Actual Finish") and pd.notna(value):
                value = pd.Timestamp(value).to_pydatetime()
            elif pd.isna(value):
                value = None
            elif key == "Progress":
                value = f"{float(value):g}%"
            ws.cell(rid, col).value = value
    write_df_sheet(wb, "Visit Records", st.session_state.visits)
    write_df_sheet(wb, "PIC Management", st.session_state.pics)
    write_df_sheet(wb, "Photo Evidence", st.session_state.photos)
    write_df_sheet(wb, "Issues & Risks", st.session_state.issues)
    wb.save(MASTER)


# ───────────────────────── UI helpers ─────────────────────────
def html(s):
    st.markdown(s, unsafe_allow_html=True)


def hero(title, sub, stats=()):
    st_html = "".join(f'<div class="hero-stat"><b>{v}</b><span>{l}</span></div>' for v, l in stats)
    html(f'<div class="hero"><div style="z-index:1"><h1>{title}</h1><p>{sub}</p></div><div class="hero-stats">{st_html}</div></div>')


def kpi(label, value, sub="", color=BLUE, icon="📊", unit=""):
    bg = {BLUE: "#DCEBFF", GREEN: "#D1FADF", AMBER: "#FEF0C7", RED: "#FEE4E2", SLATE: "#EEF2F7", CYAN: "#CFF3FC"}.get(color, "#DCEBFF")
    u = f"<small>{unit}</small>" if unit else ""
    html(f'<div class="kpi"><div class="kpi-ic" style="background:{bg}">{icon}</div><div><div class="kpi-l">{label}</div><div class="kpi-v">{value}{u}</div><div class="kpi-s">{sub}</div></div></div>')


def card_open(title, sub=""):
    s = f'<span class="card-s">{sub}</span>' if sub else ""
    return f'<div class="card"><div class="card-h"><span class="card-t">{title}</span>{s}</div>'


def pbar(label, pct, color=BLUE, extra=""):
    pct = max(0, min(100, float(pct)))
    e = f"<em>{extra}</em>" if extra else ""
    return f'<div class="pb"><div class="pb-h"><span>{label}{e}</span><b style="color:{color}">{pct:.0f}%</b></div><div class="pb-t"><div class="pb-f" style="width:{pct}%;background:linear-gradient(90deg,{color},{color}bb)"></div></div></div>'


def ring(pct, color, label, sub=""):
    pct = max(0, min(100, float(pct)))
    return f'<div class="ring-card"><div class="ring" style="background:conic-gradient({color} {pct*3.6}deg,#DCE7F8 0)"><div class="ring-in">{pct:.0f}<small>%</small></div></div><div class="ring-t"><b>{label}</b><span>{sub}</span></div></div>'


def badge(text, kind="grey"):
    return f'<span class="badge b-{kind}">{escape(str(text))}</span>'


def status_badge(s):
    return badge(s, {"Completed": "green", "In Progress": "blue", "On Hold": "amber", "Not Started": "grey"}.get(s, "grey"))


def style_fig(fig, h=360, legend=True):
    fig.update_layout(
        height=h, margin=dict(l=8, r=8, t=28, b=8), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, Microsoft YaHei, sans-serif", color=NAVY, size=12),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title=None),
        showlegend=legend,
    )
    fig.update_xaxes(gridcolor="#E6EEFB", zeroline=False, linecolor="#D5E2F7")
    fig.update_yaxes(gridcolor="#E6EEFB", zeroline=False, linecolor="#D5E2F7")
    return fig


def card_chart(title, fig, sub="", key=None):
    with st.container(border=True, key=f"card_{key}"):
        st.markdown(f'<div class="card-h"><span class="card-t">{title}</span><span class="card-s">{sub}</span></div>', unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True, key=key, config={"displayModeBar": False})


def explode_resources(df):
    d = df.copy()
    d["Resource"] = d["Resource Names"].astype(str).str.split(",")
    d = d.explode("Resource")
    d["Resource"] = d["Resource"].astype(str).str.strip()
    return d[(d["Resource"] != "") & (d["Resource"].str.lower() != "none") & (d["Resource"].str.lower() != "nan")]


def pred_ids(v):
    return [int(x) for x in re.findall(r"\d+", str(v))] if v not in (None, "") else []


def fmt_d(d, f="%d %b"):
    return pd.Timestamp(d).strftime(f) if pd.notna(d) else "–"


def priority_of(l, i):
    s = (LEVELS.index(l) + 1 if l in LEVELS else 1) * (LEVELS.index(i) + 1 if i in LEVELS else 1)
    return "High" if s >= 6 else ("Medium" if s >= 3 else "Low")


# ───────────────────────── Load + derive ─────────────────────────
get_data()
tasks_src = st.session_state.tasks
today = date.today()

with st.sidebar:
    html('<div class="brand"><div class="brand-logo">E</div><div><div class="brand-t">EMART</div><div class="brand-s">STORE OPENING OS</div></div></div>')
    html('<div class="nav-cap">WORKSPACE</div>')
    PAGES = {
        "Overview": ":material/dashboard: Overview Dashboard",
        "Checklist": ":material/checklist: Opening Checklist",
        "Departments": ":material/bar_chart: Department Progress",
        "Timeline": ":material/event: Project Timeline",
        "Visits": ":material/pin_drop: Site Visit Records",
        "Risks": ":material/warning: Issue & Risk Tracker",
        "Workload": ":material/groups: PIC & Workload",
        "Photos": ":material/photo_camera: Photo Evidence",
        "Reports": ":material/database: Data & Reports",
    }
    page = st.radio("Navigation", list(PAGES), format_func=lambda k: PAGES[k], label_visibility="collapsed")
    st.divider()

if tasks_src.empty:
    with st.sidebar:
        st.warning("No tasks loaded yet.")
    hero("Welcome to Emart Store Opening OS", "Import your opening checklist Excel (must contain a sheet named “Checklist”).")
    up = st.file_uploader("Import checklist Excel", type=["xlsx"])
    if up and st.button("Import", type="primary"):
        wb, df = read_workbook(up)
        st.session_state.wb, st.session_state.tasks = wb, df
        load_aux(up)
        save_all()
        st.rerun()
    st.stop()

T = tasks_src.copy()
T["Status"] = T["Status"].fillna("Not Started")
T["Progress"] = pd.to_numeric(T["Progress"], errors="coerce").fillna(0).clip(0, 100)
T["Eff"] = T["Progress"].where(T["Status"] != "Completed", 100)
_fin = pd.to_datetime(T["Finish"], errors="coerce")
_sta = pd.to_datetime(T["Start"], errors="coerce")
T["Overdue"] = (_fin < pd.Timestamp(today)) & (T["Status"] != "Completed")
T["Days Late"] = ((pd.Timestamp(today) - _fin).dt.days).where(T["Overdue"], 0)
T["_s"], T["_f"] = _sta, _fin
T["Delayed"] = (_sta < pd.Timestamp(today)) & (T["Status"] == "Not Started")
T["Name"], T["Res"], T["Late"] = T["Task Name"], T["Resource Names"], T["Days Late"]

N = len(T)
n_done = int((T["Status"] == "Completed").sum())
n_prog = int((T["Status"] == "In Progress").sum())
n_hold = int((T["Status"] == "On Hold").sum())
n_not = int((T["Status"] == "Not Started").sum())
n_over = int(T["Overdue"].sum())
n_delay = int(T["Delayed"].sum())
overall = round(T["Eff"].mean(), 1) if N else 0

# target opening date (auto-guess, editable)
if "open_date" not in st.session_state:
    m = T[T["Task Name"].str.contains("grand opening|opening day|store opening", case=False, na=False)]
    guess = (m["Finish"].dropna().max() if len(m) and m["Finish"].notna().any() else T["Finish"].dropna().max()) if N else today
    st.session_state.open_date = guess if isinstance(guess, date) else today
with st.sidebar:
    html('<div class="nav-cap">PROJECT</div>')
    st.session_state.open_date = st.date_input("Target opening date", value=st.session_state.open_date)
    days_left = (st.session_state.open_date - today).days
    st.caption(f"{N} tasks · {overall:.0f}% complete")
    if MASTER.exists():
        st.caption("✅ Master data saved on server")
    st.caption("Import / export Excel in **Data & Reports**.")

# ───────────────────────── Shared task dialogs ─────────────────────────
def task_view(df):
    out = df[["ID", "Category", "Task Name", "Resource Names", "Start", "Finish", "Status", "Eff", "Days Late"]].copy()
    return out.rename(columns={"Resource Names": "Department / PIC", "Eff": "Progress", "Days Late": "Days late"})


TV_CFG = {"Progress": st.column_config.ProgressColumn("Progress", min_value=0, max_value=100, format="%.0f%%"),
          "ID": st.column_config.NumberColumn("ID", width="small"),
          "Task Name": st.column_config.TextColumn("Task Name", width="large")}


def show_task_df(f):
    if f.empty:
        st.info("No tasks in this group.")
    else:
        st.dataframe(task_view(f), hide_index=True, use_container_width=True, height=min(560, 40 + 35 * len(f)), column_config=TV_CFG)


def open_tasks_dialog(title, df, by_status=False):
    @st.dialog(title, width="large")
    def _d():
        if by_status:
            groups = [("Pending", df[df["Status"] == "Not Started"]), ("In Progress", df[df["Status"] == "In Progress"]),
                      ("On Hold", df[df["Status"] == "On Hold"]), ("Overdue", df[df["Overdue"]]), ("Completed", df[df["Status"] == "Completed"])]
            for tab, (lab, f) in zip(st.tabs([f"{l} ({len(f)})" for l, f in groups]), groups):
                with tab:
                    show_task_df(f)
        else:
            html(f'<div style="margin-bottom:8px">{badge(f"{len(df)} tasks", "blue")} {badge(f"{int(df.Overdue.sum())} overdue", "red")}</div>')
            show_task_df(df)
    _d()


# ───────────────────────── PAGE: Overview ─────────────────────────
def kpi_card(col, key, label, value, sub, color, icon, df, title):
    bg = {BLUE: "#DCEBFF", GREEN: "#D1FADF", AMBER: "#FEF0C7", RED: "#FEE4E2", SLATE: "#EEF2F7", CYAN: "#CFF3FC"}.get(color, "#DCEBFF")
    with col:
        with st.container(key=f"kpi_{key}"):
            html(f'<div class="kpi2"><div class="kpi-l">{label}</div><div class="kpi2-row"><div class="kpi-v">{value}</div><span class="kpi-ic" style="background:{bg}">{icon}</span></div>'
                 f'<div class="kpi-s">{sub}</div></div>')
            if st.button("View list", key=f"kb_{key}", use_container_width=True):
                open_tasks_dialog(title, df)


def page_overview():
    hero("Store Opening Control Centre", f"Emart new store opening · live status as of {today.strftime('%d %b %Y')}",
         [(f"{overall:.0f}%", "OVERALL PROGRESS"), (str(abs(days_left)), "DAYS TO OPENING" if days_left >= 0 else "DAYS PAST TARGET")])
    cols = st.columns(6)
    kpi_card(cols[0], "total", "TOTAL TASKS", N, f"{T['Category'].nunique()} categories", BLUE, "📋", T, "All tasks")
    kpi_card(cols[1], "done", "COMPLETED", n_done, f"{n_done/N*100:.0f}% of tasks", GREEN, "✅", T[T["Status"] == "Completed"], "Completed tasks")
    kpi_card(cols[2], "prog", "IN PROGRESS", n_prog, "Currently active", CYAN, "⚡", T[T["Status"] == "In Progress"], "In-progress tasks")
    kpi_card(cols[3], "not", "NOT STARTED", n_not, "Awaiting start", SLATE, "⏳", T[T["Status"] == "Not Started"], "Not-started tasks")
    kpi_card(cols[4], "over", "OVERDUE", n_over, "Past finish date", RED, "🚨", T[T["Overdue"]], "Overdue tasks (past finish date)")
    kpi_card(cols[5], "delay", "DELAYED START", n_delay, "Start passed, not begun", AMBER, "🕒", T[T["Delayed"]], "Delayed-start tasks")
    st.write("")

    c1, c2 = st.columns([1.55, 1])
    with c1:
        s_ = pd.to_datetime(T["Start"], errors="coerce"); f_ = pd.to_datetime(T["Finish"], errors="coerce")
        months = pd.period_range(min(s_.min(), f_.min()).to_period("M"), max(s_.max(), f_.max()).to_period("M"), freq="M")
        sc = s_.dt.to_period("M").value_counts(); fc = f_.dt.to_period("M").value_counts()
        xl = [m.strftime("%b<br>%Y") for m in months]
        fig = go.Figure()
        fig.add_bar(x=xl, y=[int(sc.get(m, 0)) for m in months], name="Tasks starting", marker_color=BLUE, text=[int(sc.get(m, 0)) or "" for m in months], textposition="outside")
        fig.add_bar(x=xl, y=[int(fc.get(m, 0)) for m in months], name="Tasks finishing", marker_color="#7CC4FF", text=[int(fc.get(m, 0)) or "" for m in months], textposition="outside")
        cur = pd.Period(today, "M")
        if cur in months:
            i = list(months).index(cur)
            fig.add_vrect(x0=i - .5, x1=i + .5, fillcolor=BLUE, opacity=.07, line_width=0)
            fig.add_annotation(x=i, y=1, yref="paper", text="This month", showarrow=False, font=dict(color=BLUE, size=11), yanchor="bottom")
        fig.update_layout(barmode="group", bargap=.25, bargroupgap=.05)
        fig.update_yaxes(title="Number of tasks")
        card_chart("Monthly Task Volume", style_fig(fig, 340), "Tasks scheduled to start / finish each month", key="ov_month")
    with c2:
        html(card_open("KPI Overview", "Task health") + '<div style="display:grid;gap:12px">'
             + ring(n_done / N * 100, GREEN, "Completion rate", f"{n_done} of {N} tasks done")
             + ring(n_prog / N * 100, BLUE, "In progress", f"{n_prog} tasks active")
             + ring(n_over / N * 100, RED, "Overdue rate", f"{n_over} tasks past due")
             + '</div></div>')

    c3, c4 = st.columns([1.35, 1])
    with c3:
        rows = ""
        for r in T.groupby("Category", sort=False).agg(n=("ID", "count"), p=("Eff", "mean"), done=("Status", lambda s: (s == "Completed").sum()),
                                                       prog=("Status", lambda s: (s == "In Progress").sum()), hold=("Status", lambda s: (s == "On Hold").sum()),
                                                       notst=("Status", lambda s: (s == "Not Started").sum()), late=("Overdue", "sum"),
                                                       s=("_s", "min"), f=("_f", "max")).reset_index().itertuples():
            col = GREEN if r.p >= 80 else (BLUE if r.p >= 40 else AMBER)
            seg = "".join(f'<i style="width:{v / r.n * 100}%;background:{c}"></i>' for v, c in [(r.done, GREEN), (r.prog, BLUE), (r.hold, AMBER), (r.notst, "#D5DEEB")] if v)
            late = badge(f"{int(r.late)} overdue", "red") if r.late else badge("on track", "green")
            rows += (f'<div class="cat"><div class="cat-h"><b>{escape(r.Category)}</b><span class="cat-p" style="color:{col}">{r.p:.0f}%</span></div>'
                     f'<div class="seg">{seg}</div>'
                     f'<div class="cat-s"><span>{fmt_d(r.s, "%d %b")} → {fmt_d(r.f, "%d %b")}</span>'
                     f'<span>{badge(f"{int(r.done)} done", "green")} {badge(f"{int(r.prog)} active", "blue")} {badge(f"{int(r.notst)} pending", "grey")} {late}</span></div></div>')
        legend = f'<div class="note" style="margin-bottom:6px"><span class="dot" style="background:{GREEN}"></span>Completed <span class="dot" style="background:{BLUE}"></span>In progress <span class="dot" style="background:{AMBER}"></span>On hold <span class="dot" style="background:#D5DEEB"></span>Not started · % = average task progress</div>'
        html(card_open("Progress by Category", f"{T['Category'].nunique()} work streams") + legend + rows + "</div>")
    with c4:
        sd = T.groupby("Status").size().reindex(STATUS_LIST).dropna().reset_index(name="n")
        fig = go.Figure(go.Pie(labels=sd["Status"], values=sd["n"], hole=.68, marker=dict(colors=[STATUS_COLORS[s] for s in sd["Status"]], line=dict(color="#fff", width=3)), textinfo="value", sort=False))
        fig.add_annotation(text=f"<b>{overall:.0f}%</b><br><span style='font-size:11px;color:#6B83A8'>overall</span>", showarrow=False, font=dict(size=26, color=NAVY))
        card_chart("Task Status Distribution", style_fig(fig, 300), key="ov_pie")
        up = T[(T["Status"] != "Completed") & (pd.to_datetime(T["Finish"], errors="coerce") >= pd.Timestamp(today))].sort_values("Finish").head(5)
        rows = "".join(f'<div class="row"><div class="date-chip">{fmt_d(r.Finish)}</div><div style="flex:1;min-width:0"><div class="row-t">{escape(r.Name)}</div><div class="row-s">{escape(str(r.Res) or "")}</div></div>{status_badge(r.Status)}</div>' for r in up.itertuples()) or '<div class="note">No upcoming tasks.</div>'
        html(card_open("Upcoming Milestones", "Next to finish") + rows + "</div>")

    # Key blockers: unfinished tasks that hold up the most unfinished downstream tasks
    succ = {}
    for r in T.itertuples():
        for p in pred_ids(r.Predecessors):
            succ.setdefault(p, []).append(r.ID)
    byid = T.set_index("ID")
    blk = []
    for p, ss in succ.items():
        if p in byid.index and byid.loc[p, "Status"] != "Completed":
            open_s = [x for x in ss if x in byid.index and byid.loc[x, "Status"] != "Completed"]
            if open_s:
                blk.append((p, len(open_s)))
    blk.sort(key=lambda x: -x[1])
    rows = ""
    for p, n_b in blk[:8]:
        r = byid.loc[p]
        rows += (f'<div class="row"><div style="flex:1;min-width:0"><div class="row-t">#{p} · {escape(str(r["Task Name"]))}</div>'
                 f'<div class="row-s">{escape(str(r["Resource Names"]) or "Unassigned")} · due {fmt_d(r["Finish"], "%d %b %Y")}</div></div>'
                 f'{badge(f"blocks {n_b} tasks", "amber")} {badge(f"{int(r["Days Late"])}d late", "red") if r["Overdue"] else status_badge(r["Status"])}</div>')
    html(card_open("Key Blockers", "Unfinished tasks holding up the most downstream tasks (from Predecessors)") + (rows or '<div class="note">No dependency blockers found.</div>') + "</div>")


# ───────────────────────── PAGE: Checklist ─────────────────────────
PROG_OPTS = ["0%", "25%", "50%", "75%", "100%"]


def snap(p):
    try:
        return f"{int(round(float(p) / 25) * 25)}%"
    except Exception:
        return "0%"


def page_checklist():
    hero("Opening Checklist", "All opening tasks — edit status, dates and progress directly in the table.", [(str(N), "TASKS"), (f"{n_done}", "COMPLETED")])
    f1, f2, f3, f4 = st.columns([1.5, 1.2, 1.1, 1])
    search = f1.text_input("Search", placeholder="Task name, resource, keyword…")
    cat = f2.selectbox("Category", ["All"] + T["Category"].drop_duplicates().tolist())
    res_opts = sorted(explode_resources(T)["Resource"].unique().tolist())
    dept = f3.selectbox("Department / PIC", ["All"] + res_opts)
    stat = f4.selectbox("Status", ["All"] + STATUS_LIST + ["Overdue"])
    view = tasks_src.copy()
    flags = T.set_index("ID")
    view["_over"] = view["ID"].map(flags["Overdue"])
    if search:
        view = view[view.astype(str).apply(lambda c: c.str.contains(search, case=False, na=False)).any(axis=1)]
    if cat != "All": view = view[view["Category"] == cat]
    if dept != "All": view = view[view["Resource Names"].astype(str).str.contains(re.escape(dept), na=False)]
    if stat == "Overdue": view = view[view["_over"] == True]
    elif stat != "All": view = view[view["Status"] == stat]
    chips = f'{badge(f"{len(view)} shown", "blue")} {badge(f"{int((view.Status=="Completed").sum())} completed", "green")} {badge(f"{int((view.Status=="In Progress").sum())} in progress", "blue")} {badge(f"{int(view._over.sum())} overdue", "red")}'
    html(f'<div style="margin:6px 0 10px 0">{chips} <span class="note" style="margin-left:8px">Edit cells (Status, dates, Progress 0/25/50/75/100%) → click Save Updates</span></div>')
    cols = ["ID", "Category", "Task Name", "Local Name", "Resource Names", "Start", "Finish", "Actual Start", "Actual Finish", "Status", "Progress", "Predecessors"]
    base = view[cols].copy()
    base["Progress"] = pd.to_numeric(base["Progress"], errors="coerce").fillna(0).apply(snap)
    orig_prog = dict(zip(base["ID"], base["Progress"]))
    orig_stat = dict(zip(base["ID"], base["Status"]))
    ver = st.session_state.get("ced", 0)
    edited = st.data_editor(base, use_container_width=True, hide_index=True, height=580,
        disabled=["ID", "Category", "Task Name", "Local Name"],
        column_config={
            "ID": st.column_config.NumberColumn("ID", width="small"),
            "Task Name": st.column_config.TextColumn("Task Name", width="large"),
            "Start": st.column_config.DateColumn("Start"), "Finish": st.column_config.DateColumn("Finish"),
            "Actual Start": st.column_config.DateColumn("Actual Start"), "Actual Finish": st.column_config.DateColumn("Actual Finish"),
            "Status": st.column_config.SelectboxColumn("Status", options=STATUS_LIST, required=True),
            "Progress": st.column_config.SelectboxColumn("Progress", options=PROG_OPTS, required=True, width="small"),
        }, key=f"checklist_editor_{ver}")
    b1, b2, _ = st.columns([1, 1.4, 4])
    if b1.button("💾 Save Updates", type="primary"):
        for _, row in edited.iterrows():
            idx = st.session_state.tasks.index[st.session_state.tasks["ID"] == row["ID"]]
            if not len(idx):
                continue
            i0 = idx[0]
            for col in ["Resource Names", "Start", "Finish", "Actual Start", "Actual Finish", "Status", "Predecessors"]:
                st.session_state.tasks.loc[i0, col] = row[col]
            if row["Progress"] != orig_prog.get(row["ID"]):
                st.session_state.tasks.loc[i0, "Progress"] = float(str(row["Progress"]).rstrip("%"))
            elif row["Status"] == "Completed" and orig_stat.get(row["ID"]) != "Completed":
                st.session_state.tasks.loc[i0, "Progress"] = 100.0
        save_all()
        st.session_state.ced = ver + 1
        st.success("Saved to master workbook.")
        st.rerun()
    b2.download_button("⬇ Download filtered CSV", view.drop(columns="_over").to_csv(index=False).encode("utf-8-sig"), "opening_checklist.csv", "text/csv")


# ───────────────────────── PAGE: Departments ─────────────────────────
def dept_summary():
    e = explode_resources(T)
    g = e.groupby("Resource").agg(Total=("ID", "count"), Completed=("Status", lambda s: (s == "Completed").sum()),
                                  InProgress=("Status", lambda s: (s == "In Progress").sum()),
                                  NotStarted=("Status", lambda s: (s == "Not Started").sum()),
                                  OnHold=("Status", lambda s: (s == "On Hold").sum()),
                                  Overdue=("Overdue", "sum"), Progress=("Eff", "mean")).reset_index()
    g["Completion"] = (g["Completed"] / g["Total"] * 100).round(1)
    g["Progress"] = g["Progress"].round(1)
    return g.sort_values("Total", ascending=False)


def page_departments():
    g = dept_summary()
    hero("Department Progress", "Task ownership, completion rate and overdue exposure for every department / resource.",
         [(str(len(g)), "DEPARTMENTS"), (str(n_over), "OVERDUE TASKS")])
    if g.empty:
        st.info("No Resource Names found in the checklist.")
        return
    e = explode_resources(T)
    html('<div class="note" style="margin:0 0 8px 2px">Click <b>View tasks</b> on any department to see its Pending, In Progress, On Hold, Overdue and Completed tasks.</div>')
    show_all = st.toggle(f"Show all {len(g)} departments", value=False)
    top = g if show_all else g.head(8)
    recs = list(top.itertuples())
    for i in range(0, len(recs), 4):
        cols = st.columns(4)
        for c, r in zip(cols, recs[i:i + 4]):
            col = GREEN if r.Progress >= 80 else (BLUE if r.Progress >= 40 else AMBER)
            with c:
                with st.container(key=f"dept_{i}_{r.Resource}"):
                    html(f'{ring(r.Progress, col, escape(r.Resource), f"{r.Total} tasks · {r.Completed} done")}'
                         f'<div style="margin-top:10px;display:flex;gap:6px;flex-wrap:wrap">{badge(f"{r.InProgress} active", "blue")}{badge(f"{r.NotStarted} pending", "grey")}{badge(f"{r.OnHold} hold", "amber") if r.OnHold else ""}{badge(f"{int(r.Overdue)} overdue", "red") if r.Overdue else badge("on track", "green")}</div>')
                    if st.button("View tasks", key=f"db_{i}_{r.Resource}", use_container_width=True):
                        open_tasks_dialog(f"{r.Resource} — {r.Total} tasks", e[e["Resource"] == r.Resource], by_status=True)

    c1, c2 = st.columns([1, 1])
    top15 = g.head(15)
    with c1:
        d = top15.sort_values("Completion")
        fig = go.Figure()
        for lab, col, colr in [("Completed", "Completed", GREEN), ("In Progress", "InProgress", BLUE), ("On Hold", "OnHold", AMBER), ("Not Started", "NotStarted", "#D5DEEB")]:
            pct = (d[col] / d["Total"] * 100)
            fig.add_bar(y=d["Resource"], x=pct, name=lab, orientation="h", marker_color=colr, customdata=d[col],
                        text=[str(v) if v else "" for v in d[col]], textposition="inside", insidetextanchor="middle",
                        textfont=dict(color="#fff" if lab != "Not Started" else NAVY, size=11), hovertemplate="%{y}<br>" + lab + ": %{customdata} tasks (%{x:.0f}%)<extra></extra>")
        fig.update_layout(barmode="stack", bargap=.28)
        fig.update_xaxes(range=[0, 100], ticksuffix="%")
        card_chart("Status Mix by Department", style_fig(fig, 30 * len(d) + 110), "Share of each department's tasks · numbers = task count · top 15", key="dp_mix")
    with c2:
        ee = e[e["Resource"].isin(top15["Resource"])].copy()
        ee["M"] = pd.to_datetime(ee["Finish"], errors="coerce").dt.to_period("M")
        ee = ee.dropna(subset=["M"])
        pv = ee.pivot_table(index="Resource", columns="M", values="ID", aggfunc="count", fill_value=0).reindex(top15["Resource"])
        fig = go.Figure(go.Heatmap(z=pv.values, x=[m.strftime("%b %Y") for m in pv.columns], y=pv.index, colorscale=[[0, "#F3F8FF"], [.35, "#9CC8FF"], [1, BLUE]],
                                   text=[[str(v) if v else "" for v in row] for row in pv.values], texttemplate="%{text}", xgap=3, ygap=3, showscale=False,
                                   hovertemplate="%{y}<br>%{x}: %{z} tasks due<extra></extra>"))
        fig.update_yaxes(autorange="reversed"); fig.update_xaxes(side="top")
        card_chart("Workload Calendar", style_fig(fig, 30 * len(top15) + 110, False), "Tasks due (finish date) per department per month · darker = busier", key="dp_heat")
    st.markdown('<div class="card-t" style="margin:6px 0 8px 4px">Department summary table</div>', unsafe_allow_html=True)
    st.dataframe(g.rename(columns={"Resource": "Department", "InProgress": "In Progress", "NotStarted": "Not Started", "OnHold": "On Hold"}),
                 hide_index=True, use_container_width=True,
                 column_config={"Completion": st.column_config.ProgressColumn("Completion rate", min_value=0, max_value=100, format="%.0f%%"),
                                "Progress": st.column_config.ProgressColumn("Avg progress", min_value=0, max_value=100, format="%.0f%%")})


# ───────────────────────── PAGE: Timeline ─────────────────────────
def gantt_fig(d, labels, highlight=True, window="Full project", texts=None):
    s = pd.to_datetime(d["Start"]); f = pd.to_datetime(d["Finish"]) + pd.Timedelta(days=1)
    dur = ((f - s).dt.total_seconds() * 1000).clip(lower=3 * 86400000)
    col = d["Status"].map(STATUS_COLORS).fillna(SLATE)
    prog = d["Eff"].clip(0, 100) / 100
    hov = [f"<b>{escape(l)}</b><br>{a:%d %b %Y} → {b:%d %b %Y}<br>{escape(str(r))}<br>{st_} · {p:.0f}%" for l, a, b, r, st_, p in
           zip(labels, s, f - pd.Timedelta(days=1), d["Resource Names"], d["Status"], d["Eff"])]
    lines = [RED if (o and highlight) else "rgba(0,0,0,0)" for o in d["Overdue"]]
    fig = go.Figure()
    fig.add_bar(y=labels, x=dur, base=s, orientation="h", marker=dict(color=list(col), opacity=.30, line=dict(color=lines, width=1.6)), hovertext=hov, hoverinfo="text",
                text=texts, textposition="inside", insidetextanchor="end", textfont=dict(color=NAVY, size=11), showlegend=False)
    fig.add_bar(y=labels, x=dur * prog, base=s, orientation="h", marker=dict(color=list(col)), hoverinfo="skip", showlegend=False)
    t0 = pd.Timestamp(today)
    fig.add_shape(type="line", x0=t0, x1=t0, y0=0, y1=1, yref="paper", line=dict(color=NAVY, dash="dash", width=1.6))
    fig.add_annotation(x=t0, y=0, yref="paper", text="Today", showarrow=False, font=dict(color=NAVY, size=11), yanchor="top", yshift=-4)
    fig.update_layout(barmode="overlay", bargap=.25, margin=dict(l=8, r=12, t=52, b=30))
    fig.update_yaxes(autorange="reversed", categoryorder="array", categoryarray=labels, title="", tickfont=dict(size=11), automargin=True)
    fig.update_xaxes(type="date", side="top", dtick="M1", tickformat="%b<br>%Y", tickangle=0, gridcolor="#D5E2F7", minor=dict(dtick=7 * 86400000, showgrid=True, gridcolor="#EEF4FD"), tickfont=dict(size=12, color=NAVY))
    if window == "Focus on today":
        fig.update_xaxes(range=[t0 - pd.Timedelta(days=30), t0 + pd.Timedelta(days=75)])
    return fig


def tl_legend():
    items = "".join(f'<span><span class="dot" style="background:{c}"></span>{s}</span>' for s, c in STATUS_COLORS.items())
    html(f'<div class="note" style="margin:2px 0 8px 2px;display:flex;gap:16px;flex-wrap:wrap">{items}<span>Dark = completed portion · light = remaining · red outline = overdue</span></div>')


def tl_summary(d):
    n = len(d)
    html('<div class="strip">'
         f'<div class="chip"><b>{n}</b><span>TASKS</span></div>'
         f'<div class="chip"><b>{d["Eff"].mean():.0f}%</b><span>AVG PROGRESS</span></div>'
         f'<div class="chip"><b>{int((d["Status"] == "Completed").sum())}</b><span>COMPLETED</span></div>'
         f'<div class="chip"><b>{int((d["Status"] == "In Progress").sum())}</b><span>IN PROGRESS</span></div>'
         f'<div class="chip"><b style="color:{RED}">{int(d["Overdue"].sum())}</b><span>OVERDUE</span></div>'
         f'<div class="chip wide"><b>{fmt_d(d["Start"].min())} → {fmt_d(d["Finish"].max())}</b><span>SCHEDULE SPAN</span></div></div>')


def timeline_panel(d, key, title, sub):
    sf = st.segmented_control("Status", ["All", "Not Started", "In Progress", "Completed", "On Hold", "Overdue"], default="All", key=f"{key}_sf") or "All"
    c2, c3, _ = st.columns([1.6, 1, 2])
    win = c2.segmented_control("Time window", ["Full project", "Focus on today"], default="Full project", key=f"{key}_win") or "Full project"
    hl = c3.toggle("Mark overdue", value=True, key=f"{key}_hl")
    tl_summary(d)
    if sf == "Overdue": d = d[d["Overdue"]]
    elif sf != "All": d = d[d["Status"] == sf]
    if d.empty:
        st.info("No tasks match this filter.")
        return
    d = d.sort_values(["Start", "ID"])
    labels = [f"{i} · {n[:46]}" for i, n in zip(d["ID"], d["Task Name"])]
    tl_legend()
    card_chart(title, style_fig(gantt_fig(d, labels, hl, win), 30 * len(d) + 120, False), sub, key=f"{key}_g")


def page_timeline():
    hero("Project Timeline", "Choose a view: the whole project, one category, or one department.", [(str(n_over), "OVERDUE"), (str(abs(days_left)), "DAYS TO OPENING" if days_left >= 0 else "DAYS PAST TARGET")])
    tl = T.dropna(subset=["Start", "Finish"]).copy()
    if tl.empty:
        st.info("No valid Start / Finish dates found.")
        return
    tl["Start"] = pd.to_datetime(tl["Start"]); tl["Finish"] = pd.to_datetime(tl["Finish"])
    view = st.segmented_control("View", ["Overall", "By Category", "By Department"], default="Overall", key="tl_view", label_visibility="collapsed") or "Overall"

    if view == "Overall":
        g = tl.groupby("Category", sort=False).agg(Start=("Start", "min"), Finish=("Finish", "max"), Eff=("Eff", "mean"), n=("ID", "count"), done=("Status", lambda s: (s == "Completed").sum()),
                                                    late=("Overdue", "sum")).reset_index()
        g["Status"] = ["Completed" if p >= 99.5 else ("In Progress" if p > 0 else "Not Started") for p in g["Eff"]]
        g["Overdue"] = (g["Finish"] < pd.Timestamp(today)) & (g["Eff"] < 99.5)
        g["Resource Names"] = g["n"].astype(str) + " tasks · " + g["done"].astype(str) + " done · " + g["late"].astype(int).astype(str) + " overdue"
        g["Task Name"] = g["Category"]; g["ID"] = 0
        html(f'<div class="strip"><div class="chip"><b>{fmt_d(tl["Start"].min(), "%d %b %Y")}</b><span>PROJECT START</span></div>'
             f'<div class="chip"><b>{fmt_d(st.session_state.open_date, "%d %b %Y")}</b><span>TARGET OPENING</span></div>'
             f'<div class="chip"><b>{abs(days_left)}</b><span>{"DAYS LEFT" if days_left >= 0 else "DAYS PAST TARGET"}</span></div>'
             f'<div class="chip"><b>{overall:.0f}%</b><span>OVERALL PROGRESS</span></div>'
             f'<div class="chip"><b>{len(g)}</b><span>PHASES</span></div></div>')
        tl_legend()
        labels = g["Category"].tolist()
        texts = [f"{p:.0f}%" for p in g["Eff"]]
        card_chart("Project Phases", style_fig(gantt_fig(g, labels, True, "Full project", texts), 56 * len(g) + 120, False), "One bar per category · switch to “By Category” to drill into its tasks", key="tl_phase")
        c1, c2 = st.columns(2)
        late = T[T["Overdue"]].sort_values("Days Late", ascending=False).head(10)
        rows = "".join(f'<div class="row"><div style="flex:1;min-width:0"><div class="row-t">#{r.ID} · {escape(r.Name)}</div><div class="row-s">{escape(str(r.Res) or "Unassigned")} · due {fmt_d(r.Finish, "%d %b")}</div></div>{badge(f"{int(r.Late)}d late", "red")}</div>' for r in late.itertuples()) or '<div class="note">No delayed tasks.</div>'
        c1.markdown(card_open("Most Delayed Tasks", f"{n_over} overdue in total") + rows + "</div>", unsafe_allow_html=True)
        up = T[(T["Status"] != "Completed") & (pd.to_datetime(T["Finish"], errors="coerce") >= pd.Timestamp(today))].sort_values("Finish").head(10)
        rows = "".join(f'<div class="row"><div class="date-chip">{fmt_d(r.Finish)}</div><div style="flex:1;min-width:0"><div class="row-t">{escape(r.Name)}</div><div class="row-s">{escape(str(r.Res) or "Unassigned")}</div></div>{status_badge(r.Status)}</div>' for r in up.itertuples()) or '<div class="note">No upcoming milestones.</div>'
        c2.markdown(card_open("Upcoming Milestones", "Next 10 to finish") + rows + "</div>", unsafe_allow_html=True)
        with st.expander("🔗 Dependency check (tasks that start before their predecessor finishes)"):
            lookup = T.set_index("ID"); recs = []
            for r in T.itertuples():
                for p in pred_ids(r.Predecessors):
                    if p in lookup.index:
                        pr = lookup.loc[p]
                        if pd.notna(pr["Finish"]) and pd.notna(r.Start) and pd.Timestamp(pr["Finish"]) >= pd.Timestamp(r.Start) and "after" in str(r.Dependency).lower():
                            recs.append({"Task ID": r.ID, "Task": r.Name, "Starts": r.Start, "Predecessor ID": p, "Predecessor": pr["Task Name"], "Pred. finishes": pr["Finish"], "Pred. status": pr["Status"]})
            st.dataframe(pd.DataFrame(recs), hide_index=True, use_container_width=True) if recs else st.success("No dependency conflicts found.")

    elif view == "By Category":
        cats = tl["Category"].drop_duplicates().tolist()
        cat = st.selectbox("Category", cats, key="tl_cat_sel")
        timeline_panel(tl[tl["Category"] == cat], "tlc", f"{cat} — Task Schedule", "Sorted by start date")
    else:
        e = explode_resources(tl)
        order = e.groupby("Resource")["ID"].count().sort_values(ascending=False).index.tolist()
        dept = st.selectbox("Department / PIC", order, key="tl_dept_sel", format_func=lambda x: f"{x}  ({int((e['Resource'] == x).sum())} tasks)")
        timeline_panel(e[e["Resource"] == dept], "tld", f"{dept} — Task Schedule", "Sorted by start date")


# ───────────────────────── PAGE: Visits ─────────────────────────
def page_visits():
    v = st.session_state.visits
    hero("Site Visit Records", "Record each visit separately — previous records are always retained.", [(str(len(v)), "VISITS LOGGED"), (str(v["Visit Date"].nunique() if len(v) else 0), "VISIT DAYS")])
    tab1, tab2 = st.tabs(["➕ New Visit", "🗂 Visit History"])
    with tab1:
        with st.form("visit_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            vdate = c1.date_input("Visit date", value=today)
            tid = c2.selectbox("Related task", T["ID"].tolist(), format_func=lambda x: f'{x} · {T.loc[T.ID == x, "Task Name"].iloc[0]}')
            vstat = c3.selectbox("Observed status", STATUS_LIST)
            obs = st.text_area("Observation / issue")
            act = st.text_area("Follow-up action")
            c4, c5 = st.columns(2)
            pic = c4.text_input("PIC"); rec = c5.text_input("Recorded by")
            ok = st.form_submit_button("Add Site Visit Record", type="primary")
        if ok:
            new = {"Visit ID": f"V{len(v)+1:04d}", "Visit Date": vdate.isoformat(), "Task ID": tid, "Task Name": T.loc[T.ID == tid, "Task Name"].iloc[0], "Status": vstat, "Observation": obs, "Follow-up Action": act, "PIC": pic, "Recorded By": rec}
            st.session_state.visits = pd.concat([v, pd.DataFrame([new])], ignore_index=True)
            save_all(); st.success("Visit record added."); st.rerun()
    with tab2:
        if v.empty:
            st.info("No visits recorded yet.")
        else:
            f1, f2 = st.columns([1, 1])
            fs = f1.selectbox("Filter by status", ["All"] + STATUS_LIST)
            q = f2.text_input("Search visits")
            d = v.sort_values("Visit Date", ascending=False)
            if fs != "All": d = d[d["Status"] == fs]
            if q: d = d[d.astype(str).apply(lambda c: c.str.contains(q, case=False, na=False)).any(axis=1)]
            for r in d.head(40).to_dict("records"):
                html(f'<div class="visit"><div style="display:flex;justify-content:space-between;align-items:center"><h4>{escape(str(r["Task Name"]))}<span class="note"> · Task #{r["Task ID"]}</span></h4>{status_badge(r["Status"])}</div>'
                     f'<div class="note">{escape(str(r["Visit Date"]))} · {escape(str(r["Visit ID"]))} · PIC: {escape(str(r["PIC"] or "–"))} · Recorded by {escape(str(r["Recorded By"] or "–"))}</div>'
                     f'<p><b>Observation:</b> {escape(str(r["Observation"] or "–"))}</p><p><b>Follow-up:</b> {escape(str(r["Follow-up Action"] or "–"))}</p></div>')
            st.download_button("⬇ Download visit records (CSV)", v.to_csv(index=False).encode("utf-8-sig"), "site_visit_records.csv", "text/csv")


# ───────────────────────── PAGE: Risks ─────────────────────────
def page_risks():
    iss = st.session_state.issues.copy()
    openi = iss[iss["Status"] != "Resolved"] if len(iss) else iss
    hi = int((openi["Priority"] == "High").sum()) if len(openi) else 0
    hero("Issue & Risk Tracker", "Log issues and risks, rate likelihood × impact, and follow mitigation to closure.", [(str(len(openi)), "OPEN"), (str(hi), "HIGH PRIORITY")])
    tab1, tab2 = st.tabs(["📋 Register", "➕ Log New Issue"])
    with tab1:
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi("TOTAL LOGGED", len(iss), "All issues & risks", BLUE, "🗒️")
        with c2: kpi("OPEN", len(openi), "Not yet resolved", AMBER, "🔓")
        with c3: kpi("HIGH PRIORITY", hi, "Needs escalation", RED, "🔥")
        with c4: kpi("RESOLVED", len(iss) - len(openi), "Closed out", GREEN, "✅")
        st.write("")
        m1, m2 = st.columns([1, 1.6])
        with m1:
            cells = ""
            colors = {(3, 1): "#F79009", (3, 2): "#F04438", (3, 3): "#B42318", (2, 1): "#FDB022", (2, 2): "#F79009", (2, 3): "#F04438", (1, 1): "#32D583", (1, 2): "#FDB022", (1, 3): "#F79009"}
            for li in (3, 2, 1):
                cells += f'<div style="display:flex;align-items:center;justify-content:flex-end;padding-right:6px">{LEVELS[li-1]}</div>'
                for ii in (1, 2, 3):
                    n = int(((openi["Likelihood"] == LEVELS[li - 1]) & (openi["Impact"] == LEVELS[ii - 1])).sum()) if len(openi) else 0
                    cells += f'<div class="mx" style="background:{colors[(li, ii)]}">{n}</div>'
            html(card_open("Risk Matrix", "Open issues") + f'<div class="matrix"><div></div><div>Low</div><div>Medium</div><div>High</div>{cells}</div><div class="note" style="margin-top:10px">Rows = Likelihood · Columns = Impact</div></div>')
        with m2:
            if len(openi):
                cc = openi.groupby("Category").size().reset_index(name="n").sort_values("n")
                fig = go.Figure(go.Bar(y=cc["Category"], x=cc["n"], orientation="h", marker_color=BLUE, text=cc["n"], textposition="outside"))
                card_chart("Open Issues by Category", style_fig(fig, 250, False), key="rk_cat")
            else:
                html(card_open("Open Issues by Category") + '<div class="note">Nothing logged yet — add your first issue in the next tab.</div></div>')
        if len(iss):
            ed = iss.copy()
            ed["Date Identified"] = pd.to_datetime(ed["Date Identified"], errors="coerce").dt.date
            ed["Due Date"] = pd.to_datetime(ed["Due Date"], errors="coerce").dt.date
            edited = st.data_editor(ed, hide_index=True, use_container_width=True, num_rows="dynamic", key="issue_editor",
                disabled=["Issue ID", "Priority"],
                column_config={"Category": st.column_config.SelectboxColumn(options=ISSUE_CATS), "Likelihood": st.column_config.SelectboxColumn(options=LEVELS),
                               "Impact": st.column_config.SelectboxColumn(options=LEVELS), "Status": st.column_config.SelectboxColumn(options=ISSUE_STATUS),
                               "Date Identified": st.column_config.DateColumn(), "Due Date": st.column_config.DateColumn(),
                               "Mitigation": st.column_config.TextColumn(width="large")})
            if st.button("💾 Save Register", type="primary"):
                e2 = edited.copy()
                e2["Priority"] = [priority_of(a, b) for a, b in zip(e2["Likelihood"], e2["Impact"])]
                for i, r in e2.iterrows():
                    if not r.get("Issue ID"): e2.loc[i, "Issue ID"] = f"R{i+1:04d}"
                st.session_state.issues = e2[ISSUE_COLS]
                save_all(); st.success("Risk register saved."); st.rerun()
    with tab2:
        with st.form("issue_form", clear_on_submit=True):
            title = st.text_input("Issue / risk title")
            c1, c2, c3 = st.columns(3)
            tid = c1.selectbox("Related task", T["ID"].tolist(), format_func=lambda x: f'{x} · {T.loc[T.ID == x, "Task Name"].iloc[0]}')
            cat = c2.selectbox("Category", ISSUE_CATS)
            owner = c3.text_input("Owner (PIC)")
            c4, c5, c6, c7 = st.columns(4)
            lik = c4.selectbox("Likelihood", LEVELS, index=1); imp = c5.selectbox("Impact", LEVELS, index=1)
            due = c6.date_input("Due date", value=today + timedelta(days=14)); stat = c7.selectbox("Status", ISSUE_STATUS)
            mit = st.text_area("Mitigation / action plan")
            ok = st.form_submit_button("Log Issue", type="primary")
        if ok and title:
            new = {"Issue ID": f"R{len(iss)+1:04d}", "Date Identified": today.isoformat(), "Title": title, "Related Task ID": tid, "Category": cat, "Likelihood": lik, "Impact": imp,
                   "Priority": priority_of(lik, imp), "Owner": owner, "Status": stat, "Due Date": due.isoformat(), "Mitigation": mit}
            st.session_state.issues = pd.concat([st.session_state.issues, pd.DataFrame([new])], ignore_index=True)
            save_all(); st.success("Issue logged."); st.rerun()


# ───────────────────────── PAGE: Workload ─────────────────────────
def page_workload():
    e = explode_resources(T)
    w = e.groupby("Resource").agg(Assigned=("ID", "count"), Completed=("Status", lambda s: (s == "Completed").sum()),
                                  Overdue=("Overdue", "sum"), Avg=("Eff", "mean")).reset_index()
    w["Outstanding"] = w["Assigned"] - w["Completed"]
    w["Avg"] = w["Avg"].round(0)
    w = w.sort_values("Outstanding", ascending=False)
    hero("PIC & Workload", "Who owns what — assigned, completed and outstanding tasks per person / department.", [(str(len(w)), "PEOPLE / TEAMS"), (str(int(w["Outstanding"].max()) if len(w) else 0), "MAX OUTSTANDING")])
    tab1, tab2 = st.tabs(["📊 Workload", "📇 PIC Directory"])
    with tab1:
        if w.empty:
            st.info("No Resource Names found.")
        else:
            thr = w["Outstanding"].quantile(.75)
            c1, c2 = st.columns([1.5, 1])
            with c1:
                d = w.head(15).sort_values("Outstanding")
                fig = go.Figure()
                fig.add_bar(y=d["Resource"], x=d["Completed"], name="Completed", orientation="h", marker_color=GREEN)
                fig.add_bar(y=d["Resource"], x=d["Outstanding"] - d["Overdue"], name="Outstanding", orientation="h", marker_color=BLUE)
                fig.add_bar(y=d["Resource"], x=d["Overdue"], name="Overdue", orientation="h", marker_color=RED)
                fig.update_layout(barmode="stack")
                card_chart("Individual Workload", style_fig(fig, 480), "Top 15 by outstanding tasks", key="wl_bar")
            with c2:
                top = w.head(6)
                bars = "".join(pbar(escape(r.Resource), r.Avg, GREEN if r.Avg >= 80 else (BLUE if r.Avg >= 40 else AMBER), f"{int(r.Outstanding)} open") for r in top.itertuples())
                html(card_open("Busiest Teams", "Avg task progress") + bars + "</div>")
                heavy = w[w["Outstanding"] >= thr].head(5)
                rows = "".join(f'<div class="row"><div class="row-t">{escape(r.Resource)}</div>{badge(f"{int(r.Outstanding)} open", "amber")}</div>' for r in heavy.itertuples())
                html(card_open("High-load Watchlist", "Top quartile") + rows + "</div>")
            st.dataframe(w.rename(columns={"Resource": "PIC / Department", "Avg": "Avg Progress %"}), hide_index=True, use_container_width=True,
                         column_config={"Avg Progress %": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.0f%%")})
    with tab2:
        with st.form("pic_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            name = c1.text_input("PIC name"); dept = c2.text_input("Department / team"); email = c3.text_input("Email")
            c4, c5 = st.columns(2)
            contact = c4.text_input("Contact"); notes = c5.text_input("Notes")
            ok = st.form_submit_button("Add PIC", type="primary")
        if ok and name:
            st.session_state.pics = pd.concat([st.session_state.pics, pd.DataFrame([{"PIC Name": name, "Department": dept, "Email": email, "Contact": contact, "Notes": notes}])], ignore_index=True)
            save_all(); st.success("PIC added."); st.rerun()
        if len(st.session_state.pics):
            ep = st.data_editor(st.session_state.pics, use_container_width=True, hide_index=True, num_rows="dynamic")
            if st.button("💾 Save PIC Changes", type="primary"):
                st.session_state.pics = ep; save_all(); st.success("Saved."); st.rerun()
        else:
            st.info("No PIC directory records yet.")


# ───────────────────────── PAGE: Photos ─────────────────────────
def page_photos():
    ph = st.session_state.photos
    valid = ph[[ (PHOTO_DIR / str(f)).exists() for f in ph["File Name"] ]] if len(ph) else ph
    hero("Photo Evidence", "Upload site photos and link each one to a checklist task.", [(str(len(valid)), "PHOTOS"), (str(valid["Task ID"].nunique() if len(valid) else 0), "TASKS COVERED")])
    tab1, tab2 = st.tabs(["🖼 Gallery", "📤 Upload"])
    with tab2:
        with st.form("photo_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            pdate = c1.date_input("Photo date", value=today)
            tid = c2.selectbox("Related task", T["ID"].tolist(), format_func=lambda x: f'{x} · {T.loc[T.ID == x, "Task Name"].iloc[0]}')
            cap = st.text_input("Caption / description"); by = st.text_input("Uploaded by")
            files = st.file_uploader("Choose photo(s)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
            ok = st.form_submit_button("Save Photo Evidence", type="primary")
        if ok and files:
            tname = T.loc[T.ID == tid, "Task Name"].iloc[0]
            for f in files:
                safe = re.sub(r"[^A-Za-z0-9._-]", "_", f.name)
                nn = f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{safe}"
                (PHOTO_DIR / nn).write_bytes(f.getvalue())
                st.session_state.photos = pd.concat([st.session_state.photos, pd.DataFrame([{"Photo ID": f"P{len(st.session_state.photos)+1:04d}", "Date": pdate.isoformat(), "Task ID": tid, "Task Name": tname, "Caption": cap, "File Name": nn, "Uploaded By": by}])], ignore_index=True)
            save_all(); st.success("Photo evidence saved."); st.rerun()
    with tab1:
        if valid.empty:
            st.info("No photos yet — upload some in the next tab.")
        else:
            opts = ["All"] + sorted(valid["Task ID"].astype(str).unique().tolist())
            sel = st.selectbox("Filter by task ID", opts)
            d = valid if sel == "All" else valid[valid["Task ID"].astype(str) == sel]
            d = d.sort_values("Date", ascending=False)
            recs = d.to_dict("records")
            for i in range(0, len(recs), 4):
                cols = st.columns(4)
                for c, r in zip(cols, recs[i:i + 4]):
                    with c:
                        st.image(str(PHOTO_DIR / str(r["File Name"])), use_container_width=True)
                        html(f'<div style="margin:-6px 0 14px 0"><div class="row-t">{escape(str(r["Caption"] or "Untitled"))}</div><div class="row-s">Task #{r["Task ID"]} · {escape(str(r["Date"]))}</div></div>')


# ───────────────────────── PAGE: Reports ─────────────────────────
def build_report():
    buf = BytesIO()
    summary = pd.DataFrame({"Metric": ["Generated", "Total tasks", "Completed", "In progress", "Not started", "On hold", "Overdue", "Overall progress %", "Target opening date", "Days to opening"],
                            "Value": [today.isoformat(), N, n_done, n_prog, n_not, n_hold, n_over, overall, st.session_state.open_date.isoformat(), days_left]})
    cat = T.groupby("Category", sort=False).agg(Tasks=("ID", "count"), Completed=("Status", lambda s: (s == "Completed").sum()), Overdue=("Overdue", "sum"), AvgProgress=("Eff", "mean")).reset_index().round(1)
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        cat.to_excel(xw, sheet_name="By Category", index=False)
        dept_summary().to_excel(xw, sheet_name="By Department", index=False)
        T[T["Overdue"]][["ID", "Category", "Task Name", "Resource Names", "Finish", "Status", "Days Late"]].sort_values("Days Late", ascending=False).to_excel(xw, sheet_name="Overdue Tasks", index=False)
        for ws in xw.book.worksheets:
            for c in ws[1]:
                c.font = openpyxl.styles.Font(bold=True, color="FFFFFF"); c.fill = openpyxl.styles.PatternFill("solid", fgColor=BLUE.lstrip("#"))
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(48, max(12, max(len(str(c.value or "")) for c in col) + 2))
    return buf.getvalue()


def page_reports():
    hero("Data & Reports", "Import a workbook, export the master file and download management reports.")
    c1, c2 = st.columns(2)
    with c1:
        html(card_open("Import Excel", "Replaces current tasks") + '<div class="note">Workbook must contain a sheet named “Checklist”. Existing visit / PIC / photo / risk sheets inside the file are imported too.</div></div>')
        up = st.file_uploader("Choose .xlsx", type=["xlsx"], label_visibility="collapsed")
        if up and st.button("⬆ Import Excel", type="primary"):
            try:
                wb, df = read_workbook(up)
                st.session_state.wb, st.session_state.tasks = wb, df
                load_aux(up); save_all()
                st.success(f"Imported {len(df)} tasks."); st.rerun()
            except Exception as ex:
                st.error(f"Import failed: {ex}")
    with c2:
        html(card_open("Export", "Download current data") + '<div class="note">Master workbook keeps your original Checklist format plus the extra sheets.</div></div>')
        if MASTER.exists():
            st.download_button("⬇ Master Workbook (.xlsx)", MASTER.read_bytes(), "Emart_Store_Opening_Updated.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        else:
            st.caption("Master workbook is created after your first save.")
        st.download_button("📑 Executive Report (.xlsx)", build_report(), f"Emart_Executive_Report_{today:%Y%m%d}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        d1, d2, d3 = st.columns(3)
        d1.download_button("Checklist CSV", tasks_src.to_csv(index=False).encode("utf-8-sig"), "checklist.csv", "text/csv", use_container_width=True)
        d2.download_button("Visits CSV", st.session_state.visits.to_csv(index=False).encode("utf-8-sig"), "visits.csv", "text/csv", use_container_width=True)
        d3.download_button("Risks CSV", st.session_state.issues.to_csv(index=False).encode("utf-8-sig"), "risks.csv", "text/csv", use_container_width=True)
    st.write("")
    checks = [("Tasks without Start / Finish date", int(T[["Start", "Finish"]].isna().any(axis=1).sum())),
              ("Tasks without a resource / PIC", int((T["Resource Names"].astype(str).str.strip().isin(["", "None"])).sum())),
              ("Completed tasks with progress < 100%", int(((tasks_src["Status"] == "Completed") & (pd.to_numeric(tasks_src["Progress"], errors="coerce").fillna(0) < 100)).sum())),
              ("Not Started tasks showing progress > 0%", int(((T["Status"] == "Not Started") & (T["Progress"] > 0)).sum())),
              ("Finish date earlier than Start date", int((_fin < _sta).sum()))]
    rows = "".join(f'<div class="row"><div class="row-t">{l}</div>{badge("OK", "green") if n == 0 else badge(f"{n} to review", "amber")}</div>' for l, n in checks)
    html(card_open("Data Health Check", "Quality of the loaded checklist") + rows + "</div>")
    html('<div class="note">Photos are stored in <code>data/photos</code>; include it in backups. For multi-user live editing, deploy one shared instance with persistent storage.</div>')


{"Overview": page_overview, "Checklist": page_checklist, "Departments": page_departments, "Timeline": page_timeline,
 "Visits": page_visits, "Risks": page_risks, "Workload": page_workload, "Photos": page_photos, "Reports": page_reports}[page]()

html('<div class="note" style="text-align:center;margin-top:26px">EMART STORE OPENING OS · v2.1</div>')
