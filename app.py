import streamlit as st
import gspread
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.dates as mdates
import numpy as np
from scipy import stats
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from streamlit_echarts import st_echarts
from google.oauth2.service_account import Credentials
import os
import html
import admin_pages
from dotenv import load_dotenv

load_dotenv()

SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
SERVICE_ACCOUNT_FILE = "service_account.json"
SCOPES_SHEETS = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

st.set_page_config(page_title="Success Tutoring Dashboard", page_icon="logo.png" if os.path.exists("logo.png") else None, layout="wide")

# ── Theme ─────────────────────────────────────────────────────────────────────
BI_BG       = "#f3f5f6"   # report canvas (light grey)
BI_CARD     = "#ffffff"   # cards and panels
BI_ACCENT   = "#1f9a9a"   # Success Tutoring teal
BI_BLUE     = "#2f6fb5"
BI_GREEN    = "#2e8540"
BI_ORANGE   = "#e0562a"   # Success Tutoring orange
BI_RED      = "#c8322f"
BI_YELLOW   = "#c99a06"
BI_PURPLE   = "#7a5cb0"
BI_GRAY     = "#e3e8ea"
BI_BORDER   = "#dde5e6"
BI_TEXT     = "#1c2a30"
BI_SUBTEXT  = "#5d6d73"
BI_CHART_BG = "#ffffff"
BI_GRID     = "#e8edee"
BI_INK      = "#1c2a30"   # top bar

# One chart palette used across all reports, led by the brand teal and orange
SERIES_COLORS = [BI_ACCENT, BI_ORANGE, BI_BLUE, BI_PURPLE, BI_YELLOW,
                 "#3b8f5a", "#d4679b", "#8a5a44", "#4fb3d9", "#5d6d73"]
REGION_TO_STATE = {
    "New South Wales": "NSW",
    "Victoria": "VIC",
    "Queensland": "QLD",
    "Western Australia": "WA",
    "South Australia": "SA",
    "New Zealand": "NZ",
    "Auckland (N)": "NZ",
    "Canterbury (S)": "NZ",
    "Manawati-Whanganui": "NZ",
}

HOLIDAY_COLORS = {
    "Autumn": "rgba(255,140,0,",
    "Winter": "rgba(30,144,255,",
    "Spring": "rgba(50,205,50,",
    "Summer": "rgba(255,215,0,",
}
st.markdown(f"""
<style>
    html, body, [class*="css"] {{
        font-family: 'Segoe UI', 'Source Sans Pro', 'Source Sans 3', system-ui, sans-serif;
    }}
    .stApp {{ background-color: {BI_BG}; }}
    .main .block-container, div[data-testid="stMainBlockContainer"] {{ padding-top: 3.75rem; max-width: 1400px; }}

    /* Top bar */
    .topbar {{
        background: {BI_INK}; color: #ffffff; border-radius: 8px;
        padding: 10px 18px; margin-bottom: 18px;
        display: flex; justify-content: space-between; align-items: center; gap: 12px; flex-wrap: wrap;
    }}
    .topbar-title {{ font-size: 1.05em; font-weight: 600; }}
    .topbar-title span {{ color: #9fb3b9; font-weight: 400; }}
    .topbar-meta {{ font-size: 0.85em; color: #b9c7cb; }}
    .topbar-meta b {{ color: #ffffff; font-weight: 600; text-transform: capitalize; }}

    /* KPI cards */
    .metric-card {{
        background: {BI_CARD}; border: 1px solid {BI_BORDER};
        border-radius: 8px; padding: 14px 16px; margin: 4px 0 10px;
        box-shadow: 0 1px 2px rgba(28,42,48,0.05);
    }}
    .metric-label {{
        font-size: 0.78em; color: {BI_SUBTEXT}; margin-bottom: 4px; font-weight: 600;
        display: flex; align-items: center; gap: 6px;
    }}
    .metric-label::before {{
        content: ""; width: 8px; height: 8px; border-radius: 50%; background: {BI_SUBTEXT}; flex: none;
    }}
    .metric-card.green  .metric-label::before {{ background: {BI_ACCENT}; }}
    .metric-card.orange .metric-label::before {{ background: {BI_ORANGE}; }}
    .metric-card.red    .metric-label::before {{ background: {BI_RED}; }}
    .metric-card.blue   .metric-label::before {{ background: {BI_BLUE}; }}
    .metric-card.purple .metric-label::before {{ background: {BI_PURPLE}; }}
    .metric-value {{ font-size: 1.9em; font-weight: 700; color: {BI_TEXT}; line-height: 1.15;
                     font-variant-numeric: tabular-nums; }}
    .metric-delta {{ font-size: 0.8em; margin-top: 6px; font-weight: 600; }}
    .metric-delta span {{ color: {BI_SUBTEXT}; font-weight: 400; }}

    /* Titles */
    .report-title {{ font-size: 1.5em; font-weight: 700; color: {BI_TEXT}; margin-bottom: 0; }}
    .report-subtitle {{ font-size: 0.88em; color: {BI_SUBTEXT}; margin-bottom: 14px; }}
    .section-header {{
        font-size: 1.02em; font-weight: 700; color: {BI_TEXT};
        margin: 22px 0 10px 0; padding: 0 0 6px 0;
        border-bottom: 1px solid {BI_BORDER};
    }}
    .gauge-label {{
        text-align: center; color: {BI_SUBTEXT}; font-size: 0.82em; font-weight: 600;
        margin-top: -6px; margin-bottom: 8px;
    }}
    .filter-label {{
        font-size: 0.72em; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase;
        color: {BI_SUBTEXT}; margin-bottom: 2px;
    }}

    /* Panels: filters, expanders, tables, charts */
    div[data-testid="stVerticalBlockBorderWrapper"] {{ background: {BI_CARD}; border-radius: 8px; }}
    details[data-testid="stExpander"] {{
        background: {BI_CARD} !important; border: 1px solid {BI_BORDER} !important;
        border-radius: 8px !important; margin-bottom: 12px;
    }}
    details[data-testid="stExpander"] summary {{ font-weight: 600 !important; }}
    .stDataFrame {{ border: 1px solid {BI_BORDER}; border-radius: 6px; }}
    div[data-testid="stPlotlyChart"] {{
        background: {BI_CARD}; border: 1px solid {BI_BORDER}; border-radius: 8px; padding: 6px;
    }}
    .stButton > button {{ border-radius: 6px; font-weight: 600; }}
    hr {{ border-color: {BI_BORDER}; margin: 12px 0; }}

    /* Sidebar */
    section[data-testid="stSidebar"] {{
        background-color: {BI_CARD} !important;
        border-right: 1px solid {BI_BORDER};
        min-width: 260px !important; max-width: 260px !important; width: 260px !important;
    }}
    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] {{ gap: 2px !important; }}
    .nav-group {{
        font-size: 0.7em; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase;
        color: #8a989d; padding: 12px 6px 0;
        margin-bottom: calc(1rem + 4px);  /* offsets Streamlit's -1rem margin under markdown */
    }}
    section[data-testid="stSidebar"] .stButton > button {{
        width: 100%; justify-content: flex-start; text-align: left;
        padding: 6px 10px; min-height: 36px; border-radius: 6px;
        font-size: 0.9em; font-weight: 500;
    }}
    section[data-testid="stSidebar"] .stButton > button > div {{ justify-content: flex-start; }}
    section[data-testid="stSidebar"] .stButton > button p {{ font-size: 0.95em; }}
    section[data-testid="stSidebar"] .stButton > button[kind="secondary"],
    section[data-testid="stSidebar"] button[data-testid="stBaseButton-secondary"] {{
        background: transparent; border: 1px solid transparent; color: {BI_TEXT};
    }}
    section[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover,
    section[data-testid="stSidebar"] button[data-testid="stBaseButton-secondary"]:hover {{
        background: {BI_BG}; border-color: {BI_BORDER}; color: {BI_TEXT};
    }}
    section[data-testid="stSidebar"] .stButton > button[kind="primary"],
    section[data-testid="stSidebar"] button[data-testid="stBaseButton-primary"] {{
        background: {BI_ACCENT}; border: 1px solid {BI_ACCENT}; color: #ffffff; font-weight: 600;
    }}
    .side-footer {{ border-top: 1px solid {BI_BORDER}; margin-top: 14px; padding-top: 10px; }}
    .side-updated {{ font-size: 0.78em; color: {BI_SUBTEXT}; }}
    .user-card {{ display: flex; align-items: center; gap: 10px; padding: 10px 2px 8px; }}
    .user-avatar {{
        width: 32px; height: 32px; border-radius: 50%; flex: none;
        background: {BI_ORANGE}; color: #ffffff; display: grid; place-items: center;
        font-weight: 700; font-size: 0.9em;
    }}
    .user-text {{ min-width: 0; line-height: 1.25; }}
    .user-name {{ font-weight: 600; font-size: 0.9em; color: {BI_TEXT}; }}
    .user-mail {{ font-size: 0.75em; color: {BI_SUBTEXT}; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
</style>
""", unsafe_allow_html=True)

# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_resource
def get_sheets_client():
    try:
        # Streamlit Cloud — read from secrets
        import json
        service_account_info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES_SHEETS)
    except:
        # Local — read from file
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES_SHEETS)
    return gspread.authorize(creds)
@st.cache_data(ttl=300)
def load_sheet_data(tab_name, attempts=3):
    """Read a whole tab. Retries brief network failures (e.g. 'Response ended prematurely');
    a missing tab fails straight away."""
    import time
    for attempt in range(attempts):
        try:
            client = get_sheets_client()
            sheet = client.open_by_key(SHEET_ID)
            worksheet = sheet.worksheet(tab_name)
            return pd.DataFrame(worksheet.get_all_records())
        except gspread.exceptions.WorksheetNotFound:
            raise
        except Exception:
            if attempt == attempts - 1:
                raise
            time.sleep(2 ** attempt)

@st.cache_data(ttl=300)
def load_weekly_membership():
    df = load_sheet_data("Weekly Membership")
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
        df = df.dropna(subset=["Date"])
    num_cols = ["# Active members","# New members","# Suspended members",
                "# Cancelled members","Onboarding Members","Onboarding week","Age (Months)"]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


@st.cache_data(ttl=300)
def _load_school_holidays():
    df = load_sheet_data("School Holidays")
    df["Year"]       = pd.to_numeric(df["Year"],       errors="coerce")
    df["Start_Week"] = pd.to_numeric(df["Start_Week"], errors="coerce")
    df["End_Week"]   = pd.to_numeric(df["End_Week"],   errors="coerce")
    return df.dropna(subset=["Year","Start_Week","End_Week"])

def load_school_holidays():
    try:
        return _load_school_holidays()
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=300)
def _load_revenue():
    df = load_sheet_data("Revenue")
    df = df.rename(columns={"Location": "Success Tutoring - Business name"})
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
    for c in ["# Active Students","Total Sessions","Gross Revenue","Net Revenue",
              "Student Visits","Revenue per Session","Revenue per Student",
              "Sessions per Student","Student per Session",
              "Sessions per Student Visit","Student Visits per Session"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c].astype(str).str.replace("[$,]","",regex=True), errors="coerce").fillna(0)
    return df

def load_revenue():
    try:
        return _load_revenue()
    except Exception as e:
        st.error(f"Could not load the Revenue sheet ({e}). Click Refresh data to try again.")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def _load_vlookup():
    df = load_sheet_data("Vlookup")
    df = df.rename(columns={
        "Location": "Success Tutoring - Business name",
        "Location Start": "Location Start Date",
        "Months old": "Age (Months)",
        "Onboarding Week": "Onboarding week",
    })
    for c in ["Onboarding week","Onboarding Members","Age (Months)","Weeks old"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    if "Weeks old" in df.columns:
        # Week 1 = first week with members; months = whole months elapsed since then.
        elapsed_days = (df["Weeks old"] - 1).clip(lower=0) * 7
        df["Age (Months)"] = (elapsed_days / 30.4375).astype(int)
    return df

def load_vlookup():
    """Cached Vlookup; a failed read is not cached, so the next page load tries again."""
    try:
        return _load_vlookup()
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=300)
def load_permissions():
    client = get_sheets_client()
    sheet = client.open_by_key(SHEET_ID)
    worksheet = sheet.worksheet("Permissions")
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    permissions = {}
    for _, row in df.iterrows():
        email = str(row["email"]).strip().lower()
        tabs = [t.strip() for t in str(row["allowed_tabs"]).split(",") if t.strip()]
        gpm_filter = str(row.get("gpm_filter","")).strip()
        access_level = str(row.get("access_level","admin")).strip().lower()
        permissions[email] = {
            "tabs": tabs,
            "gpm_filter": gpm_filter,
            "access_level": access_level,
            "pin": str(row.get("pin","")).strip(),
        }
    return permissions
def log_access(email, name, action):
    try:
        client = get_sheets_client()
        sheet = client.open_by_key(SHEET_ID)
        worksheet = sheet.worksheet("Access Log")
        from datetime import datetime
        timestamp = datetime.now().strftime("%d %b %Y %I:%M %p")
        worksheet.append_row([email, name, action, timestamp])
    except Exception as e:
        pass

def flag_unknown_user(email):
    try:
        client = get_sheets_client()
        sheet = client.open_by_key(SHEET_ID)
        worksheet = sheet.worksheet("Pending Approvals")
        from datetime import datetime
        timestamp = datetime.now().strftime("%d %b %Y %I:%M %p")
        existing = worksheet.get_all_records()
        emails = [str(r.get("Email","")).strip().lower() for r in existing]
        if email.strip().lower() not in emails:
            worksheet.append_row([email, timestamp, "Pending"])
    except Exception as e:
        pass
def get_user_permissions(user_email):
    st.cache_data.clear()
    permissions = load_permissions()
    perms = permissions.get(user_email.strip().lower(), {
        "tabs": [], "gpm_filter": "", "access_level": "gpm", "pin": ""
    })

    # If GPM, derive their allowed locations from Vlookup sheet automatically
    if perms.get("access_level") == "gpm":
        gpm_name = perms.get("gpm_filter", "").strip()
        if gpm_name:
            try:
                vl = load_vlookup()
                allowed_locs = vl[
                (vl["GPM"] == gpm_name) &
                (vl["Stage"] != "Leasing")
                ]["Success Tutoring - Business name"].tolist()
                perms["allowed_locations"] = allowed_locs
            except Exception as ex:
                perms["allowed_locations"] = []
        else:
            perms["allowed_locations"] = []

    return perms

# ── Helpers ───────────────────────────────────────────────────────────────────
def churn_rate(cancelled, active):
    try:
        a = float(active)
        return round(float(cancelled) / a * 100, 1) if a > 0 else 0.0
    except:
        return 0.0

def metric_card(label, value, delta=None, color="", higher_is_better=True):
    delta_html = ""
    if delta is not None:
        arrow = "▲" if delta > 0 else "▼" if delta < 0 else "●"
        good = delta > 0 if higher_is_better else delta < 0
        dcolor = BI_SUBTEXT if delta == 0 else (BI_GREEN if good else BI_RED)
        dtext = f"{abs(delta):,.0f}" if float(delta).is_integer() or abs(delta) >= 100 else f"{abs(delta):,.1f}"
        delta_html = f'<div class="metric-delta" style="color:{dcolor}">{arrow} {dtext} <span>vs last week</span></div>'
    st.markdown(f"""<div class="metric-card {color}">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
    </div>""", unsafe_allow_html=True)

def get_age_group(months):
    try:
        m = float(months)
        if m <= 3:    return "0–3 months"
        elif m <= 6:  return "3–6 months"
        elif m <= 12: return "6–12 months"
        elif m <= 24: return "12–24 months"
        else:         return "24+ months"
    except:
        return "Unknown"

AGE_ORDER = ["0–3 months","3–6 months","6–12 months","12–24 months","24+ months"]

def gauge_chart(value, max_val, color):
    fig, ax = plt.subplots(figsize=(3.2,2.2), subplot_kw={"projection":"polar"})
    fig.patch.set_facecolor(BI_CARD); ax.set_facecolor(BI_CARD)
    theta_bg = np.linspace(np.pi,0,200)
    ax.plot(theta_bg,[0.75]*200,color=BI_GRAY,linewidth=22,solid_capstyle="round")
    pct = min(value/max(max_val,1),1.0)
    theta_val = np.linspace(np.pi,np.pi-pct*np.pi,200)
    ax.plot(theta_val,[0.75]*200,color=color,linewidth=22,solid_capstyle="round")
    ax.plot(theta_val,[0.75]*200,color=color,linewidth=30,solid_capstyle="round",alpha=0.12)
    ax.set_ylim(0,1); ax.set_xlim(0,np.pi); ax.axis("off")
    ax.text(np.pi/2,0.18,f"{value:,}",ha="center",va="center",
            fontsize=22,fontweight="bold",color=BI_TEXT,transform=ax.transData)
    pct_label = f"{pct*100:.0f}%" if max_val > 0 else ""
    ax.text(np.pi/2,-0.15,pct_label,ha="center",va="center",
            fontsize=9,color=color,transform=ax.transData,fontweight="600")
    plt.tight_layout(pad=0)
    return fig

def bi_fig(w=14,h=6):
    fig,ax = plt.subplots(figsize=(w,h))
    fig.patch.set_facecolor(BI_CHART_BG); ax.set_facecolor(BI_CHART_BG)
    ax.tick_params(colors=BI_TEXT,labelsize=9)
    for spine in ["bottom","left"]:
        ax.spines[spine].set_color(BI_BORDER); ax.spines[spine].set_linewidth(0.8)
    for spine in ["top","right"]: ax.spines[spine].set_visible(False)
    ax.grid(axis="y",color=BI_GRID,linewidth=0.5,alpha=0.6)
    return fig,ax

# ── Plotly theme ──────────────────────────────────────────────────────────────
CHART_FONT = "Source Sans Pro, Source Sans 3, Segoe UI, Arial, sans-serif"
PLOTLY_LAYOUT = dict(
    paper_bgcolor="#ffffff", plot_bgcolor="#ffffff",
    font=dict(color=BI_TEXT, family=CHART_FONT, size=12),
    xaxis=dict(gridcolor=BI_GRID, gridwidth=1, tickfont=dict(color=BI_SUBTEXT, size=11),
               linecolor=BI_BORDER, linewidth=1, showgrid=False, zeroline=False, tickangle=0),
    yaxis=dict(gridcolor="#eef2f3", gridwidth=1, tickfont=dict(color=BI_SUBTEXT, size=11),
               linecolor=BI_BORDER, showgrid=True, zeroline=False, tickformat=",", rangemode="tozero"),
    legend=dict(bgcolor="rgba(0,0,0,0)", borderwidth=0, font=dict(color=BI_TEXT, size=12),
                orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0, itemsizing="constant"),
    hovermode="x unified",
    hoverlabel=dict(bgcolor="#ffffff", bordercolor=BI_BORDER,
                    font=dict(color=BI_TEXT, size=12, family=CHART_FONT), namelength=-1),
    margin=dict(l=60, r=30, t=40, b=50), dragmode=False,
)

def std_traces(fig, df, date_col, col_name, color, label, secondary_y=False):
    """One clean line with a soft area fill (the fill is dropped by show_chart when a
    chart has several lines) and a dot on the latest point."""
    try:
        r,g,b = int(color[1:3],16),int(color[3:5],16),int(color[5:7],16)
        fill_color = f"rgba({r},{g},{b},0.10)"
    except:
        fill_color = "rgba(31,154,154,0.10)"
    kwargs = dict(secondary_y=secondary_y) if secondary_y is not False else {}
    fig.add_trace(go.Scatter(
        x=df[date_col], y=df[col_name], name=label, mode="lines",
        line=dict(color=color, width=2.25),
        fill="tozeroy", fillcolor=fill_color, xhoverformat="%d %b %Y",
        hovertemplate=f"{label}: <b>%{{y:,.1f}}</b><extra></extra>",
    ), **kwargs)
    if len(df):
        last = df.sort_values(date_col).iloc[-1]
        fig.add_trace(go.Scatter(
            x=[last[date_col]], y=[last[col_name]], mode="markers",
            marker=dict(color=color, size=7, line=dict(color="#ffffff", width=1.5)),
            showlegend=False, hoverinfo="skip",
        ), **kwargs)

def show_chart(fig, **kwargs):
    """Display a Plotly chart; drops area fills when more than one line is filled."""
    filled = [t for t in fig.data if getattr(t, "fill", None) == "tozeroy"]
    if len(filled) > 1:
        for t in filled:
            t.fill = None
    kwargs.setdefault("use_container_width", True)
    st.plotly_chart(fig, **kwargs)

def std_layout(title, yaxis_title="", height=500):
    """Shared layout. Pass title="" when a section heading above already names the chart."""
    layout = dict(PLOTLY_LAYOUT)
    if title:
        layout["title"] = dict(text=title, font=dict(color=BI_TEXT, size=14, family=CHART_FONT, weight="bold"),
                               x=0, xanchor="left", y=0.98, yanchor="top", yref="container", pad=dict(l=4))
        layout["margin"] = dict(PLOTLY_LAYOUT["margin"], t=80)
    layout["yaxis"] = dict(PLOTLY_LAYOUT["yaxis"],
                           title=dict(text=yaxis_title, font=dict(color=BI_SUBTEXT, size=11)))
    layout["height"] = height
    return layout

def plotly_line(weekly_df, date_col, series, title, yaxis_title="Members", height=500):
    fig = go.Figure()
    for col_name, color, label in series:
        if col_name not in weekly_df.columns: continue
        std_traces(fig, weekly_df, date_col, col_name, color, label)
    fig.update_layout(**std_layout(title, yaxis_title, height))
    return fig

def apply_gpm_filter(df):
    """Filter data based on logged-in user's allowed locations.
    Only admins see every location; anyone else sees only their allowed
    locations (none if the list is empty or could not be loaded)."""
    if st.session_state.get("access_level") == "admin":
        return df
    allowed_locations = st.session_state.get("allowed_locations", [])
    wm_loc_col = "Success Tutoring - Business name"
    if wm_loc_col in df.columns:
        return df[df[wm_loc_col].isin(allowed_locations)]
    return df.iloc[0:0]
def get_13m_filtered(df_wm, df_filtered):
    loc_col = "Success Tutoring - Business name"
    max_date = df_wm["Date"].max()
    cutoff = max_date - pd.DateOffset(months=13)
    df_13m = df_wm[df_wm["Date"] >= cutoff].copy()
    if loc_col in df_filtered.columns and loc_col in df_13m.columns:
        locs = df_filtered[loc_col].unique()
        if len(locs) < len(df_wm[loc_col].unique()):
            df_13m = df_13m[df_13m[loc_col].isin(locs)]
    for col in ["Country","Region","Stage","GPM","Status"]:
        if col in df_filtered.columns and col in df_13m.columns:
            vals = df_filtered[col].unique()
            if len(vals) < len(df_wm[col].dropna().unique()):
                df_13m = df_13m[df_13m[col].isin(vals)]
    return df_13m

def report_filters(df, key_prefix="", show_date=True,
                   show_country=True, show_state=True,
                   show_stage=True, show_gpm=True, show_location=False, show_status=True, panel=None):
    """Filters in one labelled panel. Pass `panel` to add more controls to the same panel."""
    loc_col = "Success Tutoring - Business name"
    if panel is None:
        panel = st.container(border=True)
    with panel:
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        n_cols = 6 if show_location else 5
        cols = st.columns(n_cols); i = 0
        if show_date and "Date" in df.columns:
            max_d = df["Date"].max(); min_d = df["Date"].min()
            with cols[i%n_cols]:
                period = st.selectbox("Period",
                    ["All Time","Last Month","Last Quarter","Last 6 Months","Last Year","Latest Week","Custom"],
                    index=0, key=f"{key_prefix}_date")
            i += 1
            if period=="Latest Week": df=df[df["Date"]==max_d]
            elif period=="Last Month": df=df[df["Date"]>=max_d-pd.DateOffset(months=1)]
            elif period=="Last Quarter": df=df[df["Date"]>=max_d-pd.DateOffset(months=3)]
            elif period=="Last 6 Months": df=df[df["Date"]>=max_d-pd.DateOffset(months=6)]
            elif period=="Last Year": df=df[df["Date"]>=max_d-pd.DateOffset(years=1)]
            elif period=="Custom":
                with cols[i%n_cols]:
                    cr=st.date_input("Range",value=(min_d.date(),max_d.date()),
                                     min_value=min_d.date(),max_value=max_d.date(),
                                     key=f"{key_prefix}_custom")
                i+=1
                if isinstance(cr,tuple) and len(cr)==2:
                    df=df[(df["Date"].dt.date>=cr[0])&(df["Date"].dt.date<=cr[1])]
        if show_country and "Country" in df.columns:
            with cols[i%n_cols]:
                opts=["All"]+sorted(df["Country"].dropna().unique().tolist())
                sel=st.selectbox("Country",opts,key=f"{key_prefix}_country")
            if sel!="All": df=df[df["Country"]==sel]
            i+=1
        if show_state and "Region" in df.columns:
            with cols[i%n_cols]:
                opts=["All"]+sorted(df["Region"].dropna().unique().tolist())
                sel=st.selectbox("State",opts,key=f"{key_prefix}_state")
            if sel!="All": df=df[df["Region"]==sel]
            i+=1
        if show_stage and "Stage" in df.columns:
            with cols[i%n_cols]:
                opts=["All"]+sorted(df["Stage"].dropna().unique().tolist())
                sel=st.selectbox("Stage",opts,key=f"{key_prefix}_stage")
            if sel!="All": df=df[df["Stage"]==sel]
            i+=1
        if show_gpm and "GPM" in df.columns:
            with cols[i%n_cols]:
                opts=["All"]+sorted(df["GPM"].dropna().unique().tolist())
                sel=st.selectbox("GPM",opts,key=f"{key_prefix}_gpm")
            if sel!="All": df=df[df["GPM"]==sel]
            i+=1
        if show_location and loc_col in df.columns:
            with cols[i%n_cols]:
                opts=sorted(df[loc_col].dropna().unique().tolist())
                sel=st.multiselect("Location",opts,key=f"{key_prefix}_loc",placeholder="All locations",
                                   format_func=lambda x: x.replace("Success Tutoring - ",""))
            if sel: df=df[df[loc_col].isin(sel)]
            i+=1
        if show_status and "Status" in df.columns:
            with cols[i%n_cols]:
                opts=["All"]+sorted(df["Status"].dropna().unique().tolist())
                sel=st.selectbox("Status",opts,key=f"{key_prefix}_status")
            if sel!="All": df=df[df["Status"]==sel]
            i+=1
    return df

def checkbox_date_filter(df, key_prefix=""):
    if "Date" not in df.columns or df.empty: return df
    all_dates = sorted(df["Date"].dropna().unique(),reverse=True)
    date_labels = [d.strftime("%d %b %Y") for d in all_dates]
    st.markdown(f"""<div class="filter-label">Weeks</div>
        <p style="color:{BI_SUBTEXT};font-size:0.8em;margin:0 0 4px 0">Newest first. Default is the latest 2 weeks.</p>""",
        unsafe_allow_html=True)
    default_sel = date_labels[:2] if len(date_labels) >= 2 else date_labels
    selected_labels = st.multiselect(
        "Select one or more weeks:",
        options=date_labels, default=default_sel,
        key=f"{key_prefix}_multidate")
    if not selected_labels:
        st.warning("Please select at least one week — defaulting to latest.")
        selected_labels = [date_labels[0]]
    selected_dt = [pd.Timestamp(d) for d in selected_labels]
    return df[df["Date"].isin(selected_dt)]

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 1 — Locations
# ══════════════════════════════════════════════════════════════════════════════
def report_locations(df_wm):
    st.markdown('<div class="report-title">Campus Locations</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Vlookup — all location metadata with stage, GPM, country and region</div>', unsafe_allow_html=True)
    loc_col = "Success Tutoring - Business name"
    df_wm = apply_gpm_filter(df_wm)
    vl_full = apply_gpm_filter(load_vlookup()).copy()
    # Label blank metadata so gaps are easy to spot (and fix in the Vlookup sheet)
    for c in ["Country", "Stage", "Region", "GPM", "Status"]:
        if c in vl_full.columns:
            vl_full[c] = vl_full[c].astype(str).str.strip().replace({"": "Not set", "nan": "Not set", "None": "Not set"})
    vl = report_filters(vl_full.copy(),key_prefix="r1vl",show_date=False,
                        show_country=True,show_state=True,show_stage=True,show_gpm=True,show_status=True)
    if not vl.empty and "Stage" in vl.columns:
        sc = vl["Stage"].value_counts()
        leasing=int(sc.get("Leasing",0)); onboarding=int(sc.get("Onboarding",0))
        growth=int(sc.get("Growth",0)); total=leasing+onboarding+growth
        gauges=[(leasing,BI_RED,"Leasing"),(onboarding,BI_ACCENT,"Onboarding"),
                (growth,BI_ORANGE,"Growth"),(total,BI_BLUE,"Total Locations")]
        cols=st.columns(4)
        for i,(val,color,label) in enumerate(gauges):
            with cols[i]:
                fig=gauge_chart(val,total,color)
                st.pyplot(fig); plt.close()
                st.markdown(f'<div class="gauge-label">{label}</div>',unsafe_allow_html=True)
    all_dates = sorted(df_wm["Date"].dropna().unique())
    latest_date = all_dates[-1] if all_dates else None
    if not vl.empty and "Country" in vl.columns and "Stage" in vl.columns:
        st.markdown('<div class="section-header">Locations by Country & Stage</div>',unsafe_allow_html=True)
        pivot = vl.groupby(["Country","Stage"]).size().unstack(fill_value=0)
        pivot["Grand Total"]=pivot.sum(axis=1); pivot.loc["Grand Total"]=pivot.sum()
        st.dataframe(pivot,use_container_width=True)
    st.markdown('<div class="section-header">All Location Details</div>',unsafe_allow_html=True)
    if not vl.empty:
        if latest_date:
            lw_all = df_wm[df_wm["Date"]==latest_date][[loc_col,
                "# Active members","# New members","# Suspended members","# Cancelled members"]].copy()
            for c in ["# Active members","# New members","# Suspended members","# Cancelled members"]:
                lw_all[c]=pd.to_numeric(lw_all[c],errors="coerce").fillna(0)
            merged = vl.merge(lw_all,on=loc_col,how="left")
            for c in ["# Active members","# New members","# Suspended members","# Cancelled members"]:
                if c in merged.columns: merged[c]=merged[c].fillna(0)
            merged["Churn %"]=merged.apply(
                lambda r: churn_rate(r.get("# Cancelled members",0),r.get("# Active members",0)),axis=1)
            show_cols=[c for c in [loc_col,"Stage","GPM","Status","Country","Region",
                                   "Age (Months)","# Active members","# New members",
                                   "# Suspended members","# Cancelled members","Churn %"] if c in merged.columns]
            tbl = merged[show_cols].rename(columns={loc_col:"Location"}).sort_values(
                "# Active members",ascending=False).reset_index(drop=True)
            tbl["Location"] = tbl["Location"].astype(str).str.replace("Success Tutoring - ","",regex=False)
            st.dataframe(tbl, use_container_width=True, hide_index=True)
        else:
            show_cols=[c for c in [loc_col,"Stage","GPM","Status","Country","Region","Age (Months)"] if c in vl.columns]
            st.dataframe(vl[show_cols].rename(columns={loc_col:"Location"}).reset_index(drop=True),
                         use_container_width=True,hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 2 — Membership
# ══════════════════════════════════════════════════════════════════════════════
MEMBER_METRICS = {
    "# Active members":    "Active members",
    "# New members":       "New members",
    "# Suspended members": "Suspended members",
    "# Cancelled members": "Cancelled members",
}
# For these, a fall is good news
LOWER_IS_BETTER = {"# Suspended members", "# Cancelled members", "Churn %"}

def pct_card(label, value, note, higher_is_better=True):
    """KPI card whose headline is a % change, coloured by whether the change is good."""
    if value is None:
        text, color = "–", BI_SUBTEXT
    else:
        good = value >= 0 if higher_is_better else value <= 0
        text = f"{'+' if value >= 0 else ''}{value:.1f}%"
        color = BI_SUBTEXT if value == 0 else (BI_GREEN if good else BI_RED)
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value" style="color:{color}">{text}</div>
        <div class="metric-delta"><span>{note}</span></div>
    </div>""", unsafe_allow_html=True)

# ── Shared 52-week year-over-year bar chart (Membership, Net Growth, Revenue) ─
MUTED_YEAR_COLORS = ["#f0a88c", "#9fb0b5", "#c9d3d6", "#dfe6e8"]

def add_week_year(df):
    """Add ISO Week and Year columns (from 'Date - Week/Year' when present, otherwise from Date)."""
    d = df.copy()
    if "Date - Week/Year" in d.columns:
        wy = d["Date - Week/Year"].astype(str).str.split("/")
        d["Week"] = pd.to_numeric(wy.str[0], errors="coerce")
        d["Year"] = pd.to_numeric(wy.str[1], errors="coerce")
    else:
        iso = pd.to_datetime(d["Date"]).dt.isocalendar()
        d["Week"] = iso["week"].astype("float"); d["Year"] = iso["year"].astype("float")
    d = d.dropna(subset=["Week", "Year"])
    d["Week"] = d["Week"].astype(int); d["Year"] = d["Year"].astype(int)
    return d

def yoy_options(key, years, metrics=None, avg_label="Average per location", avg_help=None,
                show_avg=True, show_churn=False, avg_metrics=None):
    """avg_metrics: if given, the Average toggle only appears for these metrics (rates are already per-unit)."""
    """Options row above a year-over-year chart. Returns a dict of the choices."""
    cols = st.columns([2, 2, 2, 2]) if metrics else st.columns([2, 2, 2])
    i, out = 0, {}
    if metrics:
        out["metric"] = cols[0].selectbox("Metric", list(metrics), format_func=metrics.get, key=f"{key}_metric")
        i = 1
    out["years"] = cols[i].multiselect("Years", years, default=years[:2], key=f"{key}_years")
    out["holidays"] = cols[i + 1].multiselect("School holidays", ["NSW", "VIC", "QLD", "WA", "SA", "NZ"],
                                              default=[], key=f"{key}_holidays", placeholder="None")
    with cols[i + 2]:
        out["comparative"] = st.checkbox("Comparative only", value=False, key=f"{key}_comparative",
                                         help="Only include locations open in every selected year")
        if avg_metrics is not None and out.get("metric") not in avg_metrics:
            show_avg = False
        out["average"] = st.checkbox(avg_label, value=False, key=f"{key}_average",
                                     help=avg_help or "Totals ÷ number of locations trading that week") if show_avg else False
        out["churn"] = st.checkbox("Churn %", value=False, key=f"{key}_churn",
                                   help="Plot churn % (cancelled ÷ active) as lines on the right axis") if show_churn else False
    return out

def comparative_only(d, years, loc_col):
    """Keep only (location, week) pairs present in every selected year."""
    if len(years) < 2: return d
    pairs = [set(zip(d[d["Year"] == yr][loc_col], d[d["Year"] == yr]["Week"])) for yr in years]
    common = pairs[0].intersection(*pairs[1:])
    return d[[(l, w) in common for l, w in zip(d[loc_col], d["Week"])]]

def trading_locations(d, loc_col, active_col, year):
    """Number of locations with a positive value in active_col, per week of a year."""
    x = d[(d["Year"] == year) & (d[active_col] > 0)]
    return x.groupby("Week")[loc_col].nunique()

def yoy_bar_figure(weekly, latest_year, latest_week, title, y_title, fmt=",.0f", prefix="", suffix="",
                   change="pct", holidays=None, lines=None, line_axis_title="Churn %"):
    """Weekly bars per year: this year in teal, earlier years muted, current-week marker,
    month names under week numbers. `weekly` and `lines` map year -> Series indexed by week 1..52."""
    import datetime as _dt
    all_weeks = list(range(1, 53))
    years = sorted(weekly)
    prev = weekly.get(latest_year - 1)
    fig = go.Figure()
    for yr in years:
        vals = weekly[yr].reindex(all_weeks)
        if yr == latest_year:
            color = BI_ACCENT
            if prev is not None:
                pv = prev.reindex(all_weeks)
                if change == "pts":
                    note = [f"{'+' if v-p >= 0 else ''}{v-p:.2f} pts vs {latest_year-1}" if pd.notna(v) and pd.notna(p) else ""
                            for v, p in zip(vals, pv)]
                else:
                    note = [f"{'+' if v-p >= 0 else ''}{(v-p)/p*100:.0f}% vs {latest_year-1}" if pd.notna(v) and pd.notna(p) and p != 0 else ""
                            for v, p in zip(vals, pv)]
            else:
                note = [""] * 52
            hover = f"<b>{yr}</b>: {prefix}%{{y:{fmt}}}{suffix}  %{{customdata}}<extra></extra>"
        else:
            color = MUTED_YEAR_COLORS[(latest_year - yr - 1) % len(MUTED_YEAR_COLORS)]
            note = [""] * 52
            hover = f"{yr}: {prefix}%{{y:{fmt}}}{suffix}<extra></extra>"
        fig.add_trace(go.Bar(x=all_weeks, y=vals.values, name=str(yr), marker_color=color,
                             customdata=note, hovertemplate=hover))
    if lines:
        for yr in sorted(lines):
            ln = lines[yr].reindex(all_weeks)
            latest = yr == latest_year
            fig.add_trace(go.Scatter(x=all_weeks, y=ln.values, name=f"{yr} churn %", yaxis="y2", mode="lines",
                                     line=dict(color=BI_RED if latest else "#e8a09e", width=2.25 if latest else 1.75,
                                               dash="solid" if latest else "dot"),
                                     connectgaps=False, hovertemplate=f"{yr} churn: %{{y:.1f}}%<extra></extra>"))
    if holidays:
        df_hols = load_school_holidays()
        if not df_hols.empty:
            for _, h in df_hols[df_hols["State"].isin(holidays)].iterrows():
                h_year, h_start, h_end = int(h["Year"]), int(h["Start_Week"]), int(h["End_Week"])
                base = HOLIDAY_COLORS.get(h["Holiday"], "rgba(200,200,200,")
                if h_year == latest_year:
                    fig.add_vrect(x0=h_start-0.5, x1=h_end+0.5, fillcolor=base+"0.15)", layer="below", line_width=0,
                                  annotation_text=h["Holiday"], annotation_position="top left",
                                  annotation=dict(font_size=9, font_color=BI_SUBTEXT))
                elif h_year in years:
                    fig.add_vrect(x0=h_start-0.5, x1=h_end+0.5, fillcolor=base+"0.07)", layer="below", line_width=0)
    fig.add_vline(x=latest_week + 0.5, line_dash="dash", line_color=BI_SUBTEXT, line_width=1,
                  annotation_text=f"Wk {latest_week}", annotation_position="top left",
                  annotation_font=dict(size=10, color=BI_SUBTEXT))
    ticktext, last_month = [], None
    for w in all_weeks:
        try: m = _dt.date.fromisocalendar(int(latest_year), w, 4).strftime("%b")
        except ValueError: m = None
        ticktext.append(f"{w}<br><b>{m}</b>" if m and m != last_month else str(w))
        last_month = m or last_month
    layout = std_layout(title, y_title, 460)
    layout.update(barmode="group", bargap=0.18, bargroupgap=0.04, hovermode="x unified",
                  margin=dict(l=60, r=60 if lines else 30, t=70, b=60),
                  legend=dict(PLOTLY_LAYOUT["legend"], orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    layout["xaxis"] = dict(PLOTLY_LAYOUT["xaxis"], tickmode="array", tickvals=all_weeks, ticktext=ticktext,
                           tickangle=0, range=[0.4, 52.6], title=dict(text="Week", font=dict(color=BI_SUBTEXT, size=11)))
    layout["yaxis"] = dict(layout["yaxis"], tickformat=",.2~f", tickprefix=prefix, ticksuffix=suffix,
                           rangemode="normal" if change == "pts" else "tozero")
    if lines:
        layout["yaxis2"] = dict(overlaying="y", side="right", showgrid=False, zeroline=False, rangemode="tozero",
                                ticksuffix="%", tickformat=".1f", tickfont=dict(color=BI_RED, size=11),
                                title=dict(text=line_axis_title, font=dict(color=BI_RED, size=11)))
    fig.update_layout(**layout)
    return fig

def location_comparison(d13, loc_col, measures, key, rank_col, start_col):
    """Clean lines per location for one measure, labelled at the line ends, with an optional
    network-average line. measures: label -> dict(kind="sum", col=...) or
    dict(kind="ratio", num=..., den=..., scale=1), plus optional prefix/suffix/fmt."""
    max_date = d13["Date"].max()
    ranked = (d13[d13["Date"] == max_date].groupby(loc_col)[rank_col].sum().sort_values(ascending=False).index.tolist())
    all_locs = ranked + sorted(set(d13[loc_col].dropna()) - set(ranked))
    c1, c2, c3 = st.columns([2, 5, 1.6])
    label = c1.selectbox("Measure", list(measures), key=f"{key}_measure")
    sel_locs = c2.multiselect("Locations (up to 8)", all_locs, default=all_locs[:3], max_selections=8,
                              key=f"{key}_locs", format_func=lambda x: x.replace("Success Tutoring - ", ""))
    with c3:
        st.markdown("<div style='height:1.9rem'></div>", unsafe_allow_html=True)
        show_avg = st.checkbox("Network average", value=True, key=f"{key}_avg")
    m = measures[label]
    need = [c for c in {m.get("col"), m.get("num"), m.get("den"), start_col} if c]
    fmt, pre, suf = m.get("fmt", ",.0f"), m.get("prefix", ""), m.get("suffix", "")

    def series(dd):
        g = dd.groupby("Date")[need].sum().sort_index()
        g = g[g[start_col].cumsum() > 0]
        if m["kind"] == "sum":
            return g[m["col"]]
        return (g[m["num"]] / g[m["den"]] * m.get("scale", 1)).where(g[m["den"]] > 0)

    if not sel_locs:
        st.info("Select one or more locations to compare."); return
    fig = go.Figure(); ends = []
    if show_avg:
        trading = d13[d13[start_col] > 0]
        if m["kind"] == "sum":
            avg = trading.groupby(["Date", loc_col])[m["col"]].sum().groupby(level="Date").mean()
        else:
            avg = series(trading)
        fig.add_trace(go.Scatter(x=avg.index, y=avg.values, name="Network avg", mode="lines",
                                 line=dict(color="#9fb0b5", width=2, dash="dash"),
                                 hovertemplate=f"Network avg: {pre}%{{y:{fmt}}}{suf}<extra></extra>"))
    for i, loc in enumerate(sel_locs):
        sr = series(d13[d13[loc_col] == loc]).dropna()
        if sr.empty: continue
        name = loc.replace("Success Tutoring - ", "")
        color = SERIES_COLORS[i % len(SERIES_COLORS)]
        fig.add_trace(go.Scatter(x=sr.index, y=sr.values, name=name, mode="lines", line=dict(color=color, width=2.5),
                                 hovertemplate=f"{name}: {pre}%{{y:{fmt}}}{suf}<extra></extra>"))
        fig.add_trace(go.Scatter(x=[sr.index[-1]], y=[sr.values[-1]], mode="markers", marker=dict(color=color, size=8),
                                 showlegend=False, hoverinfo="skip"))
        ends.append([sr.index[-1], float(sr.values[-1]), f"{name} {pre}{sr.values[-1]:{fmt}}{suf}", color])
    annotations = []
    if ends:
        gap = (max(abs(e[1]) for e in ends) or 1) * 0.05
        ends.sort(key=lambda e: e[1]); placed = []
        for e in ends:
            placed.append(e[1] if not placed or e[1] - placed[-1] >= gap else placed[-1] + gap)
        annotations = [dict(x=e[0], y=y, xref="x", yref="y", text=f"<b>{e[2]}</b>", showarrow=False,
                            xanchor="left", xshift=10, font=dict(color=e[3], size=11)) for e, y in zip(ends, placed)]
    layout = std_layout("", label, 460)
    layout.update(showlegend=False, hovermode="x unified", margin=dict(l=70, r=190, t=20, b=50), annotations=annotations)
    layout["xaxis"] = dict(PLOTLY_LAYOUT["xaxis"], tickangle=0, tickformat="%b %y", dtick="M1",
                           range=[d13["Date"].min() - pd.Timedelta(days=3), max_date + pd.Timedelta(days=3)])
    layout["yaxis"] = dict(layout["yaxis"], tickformat=",.2~f", tickprefix=pre, ticksuffix=suf)
    fig.update_layout(**layout)
    show_chart(fig, use_container_width=True)

def report_membership(df_wm, df_rv):
    st.markdown('<div class="report-title">Membership</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Weekly Membership — active, new, suspended & cancelled member trends</div>', unsafe_allow_html=True)
    loc_col = "Success Tutoring - Business name"
    num_cols = list(MEMBER_METRICS)
    df_wm = apply_gpm_filter(df_wm)

    # ── One filter panel drives the whole report ──────────────────────────
    panel = st.container(border=True)
    df_f = report_filters(df_wm.copy(), key_prefix="r2", show_date=False,
                          show_country=True, show_state=True, show_stage=True, show_gpm=True,
                          show_location=True, show_status=True, panel=panel)
    for c in num_cols:
        if c in df_f.columns: df_f[c] = pd.to_numeric(df_f[c], errors="coerce").fillna(0)
    with panel:
        df = checkbox_date_filter(df_f, key_prefix="r2")
    if df_f.empty:
        st.warning("No data for the selected filters."); return

    # ── KPI cards ─────────────────────────────────────────────────────────
    all_dates = sorted(df["Date"].dropna().unique())
    latest_date = all_dates[-1] if all_dates else None
    prev_date = all_dates[-2] if len(all_dates) >= 2 else None
    def ws(d, col):
        if d is None or col not in df.columns: return 0
        return df[df["Date"] == d][col].sum()
    la = ws(latest_date, "# Active members");    pa = ws(prev_date, "# Active members")
    ln = ws(latest_date, "# New members");       pn = ws(prev_date, "# New members")
    ls = ws(latest_date, "# Suspended members"); ps = ws(prev_date, "# Suspended members")
    lc = ws(latest_date, "# Cancelled members"); pc = ws(prev_date, "# Cancelled members")
    cr_latest = churn_rate(lc, la); cr_prev = churn_rate(pc, pa)
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1: metric_card("Active Members",    f"{la:,.0f}", la - pa, "green")
    with k2: metric_card("New Members",       f"{ln:,.0f}", ln - pn, "blue")
    with k3: metric_card("Suspended Members", f"{ls:,.0f}", ls - ps, "orange", higher_is_better=False)
    with k4: metric_card("Cancelled Members", f"{lc:,.0f}", lc - pc, "red", higher_is_better=False)
    with k5: metric_card("Network Churn Rate", f"{cr_latest:.1f}%", round(cr_latest - cr_prev, 2), "red", higher_is_better=False)

    # ── Year-over-year weekly comparison ──────────────────────────────────
    st.markdown('<div class="section-header">Year-over-Year Weekly Comparison</div>', unsafe_allow_html=True)
    df_yoy = add_week_year(df_f)
    available_years = sorted(df_yoy["Year"].unique().tolist(), reverse=True)
    if not available_years:
        st.info("No week/year data available for year-over-year comparison.")
    else:
        opts = yoy_options("yoy", available_years, metrics=MEMBER_METRICS, show_churn=True)
        yoy_metric, sel_years_yoy = opts["metric"], opts["years"]
        metric_name = MEMBER_METRICS[yoy_metric]
        higher_better = yoy_metric not in LOWER_IS_BETTER

        df_yoy_f = df_yoy[df_yoy["Year"].isin(sel_years_yoy)].copy()
        if opts["comparative"]:
            df_yoy_f = comparative_only(df_yoy_f, sel_years_yoy, loc_col)

        if not sel_years_yoy or df_yoy_f.empty:
            st.info("Select at least one year with data.")
        else:
            latest_year = max(sel_years_yoy)
            latest_week = int(df_yoy_f[df_yoy_f["Year"]==latest_year]["Week"].max())
            prev_year   = latest_year - 1
            weekly, churn_lines = {}, {}
            for yr in sel_years_yoy:
                d = df_yoy_f[df_yoy_f["Year"] == yr]
                g = d.groupby("Week")[num_cols].sum()
                vals = g[yoy_metric]
                if opts["average"]:
                    vals = vals / trading_locations(df_yoy_f, loc_col, "# Active members", yr)
                weekly[yr] = vals
                if opts["churn"]:
                    churn_lines[yr] = (g["# Cancelled members"] / g["# Active members"] * 100).where(g["# Active members"] > 0)
            label = f"Avg {metric_name.lower()} per location" if opts["average"] else metric_name
            fig_yoy = yoy_bar_figure(weekly, latest_year, latest_week, f"{label} by week", label,
                                     fmt=",.1f" if opts["average"] else ",.0f",
                                     holidays=opts["holidays"], lines=churn_lines or None)
            show_chart(fig_yoy, use_container_width=True)

            # ── Variance cards ────────────────────────────────────────────
            def pct(a, b):
                if a is None or b is None or pd.isna(a) or pd.isna(b) or b == 0: return None
                return round((a - b) / b * 100, 1)
            def total(year, week=None, upto=None):
                d = df_yoy_f[df_yoy_f["Year"] == year]
                if week is not None: d = d[d["Week"] == week]
                if upto is not None: d = d[d["Week"] <= upto]
                v = d[yoy_metric].sum()
                return v if v > 0 else None
            cur = total(latest_year, latest_week)
            v1, v2, v3 = st.columns(3)
            with v1: pct_card("vs last week", pct(cur, total(latest_year, latest_week - 1)),
                              f"Wk {latest_week} vs Wk {latest_week - 1}, {latest_year}", higher_better)
            with v2: pct_card("vs same week last year", pct(cur, total(prev_year, latest_week)),
                              f"Wk {latest_week}, {latest_year} vs {prev_year}", higher_better)
            with v3: pct_card("Year to date vs last year",
                              pct(total(latest_year, upto=latest_week), total(prev_year, upto=latest_week)),
                              f"Wk 1–{latest_week}, {latest_year} vs {prev_year}", higher_better)

            # ── Location table ────────────────────────────────────────────
            st.markdown(f'<div class="section-header">Location Detail — {metric_name}, Wk {latest_week} {latest_year}</div>', unsafe_allow_html=True)
            c_cur, c_prev_wk, c_vs_wk = "This Week", "Last Week", "vs LW %"
            c_prev_yr, c_vs_yr        = "Same Week LY", "vs Same Week LY %"
            c_ytd, c_ytd_prev, c_ytd_var = "YTD TY", "YTD LY", "YTD vs LY %"
            pct_cols = [c_vs_wk, c_vs_yr, c_ytd_var]

            def row_for(d, name):
                def s(year, week=None, upto=None):
                    x = d[d["Year"] == year]
                    if week is not None: x = x[x["Week"] == week]
                    if upto is not None: x = x[x["Week"] <= upto]
                    v = x[yoy_metric].sum()
                    return v if v > 0 else None
                cur_v, pw, py = s(latest_year, latest_week), s(latest_year, latest_week - 1), s(prev_year, latest_week)
                ytd, ytdp = s(latest_year, upto=latest_week), s(prev_year, upto=latest_week)
                return {"Location": name, c_cur: cur_v, c_prev_wk: pw, c_vs_wk: pct(cur_v, pw),
                        c_prev_yr: py, c_vs_yr: pct(cur_v, py),
                        c_ytd: ytd, c_ytd_prev: ytdp, c_ytd_var: pct(ytd, ytdp)}

            rows = [row_for(g, loc.replace("Success Tutoring - ", ""))
                    for loc, g in df_yoy_f.groupby(loc_col)]
            df_table = pd.DataFrame(rows).sort_values(c_cur, ascending=False, na_position="last")
            df_table = pd.concat([df_table, pd.DataFrame([row_for(df_yoy_f, "TOTAL")])], ignore_index=True)
            if prev_year not in sel_years_yoy:
                df_table = df_table.drop(columns=[c_prev_yr, c_vs_yr, c_ytd_prev, c_ytd_var])
                pct_cols = [c_vs_wk]

            def fmt_pct_cell(v):
                return "–" if v is None or pd.isna(v) else f"{'+' if v >= 0 else ''}{v:.1f}%"
            for col in pct_cols:
                df_table[col] = df_table[col].apply(fmt_pct_cell)

            good_col, bad_col = BI_GREEN, BI_RED
            if not higher_better: good_col, bad_col = bad_col, good_col
            def style_table(d):
                styles = pd.DataFrame("", index=d.index, columns=d.columns)
                for col in pct_cols:
                    for i, val in enumerate(d[col]):
                        if val.startswith("+") and val != "+0.0%": styles.iloc[i, d.columns.get_loc(col)] = f"color: {good_col}; font-weight: 600"
                        elif val.startswith("-"):                  styles.iloc[i, d.columns.get_loc(col)] = f"color: {bad_col}; font-weight: 600"
                styles.iloc[-1] = styles.iloc[-1] + "; font-weight: 700; background-color: #eef3f4"
                return styles
            count_cols = [c for c in df_table.columns if c not in ["Location"] + pct_cols]
            styled = (df_table.style.apply(style_table, axis=None)
                      .format("{:,.0f}", subset=[c for c in count_cols if c != c_cur], na_rep="–"))
            bar_max = float(df_table[c_cur].iloc[:-1].max() or 1) if len(df_table) > 1 else 1.0
            st.dataframe(styled, use_container_width=True, hide_index=True, height=min(500, 35 * (len(df_table) + 1) + 3),
                         column_config={
                             "Location": st.column_config.TextColumn("Location", pinned=True, width=180),
                             c_cur: st.column_config.ProgressColumn(c_cur, format="localized", min_value=0, max_value=bar_max),
                         })

    # ── Location comparison ───────────────────────────────────────────────
    st.markdown('<div class="section-header">Location Comparison — Last 13 Months</div>', unsafe_allow_html=True)
    max_date = df_f["Date"].max()
    df_13 = df_f[df_f["Date"] >= max_date - pd.DateOffset(months=13)].copy()
    locs_by_active = (df_13[df_13["Date"] == max_date].groupby(loc_col)["# Active members"].sum()
                      .sort_values(ascending=False).index.tolist())
    all_locs = locs_by_active + sorted(set(df_13[loc_col].dropna()) - set(locs_by_active))
    lc1, lc2, lc3 = st.columns([2, 5, 1.6])
    measure = lc1.selectbox("Measure", ["Active", "New", "Suspended", "Cancelled", "Churn %"], key="r2_cmp_measure")
    sel_locs = lc2.multiselect("Locations (up to 8)", all_locs, default=all_locs[:3], max_selections=8,
                               key="r2_cmp_locs", format_func=lambda x: x.replace("Success Tutoring - ", ""))
    with lc3:
        st.markdown("<div style='height:1.9rem'></div>", unsafe_allow_html=True)
        show_avg = st.checkbox("Network average", value=True, key="r2_cmp_avg")

    def series(d):
        """Weekly values for one location (or the whole network), starting from its first active week."""
        g = d.groupby("Date")[num_cols].sum().sort_index()
        g = g[g["# Active members"].cumsum() > 0]
        if measure == "Churn %":
            return g.apply(lambda r: churn_rate(r["# Cancelled members"], r["# Active members"]), axis=1)
        return g[f"# {measure} members"]

    if sel_locs:
        fig_c = go.Figure()
        end_labels = []
        fmt = ".1f" if measure == "Churn %" else ",.0f"
        suffix = "%" if measure == "Churn %" else ""
        if show_avg:
            open_locs = df_13[df_13["# Active members"] > 0]
            if measure == "Churn %":
                avg = series(open_locs)
            else:
                avg = open_locs.groupby("Date").apply(lambda x: x.groupby(loc_col)[f"# {measure} members"].sum().mean())
            fig_c.add_trace(go.Scatter(x=avg.index, y=avg.values, name="Network avg", mode="lines",
                                       line=dict(color="#9fb0b5", width=2, dash="dash"),
                                       hovertemplate=f"Network avg: %{{y:{fmt}}}{suffix}<extra></extra>"))
        for i, loc in enumerate(sel_locs):
            sr = series(df_13[df_13[loc_col] == loc])
            if sr.empty: continue
            name = loc.replace("Success Tutoring - ", "")
            color = SERIES_COLORS[i % len(SERIES_COLORS)]
            fig_c.add_trace(go.Scatter(x=sr.index, y=sr.values, name=name, mode="lines",
                                       line=dict(color=color, width=2.5),
                                       hovertemplate=f"{name}: %{{y:{fmt}}}{suffix}<extra></extra>"))
            fig_c.add_trace(go.Scatter(x=[sr.index[-1]], y=[sr.values[-1]], mode="markers",
                                       marker=dict(color=color, size=8), showlegend=False, hoverinfo="skip"))
            end_labels.append([sr.index[-1], float(sr.values[-1]), f"{name} {sr.values[-1]:{fmt}}{suffix}", color])
        # Spread labels at least ~5% of the axis apart, keeping their order
        if end_labels:
            y_max = max([l[1] for l in end_labels] + [0]) or 1
            gap = y_max * 0.05
            end_labels.sort(key=lambda l: l[1])
            placed = []
            for l in end_labels:
                y = l[1] if not placed or l[1] - placed[-1] >= gap else placed[-1] + gap
                placed.append(y)
            annotations = [dict(x=l[0], y=y, xref="x", yref="y", text=f"<b>{l[2]}</b>", showarrow=False,
                                xanchor="left", xshift=10, font=dict(color=l[3], size=11))
                           for l, y in zip(end_labels, placed)]
        else:
            annotations = []
        layout = std_layout("", f"{measure}" + (" (per location)" if measure != "Churn %" else ""), 460)
        layout.update(showlegend=False, hovermode="x unified", margin=dict(l=60, r=170, t=20, b=50),
                      annotations=annotations)
        layout["xaxis"] = dict(PLOTLY_LAYOUT["xaxis"], tickangle=0, tickformat="%b %y")
        layout["yaxis"] = dict(layout["yaxis"], tickformat=",", ticksuffix=suffix, rangemode="tozero")
        fig_c.update_layout(**layout)
        fig_c.update_xaxes(dtick="M1")
        show_chart(fig_c, use_container_width=True)
    else:
        st.info("Select one or more locations to compare.")

    # ── Member count by location (latest selected week) ───────────────────
    if loc_col in df.columns and latest_date is not None:
        st.markdown(f'<div class="section-header">Member Count by Location — {pd.Timestamp(latest_date).strftime("%d %b %Y")}</div>',
                    unsafe_allow_html=True)
        loc_tbl = df[df["Date"] == latest_date].groupby(loc_col).agg(
            Active=("# Active members", "sum"), New=("# New members", "sum"),
            Suspended=("# Suspended members", "sum"), Cancelled=("# Cancelled members", "sum"),
        ).reset_index()
        loc_tbl["Location"] = loc_tbl[loc_col].str.replace("Success Tutoring - ", "", regex=False)
        loc_tbl["Churn %"] = loc_tbl.apply(lambda r: churn_rate(r["Cancelled"], r["Active"]), axis=1)
        loc_tbl = loc_tbl[["Location", "Active", "New", "Suspended", "Cancelled", "Churn %"]]
        loc_tbl = loc_tbl.sort_values("Active", ascending=False).reset_index(drop=True)
        st.dataframe(loc_tbl, use_container_width=True, hide_index=True,
                     column_config={
                         "Active": st.column_config.ProgressColumn("Active", format="localized", min_value=0,
                                                                   max_value=float(loc_tbl["Active"].max() or 1)),
                         "Churn %": st.column_config.NumberColumn("Churn %", format="%.1f%%"),
                     })

    weekly_location_metrics(df_f, df_rv)
    centre_performance_index(df_f)

def weekly_location_metrics(df_f, df_rv):
    """One row per week for a single location: members and revenue side by side."""
    loc_col = "Success Tutoring - Business name"
    short = lambda x: str(x).replace("Success Tutoring - ", "")
    st.markdown('<div class="section-header">Weekly Metrics by Location</div>', unsafe_allow_html=True)
    locs = sorted(df_f[loc_col].dropna().unique().tolist())
    if not locs:
        st.info("No locations match the selected filters."); return
    c1, c2 = st.columns([3, 1])
    loc = c1.selectbox("Location (type to search)", locs, format_func=short, key="r2_wm_loc")
    period = c2.selectbox("Period", ["Last month", "Last quarter", "Last year", "All weeks"], index=1, key="r2_wm_period")

    vl = load_vlookup()
    meta = vl[vl[loc_col] == loc].iloc[0] if not vl.empty and loc_col in vl.columns and (vl[loc_col] == loc).any() else None
    if meta is not None:
        bits = [f"{k}: <b>{meta.get(k)}</b>" for k in ["Stage", "GPM", "Region", "Country"]
                if str(meta.get(k, "")).strip() not in ("", "nan", "None")]
        try: bits.append(f"Open <b>{int(float(meta.get('Age (Months)', 0)))}</b> months")
        except (TypeError, ValueError): pass
        st.markdown(f'<div class="report-subtitle">{" · ".join(bits)}</div>', unsafe_allow_html=True)

    m = df_f[df_f[loc_col] == loc].groupby("Date")[["# Active members", "# New members",
                                                    "# Suspended members", "# Cancelled members"]].sum().sort_index()
    if m.empty:
        st.info("No weekly data for this location."); return
    cutoff = {"Last month": pd.DateOffset(months=1), "Last quarter": pd.DateOffset(months=3),
              "Last year": pd.DateOffset(years=1)}.get(period)
    if cutoff is not None:
        m = m[m.index >= m.index.max() - cutoff]
    m = m.reset_index()

    # Revenue for the same location, matched to the nearest membership week (within 7 days)
    rv_cols = ["Net Revenue", "# Active Students", "Total Sessions"]
    rv = apply_gpm_filter(df_rv) if not df_rv.empty else df_rv
    if not rv.empty and loc_col in rv.columns and all(c in rv.columns for c in rv_cols):
        r = rv[rv[loc_col] == loc].groupby("Date")[rv_cols].sum().reset_index().sort_values("Date")
        m = pd.merge_asof(m.sort_values("Date"), r, on="Date", direction="nearest", tolerance=pd.Timedelta(days=7))
    for c in rv_cols:
        if c not in m.columns: m[c] = float("nan")

    a, n, s_, c = m["# Active members"], m["# New members"], m["# Suspended members"], m["# Cancelled members"]
    rev, stud, sess = m["Net Revenue"], m["# Active Students"], m["Total Sessions"]
    out = pd.DataFrame({
        "Week": m["Date"].dt.strftime("%d %b %Y"),
        "Active": a, "New": n, "Suspended": s_, "Cancelled": c,
        "Churn %": (c / a * 100).where(a > 0),
        "Net Growth %": ((n - c) / a * 100).where(a > 0),
        "Net Revenue": rev, "Active Students": stud, "Sessions": sess,
        "Rev / Session": (rev / sess).where(sess > 0),
        "Rev / Student": (rev / stud).where(stud > 0),
        "Sessions / Student": (sess / stud).where(stud > 0),
    }).iloc[::-1].reset_index(drop=True)
    avg = out.drop(columns="Week").mean(numeric_only=True)
    avg["Week"] = f"Average ({period.lower()})"
    out = pd.concat([out, pd.DataFrame([avg])], ignore_index=True)

    fmt = {"Active": "{:,.0f}", "New": "{:,.1f}", "Suspended": "{:,.1f}", "Cancelled": "{:,.1f}",
           "Churn %": "{:.1f}%", "Net Growth %": "{:+.2f}%", "Net Revenue": "${:,.0f}",
           "Active Students": "{:,.0f}", "Sessions": "{:,.0f}", "Rev / Session": "${:,.2f}",
           "Rev / Student": "${:,.2f}", "Sessions / Student": "{:.2f}"}
    whole = {"New": "{:,.0f}", "Suspended": "{:,.0f}", "Cancelled": "{:,.0f}"}
    last = len(out) - 1
    def bold_avg(row):
        return ["font-weight: 700; background-color: #eef3f4" if row.name == last else "" for _ in row]
    styled = out.style.apply(bold_avg, axis=1).format(fmt, na_rep="–")
    styled = styled.format(whole, subset=pd.IndexSlice[out.index[:-1], list(whole)], na_rep="–")
    st.dataframe(styled, use_container_width=True, hide_index=True, height=min(560, 35 * (len(out) + 1) + 3),
                 column_config={"Week": st.column_config.TextColumn("Week", pinned=True, width=150)})

def centre_performance_index(df_f):
    """Each centre's active members divided by the average of all trading centres
    (or of trading centres in the same state), for every week of the chosen year.
    Rows follow the report filters; averages always use the whole network."""
    loc_col = "Success Tutoring - Business name"
    st.markdown('<div class="section-header">Centre Performance Index</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="report-subtitle">Active members ÷ average active members per trading centre. '
                f'1.00 = average · 1.28 = 28% above average · 0.80 = 20% below.</div>', unsafe_allow_html=True)

    full = load_weekly_membership().copy()
    if "Date - Week/Year" not in full.columns or loc_col not in full.columns:
        st.info("Week/Year column not found — index unavailable."); return
    wy = full["Date - Week/Year"].astype(str).str.split("/")
    full["Week"] = pd.to_numeric(wy.str[0], errors="coerce")
    full["Year"] = pd.to_numeric(wy.str[1], errors="coerce")
    full = full.dropna(subset=["Week", "Year"])
    full["Week"] = full["Week"].astype(int); full["Year"] = full["Year"].astype(int)
    full["# Active members"] = pd.to_numeric(full["# Active members"], errors="coerce").fillna(0)
    country = full["Country"] if "Country" in full.columns else pd.Series("", index=full.index)
    region = full["Region"] if "Region" in full.columns else pd.Series("", index=full.index)
    full["State"] = [("NZ" if c == "New Zealand" else REGION_TO_STATE.get(r, r)) for c, r in zip(country, region)]

    years = sorted(full["Year"].unique().tolist(), reverse=True)
    i1, i2, _ = st.columns([1, 2, 3])
    year = i1.selectbox("Year", years, key="r2_cpi_year")
    basis = i2.radio("Compare against", ["All centres", "State"], horizontal=True, key="r2_cpi_basis")

    wk = (full[full["Year"] == year].groupby([loc_col, "State", "Week"])["# Active members"].sum().reset_index())
    wk = wk[wk["# Active members"] > 0]          # only centres trading that week
    if wk.empty:
        st.info("No active members recorded for that year."); return
    group = ["Week"] if basis == "All centres" else ["State", "Week"]
    wk["Average"] = wk.groupby(group)["# Active members"].transform("mean")
    wk["Index"] = wk["# Active members"] / wk["Average"]

    # State totals: the state's average per trading centre against the same basis (whole network)
    state_wk = wk.groupby(["State", "Week"]).agg(StateMean=("# Active members", "mean"),
                                                 Basis=("Average", "first")).reset_index()
    state_wk["Index"] = state_wk["StateMean"] / state_wk["Basis"]
    state_totals = state_wk.pivot_table(index="State", columns="Week", values="Index")

    shown = set(df_f[loc_col].dropna().unique())
    wk = wk[wk[loc_col].isin(shown)]
    if wk.empty:
        st.info("No trading centres match the selected filters for that year."); return
    centres = wk.pivot_table(index=[loc_col, "State"], columns="Week", values="Index")
    weeks = sorted(set(centres.columns) | set(state_totals.columns))
    centres = centres.reindex(columns=weeks); state_totals = state_totals.reindex(columns=weeks)
    latest = max(weeks)

    # States A–Z; centres within a state highest index first; state total after its centres
    rows, total_rows = [], []
    for state in sorted(centres.index.get_level_values("State").unique()):
        grp = centres.xs(state, level="State").sort_values(latest, ascending=False, na_position="last")
        for loc, vals in grp.iterrows():
            rows.append({"State": state, "Location": loc.replace("Success Tutoring - ", ""), **vals.to_dict()})
        total_rows.append(len(rows))
        rows.append({"State": state, "Location": f"{state} total", **state_totals.loc[state].to_dict()})
    table = pd.DataFrame(rows)
    table = table.rename(columns={w: f"Wk {w}" for w in weeks})

    def shade(v):
        if pd.isna(v): return ""
        strength = min(abs(v - 1) / 0.5, 1) * 0.55
        rgb = "46,133,64" if v >= 1 else "200,50,47"
        return f"background-color: rgba({rgb},{strength:.2f})"
    def bold_totals(row):
        style = "font-weight: 700; border-top: 1px solid #b9c7cb"
        return [style if row.name in total_rows else "" for _ in row]
    week_cols = [c for c in table.columns if c not in ("State", "Location")]
    styled = (table.style.map(shade, subset=week_cols).apply(bold_totals, axis=1)
              .format("{:.2f}", subset=week_cols, na_rep="–"))
    st.dataframe(styled, use_container_width=True, hide_index=True,
                 height=min(700, 35 * (len(table) + 1) + 3),
                 column_config={"State": st.column_config.TextColumn("State", pinned=True, width=70),
                                "Location": st.column_config.TextColumn("Location", pinned=True, width=170)})
    avg_note = "all trading centres" if basis == "All centres" else "trading centres in the same state (all of NZ counts as one)"
    st.caption(f"Averages use {avg_note} across the whole network, whatever filters are applied. "
               f"State totals compare the state's average per trading centre with that average"
               f"{' (so they are always 1.00 in State mode)' if basis == 'State' else ''}. "
               f"Weeks where a centre had no active members show –.")

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 3 — Membership by Age
# ══════════════════════════════════════════════════════════════════════════════
def report_age_combined(df_wm):
    st.markdown('<div class="report-title">Membership by Age</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Age groups from Vlookup | Member counts from Weekly Membership</div>', unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)
    loc_col="Success Tutoring - Business name"
    vl=apply_gpm_filter(load_vlookup())
    if vl.empty or "Age (Months)" not in vl.columns:
        st.warning("Age (Months) not found in Vlookup."); return
    df_vl=report_filters(vl.copy(),key_prefix="r3vl",show_date=False,
                          show_country=True,show_state=True,show_stage=True,show_gpm=True,show_status=True)
    all_dates=sorted(df_wm["Date"].dropna().unique())
    latest_date=all_dates[-1] if all_dates else None
    if latest_date is None: st.warning("No date data."); return
    latest_members=df_wm[df_wm["Date"]==latest_date].copy()
    for c in ["# Active members","# New members","# Suspended members","# Cancelled members"]:
        if c in latest_members.columns: latest_members[c]=pd.to_numeric(latest_members[c],errors="coerce").fillna(0)
    merged=df_vl[[loc_col,"Age (Months)","Stage","GPM","Country","Region"]].merge(
        latest_members[[loc_col,"# Active members","# New members","# Suspended members","# Cancelled members"]],
        on=loc_col,how="left")
    for c in ["# Active members","# New members","# Suspended members","# Cancelled members"]:
        merged[c]=merged[c].fillna(0)
    merged["Age Group"]=merged["Age (Months)"].apply(get_age_group)
    summary_rows=[]; all_loc_data={}
    for grp in AGE_ORDER:
        grp_df=merged[merged["Age Group"]==grp]
        if grp_df.empty: continue
        summary_rows.append({
            "Age Group":grp,"# Locations":len(grp_df),
            "Avg Active":round(grp_df["# Active members"].mean(),1),
            "Avg New":round(grp_df["# New members"].mean(),1),
            "Avg Suspended":round(grp_df["# Suspended members"].mean(),1),
            "Avg Cancelled":round(grp_df["# Cancelled members"].mean(),1),
            "Avg Churn %":round(churn_rate(grp_df["# Cancelled members"].sum(),grp_df["# Active members"].sum()),1)
        })
        all_loc_data[grp]=grp_df
    if not merged.empty:
        summary_rows.append({
            "Age Group":"── TOTAL / AVERAGE ──","# Locations":len(merged),
            "Avg Active":round(merged["# Active members"].mean(),1),
            "Avg New":round(merged["# New members"].mean(),1),
            "Avg Suspended":round(merged["# Suspended members"].mean(),1),
            "Avg Cancelled":round(merged["# Cancelled members"].mean(),1),
            "Avg Churn %":round(churn_rate(merged["# Cancelled members"].sum(),merged["# Active members"].sum()),1)
        })
    st.dataframe(pd.DataFrame(summary_rows),use_container_width=True,hide_index=True)
    for grp in AGE_ORDER:
        if grp not in all_loc_data: continue
        grp_df=all_loc_data[grp]
        with st.expander(f"{grp} — {len(grp_df)} locations",expanded=False):
            show_cols=[c for c in [loc_col,"# Active members","# New members","# Suspended members","# Cancelled members","Age (Months)","Stage"] if c in grp_df.columns]
            st.dataframe(grp_df[show_cols].rename(columns={loc_col:"Location"}).sort_values("# Active members",ascending=False).reset_index(drop=True),
                         use_container_width=True,hide_index=True)
    st.markdown("<br>",unsafe_allow_html=True)
    st.markdown('<div class="section-header">Age Group Trends — Last 13 Months</div>',unsafe_allow_html=True)
    df_wm_f=report_filters(df_wm.copy(),key_prefix="r3wm",show_date=False,
                            show_country=True,show_state=True,show_stage=False,show_gpm=True,show_status=True)
    max_date=df_wm_f["Date"].max(); cutoff=max_date-pd.DateOffset(months=13)
    df_chart=df_wm_f[df_wm_f["Date"]>=cutoff].copy()
    age_map=vl.set_index(loc_col)["Age (Months)"].to_dict() if loc_col in vl.columns else {}
    df_chart["Age (Months)"]=df_chart[loc_col].map(age_map)
    df_chart["Age Group"]=df_chart["Age (Months)"].apply(get_age_group)
    for c in ["# Active members","# Suspended members","# Cancelled members"]:
        if c in df_chart.columns: df_chart[c]=pd.to_numeric(df_chart[c],errors="coerce").fillna(0)
    sel_groups=st.multiselect("Select age groups:",AGE_ORDER,default=AGE_ORDER,key="r3_groups")
    c1,c2,c3=st.columns(3)
    show_a3=c1.checkbox("Active",   value=True, key="r3_active")
    show_s3=c2.checkbox("Suspended",value=False,key="r3_susp")
    show_c3=c3.checkbox("Cancelled",value=False,key="r3_canc")
    metrics_trend=[]
    if show_a3: metrics_trend.append(("# Active members","Active"))
    if show_s3: metrics_trend.append(("# Suspended members","Suspended"))
    if show_c3: metrics_trend.append(("# Cancelled members","Cancelled"))
    if not metrics_trend: st.info("Select at least one metric."); return
    age_colors=[BI_ACCENT,BI_BLUE,BI_ORANGE,BI_RED,BI_PURPLE]
    for metric_col,metric_name in metrics_trend:
        fig_age=go.Figure()
        for i,grp in enumerate(sel_groups):
            grp_df=df_chart[df_chart["Age Group"]==grp].groupby("Date")[metric_col].mean().reset_index().sort_values("Date")
            if grp_df.empty: continue
            std_traces(fig_age, grp_df, "Date", metric_col, age_colors[i%len(age_colors)], grp)
        fig_age.update_layout(**std_layout(
            "", f"Avg {metric_name}", 500))
        show_chart(fig_age,use_container_width=True)
 # ── Avg membership by Region ──
    st.markdown('<div class="section-header">Avg Active Members by Region — Last 13 Months</div>', unsafe_allow_html=True)
    if "Region" in df_chart.columns and "# Active members" in df_chart.columns:
        regions = sorted(df_chart["Region"].dropna().unique().tolist())
        sel_regions = st.multiselect("Select regions:", regions, default=regions, key="r3_regions")
        region_colors = SERIES_COLORS
        fig_region = go.Figure()
        for i, region in enumerate(sel_regions):
            reg_df = df_chart[df_chart["Region"]==region].groupby("Date")["# Active members"].mean().reset_index().sort_values("Date")
            if reg_df.empty: continue
            std_traces(fig_region, reg_df, "Date", "# Active members",
                      region_colors[i%len(region_colors)], region)
        fig_region.update_layout(**std_layout(
            "", "Avg Active Members", 500))
        show_chart(fig_region, use_container_width=True)
# ══════════════════════════════════════════════════════════════════════════════
# REPORT 8 — Net Growth Rate %
# ══════════════════════════════════════════════════════════════════════════════
def report_net_growth(df_wm):
    st.markdown('<div class="report-title">Net Growth Rate %</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Weekly Membership — Net Growth Rate % = (New − Cancelled) ÷ Active × 100</div>', unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)
    loc_col = "Success Tutoring - Business name"
    cols = ["# Active members", "# New members", "# Cancelled members"]
    df = report_filters(df_wm.copy(), key_prefix="r8ng_f", show_date=False,
                        show_country=True, show_state=True, show_stage=True,
                        show_gpm=True, show_location=True, show_status=True)
    for c in cols:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    if df.empty:
        st.warning("No data for the selected filters."); return

    # ── Year-over-year weekly NGR % ───────────────────────────────────────
    st.markdown('<div class="section-header">Net Growth Rate % — Year-over-Year by Week</div>', unsafe_allow_html=True)
    d = add_week_year(df)
    years = sorted(d["Year"].unique().tolist(), reverse=True)
    if not years:
        st.info("No week/year data available."); return
    opts = yoy_options("ngr", years, avg_label="Equal weight per location",
                       avg_help="Off: (all new − all cancelled) ÷ all active, so bigger centres count more. "
                                "On: the average of each centre's own NGR %, so every centre counts equally.")
    sel = opts["years"]
    d = d[d["Year"].isin(sel)]
    if opts["comparative"]:
        d = comparative_only(d, sel, loc_col)
    if not sel or d.empty:
        st.info("Select at least one year with data.")
    else:
        latest_year = max(sel)
        latest_week = int(d[d["Year"] == latest_year]["Week"].max())
        weekly = {}
        for yr in sel:
            x = d[d["Year"] == yr]
            if opts["average"]:
                per_loc = x.groupby(["Week", loc_col])[cols].sum()
                per_loc = per_loc[per_loc["# Active members"] > 0]
                ngr = (per_loc["# New members"] - per_loc["# Cancelled members"]) / per_loc["# Active members"] * 100
                weekly[yr] = ngr.groupby(level="Week").mean()
            else:
                g = x.groupby("Week")[cols].sum()
                weekly[yr] = ((g["# New members"] - g["# Cancelled members"]) / g["# Active members"] * 100).where(g["# Active members"] > 0)
        label = "Avg NGR % per location (equal weight)" if opts["average"] else "Network NGR %"
        fig = yoy_bar_figure(weekly, latest_year, latest_week, f"{label} by week", label,
                             fmt=".2f", suffix="%", change="pts", holidays=opts["holidays"])
        fig.add_hline(y=0, line_color=BI_SUBTEXT, line_width=1)
        show_chart(fig, use_container_width=True)

    # ── Latest week by location ───────────────────────────────────────────
    latest_wk = df["Date"].max()
    loc_tbl = df[df["Date"] == latest_wk].groupby(loc_col)[cols].sum().reset_index()
    loc_tbl["Net Growth Rate %"] = ((loc_tbl["# New members"] - loc_tbl["# Cancelled members"])
                                    / loc_tbl["# Active members"] * 100).where(loc_tbl["# Active members"] > 0)
    loc_tbl = pd.DataFrame({"Location": loc_tbl[loc_col].str.replace("Success Tutoring - ", "", regex=False),
                            "Active": loc_tbl["# Active members"].astype(int), "New": loc_tbl["# New members"].astype(int),
                            "Cancelled": loc_tbl["# Cancelled members"].astype(int),
                            "Net Growth Rate %": loc_tbl["Net Growth Rate %"]}).sort_values("Net Growth Rate %", ascending=False)
    st.markdown(f'<div class="section-header">Net Growth Rate % by Location — {latest_wk.strftime("%d %b %Y")}</div>', unsafe_allow_html=True)
    color = lambda v: "" if pd.isna(v) or v == 0 else f"color: {BI_GREEN if v > 0 else BI_RED}; font-weight: 600"
    st.dataframe(loc_tbl.style.map(color, subset=["Net Growth Rate %"]).format({"Net Growth Rate %": "{:+.2f}%"}, na_rep="–"),
                 use_container_width=True, hide_index=True, height=min(560, 35 * (len(loc_tbl) + 1) + 3))

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 9 — Onboarding Progress
# ══════════════════════════════════════════════════════════════════════════════
def report_onboarding(df_wm):
    st.markdown('<div class="report-title">Onboarding Progress</div>',unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Vlookup (Weeks old) and Weekly Membership — onboarding week progress and pre-sale member counts</div>',unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)
    vl=apply_gpm_filter(load_vlookup())
    if vl.empty: st.warning("Vlookup tab not available."); return
    vl_f=report_filters(vl.copy(),key_prefix="r7",show_date=False,
                         show_country=True,show_state=True,show_stage=True,show_gpm=True,show_status=True)
    loc_col="Success Tutoring - Business name"
    if "Stage" not in vl_f.columns: st.warning("Stage column not found."); return
    onb=vl_f[vl_f["Status"]=="Pre-Sale"].copy()
    if onb.empty: st.info("No locations in Pre-Sale status for selected filters."); return
    if "Weeks old" in onb.columns:
        # Onboarding week = weeks since the first member; members = latest week's active members.
        onb["Onboarding week"]=onb["Weeks old"]
        latest=df_wm[df_wm["Date"]==df_wm["Date"].max()].set_index(loc_col)["# Active members"]
        onb["Onboarding Members"]=onb[loc_col].map(latest).fillna(0)
    st.markdown('<div class="section-header">Onboarding Status by Week</div>',unsafe_allow_html=True)
    if "Onboarding week" in onb.columns:
        onb_w=onb.sort_values("Onboarding week")
        gpm_list=onb_w["GPM"].dropna().unique().tolist() if "GPM" in onb_w.columns else []
        gpm_color_map={g:SERIES_COLORS[i % len(SERIES_COLORS)] for i,g in enumerate(gpm_list)}
        bar_colors=[gpm_color_map.get(str(row.get("GPM","")),BI_ACCENT) for _,row in onb_w.iterrows()]
        fig,ax=bi_fig(14,max(6,len(onb_w)*0.45))
        bars=ax.barh(onb_w[loc_col],onb_w["Onboarding week"],color=bar_colors,height=0.6)
        # Dual labels: week number INSIDE bar, pre-sale members OUTSIDE
        has_members = "Onboarding Members" in onb_w.columns
        for bar,(_,row) in zip(bars,onb_w.iterrows()):
            bar_w = bar.get_width()
            bar_y = bar.get_y() + bar.get_height()/2
            week_val = int(row["Onboarding week"])
            # Week number inside the bar (white, right-aligned)
            if bar_w > 1.5:
                ax.text(bar_w - 0.2, bar_y, str(week_val),
                        va="center", ha="right", color="white",
                        fontsize=8, fontweight="bold")
            # Pre-sale member count outside bar
            if has_members:
                members_val = int(row.get("Onboarding Members", 0))
                ax.text(bar_w + 0.2, bar_y,
                        f"{members_val} members",
                        va="center", ha="left", color=BI_SUBTEXT,
                        fontsize=7.5)

        # Updated milestones
        milestones={"Pre-Sale":0,"Assessments":7,"Soft Open":9,"Grand Opening":16}
        m_colors={"Pre-Sale":BI_BLUE,"Assessments":BI_ORANGE,"Soft Open":BI_PURPLE,"Grand Opening":BI_ACCENT}
        for name,week in milestones.items():
            ax.axvline(x=week,color=m_colors[name],linestyle="--",linewidth=1.2,alpha=0.9)
            ax.text(week+0.1,len(onb_w)-0.5,name,color=m_colors[name],fontsize=7,rotation=90,va="top")

        ax.set_xlabel("Onboarding Week",color=BI_TEXT,fontsize=9)
        ax.set_title("Onboarding Status by Week",color=BI_TEXT,fontsize=11,fontweight="bold")
        ax.invert_yaxis()
        # Extra right padding so "X members" labels aren't clipped
        x_max = onb_w["Onboarding week"].max()
        ax.set_xlim(right=x_max + (5 if has_members else 2))
        if gpm_list:
            patches=[mpatches.Patch(color=gpm_color_map[g],label=g) for g in gpm_list]
            ax.legend(handles=patches,loc="center right",bbox_to_anchor=(1.18,0.5),
                      facecolor=BI_CARD,edgecolor=BI_BORDER,labelcolor=BI_TEXT,fontsize=10)
        plt.tight_layout(); st.pyplot(fig); plt.close()
    with st.expander("Onboarding Location Detail",expanded=False):
            show_cols=[c for c in [loc_col,"Onboarding week","Onboarding Members","GPM","Country","Region","Weeks old","Age (Months)"] if c in onb.columns]
            st.dataframe(onb[show_cols].sort_values("Onboarding week").reset_index(drop=True),
                         use_container_width=True,hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 10 — Revenue
# ══════════════════════════════════════════════════════════════════════════════
def report_revenue(df_rv):
    st.markdown('<div class="report-title">Revenue</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Revenue sheet — all metrics pulled directly from source</div>', unsafe_allow_html=True)

    df_rv = apply_gpm_filter(df_rv)
    loc_col = "Success Tutoring - Business name"

    # ── Filters (same panel as the other reports) ─────────────────────────
    rv_filters = st.container(border=True)
    df_base = report_filters(df_rv.copy(), key_prefix="rv", show_date=False, show_country=True, show_state=True,
                             show_stage=True, show_gpm=True, show_location=True, show_status=True, panel=rv_filters)
    df = df_base.copy()
    if df.empty:
        st.warning("No revenue data available for selected filters."); return

    # ── Date filter ───────────────────────────────────────────────────────
    with rv_filters:
        df = checkbox_date_filter(df.copy(), key_prefix="r10rv")
    all_dates = sorted(df["Date"].dropna().unique())
    if not all_dates:
        st.warning("No dates found."); return

    latest_date = all_dates[-1]
    prev_date   = all_dates[-2] if len(all_dates) >= 2 else None

    # ── Metric definitions ────────────────────────────────────────────────
    SUM_METRICS   = ["Gross Revenue","Net Revenue","# Active Students","Total Sessions","Student Visits"]
    RATIO_METRICS = ["Revenue per Session","Revenue per Student","Sessions per Student",
                     "Student per Session","Sessions per Student Visit","Student Visits per Session"]
    ALL_METRICS   = SUM_METRICS + RATIO_METRICS

    fmt = {
        "Gross Revenue":               "${:,.0f}",
        "Net Revenue":                 "${:,.0f}",
        "# Active Students":           "{:,.0f}",
        "Student Visits":              "{:,.0f}",
        "Total Sessions":              "{:,.0f}",
        "Revenue per Session":         "${:,.2f}",
        "Revenue per Student":         "${:,.2f}",
        "Sessions per Student":        "{:,.1f}",
        "Student per Session":         "{:,.1f}",
        "Sessions per Student Visit":  "{:,.1f}",
        "Student Visits per Session":  "{:,.1f}",
    }

    kpi_fmt = {
        "Net Revenue":                 ("$", "{:,.0f}"),
        "Gross Revenue":               ("$", "{:,.0f}"),
        "# Active Students":           ("",  "{:,.0f}"),
        "Student Visits":              ("",  "{:,.0f}"),
        "Total Sessions":              ("",  "{:,.0f}"),
        "Revenue per Session":         ("$", "{:,.2f}"),
        "Revenue per Student":         ("$", "{:,.2f}"),
        "Sessions per Student":        ("",  "{:,.1f}"),
        "Student per Session":         ("",  "{:,.1f}"),
        "Sessions per Student Visit":  ("",  "{:,.1f}"),
        "Student Visits per Session":  ("",  "{:,.1f}"),
    }

    # ── Week aggregation — derive ratios from totals ───────────────────────
    def week_agg(df_sub, date_val):
        if date_val is None: return {}
        w = df_sub[df_sub["Date"] == date_val]
        row = {}
        for m in SUM_METRICS:
            row[m] = pd.to_numeric(w[m], errors="coerce").sum() if m in w.columns else 0
        net_rev  = row.get("Net Revenue", 0)
        students = row.get("# Active Students", 0)
        sessions = row.get("Total Sessions", 0)
        visits   = row.get("Student Visits", 0)
        row["Revenue per Session"]        = round(net_rev  / sessions, 2) if sessions > 0 else 0
        row["Revenue per Student"]        = round(net_rev  / students, 2) if students > 0 else 0
        row["Sessions per Student"]       = round(sessions / students, 1) if students > 0 else 0
        row["Student per Session"]        = round(students / sessions, 1) if sessions > 0 else 0
        row["Sessions per Student Visit"] = round(sessions / visits,   1) if visits   > 0 else 0
        row["Student Visits per Session"] = round(visits   / sessions, 1) if sessions > 0 else 0
        return row

    latest_t = week_agg(df, latest_date)
    prior_t  = week_agg(df, prev_date)

    # ── KPI Tiles ─────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Latest Week vs Prior Week</div>', unsafe_allow_html=True)

    kpi_list = [
        ("Net Revenue",                "$",  "green"),
        ("Gross Revenue",              "$",  "green"),
        ("# Active Students",          "",   "blue"),
        ("Student Visits",             "",   "blue"),
        ("Total Sessions",             "",   "blue"),
        ("Revenue per Session",        "$",  "green"),
        ("Revenue per Student",        "$",  "green"),
        ("Sessions per Student",       "",   "blue"),
        ("Student per Session",        "",   "blue"),
        ("Sessions per Student Visit", "",   "blue"),
        ("Student Visits per Session", "",   "blue"),
    ]

    kpi_cols = st.columns(4)
    for i, (metric, prefix, color) in enumerate(kpi_list):
        val  = latest_t.get(metric, 0)
        prev = prior_t.get(metric, 0)
        pfx, f = kpi_fmt.get(metric, ("", "{:,.1f}"))
        with kpi_cols[i % 4]:
            metric_card(metric, f"{pfx}{f.format(val)}", val - prev, color)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Build df_13m_base (respects all filters except date) ─────────────
    df_13m_base = df_base.copy()

    max_date   = df_13m_base["Date"].max()
    cutoff_13m = max_date - pd.DateOffset(months=13)
    df_13m     = df_13m_base[df_13m_base["Date"] >= cutoff_13m].copy()

    for m in ALL_METRICS:
        if m in df_13m.columns:
            df_13m[m] = pd.to_numeric(df_13m[m], errors="coerce").fillna(0)

    # ── Summary table helper ──────────────────────────────────────────────
    def build_summary_table(df_src, group_col, label_col):
        if group_col not in df_src.columns: return pd.DataFrame()
        grp = df_src.groupby(group_col).agg(
            **{m: (m, "sum") for m in SUM_METRICS if m in df_src.columns}
        ).reset_index()
        net  = grp["Net Revenue"]       if "Net Revenue"       in grp.columns else pd.Series([0]*len(grp))
        stud = grp["# Active Students"] if "# Active Students" in grp.columns else pd.Series([0]*len(grp))
        sess = grp["Total Sessions"]    if "Total Sessions"    in grp.columns else pd.Series([0]*len(grp))
        vis  = grp["Student Visits"]    if "Student Visits"    in grp.columns else pd.Series([0]*len(grp))
        grp["Revenue per Session"]        = (net  / sess).round(2).where(sess > 0, 0)
        grp["Revenue per Student"]        = (net  / stud).round(2).where(stud > 0, 0)
        grp["Sessions per Student"]       = (sess / stud).round(1).where(stud > 0, 0)
        grp["Student per Session"]        = (stud / sess).round(1).where(sess > 0, 0)
        grp["Sessions per Student Visit"] = (sess / vis ).round(1).where(vis  > 0, 0)
        grp["Student Visits per Session"] = (vis  / sess).round(1).where(sess > 0, 0)
        grp = grp.rename(columns={group_col: label_col})
        sort_col = "Net Revenue" if "Net Revenue" in grp.columns else grp.columns[1]
        grp = grp.sort_values(sort_col, ascending=False).reset_index(drop=True)
        return grp

    df_latest_all = df_13m_base[df_13m_base["Date"] == max_date].copy()
    for m in SUM_METRICS:
        if m in df_latest_all.columns:
            df_latest_all[m] = pd.to_numeric(df_latest_all[m], errors="coerce").fillna(0)

    # ── Country table ─────────────────────────────────────────────────────
    st.markdown('<div class="section-header">All Metrics by Country — Latest Week</div>', unsafe_allow_html=True)
    df_country = build_summary_table(df_latest_all, "Country", "Country")
    if not df_country.empty:
        st.dataframe(df_country.style.format({k: v for k, v in fmt.items() if k in df_country.columns}),
                     use_container_width=True, hide_index=True)
    else:
        st.info("No country data available.")

    # ── Region table ──────────────────────────────────────────────────────
    st.markdown('<div class="section-header">All Metrics by Region — Latest Week</div>', unsafe_allow_html=True)
    df_region = build_summary_table(df_latest_all, "Region", "Region")
    if not df_region.empty:
        st.dataframe(df_region.style.format({k: v for k, v in fmt.items() if k in df_region.columns}),
                     use_container_width=True, hide_index=True)
    else:
        st.info("No region data available.")

    # ── Year-over-year weekly revenue ─────────────────────────────────────
    st.markdown('<div class="section-header">Year-over-Year Weekly Comparison</div>', unsafe_allow_html=True)
    RV_RATIOS = {"Revenue per Session": ("Net Revenue", "Total Sessions"), "Revenue per Student": ("Net Revenue", "# Active Students"),
                 "Sessions per Student": ("Total Sessions", "# Active Students"), "Student per Session": ("# Active Students", "Total Sessions"),
                 "Sessions per Student Visit": ("Total Sessions", "Student Visits"), "Student Visits per Session": ("Student Visits", "Total Sessions")}
    MONEY = {"Gross Revenue", "Net Revenue", "Revenue per Session", "Revenue per Student"}
    avail = [m for m in ALL_METRICS if m in df_base.columns and (m in SUM_METRICS or all(c in df_base.columns for c in RV_RATIOS.get(m, ())))]
    avail = sorted(avail, key=lambda m: m != "Net Revenue")   # Net Revenue first (default)
    dy = add_week_year(df_base)
    for m in SUM_METRICS:
        if m in dy.columns: dy[m] = pd.to_numeric(dy[m], errors="coerce").fillna(0)
    start_col = "# Active Students" if "# Active Students" in dy.columns else "Net Revenue"
    years = sorted(dy["Year"].unique().tolist(), reverse=True)
    if years and avail:
        opts = yoy_options("rvy", years, metrics={m: m.replace("# ", "") for m in avail},
                           avg_metrics=set(SUM_METRICS))
        metric, sel = opts["metric"], opts["years"]
        dy = dy[dy["Year"].isin(sel)]
        if opts["comparative"]:
            dy = comparative_only(dy, sel, loc_col)
        if not sel or dy.empty:
            st.info("Select at least one year with data.")
        else:
            latest_year = max(sel)
            latest_week = int(dy[dy["Year"] == latest_year]["Week"].max())
            weekly = {}
            for yr in sel:
                g = dy[dy["Year"] == yr].groupby("Week").sum(numeric_only=True)
                if metric in RV_RATIOS:
                    num, den = RV_RATIOS[metric]
                    weekly[yr] = (g[num] / g[den]).where(g[den] > 0)
                else:
                    weekly[yr] = g[metric]
                    if opts["average"]:
                        weekly[yr] = weekly[yr] / trading_locations(dy, loc_col, start_col, yr)
            name = metric.replace("# ", "")
            label = f"Avg {name.lower()} per location" if opts["average"] else name
            money = metric in MONEY
            fmt_ = ",.2f" if metric in RV_RATIOS or opts["average"] else ",.0f"
            fig = yoy_bar_figure(weekly, latest_year, latest_week, f"{label} by week", label, fmt=fmt_,
                                 prefix="$" if money else "", holidays=opts["holidays"])
            show_chart(fig, use_container_width=True)
    else:
        st.info("No revenue data by week available.")

    # ── Location comparison ───────────────────────────────────────────────
    st.markdown('<div class="section-header">Location Comparison — Last 13 Months</div>', unsafe_allow_html=True)
    measures = {}
    for m in avail:
        entry = {"kind": "ratio", "num": RV_RATIOS[m][0], "den": RV_RATIOS[m][1], "fmt": ",.2f"} if m in RV_RATIOS \
                else {"kind": "sum", "col": m, "fmt": ",.0f"}
        if m in MONEY: entry["prefix"] = "$"
        measures[m.replace("# ", "")] = entry
    if measures and not df_13m.empty:
        location_comparison(df_13m, loc_col, measures, "rv_cmp", "Net Revenue" if "Net Revenue" in df_13m.columns else start_col, start_col)

    # ── Performance Bar Chart ─────────────────────────────────────────────
    st.markdown('<div class="section-header">Performance by Location</div>', unsafe_allow_html=True)
    bc1, bc2 = st.columns([3, 1])
    with bc1: bar_metric = st.selectbox("Metric:", [m for m in ALL_METRICS if m in df.columns], key="rv_bar_metric")
    with bc2: bar_scope  = st.selectbox("Show:", ["Latest week only", "All selected weeks combined"], key="rv_bar_scope")

    df_bar   = df[df["Date"] == latest_date] if bar_scope == "Latest week only" else df
    bar_agg  = "sum" if bar_metric in SUM_METRICS else "mean"
    bar_data = df_bar.groupby(loc_col)[bar_metric].agg(bar_agg).reset_index().sort_values(bar_metric, ascending=False)

    prefix_map = {"Net Revenue": "$", "Gross Revenue": "$", "Revenue per Session": "$", "Revenue per Student": "$"}
    prefix_bar = prefix_map.get(bar_metric, "")

    fig_bar = go.Figure(go.Bar(
        x=bar_data[loc_col].str.replace("Success Tutoring - ", "", regex=False),
        y=bar_data[bar_metric],
        marker_color=BI_ACCENT,
        text=[f"{prefix_bar}{v:,.1f}" for v in bar_data[bar_metric]],
        textposition="outside",
    ))
    fig_bar.update_layout(
        plot_bgcolor=BI_CHART_BG, paper_bgcolor=BI_CHART_BG,
        font=dict(color=BI_TEXT),
        xaxis=dict(showgrid=False, color=BI_SUBTEXT, tickangle=-45,
                   title=dict(text="Location", font=dict(color=BI_SUBTEXT, size=11))),
        yaxis=dict(showgrid=True, gridcolor=BI_GRID, color=BI_SUBTEXT, tickprefix=prefix_bar,
                   title=dict(text=bar_metric, font=dict(color=BI_SUBTEXT, size=11))),
        margin=dict(l=40, r=20, t=30, b=160),
        height=450,
    )
    show_chart(fig_bar, use_container_width=True)

    # ── Location table ────────────────────────────────────────────────────
    st.markdown('<div class="section-header">All Metrics by Location — Latest Week</div>', unsafe_allow_html=True)
    metric_names = [m for m in ALL_METRICS if m in df.columns]
    df_table = df[df["Date"] == latest_date].groupby(loc_col).agg(
        **{m: (m, "sum") for m in SUM_METRICS if m in df.columns}
    ).reset_index()
    net  = df_table["Net Revenue"]       if "Net Revenue"       in df_table.columns else pd.Series([0]*len(df_table))
    stud = df_table["# Active Students"] if "# Active Students" in df_table.columns else pd.Series([0]*len(df_table))
    sess = df_table["Total Sessions"]    if "Total Sessions"    in df_table.columns else pd.Series([0]*len(df_table))
    vis  = df_table["Student Visits"]    if "Student Visits"    in df_table.columns else pd.Series([0]*len(df_table))
    df_table["Revenue per Session"]        = (net  / sess).round(2).where(sess > 0, 0)
    df_table["Revenue per Student"]        = (net  / stud).round(2).where(stud > 0, 0)
    df_table["Sessions per Student"]       = (sess / stud).round(1).where(stud > 0, 0)
    df_table["Student per Session"]        = (stud / sess).round(1).where(sess > 0, 0)
    df_table["Sessions per Student Visit"] = (sess / vis ).round(1).where(vis  > 0, 0)
    df_table["Student Visits per Session"] = (vis  / sess).round(1).where(sess > 0, 0)
    df_table = df_table.rename(columns={loc_col: "Location"})
    df_table["Location"] = df_table["Location"].str.replace("Success Tutoring - ", "", regex=False)
    df_table = df_table.sort_values("Net Revenue" if "Net Revenue" in df_table.columns else metric_names[0],
                                     ascending=False).reset_index(drop=True)
    st.dataframe(
        df_table.style.format({k: v for k, v in fmt.items() if k in df_table.columns}),
        use_container_width=True, hide_index=True,
    )
# ══════════════════════════════════════════════════════════════════════════════
# REPORT 11 — Outliers & Alerts
# ══════════════════════════════════════════════════════════════════════════════
def report_outliers_alerts(df_wm, df_rv):
    st.markdown('<div class="report-title">Outliers & Alerts</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Weekly Membership, Revenue and Vlookup — centres that need attention, network outliers and data checks</div>', unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)
    loc_col = "Success Tutoring - Business name"
    member_cols = ["# Active members", "# New members", "# Cancelled members", "# Suspended members"]

    df = report_filters(df_wm.copy(), key_prefix="r10", show_date=False,
                        show_country=True, show_state=True, show_stage=True,
                        show_gpm=True, show_status=True)
    for c in member_cols:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    all_dates = sorted(df["Date"].unique())
    if not all_dates: st.warning("No data."); return
    latest_date = all_dates[-1]

    # One row per location per week
    wk = (df.groupby([loc_col, "Date"])[member_cols].sum().reset_index().sort_values([loc_col, "Date"]))
    wk["Churn %"] = (wk["# Cancelled members"] / wk["# Active members"] * 100).where(wk["# Active members"] > 0)
    latest = wk[wk["Date"] == latest_date].set_index(loc_col)

    loc_latest = latest.reset_index().rename(columns={"# Active members": "Active", "# New members": "New",
                                                      "# Cancelled members": "Cancelled", "# Suspended members": "Suspended"})
    loc_latest["Churn Rate %"] = loc_latest.apply(lambda r: round(r["Cancelled"] / r["Active"] * 100, 1) if r["Active"] > 0 else 0.0, axis=1)
    loc_latest["Net Growth Rate %"] = loc_latest.apply(lambda r: round((r["New"] - r["Cancelled"]) / r["Active"] * 100, 2) if r["Active"] > 0 else 0.0, axis=1)
    vl = load_vlookup()
    if not vl.empty and "Stage" in vl.columns:
        loc_latest = loc_latest.merge(vl[[c for c in [loc_col, "Stage", "GPM", "Country", "Region", "Age (Months)"] if c in vl.columns]],
                                      on=loc_col, how="left")
    short = lambda x: str(x).replace("Success Tutoring - ", "")

    def table(d, **kw):
        raw = d.data if hasattr(d, "data") and not isinstance(d, pd.DataFrame) else d
        if raw.empty:
            st.caption("Nothing to flag this week.")
        else:
            st.dataframe(d, use_container_width=True, hide_index=True, **kw)

    trading = loc_latest[loc_latest["Active"] > 0]
    st.markdown(f'<div class="metric-card" style="padding:10px 14px;font-size:0.9em">'
                f'<b>Latest week: {pd.Timestamp(latest_date).strftime("%d %b %Y")}</b> · '
                f'{len(trading)} trading centres · Avg active: <b>{trading["Active"].mean():.0f}</b> · '
                f'Avg NGR: <b>{trading["Net Growth Rate %"].mean():.1f}%</b> · '
                f'Avg churn: <b>{trading["Churn Rate %"].mean():.1f}%</b></div>', unsafe_allow_html=True)

    # ── Watchlist ─────────────────────────────────────────────────────────
    w1, w2 = st.columns(2)
    with w1:
        # (a) Declining streaks: active members down 3+ weeks in a row, ending this week
        st.markdown('<div class="section-header">Declining Streaks — Active Down 3+ Weeks in a Row</div>', unsafe_allow_html=True)
        rows = []
        for loc, g in wk.groupby(loc_col):
            vals = g["# Active members"].tolist()
            if not vals or g["Date"].iloc[-1] != latest_date: continue
            streak = 0
            for i in range(len(vals) - 1, 0, -1):
                if vals[i] < vals[i - 1]: streak += 1
                else: break
            if streak >= 3:
                start = vals[-1 - streak]
                rows.append({"Location": short(loc), "Weeks declining": streak, "Active now": int(vals[-1]),
                             "Active before": int(start), "Change": int(vals[-1] - start)})
        table(pd.DataFrame(rows).sort_values("Weeks declining", ascending=False) if rows else pd.DataFrame())
    with w2:
        # (b) Churn spikes: this week's churn more than double the centre's own 13-week average
        st.markdown('<div class="section-header">Churn Spikes — More Than 2× Own 13-Week Average</div>', unsafe_allow_html=True)
        rows = []
        for loc, g in wk.groupby(loc_col):
            if g["Date"].iloc[-1] != latest_date: continue
            now = g["Churn %"].iloc[-1]
            prior = g["Churn %"].iloc[-14:-1].dropna()
            base = prior.mean() if len(prior) else None
            if pd.notna(now) and base and base > 0 and now > 2 * base:
                rows.append({"Location": short(loc), "Churn this week %": round(now, 1),
                             "13-week avg %": round(base, 1), "× average": round(now / base, 1),
                             "Cancelled": int(g["# Cancelled members"].iloc[-1])})
        table(pd.DataFrame(rows).sort_values("× average", ascending=False) if rows else pd.DataFrame())

    # (c) Suspension risk: suspended ÷ active this week
    st.markdown('<div class="section-header">Suspension Risk — Suspended ÷ Active This Week</div>', unsafe_allow_html=True)
    net_ratio = trading["Suspended"].sum() / trading["Active"].sum() * 100 if trading["Active"].sum() else 0
    st.caption(f"Network: {net_ratio:.1f}% of active members are suspended. Centres at more than twice that are highlighted.")
    sr = trading.assign(**{"Suspended %": (trading["Suspended"] / trading["Active"] * 100).round(1)})
    sr = sr.sort_values("Suspended %", ascending=False).head(10)
    sr_tbl = pd.DataFrame({"Location": sr[loc_col].map(short), "Suspended %": sr["Suspended %"],
                           "Suspended": sr["Suspended"].astype(int), "Active": sr["Active"].astype(int),
                           "GPM": sr["GPM"] if "GPM" in sr.columns else ""})
    if not sr_tbl.empty:
        hl = lambda v: f"background-color: rgba(200,50,47,0.18); font-weight: 600" if net_ratio and v > 2 * net_ratio else ""
        table(sr_tbl.style.map(hl, subset=["Suspended %"]).format({"Suspended %": "{:.1f}%"}))
    else:
        table(sr_tbl)

    # (e) Age-adjusted performance: active members vs centres open a similar number of months
    st.markdown('<div class="section-header">Age-Adjusted Performance — vs Centres Open a Similar Time</div>', unsafe_allow_html=True)
    if "Age (Months)" in trading.columns:
        bands = [(0, 6, "0–6 months"), (6, 12, "6–12 months"), (12, 24, "12–24 months"), (24, 36, "24–36 months"), (36, 10**6, "36+ months")]
        def band(m):
            try: m = float(m)
            except: return None
            for lo, hi, name in bands:
                if lo <= m < hi: return name
            return None
        # Band averages use every trading centre in the network, whatever the filters
        full = load_weekly_membership()
        full = full[full["Date"] == full["Date"].max()]
        full = full.groupby(loc_col)["# Active members"].sum().reset_index()
        full = full[full["# Active members"] > 0].merge(vl[[loc_col, "Age (Months)"]], on=loc_col, how="left")
        full["Band"] = full["Age (Months)"].map(band)
        band_avg = full.groupby("Band")["# Active members"].mean()
        aa = trading.copy()
        aa["Band"] = aa["Age (Months)"].map(band)
        aa = aa.dropna(subset=["Band"])
        aa["Band average"] = aa["Band"].map(band_avg).round(0)
        aa["Index"] = (aa["Active"] / aa["Band average"]).round(2)
        order = {name: i for i, (_, _, name) in enumerate(bands)}
        aa = aa.sort_values(["Band", "Index"], key=lambda c: c.map(order) if c.name == "Band" else -c)
        aa_tbl = pd.DataFrame({"Location": aa[loc_col].map(short), "Age (months)": aa["Age (Months)"].astype(int),
                               "Age band": aa["Band"], "Active": aa["Active"].astype(int),
                               "Band average": aa["Band average"].astype(int), "Index": aa["Index"]})
        st.caption("Index = active members ÷ average of trading centres in the same age band across the network. "
                   "1.00 = typical for its age · 1.28 = 28% ahead · 0.80 = 20% behind.")
        def shade(v):
            if pd.isna(v): return ""
            a = min(abs(v - 1) / 0.5, 1) * 0.55
            return f"background-color: rgba({'46,133,64' if v >= 1 else '200,50,47'},{a:.2f})"
        table(aa_tbl.style.map(shade, subset=["Index"]).format({"Index": "{:.2f}"}),
              height=min(500, 35 * (len(aa_tbl) + 1) + 3))
    else:
        st.caption("Age (Months) not found in Vlookup.")

    # ── Rankings and statistical outliers (latest week) ───────────────────
    def show_ranked(d, sort_col, title, n=5, ascending=False):
        st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)
        cols = [c for c in [sort_col, "Active", "New", "Cancelled", "Churn Rate %", "Net Growth Rate %", "Stage"] if c in d.columns]
        cols = list(dict.fromkeys(cols))
        r = d.sort_values(sort_col, ascending=ascending).head(n)
        table(pd.concat([r[loc_col].map(short).rename("Location"), r[cols]], axis=1))

    def show_outliers(d, col, title):
        st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)
        mean, std = d[col].mean(), d[col].std()
        st.caption(f"Mean {mean:.1f} · Std dev {std:.1f} · Outside {mean-2*std:.1f} to {mean+2*std:.1f}")
        o = d[(d[col] > mean + 2 * std) | (d[col] < mean - 2 * std)].copy()
        o["vs Mean"] = (o[col] - mean).round(2)
        cols = [c for c in [col, "vs Mean", "Active", "Stage"] if c in o.columns]
        o = o.sort_values(col, ascending=False)
        table(pd.concat([o[loc_col].map(short).rename("Location"), o[cols]], axis=1))

    c1, c2 = st.columns(2)
    with c1:
        show_ranked(trading, "Active", "Top 5 — Active Members")
        show_ranked(trading, "New", "Top 5 — New Members")
        show_ranked(trading, "Net Growth Rate %", "Top 5 — Net Growth Rate %")
    with c2:
        show_ranked(trading, "Churn Rate %", "Top 5 — Highest Churn Rate %")
        show_ranked(trading, "Net Growth Rate %", "Bottom 5 — Lowest Net Growth Rate %", ascending=True)
        show_outliers(trading, "Net Growth Rate %", "Outliers — Net Growth Rate % (2σ)")

    # (g-style threshold kept from before) Centres below 50 active
    st.markdown('<div class="section-header">Centres With Fewer Than 50 Active Members</div>', unsafe_allow_html=True)
    b50 = trading[trading["Active"] < 50].sort_values("Active")
    table(pd.concat([b50[loc_col].map(short).rename("Location"),
                     b50[[c for c in ["Active", "New", "Cancelled", "Churn Rate %", "Stage", "GPM"] if c in b50.columns]]], axis=1))

    # ── Weekly network summary ────────────────────────────────────────────
    st.markdown('<div class="section-header">Weekly Network Summary</div>', unsafe_allow_html=True)
    tr = wk[wk["# Active members"] > 0]
    summ = tr.groupby("Date").agg(Locations=(loc_col, "nunique"), Active=("# Active members", "sum"),
                                  Median=("# Active members", "median"), New=("# New members", "sum"),
                                  Cancelled=("# Cancelled members", "sum"), Suspended=("# Suspended members", "sum")
                                  ).reset_index().sort_values("Date", ascending=False)
    summ["Avg Active"] = (summ["Active"] / summ["Locations"]).round(1)
    summ["NGR %"] = ((summ["New"] - summ["Cancelled"]) / summ["Active"] * 100).round(1)
    summ["Churn %"] = (summ["Cancelled"] / summ["Active"] * 100).round(1)
    summ["Week"] = summ["Date"].dt.strftime("%d %b %Y")
    summ = summ[["Week", "Locations", "Active", "Avg Active", "Median", "New", "Cancelled", "Suspended", "NGR %", "Churn %"]]
    table(summ.style.format({"Active": "{:,.0f}", "Avg Active": "{:.1f}", "Median": "{:.1f}",
                             "NGR %": "{:.1f}%", "Churn %": "{:.1f}%"}), height=400)

    # ── (h) Data-quality checks ───────────────────────────────────────────
    st.markdown('<div class="section-header">Data Quality Checks</div>', unsafe_allow_html=True)
    shown = set(df[loc_col].dropna())
    checks = []
    if not vl.empty:
        for col in ["Stage", "Country", "Region", "GPM"]:
            if col in vl.columns:
                missing = vl[vl[col].astype(str).str.strip().isin(["", "nan", "None"])]
                missing = missing[missing[loc_col].isin(shown) | (st.session_state.get("access_level") == "admin")]
                for loc in missing[loc_col]:
                    checks.append({"Check": f"Missing {col} in Vlookup", "Location": short(loc) or "(blank name)", "Detail": ""})
        not_in_vl = sorted(shown - set(vl[loc_col].dropna()))
        for loc in not_in_vl:
            checks.append({"Check": "In Weekly Membership but not in Vlookup", "Location": short(loc),
                           "Detail": "Check the name matches exactly"})
    # Missing weeks: centres trading last week with no row this week, and gaps in the last 13 weeks
    recent = all_dates[-13:]
    for loc, g in wk.groupby(loc_col):
        dates = set(g["Date"])
        if len(all_dates) >= 2 and all_dates[-2] in dates and latest_date not in dates \
                and g[g["Date"] == all_dates[-2]]["# Active members"].sum() > 0:
            checks.append({"Check": "No data for latest week", "Location": short(loc),
                           "Detail": f"Had data on {pd.Timestamp(all_dates[-2]).strftime('%d %b %Y')}"})
        first = g["Date"].min()
        gaps = [d for d in recent if d >= first and d not in dates and d != latest_date]
        if gaps:
            checks.append({"Check": "Missing weeks (last 13 weeks)", "Location": short(loc),
                           "Detail": ", ".join(pd.Timestamp(d).strftime("%d %b") for d in gaps[:6]) + ("…" if len(gaps) > 6 else "")})
    # Revenue recorded with zero active members
    rv = apply_gpm_filter(df_rv) if not df_rv.empty else df_rv
    if not rv.empty and "Net Revenue" in rv.columns and loc_col in rv.columns:
        rv_latest = rv[rv["Date"] == rv["Date"].max()]
        rv_sum = rv_latest.groupby(loc_col)["Net Revenue"].sum()
        act = latest["# Active members"] if not latest.empty else pd.Series(dtype=float)
        for loc, rev in rv_sum.items():
            if rev > 0 and act.get(loc, 0) == 0 and loc in shown:
                checks.append({"Check": "Revenue with zero active members", "Location": short(loc),
                               "Detail": f"${rev:,.0f} net revenue, week of {pd.Timestamp(rv['Date'].max()).strftime('%d %b %Y')}"})
    if checks:
        st.caption(f"{len(checks)} item(s) to check in the source sheets.")
        table(pd.DataFrame(checks))
    else:
        st.caption("All checks passed.")

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ══════════════════════════════════════════════════════════════════════════════
def login_section():
    col1,col2,col3=st.columns([1,2,1])
    with col2:
        st.markdown("<br><br>",unsafe_allow_html=True)
        if os.path.exists("logo.png"):
            lc1,lc2,lc3=st.columns([1,1,1])
            with lc2: st.image("logo.png",use_container_width=True)
        else:
            st.markdown(f'<div style="color:{BI_ACCENT};font-size:3em;text-align:center"></div>',
                        unsafe_allow_html=True)
        st.markdown(f'<h2 style="color:{BI_TEXT};text-align:center;margin-top:8px">Success Tutoring Dashboard</h2>',
                    unsafe_allow_html=True)
        st.markdown(f'<p style="color:{BI_SUBTEXT};text-align:center;margin-bottom:24px">Sign in to continue</p>',
                    unsafe_allow_html=True)
        import urllib.parse
        CLIENT_ID = st.secrets["GOOGLE_CLIENT_ID"]
        CLIENT_SECRET = st.secrets["GOOGLE_CLIENT_SECRET"]
        REDIRECT_URI = st.secrets.get("REDIRECT_URI", "https://j7ky6kl5hwlbrjpxtuk8ce.streamlit.app/")
        AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
        TOKEN_URL = "https://oauth2.googleapis.com/token"
        USERINFO_URL = "https://www.googleapis.com/oauth2/v3/userinfo"
        params = st.query_params
        if "code" in params and not st.session_state.get("logged_in") and not st.session_state.get("oauth_processing"):
            st.session_state["oauth_processing"] = True
            import httpx
            code = params["code"]
            try:
                resp = httpx.post(TOKEN_URL, data={
                    "code": code, "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
                    "redirect_uri": REDIRECT_URI, "grant_type": "authorization_code",
                }, timeout=10)
                token_data = resp.json()
                access_token = token_data.get("access_token")
                if access_token:
                    user_resp = httpx.get(USERINFO_URL,
                        headers={"Authorization": f"Bearer {access_token}"}, timeout=10)
                    user_info = user_resp.json()
                    email = user_info.get("email","").lower().strip()
                    name = user_info.get("name", email.split("@")[0].title())
                    st.cache_data.clear()
                    perms = get_user_permissions(email)
                    if perms["tabs"]:
                        st.query_params.clear()
                        st.session_state["logged_in"] = True
                        st.session_state["user_email"] = email
                        st.session_state["user_name"] = name
                        st.session_state["access_level"] = perms["access_level"]
                        st.session_state["gpm_filter"] = perms["gpm_filter"]
                        st.session_state["allowed_locations"] = perms.get("allowed_locations", [])
                        st.session_state["allowed_tabs"] = perms["tabs"]
                        log_access(email, name, "Login")
                        st.rerun()
                    else:
                        st.query_params.clear()
                        flag_unknown_user(email)
                        st.error(f"{email} is not authorised. Contact your administrator.")
                else:
                    st.session_state["oauth_processing"] = False
                    st.query_params.clear()
                    st.error(f"Google auth failed: {token_data.get('error_description','Unknown error')}")
            except Exception as ex:
                st.error(f"Auth error: {ex}")
        else:
            google_auth_url = (
                f"{AUTH_URL}?response_type=code"
                f"&client_id={CLIENT_ID}"
                f"&redirect_uri={urllib.parse.quote(REDIRECT_URI)}"
                f"&scope={urllib.parse.quote('openid email profile')}"
                f"&access_type=offline"
                f"&prompt=select_account"
            )
            st.markdown(f"""
            <div style="text-align:center;margin:16px 0">
                <a href="{google_auth_url}" target="_blank" style="
                    display:inline-flex;align-items:center;gap:10px;
                    background:white;color:#1a1a2e;
                    padding:14px 32px;border-radius:8px;
                    font-weight:700;font-size:1.1em;
                    text-decoration:none;
                    border:2px solid #4285F4;
                    box-shadow:0 4px 12px rgba(66,133,244,0.4);
                ">
                    <img src="https://www.google.com/favicon.ico" width="24" height="24"/>
                    Sign in with Google
                </a>
            </div>
            """, unsafe_allow_html=True)
        st.caption("Only approved team members can access this dashboard.")

# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
if "logged_in" not in st.session_state:
    st.session_state["logged_in"]=False
if not st.session_state["logged_in"]:
    login_section(); st.stop()

st.empty()

user_email=st.session_state["user_email"]
user_name=st.session_state["user_name"]

REPORTS=[
    "1 · Campus Locations",
    "2 · Membership",
    "3 · Membership by Age",
    "8 · Net Growth Rate %",
    "9 · Onboarding Progress",
    "10 · Revenue",
    "11 · AI Outlier Analysis",
]

def report_allowed(report, allowed_tabs):
    """allowed_tabs may contain "all", report numbers ("10") or names ("Revenue")."""
    tabs = [t.strip().lower() for t in allowed_tabs]
    if "all" in tabs:
        return True
    num, name = report.split(" · ", 1)
    return num in tabs or name.lower() in tabs or report.lower() in tabs

REPORTS = [r for r in REPORTS if report_allowed(r, st.session_state.get("allowed_tabs", []))]
ADMIN_PAGES = ["⚙ Weekly Upload", "⚙ Locations", "⚙ Sheet Setup"] if st.session_state.get("access_level") == "admin" else []
if not REPORTS:
    st.warning("No reports are assigned to your account. Please contact your administrator."); st.stop()

# Report menu: (report, menu label, icon), grouped like Power BI report sections
NAV_GROUPS = [
    ("Network", [
        ("1 · Campus Locations",   "Campus Locations",    ":material/location_on:"),
        ("9 · Onboarding Progress", "Onboarding Progress", ":material/checklist:"),
    ]),
    ("Membership", [
        ("2 · Membership",          "Membership",          ":material/groups:"),
        ("3 · Membership by Age",   "Membership by Age",   ":material/stacks:"),
    ]),
    ("Growth", [
        ("8 · Net Growth Rate %",   "Net Growth Rate %",   ":material/trending_up:"),
    ]),
    ("Finance", [
        ("10 · Revenue",            "Revenue",             ":material/attach_money:"),
    ]),
    ("Insights", [
        ("11 · AI Outlier Analysis",            "Outliers & Alerts",    ":material/notification_important:"),
    ]),
]
# Admin-only pages (ADMIN_PAGES is empty for everyone else, so non-admins never see or reach them)
ADMIN_NAV = [
    ("⚙ Weekly Upload", "Weekly Upload", ":material/upload_file:"),
    ("⚙ Locations",     "Locations",     ":material/store:"),
    ("⚙ Sheet Setup",   "Sheet Setup",   ":material/table_chart:"),
]
REPORT_LABELS = {r: label for _, items in NAV_GROUPS for r, label, _ in items}
REPORT_LABELS.update({r: label for r, label, _ in ADMIN_NAV})

@st.cache_data(ttl=300)
def data_loaded_at():
    """Time the sheet data was last fetched (cleared together with the data cache)."""
    from datetime import datetime
    from zoneinfo import ZoneInfo
    return datetime.now(ZoneInfo("Australia/Sydney"))

if st.session_state.get("selected_report") not in REPORTS + ADMIN_PAGES:
    st.session_state["selected_report"] = REPORTS[0]

with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=130)
    else:
        st.markdown(f'<div style="color:{BI_ACCENT};font-size:1.1em;font-weight:700;padding:8px 0">Success Tutoring</div>',
                    unsafe_allow_html=True)
    for group, items in NAV_GROUPS + [("Admin", ADMIN_NAV)]:
        visible = [(r, label, icon) for r, label, icon in items if r in REPORTS + ADMIN_PAGES]
        if not visible:
            continue
        st.markdown(f'<div class="nav-group">{group}</div>', unsafe_allow_html=True)
        for r, label, icon in visible:
            is_active = st.session_state["selected_report"] == r
            if st.button(label, key=f"nav_{r}", icon=icon, use_container_width=True,
                         type="primary" if is_active else "secondary"):
                st.session_state["selected_report"] = r
                st.rerun()

    loaded = data_loaded_at()
    st.markdown(f'<div class="side-footer"><div class="side-updated">Data updated {loaded.strftime("%-I:%M %p").lower()} {loaded.strftime("%Z")}</div></div>',
                unsafe_allow_html=True)
    if st.button("Refresh data", icon=":material/refresh:", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    initial = (user_name or user_email or "?").strip()[:1].upper()
    st.markdown(f"""<div class="user-card">
        <div class="user-avatar">{html.escape(initial)}</div>
        <div class="user-text"><div class="user-name">{html.escape(user_name)}</div>
        <div class="user-mail">{html.escape(user_email)}</div></div>
    </div>""", unsafe_allow_html=True)
    if st.button("Log out", icon=":material/logout:", use_container_width=True):
        log_access(
            st.session_state.get("user_email",""),
            st.session_state.get("user_name",""),
            "Logout"
        )
        st.session_state.clear()
        st.rerun()

selected_report = st.session_state["selected_report"]

# Top bar: dashboard name, current report, and who is signed in
st.markdown(f"""<div class="topbar">
    <div class="topbar-title">Success Tutoring Dashboard <span>/ {REPORT_LABELS.get(selected_report, selected_report)}</span></div>
    <div class="topbar-meta"><b>{html.escape(st.session_state.get("access_level", ""))}</b> · {html.escape(user_name)}</div>
</div>""", unsafe_allow_html=True)

with st.spinner("Loading data..."):
    try:
        df_wm=load_weekly_membership()
    except Exception as e:
        st.error(f"Could not load Weekly Membership: {e}"); st.stop()
    try:
        df_rv=load_revenue()
    except Exception as e:
        st.error(f"Could not load Revenue: {e}"); st.stop()

if df_wm.empty:
    st.warning("Weekly Membership sheet is empty."); st.stop()

df_wm = apply_gpm_filter(df_wm)
if df_wm.empty:
    st.warning("No locations are assigned to your account. Please contact your administrator."); st.stop()

if selected_report=="1 · Campus Locations":           report_locations(df_wm)
elif selected_report=="2 · Membership":               report_membership(df_wm, df_rv)
elif selected_report=="3 · Membership by Age":        report_age_combined(df_wm)
elif selected_report=="8 · Net Growth Rate %":          report_net_growth(df_wm)
elif selected_report=="9 · Onboarding Progress":      report_onboarding(df_wm)
elif selected_report=="10 · Revenue":                 report_revenue(df_rv)
elif selected_report=="11 · AI Outlier Analysis":     report_outliers_alerts(df_wm, df_rv)
elif selected_report in ADMIN_PAGES:
    spreadsheet = get_sheets_client().open_by_key(SHEET_ID)
    page = {"⚙ Weekly Upload": admin_pages.page_weekly_upload,
            "⚙ Locations": admin_pages.page_locations,
            "⚙ Sheet Setup": admin_pages.page_sheet_setup}[selected_report]
    page(spreadsheet, df_wm, df_rv, st.cache_data.clear)
