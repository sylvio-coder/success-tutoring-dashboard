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
import anthropic
import os
from dotenv import load_dotenv

load_dotenv()

SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
ANTHROPIC_KEY = os.getenv("ANTHROPIC_API_KEY")
SERVICE_ACCOUNT_FILE = "service_account.json"
SCOPES_SHEETS = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

st.set_page_config(page_title="Success Tutoring Dashboard", page_icon="📊", layout="wide")

# ── Theme ─────────────────────────────────────────────────────────────────────
BI_BG       = "#1a1a2e"
BI_CARD     = "#16213e"
BI_ACCENT   = "#01b8aa"
BI_BLUE     = "#4da6ff"
BI_GREEN    = "#107c10"
BI_ORANGE   = "#fd7e14"
BI_RED      = "#d13438"
BI_YELLOW   = "#ffd700"
BI_PURPLE   = "#8764b8"
BI_GRAY     = "#2d3748"
BI_BORDER   = "#2d3748"
BI_TEXT     = "#f3f2f1"
BI_SUBTEXT  = "#a0aec0"
BI_CHART_BG = "#1f2937"
BI_GRID     = "#2d3748"

st.markdown(f"""
<style>
    html, body, [class*="css"] {{
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: {BI_TEXT};
    }}
    .stApp {{ background-color: {BI_BG}; }}
    .main .block-container {{ padding-top: 1rem; max-width: 1400px; }}
    .metric-card {{
        background: {BI_CARD}; border: 1px solid {BI_BORDER};
        border-radius: 4px; padding: 18px 16px;
        text-align: left; margin: 4px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.3);
    }}
    .metric-card.green  {{ border-top: 3px solid {BI_ACCENT}; }}
    .metric-card.orange {{ border-top: 3px solid {BI_ORANGE}; }}
    .metric-card.red    {{ border-top: 3px solid {BI_RED}; }}
    .metric-card.blue   {{ border-top: 3px solid {BI_BLUE}; }}
    .metric-card.purple {{ border-top: 3px solid {BI_PURPLE}; }}
    .metric-value {{ font-size: 2.2em; font-weight: 700; color: {BI_TEXT}; line-height: 1.1; }}
    .metric-label {{ font-size: 0.75em; color: {BI_SUBTEXT}; margin-bottom: 6px;
                     text-transform: uppercase; letter-spacing: 0.6px; font-weight: 600; }}
    .metric-delta {{ font-size: 0.82em; margin-top: 8px; font-weight: 600; }}
    .section-header {{
        font-size: 1em; font-weight: 700; color: {BI_TEXT};
        margin: 16px 0 8px 0; padding: 8px 12px;
        background: {BI_CARD}; border-left: 3px solid {BI_ACCENT};
        border-radius: 0 4px 4px 0; text-transform: uppercase; letter-spacing: 0.5px;
    }}
    .report-title {{ font-size: 1.3em; font-weight: 700; color: {BI_TEXT}; margin-bottom: 2px; }}
    .report-subtitle {{ font-size: 0.85em; color: {BI_SUBTEXT}; margin-bottom: 16px; }}
    .gauge-label {{
        text-align: center; color: {BI_SUBTEXT}; font-size: 0.82em; font-weight: 600;
        margin-top: -6px; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px;
    }}
    section[data-testid="stSidebar"] {{
        background-color: #0d1117 !important;
        border-right: 1px solid {BI_BORDER};
        min-width: 242px !important;
        max-width: 242px !important;
        width: 242px !important;
    }}
    section[data-testid="stSidebar"] * {{ color: {BI_TEXT} !important; }}
    section[data-testid="stSidebar"] div[data-testid="stHorizontalBlock"] button {{
        opacity: 1 !important; position: static !important; height: auto !important;
        margin: 0 !important; padding: 6px 8px !important; border: 1px solid {BI_BORDER} !important;
    }}
    section[data-testid="stSidebar"] .stButton > button {{
        font-size: 0.75em !important;
        padding: 3px 8px !important;
        margin: 1px 0 !important;
    }}
    section[data-testid="stSidebar"] .stRadio label {{
        font-size: 0.82em !important;
        padding: 5px 8px !important;
        border-radius: 4px !important;
        cursor: pointer;
        display: block;
        width: 100%;
    }}
    section[data-testid="stSidebar"] .stRadio label:hover {{
        background: #1a2744 !important;
        color: {BI_ACCENT} !important;
    }}
    details[data-testid="stExpander"] {{
        background: #0d2137 !important;
        border: 1px solid {BI_ACCENT} !important;
        border-radius: 6px !important;
        margin-bottom: 12px;
    }}
    details[data-testid="stExpander"] summary {{
        color: {BI_ACCENT} !important;
        font-weight: 700 !important;
        font-size: 0.9em !important;
        padding: 8px 12px !important;
        background: #0d2137 !important;
        border-radius: 6px;
    }}
    details[data-testid="stExpander"] > div {{
        background: #0d2137 !important;
        padding: 8px 12px !important;
        border-top: 1px solid {BI_BORDER};
    }}
    .stButton > button {{
        background-color: transparent; color: {BI_TEXT};
        border: 1px solid {BI_BORDER}; border-radius: 6px;
        font-weight: 600; font-size: 0.85em; padding: 6px 12px;
    }}
    .stButton > button:hover {{
        background-color: {BI_ACCENT}; color: white; border-color: {BI_ACCENT};
    }}
    section[data-testid="stSidebar"] .stButton > button {{
        position: relative !important;
        top: -36px !important;
        height: 34px !important;
        margin-bottom: -34px !important;
        margin-top: 0 !important;
        padding: 0 !important;
        opacity: 0 !important;
        width: 100% !important;
        cursor: pointer !important;
        pointer-events: all !important;
        border: none !important;
    }}
    section[data-testid="stSidebar"] .stButton {{
        margin: 0 !important;
        padding: 0 !important;
    }}
    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] {{
        gap: 0 !important;
    }}
    section[data-testid="stSidebar"] div[data-testid="element-container"] {{
        margin: 0 !important;
        padding: 0 !important;
    }}
    hr {{ border-color: {BI_BORDER}; margin: 12px 0; }}
    .stDataFrame {{ border: 1px solid {BI_BORDER}; border-radius: 2px; }}
    div[data-baseweb="select"] label {{ color: white !important; }}
    div[data-baseweb="select"] p {{ color: white !important; }}
    div[data-testid="stSelectbox"] label {{ color: white !important; }}
    div[data-testid="stSelectbox"] p {{ color: white !important; }}
    div[data-testid="stSelectbox"] [data-testid="stWidgetLabel"] * {{ color: white !important; }}
    div[data-testid="stCheckbox"] label {{ color: white !important; }}
    div[data-testid="stCheckbox"] label p {{ color: white !important; }}
    div[data-testid="stCheckbox"] span {{ color: white !important; }}
    h4 {{ color: white !important; }}
    h3 {{ color: white !important; }}
    h2 {{ color: white !important; }}
    .stMarkdown p {{ color: white !important; }}
    .stMarkdown {{ color: white !important; }}
    summary {{ color: white !important; }}
    details[data-testid="stExpander"] [data-baseweb="select"] > div {{ background-color: #1a2744 !important; color: white !important; }}
    details[data-testid="stExpander"] [data-baseweb="select"] > div > div {{ background-color: #1a2744 !important; color: white !important; }}
    details[data-testid="stExpander"] [data-baseweb="select"] span {{ color: white !important; background-color: #1a2744 !important; }}
    details[data-testid="stExpander"] [data-baseweb="select"] div {{ background-color: #1a2744 !important; color: white !important; }}
    details[data-testid="stExpander"] [data-baseweb="select"] {{ background-color: #1a2744 !important; }}
.stSelectbox > div > div {{ background-color: #1a2744 !important; color: white !important; }}
    .stSelectbox div[data-baseweb="select"] {{ background-color: #1a2744 !important; }}
.stSelectbox > div > div {{ background-color: #1a2744 !important; color: white !important; }}
    .stSelectbox div[data-baseweb="select"] {{ background-color: #1a2744 !important; }}
[data-baseweb="select"] > div {{ background-color: #1a2744 !important; }}
    [data-baseweb="select"] > div > div {{ background-color: #1a2744 !important; color: white !important; }}
    [data-baseweb="input"] {{ background-color: #1a2744 !important; }}
    input {{ background-color: #1a2744 !important; color: white !important; }}
    [class*="ValueContainer"] {{ background-color: #1a2744 !important; }}
    [class*="control"] {{ background-color: #1a2744 !important; border-color: #2d3748 !important; }}
</style>
""", unsafe_allow_html=True)

# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_resource
def get_sheets_client():
    try:
        import json
        service_account_info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES_SHEETS)
    except:
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES_SHEETS)
    return gspread.authorize(creds)

@st.cache_data(ttl=300)
def load_sheet_data(tab_name):
    client = get_sheets_client()
    sheet = client.open_by_key(SHEET_ID)
    worksheet = sheet.worksheet(tab_name)
    data = worksheet.get_all_records()
    return pd.DataFrame(data)

@st.cache_data(ttl=300)
def load_weekly_membership():
    df = load_sheet_data("Weekly Membership")
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])
    num_cols = ["# Active members","# New members","# Suspended members",
                "# Cancelled members","Onboarding Members","Onboarding week","Age (Months)"]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df

@st.cache_data(ttl=300)
def load_vlookup():
    try:
        df = load_sheet_data("Vlookup")
        df = df.rename(columns={
            "Location": "Success Tutoring - Business name",
            "Location Start": "Location Start Date",
            "Months old": "Age (Months)",
            "Onboarding Week": "Onboarding week",
        })
        for c in ["Onboarding week","Onboarding Members","Age (Months)"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        return df
    except:
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
        tabs = [t.strip() for t in str(row["allowed_tabs"]).split(",")]
        gpm_filter = str(row.get("gpm_filter","")).strip()
        access_level = str(row.get("access_level","admin")).strip().lower()
        permissions[email] = {
            "tabs": tabs,
            "gpm_filter": gpm_filter,
            "access_level": access_level,
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
def get_allowed_tabs(user_email):
    permissions = load_permissions()
    entry = permissions.get(user_email.strip().lower())
    if not entry: return []
    return entry["tabs"]

def get_user_permissions(user_email):
    permissions = load_permissions()
    return permissions.get(user_email.strip().lower(), {
        "tabs": [], "gpm_filter": "", "access_level": "gpm"
    })

# ── Helpers ───────────────────────────────────────────────────────────────────
def churn_rate(cancelled, active):
    try:
        a = float(active)
        return round(float(cancelled) / a * 100, 1) if a > 0 else 0.0
    except:
        return 0.0

def metric_card(label, value, delta=None, color=""):
    delta_html = ""
    if delta is not None:
        arrow = "▲" if delta > 0 else "▼" if delta < 0 else "●"
        dcolor = BI_ACCENT if delta > 0 else BI_RED if delta < 0 else BI_SUBTEXT
        delta_html = f'<div class="metric-delta" style="color:{dcolor}">{arrow} {abs(delta):,.0f} vs prev week</div>'
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
PLOTLY_LAYOUT = dict(
    paper_bgcolor="#ffffff", plot_bgcolor="#f8f9fa",
    font=dict(color="#1a1a2e", family="Segoe UI, Helvetica Neue, Arial, sans-serif", size=12),
    xaxis=dict(gridcolor="#1e2d3d", gridwidth=1, tickfont=dict(color="#555555", size=10),
               linecolor="#dddddd", linewidth=1, showgrid=False, zeroline=False, tickangle=-30),
    yaxis=dict(gridcolor="#e8e8e8", gridwidth=1, tickfont=dict(color="#555555", size=10),
               linecolor="#dddddd", showgrid=True, zeroline=False),
    legend=dict(bgcolor="rgba(255,255,255,0.9)", bordercolor="#cccccc", borderwidth=1,
                font=dict(color="#1a1a2e", size=11), orientation="h",
                yanchor="bottom", y=-0.35, xanchor="center", x=0.5, itemsizing="constant"),
    hovermode="x unified",
    hoverlabel=dict(bgcolor="#1a2744", bordercolor=BI_ACCENT,
                    font=dict(color="#ffffff", size=12, family="Segoe UI"), namelength=-1),
    margin=dict(l=70, r=40, t=50, b=130), dragmode=False,
)

def std_traces(fig, df, date_col, col_name, color, label, secondary_y=False):
    try:
        r,g,b = int(color[1:3],16),int(color[3:5],16),int(color[5:7],16)
        fill_color = f"rgba({r},{g},{b},0.18)"
        glow_color = f"rgba({r},{g},{b},0.4)"
    except:
        fill_color = "rgba(1,184,170,0.18)"
        glow_color = "rgba(1,184,170,0.4)"
    kwargs = dict(secondary_y=secondary_y) if secondary_y is not False else {}
    fig.add_trace(go.Scatter(
        x=df[date_col], y=df[col_name], name=label, mode="lines",
        line=dict(color=glow_color, width=6),
        showlegend=False, hoverinfo="skip",
    ), **kwargs)
    fig.add_trace(go.Scatter(
        x=df[date_col], y=df[col_name], name=label,
        mode="lines+markers",
        line=dict(color=color, width=2.5, shape="spline", smoothing=0.4),
        marker=dict(color=color, size=6, symbol="circle",
                    line=dict(color="#0f1923", width=1.5)),
        fill="tozeroy", fillcolor=fill_color,
        hovertemplate=f"<b>{label}</b><br>%{{x|%d %b %Y}}<br><b>%{{y:,.1f}}</b><extra></extra>",
    ), **kwargs)

def std_layout(title, yaxis_title="", height=500):
    layout = dict(PLOTLY_LAYOUT)
    layout["title"] = dict(text=title,
                           font=dict(color="#1a1a2e", size=15, family="Segoe UI", weight="bold"),
                           x=0, xanchor="left", pad=dict(l=0))
    layout["yaxis"] = dict(PLOTLY_LAYOUT["yaxis"],
                           title=dict(text=yaxis_title, font=dict(color="#555555", size=11)))
    layout["height"] = height
    return layout

def plotly_line(weekly_df, date_col, series, title, yaxis_title="Members", height=500):
    fig = go.Figure()
    for col_name, color, label in series:
        if col_name not in weekly_df.columns: continue
        std_traces(fig, weekly_df, date_col, col_name, color, label)
    fig.update_layout(**std_layout(title, yaxis_title, height))
    return fig

def plotly_dual_axis(weekly_df, date_col, member_series, title, height=520):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    for col_name, color, label in member_series:
        if col_name not in weekly_df.columns: continue
        std_traces(fig, weekly_df, date_col, col_name, color, label, secondary_y=False)
    if "Churn Rate %" in weekly_df.columns:
        try:
            r,g,b = int(BI_RED[1:3],16),int(BI_RED[3:5],16),int(BI_RED[5:7],16)
            glow = f"rgba({r},{g},{b},0.4)"
        except:
            glow = "rgba(209,52,56,0.4)"
        fig.add_trace(go.Scatter(
            x=weekly_df[date_col], y=weekly_df["Churn Rate %"],
            name="Churn Rate %", mode="lines",
            line=dict(color=glow, width=6),
            showlegend=False, hoverinfo="skip",
        ), secondary_y=True)
        fig.add_trace(go.Scatter(
            x=weekly_df[date_col], y=weekly_df["Churn Rate %"],
            name="Churn Rate %", mode="lines+markers",
            line=dict(color=BI_RED, width=2.5, dash="dot", shape="spline", smoothing=0.4),
            marker=dict(color=BI_RED, size=6, symbol="diamond",
                        line=dict(color="#0f1923", width=1.5)),
            hovertemplate="<b>Churn Rate</b>: %{y:.1f}%<extra></extra>",
        ), secondary_y=True)
    layout = std_layout(title, "Members", height)
    layout["yaxis2"] = dict(
        title=dict(text="Churn Rate %", font=dict(color=BI_RED, size=11)),
        tickfont=dict(color=BI_RED), gridcolor="#e8e8e8", showgrid=False, zeroline=False)
    fig.update_layout(**layout)
    return fig
def apply_gpm_filter(df):
    """Silently filter data based on logged-in user's GPM and access level."""
    gpm_filter = st.session_state.get("gpm_filter", "")
    access_level = st.session_state.get("access_level", "admin")
    loc_col = "Success Tutoring - Business name"
    if access_level == "gpm" and gpm_filter:
        vl = load_vlookup()
        if not vl.empty and "GPM" in vl.columns:
            # Only show their locations, exclude Leasing stage
            allowed_locs = vl[
                (vl["GPM"] == gpm_filter) &
                (vl["Stage"] != "Leasing")
            ][loc_col].tolist()
            if loc_col in df.columns:
                df = df[df[loc_col].isin(allowed_locs)]
    return df
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
                   show_stage=True, show_gpm=True, show_location=False, show_status=True):
    loc_col = "Success Tutoring - Business name"
    with st.expander("🔍 Filter this report", expanded=True):
        cols = st.columns(5); i = 0
        if show_date and "Date" in df.columns:
            max_d = df["Date"].max(); min_d = df["Date"].min()
            with cols[i%5]:
                period = st.selectbox("📅 Period",
                    ["All Time","Last Month","Last Quarter","Last 6 Months","Last Year","Latest Week","Custom"],
                    index=0, key=f"{key_prefix}_date")
            i += 1
            if period=="Latest Week": df=df[df["Date"]==max_d]
            elif period=="Last Month": df=df[df["Date"]>=max_d-pd.DateOffset(months=1)]
            elif period=="Last Quarter": df=df[df["Date"]>=max_d-pd.DateOffset(months=3)]
            elif period=="Last 6 Months": df=df[df["Date"]>=max_d-pd.DateOffset(months=6)]
            elif period=="Last Year": df=df[df["Date"]>=max_d-pd.DateOffset(years=1)]
            elif period=="Custom":
                with cols[i%5]:
                    cr=st.date_input("Range",value=(min_d.date(),max_d.date()),
                                     min_value=min_d.date(),max_value=max_d.date(),
                                     key=f"{key_prefix}_custom")
                i+=1
                if isinstance(cr,tuple) and len(cr)==2:
                    df=df[(df["Date"].dt.date>=cr[0])&(df["Date"].dt.date<=cr[1])]
        if show_country and "Country" in df.columns:
            with cols[i%5]:
                opts=["All"]+sorted(df["Country"].dropna().unique().tolist())
                sel=st.selectbox("🌏 Country",opts,key=f"{key_prefix}_country")
            if sel!="All": df=df[df["Country"]==sel]
            i+=1
        if show_state and "Region" in df.columns:
            with cols[i%5]:
                opts=["All"]+sorted(df["Region"].dropna().unique().tolist())
                sel=st.selectbox("📍 State",opts,key=f"{key_prefix}_state")
            if sel!="All": df=df[df["Region"]==sel]
            i+=1
        if show_stage and "Stage" in df.columns:
            with cols[i%5]:
                opts=["All"]+sorted(df["Stage"].dropna().unique().tolist())
                sel=st.selectbox("📊 Stage",opts,key=f"{key_prefix}_stage")
            if sel!="All": df=df[df["Stage"]==sel]
            i+=1
        if show_gpm and "GPM" in df.columns:
            with cols[i%5]:
                opts=["All"]+sorted(df["GPM"].dropna().unique().tolist())
                sel=st.selectbox("👤 GPM",opts,key=f"{key_prefix}_gpm")
            if sel!="All": df=df[df["GPM"]==sel]
            i+=1
        if show_location and loc_col in df.columns:
            with cols[i%5]:
                opts=["All Locations"]+sorted(df[loc_col].dropna().unique().tolist())
                sel=st.selectbox("🏢 Location",opts,key=f"{key_prefix}_loc")
            if sel!="All Locations": df=df[df[loc_col]==sel]
            i+=1
        if show_status and "Status" in df.columns:
            with cols[i%5]:
                opts=["All"]+sorted(df["Status"].dropna().unique().tolist())
                sel=st.selectbox("🔵 Status",opts,key=f"{key_prefix}_status")
            if sel!="All": df=df[df["Status"]==sel]
            i+=1
    return df

def checkbox_date_filter(df, key_prefix=""):
    if "Date" not in df.columns or df.empty: return df
    all_dates = sorted(df["Date"].dropna().unique(),reverse=True)
    date_labels = [d.strftime("%d %b %Y") for d in all_dates]
    st.markdown(f"""<div style="background:#0d2137;border:1px solid {BI_ACCENT};
                border-radius:6px;padding:12px 16px;margin-bottom:12px">
        <span style="color:{BI_ACCENT};font-weight:700;font-size:0.9em">📅 Select Weeks to Display</span>
        <p style="color:{BI_SUBTEXT};font-size:0.8em;margin:4px 0 8px 0">
        Newest first. Default is latest 2 weeks.</p>
    </div>""", unsafe_allow_html=True)
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

def draw_per_location_trend(df_13m, metric_col, metric_label, color, key_prefix):
    loc_col = "Success Tutoring - Business name"
    if metric_col not in df_13m.columns:
        st.warning(f"Column '{metric_col}' not found."); return
    vl = load_vlookup()
    all_stages = vl["Stage"].dropna().unique().tolist() if not vl.empty and "Stage" in vl.columns else []
    stage_opts = ["All"]+sorted(all_stages)
    default_s = "Growth" if "Growth" in stage_opts else stage_opts[0]
    sel_s = st.selectbox("Filter by Stage:",stage_opts,
                          index=stage_opts.index(default_s),key=f"{key_prefix}_stage")
    df_plot = df_13m.copy()
    if sel_s!="All" and not vl.empty and "Stage" in vl.columns:
        stage_locs = vl[vl["Stage"]==sel_s][loc_col].tolist()
        df_plot = df_plot[df_plot[loc_col].isin(stage_locs)]
    if df_plot.empty: st.info(f"No data for {sel_s} stage."); return
    avg_df = df_plot.groupby("Date").apply(
        lambda x: x.groupby(loc_col)[metric_col].sum().mean()
    ).reset_index(); avg_df.columns=["Date","Avg"]; avg_df=avg_df.sort_values("Date")
    stage_lbl = f"— {sel_s}" if sel_s!="All" else "— All Stages"
    fig = go.Figure()
    std_traces(fig, avg_df, "Date", "Avg", color, f"Network Avg {stage_lbl}")
    show_indiv = st.checkbox("Show individual locations",value=False,key=f"{key_prefix}_indiv")
    if show_indiv:
        locs = sorted(df_plot[loc_col].dropna().unique().tolist())
        sel_locs = st.multiselect("Select locations:",locs,default=locs[:3],
                                   max_selections=8,key=f"{key_prefix}_locs")
        colors_indiv = [BI_BLUE,BI_ORANGE,BI_PURPLE,BI_RED,BI_YELLOW,"#00b4d8","#e91e8c","#33cc99"]
        for i,loc in enumerate(sel_locs):
            loc_df = df_plot[df_plot[loc_col]==loc].groupby("Date")[metric_col].sum().reset_index().sort_values("Date")
            loc_name = loc.replace("Success Tutoring - ","")
            std_traces(fig, loc_df, "Date", metric_col, colors_indiv[i%len(colors_indiv)], loc_name)
    fig.update_layout(**std_layout(
        f"Avg {metric_label} per Location {stage_lbl} — Last 13 Months",
        f"Avg {metric_label}", 420))
    st.plotly_chart(fig, use_container_width=True)

def latest_week_table(df_wm, df_filtered, metric_col, label_col, table_title):
    loc_col = "Success Tutoring - Business name"
    latest_wk = df_wm["Date"].max()
    df_tbl = df_wm[df_wm["Date"]==latest_wk].copy()
    for col in ["Country","Region","Stage","GPM","Status"]:
        if col in df_filtered.columns and col in df_tbl.columns:
            df_tbl = df_tbl[df_tbl[col].isin(df_filtered[col].unique())]
    df_tbl[metric_col] = pd.to_numeric(df_tbl[metric_col],errors="coerce").fillna(0)
    loc = df_tbl.groupby(loc_col)[metric_col].sum().reset_index()
    loc.columns=["Location",label_col]
    loc = loc.sort_values(label_col,ascending=False).reset_index(drop=True)
    st.markdown(f'<div class="section-header">{table_title} — Latest Week ({latest_wk.strftime("%d %b %Y")})</div>',
                unsafe_allow_html=True)
    st.dataframe(loc,use_container_width=True,hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 1 — Locations
# ══════════════════════════════════════════════════════════════════════════════
def report_locations(df_wm):
    st.markdown('<div class="report-title">1 · Campus Locations</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Vlookup — all location metadata with stage, GPM, country and region</div>', unsafe_allow_html=True)
    loc_col = "Success Tutoring - Business name"
    df_wm = apply_gpm_filter(df_wm)
    vl_full = load_vlookup()
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
        with st.expander("📊 Locations by Country & Stage",expanded=True):
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
            st.dataframe(merged[show_cols].rename(columns={loc_col:"Location"}).sort_values(
                "# Active members",ascending=False).reset_index(drop=True),
                use_container_width=True,hide_index=True)
        else:
            show_cols=[c for c in [loc_col,"Stage","GPM","Status","Country","Region","Age (Months)"] if c in vl.columns]
            st.dataframe(vl[show_cols].rename(columns={loc_col:"Location"}).reset_index(drop=True),
                         use_container_width=True,hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 2 — Membership
# ══════════════════════════════════════════════════════════════════════════════
def report_membership(df_wm):
    st.markdown('<div class="report-title">2 · Membership</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Weekly Membership — active, new, suspended & cancelled member trends</div>', unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)
    df_temp = report_filters(df_wm.copy(),key_prefix="r2",show_date=False,
                              show_country=True,show_state=True,show_stage=True,show_gpm=True,show_location=True,show_status=True)
    df = checkbox_date_filter(df_temp,key_prefix="r2")
    num_cols=["# Active members","# New members","# Suspended members","# Cancelled members"]
    for c in num_cols:
        if c in df.columns: df[c]=pd.to_numeric(df[c],errors="coerce").fillna(0)
    all_dates=sorted(df["Date"].dropna().unique())
    latest_date=all_dates[-1] if all_dates else None
    prev_date=all_dates[-2] if len(all_dates)>=2 else None
    def ws(d,col):
        if d is None or col not in df.columns: return 0
        return df[df["Date"]==d][col].sum()
    la=ws(latest_date,"# Active members"); pa=ws(prev_date,"# Active members")
    ln=ws(latest_date,"# New members");    pn=ws(prev_date,"# New members")
    ls=ws(latest_date,"# Suspended members"); ps=ws(prev_date,"# Suspended members")
    lc=ws(latest_date,"# Cancelled members"); pc=ws(prev_date,"# Cancelled members")
    col1,col2,col3,col4=st.columns(4)
    with col1: metric_card("Active Members",   f"{la:,.0f}",la-pa,"green")
    with col2: metric_card("New Members",      f"{ln:,.0f}",ln-pn,"blue")
    with col3: metric_card("Suspended Members",f"{ls:,.0f}",ls-ps,"orange")
    with col4: metric_card("Cancelled Members",f"{lc:,.0f}",lc-pc,"red")
    st.markdown("<br>",unsafe_allow_html=True)
    cr_latest=churn_rate(lc,la); cr_prev=churn_rate(pc,pa)
    col1,col2=st.columns([1,3])
    with col1: metric_card("Network Churn Rate",f"{cr_latest:.1f}%",round(cr_latest-cr_prev,2),"red")
    st.markdown('<div class="section-header">Member Trend — Last 13 Months</div>',unsafe_allow_html=True)
    df_13m=get_13m_filtered(df_wm,df)
    for c in num_cols:
        if c in df_13m.columns: df_13m[c]=pd.to_numeric(df_13m[c],errors="coerce").fillna(0)
    col1,col2,col3,col4=st.columns(4)
    show_a=col1.checkbox("Active",   value=True, key="r2_active")
    show_n=col2.checkbox("New",      value=False,key="r2_new")
    show_s=col3.checkbox("Suspended",value=False,key="r2_susp")
    show_c=col4.checkbox("Cancelled",value=True, key="r2_canc")
    selected=[]
    if show_a: selected.append(("# Active members",   BI_ACCENT,"Active Members"))
    if show_n: selected.append(("# New members",      BI_BLUE,  "New Members"))
    if show_s: selected.append(("# Suspended members",BI_ORANGE,"Suspended Members"))
    if show_c: selected.append(("# Cancelled members",BI_RED,   "Cancelled Members"))
    if selected:
        cols_plot=[c for c,_,_ in selected if c in df_13m.columns]
        weekly=df_13m.groupby("Date")[cols_plot].sum().reset_index().sort_values("Date")
        fig = plotly_line(weekly,"Date",selected,"Total Member Count — Last 13 Months",height=500)
        st.plotly_chart(fig,use_container_width=True)
    else:
        st.info("Please select at least one metric.")
    st.markdown('<div class="section-header">Avg Members per Location — Growth Stage Default</div>',unsafe_allow_html=True)
    vl=load_vlookup(); loc_col="Success Tutoring - Business name"
    all_stages_11=vl["Stage"].dropna().unique().tolist() if not vl.empty and "Stage" in vl.columns else []
    stage_opts_11=["All"]+sorted(all_stages_11)
    default_11="Growth" if "Growth" in stage_opts_11 else stage_opts_11[0]
    sel_stage_11=st.selectbox("Filter by Stage:",stage_opts_11,
                               index=stage_opts_11.index(default_11),key="r11_stage")
    df_11=df_13m.copy()
    if sel_stage_11!="All" and not vl.empty and "Stage" in vl.columns:
        stage_locs=vl[vl["Stage"]==sel_stage_11][loc_col].tolist()
        df_11=df_11[df_11[loc_col].isin(stage_locs)]
    col1,col2,col3,col4=st.columns(4)
    show_a2=col1.checkbox("Active",   value=True, key="r11_active")
    show_n2=col2.checkbox("New",      value=False,key="r11_new")
    show_s2=col3.checkbox("Suspended",value=False,key="r11_susp")
    show_c2=col4.checkbox("Cancelled",value=True, key="r11_canc")
    metrics_11=[]
    if show_a2: metrics_11.append(("# Active members",   BI_ACCENT,"Active"))
    if show_n2: metrics_11.append(("# New members",      BI_BLUE,  "New"))
    if show_s2: metrics_11.append(("# Suspended members",BI_ORANGE,"Suspended"))
    if show_c2: metrics_11.append(("# Cancelled members",BI_RED,   "Cancelled"))
    if metrics_11 and loc_col in df_11.columns and not df_11.empty:
        stage_lbl=f"— {sel_stage_11}" if sel_stage_11!="All" else "— All Stages"
        fig2 = go.Figure()
        for col_name,color,label in metrics_11:
            if col_name not in df_11.columns: continue
            avg_df=df_11.groupby("Date").apply(
                lambda x,c=col_name: x.groupby(loc_col)[c].sum().mean()
            ).reset_index(); avg_df.columns=["Date","Avg"]; avg_df=avg_df.sort_values("Date")
            std_traces(fig2, avg_df, "Date", "Avg", color, f"Avg {label}")
        fig2.update_layout(**std_layout(
            f"Avg Members per Location {stage_lbl} — Last 13 Months","Avg Members/Location",500))
        st.plotly_chart(fig2,use_container_width=True)
        show_indiv=st.checkbox("Show individual locations",value=False,key="r11_indiv")
        if show_indiv:
            locs=sorted(df_11[loc_col].dropna().unique().tolist())
            sel_locs=st.multiselect("Select locations:",locs,default=locs[:3],max_selections=8,key="r11_locs")
            if sel_locs:
                fig3=go.Figure()
                colors_indiv=[BI_ACCENT,BI_BLUE,BI_ORANGE,BI_PURPLE,BI_RED,BI_YELLOW,"#00b4d8","#e91e8c"]
                for i,loc in enumerate(sel_locs):
                    for col_name,color,label in metrics_11:
                        if col_name not in df_11.columns: continue
                        loc_df=df_11[df_11[loc_col]==loc].groupby("Date")[col_name].sum().reset_index().sort_values("Date")
                        loc_name=loc.replace("Success Tutoring - ","")
                        std_traces(fig3, loc_df, "Date", col_name, colors_indiv[i%len(colors_indiv)], f"{loc_name} ({label})")
                fig3.update_layout(**std_layout("Individual Location Trends","Members",500))
                st.plotly_chart(fig3,use_container_width=True)
    with st.expander("📋 Member Count by Location (Latest Week)",expanded=True):
        if loc_col in df.columns and latest_date is not None:
            latest_df=df[df["Date"]==latest_date].copy()
            if sel_stage_11!="All" and not vl.empty and "Stage" in vl.columns:
                stage_locs=vl[vl["Stage"]==sel_stage_11][loc_col].tolist()
                latest_df=latest_df[latest_df[loc_col].isin(stage_locs)]
            loc_tbl=latest_df.groupby(loc_col).agg(
                Active=("# Active members","sum"),New=("# New members","sum"),
                Suspended=("# Suspended members","sum"),Cancelled=("# Cancelled members","sum"),
            ).reset_index().rename(columns={loc_col:"Location"})
            loc_tbl["Churn Rate %"]=loc_tbl.apply(lambda r: churn_rate(r["Cancelled"],r["Active"]),axis=1)
            loc_tbl=loc_tbl.sort_values("Active",ascending=False).reset_index(drop=True)
            st.dataframe(loc_tbl,use_container_width=True,hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 3 — Membership by Age
# ══════════════════════════════════════════════════════════════════════════════
def report_age_combined(df_wm):
    st.markdown('<div class="report-title">3 · Membership by Age</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Age groups from Vlookup | Member counts from Weekly Membership</div>', unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)
    loc_col="Success Tutoring - Business name"
    vl=load_vlookup()
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
        with st.expander(f"📂 {grp} — {len(grp_df)} locations",expanded=False):
            show_cols=[c for c in [loc_col,"# Active members","# New members","# Suspended members","# Cancelled members","Age (Months)","Stage"] if c in grp_df.columns]
            st.dataframe(grp_df[show_cols].rename(columns={loc_col:"Location"}).sort_values("# Active members",ascending=False).reset_index(drop=True),
                         use_container_width=True,hide_index=True)
    st.markdown("<br>",unsafe_allow_html=True)
    st.markdown('<div class="section-header">Age Group Trends — Last 13 Months</div>',unsafe_allow_html=True)
    df_wm_f=report_filters(df_wm.copy(),key_prefix="r3wm",show_date=False,
                            show_country=True,show_state=True,show_stage=False,show_gpm=True,show_status=True)
    max_date=df_wm["Date"].max(); cutoff=max_date-pd.DateOffset(months=13)
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
            f"Avg {metric_name} by Location Age — Last 13 Months", f"Avg {metric_name}", 500))
        st.plotly_chart(fig_age,use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# Generic trend report helper
# ══════════════════════════════════════════════════════════════════════════════
def generic_member_report(df_wm, report_num, title, metric_col, metric_label, color, key):
    st.markdown(f'<div class="report-title">{report_num} · {title}</div>',unsafe_allow_html=True)
    st.markdown(f'<div class="report-subtitle">Source: Weekly Membership — {metric_label} trends and location ranking</div>',unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)
    if metric_col not in df_wm.columns: st.warning("Column not found."); return
    df=report_filters(df_wm.copy(),key_prefix=key,show_date=False,
                      show_country=True,show_state=True,show_stage=True,show_gpm=True,show_status=True)
    df_13m=get_13m_filtered(df_wm,df)
    df_13m[metric_col]=pd.to_numeric(df_13m[metric_col],errors="coerce").fillna(0)
    st.markdown(f'<div class="section-header">{metric_label} Trend — Last 13 Months</div>',unsafe_allow_html=True)
    weekly=df_13m.groupby("Date")[metric_col].sum().reset_index().sort_values("Date")
    fig = plotly_line(weekly,"Date",[(metric_col,color,metric_label)],
                      f"{metric_label} per Week — Last 13 Months",
                      yaxis_title=metric_label,height=500)
    st.plotly_chart(fig,use_container_width=True)
    st.markdown(f'<div class="section-header">Avg {metric_label} per Location — Growth Stage Default</div>',unsafe_allow_html=True)
    draw_per_location_trend(df_13m,metric_col,metric_label,color,f"{key}plt")
    latest_week_table(df_wm,df,metric_col,metric_label,f"{metric_label} by Location")

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 7 — Membership Churn
# ══════════════════════════════════════════════════════════════════════════════
def report_churn_combined(df_wm):
    st.markdown('<div class="report-title">7 · Membership Churn</div>',unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Weekly Membership — Churn = Cancelled / Active × 100</div>',unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)
    df=report_filters(df_wm.copy(),key_prefix="r6",show_date=False,
                      show_country=True,show_state=True,show_stage=False,show_gpm=True,show_location=True,show_status=True)
    loc_col="Success Tutoring - Business name"
    for c in ["# Active members","# Cancelled members","# Suspended members","# New members"]:
        if c in df.columns: df[c]=pd.to_numeric(df[c],errors="coerce").fillna(0)
    latest_wk_ch = df_wm["Date"].max()
    df_ch_tbl = df_wm[df_wm["Date"]==latest_wk_ch].copy()
    for col in ["Country","Region","Stage","GPM","Status"]:
        if col in df.columns and col in df_ch_tbl.columns:
            df_ch_tbl = df_ch_tbl[df_ch_tbl[col].isin(df[col].unique())]
    for c in ["# Active members","# Cancelled members","# Suspended members"]:
        if c in df_ch_tbl.columns:
            df_ch_tbl[c] = pd.to_numeric(df_ch_tbl[c],errors="coerce").fillna(0)
    churn_tbl = df_ch_tbl.groupby(loc_col).agg(
        Active=("# Active members","sum"),
        Suspended=("# Suspended members","sum"),
        Cancelled=("# Cancelled members","sum")
    ).reset_index().rename(columns={loc_col:"Location"})
    churn_tbl["Churn Rate %"] = churn_tbl.apply(
        lambda r: churn_rate(r["Cancelled"],r["Active"]), axis=1)
    churn_tbl = churn_tbl.sort_values("Churn Rate %",ascending=False).reset_index(drop=True)
    st.markdown(f'<div class="section-header">Churn Rate by Location — Latest Week ({latest_wk_ch.strftime("%d %b %Y")})</div>',
                unsafe_allow_html=True)
    st.dataframe(churn_tbl, use_container_width=True, hide_index=True)
    df_13m=get_13m_filtered(df_wm,df)
    for c in ["# Active members","# New members","# Suspended members","# Cancelled members"]:
        if c in df_13m.columns: df_13m[c]=pd.to_numeric(df_13m[c],errors="coerce").fillna(0)
    st.markdown('<div class="section-header">Membership & Churn Trend</div>',unsafe_allow_html=True)
    c1,c2,c3,c4,c5=st.columns(5)
    show_a=c1.checkbox("Active",      value=True, key="r6_active")
    show_n=c2.checkbox("New",         value=False,key="r6_new")
    show_s=c3.checkbox("Suspended",   value=False,key="r6_susp")
    show_c=c4.checkbox("Cancelled",   value=False,key="r6_canc")
    show_ch=c5.checkbox("Churn Rate%",value=True, key="r6_churn")
    weekly=df_13m.groupby("Date").agg(
        Active=("# Active members","sum"),New=("# New members","sum"),
        Suspended=("# Suspended members","sum"),Cancelled=("# Cancelled members","sum")
    ).reset_index().sort_values("Date")
    weekly["Churn Rate %"]=weekly.apply(lambda r: churn_rate(r["Cancelled"],r["Active"]),axis=1)
    has_members=any([show_a,show_n,show_s,show_c])
    if has_members and show_ch:
        member_series=[]
        if show_a: member_series.append(("Active",  BI_ACCENT,"Active"))
        if show_n: member_series.append(("New",     BI_BLUE,  "New"))
        if show_s: member_series.append(("Suspended",BI_ORANGE,"Suspended"))
        if show_c: member_series.append(("Cancelled",BI_RED,  "Cancelled"))
        fig = plotly_dual_axis(weekly,"Date",member_series,"Membership Trends & Churn Rate",height=500)
        st.plotly_chart(fig,use_container_width=True)
    st.markdown('<div class="section-header">Avg Churn Rate per Location — Growth Stage Default</div>',unsafe_allow_html=True)
    df_c=df_13m.copy()
    df_c["churn_pct"]=df_c.apply(lambda r: churn_rate(r.get("# Cancelled members",0),r.get("# Active members",0)),axis=1)
    vl=load_vlookup()
    all_stages_51=vl["Stage"].dropna().unique().tolist() if not vl.empty and "Stage" in vl.columns else []
    stage_opts_51=["All"]+sorted(all_stages_51)
    default_51="Growth" if "Growth" in stage_opts_51 else stage_opts_51[0]
    sel_stage_51=st.selectbox("Filter by Stage:",stage_opts_51,
                               index=stage_opts_51.index(default_51),key="r51_stage")
    if sel_stage_51!="All" and not vl.empty and "Stage" in vl.columns:
        stage_locs=vl[vl["Stage"]==sel_stage_51][loc_col].tolist()
        df_c=df_c[df_c[loc_col].isin(stage_locs)]
    if not df_c.empty and loc_col in df_c.columns:
        avg_churn=df_c.groupby("Date").apply(
            lambda x: x.groupby(loc_col)["churn_pct"].mean().mean()
        ).reset_index(); avg_churn.columns=["Date","Avg Churn %"]; avg_churn=avg_churn.sort_values("Date")
        stage_lbl=f"— {sel_stage_51}" if sel_stage_51!="All" else "— All Stages"
        show_indiv_ch = st.checkbox("Show individual locations", value=False, key="r6_churn_indiv")
        if show_indiv_ch:
            locs = sorted(df_c[loc_col].dropna().unique().tolist())
            sel_locs_ch = st.multiselect("Select locations:", locs, default=locs[:3],
                                          max_selections=8, key="r6_churn_locs")
            if sel_locs_ch:
                fig_ci = go.Figure()
                colors_indiv = [BI_ACCENT,BI_BLUE,BI_ORANGE,BI_PURPLE,BI_RED,BI_YELLOW,"#00b4d8","#e91e8c"]
                for i, loc in enumerate(sel_locs_ch):
                    loc_df = df_c[df_c[loc_col]==loc].groupby("Date")["churn_pct"].mean().reset_index().sort_values("Date")
                    loc_name = loc.replace("Success Tutoring - ","")
                    std_traces(fig_ci, loc_df, "Date", "churn_pct", colors_indiv[i%len(colors_indiv)], loc_name)
                fig_ci.update_layout(**std_layout(
                    "Churn Rate % by Individual Location — Last 13 Months","Churn Rate %",500))
                st.plotly_chart(fig_ci, use_container_width=True)
        fig_c=go.Figure()
        std_traces(fig_c, avg_churn, "Date", "Avg Churn %", BI_RED, "Avg Churn %")
        fig_c.update_layout(**std_layout(
            f"Avg Churn Rate per Location {stage_lbl} — Last 13 Months","Avg Churn Rate %",500))
        st.plotly_chart(fig_c,use_container_width=True)
# ══════════════════════════════════════════════════════════════════════════════
# REPORT 8 — Net Growth Rate %
# ══════════════════════════════════════════════════════════════════════════════
def report_net_growth(df_wm):
    st.markdown('<div class="report-title">8 · Net Growth Rate %</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Weekly Membership — Net Growth Rate % = (New − Cancelled) / Active × 100</div>', unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)

    df = report_filters(df_wm.copy(), key_prefix="r8ng_f", show_date=False,
                        show_country=True, show_state=True, show_stage=True,
                        show_gpm=True, show_status=True)
    loc_col = "Success Tutoring - Business name"

    for c in ["# Active members","# New members","# Cancelled members"]:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df_13m = get_13m_filtered(df_wm, df)
    for c in ["# Active members","# New members","# Cancelled members"]:
        if c in df_13m.columns: df_13m[c] = pd.to_numeric(df_13m[c], errors="coerce").fillna(0)

    def calc_ngr(row):
        try:
            a = float(row.get("# Active members", 0))
            return round((float(row.get("# New members", 0)) - float(row.get("# Cancelled members", 0))) / a * 100, 2) if a > 0 else 0.0
        except: return 0.0

    # ── Network Net Growth Rate % trend ──
    st.markdown('<div class="section-header">Net Growth Rate % — Last 13 Months (Total New − Total Cancelled) / Total Active</div>', unsafe_allow_html=True)
    weekly = df_13m.groupby("Date").agg(
        New=("# New members","sum"),
        Cancelled=("# Cancelled members","sum"),
        Active=("# Active members","sum")
    ).reset_index().sort_values("Date")
    weekly["Net Growth Rate %"] = weekly.apply(
        lambda r: round((r["New"] - r["Cancelled"]) / r["Active"] * 100, 2) if r["Active"] > 0 else 0.0, axis=1)

    fig = go.Figure()
    std_traces(fig, weekly, "Date", "Net Growth Rate %", BI_ACCENT, "Net Growth Rate %")
    fig.add_hline(y=0, line_dash="dash", line_color=BI_RED, line_width=1.5,
                  annotation_text="Break-even", annotation_position="bottom right")
    fig.update_layout(**std_layout("Network Net Growth Rate % — Last 13 Months (Total New − Total Cancelled) / Total Active", "Net Growth Rate %", 500))
    st.plotly_chart(fig, use_container_width=True)

    # ── Avg Net Growth Rate % per location ──
    st.markdown('<div class="section-header">Avg Net Growth Rate % per Location — Growth Stage Default (Equal weight per location)</div>', unsafe_allow_html=True)
    vl = load_vlookup()
    all_stages = vl["Stage"].dropna().unique().tolist() if not vl.empty and "Stage" in vl.columns else []
    stage_opts = ["All"] + sorted(all_stages)
    default_s = "Growth" if "Growth" in stage_opts else stage_opts[0]
    sel_s = st.selectbox("Filter by Stage:", stage_opts,
                          index=stage_opts.index(default_s), key="r8ng_stage")
    df_plot = df_13m.copy()
    if sel_s != "All" and not vl.empty and "Stage" in vl.columns:
        stage_locs = vl[vl["Stage"] == sel_s][loc_col].tolist()
        df_plot = df_plot[df_plot[loc_col].isin(stage_locs)]

    if not df_plot.empty and loc_col in df_plot.columns:
        df_plot["ngr"] = df_plot.apply(calc_ngr, axis=1)
        avg_ngr = df_plot.groupby("Date").apply(
            lambda x: x.groupby(loc_col)["ngr"].mean().mean()
        ).reset_index(); avg_ngr.columns = ["Date","Avg NGR %"]; avg_ngr = avg_ngr.sort_values("Date")
        stage_lbl = f"— {sel_s}" if sel_s != "All" else "— All Stages"

        fig2 = go.Figure()
        std_traces(fig2, avg_ngr, "Date", "Avg NGR %", BI_BLUE, f"Avg NGR % {stage_lbl}")
        fig2.add_hline(y=0, line_dash="dash", line_color=BI_RED, line_width=1.5,
                       annotation_text="Break-even", annotation_position="bottom right")
        fig2.update_layout(**std_layout(
            f"Avg Net Growth Rate % per Location {stage_lbl} — Last 13 Months", "Avg NGR %", 500))
        st.plotly_chart(fig2, use_container_width=True)

        show_indiv = st.checkbox("Show individual locations", value=False, key="r8ng_indiv")
        if show_indiv:
            locs = sorted(df_plot[loc_col].dropna().unique().tolist())
            sel_locs = st.multiselect("Select locations:", locs, default=locs[:3],
                                       max_selections=8, key="r8ng_locs")
            if sel_locs:
                fig3 = go.Figure()
                colors_indiv = [BI_ACCENT,BI_BLUE,BI_ORANGE,BI_PURPLE,BI_RED,BI_YELLOW,"#00b4d8","#e91e8c"]
                for i, loc in enumerate(sel_locs):
                    loc_df = df_plot[df_plot[loc_col]==loc].groupby("Date")["ngr"].mean().reset_index().sort_values("Date")
                    loc_name = loc.replace("Success Tutoring - ","")
                    std_traces(fig3, loc_df, "Date", "ngr", colors_indiv[i%len(colors_indiv)], loc_name)
                fig3.add_hline(y=0, line_dash="dash", line_color=BI_RED, line_width=1.5)
                fig3.update_layout(**std_layout("NGR % by Individual Location — Last 13 Months","NGR %",500))
                st.plotly_chart(fig3, use_container_width=True)

    # ── Latest week table ──
    latest_wk = df["Date"].max()
    df_tbl = df[df["Date"] == latest_wk].copy()
    for c in ["# Active members","# New members","# Cancelled members"]:
        if c in df_tbl.columns: df_tbl[c] = pd.to_numeric(df_tbl[c], errors="coerce").fillna(0)
    df_tbl["Net Growth Rate %"] = df_tbl.apply(calc_ngr, axis=1)
    loc_tbl = df_tbl.groupby(loc_col).agg(
        Active=("# Active members","sum"),
        New=("# New members","sum"),
        Cancelled=("# Cancelled members","sum")
    ).reset_index()
    loc_tbl["Net Growth Rate %"] = loc_tbl.apply(
        lambda r: round((r["New"] - r["Cancelled"]) / r["Active"] * 100, 2) if r["Active"] > 0 else 0.0, axis=1)
    loc_tbl = loc_tbl.rename(columns={loc_col:"Location"}).sort_values("Net Growth Rate %", ascending=False).reset_index(drop=True)
    st.markdown(f'<div class="section-header">Net Growth Rate % by Location — Latest Week ({latest_wk.strftime("%d %b %Y")})</div>', unsafe_allow_html=True)
    st.dataframe(loc_tbl, use_container_width=True, hide_index=True)
# ══════════════════════════════════════════════════════════════════════════════
# REPORT 9 — Onboarding Progress
# ══════════════════════════════════════════════════════════════════════════════
def report_onboarding(df_wm):
    st.markdown('<div class="report-title">8 · Onboarding Progress</div>',unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Vlookup — onboarding week progress and pre-sale member counts</div>',unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)
    vl=load_vlookup()
    if vl.empty: st.warning("Vlookup tab not available."); return
    vl_f=report_filters(vl.copy(),key_prefix="r7",show_date=False,
                         show_country=True,show_state=True,show_stage=True,show_gpm=True,show_status=True)
    loc_col="Success Tutoring - Business name"
    if "Stage" not in vl_f.columns: st.warning("Stage column not found."); return
    onb=vl_f[vl_f["Status"]=="Pre-Sale"].copy()
    if onb.empty: st.info("No locations in Pre-Sale status for selected filters."); return
    st.markdown('<div class="section-header">Onboarding Status by Week</div>',unsafe_allow_html=True)
    if "Onboarding week" in onb.columns:
        onb_w=onb.sort_values("Onboarding week")
        gpm_list=onb_w["GPM"].dropna().unique().tolist() if "GPM" in onb_w.columns else []
        gpm_cmap=plt.cm.get_cmap("tab10",max(len(gpm_list),1))
        gpm_color_map={g:gpm_cmap(i) for i,g in enumerate(gpm_list)}
        bar_colors=[gpm_color_map.get(str(row.get("GPM","")),BI_ACCENT) for _,row in onb_w.iterrows()]
        fig,ax=bi_fig(14,max(6,len(onb_w)*0.45))
        bars=ax.barh(onb_w[loc_col],onb_w["Onboarding week"],color=bar_colors,height=0.6)
        for bar,(_,row) in zip(bars,onb_w.iterrows()):
            ax.text(bar.get_width()+0.1,bar.get_y()+bar.get_height()/2,
                    str(int(row["Onboarding week"])),va="center",color=BI_TEXT,fontsize=8)
        milestones={"Pre Sale":2,"Fit Out":5,"Assessment":12,"50 Members":15,"Grand Opening":22}
        m_colors={"Pre Sale":BI_BLUE,"Fit Out":BI_ORANGE,"Assessment":BI_BLUE,"50 Members":BI_RED,"Grand Opening":BI_ACCENT}
        for name,week in milestones.items():
            ax.axvline(x=week,color=m_colors[name],linestyle="--",linewidth=1.2,alpha=0.9)
            ax.text(week+0.1,len(onb_w)-0.5,name,color=m_colors[name],fontsize=7,rotation=90,va="top")
        ax.set_xlabel("Onboarding Week",color=BI_TEXT,fontsize=9)
        ax.set_title("Onboarding Status by Week",color=BI_TEXT,fontsize=11,fontweight="bold")
        ax.invert_yaxis()
        if gpm_list:
            patches=[mpatches.Patch(color=gpm_color_map[g],label=g) for g in gpm_list]
            ax.legend(handles=patches,loc="center right",bbox_to_anchor=(1.18,0.5),
                      facecolor=BI_CARD,edgecolor=BI_BORDER,labelcolor=BI_TEXT,fontsize=10)
        plt.tight_layout(); st.pyplot(fig); plt.close()
    st.markdown("<br>",unsafe_allow_html=True)
    st.markdown('<div class="section-header">Pre-Sale Members by Location</div>',unsafe_allow_html=True)
    if "Onboarding Members" in onb.columns:
        onb_m=onb.sort_values("Onboarding Members")
        country_colors={"Australia":BI_ACCENT,"New Zealand":BI_BLUE,"Canada":BI_RED}
        bar_colors=[country_colors.get(str(row.get("Country","")),BI_GRAY) for _,row in onb_m.iterrows()]
        fig,ax=bi_fig(14,max(6,len(onb_m)*0.45))
        bars=ax.barh(onb_m[loc_col],onb_m["Onboarding Members"],color=bar_colors,height=0.6)
        for bar,(_,row) in zip(bars,onb_m.iterrows()):
            ax.text(bar.get_width()+0.2,bar.get_y()+bar.get_height()/2,
                    str(int(row["Onboarding Members"])),va="center",color=BI_TEXT,fontsize=8)
        ax.axvline(x=50,color=BI_RED,linestyle="--",linewidth=1.5)
        ax.text(50.5,len(onb_m)-0.5,"Target (50)",color=BI_RED,fontsize=8,rotation=90,va="top")
        ax.set_xlabel("Total Presale Members",color=BI_TEXT,fontsize=9)
        ax.set_title("Pre-Sale Members by Location",color=BI_TEXT,fontsize=11,fontweight="bold")
        ax.invert_yaxis()
        patches=[mpatches.Patch(color=v,label=k) for k,v in country_colors.items()]
        ax.legend(handles=patches,loc="center right",bbox_to_anchor=(1.18,0.5),
                  facecolor=BI_CARD,edgecolor=BI_BORDER,labelcolor=BI_TEXT,fontsize=10)
        plt.tight_layout(); st.pyplot(fig); plt.close()
    with st.expander("📋 Onboarding Location Detail",expanded=False):
        show_cols=[c for c in [loc_col,"Onboarding week","Onboarding Members","GPM","Country","Region","Age (Months)"] if c in onb.columns]
        st.dataframe(onb[show_cols].sort_values("Onboarding week").reset_index(drop=True),
                     use_container_width=True,hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 10 — AI Data Analysis
# ══════════════════════════════════════════════════════════════════════════════
def report_claude_outliers(df_wm):
    st.markdown('<div class="report-title">10 · AI Data Analysis</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: Weekly Membership — latest week performance tables and network outliers</div>', unsafe_allow_html=True)
    df_wm = apply_gpm_filter(df_wm)

    df = report_filters(df_wm.copy(), key_prefix="r10", show_date=False,
                        show_country=True, show_state=True, show_stage=True,
                        show_gpm=True, show_status=True)
    loc_col = "Success Tutoring - Business name"

    for c in ["# Active members","# New members","# Cancelled members","# Suspended members"]:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    all_dates = sorted(df["Date"].unique())
    if not all_dates: st.warning("No data."); return
    latest_date = all_dates[-1]
    latest_week = df[df["Date"] == latest_date].copy()

    loc_latest = latest_week.groupby(loc_col).agg(
        Active=("# Active members","sum"),
        New=("# New members","sum"),
        Cancelled=("# Cancelled members","sum"),
        Suspended=("# Suspended members","sum"),
    ).reset_index()

    loc_latest["Churn Rate %"] = loc_latest.apply(
        lambda r: round(r["Cancelled"] / r["Active"] * 100, 1) if r["Active"] > 0 else 0.0, axis=1)
    loc_latest["Net Growth Rate %"] = loc_latest.apply(
        lambda r: round((r["New"] - r["Cancelled"]) / r["Active"] * 100, 2) if r["Active"] > 0 else 0.0, axis=1)

    vl = load_vlookup()
    if not vl.empty and "Stage" in vl.columns:
        loc_latest = loc_latest.merge(vl[[loc_col,"Stage","GPM","Country","Region"]],
                                       on=loc_col, how="left")

    def show_ranked(df_sub, sort_col, title, n=5, ascending=False):
        st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)
        other_cols = [c for c in [loc_col,"Active","New","Cancelled","Churn Rate %",
                                  "Net Growth Rate %","Stage"] if c in df_sub.columns and c != sort_col]
        cols_show = [loc_col, sort_col] + [c for c in other_cols if c != loc_col]
        cols_show = [c for c in cols_show if c in df_sub.columns]
        ranked = df_sub.sort_values(sort_col, ascending=ascending).head(n).reset_index(drop=True)
        ranked.index = ranked.index + 1
        st.dataframe(ranked[cols_show].rename(columns={loc_col:"Location"}),
                     use_container_width=True)

    def show_all(df_sub, title):
        st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)
        cols_show = [c for c in [loc_col,"Active","New","Cancelled","Churn Rate %",
                                  "Net Growth Rate %","Stage"] if c in df_sub.columns]
        st.dataframe(df_sub[cols_show].rename(columns={loc_col:"Location"}).reset_index(drop=True),
                     use_container_width=True, hide_index=True)

    def show_outliers(df_sub, col, title):
        st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)
        mean = df_sub[col].mean(); std = df_sub[col].std()
        outliers = df_sub[(df_sub[col] > mean + 2*std) | (df_sub[col] < mean - 2*std)].copy()
        outliers["vs Mean"] = (outliers[col] - mean).round(2)
        cols_show = [c for c in [loc_col, col, "vs Mean", "Active", "Stage"] if c in outliers.columns]
        st.markdown(f'<span style="color:{BI_SUBTEXT};font-size:0.82em">Mean: {mean:.1f} | Std Dev: {std:.1f} | Threshold: >{mean+2*std:.1f} or <{mean-2*std:.1f}</span>',
                    unsafe_allow_html=True)
        if outliers.empty:
            st.caption("No locations outside 2σ threshold.")
        else:
            st.dataframe(outliers[cols_show].rename(columns={loc_col:"Location"}).sort_values(col, ascending=False).reset_index(drop=True),
                         use_container_width=True, hide_index=True)

    st.markdown(f'<div style="background:{BI_CARD};border:1px solid {BI_BORDER};border-radius:4px;padding:10px 14px;margin-bottom:16px;font-size:0.88em">'
                f'<b>Latest Week: {latest_date.strftime("%d %b %Y")}</b> &nbsp;|&nbsp; '
                f'{len(loc_latest)} locations &nbsp;|&nbsp; '
                f'Avg Active: <b>{loc_latest["Active"].mean():.0f}</b> &nbsp;|&nbsp; '
                f'Avg NGR: <b>{loc_latest["Net Growth Rate %"].mean():.1f}%</b> &nbsp;|&nbsp; '
                f'Avg Churn: <b>{loc_latest["Churn Rate %"].mean():.1f}%</b>'
                f'</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        show_ranked(loc_latest, "Active",          "🥇 Top 5 — Active Members",        n=5, ascending=False)
        show_ranked(loc_latest, "New",             "🥇 Top 5 — New Members",           n=5, ascending=False)
        show_ranked(loc_latest, "Net Growth Rate %","🥇 Top 5 — Net Growth Rate %",    n=5, ascending=False)
        show_all(loc_latest[loc_latest["Active"] < 50].sort_values("Active"),
                 "⚠️ All Locations — Active Members Below 50")

    with col2:
        show_ranked(loc_latest, "Churn Rate %",    "🔴 Top 5 — Highest Churn Rate %",  n=5, ascending=False)
        show_ranked(loc_latest, "Net Growth Rate %","🔴 Bottom 5 — Lowest Net Growth Rate %",      n=5, ascending=True)
        show_outliers(loc_latest, "New",            "📊 Outliers — New Members (2σ)")
        show_outliers(loc_latest, "Net Growth Rate %","📊 Outliers — Net Growth Rate % (2σ)")
        show_outliers(loc_latest, "Churn Rate %",   "📊 Outliers — Churn Rate % (2σ)")

    if not ANTHROPIC_KEY:
        st.warning("No Anthropic API key found — AI narrative unavailable.")
        return

    if st.button("🤖 Generate Claude AI Narrative", use_container_width=True):
        def tbl(df_sub, sort_col, n=5, asc=False):
            cols = [c for c in [loc_col,"Active","New","Churn Rate %","Net Growth Rate %","Stage"] if c in df_sub.columns]
            return df_sub.sort_values(sort_col, ascending=asc).head(n)[cols].rename(
                columns={loc_col:"Location"}).to_string(index=False)

        prompt = f"""You are a sharp business analyst for Success Tutoring, an Australian tutoring franchise.
Use Australian English. Data is for the latest week only: {latest_date.strftime('%d %b %Y')}.

NETWORK SUMMARY:
Locations: {len(loc_latest)} | Avg Active: {loc_latest['Active'].mean():.0f} | Avg New: {loc_latest['New'].mean():.1f} | Avg Churn: {loc_latest['Churn Rate %'].mean():.1f}% | Avg NGR: {loc_latest['Net Growth Rate %'].mean():.1f}%

TOP 5 ACTIVE: {tbl(loc_latest,'Active')}
TOP 5 NEW: {tbl(loc_latest,'New')}
TOP 5 NGR: {tbl(loc_latest,'Net Growth Rate %')}
LOW ACTIVE (<50): {loc_latest[loc_latest['Active']<50][['Active','Stage']].to_string()}
TOP 5 CHURN: {tbl(loc_latest,'Churn Rate %')}
BOTTOM 5 NGR: {tbl(loc_latest,'Net Growth Rate %',asc=True)}

Provide a concise narrative covering:
## 🌟 Strong Performers
## ⚠️ Locations Needing Attention
## 🔴 Critical — Below 50 Members
## 💡 Leadership Recommendations — 4 actions this week"""

        client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
        with st.spinner("🤖 Claude is analysing..."):
            message = client.messages.create(
                model="claude-sonnet-4-20250514", max_tokens=1500,
                messages=[{"role":"user","content":prompt}])
        st.markdown(message.content[0].text)

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
            st.markdown(f'<div style="color:{BI_ACCENT};font-size:3em;text-align:center">📊</div>',
                        unsafe_allow_html=True)
        st.markdown(f'<h2 style="color:{BI_TEXT};text-align:center;margin-top:8px">Success Tutoring Dashboard</h2>',
                    unsafe_allow_html=True)
        st.markdown(f'<p style="color:{BI_SUBTEXT};text-align:center;margin-bottom:24px">Sign in to continue</p>',
                    unsafe_allow_html=True)
        import urllib.parse
        CLIENT_ID = "507985856717-srjmjg07sdde13anpr20io14ln46n9sf.apps.googleusercontent.com"
        CLIENT_SECRET = "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"
        REDIRECT_URI = "https://j7ky6kl5hwlbrjpxtuk8ce.streamlit.app"
        AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
        TOKEN_URL = "https://oauth2.googleapis.com/token"
        USERINFO_URL = "https://www.googleapis.com/oauth2/v3/userinfo"
        params = st.query_params
        if "code" in params and not st.session_state.get("logged_in"):
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
                    perms = get_user_permissions(email)
                    if perms["tabs"]:
                        st.query_params.clear()
                        st.session_state["logged_in"] = True
                        st.session_state["user_email"] = email
                        st.session_state["user_name"] = name
                        st.session_state["access_level"] = perms["access_level"]
                        st.session_state["gpm_filter"] = perms["gpm_filter"]
                        log_access(email, name, "Login")
                        st.rerun()
                    else:
                        st.query_params.clear()
                        flag_unknown_user(email)
                        st.error(f"⛔ {email} is not authorised. Contact your administrator.")
                else:
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
                <a href="{google_auth_url}" target="_self" style="
                    display:inline-flex;align-items:center;gap:10px;
                    background:white;color:#1a1a2e;
                    padding:12px 24px;border-radius:6px;
                    font-weight:600;font-size:0.95em;
                    text-decoration:none;
                    border:1px solid #dddddd;
                    box-shadow:0 2px 8px rgba(0,0,0,0.15);
                ">
                <img src="https://www.google.com/favicon.ico" width="20" height="20"/>
                Sign in with Google
                </a>
            </div>
            """, unsafe_allow_html=True)
            st.markdown(f'<p style="color:{BI_SUBTEXT};text-align:center;font-size:0.8em;margin-top:4px">— or use email below —</p>',
                        unsafe_allow_html=True)
        with st.expander("🔐 Sign in with email instead", expanded=False):
            email_input = st.text_input("Email",placeholder="you@successtutoring.com",
                                         label_visibility="collapsed", key="email_fallback")
            if st.button("Login with Email", use_container_width=True):
                if email_input and "@successtutoring.com" in email_input:
                    allowed = get_allowed_tabs(email_input)
                    if allowed:
                        st.session_state["logged_in"] = True
                        st.session_state["user_email"] = email_input.lower().strip()
                        st.session_state["user_name"] = email_input.split("@")[0].replace("."," ").title()
                        st.empty()
                        st.rerun()
                    else:
                        st.error("⛔ Your email is not authorised.")
                else:
                    st.error("Please enter a valid @successtutoring.com email.")
        st.caption("🔒 Only approved team members can access this dashboard.")

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
    "4 · New Members",
    "5 · Suspended Members",
    "6 · Cancelled Members",
    "7 · Membership Churn",
    "8 · Net Growth Rate %",
    "9 · Onboarding Progress",
    "10 · AI Outlier Analysis",
]

with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png",width=150)
    else:
        st.markdown(f'<div style="color:{BI_ACCENT};font-size:1.1em;font-weight:700;padding:8px 0">📊 Success Tutoring</div>',
                    unsafe_allow_html=True)
    st.markdown(f'<div style="color:#718096;font-size:0.68em;margin-bottom:8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{user_email}</div>', unsafe_allow_html=True)
    st.markdown("---")
    btn_col1,btn_col2=st.columns(2)
    with btn_col1:
        if st.button("🔄 Refresh",use_container_width=True):
            st.cache_data.clear(); st.success("✅ Refreshed!"); st.rerun()
    with btn_col2:
        if st.button("🚪 Logout",use_container_width=True):
            log_access(
                st.session_state.get("user_email",""),
                st.session_state.get("user_name",""),
                "Logout"
            )
            for key in ["connected","email","name","oauth_token"]:
                st.session_state.pop(key, None)
            st.session_state.clear()
            st.rerun()
    st.markdown("---")
    st.markdown(f'<div style="color:{BI_ACCENT};font-size:0.72em;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;padding:4px 0 6px 2px">📊 Reports</div>',
                unsafe_allow_html=True)
    if "selected_report" not in st.session_state:
        st.session_state["selected_report"] = REPORTS[0]
    for r in REPORTS:
        is_active = st.session_state["selected_report"] == r
        bg = f"rgba(255,215,0,0.12)" if is_active else "transparent"
        border = f"#ffd700" if is_active else "transparent"
        color = "#ffd700" if is_active else "#a0aec0"
        weight = "700" if is_active else "400"
        st.markdown(f"""
            <div onclick="" style="
                background:{bg};
                border-left: 3px solid {border};
                border-radius: 0 6px 6px 0;
                padding: 7px 10px;
                margin: 0;
                cursor: pointer;
                font-size: 0.85em;
                font-weight: {weight};
                color: {color};
                display: flex;
                align-items: center;
                gap: 8px;
                line-height: 1.2;
            ">{"▶ " if is_active else "　"}{r}</div>
        """, unsafe_allow_html=True)
        if st.button(r, key=f"nav_{r}", use_container_width=True):
            st.session_state["selected_report"] = r
            st.rerun()
    selected_report = st.session_state["selected_report"]

selected_report = st.session_state.get("selected_report", REPORTS[0])

hdr_col1, hdr_col2 = st.columns([1, 8])
with hdr_col1:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=110)
with hdr_col2:
    st.markdown(f'<h2 style="color:{BI_TEXT};font-weight:700;margin-bottom:2px;margin-top:8px">Success Tutoring Dashboard</h2>',
                unsafe_allow_html=True)
    st.markdown(f'<p style="color:{BI_SUBTEXT};margin-bottom:16px;font-size:0.9em">Business intelligence and performance analytics</p>',
                unsafe_allow_html=True)

with st.spinner("Loading data..."):
    try:
        df_wm=load_weekly_membership()
    except Exception as e:
        st.error(f"Could not load Weekly Membership: {e}"); st.stop()

if df_wm.empty:
    st.warning("Weekly Membership sheet is empty."); st.stop()

if selected_report=="1 · Campus Locations":           report_locations(df_wm)
elif selected_report=="2 · Membership":               report_membership(df_wm)
elif selected_report=="3 · Membership by Age":        report_age_combined(df_wm)
elif selected_report=="4 · New Members":
    generic_member_report(df_wm,"4","New Members","# New members","New Members",BI_ACCENT,"r4")
elif selected_report=="5 · Suspended Members":
    generic_member_report(df_wm,"5","Suspended Members","# Suspended members","Suspended Members",BI_ORANGE,"r5")
elif selected_report=="6 · Cancelled Members":
    generic_member_report(df_wm,"6","Cancelled Members","# Cancelled members","Cancelled Members",BI_RED,"r6")
elif selected_report=="7 · Membership Churn":         report_churn_combined(df_wm)
elif selected_report=="8 · Net Growth Rate %":          report_net_growth(df_wm)
elif selected_report=="9 · Onboarding Progress":      report_onboarding(df_wm)
elif selected_report=="10 · AI Outlier Analysis":     report_claude_outliers(df_wm)
