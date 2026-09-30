import base64
import io
import re
from pathlib import Path
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from PV_Blink_Login_Form import login_with_api


COMPANY_NAMES = ("PV-Blink Inverter", "Banga Solar Pvt Ltd")
BANGA_COLORS = {
    "#f97316": "#166534", "#ea580c": "#14532d",
    "#fb923c": "#228b4e", "#fdba74": "#86c99b",
    "#fed7aa": "#bbdfc5", "#ffedd5": "#dcf0e2",
    "#fff7ed": "#f0f8f2", "#fff8f1": "#f4faf5",
}


def company_theme(text):
    """Swap brand accents only; preserve semantic chart colors and data."""
    if st.session_state.get("dashboard_company") != "Banga Solar Pvt Ltd":
        return text
    text = re.sub(r"#[0-9a-fA-F]{6}", lambda m: BANGA_COLORS.get(m[0].lower(), m[0]), text)
    for source, target in [("249,115,22", "22,101,52"),
                           ("234,88,12", "20,83,45"),
                           ("253,186,116", "134,201,155")]:
        pattern = r"\s*,\s*".join(source.split(","))
        text = re.sub(r"(?<=\()" + pattern + r"(?=\s*,|\))", target, text)
    return text


def reset_company_data():
    """Require fresh uploads when switching companies; retain authentication."""
    for key in list(st.session_state):
        if key != "dashboard_company" and not key.startswith("pv_"):
            del st.session_state[key]


# ============================================================
# PV-BLINK ENTERPRISE BI — V11 DATA ACCURACY + UI FIXl
# ------------------------------------------------------------
# 1) NORMAL SALES EXCEL = Gross Sales / Gross QTY source.
# 2) SALES RETURN REGISTER = independent return source with:
#    Ledger Name, Order No#, Sales Return Date, Item Group,
#    Item Name, QTY, Amount. Rate is NOT required.
#    Net Sales = Gross Sales - Return Amount
#    Net QTY   = Gross QTY   - Return QTY
# 3) DASHBOARD / KPI-KRA EXCEL = Emp_Name field activity only.
#    It is NEVER deducted from Sales. Supports Visit, KM, Lead,
#    Follow-Ups, Quotation, PI, Sales-Order, Wrong KM Claim.
# 4) SERVICE = independent maintenance/ticket intelligence.
# 5) FUTURE DESIGNATIONS = when a Designation column appears in
#    Dashboard data, designation analytics activate automatically.
# 6) DATE FILTERS are independent and use inclusive end dates:
#    Sales Invoice Date, Return Date, Field Date and Service Date.
# ============================================================

# ============================================================
# PV-BLINK ENTERPRISE SALES & SERVICE INTELLIGENCE DASHBOARD
# Official Brand Theme: PVblink Technology Pvt. Ltd. (pvblink.com)
# Layout: Executive Light Theme, clean labels, four service charts, and compact controls.
# ============================================================


st.set_page_config(
    page_title="PV-BLINK | ANALYTICS",
    page_icon=str(Path(__file__).parent / "assets" / "pvblink_logo.png"),
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# LOGIN GATE — shown BEFORE the dashboard
# ============================================================
PV_LOGO_PATH = Path(__file__).parent / "assets" / "pvblink_logo.png"


@st.cache_data(show_spinner=False)
def pv_logo_data_uri():
    logo_bytes = PV_LOGO_PATH.read_bytes()
    encoded = base64.b64encode(logo_bytes).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def pv_login_gate():
    if st.session_state.get("pv_authenticated", False):
        return

    st.markdown(
        company_theme("""
        <style>
            [data-testid="stSidebar"] {
                display: none !important;
            }

            header[data-testid="stHeader"] {
                background: transparent !important;
            }

            .block-container {
                max-width: 520px !important;
                padding-top: 8vh !important;
                padding-bottom: 8vh !important;
            }

            .pv-login-logo {
                text-align: center;
                margin: 0 auto 22px auto;
                width: min(270px, 76vw);
            }

            .pv-login-logo img {
                display: block;
                width: 100%;
                height: auto;
            }

            div[data-testid="stForm"] {
                background: #ffffff;
                border: 1px solid rgba(253, 186, 116, .8);
                border-radius: 22px;
                padding: 28px 28px 24px 28px;
                box-shadow: 0 24px 70px rgba(0, 0, 0, .32);
            }

            div[data-testid="stForm"] label p,
            div[data-testid="stForm"] label {
                color: #334155 !important;
                font-size: 12px !important;
                font-weight: 800 !important;
            }

            div[data-testid="stForm"] input {
                min-height: 44px !important;
                background: #f8fafc !important;
                color: #0f172a !important;
                border: 1px solid #cbd5e1 !important;
                border-radius: 11px !important;
            }

            div[data-testid="stForm"] input:focus {
                border-color: #f97316 !important;
                box-shadow: 0 0 0 3px rgba(249, 115, 22, .14) !important;
            }

            div[data-testid="stForm"] [data-testid="stCheckbox"] {
                margin: 2px 0 12px 0;
            }

            div[data-testid="stForm"] [data-testid="stCheckbox"] label p {
                color: #64748b !important;
                font-size: 11px !important;
                font-weight: 650 !important;
            }

            div[data-testid="stFormSubmitButton"] button {
                width: 100%;
                background: linear-gradient(135deg, #fb923c, #ea580c) !important;
                color: white !important;
                border: none !important;
                border-radius: 11px !important;
                font-weight: 800 !important;
                min-height: 46px !important;
                box-shadow: 0 8px 18px rgba(234, 88, 12, .22) !important;
            }

            div[data-testid="stFormSubmitButton"] button:hover {
                filter: brightness(1.06);
                transform: translateY(-1px);
            }

            @media (max-width: 600px) {
                .block-container {
                    padding-top: 6vh !important;
                    padding-left: 16px !important;
                    padding-right: 16px !important;
                }

                div[data-testid="stForm"] {
                    padding: 22px 18px 18px 18px;
                }
            }
        </style>
        """),
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div class="pv-login-logo"><img src="{pv_logo_data_uri()}" alt="PV blink inverter"></div>',
        unsafe_allow_html=True,
    )

    remembered_email = st.session_state.get("pv_remembered_email", "")
    if remembered_email and "pv_login_username" not in st.session_state:
        st.session_state["pv_login_username"] = remembered_email

    with st.form("pv_login_form", clear_on_submit=False):
        username = st.text_input(
            "Email",
            placeholder="Enter email",
            autocomplete="username",
            key="pv_login_username",
        )
        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter password",
            autocomplete="current-password",
            key="pv_login_password",
        )
        remember_me = st.checkbox(
            "Remember me",
            key="pv_remember_me",
            help=(
                "Your email is remembered by this app. Use your browser's "
                "password manager to save the password securely."
            ),
        )

        submitted = st.form_submit_button(
            "Login",
            width="stretch",
        )

    if submitted:
        if remember_me:
            st.session_state["pv_remembered_email"] = str(username).strip()
        else:
            st.session_state.pop("pv_remembered_email", None)

        with st.spinner("Signing in..."):
            ok, message, response_data = login_with_api(
                str(username).strip(),
                str(password),
            )

        if ok:
            api_data = response_data.get("data", {})
            display_name = " ".join(
                part for part in [
                    api_data.get("firstName", ""),
                    api_data.get("lastName", ""),
                ] if part
            ) or api_data.get("email", str(username).strip())
            st.session_state["pv_authenticated"] = True
            st.session_state["pv_authenticated_user"] = display_name
            st.session_state["pv_auth_token"] = api_data.get("accessToken", "")
            st.session_state["pv_user"] = api_data
            st.session_state["pv_login_response"] = response_data
            st.rerun()
        else:
            st.error(message)

    st.stop()


pv_login_gate()

# Upgrade sessions created before the full company name was introduced.
if st.session_state.get("dashboard_company") == "Banga Solar":
    st.session_state["dashboard_company"] = "Banga Solar Pvt Ltd"
with st.container(key="company_header"):
    header_brand, header_company = st.columns([3, 1], vertical_alignment="center")
with header_company:
    company = st.selectbox(
        "COMPANY WORKSPACE", COMPANY_NAMES, key="dashboard_company",
        on_change=reset_company_data,
        help="Upload the selected company's files. Switching companies clears current uploads and filters.",
    )
company_title = company
company_export = "BangaSolar" if company == "Banga Solar Pvt Ltd" else "PVBlink"
st.set_page_config(page_title=f"{company_title} | ANALYTICS", page_icon="🌞" if company == "Banga Solar Pvt Ltd" else str(PV_LOGO_PATH))
if company == "Banga Solar Pvt Ltd":
    st.markdown("""
    <style>
    .stApp, [data-testid="stSidebar"] {
        --st-primary-color: #166534;
        --st-primary-color-bg: #f0f8f2;
        --st-primary-color-text: #14532d;
    }
    [data-baseweb="radio"]:has(input:checked) > div:first-child {
        background-color: #166534 !important;
        border-color: #166534 !important;
    }
    [data-baseweb="checkbox"]:has(input:checked) > span:first-child {
        background-color: #166534 !important;
        border-color: #166534 !important;
    }
    </style>
    """, unsafe_allow_html=True)


# Logout control shown only after successful login.
with st.sidebar:
    st.caption(
        f"🔐 Logged in: {st.session_state.get('pv_authenticated_user', 'User')}"
    )
    if st.button(
        "Logout",
        key="pv_logout_button",
        width="stretch",
    ):
        for _key in [
            "pv_authenticated",
            "pv_authenticated_user",
            "pv_auth_token",
            "pv_user",
            "pv_login_response",
            "pv_login_username",
            "pv_login_password",
        ]:
            st.session_state.pop(_key, None)
        st.rerun()

# ============================================================
# MODERN DESIGN SYSTEM & ADVANCED CSS
# ============================================================

st.markdown(
    company_theme("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');

    :root {
        --pv-orange:#f97316;
        --pv-orange-dark:#ea580c;
        --pv-orange-soft:#fff7ed;
        --pv-border:#e2e8f0;
        --pv-text:#0f172a;
        --pv-muted:#64748b;
        --pv-bg:#f8fafc;
    }

    html, body, [class*="css"] {
        font-family:'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    .stApp {
        background:#f8fafc;
        color:#0f172a;
    }

    /* ============================================================
       STREAMLIT NATIVE SHELL
       ------------------------------------------------------------
       Do not manually translate/zero the sidebar based on hidden DOM
       controls. Streamlit itself controls open/close state.
       ============================================================ */

    div[data-testid="stDecoration"] {
        display:none !important;
    }

    /* Keep native header because it owns the sidebar reopen control.
       Hide only Streamlit's right-side toolbar (Deploy / menu). */
    header[data-testid="stHeader"] {
        display:flex !important;
        visibility:visible !important;
        height:46px !important;
        min-height:46px !important;
        background:rgba(248,250,252,.98) !important;
        border-bottom:1px solid #e2e8f0 !important;
        box-shadow:none !important;
        z-index:99999 !important;
    }

    [data-testid="stToolbar"],
    [data-testid="stStatusWidget"],
    [data-testid="stHeaderActionElements"] {
        display:none !important;
    }

    /* Native sidebar toggle / reopen controls must stay accessible. */
    [data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"] {
        display:flex !important;
        visibility:visible !important;
        opacity:1 !important;
        pointer-events:auto !important;
        z-index:100001 !important;
    }

    /* OPEN sidebar only. No rules are applied to a genuinely collapsed sidebar. */
    [data-testid="stSidebar"]:not([aria-expanded="false"]) {
        width:300px !important;
        min-width:300px !important;
        max-width:300px !important;
        flex:0 0 300px !important;
        background:#ffffff !important;
        border-right:1px solid #e2e8f0 !important;
        box-shadow:5px 0 18px rgba(15,23,42,.035) !important;
        overflow:hidden !important;
    }

    [data-testid="stSidebar"] > div:first-child,
    [data-testid="stSidebar"] [data-testid="stSidebarContent"] {
        width:100% !important;
        max-width:100% !important;
        min-width:0 !important;
        box-sizing:border-box !important;
    }

    [data-testid="stSidebar"] [data-testid="stSidebarContent"] {
        padding:.75rem .72rem 1.5rem !important;
        overflow-y:auto !important;
        overflow-x:hidden !important;
    }

    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
        gap:.48rem !important;
    }

    [data-testid="stSidebar"] * {
        box-sizing:border-box !important;
        color:#334155 !important;
    }

    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea,
    [data-testid="stSidebar"] [data-baseweb="select"],
    [data-testid="stSidebar"] [data-testid="stDateInput"],
    [data-testid="stSidebar"] [data-testid="stFileUploader"],
    [data-testid="stSidebar"] section[data-testid="stFileUploaderDropzone"] {
        width:100% !important;
        min-width:0 !important;
        max-width:100% !important;
    }

    /* Let Streamlit naturally make the main area consume all remaining width. */
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"],
    [data-testid="stAppViewContainer"] > .main {
        max-width:none !important;
        min-width:0 !important;
    }

    [data-testid="stMain"] {
        flex:1 1 auto !important;
        width:auto !important;
        overflow-x:hidden !important;
    }

    [data-testid="stMainBlockContainer"],
    .block-container {
        width:100% !important;
        max-width:none !important;
        min-width:0 !important;
        margin:0 !important;
        padding:.9rem 1rem 3rem !important;
        box-sizing:border-box !important;
    }

    /* ============================================================
       SIDEBAR VISUALS
       ============================================================ */

    .sidebar-head-badge {
        background:#fff7ed;
        border:1px solid #ffedd5;
        color:#f97316 !important;
        font-size:11px;
        font-weight:800;
        letter-spacing:.8px;
        text-transform:uppercase;
        padding:4px 10px;
        border-radius:6px;
        display:inline-block;
        margin-bottom:10px;
    }

    .sidebar-section-num {
        font-size:11px;
        font-weight:800;
        color:#475569 !important;
        text-transform:uppercase;
        letter-spacing:.7px;
        margin-top:14px;
        margin-bottom:7px;
        display:flex;
        align-items:center;
        gap:5px;
    }

    [data-testid="stSidebar"] [data-testid="stFileUploader"] {
        background:#f8fafc !important;
        border:1.5px dashed #cbd5e1 !important;
        border-radius:13px !important;
        padding:7px !important;
        margin-bottom:3px !important;
    }

    [data-testid="stSidebar"] section[data-testid="stFileUploaderDropzone"] {
        background:#ffffff !important;
        border:1.5px dashed #cbd5e1 !important;
    }

    /* ============================================================
       TOP BRAND HEADER
       ============================================================ */

    .pv-navbar {
        width:100%;
        max-width:100%;
        box-sizing:border-box;
        background:#ffffff;
        border:1px solid #e2e8f0;
        border-radius:17px;
        padding:14px 20px;
        margin-bottom:18px;
        box-shadow:0 4px 20px -2px rgba(15,23,42,.04);
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:16px;
    }

    .pv-brand-title {
        font-size:27px;
        font-weight:900;
        letter-spacing:-.6px;
        color:#0f172a;
        margin:0;
        display:flex;
        align-items:center;
        gap:10px;
    }

    .pv-brand-title span { color:#f97316; }

    .pv-nav-subtext {
        font-size:12.5px;
        color:#64748b;
        margin-top:2px;
        font-weight:500;
    }

    .pv-pill-tag {
        flex:0 0 auto;
        background:#fff7ed;
        border:1px solid #ffedd5;
        border-radius:24px;
        padding:7px 14px;
        font-size:11.5px;
        font-weight:800;
        color:#ea580c;
        display:flex;
        align-items:center;
        gap:8px;
    }

    .pv-live-dot {
        width:8px;
        height:8px;
        background:#22c55e;
        border-radius:50%;
        box-shadow:0 0 8px #22c55e;
    }

    /* ============================================================
       DIRECT PAGE NAVIGATION
       Desktop = single responsive row when space permits.
       Smaller screens = clean wrapping, never cut off.
       ============================================================ */

    [data-testid="stRadio"] > div[role="radiogroup"] {
        display:flex !important;
        flex-direction:row !important;
        flex-wrap:wrap !important;
        align-items:stretch !important;
        width:100% !important;
        gap:7px !important;
        background:#ffffff !important;
        border:1px solid #e2e8f0 !important;
        border-radius:14px !important;
        padding:7px 8px !important;
        margin:0 0 18px 0 !important;
        box-shadow:0 2px 10px rgba(15,23,42,.035) !important;
        overflow:visible !important;
    }

    [data-testid="stRadio"] > div[role="radiogroup"] label {
        flex:1 1 135px !important;
        min-width:118px !important;
        max-width:none !important;
        width:auto !important;
        min-height:38px !important;
        padding:8px 10px !important;
        margin:0 !important;
        border:1px solid transparent !important;
        border-radius:10px !important;
        background:#ffffff !important;
        cursor:pointer !important;
    }

    [data-testid="stRadio"] > div[role="radiogroup"] label:hover {
        background:#fff7ed !important;
        border-color:#fed7aa !important;
    }

    [data-testid="stRadio"] > div[role="radiogroup"] label p,
    [data-testid="stRadio"] > div[role="radiogroup"] label span,
    [data-testid="stRadio"] > div[role="radiogroup"] label div {
        color:#334155 !important;
        font-size:12px !important;
        font-weight:750 !important;
        opacity:1 !important;
        visibility:visible !important;
        white-space:normal !important;
    }

    [data-testid="stRadio"] > div[role="radiogroup"] label:has(input:checked) {
        background:#fff7ed !important;
        border-color:#fdba74 !important;
        box-shadow:inset 0 -3px 0 #f97316 !important;
    }

    [data-testid="stRadio"] > div[role="radiogroup"] label:has(input:checked) * {
        color:#ea580c !important;
        font-weight:850 !important;
    }

    /* ============================================================
       KPI / CHART / TABLE DESIGN
       ============================================================ */

    [data-testid="stHorizontalBlock"] {
        width:100% !important;
        min-width:0 !important;
        flex-wrap:wrap !important;
        gap:.72rem !important;
    }

    [data-testid="column"] {
        min-width:0 !important;
        flex:1 1 175px !important;
    }

    [data-testid="column"]:has([data-testid="stPlotlyChart"]) {
        flex:1 1 430px !important;
    }

    .pv-kpi-card {
        background:#ffffff;
        border:1px solid #e2e8f0;
        border-radius:14px;
        padding:14px;
        box-shadow:0 4px 16px -2px rgba(15,23,42,.04);
        min-height:110px;
        height:100%;
        box-sizing:border-box;
        transition:.2s ease;
    }

    .pv-kpi-card:hover {
        transform:translateY(-2px);
        border-color:#fdba74;
        box-shadow:0 10px 22px -4px rgba(249,115,22,.15);
    }

    .pv-kpi-header {
        display:flex;
        align-items:center;
        gap:8px;
    }

    .pv-kpi-icon {
        width:32px;
        height:32px;
        border-radius:10px;
        background:#fff7ed;
        border:1px solid #ffedd5;
        display:flex;
        align-items:center;
        justify-content:center;
        font-size:15px;
        flex-shrink:0;
    }

    .pv-kpi-label {
        font-size:9.8px;
        font-weight:800;
        color:#64748b;
        text-transform:uppercase;
        letter-spacing:.5px;
        line-height:1.2;
    }

    .pv-kpi-value {
        font-size:18px !important;
        font-weight:850 !important;
        color:#0f172a !important;
        margin-top:6px;
        margin-bottom:4px;
        line-height:1.15 !important;
        white-space:nowrap !important;
    }

    .pv-kpi-badge {
        background:#dcfce7;
        color:#15803d;
        font-size:9.2px;
        font-weight:750;
        padding:2px 6px;
        border-radius:6px;
        width:fit-content;
        white-space:nowrap;
    }

    .pv-chart-card {
        width:100%;
        max-width:100%;
        box-sizing:border-box;
        background:#ffffff;
        border:1px solid #e2e8f0;
        border-radius:16px;
        padding:18px 20px;
        box-shadow:0 4px 18px -2px rgba(15,23,42,.03);
        margin-bottom:18px;
    }

    .pv-card-header {
        display:flex;
        align-items:center;
        justify-content:space-between;
        margin-bottom:14px;
        padding-bottom:9px;
        border-bottom:1px solid #f1f5f9;
    }

    .pv-card-title {
        font-size:15px;
        font-weight:800;
        color:#0f172a;
        display:flex;
        align-items:center;
        gap:9px;
    }

    .pv-insight-item {
        display:flex;
        align-items:center;
        gap:11px;
        background:#f8fafc;
        border:1px solid #f1f5f9;
        border-radius:12px;
        padding:11px 13px;
        margin-bottom:9px;
    }

    .pv-insight-icon { font-size:18px; }
    .pv-insight-text { font-size:12px;color:#334155;font-weight:600; }

    .stButton button,
    .stDownloadButton button {
        background:linear-gradient(135deg,#f97316 0%,#ea580c 100%) !important;
        color:#ffffff !important;
        border:none !important;
        border-radius:10px !important;
        font-weight:750 !important;
        font-size:13px !important;
        box-shadow:0 4px 12px rgba(249,115,22,.22) !important;
    }

    div[data-testid="stDataFrame"] {
        border:1px solid #e2e8f0;
        border-radius:12px;
        overflow:hidden;
    }

    [data-testid="stAlert"] { border-radius:12px !important; }
    [data-testid="stAlert"] *,
    div[role="alert"] * { color:#334155 !important; }

    div[data-testid="stPlotlyChart"],
    .js-plotly-plot,
    .plot-container,
    .plotly {
        width:100% !important;
        max-width:100% !important;
        min-width:0 !important;
    }

    .pv-footer {
        width:100%;
        max-width:100%;
        box-sizing:border-box;
        margin-top:30px;
        padding-top:16px;
        border-top:1px solid #e2e8f0;
        display:flex;
        justify-content:space-between;
        gap:14px;
        flex-wrap:wrap;
        font-size:11px;
        color:#94a3b8;
    }

    /* ============================================================
       RESPONSIVE BREAKPOINTS
       ============================================================ */

    @media (max-width:1366px) {
        [data-testid="stSidebar"]:not([aria-expanded="false"]) {
            width:280px !important;
            min-width:280px !important;
            max-width:280px !important;
            flex-basis:280px !important;
        }

        [data-testid="stMainBlockContainer"],
        .block-container {
            padding-left:.75rem !important;
            padding-right:.75rem !important;
        }

        .pv-brand-title { font-size:24px; }

        [data-testid="stRadio"] > div[role="radiogroup"] label {
            flex-basis:122px !important;
            min-width:108px !important;
        }
    }

    @media (max-width:1050px) {
        [data-testid="stSidebar"]:not([aria-expanded="false"]) {
            width:260px !important;
            min-width:260px !important;
            max-width:260px !important;
            flex-basis:260px !important;
        }

        .pv-navbar {
            flex-wrap:wrap;
            padding:12px 15px;
        }

        .pv-pill-tag { display:none; }

        [data-testid="column"]:has([data-testid="stPlotlyChart"]) {
            flex-basis:100% !important;
        }
    }

    @media (max-width:760px) {
        [data-testid="stSidebar"]:not([aria-expanded="false"]) {
            width:min(88vw,300px) !important;
            min-width:min(88vw,300px) !important;
            max-width:min(88vw,300px) !important;
        }

        [data-testid="stMainBlockContainer"],
        .block-container {
            padding:.55rem .5rem 2rem !important;
        }

        .pv-brand-title { font-size:21px; }

        [data-testid="stRadio"] > div[role="radiogroup"] label,
        [data-testid="column"],
        [data-testid="column"]:has([data-testid="stPlotlyChart"]) {
            flex:1 1 100% !important;
            min-width:100% !important;
        }
    }

    /* ============================================================
       V21 SIDEBAR SAFETY PATCH
       ============================================================ */
    [data-testid="stSidebar"] {
        z-index: 100000 !important;
    }

    [data-testid="stSidebar"]:not([aria-expanded="false"]) {
        display:block !important;
        visibility:visible !important;
        opacity:1 !important;
    }

    /* Floating fallback button injected by JS when Streamlit sidebar is hidden */
    #pv-open-filters-btn {
        font-family:'Plus Jakarta Sans',-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif !important;
    }

    @media (min-width: 1200px) {
        [data-testid="stRadio"] > div[role="radiogroup"] label {
            flex: 1 1 120px !important;
            min-width: 108px !important;
        }

        [data-testid="stRadio"] > div[role="radiogroup"] label p,
        [data-testid="stRadio"] > div[role="radiogroup"] label span,
        [data-testid="stRadio"] > div[role="radiogroup"] label div {
            font-size: 11.5px !important;
        }
    }

    /* V22 — stable page position during reruns */
    html {
        scrollbar-gutter: stable !important;
        overflow-anchor: none !important;
    }

    body,
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"] {
        overflow-anchor: none !important;
    }

    [data-testid="stMainBlockContainer"],
    .block-container {
        min-height: 100vh !important;
    }

    /* ============================================================
       V27 — PV-BLINK ANALYTICS OS 2.0
       Advanced visual polish without changing dashboard logic.
       ============================================================ */

    :root {
        --pv-accent:#f97316;
        --pv-accent-2:#fb923c;
        --pv-accent-soft:#fff7ed;
        --pv-ink:#0f172a;
        --pv-muted:#64748b;
        --pv-line:#e2e8f0;
        --pv-surface:#ffffff;
        --pv-canvas:#f5f7fb;
        --pv-success:#16a34a;
        --pv-shadow:0 10px 34px rgba(15,23,42,.07);
        --pv-shadow-hover:0 16px 38px rgba(15,23,42,.10);
    }

    html {
        scroll-behavior:smooth;
        scrollbar-gutter:stable;
    }

    body,
    .stApp {
        background:
            radial-gradient(circle at 90% 0%, rgba(249,115,22,.055), transparent 28rem),
            linear-gradient(180deg,#f8fafc 0%,#f4f7fb 100%) !important;
    }

    /* Premium compact browser-app canvas */
    [data-testid="stMainBlockContainer"],
    .block-container {
        padding-top:.78rem !important;
        padding-bottom:3.25rem !important;
    }

    /* Sidebar: control-center appearance */
    [data-testid="stSidebar"]:not([aria-expanded="false"]) {
        background:
            linear-gradient(180deg,#ffffff 0%,#fbfcfe 100%) !important;
        border-right:1px solid #dfe6ee !important;
        box-shadow:8px 0 30px rgba(15,23,42,.04) !important;
    }

    [data-testid="stSidebar"] [data-testid="stSidebarContent"] {
        scrollbar-width:thin;
        scrollbar-color:#cbd5e1 transparent;
    }

    [data-testid="stSidebar"] [data-testid="stSidebarContent"]::-webkit-scrollbar {
        width:6px;
    }

    [data-testid="stSidebar"] [data-testid="stSidebarContent"]::-webkit-scrollbar-thumb {
        background:#cbd5e1;
        border-radius:999px;
    }

    .sidebar-section-num {
        position:relative;
        padding-left:10px;
        margin-top:17px !important;
    }

    .sidebar-section-num::before {
        content:"";
        position:absolute;
        left:0;
        top:2px;
        bottom:2px;
        width:3px;
        border-radius:999px;
        background:linear-gradient(180deg,#f97316,#fb923c);
    }

    /* Hero / OS topbar */
    .pv-navbar {
        position:relative;
        overflow:hidden;
        min-height:76px;
        border:1px solid rgba(203,213,225,.78) !important;
        border-radius:18px !important;
        box-shadow:var(--pv-shadow) !important;
        background:
            linear-gradient(118deg,#ffffff 0%,#ffffff 68%,#fff8f1 100%) !important;
    }

    .pv-navbar::before {
        content:"";
        position:absolute;
        left:0;
        top:0;
        bottom:0;
        width:4px;
        background:linear-gradient(180deg,#f97316,#fb923c);
    }

    .pv-navbar::after {
        content:"";
        position:absolute;
        width:160px;
        height:160px;
        right:-70px;
        top:-90px;
        border-radius:50%;
        background:rgba(249,115,22,.06);
        pointer-events:none;
    }

    .pv-brand-title {
        letter-spacing:-.75px !important;
    }

    .pv-pill-tag {
        box-shadow:0 3px 12px rgba(249,115,22,.08);
    }

    /* App navigation */
    [data-testid="stRadio"] > div[role="radiogroup"] {
        backdrop-filter:blur(8px);
        background:rgba(255,255,255,.92) !important;
        border:1px solid rgba(203,213,225,.8) !important;
        box-shadow:0 7px 22px rgba(15,23,42,.045) !important;
    }

    [data-testid="stRadio"] > div[role="radiogroup"] label {
        transition:
            background .18s ease,
            border-color .18s ease,
            transform .18s ease,
            box-shadow .18s ease !important;
    }

    [data-testid="stRadio"] > div[role="radiogroup"] label:hover {
        transform:translateY(-1px);
        box-shadow:0 5px 14px rgba(249,115,22,.08) !important;
    }

    /* KPI cards */
    .pv-kpi-card {
        position:relative;
        overflow:hidden;
        border:1px solid rgba(203,213,225,.82) !important;
        border-radius:16px !important;
        box-shadow:0 7px 22px rgba(15,23,42,.052) !important;
        transition:
            transform .18s ease,
            border-color .18s ease,
            box-shadow .18s ease !important;
    }

    .pv-kpi-card::after {
        content:"";
        position:absolute;
        left:0;
        right:0;
        bottom:0;
        height:3px;
        opacity:.0;
        background:linear-gradient(90deg,#f97316,#fb923c);
        transition:opacity .18s ease;
    }

    .pv-kpi-card:hover {
        transform:translateY(-3px) !important;
        border-color:#fdba74 !important;
        box-shadow:var(--pv-shadow-hover) !important;
    }

    .pv-kpi-card:hover::after {
        opacity:1;
    }

    .pv-kpi-icon {
        box-shadow:inset 0 0 0 1px rgba(249,115,22,.03);
    }

    .pv-kpi-value {
        font-variant-numeric:tabular-nums;
        letter-spacing:-.35px;
    }

    /* Chart panels */
    .pv-chart-card {
        position:relative;
        border:1px solid rgba(203,213,225,.82) !important;
        border-radius:18px !important;
        box-shadow:0 8px 26px rgba(15,23,42,.05) !important;
        background:rgba(255,255,255,.97) !important;
        transition:
            border-color .18s ease,
            box-shadow .18s ease,
            transform .18s ease !important;
    }

    .pv-chart-card:hover {
        border-color:#fdba74 !important;
        box-shadow:0 13px 32px rgba(15,23,42,.075) !important;
    }

    .pv-card-header {
        min-height:34px;
        border-bottom:1px solid #edf2f7 !important;
    }

    .pv-card-title {
        letter-spacing:-.15px;
    }

    /* Actual Plotly frame: removes visual ambiguity between cards/charts */
    div[data-testid="stPlotlyChart"] {
        border-radius:13px !important;
        overflow:hidden !important;
        background:#ffffff !important;
    }

    div[data-testid="stPlotlyChart"] > div {
        border-radius:13px !important;
    }

    /* Tables */
    div[data-testid="stDataFrame"] {
        border:1px solid rgba(203,213,225,.86) !important;
        border-radius:14px !important;
        box-shadow:0 6px 18px rgba(15,23,42,.035);
    }

    /* Inputs */
    [data-baseweb="select"] > div,
    [data-testid="stDateInput"] input,
    .stTextInput input,
    .stNumberInput input {
        border-radius:10px !important;
    }

    [data-baseweb="select"] > div:focus-within,
    [data-testid="stDateInput"] input:focus {
        border-color:#fb923c !important;
        box-shadow:0 0 0 3px rgba(249,115,22,.10) !important;
    }

    /* Buttons */
    .stButton button,
    .stDownloadButton button {
        min-height:39px;
        transition:
            transform .16s ease,
            box-shadow .16s ease,
            filter .16s ease !important;
    }

    .stButton button:hover,
    .stDownloadButton button:hover {
        transform:translateY(-1px);
        filter:saturate(1.05);
        box-shadow:0 8px 18px rgba(249,115,22,.22) !important;
    }

    /* Expander / secondary surfaces */
    [data-testid="stExpander"] {
        background:rgba(255,255,255,.82);
        border:1px solid #e2e8f0 !important;
        border-radius:13px !important;
        overflow:hidden;
    }

    /* Section alerts / insight surfaces */
    [data-testid="stAlert"] {
        border-radius:13px !important;
        box-shadow:0 4px 14px rgba(15,23,42,.025);
    }

    /* Footer */
    .pv-footer {
        opacity:.92;
        padding-bottom:8px;
    }

    /* Large desktop: preserve management-dashboard density */
    @media (min-width:1500px) {
        [data-testid="stMainBlockContainer"],
        .block-container {
            padding-left:1.2rem !important;
            padding-right:1.2rem !important;
        }
    }

    /* Laptop */
    @media (max-width:1200px) {
        .pv-navbar {
            min-height:70px;
        }

        .pv-kpi-value {
            font-size:17px !important;
        }
    }

    /* Mobile/narrow window */
    @media (max-width:760px) {
        .pv-navbar {
            border-radius:14px !important;
        }

        .pv-chart-card {
            padding:14px !important;
            border-radius:14px !important;
        }

        .pv-kpi-card {
            min-height:102px;
        }
    }

    /* Full-value cursor hover: native browser tooltip, zero extra DOM load. */
    .pv-kpi-card[title] {
        cursor:help !important;
    }

    /* ============================================================
       SALES EMPLOYEE TARGET INPUT — CLEAN FIELD-STYLE UI
       ============================================================ */

    [data-testid="stNumberInput"] {
        margin-bottom: 2px !important;
    }

    [data-testid="stNumberInput"] label {
        color:#334155 !important;
        font-weight:800 !important;
        font-size:12px !important;
    }

    [data-testid="stNumberInput"] input {
        background:#111827 !important;
        color:#ffffff !important;
        border-color:#27364b !important;
        font-weight:800 !important;
    }

    [data-testid="stNumberInput"] button {
        background:#1f2937 !important;
        color:#ffffff !important;
        border-color:#27364b !important;
    }

    [data-testid="stNumberInput"] button:hover {
        background:#f97316 !important;
        border-color:#f97316 !important;
    }

    /* ============================================================
       SALES TARGET LOCK / CHANGE CONTROLS
       ============================================================ */

    .sales-target-lock-status {
        min-height:46px;
        border-radius:12px;
        padding:10px 14px;
        display:flex;
        align-items:center;
        gap:8px;
        font-size:12px;
        line-height:1.35;
        margin-top:1px;
        border:1px solid #dce5ef;
        background:#ffffff;
    }

    .sales-target-lock-status span {
        color:#64748b;
        font-weight:500;
    }

    .sales-target-lock-status.locked {
        border-color:#86efac;
        background:#f0fdf4;
        color:#166534;
    }

    .sales-target-lock-status.unlocked {
        border-color:#fed7aa;
        background:#fff7ed;
        color:#c2410c;
    }

    .sales-target-lock-status b {
        white-space:nowrap;
        font-weight:850;
    }

    .target-person-card {
        border:1px solid #e2e8f0;
        background:#ffffff;
        border-radius:14px;
        padding:10px 12px;
        margin:8px 0 10px 0;
        box-shadow:0 4px 14px rgba(15,23,42,.035);
    }

    .target-person-name {
        font-size:13px;
        font-weight:850;
        color:#0f172a;
        line-height:1.25;
    }

    .target-person-sub {
        font-size:10.5px;
        color:#64748b;
        margin-top:2px;
    }

    .target-period-summary {
        font-size:10.5px;
        color:#64748b;
        margin-top:4px;
        line-height:1.4;
    }

    .target-date-pill {
        display:inline-block;
        padding:3px 8px;
        margin:2px 4px 2px 0;
        border-radius:999px;
        border:1px solid #fed7aa;
        background:#fff7ed;
        color:#c2410c;
        font-size:9.8px;
        font-weight:750;
    }

    /* Target value uses text_input now: no +/- stepper buttons. */
    [data-testid="stTextInput"] input {
        min-height:39px !important;
    }

    /* Stable compact Target popover/dialog surface */
    div[data-baseweb="popover"] {
        max-width:min(560px,94vw) !important;
    }

    div[data-baseweb="popover"] > div {
        max-width:min(560px,94vw) !important;
    }

    [data-testid="stPopoverBody"] {
        width:min(540px,92vw) !important;
        max-width:min(540px,92vw) !important;
        max-height:78vh !important;
        overflow-y:auto !important;
        overflow-x:hidden !important;
        box-sizing:border-box !important;
    }

    [data-testid="stPopoverBody"] [data-testid="stVerticalBlock"] {
        gap:.55rem !important;
    }

    [data-testid="stPopoverBody"] [data-testid="stDataFrame"] {
        max-height:260px !important;
    }
</style>
"""),
    unsafe_allow_html=True,
)

# ============================================================
# EXTENDED ENTERPRISE UI — RETURNS + FIELD KPI/KRA MODULES
# ============================================================
st.markdown(
    company_theme("""
<style>
    .pv-net-banner {background:linear-gradient(135deg,#fff7ed,#ffffff);border:1px solid #fed7aa;border-left:5px solid #f97316;border-radius:14px;padding:13px 16px;margin:10px 0 16px 0;color:#334155;font-size:12px;}
    .pv-net-banner b {color:#0f172a;}
    .pv-return-chip {display:inline-block;background:#fff1f2;border:1px solid #fecdd3;color:#be123c;border-radius:999px;padding:4px 9px;font-size:10px;font-weight:800;margin-right:6px;}
    .pv-field-chip {display:inline-block;background:#eff6ff;border:1px solid #bfdbfe;color:#1d4ed8;border-radius:999px;padding:4px 9px;font-size:10px;font-weight:800;margin-right:6px;}
</style>
"""),
    unsafe_allow_html=True,
)



# ============================================================
# ROBUST SIDEBAR CONTROLLER
# ============================================================
st.iframe(
    company_theme("""
<script>
(function () {
    const doc = window.parent.document;

    function sidebarEl() {
        return doc.querySelector('[data-testid="stSidebar"]');
    }

    function sidebarIsVisible() {
        const sb = sidebarEl();
        if (!sb) return false;

        const r = sb.getBoundingClientRect();
        const cs = window.parent.getComputedStyle(sb);

        return (
            r.width > 120 &&
            r.right > 20 &&
            r.left < 80 &&
            cs.display !== "none" &&
            cs.visibility !== "hidden" &&
            cs.opacity !== "0"
        );
    }

    function candidateButtons() {
        return Array.from(doc.querySelectorAll('button, [role="button"]'));
    }

    function findSidebarToggle() {
        const buttons = candidateButtons();

        // Prefer explicit Streamlit testids when available.
        const explicit = doc.querySelector(
            '[data-testid="stSidebarCollapsedControl"] button,' +
            '[data-testid="collapsedControl"] button,' +
            '[data-testid="stSidebarCollapseButton"] button,' +
            '[data-testid="stSidebarCollapsedControl"],' +
            '[data-testid="collapsedControl"]'
        );

        if (explicit) return explicit;

        // Fallback across Streamlit versions.
        const keywords = [
            'sidebar',
            'open sidebar',
            'show sidebar',
            'expand sidebar',
            'collapse sidebar',
            'navigation'
        ];

        for (const btn of buttons) {
            const label = [
                btn.getAttribute('aria-label') || '',
                btn.getAttribute('title') || '',
                btn.textContent || ''
            ].join(' ').trim().toLowerCase();

            if (keywords.some(k => label.includes(k))) {
                return btn;
            }
        }

        return null;
    }

    function openSidebar() {
        if (sidebarIsVisible()) return true;

        const btn = findSidebarToggle();
        if (btn) {
            try {
                btn.click();
                return true;
            } catch (e) {}
        }
        return false;
    }

    function ensureFallbackButton() {
        let btn = doc.getElementById('pv-open-filters-btn');

        if (!btn) {
            btn = doc.createElement('button');
            btn.id = 'pv-open-filters-btn';
            btn.type = 'button';
            btn.innerHTML = '☰&nbsp;&nbsp;Filters';

            btn.style.cssText = `
                position:fixed;
                top:58px;
                left:10px;
                z-index:2147483645;
                border:1px solid #fdba74;
                background:#ffffff;
                color:#ea580c;
                height:40px;
                padding:0 13px;
                border-radius:10px;
                font-size:12px;
                font-weight:800;
                cursor:pointer;
                box-shadow:0 7px 22px rgba(15,23,42,.14);
                display:none;
                align-items:center;
                justify-content:center;
            `;

            btn.addEventListener('mouseenter', () => {
                btn.style.background = '#fff7ed';
                btn.style.borderColor = '#f97316';
            });

            btn.addEventListener('mouseleave', () => {
                btn.style.background = '#ffffff';
                btn.style.borderColor = '#fdba74';
            });

            btn.addEventListener('click', function () {
                openSidebar();

                setTimeout(function () {
                    syncButton();
                }, 250);
            });

            doc.body.appendChild(btn);
        }

        return btn;
    }

    function syncButton() {
        const btn = ensureFallbackButton();
        btn.style.display = sidebarIsVisible() ? 'none' : 'flex';
    }

    function initialOpen() {
        // Streamlit may render the sidebar control slightly after the app body.
        let attempts = 0;

        const timer = setInterval(function () {
            attempts += 1;

            if (sidebarIsVisible()) {
                clearInterval(timer);
                syncButton();
                return;
            }

            openSidebar();
            syncButton();

            if (attempts >= 20) {
                clearInterval(timer);
            }
        }, 180);
    }

    // Keep fallback button state synchronized as Streamlit rerenders.
    if (!window.parent.__pvSidebarObserverInstalled) {
        window.parent.__pvSidebarObserverInstalled = true;

        const observer = new MutationObserver(function () {
            clearTimeout(window.parent.__pvSidebarSyncTimer);
            window.parent.__pvSidebarSyncTimer = setTimeout(syncButton, 80);
        });

        observer.observe(doc.body, {
            childList:true,
            subtree:true,
            attributes:true,
            attributeFilter:['aria-expanded','style','class']
        });

        window.parent.addEventListener('resize', syncButton);
    }

    initialOpen();
    syncButton();
})();
</script>
"""),
    height=1,
)

# Standard Plotly Modebar Configuration with Zoom, Pan, Download & Fullscreen enabled
PLOTLY_CONFIG = {
    "displayModeBar": True,
    "displaylogo": False,
    "responsive": True,
    "scrollZoom": True,
    "modeBarButtonsToRemove": ["lasso2d"],
    "toImageButtonOptions": {
        "format": "png",
        "filename": "PVBlink_Analytics_Chart",
        "scale": 3,
    },
}

# ============================================================
# CONSTANTS & ALIASES
# ============================================================

CORE_REQUIRED = [
    "QTY",
    "Amount",
    "Item Name",
    "Item Group",
    "Invoice Date",
    "Ledger Name",
]

OPTIONAL_ALIASES = {
    "Qty": "QTY",
    "Quantity": "QTY",
    "QTY ": "QTY",
    "Rate ": "Rate",
    "Item": "Item Name",
    "Product": "Item Name",
    "Product Name": "Item Name",
    "Group": "Item Group",
    "Product Group": "Item Group",
    "Representative": "Representative Ref",
    "Representative Name": "Representative Ref",
    "Date": "Invoice Date",
    "Invoice Dt": "Invoice Date",
    "Customer": "Ledger Name",
    "Customer Name": "Ledger Name",
    "Invoice Number": "Invoice No#",
    "Invoice No": "Invoice No#",
}

CONTEXT_COLUMNS = [
    "Branch",
    "Ledger Name",
    "Invoice Date",
    "Invoice No#",
    "Representative Ref",
    "Representative User",
]

SERVICE_COLUMNS = [
    "Ticket No", "MTCE Type", "Company Name", "Mobile", "Email", "Product",
    "Address", "Installation Address", "Query Allocate To", "Visiting Person",
    "Closed By", "Closed Date", "Created By", "Updated By", "Follow Date",
    "MTCE Status", "Query Master", "Query", "Solution", "Remark"
]

# ============================================================
# HELPERS & FORMATTING
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    return " ".join(text.split()).strip()


def clean_label(text, max_len=24):
    """
    Prevents text label clipping:
    - Removes emails in parentheses e.g. "Kishan Pandey(kishan@pvblink.com)" -> "Kishan Pandey"
    - Truncates extra long product strings cleanly with ellipsis
    """
    if pd.isna(text):
        return ""
    s = str(text).strip()
    if "(" in s and "@" in s:
        s = s.split("(")[0].strip()
    if len(s) > max_len:
        return s[:max_len - 2] + ".."
    return s


def format_indian_currency(val):
    """Formats numeric values into Indian Cr, L, and K standards without space wrapping."""
    if pd.isna(val):
        return "₹0"
    v = float(val)
    abs_v = abs(v)
    sign = "-" if v < 0 else ""
    if abs_v >= 10_000_000:
        return f"{sign}₹{abs_v / 10_000_000:.2f}Cr"
    elif abs_v >= 100_000:
        return f"{sign}₹{abs_v / 100_000:.2f}L"
    elif abs_v >= 1_000:
        return f"{sign}₹{abs_v / 1_000:.1f}K"
    else:
        return f"{sign}₹{abs_v:,.0f}"



def format_indian_number_full(value, decimals=None):
    """Exact Indian number grouping for hover text."""
    try:
        if pd.isna(value):
            return "0"
        n = float(value)
    except Exception:
        return clean_text(value)

    sign = "-" if n < 0 else ""
    n = abs(n)

    if decimals is None:
        decimals = 0 if abs(n - round(n)) < 1e-9 else 2

    raw = f"{n:.{decimals}f}"
    integer_part, dot, decimal_part = raw.partition(".")

    if len(integer_part) <= 3:
        grouped = integer_part
    else:
        last3 = integer_part[-3:]
        rest = integer_part[:-3]
        groups = []
        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            groups.insert(0, rest)
        grouped = ",".join(groups + [last3])

    if decimals > 0:
        grouped += "." + decimal_part

    return sign + grouped


def format_indian_currency_full(value):
    """Exact rupee value for cursor hover, without Cr/L/K abbreviation."""
    return "₹" + format_indian_number_full(value)


def _html_attr(value):
    """Escape values placed inside HTML attributes."""
    text = "" if value is None else str(value)
    return (
        text.replace("&", "&amp;")
            .replace('"', "&quot;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
    )


APPROVED_PRODUCT_KW = [2, 3, 3.3, 3.6, 4, 4.6, 5, 5.5, 6, 8, 10, 12, 15, 18, 20, 25, 30]

PRODUCT_CHART_CATALOG = [
    "10 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT10K-M1)-PVBLINK",
    "12 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT12K-M1)-PVBLINK",
    "15 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT15K-M1)-PVBLINK",
    "18 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT18K-M1)-PVBLINK",
    "2 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-DC-CT (PVBSM2K-M1)(G)-PVBLINK",
    "20 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT20K-M1)-PVBLINK",
    "25 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT25K-M1)-PVBLINK",
    "3 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT (PVBSM3K-M1)(G)-PVBLINK",
    "3 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-CT (PVBSM3K-M1)(G)-PVBLINK",
    "3 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-DC-CT (PVBSM3K-M1)(G)-PVBLINK",
    "3.3 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-CT (PVBS3.3K-M1)(G)-PVBLINK",
    "3.3 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-DC-CT (PVBS3.3K-M1)(G)-PVBLINK",
    "3.6 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-CT (PVBS3.6K-M1)(G)-PVBLINK",
    "3.6 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-DC-CT (PVBS3.6K-M1)(G)-PVBLINK",
    "30 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT30K-M1)-PVBLINK",
    "4 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT (PVBSM4K-M1)(G)-PVBLINK",
    "4 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-CT (PVBSM4K-M1)(G)-PVBLINK",
    "4 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-DC-CT (PVBSM4K-M1)(G)-PVBLINK",
    "4.6 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-CT (PVBSM4.6K-M1)-PVBLINK",
    "4.6 KW 1 Phase Ongrid PV Solar Inverter 2 MPPT-W/O-CT (PVBS4.6KPro-M1)-PVBLINK",
    "5 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-CT (PVBSM5K-M1)-PVBLINK",
    "5 KW 1 Phase Ongrid PV Solar Inverter 2 MPPT (PVBS5K-M1)-PVBLINK",
    "5 KW 1 Phase Ongrid PV Solar Inverter 2 MPPT-W/O-CT (PVBS5K-M1)-PVBLINK",
    "5 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT5KPro-M1)-PVBLINK",
    "5.5 KW 1 Phase Ongrid PV Solar Inverter 1 MPPT-W/O-CT (PVBSM5.5K-M1)-PVBLINK",
    "6 KW 1 Phase Ongrid PV Solar Inverter 2 MPPT-W/O-CT (PVBS6K-M1)-PVBLINK",
    "6 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT6KPro-M1)-PVBLINK",
    "8 KW 3 Phase Ongrid PV Solar Inverter 3 MPPT (PVBT8K-M1)-PVBLINK",
    "LFP 51.2V X 100AH LV BATTERY PACK (PVBBS-5.12K-LV-M1)",
]


def service_product_key(value):
    """Match listed service models despite spacing, HTML breaks or the (G) suffix."""
    text = re.sub(r"<br\s*/?>", " ", clean_text(value), flags=re.IGNORECASE)
    text = re.sub(r"\(\s*g\s*\)", "", text, flags=re.IGNORECASE)
    return re.sub(r"\s+", "", text).casefold()


def service_product_counts(frame):
    """Count tickets for listed products, grouped by KW, phase and MPPT."""
    product_groups = {
        service_product_key(name): (
            "Battery Pack" if "BATTERY PACK" in name
            else name.split(" Ongrid", 1)[0] + " " + re.search(r"\d+ MPPT\b", name).group(0)
        )
        for name in PRODUCT_CHART_CATALOG
    }
    groups = frame["Product"].map(service_product_key).map(product_groups)
    counts = groups.dropna().value_counts().rename_axis("Product").reset_index(name="Count")
    return counts.sort_values(["Count", "Product"], ascending=[False, True]).reset_index(drop=True)


def normalize_product_kw(value):
    """Return only the approved inverter size label used by the Top Product QTY chart."""
    raw = clean_text(value).upper().replace("K.W", "KW")
    m = re.search(r"(?<!\d)(\d+(?:\.\d+)?)\s*KW\b", raw)
    if not m:
        return pd.NA
    try:
        v = float(m.group(1))
    except Exception:
        return pd.NA
    if v not in APPROVED_PRODUCT_KW:
        return pd.NA
    return f"{v:g} KW"


def normalize_columns(df):
    df = df.copy()
    df.columns = [clean_text(c) for c in df.columns]

    for old, new in OPTIONAL_ALIASES.items():
        if old in df.columns and new not in df.columns:
            df.rename(columns={old: new}, inplace=True)

    seen = {}
    new_columns = []
    for col in df.columns:
        if col not in seen:
            seen[col] = 0
            new_columns.append(col)
        else:
            seen[col] += 1
            new_columns.append(f"{col}_{seen[col]}")
    df.columns = new_columns
    return df


def find_header_row(raw):
    wanted = {
        "Rate", "QTY", "Qty", "Quantity", "Amount",
        "Item Name", "Item", "Item Group",
        "Representative Ref", "Invoice Date", "Ledger Name",
        "Invoice No#", "Customer",
    }

    best_row = 0
    best_score = -1
    scan_limit = min(150, len(raw))

    for row in range(scan_limit):
        values = {clean_text(v) for v in raw.iloc[row].tolist()}
        score = len(values.intersection(wanted))
        if score > best_score:
            best_score = score
            best_row = row

    return best_row, best_score


def load_excel(uploaded_file):
    name = getattr(uploaded_file, "name", "file.xlsx").lower()
    uploaded_file.seek(0)

    if name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file, header=None, dtype=object)
    else:
        raw = pd.read_excel(uploaded_file, header=None, dtype=object)

    if raw is None or raw.empty:
        raise ValueError("The uploaded file is empty.")

    header_row, score = find_header_row(raw)

    if score < 3:
        raise ValueError(
            "Could not identify the Excel header row. "
            "Expected fields such as Rate, QTY, Amount, Item Name, "
            "Item Group, Invoice Date and Ledger Name."
        )

    headers = [clean_text(v) for v in raw.iloc[header_row].tolist()]

    used = {}
    safe_headers = []
    for i, header in enumerate(headers):
        header = header if header else f"Unnamed_{i+1}"
        if header not in used:
            used[header] = 0
            safe_headers.append(header)
        else:
            used[header] += 1
            safe_headers.append(f"{header}_{used[header]}")

    df = raw.iloc[header_row + 1:].copy()
    df.columns = safe_headers
    df = normalize_columns(df)

    df = df.dropna(how="all").dropna(axis=1, how="all").copy()
    return df, header_row, score


def numeric_clean(series):
    s = series.astype(str).str.strip()
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace("₹", "", regex=False)
    s = s.str.replace("$", "", regex=False)
    s = s.str.replace("€", "", regex=False)
    s = s.str.replace("£", "", regex=False)
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    return pd.to_numeric(s, errors="coerce")


def parse_mixed_date_series(series):
    """Parse normal Excel dates, text dates and Excel serial dates safely."""
    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    numeric = pd.to_numeric(series, errors="coerce")
    serial_mask = numeric.notna() & numeric.between(1, 60000)
    if serial_mask.any():
        out.loc[serial_mask] = pd.to_datetime(
            numeric.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce"
        )
    text_mask = ~serial_mask
    if text_mask.any():
        # Pandas can lock onto the first text-date pattern in a Series and then
        # reject later valid formats/month names. format="mixed" prevents that.
        try:
            parsed = pd.to_datetime(
                series.loc[text_mask], errors="coerce", dayfirst=True, format="mixed"
            )
        except (TypeError, ValueError):
            parsed = series.loc[text_mask].map(
                lambda x: pd.to_datetime(x, errors="coerce", dayfirst=True)
            )
        out.loc[text_mask] = parsed
    return out


def fill_invoice_context(df):
    df = df.copy()
    for col in CONTEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].replace(r"^\s*$", np.nan, regex=True)
            df[col] = df[col].ffill()
    return df


def process_data(uploaded_file, company_name="PV-Blink"):
    """
    SALES INVOICE REGISTER — EXACT MANUAL CLEANING WORKFLOW

    This reproduces the user's Excel cleaning process:

    1) Remove the first metadata rows by detecting the real header
       (raw PV-Blink report = first 6 rows removed, header on row 7).
    2) For the RAW report structure only, shift the cells from
       Item Group through Net Amount UP by one row.
    3) Filter Item Group = blank and remove those rows.
    4) Find Representative Ref = blank (do not delete).
    5) Within those rows keep Ledger Name = non-blank.
    6) Put the company name into missing Representative Ref and Representative User.
    7) Release filters.
    8) Fill down:
       Branch, Ledger Name, Invoice Date, Invoice No#,
       Representative Ref, Representative User.
    9) Filter Amount = 0 and remove those rows.

    Important:
    - Representative Ref is the authoritative employee field for Sales analytics.
    - Representative User is retained as source/display data only.
    - No blind drop_duplicates() is used.
    - Rate = 0 is NOT the deletion rule anymore.
    - Amount = 0 is the deletion rule.
    - Already-cleaned Sales files are detected and are NOT shifted again.
    """

    df, header_row, header_score = load_excel(uploaded_file)

    missing = [c for c in CORE_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(
            "Required columns missing:\n\n"
            + "\n".join(f"• {c}" for c in missing)
        )

    original_rows = int(len(df))

    # --------------------------------------------------------
    # Basic text cleanup only — do NOT fill invoice context yet.
    # The exact manual sequence requires PV-Blink assignment first.
    # --------------------------------------------------------
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].map(clean_text)

    df = df.replace(r"^\s*$", np.nan, regex=True)

    def _blank_mask(series):
        return (
            series.isna()
            | series.astype("string").fillna("").str.strip().eq("")
        )

    # --------------------------------------------------------
    # STEP 2
    # Shift Item Group -> Net Amount UP by one row.
    #
    # Raw report signature:
    #   invoice/header rows = Invoice No present + Item Group blank
    #   detail rows         = Invoice No blank   + Item Group present
    #
    # If the file is already cleaned, do not shift again.
    # --------------------------------------------------------
    invoice_blank = _blank_mask(df["Invoice No#"]) if "Invoice No#" in df.columns else pd.Series(True, index=df.index)
    item_group_blank_before = _blank_mask(df["Item Group"])

    raw_invoice_header_rows = (
        (~invoice_blank) & item_group_blank_before
    )
    raw_detail_rows = (
        invoice_blank & (~item_group_blank_before)
    )

    detail_shift_applied = bool(
        raw_invoice_header_rows.any()
        and raw_detail_rows.any()
    )

    shifted_columns_count = 0

    if detail_shift_applied:
        cols = list(df.columns)
        start_idx = cols.index("Item Group")

        if "Net Amount" in cols:
            end_idx = cols.index("Net Amount")
        else:
            # Fallback for a report version without Net Amount:
            # shift every detail-side column from Item Group onward.
            end_idx = len(cols) - 1

        detail_cols = cols[start_idx:end_idx + 1]
        shifted_columns_count = len(detail_cols)

        # Exact Excel behavior: delete first detail-side cells and SHIFT UP.
        df.loc[:, detail_cols] = df.loc[:, detail_cols].shift(-1)

    # --------------------------------------------------------
    # STEP 3
    # Item Group blank -> DELETE the complete row.
    # --------------------------------------------------------
    item_group_blank_mask = _blank_mask(df["Item Group"])
    item_group_blank_removed = int(item_group_blank_mask.sum())
    df = df.loc[~item_group_blank_mask].copy()

    # --------------------------------------------------------
    # STEPS 4, 5, 6
    # Representative Ref blank + Ledger Name non-blank
    # -> set BOTH representative fields to PV-Blink.
    # --------------------------------------------------------
    if "Representative Ref" not in df.columns:
        df["Representative Ref"] = pd.Series(index=df.index, dtype="object")
    if "Representative User" not in df.columns:
        df["Representative User"] = pd.Series(index=df.index, dtype="object")

    rep_blank = _blank_mask(df["Representative Ref"])
    ledger_nonblank = ~_blank_mask(df["Ledger Name"])
    pv_blink_mask = rep_blank & ledger_nonblank

    pv_blink_assigned = int(pv_blink_mask.sum())

    df.loc[pv_blink_mask, "Representative Ref"] = company_name
    df.loc[pv_blink_mask, "Representative User"] = company_name

    # --------------------------------------------------------
    # STEP 8
    # Excel Ctrl+G / fill-down equivalent for invoice context.
    # --------------------------------------------------------
    fill_columns = [
        "Branch",
        "Ledger Name",
        "Invoice Date",
        "Invoice No#",
        "Representative Ref",
        "Representative User",
    ]

    for col in fill_columns:
        if col in df.columns:
            df[col] = (
                df[col]
                .replace(r"^\s*$", np.nan, regex=True)
                .ffill()
            )

    # --------------------------------------------------------
    # Numeric cleaning AFTER the detail-column shift.
    # --------------------------------------------------------
    for col in ["Rate", "QTY", "Amount"]:
        if col in df.columns:
            df[col] = numeric_clean(df[col])

    for col in ["CGST", "SGST", "IGST", "CESS", "Net Amount"]:
        if col in df.columns:
            df[col] = numeric_clean(df[col])

    # --------------------------------------------------------
    # STEP 9
    # Amount = 0 -> DELETE.
    # This replaces the old "Rate = 0" deletion logic.
    # --------------------------------------------------------
    amount_zero_mask = df["Amount"].eq(0)
    amount_zero_removed = int(amount_zero_mask.sum())
    df = df.loc[~amount_zero_mask].copy()

    # --------------------------------------------------------
    # Date parsing and minimal safety validation.
    # These do not alter the confirmed PV-Blink manual result
    # for the supplied Sales Invoice workbook.
    # --------------------------------------------------------
    df["Invoice Date"] = parse_mixed_date_series(df["Invoice Date"])

    invalid_date_mask = df["Invoice Date"].isna()
    invalid_dates_removed = int(invalid_date_mask.sum())
    df = df.loc[~invalid_date_mask].copy()

    invalid_core_mask = (
        df["QTY"].isna()
        | df["Amount"].isna()
        | df["Item Name"].isna()
        | df["Ledger Name"].isna()
    )
    core_removed = int(invalid_core_mask.sum())
    df = df.loc[~invalid_core_mask].copy()

    # IMPORTANT:
    # Do NOT blindly remove duplicate-looking commercial lines.
    # Two valid invoice item rows can look identical.
    duplicate_removed = 0

    # --------------------------------------------------------
    # Dashboard helper columns.
    # --------------------------------------------------------
    df["Year"] = df["Invoice Date"].dt.year.astype("Int64")
    df["Quarter"] = "Q" + df["Invoice Date"].dt.quarter.astype(str)
    df["Month"] = df["Invoice Date"].dt.month_name()
    df["Month No"] = df["Invoice Date"].dt.month.astype("Int64")
    df["Year-Month"] = df["Invoice Date"].dt.to_period("M").astype(str)
    df["Day"] = df["Invoice Date"].dt.day.astype("Int64")
    df["Week"] = df["Invoice Date"].dt.isocalendar().week.astype("Int64")

    # Helper used ONLY for normalized Top Product Quantity charts.
    df["Product KW"] = df["Item Name"].map(normalize_product_kw)

    if "Rate" not in df.columns:
        df["Rate"] = np.nan
    gross_value = df["Rate"].fillna(0) * df["QTY"].fillna(0)
    df["Discount Amount"] = (
        gross_value - df["Amount"].fillna(0)
    ).clip(lower=0)

    df["Discount %"] = np.where(
        gross_value > 0,
        (df["Discount Amount"] / gross_value) * 100,
        0,
    )

    preferred = [
        "Branch", "Ledger Name", "Invoice Date", "Invoice No#",
        "Representative Ref", "Representative User",
        "Item Group", "Item Name",
        "Hsn Code", "Rate", "QTY", "Amount",
        "Discount Amount", "Discount %",
        "CGST", "SGST", "IGST", "CESS", "Net Amount",
        "Narration",
        "Year", "Quarter", "Month", "Month No",
        "Year-Month", "Day", "Week",
    ]

    ordered = [c for c in preferred if c in df.columns]
    ordered += [c for c in df.columns if c not in ordered]
    df = df[ordered].reset_index(drop=True)

    summary = {
        "original": int(original_rows),

        # New exact-cleaning audit fields.
        "metadata_rows_removed": int(header_row),
        "detail_shift_applied": int(detail_shift_applied),
        "shifted_columns_count": int(shifted_columns_count),
        "item_group_blank_removed": int(item_group_blank_removed),
        "pv_blink_assigned": int(pv_blink_assigned),
        "amount_zero_removed": int(amount_zero_removed),

        # Existing compatibility keys used elsewhere in the dashboard.
        # rate_removed now mirrors Amount=0 removal because the cleaning
        # business rule has changed from Rate=0 to Amount=0.
        "rate_removed": int(amount_zero_removed),
        "invalid_dates_removed": int(invalid_dates_removed),
        "core_removed": int(core_removed),
        "duplicates": int(duplicate_removed),

        "final": int(len(df)),
        "header_row": int(header_row + 1),
        "header_score": int(header_score),
    }

    return df, summary


def safe_unique(df, column):
    if column not in df.columns:
        return []
    values = (
        df[column]
        .dropna()
        .astype(str)
        .map(clean_text)
    )
    values = values[values != ""]
    return sorted(values.unique().tolist(), key=lambda x: x.lower())


def _clean_plotly_text(value):
    """Return safe display text for Plotly layout/trace fields."""
    if value is None:
        return ""
    try:
        text = str(value).strip()
    except Exception:
        return ""
    if text.casefold() in {"", "undefined", "none", "null", "nan", "<na>"}:
        return ""
    return text


def clean_plotly_fig(fig, height=350, is_horizontal=False, max_val=None):
    """
    V27 universal enterprise Plotly formatter.

    Fixes:
    - literal `undefined` appearing above charts
    - undefined/None axis, legend and annotation text
    - cramped labels in normal/full-screen mode
    - inconsistent Plotly canvas styling
    """
    margin_left = 218 if is_horizontal else 66
    margin_right = 220 if is_horizontal else 62

    # Read the incoming title safely. Plotly may hold a title object whose text
    # is None/undefined even when no visible title was requested.
    try:
        incoming_title = _clean_plotly_text(fig.layout.title.text)
    except Exception:
        incoming_title = ""

    # Explicit blank title object is important. Using title=None can serialize
    # differently across Plotly versions and may render as literal "undefined".
    safe_title = (
        dict(
            text=incoming_title,
            x=0.015,
            xanchor="left",
            y=0.985,
            yanchor="top",
            font=dict(size=15, color="#0f172a", family="Plus Jakarta Sans"),
        )
        if incoming_title
        else dict(text="")
    )

    fig.update_layout(
        template="plotly_white",
        title=safe_title,
        height=height,
        autosize=True,
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(color="#334155", family="Plus Jakarta Sans", size=11),
        margin=dict(
            l=margin_left,
            r=margin_right,
            t=58 if incoming_title else 26,
            b=64,
            pad=2,
        ),
        legend=dict(
            bgcolor="rgba(255,255,255,0.94)",
            bordercolor="rgba(226,232,240,.85)",
            borderwidth=0,
            font=dict(color="#334155", size=10.5),
            title=dict(text=""),
        ),
        hoverlabel=dict(
            bgcolor="#0f172a",
            bordercolor="#0f172a",
            font_color="#ffffff",
            font_family="Plus Jakarta Sans",
            font_size=11,
        ),
        uniformtext_minsize=9,
        uniformtext_mode="show",
        transition_duration=0,
        hovermode="closest",
        modebar=dict(
            bgcolor="rgba(255,255,255,.88)",
            color="#64748b",
            activecolor=company_theme("#f97316"),
        ),
    )

    # Axis titles are intentionally externalized in the dashboard UI.
    fig.update_xaxes(
        title=dict(text=""),
        showgrid=True,
        gridcolor="#edf2f7",
        gridwidth=1,
        tickfont=dict(color="#475569", size=10.5),
        zeroline=False,
        automargin=True,
        showline=False,
        ticks="",
    )
    fig.update_yaxes(
        title=dict(text=""),
        showgrid=not is_horizontal,
        gridcolor="#edf2f7",
        gridwidth=1,
        tickfont=dict(color="#334155", size=10.5),
        zeroline=False,
        automargin=True,
        showline=False,
        ticks="",
    )

    if is_horizontal and max_val is not None:
        try:
            max_number = float(max_val)
            if max_number > 0:
                fig.update_xaxes(range=[0, max_number * 1.62])
        except Exception:
            pass

    # Remove bad annotation strings such as "undefined".
    try:
        annotations = []
        for ann in list(fig.layout.annotations or []):
            cleaned = _clean_plotly_text(getattr(ann, "text", ""))
            if cleaned:
                ann.text = cleaned
                annotations.append(ann)
        fig.update_layout(annotations=annotations)
    except Exception:
        pass

    # Sanitize traces comprehensively.
    for trace in fig.data:
        try:
            trace.name = _clean_plotly_text(getattr(trace, "name", ""))
        except Exception:
            pass

        try:
            legendgroup = _clean_plotly_text(getattr(trace, "legendgroup", ""))
            if getattr(trace, "legendgroup", None) is not None:
                trace.legendgroup = legendgroup
        except Exception:
            pass

        try:
            ht = getattr(trace, "hovertemplate", None)
            if isinstance(ht, str):
                trace.hovertemplate = ht.replace("undefined", "").replace("Undefined", "")
        except Exception:
            pass

        # Clean chart text arrays too, without changing numeric labels.
        try:
            trace_text = getattr(trace, "text", None)
            if trace_text is not None and not isinstance(trace_text, str):
                cleaned_text = []
                for item in trace_text:
                    if isinstance(item, str):
                        cleaned_text.append(_clean_plotly_text(item))
                    else:
                        cleaned_text.append(item)
                trace.text = cleaned_text
            elif isinstance(trace_text, str):
                trace.text = _clean_plotly_text(trace_text)
        except Exception:
            pass

        if hasattr(trace, "cliponaxis"):
            try:
                trace.cliponaxis = False
            except Exception:
                pass

        if getattr(trace, "type", "") == "bar":
            try:
                trace.constraintext = "none"
            except Exception:
                pass
            try:
                if trace.text is not None:
                    trace.textfont = dict(
                        size=10.5,
                        color="#0f172a",
                        family="Plus Jakarta Sans",
                    )
                    if not getattr(trace, "textposition", None):
                        trace.textposition = "outside"
            except Exception:
                pass

        elif getattr(trace, "type", "") == "scatter":
            try:
                if trace.text is not None:
                    trace.textfont = dict(
                        size=10.5,
                        color="#0f172a",
                        family="Plus Jakarta Sans",
                    )
                    trace.cliponaxis = False
            except Exception:
                pass

    # Final defensive pass: force blank text, never a null title object.
    try:
        if not _clean_plotly_text(fig.layout.title.text):
            fig.layout.title.text = ""
    except Exception:
        fig.update_layout(title=dict(text=""))

    return fig

def kpi_card_html(icon, label, value, badge_text="", full_value=None):
    """
    Card keeps the compact dashboard value.
    Move/hold the cursor on the card to see the exact full value.
    """
    exact = value if full_value is None else full_value
    tooltip = _html_attr(f"{label}: {exact}")

    return f'''
    <div class="pv-kpi-card" title="{tooltip}" style="cursor:help;">
        <div class="pv-kpi-header">
            <div class="pv-kpi-icon">{icon}</div>
            <div class="pv-kpi-label">{label}</div>
        </div>
        <div class="pv-kpi-value">{value}</div>
        <div class="pv-kpi-badge">{badge_text}</div>
    </div>
    '''


def make_excel_bytes(cleaned_df, filtered_df, summary):
    output = BytesIO()

    quality_rows = [
        ["Metric", "Value"],
        ["Original Source Rows", summary.get("original", 0)],
        ["Top Metadata Rows Removed", summary.get("metadata_rows_removed", 0)],
        ["Detail Columns Shift-Up Applied", summary.get("detail_shift_applied", 0)],
        ["Blank Item Group Rows Removed", summary.get("item_group_blank_removed", 0)],
        ["Company Representative Rows Assigned", summary.get("pv_blink_assigned", 0)],
        ["Amount = 0 Rows Removed", summary.get("amount_zero_removed", summary.get("rate_removed", 0))],
        ["Invalid Date Rows Removed", summary.get("invalid_dates_removed", 0)],
        ["Missing Core Field Rows Removed", summary.get("core_removed", 0)],
        ["Duplicate Rows Removed", summary.get("duplicates", 0)],
        ["Final Cleaned Rows", summary.get("final", 0)],
        ["Detected Header Row", summary.get("header_row", 0)],
        ["Header Detection Score", summary.get("header_score", 0)],
    ]
    quality_df = pd.DataFrame(quality_rows[1:], columns=quality_rows[0])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cleaned_df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
        filtered_df.to_excel(writer, sheet_name="Filtered_Data", index=False)
        quality_df.to_excel(writer, sheet_name="Data_Quality", index=False)

    output.seek(0)
    wb = load_workbook(output)

    header_fill = PatternFill("solid", fgColor="0F172A")
    header_font = Font(color="FFFFFF", bold=True, name="Calibri")
    thin = Side(style="thin", color="F97316")
    border = Border(bottom=thin)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        ws.row_dimensions[1].height = 28

        for col_cells in ws.columns:
            col_letter = get_column_letter(col_cells[0].column)
            header = str(col_cells[0].value or "")
            max_len = len(header)

            for cell in col_cells[1:200]:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))

            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 44)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if hasattr(cell.value, "year") and hasattr(cell.value, "month"):
                    cell.number_format = "dd-mmm-yyyy"
                cell.alignment = Alignment(vertical="center", wrap_text=False)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ============================================================
# SERVICE DATA PROCESSING
# ============================================================

def normalize_service_columns(df):
    df = df.copy()
    cols = []
    seen = {}
    for i, c in enumerate(df.columns):
        c = clean_text(c)
        if not c:
            c = f"Unnamed_{i+1}"
        if c in seen:
            seen[c] += 1
            c = f"{c}_{seen[c]}"
        else:
            seen[c] = 0
        cols.append(c)
    df.columns = cols
    return df


def clean_service_file(uploaded_file):
    uploaded_file.seek(0)
    file_name = getattr(uploaded_file, "name", "Service_Report.xlsx")
    raw = pd.read_excel(uploaded_file, header=None, dtype=object)
    if raw is None or raw.shape[0] <= 6:
        raise ValueError(f"{file_name}: file has no data after the first 6 rows.")

    header = [clean_text(x) for x in raw.iloc[6].tolist()]
    used = {}
    safe = []
    for i, c in enumerate(header):
        c = c if c else f"Unnamed_{i+1}"
        if c in used:
            used[c] += 1
            c = f"{c}_{used[c]}"
        else:
            used[c] = 0
        safe.append(c)

    df = raw.iloc[7:].copy()
    df.columns = safe
    df = normalize_service_columns(df)

    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(lambda x: clean_text(x) if not pd.isna(x) else np.nan)

    for c in SERVICE_COLUMNS:
        if c not in df.columns:
            df[c] = np.nan

    df["Source File"] = file_name
    df["MTCE Status Raw"] = df["MTCE Status"].fillna("")
    mtce_clean = df["MTCE Status Raw"].map(lambda x: clean_text(x).lower())

    df["Status"] = np.where(
        mtce_clean.str.contains("in-progress|in progress|open|pending", regex=True, na=False),
        "Open",
        np.where(
            mtce_clean.str.contains("closed|close", regex=True, na=False),
            "Closed",
            np.where(mtce_clean.eq(""), "Unknown", mtce_clean.str.title())
        )
    )

    for c in ["Closed Date", "Follow Date", "Created Date", "Date"]:
        if c in df.columns:
            df[c + "_Parsed"] = parse_mixed_date_series(df[c])

    df["Ticket Valid"] = df["Ticket No"].notna() & df["Ticket No"].astype(str).str.strip().ne("")
    # Repeated follow-up rows for the same otherwise identical ticket are one
    # service query. Retain the latest follow-up without merging differing statuses.
    if "Follow Date_Parsed" in df.columns:
        ticket_columns = [c for c in df.columns if c not in ["Follow Date", "Follow Date_Parsed"]]
        df = (
            df.sort_values("Follow Date_Parsed", na_position="first", kind="stable")
            .drop_duplicates(subset=ticket_columns, keep="last")
            .sort_index()
        )
    return df


def load_service_files(files):
    frames = []
    details = []
    for f in files or []:
        try:
            d = clean_service_file(f)
            frames.append(d)
            details.append({
                "File": getattr(f, "name", "Service.xlsx"),
                "Rows After Header": len(d),
                "Columns": len(d.columns),
                "Open": int((d["Status"] == "Open").sum()),
                "Closed": int((d["Status"] == "Closed").sum()),
                "Blank Ticket": int((~d["Ticket Valid"]).sum()),
                "Cleaning": "First 6 rows removed; blank cells preserved"
            })
        except Exception as e:
            details.append({"File": getattr(f, "name", "Service.xlsx"), "Error": str(e)})

    if not frames:
        return pd.DataFrame(), pd.DataFrame(details)

    out = pd.concat(frames, ignore_index=True, sort=False)
    out = out.drop_duplicates(keep="first")
    return out, pd.DataFrame(details)

# ============================================================
# TARGETS & PERFORMANCE EDITOR
# ============================================================

def _target_employee_key(value):
    """
    Matching key for pasted employee names.
    Also matches a source like:
    Nilesh Soni(nileshsoni@pvblink.com)
    with pasted text:
    Nilesh Soni
    """
    text = clean_text(value).lower()
    text = re.sub(r"\([^)]*@[^)]*\)", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _parse_target_amount(value):
    """Parse target values such as 5000000, ₹50,00,000, 50L, 1.25Cr."""
    if value is None:
        return np.nan

    raw = clean_text(value).lower()
    if not raw:
        return np.nan

    raw = (
        raw.replace("₹", "")
           .replace(",", "")
           .replace("rs.", "")
           .replace("rs", "")
           .strip()
    )

    multiplier = 1.0

    if raw.endswith("crore"):
        multiplier = 10_000_000
        raw = raw[:-5].strip()
    elif raw.endswith("cr"):
        multiplier = 10_000_000
        raw = raw[:-2].strip()
    elif raw.endswith("lakh"):
        multiplier = 100_000
        raw = raw[:-4].strip()
    elif raw.endswith("lac"):
        multiplier = 100_000
        raw = raw[:-3].strip()
    elif raw.endswith("l"):
        multiplier = 100_000
        raw = raw[:-1].strip()
    elif raw.endswith("k"):
        multiplier = 1_000
        raw = raw[:-1].strip()

    try:
        return float(raw) * multiplier
    except Exception:
        return np.nan


def _parse_pasted_targets(text, employee_names):
    """
    Accepts copied Excel rows:
        Employee<TAB>Target
    Also accepts comma / semicolon separated values.
    """
    if not text or not clean_text(text):
        return {}, []

    employees = [str(x).strip() for x in employee_names if str(x).strip()]

    exact_lookup = {
        clean_text(emp).lower(): emp
        for emp in employees
    }
    key_lookup = {}
    for emp in employees:
        key_lookup.setdefault(_target_employee_key(emp), emp)

    updates = {}
    unmatched = []

    for raw_line in str(text).splitlines():
        line = raw_line.strip()
        if not line:
            continue

        # Excel copy/paste normally uses TAB.
        if "\t" in line:
            parts = [p.strip() for p in line.split("\t")]
        elif ";" in line:
            parts = [p.strip() for p in line.split(";")]
        elif "," in line:
            # Split on the LAST comma only when possible because
            # target amounts may also use Indian comma grouping.
            m = re.match(r"^(.*?)[,\s]+([₹Rsrs\.\d,]+(?:\.\d+)?\s*(?:cr|crore|l|lac|lakh|k)?)$", line, re.I)
            if m:
                parts = [m.group(1).strip(), m.group(2).strip()]
            else:
                parts = [p.strip() for p in line.split(",", 1)]
        else:
            # Fallback: last whitespace-delimited token = target.
            m = re.match(r"^(.*?)\s+([₹Rsrs\.\d,]+(?:\.\d+)?\s*(?:cr|crore|l|lac|lakh|k)?)$", line, re.I)
            if not m:
                continue
            parts = [m.group(1).strip(), m.group(2).strip()]

        if len(parts) < 2:
            continue

        pasted_emp = parts[0].strip()
        pasted_target = parts[-1].strip()

        # Skip common header row.
        if _target_employee_key(pasted_emp) in {
            "employee", "emp name", "emp_name", "representative",
            "sales representative", "name"
        }:
            continue

        amount = _parse_target_amount(pasted_target)
        if pd.isna(amount):
            continue

        exact_match = exact_lookup.get(clean_text(pasted_emp).lower())
        matched_emp = exact_match or key_lookup.get(_target_employee_key(pasted_emp))

        if matched_emp:
            updates[matched_emp] = max(float(amount), 0.0)
        else:
            unmatched.append(pasted_emp)

    return updates, unmatched


def canonical_sales_employee(value):
    """
    Canonical employee name used ONLY for Sales Target Performance.

    Examples:
      KAUSHAL THAKKAR(kaushal@pvblink.com)
      KAUSHAL THAKKAR(kaushalthakkar9298@gmail.com)
    both become:
      Kaushal Thakkar

    This prevents the same employee appearing twice because the source
    contains different email IDs / case formatting.
    """
    text = clean_text(value)

    if not text:
        return ""

    # Remove email/address text inside parentheses.
    text = re.sub(
        r"\s*\([^)]*@[^)]*\)\s*",
        " ",
        text,
        flags=re.I,
    )

    # Also remove a plain email if one appears outside parentheses.
    text = re.sub(
        r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b",
        " ",
        text,
        flags=re.I,
    )

    text = re.sub(r"\s+", " ", text).strip()

    if not text:
        return ""

    # Consistent display casing.
    return text.title()


def canonical_sales_employee_key(value):
    """Stable matching key for canonical employee names."""
    return re.sub(
        r"[^a-z0-9]+",
        "",
        canonical_sales_employee(value).lower(),
    )


# ============================================================
# SALES TARGET PERFORMANCE — EMPLOYEE ROSTER OVERRIDES
# ONLY affects the Sales Target Performance module.
# ============================================================
SALES_TARGET_EMPLOYEES_REMOVE = {
    canonical_sales_employee_key("Atulraj Rajput"),
    canonical_sales_employee_key("Param Patel"),
}

SALES_TARGET_EMPLOYEES_ADD = [
    "Valand Krina",
    "Ashif P A",
    "Aditya Yogi",
    "Mahadevappa Jakati",
    "Mahadeshwara Setty",
]


def sales_target_roster(employee_names):
    """
    Build the Sales Target Performance roster only.

    - Keep all employees found in Sales Invoice Register.
    - Remove Atulraj Rajput and Param Patel from Target Performance only.
    - Always add the five requested employees.
    - De-duplicate employee/email/case variants.
    """
    combined = list(employee_names or []) + SALES_TARGET_EMPLOYEES_ADD

    output = []
    seen = set()

    for value in combined:
        employee = canonical_sales_employee(value)
        key = canonical_sales_employee_key(employee)

        if not employee or not key:
            continue

        if employee.lower() in {"-", "nan", "none", "<na>", "unassigned"}:
            continue

        if key in SALES_TARGET_EMPLOYEES_REMOVE:
            continue

        if key in seen:
            continue

        seen.add(key)
        output.append(employee)

    return sorted(output, key=lambda x: x.lower())


def parse_target_amount(value):
    """
    Paste-friendly target parser.
    Accepts:
      80000000
      8,00,00,000
      ₹8,00,00,000
      8Cr / 8 Cr
      75L / 75 lakh
    """
    if value is None:
        return 0.0

    text = clean_text(value).lower()
    if not text:
        return 0.0

    text = (
        text.replace("₹", "")
        .replace(",", "")
        .replace(" ", "")
    )

    multiplier = 1.0

    if text.endswith("crore"):
        multiplier = 10_000_000
        text = text[:-5]
    elif text.endswith("cr"):
        multiplier = 10_000_000
        text = text[:-2]
    elif text.endswith("lakhs"):
        multiplier = 100_000
        text = text[:-5]
    elif text.endswith("lakh"):
        multiplier = 100_000
        text = text[:-4]
    elif text.endswith("lac"):
        multiplier = 100_000
        text = text[:-3]
    elif text.endswith("l"):
        multiplier = 100_000
        text = text[:-1]

    try:
        return max(float(text) * multiplier, 0.0)
    except Exception:
        return 0.0


def _target_years_for_editor(selected_start=None, selected_end=None):
    current_year = pd.Timestamp.today().year

    years = set(range(current_year - 5, current_year + 4))

    if selected_start is not None:
        years.add(pd.Timestamp(selected_start).year)

    if selected_end is not None:
        years.add(pd.Timestamp(selected_end).year)

    return sorted(years)


def _normalise_target_rule(rule):
    """
    Normalize one employee target rule.

    FINAL supported types:
      - Year
      - Date Range
    """
    if not isinstance(rule, dict):
        return None

    employee_key = clean_text(rule.get("employee_key", ""))
    basis = clean_text(rule.get("basis", "Year")).title()

    if basis not in {"Year", "Date"}:
        return None

    target = parse_target_amount(rule.get("target", 0))

    try:
        year = int(rule.get("year")) if rule.get("year") is not None else None
    except Exception:
        year = None

    start_date = rule.get("start_date")
    end_date = rule.get("end_date")

    if start_date is not None:
        try:
            start_date = pd.Timestamp(start_date).date()
        except Exception:
            start_date = None

    if end_date is not None:
        try:
            end_date = pd.Timestamp(end_date).date()
        except Exception:
            end_date = None

    if start_date and end_date and end_date < start_date:
        start_date, end_date = end_date, start_date

    if basis == "Year" and year is None:
        return None

    if basis == "Date" and (start_date is None or end_date is None):
        return None

    return {
        "employee_key": employee_key,
        "basis": basis,
        "year": year,
        "start_date": start_date,
        "end_date": end_date,
        "target": float(target),
        "locked": bool(rule.get("locked", True)),
    }


def _target_rule_id(rule):
    rule = _normalise_target_rule(rule)

    if rule is None:
        return ""

    if rule["basis"] == "Year":
        return (
            f"{rule['employee_key']}|YEAR|"
            f"{int(rule['year'])}"
        )

    return (
        f"{rule['employee_key']}|DATE|"
        f"{rule['start_date'].isoformat()}|"
        f"{rule['end_date'].isoformat()}"
    )


def _target_rule_label(rule):
    rule = _normalise_target_rule(rule)

    if rule is None:
        return ""

    if rule["basis"] == "Year":
        return str(rule["year"])

    return (
        f"{rule['start_date'].strftime('%d/%m/%Y')} to "
        f"{rule['end_date'].strftime('%d/%m/%Y')}"
    )


def _rule_date_bounds(rule):
    rule = _normalise_target_rule(rule)

    if rule is None:
        return None, None

    if rule["basis"] == "Year":
        return (
            pd.Timestamp(rule["year"], 1, 1).date(),
            pd.Timestamp(rule["year"], 12, 31).date(),
        )

    return rule["start_date"], rule["end_date"]


def _ranges_overlap(start_a, end_a, start_b, end_b):
    return max(start_a, start_b) <= min(end_a, end_b)


def _employee_target_for_date_range(
    target_rules,
    employee_key,
    selected_start,
    selected_end,
):
    """
    FINAL target denominator logic.

    IMPORTANT:
    Achievement % = Achievement / FIXED TARGET * 100

    YEAR target:
      If the selected dashboard date range touches that year,
      use the FULL fixed Year target.
      Do NOT prorate it by months or days.

      Example:
        Harsh 2026 Year Target = ₹8Cr
        Dashboard selected dates = 01/04/2026 to 30/06/2026
        Achievement = ₹3,86,86,380

        Target denominator = ₹8,00,00,000
        Achievement % = 3,86,86,380 / 8,00,00,000 * 100
                      = 48.36%

    DATE RANGE target:
      If the selected dashboard date range overlaps a fixed Date Range
      target for the employee, use the FULL fixed Date Range target.
      Do NOT prorate that target.

    Priority:
      Date Range target overrides Year target for the overlapping analysis.

    Multiple years:
      If the selected dashboard period spans multiple years and no Date Range
      override applies, add the full fixed Year targets for those touched years.
    """
    if selected_start is None or selected_end is None:
        return 0.0

    selected_start = pd.Timestamp(
        selected_start
    ).date()

    selected_end = pd.Timestamp(
        selected_end
    ).date()

    if selected_end < selected_start:
        selected_start, selected_end = (
            selected_end,
            selected_start,
        )

    emp_rules = []

    for raw_rule in list(target_rules or []):
        rule = _normalise_target_rule(
            raw_rule
        )

        if (
            rule is not None
            and rule["employee_key"] == employee_key
            and float(rule["target"]) > 0
            and rule.get("locked", True)
        ):
            emp_rules.append(rule)

    if not emp_rules:
        return 0.0

    # --------------------------------------------------------
    # 1) DATE RANGE target has first priority.
    # Use FULL fixed target if selected dashboard dates overlap it.
    # --------------------------------------------------------
    matching_date_rules = []

    for rule in emp_rules:
        if rule["basis"] != "Date":
            continue

        if _ranges_overlap(
            selected_start,
            selected_end,
            rule["start_date"],
            rule["end_date"],
        ):
            matching_date_rules.append(
                rule
            )

    if matching_date_rules:
        # Overlapping Date Range rules are blocked during target creation,
        # so simply add all matching fixed ranges if the dashboard selection
        # intentionally spans multiple separate saved target periods.
        return float(
            sum(
                float(rule["target"])
                for rule in matching_date_rules
            )
        )

    # --------------------------------------------------------
    # 2) Otherwise use FULL YEAR targets for every year touched
    # by the dashboard Date Range.
    # NO PRORATION.
    # --------------------------------------------------------
    selected_years = set(
        range(
            int(selected_start.year),
            int(selected_end.year) + 1,
        )
    )

    year_target_total = 0.0

    for rule in emp_rules:
        if (
            rule["basis"] == "Year"
            and int(rule["year"])
            in selected_years
        ):
            year_target_total += float(
                rule["target"]
            )

    return float(year_target_total)


def target_editor(
    employee_names,
    selected_start=None,
    selected_end=None,
):
    """
    FINAL compact employee-wise target setup.

    Supported:
      - Year
      - Date Range

    Workflow:
      1. Select Employee
      2. Select Year OR Date Range
      3. Paste Target
      4. Fix Target
      5. Change / Remove later

    Percentage uses:
      Achievement % = Achievement / Target × 100
    """
    names = sales_target_roster(
        employee_names
    )

    if not names:
        st.info(
            "No Sales employees are available for Target setup."
        )
        return pd.DataFrame(
            columns=["Employee", "Target"]
        )

    # Reset only the obsolete experimental target modes once.
    mode_version = "YEAR_AND_DATE_RANGE_V1"

    if (
        st.session_state.get(
            "pv_target_mode_version"
        )
        != mode_version
    ):
        st.session_state[
            "pv_target_mode_version"
        ] = mode_version

        old_rules = st.session_state.get(
            "pv_target_rules",
            [],
        )

        kept_rules = []

        for raw_rule in list(old_rules or []):
            rule = _normalise_target_rule(raw_rule)

            if rule is not None:
                kept_rules.append(rule)

        st.session_state[
            "pv_target_rules"
        ] = kept_rules

        for old_key in [
            "pv_annual_target_store",
            "pv_annual_target_locks",
            "pv_target_range_store",
            "pv_target_range_locks",
            "pv_target_schedule_store",
            "pv_sales_targets_locked",
        ]:
            st.session_state.pop(
                old_key,
                None,
            )

    target_rules = [
        r
        for r in (
            _normalise_target_rule(x)
            for x in st.session_state.get(
                "pv_target_rules",
                [],
            )
        )
        if r is not None
    ]

    st.caption(
        "Select Employee → Year / Date Range → paste Target → Fix Target."
    )

    c1, c2 = st.columns([1.35, 1.0])

    with c1:
        selected_employee = st.selectbox(
            "👤 Employee",
            options=names,
            key="pv_target_setup_employee",
        )

    with c2:
        target_basis_ui = st.selectbox(
            "🎯 Target Type",
            options=["Year", "Date Range"],
            key="pv_target_setup_basis",
        )

    employee_key = canonical_sales_employee_key(
        selected_employee
    )

    years = _target_years_for_editor(
        selected_start,
        selected_end,
    )

    default_year = (
        pd.Timestamp(selected_start).year
        if selected_start is not None
        else pd.Timestamp.today().year
    )

    selected_year = None
    from_date = None
    to_date = None

    if target_basis_ui == "Year":
        selected_year = st.selectbox(
            "📅 Target Year",
            options=years,
            index=(
                years.index(default_year)
                if default_year in years
                else 0
            ),
            key="pv_target_setup_year",
        )

        draft_rule = {
            "employee_key": employee_key,
            "basis": "Year",
            "year": int(selected_year),
            "start_date": None,
            "end_date": None,
            "target": 0.0,
            "locked": True,
        }

    else:
        if (
            selected_start is not None
            and selected_end is not None
        ):
            default_start = pd.Timestamp(
                selected_start
            ).date()
            default_end = pd.Timestamp(
                selected_end
            ).date()
        else:
            default_start = pd.Timestamp.today().date()
            default_end = default_start

        d1, d2 = st.columns(2)

        with d1:
            from_date = st.date_input(
                "📅 From Date",
                value=default_start,
                key="pv_target_setup_from_date",
                format="DD/MM/YYYY",
            )

        with d2:
            to_date = st.date_input(
                "📅 To Date",
                value=default_end,
                key="pv_target_setup_to_date",
                format="DD/MM/YYYY",
            )

        from_date = pd.Timestamp(from_date).date()
        to_date = pd.Timestamp(to_date).date()

        if to_date < from_date:
            from_date, to_date = (
                to_date,
                from_date,
            )

            st.warning(
                "To Date was before From Date, so the dates were corrected automatically."
            )

        draft_rule = {
            "employee_key": employee_key,
            "basis": "Date",
            "year": None,
            "start_date": from_date,
            "end_date": to_date,
            "target": 0.0,
            "locked": True,
        }

    draft_id = _target_rule_id(
        draft_rule
    )

    existing_rule = next(
        (
            r
            for r in target_rules
            if _target_rule_id(r)
            == draft_id
        ),
        None,
    )

    existing_locked = bool(
        existing_rule is not None
        and existing_rule.get(
            "locked",
            True,
        )
    )

    widget_key = (
        "pv_target_value_"
        + re.sub(
            r"[^a-zA-Z0-9]+",
            "_",
            draft_id,
        )[:120]
    )

    if widget_key not in st.session_state:
        existing_target = (
            float(
                existing_rule["target"]
            )
            if existing_rule
            else 0.0
        )

        st.session_state[
            widget_key
        ] = (
            ""
            if existing_target <= 0
            else format_indian_number_full(
                existing_target,
                decimals=0,
            )
        )

    target_text = st.text_input(
        "Target Amount",
        key=widget_key,
        disabled=existing_locked,
        placeholder="Paste 8Cr / 75L / 80000000",
        help=(
            "Year = annual target. "
            "Date Range = total target for that exact period."
        ),
    )

    entered_target = parse_target_amount(
        target_text
    )

    b1, b2, b3 = st.columns(
        [1.0, 1.0, 0.82]
    )

    with b1:
        fix_clicked = st.button(
            "🔒 Fix Target",
            width="stretch",
            disabled=existing_locked,
            key=f"fix_{widget_key}",
        )

    with b2:
        change_clicked = st.button(
            "✏️ Change",
            width="stretch",
            disabled=not existing_locked,
            key=f"change_{widget_key}",
        )

    with b3:
        remove_clicked = st.button(
            "🗑 Remove",
            width="stretch",
            disabled=existing_rule is None,
            key=f"remove_{widget_key}",
        )

    if fix_clicked:
        if entered_target <= 0:
            st.error(
                "Enter a Target greater than ₹0 before fixing."
            )

        else:
            # Prevent overlapping Date Range targets for same employee.
            # Year + Date Range is allowed because Date Range overrides Year.
            overlap_error = []

            if draft_rule["basis"] == "Date":
                for rule in target_rules:
                    if (
                        rule["employee_key"]
                        != employee_key
                    ):
                        continue

                    if (
                        rule["basis"]
                        != "Date"
                    ):
                        continue

                    if (
                        _target_rule_id(rule)
                        == draft_id
                    ):
                        continue

                    if _ranges_overlap(
                        draft_rule["start_date"],
                        draft_rule["end_date"],
                        rule["start_date"],
                        rule["end_date"],
                    ):
                        overlap_error.append(
                            rule
                        )

            if overlap_error:
                st.error(
                    "This employee already has an overlapping Date Range Target: "
                    + ", ".join(
                        _target_rule_label(r)
                        for r in overlap_error
                    )
                )

            else:
                new_rule = {
                    **draft_rule,
                    "target": float(
                        entered_target
                    ),
                    "locked": True,
                }

                target_rules = [
                    r
                    for r in target_rules
                    if _target_rule_id(r)
                    != draft_id
                ]

                target_rules.append(
                    new_rule
                )

                st.session_state[
                    "pv_target_rules"
                ] = target_rules

                st.success(
                    f"✓ {selected_employee} • "
                    f"{_target_rule_label(new_rule)} • "
                    f"{format_indian_currency_full(entered_target)} fixed."
                )

                st.rerun()

    if (
        change_clicked
        and existing_rule is not None
    ):
        updated_rules = []

        for rule in target_rules:
            if (
                _target_rule_id(rule)
                == draft_id
            ):
                rule = dict(rule)
                rule["locked"] = False

            updated_rules.append(
                rule
            )

        st.session_state[
            "pv_target_rules"
        ] = updated_rules

        st.rerun()

    if (
        remove_clicked
        and existing_rule is not None
    ):
        st.session_state[
            "pv_target_rules"
        ] = [
            r
            for r in target_rules
            if _target_rule_id(r)
            != draft_id
        ]

        st.session_state.pop(
            widget_key,
            None,
        )

        st.rerun()

    current_rules = [
        r
        for r in (
            _normalise_target_rule(x)
            for x in st.session_state.get(
                "pv_target_rules",
                [],
            )
        )
        if r is not None
    ]

    # --------------------------------------------------------
    # Current Target Rules
    # --------------------------------------------------------
    st.markdown(
        "##### Current Target Rules"
    )

    if current_rules:
        name_lookup = {
            canonical_sales_employee_key(
                name
            ): name
            for name in names
        }

        rows = []

        for rule in sorted(
            current_rules,
            key=lambda r: (
                name_lookup.get(
                    r["employee_key"],
                    r["employee_key"],
                ).lower(),
                0 if r["basis"] == "Year" else 1,
                _target_rule_label(r),
            ),
        ):
            rows.append(
                {
                    "Employee": name_lookup.get(
                        rule["employee_key"],
                        rule["employee_key"],
                    ),
                    "Type": (
                        "Year"
                        if rule["basis"] == "Year"
                        else "Date Range"
                    ),
                    "Period": _target_rule_label(
                        rule
                    ),
                    "Target": (
                        format_indian_currency_full(
                            rule["target"]
                        )
                    ),
                    "Status": (
                        "🔒 Fixed"
                        if rule.get(
                            "locked",
                            True,
                        )
                        else "✏️ Editable"
                    ),
                }
            )

        st.dataframe(
            pd.DataFrame(rows),
            width="stretch",
            hide_index=True,
            height=min(
                270,
                max(
                    105,
                    36 * len(rows)
                    + 40,
                ),
            ),
            column_config={
                "Employee": st.column_config.TextColumn(
                    "Employee",
                    width="large",
                ),
                "Type": st.column_config.TextColumn(
                    "Type",
                    width="small",
                ),
                "Period": st.column_config.TextColumn(
                    "Period",
                    width="large",
                ),
                "Target": st.column_config.TextColumn(
                    "Target",
                    width="medium",
                ),
                "Status": st.column_config.TextColumn(
                    "Status",
                    width="small",
                ),
            },
        )

    else:
        st.caption(
            "No fixed targets yet."
        )

    # --------------------------------------------------------
    # Target for currently selected MAIN dashboard date range.
    # --------------------------------------------------------
    result_rows = []

    for employee in names:
        emp_key = canonical_sales_employee_key(
            employee
        )

        target_value = (
            _employee_target_for_date_range(
                current_rules,
                emp_key,
                selected_start,
                selected_end,
            )
        )

        result_rows.append(
            {
                "Employee": employee,
                "Target": target_value,
            }
        )

    result = pd.DataFrame(
        result_rows
    )

    st.session_state[
        "pv_target_table"
    ] = result.copy()

    return result


def sales_representative_column(df):
    """
    Authoritative Sales employee field.

    Representative Ref is the business owner/name field used after the
    confirmed manual Sales Invoice cleaning process.

    Representative User is used only as a fallback when Ref is genuinely
    unavailable in another report version.
    """
    if df is not None and "Representative Ref" in df.columns:
        if df["Representative Ref"].notna().any():
            return "Representative Ref"

    if df is not None and "Representative User" in df.columns:
        return "Representative User"

    return None


def sales_invoice_target_performance(sales_df, target_df):
    """
    Employee target performance using ONLY Sales Invoice Register data.

    Achievement = SUM(Amount) from the current filtered Sales Invoice Register.
    Sales Return Register is intentionally NOT deducted in this analysis.
    """
    if sales_df is None or sales_df.empty:
        return pd.DataFrame(
            columns=[
                "Employee", "Target", "Achievement", "QTY",
                "Invoices", "Achievement %", "Gap", "Status"
            ]
        )

    rep_col = sales_representative_column(sales_df)

    if rep_col not in sales_df.columns:
        return pd.DataFrame(
            columns=[
                "Employee", "Target", "Achievement", "QTY",
                "Invoices", "Achievement %", "Gap", "Status"
            ]
        )

    work = sales_df.copy()

    # Target page may already provide this helper column.
    # Otherwise create it directly from the representative source.
    if "_Target Employee" not in work.columns:
        work["_Target Employee"] = (
            work[rep_col]
            .astype("string")
            .fillna("")
            .map(canonical_sales_employee)
        )

    work["_Target Employee"] = (
        work["_Target Employee"]
        .astype("string")
        .fillna("")
        .str.strip()
    )

    work = work[
        work["_Target Employee"].ne("")
        & ~work["_Target Employee"].str.lower().isin(
            ["nan", "none", "<na>", "-", "unassigned"]
        )
    ].copy()

    if work.empty:
        return pd.DataFrame(
            columns=[
                "Employee", "Target", "Achievement", "QTY",
                "Invoices", "Achievement %", "Gap", "Status"
            ]
        )

    if "Invoice No#" in work.columns:
        perf = (
            work.groupby("_Target Employee", as_index=False)
            .agg(
                Achievement=("Amount", "sum"),
                QTY=("QTY", "sum"),
                Invoices=("Invoice No#", "nunique"),
            )
            .rename(columns={"_Target Employee": "Employee"})
        )
    else:
        perf = (
            work.groupby("_Target Employee", as_index=False)
            .agg(
                Achievement=("Amount", "sum"),
                QTY=("QTY", "sum"),
                Invoices=("Amount", "size"),
            )
            .rename(columns={"_Target Employee": "Employee"})
        )

    if target_df is not None and not target_df.empty:
        targets = target_df[["Employee", "Target"]].copy()
        targets["Employee"] = (
            targets["Employee"]
            .astype(str)
            .map(canonical_sales_employee)
        )
        targets["Target"] = pd.to_numeric(
            targets["Target"], errors="coerce"
        ).fillna(0.0)

        targets = (
            targets[targets["Employee"].ne("")]
            .groupby("Employee", as_index=False)["Target"]
            .max()
        )

        # Full outer merge keeps:
        # - employees with Sales achievement but no target
        # - employees with Target but no achievement in the selected range
        perf = targets.merge(
            perf,
            on="Employee",
            how="outer",
        )
    else:
        perf["Target"] = 0.0

    # A target-only employee has no Sales rows in this selected range.
    for _col in ["Achievement", "QTY", "Invoices"]:
        if _col not in perf.columns:
            perf[_col] = 0
        perf[_col] = pd.to_numeric(
            perf[_col],
            errors="coerce",
        ).fillna(0)

    perf["Target"] = pd.to_numeric(
        perf["Target"], errors="coerce"
    ).fillna(0.0)

    perf["Achievement"] = pd.to_numeric(
        perf["Achievement"], errors="coerce"
    ).fillna(0.0)

    # FINAL percentage rule:
    # Achievement % = Total Achievement / Fixed Target * 100
    # Both values are raw rupee numbers.
    perf["Target"] = pd.to_numeric(
        perf["Target"],
        errors="coerce",
    ).fillna(0.0)

    perf["Achievement"] = pd.to_numeric(
        perf["Achievement"],
        errors="coerce",
    ).fillna(0.0)

    perf["Achievement %"] = np.where(
        perf["Target"] > 0,
        (perf["Achievement"] / perf["Target"]) * 100.0,
        np.nan,
    )

    # Positive Gap = still required to achieve target.
    perf["Gap"] = (perf["Target"] - perf["Achievement"]).clip(lower=0)

    perf["Status"] = np.select(
        [
            perf["Target"].le(0),
            perf["Achievement"].ge(perf["Target"]),
            perf["Achievement %"].ge(75),
        ],
        [
            "Target Not Set",
            "Target Achieved",
            "Watch",
        ],
        default="Below Target",
    )

    return perf.sort_values(
        ["Achievement %", "Achievement"],
        ascending=[False, False],
        na_position="last",
    ).reset_index(drop=True)


# ============================================================
# SALES FILTER HELPERS
# ============================================================

def apply_sales_filters(df):
    """Original working Sales filter behaviour, with a fresh V9 widget key."""
    if df.empty:
        return df
    f = df.copy()

    with st.sidebar:
        st.markdown('<div class="sidebar-section-num">📅 3. SALES DATE FILTER</div>', unsafe_allow_html=True)
        dates = f["Invoice Date"].dropna()
        if not dates.empty:
            mn, mx = dates.min().date(), dates.max().date()
            dv = st.date_input(
                "Sales Invoice Date Range",
                value=(mn, mx),
                min_value=mn,
                max_value=mx,
                key="sales_dates_v9",
                format="DD/MM/YYYY",
            )
            if isinstance(dv, (tuple, list)) and len(dv) == 2:
                start_date = pd.Timestamp(dv[0]).normalize()
                end_date = pd.Timestamp(dv[1]).normalize() + pd.Timedelta(days=1)
                f = f[(f["Invoice Date"] >= start_date) & (f["Invoice Date"] < end_date)]

        st.markdown('<div class="sidebar-section-num">🎛️ 4. SALES BUSINESS FILTERS</div>', unsafe_allow_html=True)
        for col, label, key in [
            ("Ledger Name", "Ledger Name (Customer)", "sales_customer_v9"),
            ("Item Group", "Item Group (Category)", "sales_group_v9"),
            ("Item Name", "Item Name (Product)", "sales_item_v9"),
        ]:
            vals = safe_unique(f, col)
            if vals:
                sel = st.multiselect(label, vals, key=key)
                if sel:
                    f = f[f[col].astype(str).isin([str(x) for x in sel])]

        rep_col = sales_representative_column(f)
        if rep_col in f.columns:
            vals = safe_unique(f, rep_col)
            if vals:
                sel = st.multiselect("Sales Representative", vals, key="sales_rep_v9")
                if sel:
                    f = f[f[rep_col].astype(str).isin([str(x) for x in sel])]

        st.markdown('<div class="sidebar-section-num">⚙️ 5. FILTER ACTIONS</div>', unsafe_allow_html=True)
        if st.button("↻ Reset Sales Filters", width="stretch", key="reset_sales_v9"):
            for k in [
                "sales_dates_v9", "sales_rep_v9", "sales_group_v9",
                "sales_item_v9", "sales_customer_v9", "return_dates_v9",
                "return_date_scope_v9"
            ]:
                st.session_state.pop(k, None)
            st.rerun()

    return f


# ============================================================
# SALES RETURN REGISTER MODULE
# ============================================================
RETURN_REQUIRED = [
    "Ledger Name", "Order No#", "Sales Return Date",
    "Item Group", "Item Name", "QTY", "Amount"
]
RETURN_CONTEXT_COLUMNS = [
    "Branch", "Ledger Name", "Order No#", "Sales Return Date",
    "Representative Ref", "Representative User"
]


def find_return_header_row(raw):
    wanted = {
        "Ledger Name", "Order No#", "Sales Return Date", "Item Group",
        "Item Name", "QTY", "Amount", "Rate", "Branch"
    }
    best_row, best_score = 0, -1
    for row in range(min(150, len(raw))):
        values = {clean_text(v) for v in raw.iloc[row].tolist()}
        score = len(values.intersection(wanted))
        if score > best_score:
            best_row, best_score = row, score
    return best_row, best_score


def load_return_excel(uploaded_file):
    name = getattr(uploaded_file, "name", "return.xlsx").lower()
    uploaded_file.seek(0)
    if name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file, header=None, dtype=object)
    else:
        raw = pd.read_excel(uploaded_file, header=None, dtype=object)
    if raw is None or raw.empty:
        raise ValueError("The Sales Return Register is empty.")

    header_row, score = find_return_header_row(raw)
    if score < 4:
        raise ValueError(
            "Could not identify Sales Return header. Expected Ledger Name, "
            "Order No#, Sales Return Date, Item Group, Item Name, QTY and Amount."
        )

    headers = [clean_text(v) for v in raw.iloc[header_row].tolist()]
    used, safe = {}, []
    for i, header in enumerate(headers):
        header = header or f"Unnamed_{i+1}"
        if header not in used:
            used[header] = 0
            safe.append(header)
        else:
            used[header] += 1
            safe.append(f"{header}_{used[header]}")

    df = raw.iloc[header_row + 1:].copy()
    df.columns = safe
    df = df.dropna(how="all").dropna(axis=1, how="all").copy()

    # Flexible return aliases: tolerate export spelling/case/spacing differences.
    canonical = {
        "ledger name": "Ledger Name", "customer": "Ledger Name", "customer name": "Ledger Name",
        "order no#": "Order No#", "order no": "Order No#", "invoice no#": "Order No#", "invoice no": "Order No#",
        "sales return date": "Sales Return Date", "return date": "Sales Return Date", "invoice date": "Sales Return Date",
        "item group": "Item Group", "group": "Item Group", "product group": "Item Group",
        "item name": "Item Name", "item": "Item Name", "product": "Item Name", "product name": "Item Name",
        "qty": "QTY", "quantity": "QTY", "amount": "Amount", "rate": "Rate",
        "representative ref": "Representative Ref", "representative user": "Representative User",
    }
    rename = {}
    for c in df.columns:
        key = clean_text(c).casefold().replace("  ", " ")
        if key in canonical and canonical[key] not in df.columns:
            rename[c] = canonical[key]
    if rename:
        df = df.rename(columns=rename)
    return df, header_row, score


def process_return_data(uploaded_file):
    """Clean Sales Return Register with the same commercial rules as Sales."""
    df, header_row, header_score = load_return_excel(uploaded_file)
    missing = [c for c in RETURN_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("Return Register columns missing: " + ", ".join(missing))

    original_rows = len(df)
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].map(clean_text)
    df = df.replace("", np.nan)

    for col in RETURN_CONTEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].replace(r"^\s*$", np.nan, regex=True).ffill()

    for col in ["Rate", "QTY", "Amount", "CGST", "SGST", "IGST", "CESS", "Net Amount"]:
        if col in df.columns:
            df[col] = numeric_clean(df[col])

    df["Sales Return Date"] = parse_mixed_date_series(df["Sales Return Date"])
    invalid_date_mask = df["Sales Return Date"].isna()
    invalid_dates_removed = int(invalid_date_mask.sum())
    df = df.loc[~invalid_date_mask].copy()

    # E-Khata return exports contain one order-summary line (Item Name blank)
    # followed by the item detail lines. The summary line is used ONLY as a
    # reconciliation check; it must not be counted again in analytics.
    summary_mask = (
        df["Item Name"].isna()
        & df["QTY"].notna()
        & df["Amount"].notna()
        & df["Order No#"].notna()
    )
    report_total_qty = float(df.loc[summary_mask, "QTY"].abs().sum()) if summary_mask.any() else 0.0
    report_total_amount = float(df.loc[summary_mask, "Amount"].abs().sum()) if summary_mask.any() else 0.0

    # Return Register does NOT require Rate.
    rate_removed = 0

    invalid_core_mask = (
        df["QTY"].isna()
        | df["Amount"].isna()
        | df["Item Name"].isna()
        | df["Ledger Name"].isna()
        | df["Order No#"].isna()
    )
    core_removed = int(invalid_core_mask.sum())
    df = df.loc[~invalid_core_mask].copy()

    # Preserve raw values for audit. Return exports can store signs differently;
    # analytics uses positive return magnitude and subtracts exactly once later.
    df["Return QTY Raw"] = df["QTY"]
    df["Return Amount Raw"] = df["Amount"]
    df["QTY"] = df["QTY"].abs()
    df["Amount"] = df["Amount"].abs()

    # V11 duplicate protection: do not blindly drop repeated-looking return lines.
    # Only auto-reconcile exact business-line duplicates when the report contains a
    # summary total AND removing those duplicates makes BOTH QTY and Amount match it.
    dup_cols = [c for c in [
        "Sales Return Date", "Order No#", "Ledger Name", "Item Group", "Item Name",
        "Hsn Code", "Rate", "QTY", "Amount"
    ] if c in df.columns]
    duplicate_flagged = int(df.duplicated(subset=dup_cols, keep=False).sum()) if dup_cols else 0
    duplicate_removed = 0
    detail_qty_before = float(df["QTY"].sum())
    detail_amount_before = float(df["Amount"].sum())
    if report_total_qty and report_total_amount and dup_cols and duplicate_flagged:
        candidate = df.drop_duplicates(subset=dup_cols, keep="first").copy()
        cq = float(candidate["QTY"].sum())
        ca = float(candidate["Amount"].sum())
        if abs(cq - report_total_qty) < 0.001 and abs(ca - report_total_amount) < 0.01:
            duplicate_removed = int(len(df) - len(candidate))
            df = candidate


    df["Return Year"] = df["Sales Return Date"].dt.year.astype("Int64")
    df["Return Month"] = df["Sales Return Date"].dt.month_name()
    df["Return Year-Month"] = df["Sales Return Date"].dt.to_period("M").astype(str)
    df["Product KW"] = df["Item Name"].map(normalize_product_kw)

    preferred = [
        "Branch", "Ledger Name", "Order No#", "Sales Return Date",
        "Representative Ref", "Representative User", "Item Group", "Item Name",
        "Hsn Code", "Rate", "QTY", "Amount", "CGST", "SGST", "IGST", "CESS",
        "Net Amount", "Return Year", "Return Month", "Return Year-Month"
    ]
    ordered = [c for c in preferred if c in df.columns]
    ordered += [c for c in df.columns if c not in ordered]
    df = df[ordered].reset_index(drop=True)

    summary = {
        "original": int(original_rows),
        "rate_removed": rate_removed,
        "invalid_dates_removed": invalid_dates_removed,
        "core_removed": core_removed,
        "duplicates": int(duplicate_removed),
        "duplicate_rows_flagged": duplicate_flagged,
        "duplicates_reconciled": int(duplicate_removed),
        "detail_qty_before_reconcile": detail_qty_before,
        "detail_amount_before_reconcile": detail_amount_before,
        "final": int(len(df)),
        "header_row": int(header_row + 1),
        "header_score": int(header_score),
        "return_qty": float(df["QTY"].sum()) if not df.empty else 0.0,
        "return_amount": float(df["Amount"].sum()) if not df.empty else 0.0,
        "report_total_qty": report_total_qty,
        "report_total_amount": report_total_amount,
        "qty_check_diff": (float(df["QTY"].sum()) - report_total_qty) if report_total_qty else 0.0,
        "amount_check_diff": (float(df["Amount"].sum()) - report_total_amount) if report_total_amount else 0.0,
    }
    summary["validation_status"] = (
        "PASS"
        if (not report_total_qty or abs(summary["qty_check_diff"]) < 0.001)
        and (not report_total_amount or abs(summary["amount_check_diff"]) < 0.01)
        else "CHECK"
    )
    return df, summary


def load_register_files(files, processor):
    """Clean each register independently before combining its commercial lines."""
    frames, audits, errors = [], [], []
    for uploaded_file in files or []:
        name = getattr(uploaded_file, "name", "Register.xlsx")
        try:
            frame, audit = processor(uploaded_file)
        except Exception as exc:
            errors.append(f"{name}: {exc}")
            continue
        frame = frame.copy()
        frame["Source File"] = name
        frames.append(frame)
        audits.append({"File": name, **audit})
    if not frames:
        return pd.DataFrame(), {}, errors
    combined = pd.concat(frames, ignore_index=True, sort=False)
    summary = {}
    metadata = {"File", "header_row", "header_score", "validation_status"}
    for key in set().union(*(audit.keys() for audit in audits)) - metadata:
        summary[key] = sum(audit.get(key, 0) for audit in audits)
    for key in ["header_row", "header_score"]:
        values = list(dict.fromkeys(audit[key] for audit in audits if key in audit))
        summary[key] = values[0] if len(values) == 1 else "See file audit"
    if any("validation_status" in audit for audit in audits):
        # Opposite reconciliation differences in separate files must not cancel a failure.
        summary["validation_status"] = "PASS" if all(audit.get("validation_status") == "PASS" for audit in audits) else "CHECK"
    summary["files"] = audits
    return combined, summary, errors


def reset_register_upload_filters(kind):
    keys = {
        "sales": ["sales_dates_v9", "sales_rep_v9", "sales_group_v9", "sales_item_v9",
                  "sales_customer_v9", "sales_target_date_filter", "sales_target_emp_filter"],
        "returns": ["return_dates_v9", "return_date_scope_v9"],
        "service": ["service_dates_v9", "service_date_basis_v9", "service_date_filter_enabled"],
    }
    for key in keys[kind]:
        st.session_state.pop(key, None)


def apply_return_filters(return_df):
    """
    Filter Return Register independently by return date, while reusing the active
    Sales customer/group/product selections. Default = all uploaded return dates,
    so uploading a Return Register immediately deducts it from Gross Sales.
    """
    if return_df.empty:
        return return_df.copy()
    f = return_df.copy()

    with st.sidebar:
        st.markdown('<div class="sidebar-section-num">↩️ 6. RETURN REGISTER FILTER</div>', unsafe_allow_html=True)
        scope = st.radio(
            "Return Date Scope",
            ["All / Return Date Range", "Match Sales Date Range"],
            key="return_date_scope_v9",
            help="Use All / Return Date Range to deduct the uploaded Return Register independently. Use Match Sales Date Range for period-by-period reconciliation.",
        )

        if scope == "Match Sales Date Range":
            dv = st.session_state.get("sales_dates_v9")
            if isinstance(dv, (tuple, list)) and len(dv) == 2:
                start_date = pd.Timestamp(dv[0]).normalize()
                end_date = pd.Timestamp(dv[1]).normalize() + pd.Timedelta(days=1)
                f = f[(f["Sales Return Date"] >= start_date) & (f["Sales Return Date"] < end_date)]
        else:
            valid_dates = f["Sales Return Date"].dropna()
            if not valid_dates.empty:
                mn, mx = valid_dates.min().date(), valid_dates.max().date()
                rv = st.date_input(
                    "Sales Return Date Range",
                    value=(mn, mx),
                    min_value=mn,
                    max_value=mx,
                    key="return_dates_v9",
                    format="DD/MM/YYYY",
                )
                if isinstance(rv, (tuple, list)) and len(rv) == 2:
                    start_date = pd.Timestamp(rv[0]).normalize()
                    end_date = pd.Timestamp(rv[1]).normalize() + pd.Timedelta(days=1)
                    f = f[(f["Sales Return Date"] >= start_date) & (f["Sales Return Date"] < end_date)]

    # Reuse active Sales business selections so customer/product views reconcile.
    for state_key, col in [
        ("sales_customer_v9", "Ledger Name"),
        ("sales_group_v9", "Item Group"),
        ("sales_item_v9", "Item Name"),
    ]:
        selected = st.session_state.get(state_key, [])
        if selected and col in f.columns:
            selected_norm = {clean_text(x).casefold() for x in selected}
            f = f[f[col].map(lambda x: clean_text(x).casefold()).isin(selected_norm)]

    return f.copy()


def _match_key(value):
    if pd.isna(value):
        return ""
    text = clean_text(value).casefold()
    # Excel often turns 12345 into 12345.0; normalize that harmless difference.
    text = re.sub(r"\.0$", "", text)
    return re.sub(r"\s+", "", text)


def assign_return_representative(return_df, sales_df):
    """Map returns to a sales person without pandas dtype-upcast failures.

    Priority:
    1. Representative already present in the Return Register.
    2. Order No# -> Invoice No# exact normalized match.
    3. Ledger + Item fallback only where the Sales source identifies exactly one person.
    """
    if return_df.empty:
        out = return_df.copy()
        out["Return Representative"] = pd.Series(index=out.index, dtype="object")
        return out

    out = return_df.copy()
    # IMPORTANT: object dtype from the beginning. Initializing with np.nan alone creates
    # float64 in recent pandas and later string assignment raises TypeError.
    out["Return Representative"] = pd.Series(pd.NA, index=out.index, dtype="object")

    sales_rep_col = next(
        (c for c in ["Representative Ref", "Representative User"]
         if c in sales_df.columns and sales_df[c].notna().any()),
        None,
    )

    direct_col = next(
        (c for c in ["Representative Ref", "Representative User"]
         if c in out.columns and out[c].notna().any()),
        None,
    )
    if direct_col:
        direct = out[direct_col].map(clean_text).astype("object")
        mask = direct.ne("")
        out.loc[mask, "Return Representative"] = direct.loc[mask].to_numpy(dtype=object)

    if sales_rep_col and not sales_df.empty:
        sales = sales_df.copy()
        sales[sales_rep_col] = sales[sales_rep_col].map(clean_text).astype("object")

        def unique_person(series):
            vals = series.map(clean_text)
            vals = vals[vals.ne("")]
            uniq = pd.unique(vals)
            return uniq[0] if len(uniq) == 1 else pd.NA

        # Highest-confidence order/invoice mapping.
        if "Order No#" in out.columns and "Invoice No#" in sales.columns:
            sales["__order_key"] = sales["Invoice No#"].map(_match_key)
            order_map = (
                sales[sales["__order_key"].ne("")]
                .groupby("__order_key", dropna=False)[sales_rep_col]
                .agg(unique_person)
            )
            blank = out["Return Representative"].isna()
            mapped = out.loc[blank, "Order No#"].map(
                lambda x: order_map.get(_match_key(x), pd.NA)
            ).astype("object")
            out.loc[blank, "Return Representative"] = mapped.to_numpy(dtype=object)

        # Conservative customer+product fallback.
        if all(c in sales.columns for c in ["Ledger Name", "Item Name"]):
            base = sales[["Ledger Name", "Item Name", sales_rep_col]].copy()
            base["Ledger Name"] = base["Ledger Name"].map(clean_text)
            base["Item Name"] = base["Item Name"].map(clean_text)
            pair_map = base.groupby(["Ledger Name", "Item Name"], dropna=False)[sales_rep_col].agg(unique_person)

            blank = out["Return Representative"].isna()
            if blank.any():
                keys = list(zip(
                    out.loc[blank, "Ledger Name"].map(clean_text),
                    out.loc[blank, "Item Name"].map(clean_text),
                ))
                mapped = pd.Series(
                    [pair_map.get(k, pd.NA) for k in keys],
                    index=out.index[blank],
                    dtype="object",
                )
                out.loc[blank, "Return Representative"] = mapped.to_numpy(dtype=object)

    out["Return Representative"] = (
        out["Return Representative"]
        .astype("object")
        .where(out["Return Representative"].notna(), "Unallocated Return")
    )
    return out


def net_metrics(sales_frame, return_frame):
    gross_amount = float(sales_frame["Amount"].sum()) if not sales_frame.empty else 0.0
    gross_qty = float(sales_frame["QTY"].sum()) if not sales_frame.empty else 0.0
    return_amount = float(return_frame["Amount"].sum()) if not return_frame.empty else 0.0
    return_qty = float(return_frame["QTY"].sum()) if not return_frame.empty else 0.0
    return {
        "gross_amount": gross_amount,
        "gross_qty": gross_qty,
        "return_amount": return_amount,
        "return_qty": return_qty,
        "net_amount": gross_amount - return_amount,
        "net_qty": gross_qty - return_qty,
    }


def net_by_dimension(sales_frame, return_frame, dimension):
    if sales_frame.empty or dimension not in sales_frame.columns:
        return pd.DataFrame(columns=[dimension, "Gross Sales", "Return Amount", "Net Sales", "Gross QTY", "Return QTY", "Net QTY"])
    gross = sales_frame.groupby(dimension, dropna=False, as_index=False).agg(
        **{"Gross Sales": ("Amount", "sum"), "Gross QTY": ("QTY", "sum")}
    )
    if return_frame.empty or dimension not in return_frame.columns:
        gross["Return Amount"] = 0.0
        gross["Return QTY"] = 0.0
        result = gross
    else:
        ret = return_frame.groupby(dimension, dropna=False, as_index=False).agg(
            **{"Return Amount": ("Amount", "sum"), "Return QTY": ("QTY", "sum")}
        )
        result = gross.merge(ret, on=dimension, how="outer")
        for c in ["Gross Sales", "Gross QTY", "Return Amount", "Return QTY"]:
            result[c] = pd.to_numeric(result[c], errors="coerce").fillna(0)
    result["Net Sales"] = result["Gross Sales"] - result["Return Amount"]
    result["Net QTY"] = result["Gross QTY"] - result["Return QTY"]
    return result


def employee_net_performance(sales_frame, return_frame):
    rep_col = sales_representative_column(sales_frame)
    if sales_frame.empty or rep_col not in sales_frame.columns:
        return pd.DataFrame(), rep_col
    gross = sales_frame.groupby(rep_col, as_index=False).agg(
        Gross_Achievement=("Amount", "sum"),
        Gross_QTY=("QTY", "sum"),
        Invoices=("Invoice No#", "nunique") if "Invoice No#" in sales_frame.columns else ("Amount", "size")
    ).rename(columns={rep_col: "Employee"})
    if return_frame.empty:
        gross["Return Amount"] = 0.0
        gross["Return QTY"] = 0.0
    else:
        r = return_frame.groupby("Return Representative", as_index=False).agg(
            **{"Return Amount": ("Amount", "sum"), "Return QTY": ("QTY", "sum")}
        ).rename(columns={"Return Representative": "Employee"})
        gross = gross.merge(r, on="Employee", how="left")
        gross[["Return Amount", "Return QTY"]] = gross[["Return Amount", "Return QTY"]].fillna(0)
    gross["Achievement"] = gross["Gross_Achievement"] - gross["Return Amount"]
    gross["QTY"] = gross["Gross_QTY"] - gross["Return QTY"]
    return gross, rep_col


# ============================================================
# FIELD DASHBOARD / KPI-KRA MODULE
# ============================================================
FIELD_METRICS = [
    "Visit", "KM", "Lead", "Follow-Ups", "Quotation",
    "PI", "Sales-Order", "Wrong KM Claim"
]
FIELD_ALIASES = {
    "visit": "Visit", "visits": "Visit", "km": "KM", "kms": "KM",
    "lead": "Lead", "leads": "Lead", "follow-up": "Follow-Ups",
    "follow up": "Follow-Ups", "followup": "Follow-Ups", "follow-ups": "Follow-Ups",
    "quotation": "Quotation", "quotations": "Quotation", "pi": "PI",
    "sales order": "Sales-Order", "sales-order": "Sales-Order", "salesorder": "Sales-Order",
    "wrong claims km": "Wrong KM Claim", "wrong claim km": "Wrong KM Claim",
    "wrong km claim": "Wrong KM Claim", "wrong km claims": "Wrong KM Claim"
}
FIELD_DEFAULT_TARGETS = {
    "Visit": 75, "Lead": 20, "Follow-Ups": 0,
    "Quotation": 15, "PI": 10, "Sales-Order": 8, "KM": 500
}


def field_metric_name(value):
    raw = clean_text(value)
    return FIELD_ALIASES.get(raw.lower(), raw)


def field_number(value):
    if value is None or pd.isna(value):
        return 0.0
    if isinstance(value, (int, float, np.integer, np.floating)):
        return 0.0 if pd.isna(value) else float(value)
    raw = clean_text(value)
    if raw.lower() in {"", "check", "-", "--", "n/a", "na", "nan", "none", "null"}:
        return 0.0
    raw = raw.replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", raw)
    return float(match.group()) if match else 0.0


def field_date(value):
    """Parse Excel serials, real datetime cells and mixed text dates safely.

    Important: matrix headers with the special PV-Blink YYYY/DD/MM pattern are
    normalized by parse_dashboard_header_dates() before this generic parser runs.
    """
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value).normalize()
    if isinstance(value, np.datetime64):
        return pd.Timestamp(value).normalize()
    if isinstance(value, (int, float, np.integer, np.floating)):
        v = float(value)
        if 1 <= v <= 60000:
            parsed = pd.to_datetime(v, unit="D", origin="1899-12-30", errors="coerce")
            return parsed.normalize() if pd.notna(parsed) else None
    raw = clean_text(value)
    if not raw:
        return None
    # ISO dates in flat reports are year-month-day, even when both numbers <= 12.
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
        parsed = pd.to_datetime(raw, format="%Y-%m-%d", errors="coerce")
        return parsed.normalize() if pd.notna(parsed) else None
    # Prefer pandas mixed-format parsing. dayfirst=True is safest for Indian reports.
    try:
        parsed = pd.to_datetime(raw, errors="coerce", dayfirst=True, format="mixed")
    except (TypeError, ValueError):
        parsed = pd.to_datetime(raw, errors="coerce", dayfirst=True)
    return pd.Timestamp(parsed).normalize() if pd.notna(parsed) else None


def parse_dashboard_header_dates(headers, excluded):
    """Parse Field KPI/KRA matrix date headers without month/day swapping.

    PV-Blink source files can contain headers in either YYYY/DD/MM or YYYY/MM/DD.
    Excel may auto-convert only the first 12 ambiguous headers to real datetimes.
    This routine tests BOTH interpretations and chooses the interpretation that
    forms the strongest single-month, day-by-day reporting sequence.
    """
    candidates = [c for c in headers if c not in excluded]

    def parts(value):
        # Preserve real Excel/Pandas date cells as their visible Y-M-D components.
        if isinstance(value, (pd.Timestamp, datetime, np.datetime64)):
            t = pd.Timestamp(value)
            return int(t.year), int(t.month), int(t.day)

        raw = clean_text(value)
        m = re.fullmatch(
            r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T]\d{1,2}:\d{2}:\d{2}(?:\.\d+)?)?",
            raw,
        )
        if not m:
            return None
        return int(m.group(1)), int(m.group(2)), int(m.group(3))

    parsed = [(c, parts(c)) for c in candidates]
    parsed = [(c, p) for c, p in parsed if p is not None]

    def build(mode):
        rows = []
        for c, (y, a, b) in parsed:
            try:
                if mode == "YDM":
                    dt = pd.Timestamp(year=y, month=b, day=a).normalize()
                else:  # YMD
                    dt = pd.Timestamp(year=y, month=a, day=b).normalize()
                rows.append((c, dt))
            except (ValueError, TypeError):
                pass
        return rows

    def sequence_score(rows):
        if not rows:
            return (-1, -1, -1, -1)
        dates = [d for _, d in rows]
        ym = [(d.year, d.month) for d in dates]
        mode_ym = pd.Series(ym).mode()
        common_ym = mode_ym.iloc[0] if not mode_ym.empty else None
        same_month = [d for d in dates if (d.year, d.month) == common_ym]
        days = sorted(set(d.day for d in same_month))
        consecutive = sum(1 for a, b in zip(days, days[1:]) if b - a == 1)
        # Strong preference: many unique dates in one month, then continuity.
        return (len(same_month), len(days), consecutive, len(rows))

    ydm = build("YDM")
    ymd = build("YMD")
    score_ydm = sequence_score(ydm)
    score_ymd = sequence_score(ymd)

    # Require at least 3 structured date headers before overriding generic parsing.
    structured = ydm if len(parsed) >= 3 and score_ydm > score_ymd else ymd if len(parsed) >= 3 else []

    result = []
    seen = set()
    structured_map = {c: d for c, d in structured}

    for c in candidates:
        d = structured_map.get(c)
        if d is None:
            d = field_date(c)
        if d is not None:
            d = pd.Timestamp(d).normalize()
            if d not in seen:
                result.append((c, d))
                seen.add(d)

    return result

def unique_headers(values):
    seen, result = {}, []
    for i, value in enumerate(values):
        name = clean_text(value) or f"Column_{i+1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        result.append(name)
    return result


def process_field_dashboard(uploaded_file):
    """Load Dashboard KPI/KRA data accurately and preserve the full employee roster.

    Supported sources:
    A) Matrix: Sr.No | Emp_Name | [Designation] | Details | date columns | Total
    B) Flat:   Emp_Name | [Designation] | [Date] | Visit | KM | Lead | ...

    Accuracy rules:
    - Employee names are normalized case-insensitively but their first source spelling is kept.
    - All roster employees are retained, even when an employee has missing KPI rows / zero activity.
    - Date-column values are read from cached Excel values when available.
    - Duplicate employee+date+metric cells are SUMMED, never silently discarded.
    - Missing KPI rows remain zero and are exposed through coverage diagnostics.
    """
    uploaded_file.seek(0)
    file_bytes = uploaded_file.read()
    excel_buf = io.BytesIO(file_bytes)
    with pd.ExcelFile(excel_buf) as xls:
        excel_engine = xls.engine
        dashboard_sheet = next((s for s in xls.sheet_names if clean_text(s).lower() == "dashboard"), None)
        if dashboard_sheet is None:
            dashboard_sheet = next((s for s in xls.sheet_names if "dashboard" in clean_text(s).lower()), None)
        if dashboard_sheet is None:
            raise ValueError("Dashboard sheet was not found in the KPI/KRA Excel workbook.")
        raw_formula = xls.parse(sheet_name=dashboard_sheet, header=None, dtype=object)

    # Cached values are important when date cells / totals are formula-driven.
    if excel_engine == "openpyxl":
        wb_values = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        try:
            ws_values = wb_values[dashboard_sheet]
            raw_values = pd.DataFrame(ws_values.values)
        finally:
            wb_values.close()
    else:
        # xlrd already supplies cached values for legacy XLS workbooks.
        raw_values = raw_formula.copy()

    # Keep the same physical row alignment between formula and cached-value views.
    nonempty_mask = ~raw_formula.isna().all(axis=1)
    raw_formula = raw_formula.loc[nonempty_mask].reset_index(drop=True)
    raw_values = raw_values.reindex(index=nonempty_mask.index)
    raw_values = raw_values.loc[nonempty_mask].reset_index(drop=True)

    if raw_formula.empty:
        raise ValueError("Dashboard sheet is empty.")

    matrix_header = None
    flat_header = None
    for r in range(min(60, len(raw_formula))):
        vals = [clean_text(v) for v in raw_formula.iloc[r].tolist()]
        lower = [v.lower() for v in vals]
        if "emp_name" in lower and "details" in lower:
            matrix_header = r
            break
        if "emp_name" in lower and any(field_metric_name(v) in FIELD_METRICS for v in vals):
            flat_header = r
            break

    def normalize_employee(value, name_map):
        name = re.sub(r"\s+", " ", clean_text(value)).strip()
        if not name:
            return ""
        key = name.casefold()
        return name_map.setdefault(key, name)

    if matrix_header is not None:
        headers = unique_headers(raw_formula.iloc[matrix_header].tolist())
        width = len(headers)
        body_formula = raw_formula.iloc[matrix_header + 1:, :width].copy().reset_index(drop=True)
        body_values = raw_values.iloc[matrix_header + 1:, :width].copy().reset_index(drop=True)
        body_formula.columns = headers
        body_values.columns = headers

        emp_col = next((c for c in headers if clean_text(c).lower() == "emp_name"), None)
        details_col = next((c for c in headers if clean_text(c).lower() == "details"), None)
        sr_col = next((c for c in headers if clean_text(c).lower() in {"sr.no", "sr no", "sr.no."}), None)
        designation_col = next((c for c in headers if clean_text(c).lower() == "designation"), None)
        total_col = next((c for c in headers if clean_text(c).lower() == "total"), None)
        if emp_col is None or details_col is None:
            raise ValueError("Dashboard matrix requires Emp_Name and Details columns.")

        # Forward-fill identifiers only, matching the source block structure.
        for c in [emp_col, sr_col, designation_col]:
            if c:
                body_formula[c] = body_formula[c].ffill()
                body_values[c] = body_values[c].ffill()

        excluded = {emp_col, details_col}
        for c in [sr_col, total_col, designation_col]:
            if c:
                excluded.add(c)

        date_cols = parse_dashboard_header_dates(headers, excluded)

        name_map = {}
        roster = []
        employee_designation = {}
        metric_rows_by_employee = {}
        records = []
        unknown_metric_rows = 0

        # Build roster independently from recognized KPI rows. This is the key fix for 24 vs 27.
        for _, row in body_formula.iterrows():
            employee = normalize_employee(row.get(emp_col), name_map)
            if not employee:
                continue
            # Avoid accidental footer propagation: a roster employee must belong to a source block
            # with either an Sr.No or a nonblank Details cell.
            sr_ok = True
            if sr_col:
                sr_ok = clean_text(row.get(sr_col)) != ""
            details_raw = clean_text(row.get(details_col))
            if sr_ok and employee not in roster:
                roster.append(employee)
            if designation_col:
                des = clean_text(row.get(designation_col))
                if des:
                    employee_designation[employee] = des

            metric = field_metric_name(details_raw)
            if metric not in FIELD_METRICS:
                if details_raw:
                    unknown_metric_rows += 1
                continue
            metric_rows_by_employee.setdefault(employee, set()).add(metric)

        for i, row in body_formula.iterrows():
            employee = normalize_employee(row.get(emp_col), name_map)
            metric = field_metric_name(row.get(details_col))
            if not employee or metric not in FIELD_METRICS:
                continue

            if date_cols:
                for col, dt in date_cols:
                    # Prefer cached displayed values; fall back to formula view only if needed.
                    value = body_values.iloc[i].get(col) if i < len(body_values) else row.get(col)
                    if pd.isna(value) and not pd.isna(row.get(col)):
                        value = row.get(col)
                    records.append({
                        "Employee": employee,
                        "Date": dt,
                        "Metric": metric,
                        "Value": field_number(value),
                    })
            elif total_col:
                value = body_values.iloc[i].get(total_col) if i < len(body_values) else row.get(total_col)
                if pd.isna(value) and not pd.isna(row.get(total_col)):
                    value = row.get(total_col)
                records.append({
                    "Employee": employee,
                    "Date": pd.NaT,
                    "Metric": metric,
                    "Value": field_number(value),
                })

        if not records and not roster:
            raise ValueError("No employee roster or valid KPI/KRA activity records were found in Dashboard sheet.")

        if records:
            long_df = pd.DataFrame(records)
            long_df["Value"] = pd.to_numeric(long_df["Value"], errors="coerce").fillna(0.0)
            if long_df["Date"].notna().any():
                wide = long_df.pivot_table(
                    index=["Employee", "Date"], columns="Metric", values="Value", aggfunc="sum", fill_value=0
                ).reset_index()
            else:
                wide = long_df.pivot_table(
                    index=["Employee"], columns="Metric", values="Value", aggfunc="sum", fill_value=0
                ).reset_index()
                wide["Date"] = pd.NaT
            wide.columns.name = None
        else:
            wide = pd.DataFrame(columns=["Employee", "Date"])

        # Preserve every employee even if all KPI rows are missing / zero.
        present = set(wide["Employee"].astype(str)) if not wide.empty else set()
        missing_roster = [e for e in roster if e not in present]
        if missing_roster:
            zero_rows = pd.DataFrame({"Employee": missing_roster, "Date": [pd.NaT] * len(missing_roster)})
            wide = pd.concat([wide, zero_rows], ignore_index=True, sort=False)

        for m in FIELD_METRICS:
            if m not in wide.columns:
                wide[m] = 0.0
            wide[m] = pd.to_numeric(wide[m], errors="coerce").fillna(0.0)

        if employee_designation:
            wide["Designation"] = wide["Employee"].map(employee_designation)

        coverage_rows = []
        for employee in roster:
            reported = metric_rows_by_employee.get(employee, set())
            coverage_rows.append({
                "Employee": employee,
                "Reported Metrics": len(reported),
                "Missing Metric Rows": len(FIELD_METRICS) - len(reported),
                "Coverage %": round(len(reported) / len(FIELD_METRICS) * 100, 1),
            })
        coverage_df = pd.DataFrame(coverage_rows)
        source_format = "Dashboard matrix (Emp_Name + Details + Date columns)"
        header_row = matrix_header + 1

    elif flat_header is not None:
        headers = unique_headers(raw_formula.iloc[flat_header].tolist())
        body = raw_values.iloc[flat_header + 1:, :len(headers)].copy()
        body.columns = headers
        rename = {}
        for c in body.columns:
            low = clean_text(c).lower()
            if low == "emp_name":
                rename[c] = "Employee"
            elif low == "designation":
                rename[c] = "Designation"
            elif low == "date":
                rename[c] = "Date"
            else:
                metric = field_metric_name(c)
                if metric in FIELD_METRICS:
                    rename[c] = metric
        body = body.rename(columns=rename)
        if "Employee" not in body.columns:
            raise ValueError("Flat Dashboard format requires Emp_Name.")

        name_map = {}
        body["Employee"] = body["Employee"].map(lambda v: normalize_employee(v, name_map))
        body = body[body["Employee"].ne("")].copy()
        roster = list(dict.fromkeys(body["Employee"].tolist()))

        for m in FIELD_METRICS:
            if m not in body.columns:
                body[m] = 0.0
            body[m] = body[m].map(field_number)
        if "Date" in body.columns:
            body["Date"] = pd.to_datetime(body["Date"].map(field_date), errors="coerce")
        else:
            body["Date"] = pd.NaT
        keep = ["Employee", "Date"] + (["Designation"] if "Designation" in body.columns else []) + FIELD_METRICS
        wide = body[keep].copy()
        coverage_df = pd.DataFrame({
            "Employee": roster,
            "Reported Metrics": [len([m for m in FIELD_METRICS if m in body.columns])] * len(roster),
            "Missing Metric Rows": [0] * len(roster),
            "Coverage %": [100.0] * len(roster),
        })
        unknown_metric_rows = 0
        source_format = "Flat KPI table (Emp_Name + KPI columns)"
        header_row = flat_header + 1
    else:
        raise ValueError(
            "Dashboard format not recognized. Expected Sr.No | Emp_Name | Details | Date columns | Total "
            "or a flat Emp_Name KPI table."
        )

    for m in FIELD_METRICS:
        wide[m] = pd.to_numeric(wide[m], errors="coerce").fillna(0.0)
    wide["Valid KM"] = (wide["KM"] - wide["Wrong KM Claim"]).clip(lower=0)

    # Stable identity diagnostics for UI/audit.
    roster = list(dict.fromkeys([clean_text(e) for e in roster if clean_text(e)]))
    info = {
        "sheet": dashboard_sheet,
        "format": source_format,
        "header_row": header_row,
        "rows": int(len(wide)),
        "employees": int(len(roster)),
        "roster_employees": int(len(roster)),
        "active_employees": int(wide.loc[wide[FIELD_METRICS].sum(axis=1) > 0, "Employee"].nunique()),
        "dates": int(wide["Date"].nunique()) if wide["Date"].notna().any() else 0,
        "unknown_metric_rows": int(unknown_metric_rows),
        "employee_roster": roster,
        "coverage": coverage_df,
    }
    return wide.reset_index(drop=True), info


def reset_field_upload_filters():
    for key in ["field_emp_filter", "field_date_filter_v26"]:
        st.session_state.pop(key, None)


def load_field_files(files):
    """Combine monthly field workbooks without discarding legitimate activity."""
    frames, coverages, reports, errors = [], [], [], []
    employee_names = {}
    for uploaded_file in files or []:
        name = getattr(uploaded_file, "name", "Field_Report.xlsx")
        try:
            frame, info = process_field_dashboard(uploaded_file)
        except Exception as exc:
            errors.append(f"{name}: {exc}")
            continue
        for employee in info["employee_roster"]:
            employee_names.setdefault(clean_text(employee).casefold(), clean_text(employee))
        frame = frame.copy()
        frame["Employee"] = frame["Employee"].map(lambda v: employee_names[clean_text(v).casefold()])
        frame["Source File"] = name
        frames.append(frame)
        coverage = info["coverage"].copy()
        coverage["Employee"] = coverage["Employee"].map(lambda v: employee_names[clean_text(v).casefold()])
        coverage["Source File"] = name
        coverages.append(coverage)
        reports.append({"File": name, "Sheet": info["sheet"], "Header Row": info["header_row"], "Rows": len(frame)})
    if not frames:
        return pd.DataFrame(), {}, errors
    combined = pd.concat(frames, ignore_index=True, sort=False)
    roster = list(employee_names.values())
    info = {
        "sheet": ", ".join(dict.fromkeys(r["Sheet"] for r in reports)),
        "format": f"{len(reports)} Field - Visit workbook(s) combined",
        "header_row": ", ".join(str(r["Header Row"]) for r in reports),
        "rows": len(combined),
        "employees": len(roster),
        "roster_employees": len(roster),
        "active_employees": int(combined.loc[combined[FIELD_METRICS].sum(axis=1) > 0, "Employee"].nunique()),
        "dates": int(combined["Date"].nunique()),
        "employee_roster": roster,
        "coverage": pd.concat(coverages, ignore_index=True),
        "files": pd.DataFrame(reports),
    }
    return combined, info, errors


def field_employee_summary(field_frame):
    if field_frame.empty:
        return pd.DataFrame()
    agg_cols = FIELD_METRICS + ["Valid KM"]
    emp = field_frame.groupby("Employee", as_index=False)[agg_cols].sum()
    if "Designation" in field_frame.columns:
        des = field_frame.groupby("Employee")["Designation"].agg(
            lambda s: next((clean_text(v) for v in s if clean_text(v)), "Unassigned")
        )
        emp["Designation"] = emp["Employee"].map(des)
    emp["Lead / Visit %"] = np.where(emp["Visit"] > 0, emp["Lead"] / emp["Visit"] * 100, 0)
    emp["PI / Lead %"] = np.where(emp["Lead"] > 0, emp["PI"] / emp["Lead"] * 100, 0)
    emp["Order / Visit %"] = np.where(emp["Visit"] > 0, emp["Sales-Order"] / emp["Visit"] * 100, 0)
    emp["Order / PI %"] = np.where(emp["PI"] > 0, emp["Sales-Order"] / emp["PI"] * 100, 0)
    return emp


def add_kra_scores(emp, targets):
    out = emp.copy()
    active = [(m, float(t)) for m, t in targets.items() if float(t) > 0 and m in out.columns]
    if not active:
        out["KRA Score"] = 0.0
    else:
        pieces = [(out[m] / target * 100).clip(0, 100) for m, target in active]
        out["KRA Score"] = pd.concat(pieces, axis=1).mean(axis=1).round(1)
    out["KRA Performance"] = pd.cut(
        out["KRA Score"], [-np.inf, 40, 60, 75, 90, np.inf],
        labels=["Critical", "Needs Improvement", "Good", "Strong", "Excellent"]
    ).astype(str)
    return out


def simple_excel_bytes(dataframes):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, frame in dataframes.items():
            (frame.copy() if frame is not None else pd.DataFrame()).to_excel(writer, sheet_name=sheet[:31], index=False)
    output.seek(0)
    wb = load_workbook(output)
    header_fill = PatternFill("solid", fgColor="0F172A")
    header_font = Font(color="FFFFFF", bold=True)
    for ws in wb.worksheets:
        if ws.max_row >= 1:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ============================================================
# SERVICE DATE FILTER
# ============================================================
def apply_service_filters(df):
    """Independent, reliable Service date filter with selectable date basis."""
    if df.empty:
        return df
    f = df.copy()

    date_options = []
    if "Date_Parsed" in f.columns and f["Date_Parsed"].notna().any():
        date_options.append(("Query Created Date", "Date_Parsed"))
    elif "Created Date_Parsed" in f.columns and f["Created Date_Parsed"].notna().any():
        date_options.append(("Query Created Date", "Created Date_Parsed"))
    if "Closed Date_Parsed" in f.columns and (f["Status"].eq("Closed") & f["Closed Date_Parsed"].notna()).any():
        date_options.append(("Query Closed Date", "Closed Date_Parsed"))
    if "Follow Date_Parsed" in f.columns and f["Follow Date_Parsed"].notna().any():
        date_options.append(("Follow Date", "Follow Date_Parsed"))

    with st.sidebar:
        st.markdown('<div class="sidebar-section-num">🗓️ 7. SERVICE DATE FILTER</div>', unsafe_allow_html=True)
        if date_options:
            if not st.checkbox("Filter service tickets by date", value=False, key="service_date_filter_enabled"):
                st.caption("Showing all service tickets, including those without dates.")
                return f
            labels = [x[0] for x in date_options]
            basis = st.selectbox("Service Date Basis", labels, key="service_date_basis_v9")
            date_col = dict(date_options)[basis]
            if basis == "Query Closed Date":
                f = f[f["Status"].eq("Closed")]
            valid = f[date_col].dropna()
            if not valid.empty:
                mn, mx = valid.min().date(), valid.max().date()
                dv = st.date_input(
                    "Service Date Range",
                    value=(mn, mx),
                    min_value=mn,
                    max_value=mx,
                    key="service_dates_v9",
                    format="DD/MM/YYYY",
                )
                if isinstance(dv, (tuple, list)) and len(dv) == 2:
                    a = pd.Timestamp(dv[0]).normalize()
                    b = pd.Timestamp(dv[1]).normalize() + pd.Timedelta(days=1)
                    f = f[(f[date_col] >= a) & (f[date_col] < b)]
        else:
            st.caption("No usable Service date field found in the uploaded source.")
    return f


# ============================================================
# MAIN TOP HEADER NAVBAR
# ============================================================

st.markdown(company_theme("""
<style>
[data-testid="stMainBlockContainer"] { padding-top:4.4rem !important; }
.st-key-company_header {
    background:linear-gradient(115deg,#ffffff 55%,#fff7ed 100%);
    border:1px solid #e2e8f0; border-left:5px solid #f97316;
    border-radius:20px; padding:26px 24px; margin-bottom:16px;
    box-shadow:0 8px 28px rgba(15,23,42,.06);
}
.company-identity { min-height:82px; display:flex; align-items:center; }
.company-identity .pv-nav-subtext { line-height:1.7; margin-top:7px; }
.company-header-dot { color:#f97316; padding:0 5px; }
.st-key-company_header [data-testid="stSelectbox"] label p {
    font-size:10px; font-weight:800; letter-spacing:1.2px; color:#64748b;
}
.st-key-company_header [data-baseweb="select"] > div {
    min-height:48px; border:1px solid #fed7aa !important;
    border-radius:12px !important; background:#ffffff !important;
    box-shadow:0 2px 8px rgba(15,23,42,.03);
}
.st-key-company_header [data-baseweb="select"]:focus-within > div {
    border-color:#f97316 !important; box-shadow:0 0 0 3px #ffedd5;
}
@media(max-width:760px) {
    .st-key-company_header { padding:18px 14px; }
    .company-identity img { width:90px !important; }
    .company-identity { min-height:70px; }
}
</style>
"""), unsafe_allow_html=True)

company_logo = (
    '<div style="font-size:26px;font-weight:850;color:#166534;line-height:1.15;">BANGA<br>SOLAR</div>'
    if company == "Banga Solar Pvt Ltd" else
    f'<img src="{pv_logo_data_uri()}" alt="PV blink inverter" style="width:132px;height:auto;display:block;">'
)
header_brand.markdown(
    f"""
<div class="company-identity">
    <div style="display: flex; align-items: center;">
        {company_logo}
        <div style="margin-left: 18px; border-left: 2px solid #e2e8f0; padding-left: 18px;">
            <div style="font-size: 19px; font-weight: 850; color: #0f172a;">{company_title}</div>
            <div class="pv-nav-subtext">Sales-Report • Sales-Returns • Field-Visit • Employee Performance • Service-Report</div>
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# Sidebar Data Upload Section
with st.sidebar:
    st.markdown(f'<div class="sidebar-head-badge">{company_title} ENTERPRISE</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-section-num">📂 1. SALES & FIELD DATA</div>', unsafe_allow_html=True)
    st.caption("Three independent inputs. Field - Visit Report KPI/KRA data is never subtracted from Sales.")
    sales_files = st.file_uploader(
        "1️⃣ Sales Invoice Register", type=["xlsx", "xls", "csv"], key="sales_upload",
        accept_multiple_files=True, on_change=reset_register_upload_filters, args=("sales",),
        help="Upload multiple Sales Invoice files together, or browse again to add another month. Upload each report only once; all commercial lines are combined."
    )
    return_files = st.file_uploader(
        "2️⃣ Sales Return Register", type=["xlsx", "xls", "csv"], key="return_upload",
        accept_multiple_files=True, on_change=reset_register_upload_filters, args=("returns",),
        help="Upload multiple Sales Return files together, or browse again to add another month. Combined returns are deducted from Sales. Avoid overlapping reports."
    )
    field_files = st.file_uploader(
        "3️⃣ Field - Visit Reports", type=["xlsx", "xls"], key="field_dashboard_upload",
        accept_multiple_files=True, on_change=reset_field_upload_filters,
        help="Upload May, June, and other monthly workbooks together, or browse again to add files. Each needs a Dashboard sheet. Upload each report only once."
    )
    st.markdown('<div class="sidebar-section-num">🛠️ 2. SERVICE DATA</div>', unsafe_allow_html=True)
    service_files = st.file_uploader(
        "Upload Service Excel File(s)", type=["xlsx", "xls"], accept_multiple_files=True, key="service_upload",
        on_change=reset_register_upload_filters, args=("service",),
        help="Upload multiple Service files together, or browse again to add files. All valid files are combined. Avoid overlapping reports."
    )

# Load Datasets
sales_df, sales_cleaning, sales_errors = load_register_files(
    sales_files,
    lambda uploaded_file: process_data(uploaded_file, "Banga Solar Pvt Ltd" if company == "Banga Solar Pvt Ltd" else "PV-Blink"),
)
return_df, return_cleaning, return_errors = load_register_files(return_files, process_return_data)
if not sales_df.empty and not return_df.empty:
    sales_start, sales_end = sales_df["Invoice Date"].min(), sales_df["Invoice Date"].max()
    returns_start, returns_end = return_df["Sales Return Date"].min(), return_df["Sales Return Date"].max()
    if sales_end < returns_start or returns_end < sales_start:
        st.warning(
            f"Uploaded sales ({sales_start:%d %b %Y}–{sales_end:%d %b %Y}) and returns "
            f"({returns_start:%d %b %Y}–{returns_end:%d %b %Y}) cover different, non-overlapping periods. "
            "Upload reports for the same period, or select ?Match Sales Date Range? in the sidebar. "
            "Net Sales currently subtracts all uploaded returns."
        )

field_df = pd.DataFrame()
field_info = {}
field_errors = []
if field_files:
    field_df, field_info, field_errors = load_field_files(field_files)

service_df = pd.DataFrame()
service_detail = pd.DataFrame()
if service_files:
    service_df, service_detail = load_service_files(service_files)
    service_df = apply_service_filters(service_df)

for sales_error in sales_errors:
    st.error(f"Sales Invoice file could not be loaded: {sales_error}")
for return_error in return_errors:
    st.error(f"Sales Return file could not be loaded: {return_error}")
if "Error" in service_detail.columns:
    for _, failed_file in service_detail[service_detail["Error"].notna()].iterrows():
        st.error(f"Service file could not be loaded: {failed_file['File']}: {failed_file['Error']}")
for field_error in field_errors:
    st.error(f"Field - Visit Report could not be loaded: {field_error}")



# ============================================================
# ALL DASHBOARD CLICKS — SHORT 150MS LOADING
# ============================================================
st.iframe(
    company_theme("""
<script>
(function () {
    const doc = window.parent.document;

    function getOverlay() {
        let overlay = doc.getElementById("pv-upload-loading-overlay");

        if (!overlay) {
            overlay = doc.createElement("div");
            overlay.id = "pv-upload-loading-overlay";

            overlay.innerHTML = `
                <div style="
                    display:flex;
                    align-items:center;
                    gap:12px;
                    min-width:220px;
                    background:#ffffff;
                    border:1px solid #fed7aa;
                    border-left:4px solid #f97316;
                    border-radius:13px;
                    padding:13px 16px;
                    box-shadow:0 16px 36px rgba(15,23,42,.16);
                    font-family:'Segoe UI',Arial,sans-serif;
                    color:#0f172a;">
                    <div style="
                        width:22px;
                        height:22px;
                        flex:0 0 22px;
                        border:3px solid #ffedd5;
                        border-top-color:#f97316;
                        border-radius:50%;
                        animation:pvUploadSpin .65s linear infinite;">
                    </div>
                    <div>
                        <div style="
                            font-size:12.5px;
                            line-height:1.2;
                            font-weight:800;">
                            Loading Excel...
                        </div>
                        <div style="
                            margin-top:2px;
                            font-size:10px;
                            color:#64748b;
                            font-weight:500;">
                            Reading uploaded file
                        </div>
                    </div>
                </div>`;

            overlay.style.cssText = `
                position:fixed;
                inset:0;
                z-index:2147483646;
                display:none;
                align-items:center;
                justify-content:center;
                background:rgba(248,250,252,.20);
                backdrop-filter:blur(.5px);
                pointer-events:none;
            `;

            if (!doc.getElementById("pv-upload-loading-style")) {
                const style = doc.createElement("style");
                style.id = "pv-upload-loading-style";
                style.textContent =
                    "@keyframes pvUploadSpin{to{transform:rotate(360deg)}}";
                doc.head.appendChild(style);
            }

            doc.body.appendChild(overlay);
        }

        return overlay;
    }

    function showShortClickLoading() {
        const overlay = getOverlay();
        overlay.style.display = "flex";

        clearTimeout(window.parent.__pvShortClickTimer);

        // Very short feedback only: 150 ms.
        window.parent.__pvShortClickTimer = setTimeout(function () {
            overlay.style.display = "none";
        }, 150);
    }

    // Never leave an old overlay visible after a rerun.
    const overlay = getOverlay();
    overlay.style.display = "none";

    if (!window.parent.__pvAllClickLoadingBound) {
        window.parent.__pvAllClickLoadingBound = true;

        // Show the short loading feedback for interactive dashboard clicks.
        doc.addEventListener("click", function (event) {
            const target = event.target;
            if (!target || !target.closest) return;

            const interactive = target.closest(
                'button, [role="button"], [role="tab"], ' +
                '[data-baseweb="select"], [data-baseweb="checkbox"], ' +
                '[data-baseweb="radio"], [data-testid="stDateInput"], ' +
                '[data-testid="stDownloadButton"], a'
            );

            if (!interactive) return;

            const label = (interactive.innerText || interactive.textContent || "")
                .trim()
                .toLowerCase();
            const isAuthAction = label.includes("login") || label.includes("logout");

            if (isAuthAction) {
                const overlay = doc.getElementById("pv-upload-loading-overlay");
                if (overlay) overlay.style.display = "none";
                return;
            }

            showShortClickLoading();
        }, true);

        // File selection, dropdown/date changes, checkbox/radio changes.
        doc.addEventListener("change", function (event) {
            const target = event.target;
            if (!target || !target.matches) return;

            if (
                target.matches('input[type="file"]') ||
                target.matches('input') ||
                target.matches('select')
            ) {
                showShortClickLoading();
            }
        }, true);
    }
})();
</script>
"""),
    height=1,
)

# Filter Sales first; Return Register follows the SAME Sales filters.
filtered_sales = sales_df.copy()
if not sales_df.empty:
    filtered_sales = apply_sales_filters(sales_df)
filtered_returns = apply_return_filters(return_df) if not return_df.empty else pd.DataFrame()
filtered_returns = assign_return_representative(filtered_returns, sales_df) if not filtered_returns.empty else filtered_returns
net = net_metrics(filtered_sales, filtered_returns)

# ============================================================
# MAIN NAVIGATION — robust text navigation (no Streamlit tab overflow)
# ============================================================
PAGE_OPTIONS = [
    "Executive Overview",
    "Sales Intelligence",
    "Sales Target Performance",
    "Sales Returns & Net",
    "Field KPI / KRA",
    "Service Operations",
    "Quality & Export",
]

selected_page = st.radio(
    "Dashboard navigation",
    PAGE_OPTIONS,
    horizontal=True,
    label_visibility="collapsed",
    key="pv_main_navigation",
)

# ============================================================
# 1. EXECUTIVE OVERVIEW
# ============================================================
if selected_page == "Executive Overview":
    sales_amount = net["gross_amount"]
    sales_qty = net["gross_qty"]
    return_amount = net["return_amount"]
    return_qty = net["return_qty"]
    net_sales_amount = net["net_amount"]
    net_sales_qty = net["net_qty"]
    invoices_count = int(filtered_sales["Invoice No#"].nunique()) if not filtered_sales.empty and "Invoice No#" in filtered_sales else len(filtered_sales)
    avg_invoice_val = net_sales_amount / invoices_count if invoices_count else 0
    customers_count = int(filtered_sales["Ledger Name"].nunique()) if not filtered_sales.empty else 0
    total_service = int(service_df.loc[service_df["Ticket Valid"]].shape[0]) if not service_df.empty else 0
    field_employees = int(field_info.get("roster_employees", field_df["Employee"].nunique() if not field_df.empty else 0))

    st.markdown(
        '<div class="pv-net-banner"><b>Net Sales:</b> Normal Sales − Sales Return Register.'
        ' Field - Visit Report activity is independent and is never deducted from sales.</div>',
        unsafe_allow_html=True
    )
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1:
        st.markdown(kpi_card_html("₹", "GROSS SALES", format_indian_currency(sales_amount), "Normal Sales", format_indian_currency_full(sales_amount)), unsafe_allow_html=True)
    with k2:
        st.markdown(kpi_card_html("↩️", "RETURN AMOUNT", format_indian_currency(return_amount), "Return Register", format_indian_currency_full(return_amount)), unsafe_allow_html=True)
    with k3:
        st.markdown(kpi_card_html("✅", "NET SALES", format_indian_currency(net_sales_amount), "Gross − Return", format_indian_currency_full(net_sales_amount)), unsafe_allow_html=True)
    with k4:
        st.markdown(kpi_card_html("📦", "GROSS QTY", f"{sales_qty:,.0f}", "Normal Sales QTY", format_indian_number_full(sales_qty)), unsafe_allow_html=True)
    with k5:
        st.markdown(kpi_card_html("↩️", "RETURN QTY", f"{return_qty:,.0f}", "Returned QTY", format_indian_number_full(return_qty)), unsafe_allow_html=True)
    with k6:
        st.markdown(kpi_card_html("⚡", "NET QTY", f"{net_sales_qty:,.0f}", "Gross − Return", format_indian_number_full(net_sales_qty)), unsafe_allow_html=True)

    st.write("")
    x1, x2, x3, x4 = st.columns(4)
    with x1:
        st.markdown(kpi_card_html("📄", "INVOICES", f"{invoices_count:,}", "Gross invoices", format_indian_number_full(invoices_count)), unsafe_allow_html=True)
    with x2:
        st.markdown(kpi_card_html("👥", "CUSTOMERS", f"{customers_count:,}", "Active ledgers", format_indian_number_full(customers_count)), unsafe_allow_html=True)
    with x3:
        st.markdown(kpi_card_html("📈", "FIELD EMPLOYEES", f"{field_employees:,}", "Field Visit source", format_indian_number_full(field_employees)), unsafe_allow_html=True)
    with x4:
        st.markdown(kpi_card_html("🛠️", "SERVICE TICKETS", f"{total_service:,}", "Maintenance workload", format_indian_number_full(total_service)), unsafe_allow_html=True)

    st.write("")

    if not filtered_sales.empty:
        # Row 1 Header with Integrated Sleek Time Grain Selectbox
        hdr_col1, hdr_col2 = st.columns([3, 1])
        with hdr_col1:
            st.markdown('<div style="font-size: 17px; font-weight: 800; color: #0f172a; margin-top: 4px;">📈 Executive Revenue Performance Dashboard</div>', unsafe_allow_html=True)
        with hdr_col2:
            time_grain = st.selectbox(
                "Granularity",
                ["Monthly", "Weekly", "Quarterly", "Yearly"],
                key="exec_time_grain",
                label_visibility="collapsed"
            )

        df_trend = filtered_sales.copy()
        if time_grain == "Weekly":
            df_trend["GrainLabel"] = "W" + df_trend["Week"].astype(str) + " (" + df_trend["Year"].astype(str) + ")"
            trend_group = df_trend.groupby(["Year", "Week", "GrainLabel"], as_index=False)["Amount"].sum().sort_values(["Year", "Week"])
        elif time_grain == "Quarterly":
            df_trend["GrainLabel"] = df_trend["Year"].astype(str) + " " + df_trend["Quarter"]
            trend_group = df_trend.groupby(["Year", "Quarter", "GrainLabel"], as_index=False)["Amount"].sum().sort_values(["Year", "Quarter"])
        elif time_grain == "Yearly":
            df_trend["GrainLabel"] = df_trend["Year"].astype(str)
            trend_group = df_trend.groupby(["Year", "GrainLabel"], as_index=False)["Amount"].sum().sort_values("Year")
        else:
            df_trend["GrainLabel"] = pd.to_datetime(df_trend["Year-Month"].astype(str)).dt.strftime("%b %Y")
            trend_group = df_trend.groupby(["Year-Month", "GrainLabel"], as_index=False)["Amount"].sum().sort_values("Year-Month")

        trend_group["FormattedSales"] = trend_group["Amount"].map(format_indian_currency)

        # Row 1: Revenue Trend, Sales by Item Group, Sales by Customer
        c1, c2, c3 = st.columns([1.15, 0.85, 1.0])

        with c1:
            st.markdown(f'''
            <div class="pv-chart-card">
                <div class="pv-card-header">
                    <div class="pv-card-title">📈 Revenue Trend ({time_grain})</div>
                </div>
            ''', unsafe_allow_html=True)
            fig = px.line(trend_group, x="GrainLabel", y="Amount", text="FormattedSales", markers=True, color_discrete_sequence=[company_theme("#f97316")])
            fig.update_traces(
                mode="lines+markers+text",
                textposition="top center",
                textfont=dict(size=10.5, color="#0f172a", weight=700),
                line=dict(width=3),
                marker=dict(size=8, color=company_theme("#f97316"), symbol="circle")
            )
            fig = clean_plotly_fig(fig, height=330)
            fig.update_yaxes(tickprefix="₹", separatethousands=True)
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
            st.markdown('</div>', unsafe_allow_html=True)

        with c2:
            st.markdown('''
            <div class="pv-chart-card">
                <div class="pv-card-header">
                    <div class="pv-card-title">🍩 Item Group Share (Top 10)</div>
                </div>
            ''', unsafe_allow_html=True)
            group_df = filtered_sales.groupby("Item Group", as_index=False)["Amount"].sum().sort_values("Amount", ascending=False).head(10)
            group_df["CleanGroup"] = group_df["Item Group"].map(lambda x: clean_label(x, 20))
            group_df["FormattedVal"] = group_df["Amount"].map(format_indian_currency)

            fig = px.pie(
                group_df, values="Amount", names="CleanGroup", hole=0.48,
                color_discrete_sequence=[company_theme("#f97316"), company_theme("#fb923c"), company_theme("#fdba74"), "#38bdf8", "#818cf8", "#cbd5e1", "#f43f5e", "#a855f7"]
            )
            fig = clean_plotly_fig(fig, height=330)
            fig.update_traces(
                textinfo="percent",
                textposition="inside",
                insidetextfont=dict(color="#ffffff", size=10, weight=700)
            )
            fig.update_layout(legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=10)))
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
            st.markdown('</div>', unsafe_allow_html=True)

        with c3:
            st.markdown('''
            <div class="pv-chart-card">
                <div class="pv-card-header">
                    <div class="pv-card-title">🏢 Top Customers by Revenue</div>
                </div>
            ''', unsafe_allow_html=True)
            cust_df = filtered_sales.groupby("Ledger Name", as_index=False)["Amount"].sum().sort_values("Amount", ascending=False).head(10)
            cust_df["CleanLabel"] = cust_df["Ledger Name"].map(lambda x: clean_label(x, 22))
            cust_df["FormattedVal"] = cust_df["Amount"].map(format_indian_currency)

            fig = px.bar(
                cust_df.sort_values("Amount"), x="Amount", y="Ledger Name", orientation="h",
                color_discrete_sequence=[company_theme("#f97316")]
            )
            max_c_val = float(cust_df["Amount"].max()) if not cust_df.empty else 0
            fig = clean_plotly_fig(fig, height=max(400, len(cust_df) * 36 + 110), is_horizontal=True, max_val=max_c_val)
            fig.update_yaxes(type="category", dtick=1, automargin=True)
            fig.update_traces(
                text=cust_df.sort_values("Amount")["FormattedVal"],
                textposition="outside",
                textfont=dict(size=10.5, color="#0f172a", weight=700),
                marker_line_color=company_theme("#ea580c"),
                marker_line_width=1
            )
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
            st.markdown('</div>', unsafe_allow_html=True)

        # Row 2: Top Items Table, Day of Week, Quick Insights
        c1, c2, c3 = st.columns([1.2, 0.9, 0.9])

        with c1:
            st.markdown('''
            <div class="pv-chart-card">
                <div class="pv-card-header">
                    <div class="pv-card-title">🏆 Top 10 Products by Sales Revenue</div>
                </div>
            ''', unsafe_allow_html=True)
            top_items = filtered_sales.groupby(["Item Name", "Item Group"], as_index=False).agg(
                Qty=("QTY", "sum"), Sales=("Amount", "sum")
            ).sort_values("Sales", ascending=False).head(10).reset_index(drop=True)
            top_items.index = top_items.index + 1
            top_items["Sales (₹)"] = top_items["Sales"].map(format_indian_currency)
            top_items["Qty"] = top_items["Qty"].map(lambda x: f"{x:,.0f}")
            st.dataframe(top_items[["Item Name", "Item Group", "Qty", "Sales (₹)"]], width="stretch", height=270)
            st.markdown('</div>', unsafe_allow_html=True)

        with c2:
            st.markdown('''
            <div class="pv-chart-card">
                <div class="pv-card-header">
                    <div class="pv-card-title">📅 Revenue by Day of Week</div>
                </div>
            ''', unsafe_allow_html=True)
            df_day = filtered_sales.copy()
            df_day["DayName"] = df_day["Invoice Date"].dt.day_name()
            day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            day_sales = df_day.groupby("DayName", as_index=False)["Amount"].sum()
            day_sales["DayName"] = pd.Categorical(day_sales["DayName"], categories=day_order, ordered=True)
            day_sales = day_sales.sort_values("DayName")
            day_sales["FormattedVal"] = day_sales["Amount"].map(format_indian_currency)

            fig = px.bar(day_sales, x="DayName", y="Amount", color_discrete_sequence=[company_theme("#fb923c")])
            fig = clean_plotly_fig(fig, height=270)
            fig.update_traces(
                text=day_sales["FormattedVal"],
                textposition="outside",
                textfont=dict(size=10, color="#0f172a", weight=700)
            )
            fig.update_yaxes(tickprefix="₹", separatethousands=True)
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
            st.markdown('</div>', unsafe_allow_html=True)

        with c3:
            st.markdown('''
            <div class="pv-chart-card">
                <div class="pv-card-header">
                    <div class="pv-card-title">💡 Strategic Executive Insights</div>
                </div>
            ''', unsafe_allow_html=True)
            top_grp_name = group_df.iloc[0]["Item Group"] if not group_df.empty else "N/A"
            top_cust_name = cust_df.iloc[0]["Ledger Name"] if not cust_df.empty else "N/A"
            top_day_name = day_sales.sort_values("Amount", ascending=False).iloc[0]["DayName"] if not day_sales.empty else "N/A"

            st.markdown(f'''
            <div class="pv-insight-item"><div class="pv-insight-icon">📈</div><div class="pv-insight-text">Net Revenue After Returns: <b>{format_indian_currency(net_sales_amount)}</b></div></div>
            <div class="pv-insight-item"><div class="pv-insight-icon">📦</div><div class="pv-insight-text">Highest Volume Category: <b>{top_grp_name}</b></div></div>
            <div class="pv-insight-item"><div class="pv-insight-icon">🏢</div><div class="pv-insight-text">Top Key Customer: <b>{clean_label(top_cust_name, 26)}</b></div></div>
            <div class="pv-insight-item"><div class="pv-insight-icon">📅</div><div class="pv-insight-text">Peak Sales Day: <b>{top_day_name}</b></div></div>
            ''', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("💡 Upload an E-Khata Sales Excel file from the left sidebar to generate Executive Insights.")

# ============================================================
# 2. SALES INTELLIGENCE & DEEP ANALYTICS
# ============================================================
if selected_page == "Sales Intelligence":
    st.markdown('''
    <div class="pv-chart-card">
        <div class="pv-card-header">
            <div class="pv-card-title">💰 Product & Customer Commercial Intelligence</div>
        </div>
    ''', unsafe_allow_html=True)
    if filtered_sales.empty:
        st.info("Upload Sales Excel data to view detailed product and customer performance analytics.")
    else:
        st.markdown(
            f'<div class="pv-net-banner"><b>Filtered Net Sales:</b> {format_indian_currency(net["net_amount"])}'
            f' &nbsp;|&nbsp; <b>Net QTY:</b> {net["net_qty"]:,.0f}'
            f' &nbsp;|&nbsp; Return deducted: {format_indian_currency(net["return_amount"])}.'
            'Charts below preserve your original Gross Sales analytics; reconciled Net analytics are in the Returns & Net tab.</div>',
            unsafe_allow_html=True
        )
        with st.container():
            st.caption("Showing the specified inverter products by KW size, phase, and MPPT count, plus Battery Pack, all representatives, and the top 15 customers / ledgers within the current filters.")
            st.markdown('##### 🏆 All Products by Quantity — KW, Phase & MPPT / Battery Pack')
            # This allowlist applies only to this chart. Match complete item names,
            # tolerating case, extra whitespace and exported HTML line breaks.
            chart_product_names = PRODUCT_CHART_CATALOG
            chart_product_groups = {
                name.casefold(): (
                    "Battery Pack" if "BATTERY PACK" in name
                    else name.split(" Ongrid", 1)[0] + " " + re.search(r"\d+ MPPT\b", name).group(0)
                )
                for name in chart_product_names
            }
            chart_item_keys = (
                filtered_sales["Item Name"].astype("string")
                .str.replace(r"(?i)<br\s*/?>", " ", regex=True)
                .str.replace(r"\s+", " ", regex=True)
                .str.strip().str.casefold()
            )
            chart_groups = chart_item_keys.map(chart_product_groups)
            prod_source = filtered_sales.loc[chart_groups.notna()].copy()
            prod_source["Product"] = chart_groups.loc[chart_groups.notna()]
            prod = (
                prod_source.groupby("Product", as_index=False, dropna=False)
                .agg(QTY=("QTY", "sum"), Sales=("Amount", "sum"))
                .sort_values("QTY", ascending=False)
            )
            prod["FormattedQty"] = prod["QTY"].map(lambda x: f"{x:,.0f} units")
            chart_prod = prod.sort_values("QTY", ascending=True)
            fig = px.bar(chart_prod, x="QTY", y="Product", orientation="h", color_discrete_sequence=[company_theme("#f97316")])
            max_q_val = float(prod["QTY"].max()) if not prod.empty else 0
            fig = clean_plotly_fig(fig, height=max(440, len(prod) * 34 + 120), is_horizontal=True, max_val=max_q_val)
            fig.update_yaxes(type="category", dtick=1, automargin=True)
            fig.update_traces(
                text=chart_prod["FormattedQty"],
                textposition="outside",
                textfont=dict(size=11, color="#0f172a", weight=700),
                hovertemplate="<b>%{y}</b><br>QTY: %{x:,.0f}<extra></extra>",
            )
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        with st.container():
            st.markdown('##### 🏢 Top 15 Customers / Ledgers by Revenue')
            cust = filtered_sales.groupby("Ledger Name", as_index=False, dropna=False).agg(Sales=("Amount", "sum")).sort_values("Sales", ascending=False).head(15)
            cust["FormattedSales"] = cust["Sales"].map(format_indian_currency)

            fig = px.bar(cust.sort_values("Sales"), x="Sales", y="Ledger Name", orientation="h", color_discrete_sequence=[company_theme("#fb923c")])
            max_s_val = float(cust["Sales"].max()) if not cust.empty else 0
            fig = clean_plotly_fig(fig, height=max(440, len(cust) * 34 + 120), is_horizontal=True, max_val=max_s_val)
            fig.update_yaxes(type="category", dtick=1, automargin=True)
            fig.update_traces(
                text=cust.sort_values("Sales")["FormattedSales"],
                textposition="outside",
                textfont=dict(size=10.5, color="#0f172a", weight=700)
            )
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        rep_col = sales_representative_column(filtered_sales)
        rep = filtered_sales.groupby(rep_col, as_index=False, dropna=False).agg(Sales=("Amount", "sum"), QTY=("QTY", "sum")).sort_values("Sales", ascending=False)
        rep["FormattedSales"] = rep["Sales"].map(format_indian_currency)
        rep["FormattedQty"] = rep["QTY"].map(lambda x: f"{x:,.0f} units")

        with st.container():
            st.markdown('##### 👤 Sales Contribution by Representative (₹)')
            fig = px.bar(rep.sort_values("Sales"), x="Sales", y=rep_col, orientation="h", color_discrete_sequence=[company_theme("#f97316")])
            max_rep_s = float(rep["Sales"].max()) if not rep.empty else 0
            fig = clean_plotly_fig(fig, height=max(480, len(rep) * 34 + 120), is_horizontal=True, max_val=max_rep_s)
            fig.update_yaxes(type="category", dtick=1, automargin=True)
            fig.update_traces(
                text=rep.sort_values("Sales")["FormattedSales"],
                textposition="outside",
                textfont=dict(size=10.5, color="#0f172a", weight=700)
            )
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        with st.container():
            st.markdown('##### 📦 Unit Sales Volume by Representative')
            fig = px.bar(rep.sort_values("QTY"), x="QTY", y=rep_col, orientation="h", color_discrete_sequence=["#0284c7"])
            max_rep_q = float(rep["QTY"].max()) if not rep.empty else 0
            fig = clean_plotly_fig(fig, height=max(480, len(rep) * 34 + 120), is_horizontal=True, max_val=max_rep_q)
            fig.update_yaxes(type="category", dtick=1, automargin=True)
            fig.update_traces(
                text=rep.sort_values("QTY")["FormattedQty"],
                textposition="outside",
                textfont=dict(size=10.5, color="#0f172a", weight=700)
            )
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        # -------- Advanced Net Sales Analytics (Gross - Return) --------
        st.markdown("##### ⚡ Advanced Net Sales Analytics")

        with st.container():
            sales_monthly_net = filtered_sales.groupby("Year-Month", as_index=False).agg(
                **{"Gross Sales": ("Amount", "sum"), "Gross QTY": ("QTY", "sum")}
            )
            if not filtered_returns.empty:
                ret_month = filtered_returns.groupby("Return Year-Month", as_index=False).agg(
                    **{"Return Amount": ("Amount", "sum"), "Return QTY": ("QTY", "sum")}
                ).rename(columns={"Return Year-Month": "Year-Month"})
                sales_monthly_net = sales_monthly_net.merge(ret_month, on="Year-Month", how="outer")
            else:
                sales_monthly_net["Return Amount"] = 0.0
                sales_monthly_net["Return QTY"] = 0.0
            for cc in ["Gross Sales", "Gross QTY", "Return Amount", "Return QTY"]:
                sales_monthly_net[cc] = pd.to_numeric(sales_monthly_net[cc], errors="coerce").fillna(0)
            sales_monthly_net["Net Sales"] = sales_monthly_net["Gross Sales"] - sales_monthly_net["Return Amount"]
            sales_monthly_net = sales_monthly_net.sort_values("Year-Month")
            sales_monthly_net["Month"] = pd.to_datetime(sales_monthly_net["Year-Month"]).dt.strftime("%b %Y")
            fig = px.line(
                sales_monthly_net, x="Month", y=["Gross Sales", "Net Sales"], markers=True,
                color_discrete_sequence=[company_theme("#f97316"), "#16a34a"]
            )
            fig = clean_plotly_fig(fig, height=390)
            fig.update_yaxes(tickprefix="₹", separatethousands=True)
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        with st.container():
            emp_net, _ = employee_net_performance(filtered_sales, filtered_returns)
            if not emp_net.empty:
                emp_net = emp_net.sort_values("Achievement", ascending=False)
                emp_net["Net Label"] = emp_net["Achievement"].map(format_indian_currency)
                fig = px.bar(
                    emp_net.sort_values("Achievement"), x="Achievement", y="Employee",
                    orientation="h", color_discrete_sequence=["#16a34a"]
                )
                max_emp = float(emp_net["Achievement"].clip(lower=0).max()) if not emp_net.empty else 0
                fig = clean_plotly_fig(fig, height=max(390, len(emp_net) * 34 + 120), is_horizontal=True, max_val=max_emp if emp_net["Achievement"].ge(0).all() else None)
                fig.update_yaxes(type="category", dtick=1, automargin=True)
                fig.update_traces(
                    text=emp_net.sort_values("Achievement")["Net Label"], textposition="outside"
                )
                st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
            else:
                st.info("Representative data is not available for Net Sales analytics.")

        # Advanced Branch & Profitability Deep Dive
        if "Branch" in filtered_sales.columns:
            st.markdown('##### 🏢 Branch Performance Breakdown')
            branch_df = filtered_sales.groupby("Branch", as_index=False).agg(
                Sales=("Amount", "sum"), Qty=("QTY", "sum"), Invoices=("Invoice Date", "count")
            ).sort_values("Sales", ascending=False)
            branch_df["Sales (₹)"] = branch_df["Sales"].map(format_indian_currency)
            branch_df["Qty"] = branch_df["Qty"].map(lambda x: f"{x:,.0f}")
            st.dataframe(branch_df[["Branch", "Qty", "Invoices", "Sales (₹)"]], width="stretch", hide_index=True)

    st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# 3. SALES TARGET PERFORMANCE — SALES INVOICE REGISTER ONLY
# ============================================================
if selected_page == "Sales Target Performance":
    st.markdown("""
    <div class="pv-chart-card">
        <div class="pv-card-header">
            <div class="pv-card-title">🎯 Sales Invoice Register • Employee Target Performance</div>
        </div>
    """, unsafe_allow_html=True)

    st.caption(
        "Target Performance uses Sales Invoice Register only. "
        "Achievement = Sales Invoice Register Amount. "
        "Sales Return Register is NOT deducted."
    )

    if sales_df.empty:
        st.info(
            "Upload Sales Invoice Register data to activate "
            "employee target performance."
        )
    else:
        rep_col = sales_representative_column(sales_df)

        if rep_col not in sales_df.columns:
            st.warning(
                "Representative Ref / Representative User is not "
                "available in the Sales Invoice Register."
            )
        else:
            # ========================================================
            # COMPLETE SALES EMPLOYEE ROSTER
            # ========================================================
            sales_target_source = sales_df.copy()
            sales_target_source["_Target Employee"] = (
                sales_target_source[rep_col]
                .astype("string")
                .fillna("")
                .map(canonical_sales_employee)
            )

            all_target_emp_names = sales_target_roster(
                sales_target_source["_Target Employee"]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )

            # ========================================================
            # PAGE-SPECIFIC FILTERS
            # ========================================================
            st.markdown("##### 🎛️ Target Performance Filters")

            target_work = sales_target_source.copy()

            target_dates = pd.to_datetime(
                target_work["Invoice Date"],
                errors="coerce",
            ).dropna()

            target_start_date = None
            target_end_date = None

            # Visual order:
            # 1) Target button
            # 2) Employee choose options
            # 3) ONE date range
            # 4) Reset
            f_target, f_emp, f_date, f_reset = st.columns(
                [0.82, 1.45, 1.75, 0.62]
            )

            # Execute date input first so the popup can use the selected
            # range while still appearing visually as the third control.
            with f_date:
                if not target_dates.empty:
                    target_min_date = target_dates.min().date()
                    target_max_date = target_dates.max().date()

                    target_date_value = st.date_input(
                        "📅 Sales Invoice Date Range",
                        value=(target_min_date, target_max_date),
                        min_value=target_min_date,
                        max_value=target_max_date,
                        key="sales_target_date_filter",
                        format="DD/MM/YYYY",
                    )

                    if (
                        isinstance(target_date_value, (tuple, list))
                        and len(target_date_value) == 2
                    ):
                        target_start_date = pd.Timestamp(
                            target_date_value[0]
                        )
                        target_end_date = (
                            pd.Timestamp(target_date_value[1])
                            + pd.Timedelta(days=1)
                        )
                    else:
                        target_start_date = pd.Timestamp(
                            target_date_value
                        )
                        target_end_date = (
                            target_start_date
                            + pd.Timedelta(days=1)
                        )
                else:
                    st.info("Invoice Date is not available.")

            with f_emp:
                selected_target_emps = st.multiselect(
                    "👤 Employee",
                    options=all_target_emp_names,
                    default=[],
                    key="sales_target_emp_filter",
                    help="Leave blank to analyse all Sales employees.",
                )

            # Compact EMPLOYEE-WISE Year / Month / Date target popup LEFT of Choose options.
            with f_target:
                st.write("")
                if hasattr(st, "popover"):
                    with st.popover(
                        "🎯 Targets",
                        width="stretch",
                    ):
                        target_df = target_editor(
                            all_target_emp_names,
                            selected_start=target_start_date,
                            selected_end=(
                                target_end_date - pd.Timedelta(days=1)
                                if target_end_date is not None
                                else None
                            ),
                        )
                else:
                    # Fallback for old Streamlit versions.
                    with st.expander("🎯 Targets", expanded=False):
                        target_df = target_editor(
                            all_target_emp_names,
                            selected_start=target_start_date,
                            selected_end=(
                                target_end_date - pd.Timedelta(days=1)
                                if target_end_date is not None
                                else None
                            ),
                        )

            with f_reset:
                st.write("")
                st.write("")
                if st.button(
                    "↻ Reset",
                    width="stretch",
                    key="sales_target_reset_filters",
                ):
                    st.session_state.pop(
                        "sales_target_emp_filter", None
                    )
                    st.session_state.pop(
                        "sales_target_date_filter", None
                    )
                    st.rerun()

            # Apply only the two requested Target Performance filters.
            if target_start_date is not None:
                invoice_dates = pd.to_datetime(
                    target_work["Invoice Date"],
                    errors="coerce",
                )
                target_work = target_work[
                    invoice_dates.ge(target_start_date)
                    & invoice_dates.lt(target_end_date)
                ].copy()

            if selected_target_emps:
                target_work = target_work[
                    target_work["_Target Employee"]
                    .astype(str)
                    .isin(selected_target_emps)
                ].copy()

            if target_work.empty:
                st.warning(
                    "No Sales Invoice data matches the selected "
                    "Employee / Date filters."
                )
            else:
                perf = sales_invoice_target_performance(
                    target_work,
                    target_df,
                )

                overall_target = (
                    float(perf["Target"].sum())
                    if not perf.empty
                    else 0.0
                )
                overall_achievement = (
                    float(perf["Achievement"].sum())
                    if not perf.empty
                    else 0.0
                )
                overall_pct = (
                    overall_achievement / overall_target * 100
                    if overall_target > 0
                    else np.nan
                )
                overall_gap = max(
                    overall_target - overall_achievement,
                    0.0,
                )
                total_qty = float(
                    pd.to_numeric(
                        target_work["QTY"],
                        errors="coerce",
                    ).fillna(0).sum()
                )
                product_count = (
                    int(target_work["Item Name"].nunique())
                    if "Item Name" in target_work.columns
                    else 0
                )
                group_count = (
                    int(target_work["Item Group"].nunique())
                    if "Item Group" in target_work.columns
                    else 0
                )
                selected_employee_count = (
                    int(target_work["_Target Employee"].nunique())
                )

                st.markdown(
                    "##### 📊 Overall Target & Achievement"
                )

                k1, k2 = st.columns(2)
                k3, k4 = st.columns(2)

                with k1:
                    st.markdown(
                        kpi_card_html(
                            "🎯",
                            "OVERALL TARGET",
                            format_indian_currency(
                                overall_target
                            ),
                            "Selected employee target",
                            format_indian_currency_full(
                                overall_target
                            ),
                        ),
                        unsafe_allow_html=True,
                    )

                with k2:
                    st.markdown(
                        kpi_card_html(
                            "✅",
                            "ACHIEVEMENT",
                            format_indian_currency(
                                overall_achievement
                            ),
                            "Sales Invoice Amount",
                            format_indian_currency_full(
                                overall_achievement
                            ),
                        ),
                        unsafe_allow_html=True,
                    )

                pct_text = (
                    f"{overall_pct:.2f}%"
                    if pd.notna(overall_pct)
                    else "—"
                )

                with k3:
                    st.markdown(
                        kpi_card_html(
                            "📈",
                            "ACHIEVEMENT %",
                            pct_text,
                            "Achievement ÷ Target",
                            pct_text,
                        ),
                        unsafe_allow_html=True,
                    )

                with k4:
                    st.markdown(
                        kpi_card_html(
                            "⚡",
                            "TARGET GAP",
                            format_indian_currency(
                                overall_gap
                            ),
                            "Remaining target",
                            format_indian_currency_full(
                                overall_gap
                            ),
                        ),
                        unsafe_allow_html=True,
                    )

                # ========================================================
                # SALES ACTIVITY KPI STRIP
                # ========================================================
                st.markdown("##### 📦 Sales Mix & Quantity Overview")

                q1, q2, q3, q4 = st.columns(4)

                with q1:
                    st.markdown(
                        kpi_card_html(
                            "👥",
                            "EMPLOYEES",
                            f"{selected_employee_count:,}",
                            "Current selection",
                            format_indian_number_full(
                                selected_employee_count
                            ),
                        ),
                        unsafe_allow_html=True,
                    )

                with q2:
                    st.markdown(
                        kpi_card_html(
                            "📦",
                            "TOTAL QTY",
                            f"{total_qty:,.0f}",
                            "Sales Invoice QTY",
                            format_indian_number_full(total_qty),
                        ),
                        unsafe_allow_html=True,
                    )

                with q3:
                    st.markdown(
                        kpi_card_html(
                            "⚙️",
                            "PRODUCTS",
                            f"{product_count:,}",
                            "Unique Item Name",
                            format_indian_number_full(
                                product_count
                            ),
                        ),
                        unsafe_allow_html=True,
                    )

                with q4:
                    st.markdown(
                        kpi_card_html(
                            "🧩",
                            "ITEM GROUPS",
                            f"{group_count:,}",
                            "Unique Item Group",
                            format_indian_number_full(
                                group_count
                            ),
                        ),
                        unsafe_allow_html=True,
                    )

                # ========================================================
                # EMPLOYEE TARGET VS ACHIEVEMENT — TARGET SET ONLY
                # ========================================================
                st.markdown(
                    "##### 👤 Employee Target vs Achievement"
                )

                if perf.empty:
                    st.info(
                        "No employee performance data is available."
                    )
                else:
                    target_set_df = perf[
                        perf["Target"] > 0
                    ].copy()

                    not_set_count = int(
                        (perf["Target"] <= 0).sum()
                    )

                    if target_set_df.empty:
                        st.info(
                            "Enter a non-zero target for at least one employee. "
                            "Employees with ₹0 target are intentionally excluded "
                            "from the Target vs Achievement comparison."
                        )
                    else:
                        if not_set_count > 0:
                            st.caption(
                                f"{not_set_count} employee(s) with ₹0 target are "
                                "excluded from this comparison chart so the scale "
                                "remains accurate and readable."
                            )

                        chart_df = target_set_df.copy()
                        chart_df["Clean Employee"] = (
                            chart_df["Employee"].map(
                                lambda x: clean_label(x, 28)
                            )
                        )

                        chart_df["Target Label"] = (
                            chart_df["Target"].map(
                                format_indian_currency
                            )
                        )
                        chart_df["Achievement Label"] = (
                            chart_df["Achievement"].map(
                                format_indian_currency
                            )
                        )
                        chart_df["Pct Label"] = (
                            chart_df["Achievement %"].map(
                                lambda x: f"{x:.1f}%"
                                if pd.notna(x)
                                else "—"
                            )
                        )

                        # Highest achievement percentage first.
                        chart_df = chart_df.sort_values(
                            ["Achievement %", "Achievement"],
                            ascending=[True, True],
                            na_position="first",
                        )

                        fig = go.Figure()

                        fig.add_trace(
                            go.Bar(
                                name="Target",
                                y=chart_df["Clean Employee"],
                                x=chart_df["Target"],
                                orientation="h",
                                marker_color="#0284c7",
                                text=chart_df["Target Label"],
                                textposition="outside",
                                customdata=np.stack(
                                    [
                                        chart_df["Employee"],
                                        chart_df["Target"],
                                        chart_df["Achievement"],
                                        chart_df["Achievement %"],
                                    ],
                                    axis=-1,
                                ),
                                hovertemplate=(
                                    "<b>%{customdata[0]}</b>"
                                    "<br>Target: ₹%{customdata[1]:,.2f}"
                                    "<br>Achievement: ₹%{customdata[2]:,.2f}"
                                    "<br>Achievement %: %{customdata[3]:.1f}%"
                                    "<extra></extra>"
                                ),
                            )
                        )

                        fig.add_trace(
                            go.Bar(
                                name="Achievement",
                                y=chart_df["Clean Employee"],
                                x=chart_df["Achievement"],
                                orientation="h",
                                marker_color=company_theme("#f97316"),
                                text=chart_df["Achievement Label"],
                                textposition="outside",
                                customdata=np.stack(
                                    [
                                        chart_df["Employee"],
                                        chart_df["Target"],
                                        chart_df["Achievement"],
                                        chart_df["Achievement %"],
                                    ],
                                    axis=-1,
                                ),
                                hovertemplate=(
                                    "<b>%{customdata[0]}</b>"
                                    "<br>Target: ₹%{customdata[1]:,.2f}"
                                    "<br>Achievement: ₹%{customdata[2]:,.2f}"
                                    "<br>Achievement %: %{customdata[3]:.1f}%"
                                    "<extra></extra>"
                                ),
                            )
                        )

                        max_value = float(
                            chart_df[
                                ["Target", "Achievement"]
                            ].max().max()
                        ) if not chart_df.empty else 0.0

                        fig = clean_plotly_fig(
                            fig,
                            height=max(
                                390,
                                min(
                                    650,
                                    len(chart_df) * 45 + 150,
                                ),
                            ),
                            is_horizontal=True,
                            max_val=max_value,
                        )

                        fig.update_layout(
                            barmode="group",
                            bargap=0.22,
                            bargroupgap=0.08,
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=1.02,
                                xanchor="right",
                                x=1,
                            ),
                        )

                        # Percentage shown beside employee name via annotation.
                        for row_i, row in chart_df.iterrows():
                            fig.add_annotation(
                                x=0,
                                y=row["Clean Employee"],
                                text=f'  {row["Pct Label"]}',
                                showarrow=False,
                                xanchor="right",
                                xshift=-6,
                                font=dict(
                                    size=9,
                                    color="#16a34a"
                                    if pd.notna(row["Achievement %"])
                                    and row["Achievement %"] >= 100
                                    else "#64748b",
                                ),
                            )

                        st.plotly_chart(
                            fig,
                            width="stretch",
                            config=PLOTLY_CONFIG,
                        )

                # ========================================================
                # PRODUCT + ITEM GROUP + QTY ANALYTICS
                # ========================================================
                st.markdown(
                    "##### 📦 Product, Item Group & QTY Analytics"
                )

                product_col, group_col = st.columns(2)

                with product_col:
                    st.markdown(
                        "###### 🏆 Product Quantity"
                    )

                    if (
                        "Product KW" in target_work.columns
                        and target_work["Product KW"].notna().any()
                    ):
                        prod_qty = (
                            target_work[
                                target_work[
                                    "Product KW"
                                ].notna()
                            ]
                            .groupby(
                                "Product KW",
                                as_index=False,
                            )
                            .agg(QTY=("QTY", "sum"))
                            .sort_values(
                                "QTY",
                                ascending=False,
                            )
                        )

                        order_map = {
                            f"{v:g} KW": i
                            for i, v in enumerate(
                                APPROVED_PRODUCT_KW
                            )
                        }
                        prod_qty["KW Order"] = (
                            prod_qty["Product KW"].map(
                                order_map
                            )
                        )
                        prod_qty["Label"] = (
                            prod_qty["QTY"].map(
                                lambda x: f"{x:,.0f}"
                            )
                        )
                        prod_chart = prod_qty.sort_values(
                            "QTY",
                            ascending=True,
                        )

                        fig = px.bar(
                            prod_chart,
                            x="QTY",
                            y="Product KW",
                            orientation="h",
                            color_discrete_sequence=[
                                company_theme("#f97316")
                            ],
                        )

                        max_prod_qty = (
                            float(prod_qty["QTY"].max())
                            if not prod_qty.empty
                            else 0
                        )

                        fig = clean_plotly_fig(
                            fig,
                            height=max(
                                390,
                                min(
                                    650,
                                    len(prod_qty) * 30 + 120,
                                ),
                            ),
                            is_horizontal=True,
                            max_val=max_prod_qty,
                        )
                        fig.update_traces(
                            text=prod_chart["Label"],
                            textposition="outside",
                            hovertemplate=(
                                "<b>%{y}</b>"
                                "<br>QTY: %{x:,.0f}"
                                "<extra></extra>"
                            ),
                        )

                        st.plotly_chart(
                            fig,
                            width="stretch",
                            config=PLOTLY_CONFIG,
                        )

                    elif "Item Name" in target_work.columns:
                        prod_qty = (
                            target_work.groupby(
                                "Item Name",
                                as_index=False,
                            )
                            .agg(QTY=("QTY", "sum"))
                            .sort_values(
                                "QTY",
                                ascending=False,
                            )
                            .head(12)
                        )
                        prod_qty["Clean Product"] = (
                            prod_qty["Item Name"].map(
                                lambda x: clean_label(x, 28)
                            )
                        )
                        prod_chart = prod_qty.sort_values(
                            "QTY",
                            ascending=True,
                        )

                        fig = px.bar(
                            prod_chart,
                            x="QTY",
                            y="Clean Product",
                            orientation="h",
                            color_discrete_sequence=[
                                company_theme("#f97316")
                            ],
                        )
                        fig = clean_plotly_fig(
                            fig,
                            height=470,
                            is_horizontal=True,
                            max_val=float(
                                prod_qty["QTY"].max()
                            ) if not prod_qty.empty else 0,
                        )
                        fig.update_traces(
                            text=prod_chart["QTY"].map(
                                lambda x: f"{x:,.0f}"
                            ),
                            textposition="outside",
                        )
                        st.plotly_chart(
                            fig,
                            width="stretch",
                            config=PLOTLY_CONFIG,
                        )

                with group_col:
                    st.markdown(
                        "###### 🧩 Item Group Quantity"
                    )

                    if "Item Group" in target_work.columns:
                        group_qty = (
                            target_work.groupby(
                                "Item Group",
                                as_index=False,
                            )
                            .agg(QTY=("QTY", "sum"))
                            .sort_values(
                                "QTY",
                                ascending=False,
                            )
                            .head(15)
                        )
                        group_qty["Clean Group"] = (
                            group_qty["Item Group"].map(
                                lambda x: clean_label(x, 28)
                            )
                        )
                        group_chart = group_qty.sort_values(
                            "QTY",
                            ascending=True,
                        )

                        fig = px.bar(
                            group_chart,
                            x="QTY",
                            y="Clean Group",
                            orientation="h",
                            color_discrete_sequence=[
                                "#0284c7"
                            ],
                        )
                        fig = clean_plotly_fig(
                            fig,
                            height=max(
                                390,
                                min(
                                    650,
                                    len(group_qty) * 30 + 120,
                                ),
                            ),
                            is_horizontal=True,
                            max_val=float(
                                group_qty["QTY"].max()
                            ) if not group_qty.empty else 0,
                        )
                        fig.update_traces(
                            text=group_chart["QTY"].map(
                                lambda x: f"{x:,.0f}"
                            ),
                            textposition="outside",
                            hovertemplate=(
                                "<b>%{y}</b>"
                                "<br>QTY: %{x:,.0f}"
                                "<extra></extra>"
                            ),
                        )
                        st.plotly_chart(
                            fig,
                            width="stretch",
                            config=PLOTLY_CONFIG,
                        )

                # ========================================================
                # ACHIEVEMENT %
                # ========================================================
                st.markdown(
                    "##### 📈 Employee Achievement %"
                )

                pct_df = perf[perf["Target"] > 0].copy()

                if pct_df.empty:
                    st.info(
                        "Enter non-zero employee targets to "
                        "activate Achievement % analysis."
                    )
                else:
                    pct_df["Clean Employee"] = (
                        pct_df["Employee"].map(
                            lambda x: clean_label(x, 27)
                        )
                    )
                    pct_df["Pct Label"] = (
                        pct_df["Achievement %"].map(
                            lambda x: f"{x:.1f}%"
                        )
                    )
                    pct_df = pct_df.sort_values(
                        "Achievement %",
                        ascending=True,
                    )

                    fig = px.bar(
                        pct_df,
                        x="Achievement %",
                        y="Clean Employee",
                        orientation="h",
                        color_discrete_sequence=[
                            "#16a34a"
                        ],
                    )

                    max_pct_raw = float(
                        pct_df["Achievement %"].max()
                    )
                    max_pct = max(
                        100.0,
                        min(max_pct_raw * 1.12, 250.0),
                    )

                    fig = clean_plotly_fig(
                        fig,
                        height=max(
                            400,
                            min(
                                740,
                                len(pct_df) * 32 + 140,
                            ),
                        ),
                        is_horizontal=True,
                        max_val=max_pct,
                    )

                    fig.update_traces(
                        text=pct_df["Pct Label"],
                        textposition="outside",
                        customdata=np.stack(
                            [
                                pct_df["Employee"],
                                pct_df["Target"],
                                pct_df["Achievement"],
                                pct_df["Gap"],
                            ],
                            axis=-1,
                        ),
                        hovertemplate=(
                            "<b>%{customdata[0]}</b>"
                            "<br>Target: ₹%{customdata[1]:,.2f}"
                            "<br>Achievement: ₹%{customdata[2]:,.2f}"
                            "<br>Remaining Gap: ₹%{customdata[3]:,.2f}"
                            "<br>Achievement %: %{x:.1f}%"
                            "<extra></extra>"
                        ),
                    )

                    fig.add_vline(
                        x=100,
                        line_dash="dash",
                        line_color="#16a34a",
                    )

                    st.plotly_chart(
                        fig,
                        width="stretch",
                        config=PLOTLY_CONFIG,
                    )

                # ========================================================
                # DETAIL TABLE
                # ========================================================
                with st.expander(
                    "📋 Employee Target Performance Detail",
                    expanded=True,
                ):
                    detail = perf[
                        [
                            "Employee",
                            "Target",
                            "Achievement",
                            "QTY",
                            "Invoices",
                            "Achievement %",
                            "Gap",
                            "Status",
                        ]
                    ].copy()

                    # Target-set employees first, then highest achievement.
                    detail["_Target Set"] = (
                        pd.to_numeric(
                            detail["Target"],
                            errors="coerce",
                        ).fillna(0) > 0
                    )

                    detail["_Achievement Sort"] = pd.to_numeric(
                        detail["Achievement"],
                        errors="coerce",
                    ).fillna(0)

                    detail = detail.sort_values(
                        ["_Target Set", "_Achievement Sort"],
                        ascending=[False, False],
                    ).drop(
                        columns=[
                            "_Target Set",
                            "_Achievement Sort",
                        ]
                    )

                    # Full exact values — no Cr/L abbreviation in the detail table.
                    detail["Target"] = detail[
                        "Target"
                    ].map(format_indian_currency_full)

                    detail["Achievement"] = detail[
                        "Achievement"
                    ].map(format_indian_currency_full)

                    detail["Gap"] = detail[
                        "Gap"
                    ].map(format_indian_currency_full)

                    detail["QTY"] = pd.to_numeric(
                        detail["QTY"],
                        errors="coerce",
                    ).fillna(0).map(
                        lambda x: format_indian_number_full(
                            x
                        )
                    )

                    detail["Invoices"] = pd.to_numeric(
                        detail["Invoices"],
                        errors="coerce",
                    ).fillna(0).astype(int)

                    detail["Achievement %"] = detail[
                        "Achievement %"
                    ].map(
                        lambda x: (
                            f"{x:.1f}%"
                            if pd.notna(x)
                            else "—"
                        )
                    )

                    st.caption(
                        "Target is calculated employee-wise from Year / Month / Custom Date rules for the selected Date Range. "
                        "Month and Custom Date Range targets override the Year target only for their own dates."
                    )

                    st.dataframe(
                        detail,
                        width="stretch",
                        hide_index=True,
                        height=min(
                            650,
                            max(
                                300,
                                len(detail) * 36 + 80,
                            ),
                        ),
                        column_config={
                            "Employee": st.column_config.TextColumn(
                                "Employee",
                                width="large",
                            ),
                            "Target": st.column_config.TextColumn(
                                "Target",
                                width="medium",
                            ),
                            "Achievement": st.column_config.TextColumn(
                                "Achievement",
                                width="medium",
                            ),
                            "QTY": st.column_config.TextColumn(
                                "QTY",
                                width="small",
                            ),
                            "Invoices": st.column_config.NumberColumn(
                                "Invoices",
                                width="small",
                                format="%d",
                            ),
                            "Achievement %": st.column_config.TextColumn(
                                "Achievement %",
                                width="small",
                            ),
                            "Gap": st.column_config.TextColumn(
                                "Gap",
                                width="medium",
                            ),
                            "Status": st.column_config.TextColumn(
                                "Status",
                                width="medium",
                            ),
                        },
                    )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# 4. SALES RETURNS & NET SALES RECONCILIATION
# ============================================================
if selected_page == "Sales Returns & Net":
    st.markdown("""
    <div class="pv-chart-card">
        <div class="pv-card-header">
            <div class="pv-card-title">↩️ Sales Return Register & Net Sales Reconciliation</div>
        </div>
    """, unsafe_allow_html=True)

    if sales_df.empty:
        st.info("Upload Normal Sales Excel first. Return Register can then be reconciled against the same Sales filters.")
    else:
        n = net_metrics(filtered_sales, filtered_returns)
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        cards = [
            (c1, "₹", "GROSS SALES", format_indian_currency(n["gross_amount"]), "Normal Sales", format_indian_currency_full(n["gross_amount"])),
            (c2, "↩️", "RETURN", format_indian_currency(n["return_amount"]), "Return Register", format_indian_currency_full(n["return_amount"])),
            (c3, "✅", "NET SALES", format_indian_currency(n["net_amount"]), "Gross − Return", format_indian_currency_full(n["net_amount"])),
            (c4, "📦", "GROSS QTY", f'{n["gross_qty"]:,.0f}', "Normal Sales", format_indian_number_full(n["gross_qty"])),
            (c5, "↩️", "RETURN QTY", f'{n["return_qty"]:,.0f}', "Returned units", format_indian_number_full(n["return_qty"])),
            (c6, "⚡", "NET QTY", f'{n["net_qty"]:,.0f}', "Gross − Return", format_indian_number_full(n["net_qty"])),
        ]
        for col, icon, label, value, badge, full_value in cards:
            with col:
                st.markdown(
                    kpi_card_html(icon, label, value, badge, full_value),
                    unsafe_allow_html=True
                )

        if not return_files:
            st.info("No Sales Return Register uploaded. Net Sales currently equals Gross Sales.")
        elif filtered_returns.empty:
            st.warning(
                "Return Register is loaded, but the current Return Date / Sales business filters leave 0 return rows. "
                "Set Return Date Scope to 'All / Return Date Range' and use the full Return Register date range."
            )
        else:
            st.markdown(
                f'<div class="pv-net-banner"><span class="pv-return-chip">RETURN REGISTER ACTIVE</span> '
                f'{len(filtered_returns):,} cleaned return line(s) are deducted from the current Sales selection.</div>',
                unsafe_allow_html=True
            )
            if return_cleaning:
                validation = return_cleaning.get("validation_status", "")
                report_qty = return_cleaning.get("report_total_qty", 0)
                report_amount = return_cleaning.get("report_total_amount", 0)
                if report_qty or report_amount:
                    msg = (
                        f"Return source validation: {validation} • Report total QTY {report_qty:,.0f} • "
                        f"Clean detail QTY {return_cleaning.get('return_qty',0):,.0f} • "
                        f"Report amount {format_indian_currency(report_amount)} • "
                        f"Clean detail amount {format_indian_currency(return_cleaning.get('return_amount',0))} • "
                        f"Exact duplicate lines reconciled {return_cleaning.get('duplicates_reconciled',0):,}."
                    )
                    (st.success if validation == "PASS" else st.warning)(msg)

        customer_net = net_by_dimension(filtered_sales, filtered_returns, "Ledger Name")
        product_net = net_by_dimension(filtered_sales, filtered_returns, "Item Name")
        group_net = net_by_dimension(filtered_sales, filtered_returns, "Item Group")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 🏢 Net Sales by Customer")
            top = customer_net.sort_values("Net Sales", ascending=False).head(12).copy()
            top["CleanLabel"] = top["Ledger Name"].map(lambda x: clean_label(x, 28))
            top["NetLabel"] = top["Net Sales"].map(format_indian_currency)
            fig = px.bar(top.sort_values("Net Sales"), x="Net Sales", y="Ledger Name", orientation="h", color_discrete_sequence=["#16a34a"])
            max_v = float(top["Net Sales"].clip(lower=0).max()) if not top.empty else 0
            fig = clean_plotly_fig(fig, height=450, is_horizontal=True, max_val=max_v)
            fig.update_traces(text=top.sort_values("Net Sales")["NetLabel"], textposition="outside", textfont=dict(size=10, weight=700))
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        with c2:
            st.markdown("##### 📦 Net Quantity by Standard KW Product")
            sales_kw = filtered_sales[filtered_sales["Product KW"].notna()].groupby("Product KW", as_index=False).agg(**{"Gross QTY": ("QTY", "sum")})
            if not filtered_returns.empty and "Product KW" in filtered_returns.columns:
                ret_kw = filtered_returns[filtered_returns["Product KW"].notna()].groupby("Product KW", as_index=False).agg(**{"Return QTY": ("QTY", "sum")})
                top = sales_kw.merge(ret_kw, on="Product KW", how="left")
                top["Return QTY"] = top["Return QTY"].fillna(0)
            else:
                top = sales_kw.copy()
                top["Return QTY"] = 0.0
            top["Net QTY"] = top["Gross QTY"] - top["Return QTY"]
            top = top.sort_values("Net QTY", ascending=False)
            chart_top = top.sort_values("Net QTY")
            fig = px.bar(chart_top, x="Net QTY", y="Product KW", orientation="h", color_discrete_sequence=["#0284c7"])
            max_v = float(top["Net QTY"].clip(lower=0).max()) if not top.empty else 0
            fig = clean_plotly_fig(fig, height=max(450, len(top) * 30 + 120), is_horizontal=True, max_val=max_v)
            fig.update_traces(text=chart_top["Net QTY"], texttemplate="%{text:,.0f}", textposition="outside", hovertemplate="<b>%{y}</b><br>Net QTY: %{x:,.0f}<extra></extra>")
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        st.markdown("##### 🧩 Item Group Net Reconciliation")
        group_display = group_net.sort_values("Net Sales", ascending=False).copy()
        for col in ["Gross Sales", "Return Amount", "Net Sales"]:
            group_display[col] = group_display[col].map(format_indian_currency)
        for col in ["Gross QTY", "Return QTY", "Net QTY"]:
            group_display[col] = group_display[col].map(lambda x: f"{x:,.0f}")
        st.dataframe(group_display, width="stretch", hide_index=True, height=330)

        if not filtered_returns.empty:
            unallocated = filtered_returns[filtered_returns["Return Representative"] == "Unallocated Return"]
            unallocated_amount = float(unallocated["Amount"].sum()) if not unallocated.empty else 0.0
            unallocated_qty = float(unallocated["QTY"].sum()) if not unallocated.empty else 0.0
            if unallocated_amount > 0 or unallocated_qty > 0:
                st.warning(
                    f"Person-wise return audit: {format_indian_currency(unallocated_amount)} / {unallocated_qty:,.0f} QTY "
                    "could not be assigned to a sales representative with certainty. Global Net Sales still deducts it correctly; "
                    "employee Net Achievement deducts only accurately allocated returns."
                )

            st.markdown("##### 📋 Clean Sales Return Register")
            show_cols = [c for c in [
                "Sales Return Date", "Order No#", "Ledger Name", "Item Group", "Item Name",
                "QTY", "Rate", "Amount", "Return Representative"
            ] if c in filtered_returns.columns]
            st.dataframe(filtered_returns[show_cols], width="stretch", hide_index=True, height=400)

    st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# 4. FIELD DASHBOARD — EMPLOYEE KPI / KRA
# ============================================================
if selected_page == "Field KPI / KRA":
    st.markdown("""
    <div class="pv-chart-card">
        <div class="pv-card-header">
            <div class="pv-card-title">📈 Field Dashboard — Employee KPI / KRA Intelligence</div>
        </div>
    """, unsafe_allow_html=True)

    if field_df.empty:
        st.info(
            "Upload one or more Field - Visit Reports. Supported formats: "
            "(1) Sr.No | Emp_Name | Details | Date columns | Total, or "
            "(2) Emp_Name | Visit | KM | Lead | Follow-Ups | Quotation | PI | Sales-Order | Wrong KM Claim."
        )
    else:
        st.markdown(
            f'<div class="pv-net-banner"><span class="pv-field-chip">FIELD VISIT SOURCE</span> '
            f'<b>{field_info.get("roster_employees", field_info.get("employees",0))}</b> roster employee(s) • '
            f'<b>{field_info.get("active_employees",0)}</b> with recorded activity • {field_info.get("format","")} • '
            f'{field_info.get("dates",0)} reporting date(s). KPI/KRA data is independent of Sales/Returns.</div>',
            unsafe_allow_html=True
        )

        if field_df["Date"].notna().any():
            _fd_min = field_df["Date"].min()
            _fd_max = field_df["Date"].max()
            st.caption(
                f"Detected field reporting period: {_fd_min.strftime('%d/%m/%Y')} to {_fd_max.strftime('%d/%m/%Y')} "
                f"• {field_df['Date'].nunique()} date(s)."
            )

        field_filtered = field_df.copy()
        filter_col1, filter_col2 = st.columns([1.15, 0.85])
        with filter_col1:
            employees = sorted(
                field_info.get("employee_roster", field_filtered["Employee"].dropna().astype(str).unique().tolist()),
                key=str.casefold,
            )
            emp_sel = st.multiselect("Employee Filter", employees, key="field_emp_filter")
            if emp_sel:
                field_filtered = field_filtered[field_filtered["Employee"].isin(emp_sel)]
        with filter_col2:
            if field_filtered["Date"].notna().any():
                min_d = field_filtered["Date"].min().date()
                max_d = field_filtered["Date"].max().date()
                fr = st.date_input("Field Date Range", value=(min_d, max_d), min_value=min_d, max_value=max_d, key="field_date_filter_v26", format="DD/MM/YYYY")
                if isinstance(fr, (tuple, list)) and len(fr) == 2:
                    field_filtered = field_filtered[
                        (field_filtered["Date"] >= pd.Timestamp(fr[0]))
                        & (field_filtered["Date"] < pd.Timestamp(fr[1]) + pd.Timedelta(days=1))
                    ]

        emp = field_employee_summary(field_filtered)
        totals = {m: float(field_filtered[m].sum()) for m in FIELD_METRICS}
        active_employees = int(
            emp.loc[emp[FIELD_METRICS].sum(axis=1) > 0, "Employee"].nunique()
        ) if not emp.empty else 0
        roster_employees = len(emp_sel) if emp_sel else int(field_info.get("roster_employees", len(employees)))

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        for col, icon, label, value, note, full_value in [
            (k1, "👥", "EMPLOYEES", f"{roster_employees:,}", f"Roster • {active_employees} active in selection", format_indian_number_full(roster_employees)),
            (k2, "🚗", "VISITS", f'{totals["Visit"]:,.0f}', "Customer/site visits", format_indian_number_full(totals["Visit"])),
            (k3, "🎯", "LEADS", f'{totals["Lead"]:,.0f}', "Lead generation", format_indian_number_full(totals["Lead"])),
            (k4, "🔁", "FOLLOW-UPS", f'{totals["Follow-Ups"]:,.0f}', "Follow-up activity", format_indian_number_full(totals["Follow-Ups"])),
            (k5, "📄", "PI", f'{totals["PI"]:,.0f}', "Proforma invoices", format_indian_number_full(totals["PI"])),
            (k6, "✅", "SALES ORDERS", f'{totals["Sales-Order"]:,.0f}', "Orders generated", format_indian_number_full(totals["Sales-Order"])),
        ]:
            with col:
                st.markdown(
                    kpi_card_html(icon, label, value, note, full_value),
                    unsafe_allow_html=True
                )

        sub1, sub2, sub3, sub4 = st.columns(4)
        conversion_values = [
            (sub1, "Lead / Visit", (totals["Lead"] / totals["Visit"] * 100) if totals["Visit"] else 0),
            (sub2, "PI / Lead", (totals["PI"] / totals["Lead"] * 100) if totals["Lead"] else 0),
            (sub3, "Order / PI", (totals["Sales-Order"] / totals["PI"] * 100) if totals["PI"] else 0),
            (sub4, "Order / Visit", (totals["Sales-Order"] / totals["Visit"] * 100) if totals["Visit"] else 0),
        ]
        for col, label, value in conversion_values:
            with col:
                st.metric(label, f"{value:.1f}%")

        coverage = field_info.get("coverage")
        if isinstance(coverage, pd.DataFrame) and not coverage.empty:
            incomplete = coverage[coverage["Missing Metric Rows"] > 0]
            if not incomplete.empty:
                st.markdown(
                    f'<div class="pv-net-banner"><b>KPI/KRA source audit:</b> {incomplete["Employee"].nunique()} employee(s) have one or more missing KPI metric rows in an uploaded file. '
                    'All roster employees remain visible; missing metrics are treated as 0 and can be reviewed below.</div>',
                    unsafe_allow_html=True
                )
                with st.expander("View KPI/KRA metric-row coverage", expanded=False):
                    st.dataframe(coverage, width="stretch", hide_index=True)

        st.markdown("##### 🎯 KPI/KRA Target Settings")
        with st.expander("Edit KPI/KRA targets", expanded=False):
            tcols = st.columns(3)
            targets = {}
            for i, (metric, default) in enumerate(FIELD_DEFAULT_TARGETS.items()):
                with tcols[i % 3]:
                    targets[metric] = st.number_input(
                        f"{metric} Target / Employee", min_value=0.0, value=float(default), step=1.0,
                        key=f"field_target_{metric}"
                    )
        emp = add_kra_scores(emp, targets)

        c1, c2 = st.columns(2)
        with c1:
            metric_choice = st.selectbox(
                "Employee KPI to compare", ["Visit", "Lead", "Follow-Ups", "PI", "Sales-Order"],
                key="field_metric_compare"
            )
            ranking = emp.sort_values(metric_choice, ascending=False).head(27).sort_values(metric_choice).copy()
            ranking["Employee Display"] = ranking["Employee"].map(lambda x: clean_label(x, 30))
            fig = px.bar(ranking, x=metric_choice, y="Employee Display", orientation="h", text=metric_choice, hover_name="Employee", color_discrete_sequence=[company_theme("#f97316")])
            max_v = float(ranking[metric_choice].max()) if not ranking.empty else 0
            fig = clean_plotly_fig(fig, height=max(460, len(ranking) * 31 + 120), is_horizontal=True, max_val=max_v)
            fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textfont=dict(size=10, weight=700))
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        with c2:
            kra = emp.sort_values("KRA Score", ascending=False).head(27).sort_values("KRA Score").copy()
            kra["Employee Display"] = kra["Employee"].map(lambda x: clean_label(x, 30))
            fig = px.bar(kra, x="KRA Score", y="Employee Display", orientation="h", text="KRA Score", hover_name="Employee", color_discrete_sequence=["#16a34a"])
            fig = clean_plotly_fig(fig, height=max(460, len(kra) * 31 + 120), is_horizontal=True, max_val=100)
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside", textfont=dict(size=10, weight=700))
            fig.add_vline(x=75, line_dash="dash", line_color=company_theme("#f97316"))
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        funnel_df = pd.DataFrame({
            "Stage": ["Visit", "Lead", "Follow-Ups", "Quotation", "PI", "Sales-Order"],
            "Count": [totals["Visit"], totals["Lead"], totals["Follow-Ups"], totals["Quotation"], totals["PI"], totals["Sales-Order"]]
        })
        c1, c2 = st.columns(2)
        with c1:
            fig = px.funnel(funnel_df, x="Count", y="Stage", title="Field Sales Activity Funnel", color_discrete_sequence=[company_theme("#f97316")])
            fig = clean_plotly_fig(fig, height=420)
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
        with c2:
            if field_filtered["Date"].notna().any():
                daily_metric = st.selectbox("Daily KPI Trend", ["Visit", "Lead", "Follow-Ups", "PI", "Sales-Order"], key="field_daily_metric")
                daily = field_filtered.groupby("Date", as_index=False)[daily_metric].sum().sort_values("Date")
                daily["Data Label"] = daily[daily_metric].map(lambda v: "" if float(v) == 0 else f"{float(v):,.0f}")
                fig = px.line(daily, x="Date", y=daily_metric, markers=True, text="Data Label", title=f"Daily {daily_metric}", color_discrete_sequence=["#0284c7"])
                fig.update_traces(textposition="top center", textfont=dict(size=10, color="#0f172a"), cliponaxis=False)
                fig.update_xaxes(tickformat="%d/%m", dtick="D1" if len(daily) <= 16 else None)
                max_daily = float(daily[daily_metric].max()) if not daily.empty else 0
                if max_daily > 0:
                    fig.update_yaxes(range=[0, max_daily * 1.20])
                fig = clean_plotly_fig(fig, height=420)
                st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
            else:
                st.info("This Field Visit source has no date-level fields. Employee totals and KPI/KRA analytics are still available.")

        if "Designation" in emp.columns and emp["Designation"].notna().any():
            st.markdown("##### 🧑‍💼 Designation Analytics")
            designation = emp.groupby("Designation", as_index=False).agg(
                Employees=("Employee", "nunique"), Visits=("Visit", "sum"), Leads=("Lead", "sum"),
                PI=("PI", "sum"), Orders=("Sales-Order", "sum")
            ).sort_values("Orders", ascending=False)
            st.dataframe(designation, width="stretch", hide_index=True)
        else:
            st.caption("Future-ready: when a Designation column is added to the Field Visit source, designation-wise analytics will activate automatically.")

        st.markdown("##### 📋 Employee KPI / KRA Scorecard")
        display_cols = [
            "Employee", "Visit", "Lead", "Follow-Ups", "Quotation", "PI", "Sales-Order",
            "KM", "Wrong KM Claim", "Valid KM", "Lead / Visit %", "Order / Visit %", "KRA Score", "KRA Performance"
        ]
        if "Designation" in emp.columns:
            display_cols.insert(1, "Designation")
        st.dataframe(emp[display_cols].sort_values(["KRA Score", "Sales-Order"], ascending=False), width="stretch", hide_index=True, height=520)

    st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# 6. EMPLOYEE QUOTAS & TARGET PERFORMANCE (ZERO-OVERLAP FIX)
# ============================================================
if selected_page == "Quotas & Targets":
    st.markdown('''
    <div class="pv-chart-card">
        <div class="pv-card-header">
            <div class="pv-card-title">🎯 Quotas & Representative Achievement Tracking</div>
        </div>
    ''', unsafe_allow_html=True)
    if filtered_sales.empty:
        st.info("Upload Sales Excel data to enter employee targets and track quota achievement.")
    else:
        rep_col = sales_representative_column(filtered_sales)
        emp_names = sorted([x for x in filtered_sales[rep_col].dropna().astype(str).str.strip().unique() if x and x != "-"])

        st.markdown("##### 🎯 Representative Monthly Target Input Table")
        target_df = target_editor(emp_names)

        # Net employee Achievement = Gross employee sales − accurately attributable returns.
        ach, _ = employee_net_performance(filtered_sales, filtered_returns)
        if ach.empty:
            ach = pd.DataFrame(columns=["Employee", "Achievement", "QTY", "Invoices", "Gross_Achievement", "Return Amount", "Return QTY"])

        if not target_df.empty:
            target_df["Target"] = pd.to_numeric(target_df["Target"], errors="coerce").fillna(0)
            ach = ach.merge(target_df, on="Employee", how="left")
        else:
            ach["Target"] = 0

        ach["Target"] = pd.to_numeric(ach["Target"], errors="coerce").fillna(0)
        ach["Achievement %"] = np.where(ach["Target"] > 0, (ach["Achievement"] / ach["Target"]) * 100, np.nan)
        ach["Gap"] = ach["Achievement"] - ach["Target"]
        ach["CleanEmp"] = ach["Employee"].map(lambda x: clean_label(x, 26))
        ach["FormattedAch"] = ach["Achievement"].map(format_indian_currency)
        # FIX ZERO TEXT COLLISION: Only show Target text if Target > 0
        ach["FormattedTgt"] = ach["Target"].map(lambda x: format_indian_currency(x) if x > 0 else "")
        ach["FormattedPct"] = ach["Achievement %"].map(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown('##### 📊 Quota Achievement % by Representative')
            plot_df = ach[ach["Target"] > 0].sort_values("Achievement %")
            if not plot_df.empty:
                max_pct = float(plot_df["Achievement %"].max()) if not plot_df.empty else 100
                fig = px.bar(plot_df, x="Achievement %", y="CleanEmp", orientation="h", color_discrete_sequence=[company_theme("#f97316")])
                fig = clean_plotly_fig(fig, height=520, is_horizontal=True, max_val=max_pct)
                fig.update_traces(
                    text=plot_df["FormattedPct"],
                    textposition="outside",
                    textfont=dict(size=10.5, color="#0f172a", weight=700)
                )
                fig.add_vline(x=100, line_dash="dash", line_color="#16a34a")
                st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
            else:
                st.info("💡 Enter non-zero targets in the table above to visualize Quota Achievement %.")

        with c2:
            st.markdown('##### 🎯 Target vs Actual Sales Revenue (Spaced Layout)')
            plot_comp = ach.sort_values("Achievement", ascending=True)
            dynamic_height = max(500, len(plot_comp) * 32)
            fig = go.Figure()
            fig.add_trace(go.Bar(
                name="Actual Achievement",
                y=plot_comp["CleanEmp"],
                x=plot_comp["Achievement"],
                orientation="h",
                marker_color=company_theme("#f97316"),
                text=plot_comp["FormattedAch"],
                textposition="outside"
            ))
            fig.add_trace(go.Bar(
                name="Target Quota",
                y=plot_comp["CleanEmp"],
                x=plot_comp["Target"],
                orientation="h",
                marker_color="#0284c7",
                text=plot_comp["FormattedTgt"],
                textposition="outside"
            ))
            max_ach_val = float(plot_comp[["Achievement", "Target"]].max().max()) if not plot_comp.empty else 0
            fig = clean_plotly_fig(fig, height=dynamic_height, is_horizontal=True, max_val=max_ach_val)
            fig.update_layout(barmode="group", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

        st.markdown("##### 📋 Complete Performance Audit Matrix")
        display_cols_perf = [c for c in ["Employee", "Gross_Achievement", "Return Amount", "Target", "Achievement", "Gap", "Achievement %"] if c in ach.columns]
        display_df = ach[display_cols_perf].copy()
        display_df["Target"] = display_df["Target"].map(format_indian_currency)
        if "Gross_Achievement" in display_df.columns:
            display_df["Gross_Achievement"] = display_df["Gross_Achievement"].map(format_indian_currency)
        if "Return Amount" in display_df.columns:
            display_df["Return Amount"] = display_df["Return Amount"].map(format_indian_currency)
        display_df["Achievement"] = display_df["Achievement"].map(format_indian_currency)
        display_df["Gap"] = display_df["Gap"].map(format_indian_currency)
        display_df["Achievement %"] = display_df["Achievement %"].map(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")
        st.dataframe(display_df, width="stretch", hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
# 7. FULL SERVICE OPERATIONS & TICKET ANALYTICS (4 COMPLETE CHARTS)
# ============================================================
if selected_page == "Service Operations":
    st.markdown('''
    <div class="pv-chart-card">
        <div class="pv-card-header">
            <div class="pv-card-title">🛠️ Service Maintenance Operations & Ticket Intelligence</div>
        </div>
    ''', unsafe_allow_html=True)
    if service_df.empty:
        st.info("Upload Service Excel files from the left sidebar to activate Maintenance Intelligence.")
    else:
        sf = service_df[service_df["Ticket Valid"]].copy()
        if sf.empty:
            st.warning("Service files loaded, but no valid ticket numbers were detected.")
        else:
            total_tickets = len(sf)
            closed_tickets = int((sf["Status"] == "Closed").sum())
            open_tickets = int((sf["Status"] == "Open").sum())
            service_products = service_product_counts(sf)
            top_product = service_products.iloc[0]["Product"] if not service_products.empty else "N/A"

            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(kpi_card_html("🛠️", "TOTAL TICKETS", f"{total_tickets:,}", "100% Logged", format_indian_number_full(total_tickets)), unsafe_allow_html=True)
            with k2:
                st.markdown(kpi_card_html("✅", "CLOSED TICKETS", f"{closed_tickets:,}", f"{closed_tickets/total_tickets*100:.1f}% Resolved", format_indian_number_full(closed_tickets)), unsafe_allow_html=True)
            with k3:
                st.markdown(kpi_card_html("⏳", "OPEN TICKETS", f"{open_tickets:,}", f"{open_tickets/total_tickets*100:.1f}% In Progress", format_indian_number_full(open_tickets)), unsafe_allow_html=True)
            with k4:
                st.markdown(kpi_card_html("📦", "TOP SERVICE PRODUCT", clean_text(top_product), "High Maintenance", clean_text(top_product)), unsafe_allow_html=True)

            st.write("")

            # 4 COMPLETE SERVICE ANALYTICS CHARTS (2x2 GRID)
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("##### 🍩 1. Ticket Resolution Status Breakdown")
                status_counts = sf["Status"].value_counts().reset_index()
                status_counts.columns = ["Status", "Count"]
                fig = px.pie(status_counts, values="Count", names="Status", hole=0.48,
                             color_discrete_map={"Open": company_theme("#f97316"), "Closed": "#16a34a", "Unknown": "#94a3b8"})
                fig = clean_plotly_fig(fig, height=350)
                fig.update_traces(textinfo="percent+value", textposition="inside", insidetextfont=dict(color="#ffffff", size=11, weight=700))
                st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

            with c2:
                st.markdown("##### 📈 2. Monthly Ticket Creation / Closure Trend")
                created_col = next((c for c in ["Date_Parsed", "Created Date_Parsed"] if c in sf.columns and sf[c].notna().any()), None)
                closed_col = "Closed Date_Parsed" if "Closed Date_Parsed" in sf.columns and (sf["Status"].eq("Closed") & sf["Closed Date_Parsed"].notna()).any() else None
                pieces = []
                if created_col:
                    x = sf[sf[created_col].notna()].copy()
                    x["Month"] = x[created_col].dt.to_period("M").astype(str)
                    x = x.groupby("Month", as_index=False).size().rename(columns={"size": "Query Created"})
                    pieces.append(x)
                if closed_col:
                    x = sf[sf["Status"].eq("Closed") & sf[closed_col].notna()].copy()
                    x["Month"] = x[closed_col].dt.to_period("M").astype(str)
                    x = x.groupby("Month", as_index=False).size().rename(columns={"size": "Query Closed"})
                    pieces.append(x)
                if pieces:
                    t_trend = pieces[0]
                    for x in pieces[1:]:
                        t_trend = t_trend.merge(x, on="Month", how="outer")
                    t_trend = t_trend.fillna(0).sort_values("Month")
                    t_trend["Month Label"] = pd.to_datetime(t_trend["Month"]).dt.strftime("%b %Y")
                    value_cols = [c for c in ["Query Created", "Query Closed"] if c in t_trend.columns]
                    fig = px.line(t_trend, x="Month Label", y=value_cols, markers=True)
                    fig.update_traces(mode="lines+markers+text", texttemplate="%{y:.0f}", textposition="top center")
                    fig = clean_plotly_fig(fig, height=350)
                    st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
                else:
                    st.info("Monthly trends need a creation date or a closed ticket with a closure date. This report has neither; all tickets are still included in the other charts.")

            with st.container():
                st.markdown("##### 📦 3. Top 25 Products Requiring Maintenance")
                st.caption("Listed products grouped by KW, phase and MPPT; naming variants are combined. Battery Pack is shown separately.")
                p_counts = service_products.head(25).copy()
                fig = px.bar(p_counts.sort_values("Count"), x="Count", y="Product", orientation="h", color_discrete_sequence=[company_theme("#f97316")])
                max_p = float(p_counts["Count"].max()) if not p_counts.empty else 0
                fig = clean_plotly_fig(fig, height=max(360, len(p_counts) * 34 + 120), is_horizontal=True, max_val=max_p)
                fig.update_yaxes(type="category", dtick=1, automargin=True)
                fig.update_traces(text=p_counts.sort_values("Count")["Count"], textposition="outside", textfont=dict(size=10.5, weight=700))
                st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

            with st.container():
                st.markdown("##### 🛠️ 4. Query Master / Issue Category Distribution")
                qm_counts = sf["Query Master"].dropna().value_counts().reset_index()
                qm_counts.columns = ["Query Master", "Count"]
                fig = px.bar(qm_counts.sort_values("Count"), x="Count", y="Query Master", orientation="h", color_discrete_sequence=[company_theme("#fb923c")])
                max_qm = float(qm_counts["Count"].max()) if not qm_counts.empty else 0
                fig = clean_plotly_fig(fig, height=max(360, len(qm_counts) * 34 + 120), is_horizontal=True, max_val=max_qm)
                fig.update_yaxes(type="category", dtick=1, automargin=True)
                fig.update_traces(text=qm_counts.sort_values("Count")["Count"], textposition="outside", textfont=dict(size=10.5, weight=700))
                st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

            st.markdown("##### 👥 Person-wise Query Created & Query Closed")
            with st.container():
                created_people = sf["Created By"].dropna().map(clean_text)
                created_people = created_people[created_people.ne("")].value_counts().reset_index()
                created_people.columns = ["Person", "Query Created"]
                if not created_people.empty:
                    fig = px.bar(created_people.sort_values("Query Created"), x="Query Created", y="Person", orientation="h", color_discrete_sequence=["#0284c7"])
                    fig = clean_plotly_fig(fig, height=max(360, len(created_people) * 34 + 120), is_horizontal=True, max_val=float(created_people["Query Created"].max()))
                    fig.update_yaxes(type="category", dtick=1, automargin=True)
                    fig.update_traces(text=created_people.sort_values("Query Created")["Query Created"], textposition="outside")
                    st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
                else:
                    st.info("Created By data is not available.")
            with st.container():
                closed_people = sf.loc[sf["Status"].eq("Closed"), "Closed By"].dropna().map(clean_text)
                closed_people = closed_people[closed_people.ne("")].value_counts().reset_index()
                closed_people.columns = ["Person", "Query Closed"]
                if not closed_people.empty:
                    fig = px.bar(closed_people.sort_values("Query Closed"), x="Query Closed", y="Person", orientation="h", color_discrete_sequence=["#16a34a"])
                    fig = clean_plotly_fig(fig, height=max(360, len(closed_people) * 34 + 120), is_horizontal=True, max_val=float(closed_people["Query Closed"].max()))
                    fig.update_yaxes(type="category", dtick=1, automargin=True)
                    fig.update_traces(text=closed_people.sort_values("Query Closed")["Query Closed"], textposition="outside")
                    st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
                else:
                    st.info("Closed By data is not available.")

            st.markdown("##### 📋 Maintenance Service Ticket Log Data")
            service_cols_display = [c for c in sf.columns if not c.endswith("_Parsed") and c not in ["Source File", "Ticket Valid", "MTCE Status Raw", "ServiceMonth"]]
            st.dataframe(sf[service_cols_display], width="stretch", height=400, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
# 5. DATA QUALITY & EXPORT
# ============================================================
if selected_page == "Quality & Export":
    st.markdown('''
    <div class="pv-chart-card">
        <div class="pv-card-header">
            <div class="pv-card-title">🧹 Data Cleaning Audit Log & Excel Export Hub</div>
        </div>
    ''', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if sales_cleaning:
            st.dataframe(pd.DataFrame(sales_cleaning["files"]), width="stretch", hide_index=True)
            audit_items = [
                ("Original Source Rows", sales_cleaning.get("original", 0)),
                ("Top Metadata Rows Removed", sales_cleaning.get("metadata_rows_removed", 0)),
                ("Detail Shift-Up Applied", sales_cleaning.get("detail_shift_applied", 0)),
                ("Blank Item Group Rows Removed", sales_cleaning.get("item_group_blank_removed", 0)),
                ("Company Representative Rows Assigned", sales_cleaning.get("pv_blink_assigned", 0)),
                ("Amount = 0 Rows Removed", sales_cleaning.get("amount_zero_removed", 0)),
                ("Invalid Date Rows Removed", sales_cleaning.get("invalid_dates_removed", 0)),
                ("Missing Core Field Rows Removed", sales_cleaning.get("core_removed", 0)),
                ("Duplicate Rows Removed", sales_cleaning.get("duplicates", 0)),
                ("Final Cleaned Rows", sales_cleaning.get("final", 0)),
                ("Detected Header Row Index", sales_cleaning.get("header_row", 0)),
                ("Header Matching Score", sales_cleaning.get("header_score", 0)),
            ]
            sales_audit = pd.DataFrame(audit_items, columns=["Quality Metric", "Count"])
            sales_audit["Count"] = sales_audit["Count"].astype(str)
            st.dataframe(sales_audit, width="stretch", hide_index=True)

    with c2:
        if not service_detail.empty:
            st.dataframe(service_detail, width="stretch", hide_index=True)

    if return_cleaning:
        st.markdown("##### ↩️ Sales Return Quality Audit")
        st.dataframe(pd.DataFrame(return_cleaning["files"]), width="stretch", hide_index=True)
        return_audit = pd.DataFrame([
            ("Original Source Rows", return_cleaning.get("original", 0)),
            ("Detected Header Row", return_cleaning.get("header_row", 0)),
            ("Zero/Blank Rate or Zero-value Rows Removed", return_cleaning.get("rate_removed", 0)),
            ("Invalid Date Rows Removed", return_cleaning.get("invalid_dates_removed", 0)),
            ("Missing Core Rows Removed", return_cleaning.get("core_removed", 0)),
            ("Duplicate Rows Removed", return_cleaning.get("duplicates", 0)),
            ("Duplicate-looking Rows Flagged (Retained)", return_cleaning.get("duplicate_rows_flagged", 0)),
            ("Final Return Lines", return_cleaning.get("final", 0)),
            ("Clean Return QTY", return_cleaning.get("return_qty", 0)),
            ("Clean Return Amount", return_cleaning.get("return_amount", 0)),
            ("Report Summary QTY", return_cleaning.get("report_total_qty", 0)),
            ("Report Summary Amount", return_cleaning.get("report_total_amount", 0)),
            ("QTY Reconciliation Difference", return_cleaning.get("qty_check_diff", 0)),
            ("Amount Reconciliation Difference", return_cleaning.get("amount_check_diff", 0)),
            ("Return Validation", return_cleaning.get("validation_status", "")),
        ], columns=["Return Quality Metric", "Value"])
        return_audit["Value"] = return_audit["Value"].astype(str)
        st.dataframe(return_audit, width="stretch", hide_index=True)

    if field_info:
        st.markdown("##### 📈 Dashboard KPI/KRA Source Audit")
        st.dataframe(field_info["files"], width="stretch", hide_index=True)
        field_audit = pd.DataFrame([
            ("Dashboard Sheet", field_info.get("sheet", "")),
            ("Detected Format", field_info.get("format", "")),
            ("Detected Header Row", field_info.get("header_row", 0)),
            ("Clean Activity Rows", field_info.get("rows", 0)),
            ("Employees", field_info.get("employees", 0)),
            ("Reporting Dates", field_info.get("dates", 0)),
        ], columns=["Dashboard Quality Metric", "Value"])
        field_audit["Value"] = field_audit["Value"].astype(str)
        st.dataframe(field_audit, width="stretch", hide_index=True)

    st.write("")
    if not sales_df.empty:
        try:
            excel_bytes = make_excel_bytes(sales_df, filtered_sales, sales_cleaning)
            st.download_button(
                label="📥 Download Formatted Sales Excel Workbook (Openpyxl)",
                data=excel_bytes,
                file_name=f"{company_export}_EKhata_Sales_Analytics.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )
        except Exception as e:
            st.error(f"Error generating Excel download: {e}")

    d1, d2 = st.columns(2)
    with d1:
        if not return_df.empty:
            return_export = filtered_returns.drop(columns=["Return Year", "Return Month", "Return Year-Month"], errors="ignore")
            st.download_button(
                "↩️ Download Clean Return Register",
                data=simple_excel_bytes({"Return_Cleaned": return_export}),
                file_name=f"{company_export}_Sales_Return_Cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch", key="download_return_excel"
            )
    with d2:
        if not field_df.empty:
            field_summary_export = field_employee_summary(field_df)
            st.download_button(
                "📈 Download KPI-KRA Clean Data",
                data=simple_excel_bytes({"Field_Daily": field_df, "Employee_Summary": field_summary_export}),
                file_name=f"{company_export}_Field_KPI_KRA_Cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch", key="download_field_excel"
            )
    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown(
    f"""
<div class="pv-footer">
    <div>© 2025 {company_title}. All rights reserved.</div>
    <div>Sales • Returns • KPI/KRA • Employee Performance • Service Intelligence</div>
</div>
""",
    unsafe_allow_html=True,
)
