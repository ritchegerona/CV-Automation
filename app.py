import os
import io
import re
import zipfile
import subprocess
import shutil
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
import base64

# Set up page config
st.set_page_config(
    page_title="GCC CV Studio",
    page_icon=":material/badge:",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Slate background & Deep Indigo accents)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,400;14..32,500;14..32,600;14..32,700&display=swap');
    @import url('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css');

    /* =========================================================
       DESIGN TOKENS  (centralized so colors can be changed here)
       ========================================================= */
    :root {
        --brand:        #2b6bff;
        --brand-dark:   #1e4fcf;
        --brand-soft:   #eef4ff;
        --brand-light:  #f0f5ff;
        --brand-ink:    #1e4fcf;
        --navy:         #0b1120;
        --navy-soft:    #1e2940;
        --navy-text:    #b0bcdb;
        --navy-faint:   #5f6d92;
        --navy-lab:     #4b5880;
        --text-dark:    #0b1120;
        --text-body:    #334155;
        --text-secondary:#6d7a9e;
        --text-muted:   #8491a5;
        --text-faint:   #94a3b8;
        --bg-canvas:    #f4f6fc;
        --bg-card:      #ffffff;
        --bg-subtle:    #e8edf8;
        --border:       #dce3ef;
        --border-soft:  #e8edf8;
        --green:        #22c55e;
        --green-soft:   #e6f9ee;
        --green-ink:    #159447;
        --amber:        #f59e0b;
        --amber-soft:   #fef7e6;
        --red:          #e91e63;
        --red-soft:     #ffeaf2;
        --gray:         #6b7280;
        --radii-sm:     10px;
        --radius:       16px;
        --radius-lg:    20px;
        --radius-xl:    24px;
        --shadow-sm:    0 2px 10px rgba(0,0,0,0.04);
        --shadow:       0 4px 18px rgba(0,0,0,0.05);
        --shadow-lg:    0 8px 28px rgba(0,0,0,0.08);
        --sp-1: 4px; --sp-2: 8px; --sp-3: 12px; --sp-4: 16px;
        --sp-6: 24px; --sp-8: 32px; --sp-12: 48px; --sp-16: 64px;
        --trans: 0.2s ease;
    }

    html, body, [class*="css"] {
        font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        color: var(--text-body);
    }
    .stApp { background-color: var(--bg-canvas); }

    /* =========================================================
       MATERIAL SYMBOL HELPER
       ========================================================= */
    .m {
        font-family: 'Material Symbols Outlined', sans-serif;
        font-weight: 400; font-style: normal;
        display: inline-block; line-height: 1;
        vertical-align: -0.14em; letter-spacing: normal;
        text-transform: none; white-space: nowrap; word-wrap: normal;
        -webkit-font-feature-settings: 'liga'; font-feature-settings: 'liga';
        -webkit-font-smoothing: antialiased;
        font-variation-settings: 'FILL' 0, 'wght' 500, 'GRAD' 0, 'opsz' 24;
        color: var(--text-body);
    }
    .m-filled { font-variation-settings: 'FILL' 1, 'wght' 500, 'GRAD' 0, 'opsz' 24; }
    .m-soft { color: var(--brand); }
    .m-muted { color: var(--text-muted); }

    /* =========================================================
       JAVA / FORMS  (native Streamlit controls)
       ========================================================= */
    .stApp [data-testid="stToolbar"] { display: none; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    .main .block-container {
        padding-top: 1.25rem; padding-bottom: 3rem;
        max-width: 1160px; margin: 0 auto;
    }
    .stButton > button, .stDownloadButton > button {
        border-radius: var(--radius-sm); font-weight: 600;
        transition: all var(--trans);
        border: 1px solid var(--border);
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    .stButton > button:hover { transform: translateY(-1px); box-shadow: var(--shadow); }
    .stButton > button[kind="primary"] {
        background-color: var(--brand); border-color: var(--brand); color: #fff;
    }
    .stButton > button[kind="primary"]:hover { background-color: var(--brand-dark); }
    .stTextInput input, .stTextArea textarea, .stSelectbox > div > div,
    .stNumberInput input { border-radius: var(--radius-sm) !important; font-family: 'Plus Jakarta Sans', sans-serif; }
    .stProgress > div > div > div > div { background-color: var(--brand); }
    .stFileUploader [data-testid="stFileUploaderDropzone"] {
        border: 1.5px dashed #86D4D0; border-radius: 14px;
        background: linear-gradient(110deg, #F3FBFA 0%, #FFFFFF 50%, #F3FBFA 100%);
        padding: 1.5rem; min-height: 168px; display: flex;
        align-items: center; justify-content: center;
        transition: all var(--trans);
    }
    .stFileUploader [data-testid="stFileUploaderDropzone"]:hover {
        border-color: var(--brand); background: #E6F7F5;
        box-shadow: 0 0 0 4px rgba(14,165,164,0.08);
    }
    .stFileUploader [data-testid="stFileUploaderDropzone"] button {
        border-radius: 999px; background: #fff; border: 1px solid #D6DEE7;
        color: var(--brand); font-weight: 600; padding: 0.3rem 1.2rem; height: 44px;
    }
    .stFileUploader [data-testid="stFileUploaderDropzone"] small {
        font-size: 0.85rem; color: var(--text-body); font-weight: 600;
    }
    .stFileUploader [data-testid="stFileUploaderFile"] {
        border-radius: var(--radius-sm); border: 1px solid var(--border-soft);
    }
    [data-testid="stExpander"] { border: 1px solid var(--border-soft); border-radius: var(--radius); }
    hr { border-color: var(--border-soft); }

    /* =========================================================
       LAYOUT SHELL / TOP BAR  (reference design)
       ========================================================= */
    .tb-title { font-size: 26px; font-weight: 700; color: #0b1120; letter-spacing: -0.4px; }
    .tb-sub { font-size: 14px; color: #6d7a9e; margin-top: 2px; }
    .tb-badge {
        background: #e8edf8; padding: 6px 18px; border-radius: 30px; font-size: 13px;
        font-weight: 500; color: #0b1120; display: inline-flex; align-items: center; gap: 8px;
        white-space: nowrap;
    }
    .tb-avatar {
        width: 40px; height: 40px; border-radius: 50%; background: #2b6bff; color: #fff;
        display: inline-flex; align-items: center; justify-content: center;
        font-weight: 600; font-size: 16px;
    }
    /* Top-bar icon-only action buttons rendered as circular icon buttons */
    [data-testid="stAppViewContainer"] [data-testid="stButton"] > button:has(> span[class*="material"]:only-child) {
        width: 42px !important; height: 42px !important; min-width: 42px !important;
        padding: 0 !important; border-radius: 50%; display: inline-flex;
        align-items: center; justify-content: center;
        background: #ffffff !important; border: 1px solid #dce3ef !important;
        color: #6d7a9e !important; font-size: 1.15rem; box-shadow: none !important;
    }
    [data-testid="stAppViewContainer"] [data-testid="stButton"] > button:has(> span[class*="material"]:only-child):hover {
        background: #eef4ff !important; color: #2b6bff !important;
        border-color: #2b6bff !important; transform: none;
    }
    .page-heading { margin: 0 0 var(--sp-2); }

    /* =========================================================
       DASHBOARD  (reference design)
       ========================================================= */
    .stats-row { display: grid; grid-template-columns: 180px 1fr; gap: 20px; margin-bottom: 28px; }
    .stats-card { background:#ffffff; border-radius:16px; padding:20px 24px; box-shadow:0 2px 10px rgba(0,0,0,0.04); border:1px solid #e8edf8; }
    .stats-card .label { font-size:13px; font-weight:600; color:#6d7a9e; text-transform:uppercase; letter-spacing:0.5px; }
    .stats-card .number { font-size:42px; font-weight:700; color:#0b1120; line-height:1.1; margin-top:2px; }
    .stats-card .number span { font-size:18px; font-weight:400; color:#6d7a9e; margin-left:6px; }
    .status-card { background:#ffffff; border-radius:16px; padding:20px 24px; box-shadow:0 2px 10px rgba(0,0,0,0.04); border:1px solid #e8edf8; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:8px 20px; }
    .status-card .status-item { display:flex; align-items:center; gap:12px; }
    .status-card .status-item .dot { width:8px; height:8px; border-radius:50%; background:#22c55e; flex-shrink:0; }
    .status-card .status-item .dot.inactive { background:#b0bcdb; }
    .status-card .status-item .label { font-size:13px; color:#6d7a9e; font-weight:500; }
    .status-card .status-item .value { font-size:14px; font-weight:600; color:#0b1120; }
    .status-card .status-item .value .highlight { color:#2b6bff; }

    .studio-card { background:#ffffff; border-radius:20px; padding:34px 38px 38px; box-shadow:0 4px 18px rgba(0,0,0,0.04); border:1px solid #e8edf8; margin-bottom:28px; }
    .studio-card .studio-header { display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:12px; margin-bottom:26px; }
    .studio-card .studio-header .title-group h3 { font-size:22px; font-weight:700; color:#0b1120; letter-spacing:-0.2px; }
    .studio-card .studio-header .title-group p { font-size:14px; color:#6d7a9e; margin-top:4px; max-width:520px; }
    .studio-card .studio-header .template-badge { background:#eef4ff; color:#1e4fcf; padding:6px 20px; border-radius:30px; font-size:13px; font-weight:600; display:flex; align-items:center; gap:8px; white-space:nowrap; }

    [data-testid="stKeyedContainer"].studio-card-wrap { background:#ffffff; border:1px solid #e8edf8; border-radius:20px; padding:34px 38px 38px; box-shadow:0 4px 18px rgba(0,0,0,0.04); margin-bottom:20px; }
    [data-testid="stKeyedContainer"].studio-card-wrap [data-testid="stFileUploader"] { margin:0 0 18px; }
    [data-testid="stKeyedContainer"].studio-card-wrap [data-testid="stVerticalBlockBorderWrapper"] { background:transparent !important; border:none !important; box-shadow:none !important; }
    .studio-card .studio-header, [data-testid="stKeyedContainer"].studio-card-wrap .studio-header { display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:12px; margin-bottom:20px; }
    .studio-card .studio-header .title-group h3, [data-testid="stKeyedContainer"].studio-card-wrap .studio-header .title-group h3 { font-size:22px; font-weight:700; color:#0b1120; letter-spacing:-0.2px; }
    .studio-card .studio-header .title-group p, [data-testid="stKeyedContainer"].studio-card-wrap .studio-header .title-group p { font-size:14px; color:#6d7a9e; margin-top:4px; max-width:520px; }

    .dz-guide { text-align:center; padding:6px 0 2px; }
    .dz-guide .dz-ic { width:68px; height:68px; background:#e8edf8; border-radius:50%; display:flex; align-items:center; justify-content:center; margin:0 auto 16px; font-size:30px; color:#2b6bff; }
    .dz-guide h4 { font-size:20px; font-weight:600; color:#0b1120; margin-bottom:4px; }
    .dz-guide p { font-size:14px; color:#6d7a9e; margin-bottom:14px; }
    .formats { display:flex; justify-content:center; gap:12px; flex-wrap:wrap; margin-bottom:18px; }
    .formats .format-tag { background:#ffffff; border:1px solid #dce3ef; padding:5px 18px; border-radius:30px; font-size:13px; font-weight:600; color:#0b1120; display:inline-flex; align-items:center; gap:6px; transition:0.15s; }
    .formats .format-tag i { color:#2b6bff; font-size:14px; }
    .formats .format-tag:hover { border-color:#2b6bff; background:#f0f5ff; }

    /* Native uploader styled as the reference drop-zone */
    .studio-card [data-testid="stFileUploaderDropzone"] {
        border:2px dashed #dce3ef !important; border-radius:16px !important;
        background:#fafbff !important; padding:34px 20px !important;
        transition:all 0.25s !important; text-align:center !important;
    }
    .studio-card [data-testid="stFileUploaderDropzone"]:hover {
        border-color:#2b6bff !important; background:#f0f5ff !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        background:#0b1120 !important; color:#fff !important; border:none !important;
        padding:8px 32px !important; border-radius:30px !important; font-weight:600 !important; font-size:14px !important; height:auto !important;
    }
    [data-testid="stFileUploaderDropzone"] button:hover { background:#1a2640 !important; cursor:pointer; }
    [data-testid="stFileUploaderDropzone"] small { font-size:15px !important; color:#0b1120 !important; font-weight:600 !important; }

    .feature-grid { display:grid; grid-template-columns:repeat(auto-fit, minmax(210px, 1fr)); gap:18px; margin-top:26px; }
    .feature-card { background:#ffffff; border-radius:16px; padding:22px 20px 20px; border:1px solid #e8edf8; transition:0.2s; box-shadow:0 2px 6px rgba(0,0,0,0.02); }
    .feature-card:hover { border-color:#cbd6e8; box-shadow:0 6px 14px rgba(0,0,0,0.05); transform:translateY(-3px); }
    .feature-card .f-icon { font-size:20px; color:#2b6bff; margin-bottom:12px; display:inline-block; background:#eef4ff; padding:9px; border-radius:12px; line-height:1; width:44px; text-align:center; }
    .feature-card h5 { font-size:15px; font-weight:600; color:#0b1120; margin-bottom:4px; }
    .feature-card p { font-size:13px; color:#6d7a9e; line-height:1.5; margin-bottom:10px; }
    .feature-card .ftag { display:inline-block; margin-top:2px; font-size:11px; font-weight:600; color:#22c55e; background:#e6f9ee; padding:2px 14px; border-radius:30px; letter-spacing:0.3px; }
    .feature-card .ftag.blue { color:#1e4fcf; background:#eef4ff; }
    .feature-card .ftag.gold { color:#b8860b; background:#fef7e6; }
    .feature-card .ftag i { margin-right:4px; }

    .bottom-meta { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-top:26px; padding-top:20px; border-top:1px solid #e8edf8; font-size:13px; color:#6d7a9e; }
    .bottom-meta .left { display:flex; align-items:center; gap:20px; flex-wrap:wrap; }
    .bottom-meta .left .pill { background:#e8edf8; padding:4px 16px; border-radius:30px; font-weight:500; color:#0b1120; font-size:12px; }
    .bottom-meta .right { display:flex; align-items:center; gap:16px; }
    .bottom-meta .right i { color:#b0bcdb; }
    .page-title { font-size: 1.7rem; font-weight: 800; color: var(--text-dark); letter-spacing: -0.02em; line-height: 1.15; }
    .page-title .accent { color: var(--brand); }
    .page-sub { font-size: 0.95rem; color: var(--text-muted); margin: var(--sp-1) 0 0; max-width: 640px; line-height: 1.55; }
    .page-section { font-size: 1.02rem; font-weight: 700; color: var(--text-dark); margin: var(--sp-6) 0 var(--sp-3); display: flex; align-items: center; gap: 0.5rem; }

    /* =========================================================
       SIDEBAR  (dark navy, reference design)
       ========================================================= */
    [data-testid="stSidebar"] {
        background: #0b1120; color: #b0bcdb; border-right: none;
        min-width: 224px; padding: 14px 12px 22px;
    }
    [data-testid="stSidebar"] .block-container { padding-top: 0.25rem; padding-right: 6px; padding-left: 6px; }
    [data-testid="stSidebarCollapseButton"], [data-testid="stSidebarCollapsedControl"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }

    .sb-logo { display: flex; align-items: center; gap: 9px; margin-bottom: 4px; }
    .sb-logo-icon {
        font-size: 17px; color: #2b6bff; background: rgba(43,107,255,0.15);
        padding: 6px; border-radius: 9px; line-height: 1;
    }
    .sb-logo-name { font-size: 17px; font-weight: 700; color: #ffffff; letter-spacing: -0.3px; }
    .sb-logo-name span { color: #2b6bff; }
    .sb-subtitle {
        font-size: 11px; color: #6d7a9e; margin: 2px 0 14px; padding-left: 10px;
        border-left: 3px solid #2b6bff; font-weight: 400; line-height: 1.45;
    }
    .sb-subtitle strong { color: #d0d9f0; font-weight: 500; }
    .sb-system-divider { height: 1px; background: #1e2940; margin: 9px 2px 7px; }
    .sb-bottom-label { font-size: 10px; font-weight: 600; color: #4b5880; text-transform: uppercase; letter-spacing: 0.7px; padding: 5px 10px 2px; }
    .sb-system-link {
        display: flex; align-items: center; gap: 12px; padding: 7px 10px; border-radius: 8px;
        color: #b0bcdb; font-size: 13px; font-weight: 500;
    }
    .sb-system-link i { width: 17px; font-size: 14px; text-align: center; color: #5f6d92; }
    .sb-system-link:hover { background: rgba(43,107,255,0.08); color: #ffffff; }
    .sb-system-link:hover i { color: #2b6bff; }

    /* Radio used as the reference nav (dark) */
    [data-testid="stRadio"] > div { gap: 3px; }
    [data-testid="stRadio"] [role="radiogroup"] { gap: 3px; }
    [data-testid="stRadio"] label {
        display: flex; align-items: center; gap: 12px; height: 36px;
        width: 100%; max-width: 100%;
        padding: 0 12px !important; border-radius: 8px !important;
        font-size: 13.5px !important; font-weight: 600 !important; color: #dce4f7 !important;
        transition: background 0.2s !important; cursor: pointer; margin: 0 !important;
    }
    [data-testid="stRadio"] label:hover { background: rgba(43,107,255,0.14) !important; color: #ffffff !important; }
    [data-testid="stRadio"] label:has(input:checked) { background: rgba(43,107,255,0.18) !important; color: #ffffff !important; }
    [data-testid="stRadio"] label > div:first-child { display: none; }
    [data-testid="stRadio"] label::before {
        font-family: 'Font Awesome 6 Free', sans-serif; font-weight: 900;
        font-size: 13px; width: 16px; text-align: center; color: #aab8da;
    }
    [data-testid="stRadio"] label:nth-child(1)::before { content: '\f1ec'; } /* chart-pie */
    [data-testid="stRadio"] label:nth-child(2)::before { content: '\f1da'; } /* clock-rotate-left */
    [data-testid="stRadio"] label:nth-child(3)::before { content: '\f0c5'; } /* copy */
    [data-testid="stRadio"] label:nth-child(4)::before { content: '\f1de'; } /* sliders */
    [data-testid="stRadio"] label:nth-child(5)::before { content: '\f059'; } /* circle-question */
    [data-testid="stRadio"] label:has(input:checked)::before { color: #2b6bff; }
    [data-testid="stRadio"] label:hover::before { color: #2b6bff; }
    .sb-metric {
        background: transparent; border: none; border-radius: var(--radius);
        padding: 3px 2px; margin-bottom: 0.4rem; color: #b0bcdb;
    }
    .sb-metric-val { font-size: 24px; font-weight: 700; color: #ffffff; line-height: 1; }
    .sb-metric-lbl { font-size: 10px; color: #6d7a9e; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 0.2rem; font-weight: 600; }
    .sb-progress { height: 4px; background: #1e2940; border-radius: 999px; overflow: hidden; margin-top: 6px; max-width: 100%; }
    .sb-progress > div { height: 100%; background: #2b6bff; border-radius: 999px; transition: width 0.4s ease; }
    .sb-status-card { display: flex; align-items: flex-start; gap: 0.5rem; background: #111a2e; border: 1px solid #1e2940; border-radius: 9px; padding: 0.5rem 0.65rem; margin-bottom: 0.4rem; }
    .sb-status-ic { width: 18px; height: 18px; border-radius: 50%; background: rgba(34,197,94,0.15); color: #22c55e; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 0.1rem; }
    .sb-status-ic .m { font-size: 0.7rem; color: #22c55e; }
    .sb-status-card strong { display: block; font-size: 0.76rem; color: #e8edf8; font-weight: 600; }
    .sb-status-card span { font-size: 0.7rem; color: #5f8f7a; line-height: 1.4; }
    .sb-status-card span.miss { color: #f4648c; }
    .sb-copyright { font-size: 0.7rem; color: #4b5880; padding: 2px 2px; line-height: 1.4; }

    /* Sidebar text readability (sidebar is always navy bg) */
    [data-testid="stSidebar"] .sb-subtitle,
    [data-testid="stSidebar"] .sb-subtitle strong,
    [data-testid="stSidebar"] .sb-system-link,
    [data-testid="stSidebar"] .sb-status-card strong { color: #ffffff !important; }
    [data-testid="stSidebar"] .sb-bottom-label,
    [data-testid="stSidebar"] .sb-system-link i,
    [data-testid="stSidebar"] .sb-metric-lbl,
    [data-testid="stSidebar"] .sb-status-card span,
    [data-testid="stSidebar"] .stat-label-sm { color: #c3ceea !important; }
    [data-testid="stSidebar"] .sb-status-card span.miss { color: #ff8aa8 !important; }
    [data-testid="stSidebar"] .sb-copyright { color: #8b9ac0 !important; }
    [data-testid="stSidebar"] .sb-status-card .sb-status-ic span.m { color: #34d780 !important; }
    /* Navigation labels (unreadable at low contrast — force readable) */
    [data-testid="stSidebar"] [data-testid="stRadio"] label { color: #dce4f7 !important; }
    [data-testid="stSidebar"] [data-testid="stRadio"] label:hover,
    [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) { color: #ffffff !important; }
    [data-testid="stSidebar"] [data-testid="stRadio"] label::before { color: #aab8da !important; }
    [data-testid="stSidebar"] [data-testid="stRadio"] label:hover::before,
    [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked)::before { color: #2b6bff !important; }
    [data-testid="stSidebar"] [data-testid="stRadio"] label p,
    [data-testid="stSidebar"] [data-testid="stRadio"] label span { color: inherit !important; }

    /* =========================================================
       CARDS / BADGES / CHIPS
       ========================================================= */
    .card { background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); box-shadow: var(--shadow-sm); }
    .card-pad { padding: 1.25rem 1.35rem; }
    .card-title { font-size: 0.98rem; font-weight: 700; color: var(--text-dark); margin-bottom: 0.4rem; display: flex; align-items: center; gap: 0.5rem; }
    .card-sub { font-size: 0.82rem; color: var(--text-muted); }
    .format-badge { display: inline-block; padding: 0.28rem 0.7rem; border-radius: 999px; font-size: 0.72rem; font-weight: 700; letter-spacing: 0.04em; border: 1px solid; }
    .badge-pdf { color: #E91E63; border-color: #F8BBD9; background: #FFF0F5; }
    .badge-docx { color: #3182CE; border-color: #B1D4F2; background: #EEF6FF; }
    .badge-doc { color: #F59E0B; border-color: #FDE2B3; background: #FFF7E8; }
    .badge-txt { color: #65B96B; border-color: #C6E8C8; background: #EFF9F0; }
    .pill { display: inline-flex; align-items: center; gap: 0.35rem; padding: 0.2rem 0.6rem; border-radius: 999px; font-size: 0.72rem; font-weight: 700; }
    .pill.ok { background: var(--green-soft); color: var(--green); }
    .pill.warn { background: var(--amber-soft); color: var(--amber); }
    .pill.err { background: var(--red-soft); color: var(--red); }
    .pill.neutral { background: var(--bg-subtle); color: var(--text-muted); }
    .chip { display: inline-block; background: var(--brand-soft); color: var(--brand-ink); border: 1px solid rgba(14,165,164,0.35); border-radius: 999px; padding: 0.2rem 0.7rem; font-size: 0.78rem; font-weight: 600; margin: 0.15rem 0.25rem 0.15rem 0; }
    .chip-group { display: flex; flex-wrap: wrap; gap: 0.3rem; margin-top: 0.6rem; }

    /* =========================================================
       FEATURE / CAPABILITY
       ========================================================= */
    .category-view .feature-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; }
    .category-view .feature-card { background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); padding: 1.4rem 1.3rem; box-shadow: var(--shadow-sm); position: relative; overflow: hidden; transition: transform var(--trans), box-shadow var(--trans); }
    .category-view .feature-card::after { content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 4px; background: var(--brand); }
    .category-view .feature-card:hover { transform: translateY(-2px); box-shadow: var(--shadow); }
    .category-view .feature-icon { width: 42px; height: 42px; border-radius: 10px; background: var(--brand-soft); color: var(--brand); display: flex; align-items: center; justify-content: center; font-size: 1.4rem; margin-bottom: 0.8rem; }
    .category-view .feature-title { font-size: 1rem; font-weight: 700; color: var(--text-dark); margin-bottom: 0.35rem; }
    .category-view .feature-desc { font-size: 0.82rem; color: var(--text-muted); line-height: 1.5; margin-bottom: 0.7rem; }
    .capability-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; }
    .capability { display: flex; align-items: flex-start; gap: 0.7rem; padding: 0 0.9rem; border-left: 1px solid var(--border-soft); }
    .capability:first-child { border-left: none; padding-left: 0; }
    .capability .m { font-size: 1.3rem; color: var(--brand); margin-top: 0.05rem; }
    .capability strong { display: block; font-size: 0.86rem; color: var(--text-dark); font-weight: 700; }
    .capability span { font-size: 0.76rem; color: var(--text-muted); line-height: 1.4; }

    /* =========================================================
       UPLOAD ZONE
       ========================================================= */
    .dropzone {
        border: 1.75px dashed var(--border); border-radius: var(--radius-lg);
        background: linear-gradient(180deg, #FDFEFF, #F4FAFA);
        padding: 2.75rem 1.5rem; text-align: center; transition: all var(--trans);
    }
    .dropzone:hover, .dropzone.drag {
        border-color: var(--brand); background: var(--brand-soft);
        box-shadow: 0 0 0 4px rgba(14,165,164,0.08);
    }
    .dz-icon { font-size: 2.6rem; color: var(--brand); margin-bottom: 0.6rem; }
    .dz-title { font-size: 1.15rem; font-weight: 800; color: var(--text-dark); margin-bottom: 0.25rem; }
    .dz-sub { font-size: 0.88rem; color: var(--text-muted); margin-bottom: 1rem; }
    .dz-meta { font-size: 0.78rem; color: var(--text-faint); margin-top: 0.9rem; }

    /* =========================================================
       DECORATIVE FOLDER ILLUSTRATION (upload dropzone right side)
       ========================================================= */
    .folder-illus { position: relative; width: 100%; max-width: 240px; height: 150px; display: flex; align-items: center; justify-content: center; margin: 0 auto; }
    .folder-sheet { position: absolute; left: 50%; top: 22px; width: 88px; height: 112px; background: #fff; border: 1px solid #E0E7EF; border-radius: 8px; box-shadow: 0 4px 12px rgba(15,23,42,0.08); transform: translateX(-50%); }
    .folder-sheet.s1 { left: 39%; transform: rotate(-8deg); }
    .folder-sheet.s2 { left: 61%; transform: rotate(8deg); }
    .folder-main { position: relative; width: 150px; height: 102px; background: linear-gradient(180deg, #2FB9B4, #0E9F9A); border-radius: 12px 12px 14px 14px; box-shadow: 0 8px 20px rgba(14,159,154,0.28); display: flex; align-items: center; justify-content: center; z-index: 2; }
    .folder-tab { position: absolute; top: -14px; left: 14px; width: 46px; height: 18px; background: #0E9F9A; border-radius: 8px 8px 0 0; }
    .folder-cv { color: #fff; font-size: 1.35rem; font-weight: 800; letter-spacing: 0.02em; font-family: 'Inter', sans-serif; z-index: 3; }
    .folder-badge { position: absolute; z-index: 5; font-size: 0.62rem; font-weight: 700; border-radius: 999px; padding: 0.22rem 0.55rem; border: 1px solid; box-shadow: 0 2px 6px rgba(15,23,42,0.1); background: #fff; }
    .folder-badge.b-pdf { top: 4px; right: 8px; color: #E91E63; border-color: #F8BBD9; }
    .folder-badge.b-docx { top: 40px; left: 0; color: #3182CE; border-color: #B1D4F2; }
    .folder-badge.b-doc { bottom: 6px; right: 2px; color: #F59E0B; border-color: #FDE2B3; }
    .folder-badge.b-txt { bottom: -2px; left: 6px; color: #65B96B; border-color: #C6E8C8; }

    @media (max-width: 899px) {
        .folder-illus { display: none; }
    }


    /* =========================================================
       FILE QUEUE / PROCESS CARD
       ========================================================= */
    .file-row { display: flex; align-items: center; gap: 0.8rem; background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius-sm); padding: 0.7rem 0.95rem; margin-bottom: 0.5rem; box-shadow: var(--shadow-sm); }
    .file-row .f-icon { width: 34px; height: 34px; border-radius: 8px; background: var(--brand-soft); color: var(--brand); display: flex; align-items: center; justify-content: center; font-size: 1.1rem; flex-shrink: 0; }
    .file-row .f-name { font-weight: 600; color: var(--text-dark); font-size: 0.88rem; }
    .file-row .f-size { font-size: 0.76rem; color: var(--text-faint); }
    .file-row .f-status { margin-left: auto; }
    .proc-steps { display: flex; align-items: center; gap: 0.4rem; font-size: 0.72rem; color: var(--text-faint); }
    .proc-step { display: inline-flex; align-items: center; gap: 0.25rem; }
    .proc-step.on { color: var(--brand); font-weight: 700; }
    .proc-step.done { color: var(--green); }
    .proc-arrow { color: var(--border); }
    .workflow { display: flex; align-items: center; gap: 0.3rem; flex-wrap: wrap; background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); padding: 0.7rem 0.9rem; box-shadow: var(--shadow-sm); font-size: 0.72rem; }
    .workflow .wf-label { font-weight: 700; color: var(--text-muted); margin-right: 0.35rem; letter-spacing: 0.04em; }
    .workflow .wf-step { display: inline-flex; align-items: center; gap: 0.25rem; color: var(--text-faint); font-weight: 600; }
    .workflow .wf-step.done { color: var(--green); }
    .workflow .wf-step.on { color: var(--brand); font-weight: 700; }
    .workflow .wf-arrow { color: var(--border); }

    /* =========================================================
       RESULTS / TABLE / CANDIDATE
       ========================================================= */
    .data-table { width: 100%; border-collapse: collapse; background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); overflow: hidden; font-size: 0.85rem; }
    .data-table th { text-align: left; padding: 0.7rem 1rem; background: var(--bg-subtle); color: var(--text-muted); font-weight: 700; font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.04em; border-bottom: 1px solid var(--border-soft); }
    .data-table td { padding: 0.7rem 1rem; border-bottom: 1px solid var(--border-soft); color: var(--text-body); }
    .data-table tr:last-child td { border-bottom: none; }
    .data-table tr:hover td { background: var(--brand-soft); }
    .candidate-card { display: flex; align-items: center; gap: 1rem; background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); padding: 0.9rem 1rem; margin-bottom: 0.6rem; box-shadow: var(--shadow-sm); }
    .candidate-card.active { border-color: var(--brand); box-shadow: 0 0 0 3px rgba(14,165,164,0.10); }
    .candidate-avatar { width: 44px; height: 44px; border-radius: 10px; background: linear-gradient(135deg, var(--brand), var(--brand-dark)); color: #fff; display: flex; align-items: center; justify-content: center; font-size: 1rem; font-weight: 700; flex-shrink: 0; }
    .candidate-name { font-size: 0.95rem; font-weight: 700; color: var(--text-dark); }
    .candidate-meta { font-size: 0.76rem; color: var(--text-muted); }

    /* =========================================================
       EMPTY / ERROR / SUCCESS STATES
       ========================================================= */
    .state-block { text-align: center; padding: 3rem 1.5rem; }
    .state-icon { width: 64px; height: 64px; border-radius: 50%; background: var(--brand-soft); color: var(--brand); font-size: 2rem; display: inline-flex; align-items: center; justify-content: center; margin-bottom: 1rem; }
    .state-title { font-size: 1.15rem; font-weight: 800; color: var(--text-dark); margin-bottom: 0.35rem; }
    .state-desc { font-size: 0.88rem; color: var(--text-muted); max-width: 420px; margin: 0 auto; line-height: 1.5; }

    /* =========================================================
       STATS
       ========================================================= */
    .stat-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 0.8rem; }
    .stat-box { background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); padding: 0.9rem; text-align: center; box-shadow: var(--shadow-sm); }
    .stat-number { font-size: 1.5rem; font-weight: 800; color: var(--brand); }
    .stat-label-sm { font-size: 0.7rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.04em; }

    /* =========================================================
       PRIVACY NOTICE
       ========================================================= */
    .privacy-note { display: flex; align-items: center; gap: 0.5rem; font-size: 0.78rem; color: var(--text-muted); background: var(--bg-subtle); border: 1px solid var(--border-soft); border-radius: var(--radius-sm); padding: 0.55rem 0.85rem; }
    .privacy-note .m { color: var(--brand); }

    /* =========================================================
       STAT CARD / RECORD ROW / STATUS PILL / EMPTY STATE
       ========================================================= */
    .stat-card { background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); padding: 0.9rem; text-align: center; box-shadow: var(--shadow-sm); }
    .record-row { display: flex; align-items: center; gap: 0.8rem; background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius-sm); padding: 0.7rem 0.95rem; margin-bottom: 0.5rem; box-shadow: var(--shadow-sm); }
    .record-row .f-icon { width: 34px; height: 34px; border-radius: 8px; background: var(--brand-soft); color: var(--brand); display: flex; align-items: center; justify-content: center; font-size: 1.1rem; flex-shrink: 0; }
    .record-row .f-name { font-weight: 600; color: var(--text-dark); font-size: 0.88rem; }
    .record-row .f-size { font-size: 0.76rem; color: var(--text-faint); }
    .record-row .f-status { margin-left: auto; }
    .status-pill { display: inline-flex; align-items: center; gap: 0.3rem; padding: 0.25rem 0.6rem; border-radius: 999px; font-size: 0.72rem; font-weight: 700; border: 1px solid; white-space: nowrap; }
    .status-pill.status-processing { color: var(--amber); background: var(--amber-soft); border-color: #FDBA74; }
    .status-pill.status-success { color: var(--green); background: var(--green-soft); border-color: #B9F6CA; }
    .status-pill.status-error { color: var(--red); background: var(--red-soft); border-color: #FECACA; }
    .status-pill.status-warn { color: var(--amber); background: var(--amber-soft); border-color: #FDBA74; }
    .empty-state { text-align: center; padding: 3rem 1.5rem; background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); }
    .empty-title { font-size: 1.15rem; font-weight: 800; color: var(--text-dark); margin: 0.5rem 0 0.25rem; }
    .empty-desc { font-size: 0.88rem; color: var(--text-muted); max-width: 420px; margin: 0 auto; line-height: 1.5; }

    /* =========================================================
       PROCESS CARD / EXPORT PANEL / PROFILE HERO
       ========================================================= */
    .process-card { display: flex; align-items: center; gap: 0.6rem; background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius-sm); padding: 0.7rem 0.95rem; margin-bottom: 0.5rem; box-shadow: var(--shadow-sm); font-size: 0.85rem; }
    .process-card-name { flex: 1; font-weight: 600; color: var(--text-dark); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .process-card-status { display: inline-flex; align-items: center; gap: 0.3rem; font-weight: 700; font-size: 0.72rem; }
    .process-card-status .m { font-size: 1rem; }
    .process-card-status.status-processing { color: var(--amber); }
    .process-card-status.status-success { color: var(--green); }
    .process-card-status.status-error { color: var(--red); }
    .export-panel { background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); padding: 1.25rem 1.35rem; box-shadow: var(--shadow-sm); }
    .export-panel-title { font-size: 0.92rem; font-weight: 600; color: var(--text-dark); margin-bottom: 0.9rem; display: flex; align-items: center; gap: 0.5rem; }
    .profile-hero { background: linear-gradient(120deg, var(--brand-soft), var(--bg-card)); border: 1px solid var(--border-soft); border-radius: var(--radius-lg); padding: 1.6rem 1.8rem; margin: 0.4rem 0 1.25rem; box-shadow: var(--shadow-sm); }
    .profile-hero-name { font-size: 1.5rem; font-weight: 800; color: var(--text-dark); letter-spacing: -0.02em; }
    .profile-hero-sub { font-size: 0.9rem; font-weight: 600; color: var(--brand); margin: 0.25rem 0; }
    .profile-hero-summary { font-size: 0.88rem; color: var(--text-muted); line-height: 1.55; margin: 0.35rem 0 0.75rem; }
    .profile-hero .chip-group { margin-top: 0.6rem; }

    /* =========================================================
       ANIMATION
       ========================================================= */
    @keyframes fadeSlideIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
    .fade-in { animation: fadeSlideIn 0.25s ease-out forwards; }
    .pulse { animation: pulse 1.4s ease-in-out infinite; }
    @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.55; } }

    /* =========================================================
       FLOATING MANAGE APP BUTTON
       ========================================================= */
    .floating-manage {
        position: fixed; right: 28px; bottom: 24px; z-index: 999;
    }
    .floating-manage > button {
        height: 48px; padding: 0 20px; border-radius: 14px;
        background: var(--brand) !important; color: #fff !important;
        border: none !important; box-shadow: 0 8px 24px rgba(14,165,164,0.25) !important;
        font-weight: 700 !important; font-size: 0.9rem !important;
        display: inline-flex !important; align-items: center; gap: 0.5rem;
        transition: all var(--trans) !important;
    }
    .floating-manage > button:hover { background: var(--brand-dark) !important; transform: translateY(-2px); }

    /* =========================================================
       RESPONSIVE
       ========================================================= */
    @media (max-width: 900px) {
        .feature-grid { grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); }
        .stats-row { grid-template-columns: 1fr; }
        .studio-card, [data-testid="stKeyedContainer"].studio-card-wrap { padding: 24px 20px; }
    }
    @media (max-width: 600px) {
        .feature-grid { grid-template-columns: 1fr 1fr; }
        .dz-guide { padding: 2px 0; }
        .formats { gap: 6px; }
        .formats .format-tag { font-size: 12px; padding: 4px 12px; }
        .status-card { flex-direction: column; align-items: flex-start; gap: 8px; }
        .bottom-meta { flex-direction: column; align-items: flex-start; }
    }
    @media (max-width: 420px) {
        .feature-grid { grid-template-columns: 1fr; }
    }
    .stColumn .stButton > button { width: 100%; }

    /* =========================================================
       SAAS POLISH  (hero, panels, result cards, footer)
       ========================================================= */
    .hero { padding: 8px 0 22px; margin-bottom: 20px; }
    .hero-kicker { display: inline-flex; align-items: center; gap: 8px; font-size: 12px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; color: var(--brand); background: var(--brand-light); padding: 5px 14px; border-radius: 999px; margin-bottom: 14px; }
    .hero-title { font-size: 40px; font-weight: 800; color: var(--text-dark); letter-spacing: -0.03em; line-height: 1.1; margin: 0; }
    .hero-title span { color: var(--brand); }
    .hero-underline { height: 5px; width: 118px; background: linear-gradient(90deg, var(--brand), #7ba2ff, transparent); border-radius: 999px; margin-top: 16px; }
    .hero-sub { font-size: 16px; color: var(--text-secondary); margin-top: 14px; max-width: 660px; line-height: 1.6; }
    .hero-actions { margin-top: 20px; display: flex; gap: 10px; flex-wrap: wrap; }

    .panel { width: 100%; }
    .result-card { background: var(--bg-card); border: 1px solid var(--border-soft); border-left: 4px solid var(--brand); border-radius: var(--radius); padding: 18px 20px; box-shadow: var(--shadow-sm); margin-bottom: 14px; }
    .result-card.green { border-left-color: var(--green); }
    .result-card.amber { border-left-color: var(--amber); }
    .result-card .rc-title { font-size: 15px; font-weight: 700; color: var(--text-dark); display: flex; align-items: center; gap: 8px; margin-bottom: 6px; }
    .result-card .rc-body { font-size: 13px; color: var(--text-secondary); line-height: 1.55; }

    .tip-panel { background: var(--bg-card); border: 1px solid var(--border-soft); border-radius: var(--radius); padding: 18px; box-shadow: var(--shadow); margin-bottom: 16px; }
    .tip-panel .tp-title { font-size: 13px; font-weight: 700; color: var(--text-dark); text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }
    .tip-panel .tp-title .m { color: var(--brand); }
    .tip-item { display: flex; gap: 10px; align-items: flex-start; font-size: 13px; color: var(--text-secondary); line-height: 1.5; margin-bottom: 10px; }
    .tip-item .m { font-size: 16px; color: var(--brand); margin-top: 1px; }

    .app-footer { margin-top: 40px; padding-top: 18px; border-top: 1px solid var(--border-soft); text-align: center; font-size: 13px; color: var(--text-muted); }
    .app-footer .h { color: #e91e63; }

    .sidebar-upload { background:#111a2e; border:1px solid #1e2940; border-radius:10px; padding:11px 10px; margin: 6px 0 4px; }
    .sidebar-upload .su-title { font-size:11px; font-weight:700; color:#fff; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:9px; display:flex; align-items:center; gap:7px; }
    .sidebar-upload .su-title .m { color:#2b6bff; }
    .sidebar-upload [data-testid="stFileUploaderDropzone"] {
        background:#0f1e3d !important; border:1.5px dashed #4568c7 !important; border-radius:9px !important;
        padding:12px 10px !important; min-height: 74px !important;
    }
    .sidebar-upload [data-testid="stFileUploaderDropzone"]:hover { border-color:#2b6bff !important; background:#12263f !important; }
    .sidebar-upload [data-testid="stFileUploaderDropzone"] button {
        background:#2b6bff !important; color:#fff !important; border:none !important;
        border-radius:999px !important; font-weight:600 !important; font-size:12px !important; height:auto !important; padding:5px 14px !important;
    }
    .sidebar-upload [data-testid="stFileUploaderDropzone"] button:hover { background:#3b7bff !important; }
    .sidebar-upload [data-testid="stFileUploaderDropzone"] small { font-size:12px !important; color:#c3ceea !important; font-weight:600 !important; }
    .sidebar-upload .sidebar-divider { height:1px; background:#1e2940; margin:10px 0; }
    .sidebar-upload .sidebar-divider + .stButton, .sidebar-upload .stButton { margin-top: 4px; }

    [data-testid="stSidebar"] [data-testid="stToggle"] { padding-top: 2px; padding-bottom: 2px; }
    [data-testid="stSidebar"] [data-testid="stToggle"] label { padding: 2px 0; gap: 10px; }
    [data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] [data-testid="stCaptionContainer"] { margin-top: 2px !important; font-size: 11px !important; }
    [data-testid="stSidebar"] .block-container > div { margin-top: 0.15rem; margin-bottom: 0.15rem; }

    .stButton > button, .stDownloadButton > button { padding: 0.6rem 1.1rem; }
    .stButton > button:hover, .stDownloadButton > button:hover { transform: translateY(-1px); box-shadow: var(--shadow); border-color: var(--brand); }
</style>
""", unsafe_allow_html=True)

# --- DARK MODE override (injected after light styles so tokens win) ---
if st.session_state.get("dark_mode", False):
    st.markdown("""<style>
    :root {
        --brand:#2b6bff; --brand-dark:#3b7bff; --brand-soft:#0f1e3d; --brand-light:#0a1730; --brand-ink:#7ba2ff;
        --text-dark:#E9EFF7; --text-body:#C4CEDC; --text-secondary:#96A4B8; --text-muted:#7A879B; --text-faint:#64748B;
        --bg-canvas:#0D1320; --bg-card:#151D2D; --bg-subtle:#1C2537;
        --border:#26314A; --border-soft:#1E2840;
        --green:#34D780; --green-soft:#0F2E20; --green-ink:#3FE08C;
        --amber:#F5B85C; --amber-soft:#33250F;
        --red:#F4648C; --red-soft:#33141F;
        --gray:#8A94A6;
        --shadow-sm:0 1px 3px rgba(0,0,0,0.45); --shadow:0 4px 18px rgba(0,0,0,0.5); --shadow-lg:0 8px 28px rgba(0,0,0,0.55);
    }
    [data-testid="stAppViewContainer"] { background-color:#0d1320 !important; }
    [data-testid="stSidebar"] { background:#0b1120 !important; }
    [data-testid="stHeader"] { background:transparent !important; }

    .tb-title { color:#e9eff7 !important; }
    .tb-badge { background:#1c2537 !important; color:#c4cedc !important; }
    [data-testid="stAppViewContainer"] [data-testid="stButton"] > button:has(> span[class*="material"]:only-child) {
        background:#151d2d !important; border-color:#26314a !important; color:#96a4b8 !important;
    }
    [data-testid="stAppViewContainer"] [data-testid="stButton"] > button:has(> span[class*="material"]:only-child):hover {
        background:#1c2537 !important; color:#2b6bff !important;
    }

    [data-testid="stRadio"] label { color:#b0bcdb !important; }
    [data-testid="stRadio"] label:hover { background:rgba(43,107,255,0.08) !important; }
    [data-testid="stRadio"] label::before { color:#5f6d92 !important; }
    [data-testid="stRadio"] label:has(input:checked) { background:rgba(43,107,255,0.12) !important; color:#ffffff !important; }
    [data-testid="stRadio"] label:has(input:checked)::before { color:#2b6bff !important; }

    .stats-card, .status-card, .studio-card, .feature-card,
    [data-testid="stKeyedContainer"].studio-card-wrap { background:#151d2d !important; border-color:#26314a !important; }
    .stats-card .label, .status-card .status-item .label, .studio-card .studio-header .title-group p,
    [data-testid="stKeyedContainer"].studio-card-wrap .studio-header .title-group p,
    .feature-card p, .bottom-meta { color:#96a4b8 !important; }
    .stats-card .number, .status-card .status-item .value, .studio-card .studio-header .title-group h3,
    [data-testid="stKeyedContainer"].studio-card-wrap .studio-header .title-group h3,
    .feature-card h5 { color:#e9eff7 !important; }
    .feature-card:hover { border-color:#2b6bff !important; box-shadow:0 6px 14px rgba(0,0,0,0.4) !important; }
    .status-card .status-item .value .highlight { color:#7ba2ff !important; }
    .template-badge { background:#0f1e3d !important; color:#7ba2ff !important; }
    .formats .format-tag { background:#151d2d !important; border-color:#26314a !important; color:#c4cedc !important; }
    .formats .format-tag i { color:#2b6bff !important; }
    .dz-guide h4 { color:#e9eff7 !important; }
    .dz-guide p { color:#96a4b8 !important; }
    .dz-guide .dz-ic { background:#1c2537 !important; color:#2b6bff !important; }
    .bottom-meta .left .pill { background:#1c2537 !important; color:#c4cedc !important; }
    .feature-card .ftag { background:#0f2e20 !important; color:#34d780 !important; }
    .feature-card .ftag.blue { background:#0f1e3d !important; color:#7ba2ff !important; }
    .feature-card .ftag.gold { background:#33250f !important; color:#f5b85c !important; }

    .studio-card [data-testid="stFileUploaderDropzone"] {
        background:linear-gradient(110deg,#0e1a2e 0%,#151d2d 50%,#0e1a2e 100%) !important;
        border-color:#2b6bff !important;
    }
    .studio-card [data-testid="stFileUploaderDropzone"]:hover { background:#0f2036 !important; }
    [data-testid="stFileUploaderDropzone"] button { background:#2b6bff !important; color:#fff !important; border-color:#2b6bff !important; }
    [data-testid="stFileUploaderDropzone"] small { color:#c4cedc !important; }

    .card, .navbar, .process-card, .export-panel, .stat-card { background:#151d2d; border-color:#26314a; }
    .capability-chip, .pill, .badge { border-color:#26314a; }

    .folder-sheet, .folder-main { box-shadow:0 3px 14px rgba(0,0,0,0.5); }

    .stExpander, [data-testid="stExpander"] { border-color:#26314a; }
    .stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab"] { color:#c4cedc; }
    input, [data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea,
    .stSelectbox [data-baseweb="select"] > div, .stNumberInput input {
        color:#e9eff7; background-color:#151d2d !important;
    }
    .stSelectbox [data-baseweb="select"] > div { border-color:#26314a !important; }

    .hero-title { color:#e9eff7 !important; }
    .hero-kicker { background:#0f1e3d !important; color:#7ba2ff !important; }
    .hero-sub { color:#96a4b8 !important; }
    .result-card, .tip-panel { background:#151d2d !important; border-color:#26314a !important; }
    .result-card .rc-title, .tip-panel .tp-title { color:#e9eff7 !important; }
    .result-card .rc-body, .tip-item { color:#96a4b8 !important; }
    .app-footer { border-top-color:#26314a !important; color:#7a879b !important; }
    </style>""", unsafe_allow_html=True)


# Define directories and file targets
WORKSPACE_DIR = os.getcwd()
SUMMARY_DIR = os.path.join(WORKSPACE_DIR, "Candidate CV Summary")
GCC_TEMPLATE_NAME = "GCC_CV_FORMAT.doc"
GCC_TEMPLATE_PATH = os.path.join(WORKSPACE_DIR, GCC_TEMPLATE_NAME)
LOGO_PATH = os.path.join(WORKSPACE_DIR, "GCC_Header.png")

def _logo_base64():
    with open(LOGO_PATH, "rb") as f:
        return base64.b64encode(f.read()).decode()

# Ensure the output summary directory exists
os.makedirs(SUMMARY_DIR, exist_ok=True)

# Initialize Session State
if "processed_count" not in st.session_state:
    st.session_state.processed_count = 0
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "view" not in st.session_state:
    st.session_state.view = "Dashboard"
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

# --- SIDEBAR (enterprise navigation) ---
NAV_ITEMS = [
    ("dashboard",        "Dashboard"),
    ("history",          "Processing History"),
    ("description",      "Templates"),
    ("settings",         "Settings"),
    ("help",             "Help & Guide"),
]

with st.sidebar:
    # Logo + subtitle
    st.markdown("""<div class='sb-logo'>
      <div class='sb-logo-icon'><i class='fas fa-file-pen'></i></div>
      <div class='sb-logo-name'>GCC <span>CV Studio</span></div>
    </div>
    <div class='sb-subtitle'>Enterprise CV Parser &amp; Standardizer<br/><strong>Developed by Ritche Gerona</strong></div>""", unsafe_allow_html=True)

    # Navigation (reference nav items via radio)
    nav_values = [lbl for _, lbl in NAV_ITEMS]
    st.radio(
        "Navigation",
        nav_values,
        index=nav_values.index(st.session_state.view),
        key="nav_select",
        label_visibility="collapsed",
        on_change=lambda: setattr(st.session_state, "view", st.session_state.nav_select),
    )

    st.markdown("<div class='sb-system-divider'></div>", unsafe_allow_html=True)

    # --- UPLOAD & CONFIG (control panel) ---
    st.markdown("""<div class='sidebar-upload'>
      <div class='su-title'><span class='m'>cloud_upload</span> Upload CVs</div>""", unsafe_allow_html=True)

    cv_uploader_key = f"cv_uploader_{st.session_state.uploader_key}"
    uploaded_files = st.file_uploader(
        "Upload raw CVs",
        type=["pdf", "docx", "doc", "txt"],
        accept_multiple_files=True,
        help="Supported files: PDF (.pdf), Microsoft Word (.docx, .doc), and plain text (.txt)",
        key=cv_uploader_key,
        label_visibility="collapsed",
    )

    uploaded_photo = None
    if uploaded_files and len(uploaded_files) == 1:
        uploaded_photo = st.file_uploader(
            "Custom profile photo (JPEG/PNG)",
            type=["jpg", "jpeg", "png"],
            help="Optional — if omitted, the studio will try to extract the candidate photo from the raw CV.",
            key=f"photo_uploader_{st.session_state.uploader_key}",
        )

    st.markdown("<div class='sidebar-divider'></div>", unsafe_allow_html=True)

    _u1, _u2 = st.columns([1, 1])
    with _u1:
        st.markdown("##### Batch")
        if uploaded_files:
            if st.button(":material/close:" " Clear", key="clear_files_btn", help="Clear all uploaded files", use_container_width=True):
                st.session_state.uploader_key += 1
                for _k in ["batch_results", "selected_candidate_idx"]:
                    if _k in st.session_state:
                        del st.session_state[_k]
                st.rerun()
    with _u2:
        process_btn = st.button(":material/auto_fix_high:" " Standardize", key="sb_process_btn", type="primary", use_container_width=True, disabled=not uploaded_files)

    st.markdown("</div>", unsafe_allow_html=True)

    # --- PREFERENCES (config widgets) ---
    st.markdown("<div class='sb-bottom-label'>Preferences</div>", unsafe_allow_html=True)
    if "pref_auto_pdf" not in st.session_state:
        st.session_state.pref_auto_pdf = True
    if "pref_save_profiles" not in st.session_state:
        st.session_state.pref_save_profiles = True
    st.toggle("Auto-generate PDF per candidate", key="pref_auto_pdf")
    st.toggle("Save local profile .txt files", key="pref_save_profiles")
    st.caption("Preferences persist for the current session.")

    st.markdown("<div class='sb-system-divider'></div>", unsafe_allow_html=True)
    st.markdown("<div class='sb-bottom-label'>System</div>", unsafe_allow_html=True)
    st.markdown("""<div class='sb-system-link'><i class='fas fa-server'></i> Status</div>
<div class='sb-system-link'><i class='fas fa-user-shield'></i> Security</div>""", unsafe_allow_html=True)
    st.markdown("<div class='sb-system-divider'></div>", unsafe_allow_html=True)

    # Session stats
    _br = st.session_state.get("batch_results", {})
    _total_done = st.session_state.processed_count
    st.markdown(f"""<div class='sb-metric'>
      <div class='sb-metric-val'>{_total_done}</div>
      <div class='sb-metric-lbl'>Processed This Session</div>
      <div class='sb-progress'><div style='width:{min(100, _total_done * 10)}%'></div></div>
    </div>""", unsafe_allow_html=True)
    if _br:
        _ok = sum(1 for r in _br.values() if r.get("status") == "Success")
        _err = sum(1 for r in _br.values() if r.get("status") == "Error")
        st.markdown(f"""<div class='stat-row'>
          <div class='stat-box'><div class='stat-number'>{_ok}</div><div class='stat-label-sm'>Success</div></div>
          <div class='stat-box' style='color:var(--red)'><div class='stat-number' style='color:var(--red)'>{_err}</div><div class='stat-label-sm'>Errors</div></div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div class='sb-system-divider'></div>", unsafe_allow_html=True)

    # System status
    summary_exists = os.path.exists(SUMMARY_DIR)
    gcc_exists = os.path.exists(GCC_TEMPLATE_PATH)
    st.markdown(f"""<div class='sb-status-card'>
      <div class='sb-status-ic'><span class='m'>check</span></div>
      <div><strong>Local Folder Status</strong>
      <span>{'Candidate CV Summary/ is active' if summary_exists else 'Candidate CV Summary/ missing'}</span></div>
    </div>""", unsafe_allow_html=True)
    st.markdown(f"""<div class='sb-status-card'>
      <div class='sb-status-ic'><span class='m'>check</span></div>
      <div><strong>CV Format</strong>
      <span>{'GCC_CV_FORMAT.doc is loaded as the template' if gcc_exists else 'GCC_CV_FORMAT.doc missing'}</span></div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<div class='sb-system-divider'></div>", unsafe_allow_html=True)
    st.markdown("<div class='sb-copyright'>&copy; 2025 GCC CV Studio<br/>All rights reserved.</div>", unsafe_allow_html=True)


# --- Toolbar action dialogs ---
@st.dialog("Share this App")
def _share_dialog():
    _s_url = "https://github.com/ritchegerona/CV-Automation"
    st.markdown("<div class='share-dlg'><span class='m' style='color:var(--brand);font-size:1.7rem;'>share</span><h4 style='margin:0 0 0.2rem;'>Share GCC CV Studio</h4><p style='margin:0;color:var(--text-muted);font-size:0.85rem;'>Invite teammates to this CV processing studio.</p></div>", unsafe_allow_html=True)
    st.markdown("<div class='sb-divider'></div>", unsafe_allow_html=True)
    st.caption("App link")
    st.code(_s_url, language="text")
    _email_target = st.text_input("Send to (optional)", placeholder="name@company.com")
    st.download_button(
        ":material/link:" "  Download share link (.txt)",
        data=f"GCC CV Studio\nShare this link:\n{_s_url}",
        file_name="msr-cv-studio-share.txt",
        mime="text/plain",
        use_container_width=True,
    )

@st.dialog("Notifications")
def _notifications_dialog():
    st.markdown("<h4 style='margin:0 0 0.6rem;'>Notifications</h4>", unsafe_allow_html=True)
    _n = st.session_state.processed_count
    if _n > 0:
        st.success(f":material/task_alt:  {_n} CV(s) processed this session.")
    else:
        st.info(":material/inbox:  No notifications yet. Process a CV to see updates here.")
    st.caption("All clear — no new alerts.")

@st.dialog("Profile")
def _profile_dialog():
    st.markdown(f"""<div class='profile-hero'>
      <div class='profile-hero-avatar'><span class='m'>person</span><span class='online-dot'></span></div>
      <div>
        <div class='profile-hero-name'>Ritche Gerona</div>
        <div class='profile-hero-role'>Developer · GCC CV Studio</div>
      </div>
    </div>""", unsafe_allow_html=True)
    st.markdown("<div class='sb-divider'></div>", unsafe_allow_html=True)
    st.write("**Studio:** Enterprise CV Parser & Standardizer")
    st.write(f"**Processed this session:** {st.session_state.processed_count}")
    st.write(f"**Theme:** {'Dark' if st.session_state.dark_mode else 'Light'}")

# --- HERO HEADER + action controls ---
_hx1, _hx2 = st.columns([9, 3], vertical_alignment="center")
with _hx1:
    st.markdown("""<div class='hero'>
      <span class='hero-kicker'><i class='fas fa-bolt'></i> Enterprise CV Automation</span>
      <div class='hero-title'>GCC <span>CV Studio</span></div>
      <div class='hero-underline'></div>
      <div class='hero-sub'>Parse, standardize, and export GCC-aligned, corporate-ready resumes — fully offline, in batch.</div>
    </div>""", unsafe_allow_html=True)
with _hx2:
    st.markdown("<div class='hero-actions'>", unsafe_allow_html=True)
    _ha1, _ha2, _ha3, _ha4, _ha5 = st.columns([1, 1, 1, 1, 2], vertical_alignment="center")
    with _ha1:
        if st.button(":material/share:", key="share_btn", help="Share this app", use_container_width=True):
            _share_dialog()
    with _ha2:
        theme_clicked = st.button(
            ":material/dark_mode:" if not st.session_state.dark_mode else ":material/light_mode:",
            key="theme_toggle_btn",
            help="Toggle dark mode",
            use_container_width=True,
        )
    with _ha3:
        if st.button(":material/notifications:", key="notif_btn", help="Notifications", use_container_width=True):
            _notifications_dialog()
    with _ha4:
        if st.button(":material/person:", key="profile_btn", help="Profile", use_container_width=True):
            _profile_dialog()
    with _ha5:
        st.markdown("<span class='tb-badge'><i class='fas fa-circle-check' style='color:#2b6bff;'></i> v2.4.1</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div class='app-top-spacer' style='height:4px;'></div>", unsafe_allow_html=True)

if theme_clicked:
    st.session_state.dark_mode = not st.session_state.dark_mode
    st.rerun()

# Helper functions
def parse_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                full_text.append(cell.text)
    return '\n'.join(list(dict.fromkeys(full_text)))  # Remove duplicates preserving order

def parse_txt(file_bytes):
    return file_bytes.decode("utf-8", errors="ignore")

def parse_pdf(file_bytes):
    import pypdf
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    full_text = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text.append(text)
    return "\n".join(full_text)

def parse_doc(file_bytes, filename):
    import tempfile
    soffice_bin = get_soffice_path()
    if not soffice_bin:
        raise Exception("LibreOffice (soffice) is required to convert .doc files but was not found on this system.")
    with tempfile.TemporaryDirectory() as temp_dir:
        safe_doc_name = "input.doc"
        doc_path = os.path.join(temp_dir, safe_doc_name)
        with open(doc_path, "wb") as f:
            f.write(file_bytes)
        cmd = [
            soffice_bin,
            "--headless",
            "-env:UserInstallation=file:///" + os.path.join(temp_dir, "soffice_profile"),
            "--convert-to",
            "docx",
            "--outdir",
            temp_dir,
            doc_path
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise Exception(f"Failed to convert .doc to .docx: {result.stderr}")
        docx_filename = "input.docx"
        docx_path = os.path.join(temp_dir, docx_filename)
        if os.path.exists(docx_path):
            with open(docx_path, "rb") as f:
                docx_bytes = f.read()
            return parse_docx(docx_bytes), docx_bytes
        else:
            raise Exception("Converted .docx file not found in temporary directory.")

@st.cache_data(show_spinner=False)
def get_soffice_path():
    soffice_path = shutil.which("soffice")
    if soffice_path:
        return soffice_path
    mac_paths = [
        "/opt/homebrew/bin/soffice",
        "/usr/local/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    ]
    for path in mac_paths:
        if os.path.exists(path):
            return path
    return None

def extract_picture(file_bytes, file_ext, temp_docx_path=None):
    try:
        if file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    img_data = rel.target_part.blob
                    if len(img_data) > 5000:  # Filter out tiny icons
                        return img_data
        elif file_ext == "doc" and temp_docx_path:
            if os.path.exists(temp_docx_path):
                doc = Document(temp_docx_path)
                for rel in doc.part.rels.values():
                    if "image" in rel.target_ref:
                        img_data = rel.target_part.blob
                        if len(img_data) > 5000:
                            return img_data
        elif file_ext == "pdf":
            import pypdf
            reader = pypdf.PdfReader(io.BytesIO(file_bytes))
            for page in reader.pages:
                for image_file_object in page.images:
                    img_data = image_file_object.data
                    if len(img_data) > 5000:
                        return img_data
    except Exception as e:
        print(f"Error extracting picture: {e}")
    return None

def _add_form_field(doc, label, value):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r0 = p.add_run(f"{label}: ")
    r0.bold = True
    r0.font.name = 'Arial'
    r0.font.size = Pt(10.5)
    r1 = p.add_run(value or " ")
    r1.font.name = 'Arial'
    r1.font.size = Pt(10.5)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'dotted')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_form_table(doc, title, pairs):
    # Section title
    tp = doc.add_paragraph()
    tp.paragraph_format.space_before = Pt(10)
    tp.paragraph_format.space_after = Pt(2)
    tr = tp.add_run(title)
    tr.bold = True
    tr.font.name = 'Arial'
    tr.font.size = Pt(12)
    tr.font.color.rgb = RGBColor(0x00, 0x33, 0x33)
    # Details as clean fill-in form fields (no visible table grid)
    for k, v in pairs:
        _add_form_field(doc, k, v)


def _build_document(logo_path):
    from docx.shared import Twips, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document()
    for section in doc.sections:
        section.page_width = Twips(11906)   # A4 width
        section.page_height = Twips(16838)  # A4 height
        section.left_margin = Inches(0.87)
        section.right_margin = Inches(0.86)
        section.top_margin = Inches(1.15)
        section.bottom_margin = Inches(0.43)
        section.header_distance = Inches(0.42)
        section.footer_distance = Inches(0.42)

    if 'Normal' in doc.styles:
        doc.styles['Normal'].font.name = 'Arial'
        doc.styles['Normal'].font.size = Pt(10.5)

    # --- Header letterhead ---
    if os.path.exists(logo_path):
        hdr = doc.sections[0].header
        tbl = hdr.add_table(rows=1, cols=2, width=Inches(6.9))
        tbl.autofit = False
        tbl.allow_autofit = False
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tb = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0'); el.set(qn('w:space'), '0')
            tb.append(el)
        tbl._tbl.tblPr.append(tb)

        def set_cell_width(cell, inches):
            cell.width = Inches(inches)
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW'); tcPr.append(tcW)
            tcW.set(qn('w:w'), str(int(Inches(inches).emu / 635)))
            tcW.set(qn('w:type'), 'dxa')

        cellL = tbl.cell(0, 0)
        set_cell_width(cellL, 3.2)
        pL = cellL.paragraphs[0]
        pL.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pL.add_run().add_picture(logo_path, height=Inches(0.68))

        cellR = tbl.cell(0, 1)
        set_cell_width(cellR, 3.4)
        pR = cellR.paragraphs[0]
        pR.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        for i, (txt, bold) in enumerate([
            ("A Leading Recruitment Agency", True),
            ("For Medical Professionals", False),
            ("DMW-657-LB-10132025-R", False),
        ]):
            if i > 0:
                pR = cellR.add_paragraph()
                pR.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = pR.add_run(txt)
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
            run.font.bold = bold

        bp = hdr.add_paragraph()
        bp.paragraph_format.space_before = Pt(2)
        pPr = bp._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'dotted')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), '009999')
        pBdr.append(bottom)
        pPr.append(pBdr)

    doc.add_paragraph("")
    return doc


def generate_docx_document(profile_data, output_path, photo_bytes=None, logo_path=None):
    if not logo_path:
        logo_path = LOGO_PATH
    doc = _build_document(logo_path)

    # Override style fonts to ensure 100% uniformity
    try:
        for style_name in ['Normal', 'List Bullet']:
            if style_name in doc.styles:
                doc.styles[style_name].font.name = 'Arial'
                doc.styles[style_name].font.size = Pt(10.5)
    except Exception as e:
        print(f"Error setting default styles: {e}")

    last_name = profile_data.get("lastName", "LAST NAME")
    first_name = profile_data.get("firstName", "FIRST NAME")
    details = profile_data.get("details", {})

    # Candidate name header (all caps, centered, underlined style)
    name_para = doc.add_paragraph()
    name_para.alignment = 1  # center
    name_para.paragraph_format.space_before = Pt(0)
    name_para.paragraph_format.space_after = Pt(6)
    name_run = name_para.add_run(f"{last_name}".upper() + ", " + f"{first_name}".upper() + ", MIDDLE NAME")
    name_run.bold = True
    name_run.font.name = 'Arial'
    name_run.font.size = Pt(13)
    name_run.underline = True

    # Photo right-aligned
    if photo_bytes:
        try:
            p_img = doc.add_paragraph()
            p_img.alignment = 2  # Right
            p_img.paragraph_format.space_before = Pt(0)
            p_img.paragraph_format.space_after = Pt(8)
            run = p_img.add_run()
            run.add_picture(io.BytesIO(photo_bytes), width=Pt(95))
        except Exception as e:
            print(f"Error adding picture to document: {e}")

    # EXPERIENCE SUMMARY
    experience = details.get("experience", [])
    if not experience:
        experience = [{"employer": " ", "bedCapacity": " ", "areaExposure": " ",
                       "position": " ", "duration": " ", "duties": []}]
    for exp in experience:
        pairs = [
            ("Name of Employer", exp.get("employer", " ")),
            ("Bed Capacity", exp.get("bedCapacity", " ")),
            ("Area of Exposure", exp.get("areaExposure", " ")),
            ("Position", exp.get("position", " ")),
            ("Duration", exp.get("duration", " ")),
        ]
        _add_form_table(doc, "EXPERIENCE SUMMARY", pairs)
        # Duties and responsibilities
        dp = doc.add_paragraph()
        dp.paragraph_format.space_before = Pt(2)
        dp.paragraph_format.space_after = Pt(2)
        dr = dp.add_run("Duties and Responsibilities: ")
        dr.bold = True
        dr.font.name = 'Arial'
        dr.font.size = Pt(10.5)
        duties = exp.get("duties", [])
        for duty in duties:
            b = doc.add_paragraph(style='List Bullet')
            b.paragraph_format.space_before = Pt(0)
            b.paragraph_format.space_after = Pt(1)
            br = b.add_run(duty)
            br.font.name = 'Arial'
            br.font.size = Pt(10.5)

    # PERSONAL DETAILS
    _add_form_table(doc, "PERSONAL DETAILS", [
        ("Date of Birth", details.get("dob", " ")),
        ("Nationality", details.get("nationality", " ")),
        ("Gender", details.get("gender", " ")),
        ("Marital Status", details.get("maritalStatus", " ")),
        ("Religion", details.get("religion", " ")),
        ("Height", details.get("height", " ")),
        ("Weight", details.get("weight", " ")),
        ("BMI", details.get("bmi", " ")),
    ])

    # PASSPORT DETAILS
    _add_form_table(doc, "PASSPORT DETAILS", [
        ("Passport No.", details.get("passportNo", " ")),
        ("Place of Issue", details.get("placeOfIssue", " ")),
        ("Date of Issue", details.get("dateOfIssue", " ")),
        ("Date of Expiry", details.get("dateOfExpiry", " ")),
    ])

    # EDUCATIONAL DETAILS
    education = details.get("education", [])
    edu_pairs = []
    if education:
        edu_pairs.append(("College/University Attended", education[0].get("school", " ")))
        edu_pairs.append(("Qualification", education[0].get("qualification", " ")))
    else:
        edu_pairs.append(("College/University Attended", " "))
        edu_pairs.append(("Qualification", " "))
    edu_pairs.append(("Graduation Date", " "))
    _add_form_table(doc, "EDUCATIONAL DETAILS", edu_pairs)

    # REGISTRATION DETAILS
    reg_pairs = [
        ("Registration Authority", "Professional Regulation Commission"),
        ("Registration No.", details.get("prcRegNo", " ")),
        ("Registration Date", details.get("prcRegDate", " ")),
        ("Validity Date", details.get("prcValidity", " ")),
    ]
    reg_pairs_2 = [
        ("Registration Authority", "Saudi Commission for Health Specialties"),
        ("Registration No.", details.get("scfhsRegNo", " ")),
        ("Registration Date", details.get("scfhsRegDate", " ")),
        ("Validity Date", details.get("scfhsValidity", " ")),
    ]
    _add_form_table(doc, "REGISTRATION DETAILS", reg_pairs)
    _add_form_table(doc, "REGISTRATION DETAILS", reg_pairs_2)

    doc.save(output_path)


def _pdf_form_section(story, title, pairs):
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, HRFlowable, Spacer
    from reportlab.lib.styles import ParagraphStyle

    section_style = ParagraphStyle('sec', fontName='Helvetica-Bold', fontSize=12.5,
                                   spaceBefore=12, spaceAfter=5, textColor=colors.HexColor('#006666'))
    field_style = ParagraphStyle('fld', fontName='Helvetica', fontSize=11,
                                 leading=16, spaceBefore=2.5, spaceAfter=1, textColor=colors.black)

    story.append(Paragraph(title, section_style))
    for k, v in pairs:
        story.append(Paragraph(f"<b>{k}:</b>&nbsp;&nbsp;{(v or ' ')}", field_style))
        story.append(Spacer(1, 4))


@st.cache_data(show_spinner=False)
def _white_background_logo(logo_path=None):
    if not logo_path:
        logo_path = LOGO_PATH
    try:
        from PIL import Image
        import io as _io
        img = Image.open(logo_path).convert("RGBA")
        bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
        bg.alpha_composite(img)
        buf = _io.BytesIO()
        bg.convert("RGB").save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        with open(logo_path, "rb") as f:
            return f.read()


def generate_pdf_direct(profile_data, output_path, photo_bytes=None, logo_path=None):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.pdfgen import canvas
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER

    if not logo_path:
        logo_path = LOGO_PATH

    margin = 36  # ~0.5in
    ch = 0.68 * inch
    header_pad = 1.35 * inch

    def _draw_header(c, _doc):
        w, h = A4
        c.saveState()
        if os.path.exists(logo_path):
            try:
                from reportlab.lib.utils import ImageReader
                logo_data = _white_background_logo(logo_path)
                c.drawImage(ImageReader(io.BytesIO(logo_data)), margin, h - margin - ch,
                            width=2.0 * inch, height=ch, preserveAspectRatio=True, anchor='sw')
            except Exception:
                pass
        ty = h - margin - 14
        for txt, f in [("A Leading Recruitment Agency", 'Helvetica-Bold'),
                       ("For Medical Professionals", 'Helvetica'),
                       ("DMW-657-LB-10132025-R", 'Helvetica')]:
            c.setFont(f, 11)
            c.drawRightString(w - margin, ty, txt)
            ty -= 13
        c.restoreState()

    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            leftMargin=margin, rightMargin=margin,
                            topMargin=header_pad, bottomMargin=margin,
                            title="Curriculum Vitae", author="GCC CV Studio")

    story = []
    last_name = profile_data.get("lastName", "LAST NAME")
    first_name = profile_data.get("firstName", "FIRST NAME")
    details = profile_data.get("details", {})

    name_style = ParagraphStyle('nm', fontName='Helvetica-Bold', fontSize=13,
                                alignment=TA_CENTER, spaceAfter=8, textColor=colors.black)
    story.append(Paragraph(
        f"{last_name.upper()}, {first_name.upper()}, MIDDLE NAME", name_style))

    if photo_bytes:
        try:
            from reportlab.lib.utils import ImageReader
            img = ImageReader(io.BytesIO(photo_bytes))
            iw, ih = img.getSize()
            tw = 1.4 * inch
            th = tw * ih / iw
            pt = Table([['', Image(io.BytesIO(photo_bytes), width=tw, height=th)]],
                       colWidths=[doc.width - tw, tw])
            pt.setStyle(TableStyle([
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (0, 0), doc.width - 2.4 * 72),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))
            story.append(pt)
            story.append(Spacer(1, 4))
        except Exception:
            pass

    experience = details.get("experience", [])
    if not experience:
        experience = [{"employer": " ", "position": " ", "duration": " ", "duties": []}]
    for exp in experience:
        _pdf_form_section(story, "EXPERIENCE SUMMARY", [
            ("Name of Employer", exp.get("employer", " ")),
            ("Bed Capacity", exp.get("bedCapacity", " ")),
            ("Area of Exposure", exp.get("areaExposure", " ")),
            ("Position", exp.get("position", " ")),
            ("Duration", exp.get("duration", " ")),
        ])
        duty_label = Paragraph("Duties and Responsibilities:", ParagraphStyle(
            'dl', fontName='Helvetica-Bold', fontSize=10.5))
        story.append(duty_label)
        story.append(Spacer(1, 2))
        for duty in exp.get("duties", []):
            cleaned = re.sub(r'^[\s\u2022*\-\u2023\u25aa\u25cf]+', '', str(duty)).strip()
            if cleaned:
                story.append(Paragraph(f"\u2022 {cleaned}", ParagraphStyle(
                    'dd', fontName='Helvetica', fontSize=10, leftIndent=12)))
        story.append(Spacer(1, 4))

    _pdf_form_section(story, "PERSONAL DETAILS", [
        ("Date of Birth", details.get("dob", " ")),
        ("Nationality", details.get("nationality", " ")),
        ("Gender", details.get("gender", " ")),
        ("Marital Status", details.get("maritalStatus", " ")),
        ("Religion", details.get("religion", " ")),
        ("Height", details.get("height", " ")),
        ("Weight", details.get("weight", " ")),
        ("BMI", details.get("bmi", " ")),
    ])

    _pdf_form_section(story, "PASSPORT DETAILS", [
        ("Passport No.", details.get("passportNo", " ")),
        ("Place of Issue", details.get("placeOfIssue", " ")),
        ("Date of Issue", details.get("dateOfIssue", " ")),
        ("Date of Expiry", details.get("dateOfExpiry", " ")),
    ])

    education = details.get("education", [])
    if education:
        school = education[0].get("school", " ")
        qualification = education[0].get("qualification", " ")
    else:
        school, qualification = " ", " "
    _pdf_form_section(story, "EDUCATIONAL DETAILS", [
        ("College/University Attended", school),
        ("Qualification", qualification),
        ("Graduation Date", " "),
    ])

    _pdf_form_section(story, "REGISTRATION DETAILS", [
        ("Registration Authority", "Professional Regulation Commission"),
        ("Registration No.", details.get("prcRegNo", " ")),
        ("Registration Date", details.get("prcRegDate", " ")),
        ("Validity Date", details.get("prcValidity", " ")),
    ])
    _pdf_form_section(story, "REGISTRATION DETAILS", [
        ("Registration Authority", "Saudi Commission for Health Specialties"),
        ("Registration No.", details.get("scfhsRegNo", " ")),
        ("Registration Date", details.get("scfhsRegDate", " ")),
        ("Validity Date", details.get("scfhsValidity", " ")),
    ])

    doc.build(story, onFirstPage=_draw_header, onLaterPages=_draw_header)


def extract_name_offline(raw_text):
    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
    candidate_name = "Candidate Name"
    first_name = "Firstname"
    last_name = "Lastname"
    
    skip_keywords = {
        "cv", "resume", "curriculum vitae", "summary", "contact", "about me",
        "experience", "education", "educational", "skills", "key skills",
        "certifications", "certification", "licenses", "license", "references",
        "referees", "personal details", "details", "personal", "information",
        "passport details", "registration details", "employment history",
        "professional experience", "training", "seminar", "additional",
        "objective", "profile", "about", "gender", "civil status", "civil",
        "religion", "nationality", "marital", "birth", "dob", "date of birth",
        "height", "weight", "bmi", "passport", "registration", "validity",
        "place of issue", "date of issue", "date of expiry", "expiry",
        "college", "university", "school", "degree", "qualification",
        "employer", "position", "duration", "duties", "bed capacity",
        "area of exposure", "name of employer", "job description",
        "duties and responsibilities", "general",
        "name", "candidate", "referee", "credit", "character",
        "tel", "telephone", "mobile", "email", "website", "linkedin",
    }
    address_keywords = {
        "city", "philippines", "street", "road", "province", "zip",
        "address", "brgy", "barangay", "block", "lot", "zone",
        "region", "country", "location",
    }
    
    for line in lines[:8]:
        lower_line = line.lower()
        if '@' in line or 'mailto' in line or 'http' in line or 'www' in line:
            continue
        if any(char.isdigit() for char in line) and len(line) < 16:
            continue
        if len(line) > 40:
            continue
        if ':' in line or ' : ' in line:
            continue
        if lower_line in skip_keywords:
            continue
        words = lower_line.replace(',', ' ').replace('.', ' ').split()
        if words and all(w in skip_keywords or len(w) <= 1 for w in words):
            continue
        if any(kw in lower_line for kw in address_keywords):
            continue
        if line.isupper() and len(line) < 30 and ',' not in line:
            continue
        if any(lower_line.endswith(s) for s in [":", ".", ": ", ". ", "name", "candidate"]):
            continue
        name_tokens = re.sub(r'[.,]', ' ', line).split()
        alpha_count = sum(1 for t in name_tokens if t.isalpha())
        if alpha_count < 2:
            continue
        
        candidate_name = line
        break
        
    # Split name into first and last
    if ',' in candidate_name:
        parts = candidate_name.split(',', 1)
        last_name = parts[0].strip()
        first_name = parts[1].strip()
    else:
        parts = candidate_name.split()
        if len(parts) >= 2:
            first_name = parts[0]
            last_name = " ".join(parts[1:])
        elif len(parts) == 1:
            first_name = parts[0]
            last_name = ""
        
    full_name_upper = f"{last_name.upper()}, {first_name.upper()}" if last_name else first_name.upper()
    return full_name_upper, last_name, first_name

def calculate_years_exp_offline(raw_text):
    from datetime import datetime
    current_year = datetime.now().year
    
    # Find all patterns of YYYY - YYYY or YYYY - Present
    pattern = r'\b(19\d{2}|20\d{2})\s*[-–—to\s]+\s*(Present|present|Current|current|20\d{2}|19\d{2})\b'
    matches = re.findall(pattern, raw_text)
    
    ranges = []
    
    for start, end in matches:
        start_yr = int(start)
        if end.lower() in ['present', 'current']:
            end_yr = current_year
        else:
            end_yr = int(end)
            
        if 0 <= (end_yr - start_yr) <= 40:
            ranges.append((start_yr, end_yr))
            
    if not ranges:
        years = [int(y) for y in re.findall(r'\b(19\d{2}|20\d{2})\b', raw_text)]
        if years:
            years = [y for y in years if y <= current_year]
            if len(years) >= 2:
                span = max(years) - min(years)
                est = max(0, span - 4) if span > 4 else span
                return min(25, max(1, est))
            return 1
        return 0
        
    ranges.sort(key=lambda x: x[0])
    merged_ranges = []
    for r in ranges:
        if not merged_ranges or merged_ranges[-1][1] < r[0]:
            merged_ranges.append(list(r))
        else:
            merged_ranges[-1][1] = max(merged_ranges[-1][1], r[1])
            
    total_years = sum(end - start for start, end in merged_ranges)
    return max(1, total_years)

def structure_cv_offline(raw_text):
    lines = raw_text.split('\n')
    sections = {
        "Contact Information": [],
        "Professional Summary": [],
        "Work Experience": [],
        "Education": [],
        "Key Skills": [],
        "Certifications & Licenses": []
    }
    
    current_section = "Contact Information"
    
    keywords = {
        "work": "Work Experience",
        "experience": "Work Experience",
        "employment": "Work Experience",
        "history": "Work Experience",
        "job": "Work Experience",
        "education": "Education",
        "academic": "Education",
        "university": "Education",
        "school": "Education",
        "degree": "Education",
        "skills": "Key Skills",
        "expertise": "Key Skills",
        "abilities": "Key Skills",
        "competencies": "Key Skills",
        "summary": "Professional Summary",
        "objective": "Professional Summary",
        "profile": "Professional Summary",
        "about": "Professional Summary",
        "certification": "Certifications & Licenses",
        "license": "Certifications & Licenses",
        "credential": "Certifications & Licenses",
        "contact": "Contact Information",
        "phone": "Contact Information",
        "email": "Contact Information",
        "address": "Contact Information"
    }
    
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue
            
        is_header = False
        lower_line = line_stripped.lower()
        
        if len(line_stripped) < 40:
            for kw, sec_name in keywords.items():
                if kw in lower_line:
                    if any(phrase in lower_line for phrase in [
                        "work experience", "employment", "history", "education", 
                        "skills", "summary", "objective", "certifications", 
                        "licenses", "contact info", "about me", "professional summary"
                    ]) or line_stripped.isupper() or len(line_stripped.split()) < 4:
                        current_section = sec_name
                        is_header = True
                        break
            if is_header:
                continue
                
        sections[current_section].append(line_stripped)
        
    markdown_parts = []
    
    markdown_parts.append("## Contact Information")
    if sections["Contact Information"]:
        for line in sections["Contact Information"]:
            markdown_parts.append(line)
    else:
        emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', raw_text)
        phones = re.findall(r'\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b', raw_text)
        if emails:
            markdown_parts.append(f"**Email:** {emails[0]}")
        if phones:
            markdown_parts.append(f"**Phone:** {phones[0]}")
    markdown_parts.append("")
    
    markdown_parts.append("## Professional Summary")
    if sections["Professional Summary"]:
        for line in sections["Professional Summary"]:
            markdown_parts.append(line)
    else:
        markdown_parts.append("Results-driven professional with experience in the industry. Possesses strong analytical, technical, and communication skills to drive project success and organizational growth.")
    markdown_parts.append("")
    
    markdown_parts.append("## Work Experience")
    if sections["Work Experience"]:
        for line in sections["Work Experience"]:
            if line.startswith("-") or line.startswith("*"):
                markdown_parts.append(line)
            else:
                if any(yr in line for yr in ["201", "202", "199", "Present"]):
                    markdown_parts.append(f"\n**{line}**")
                else:
                    markdown_parts.append(line)
    else:
        markdown_parts.append("Refer to raw CV for work history details.")
    markdown_parts.append("")
    
    markdown_parts.append("## Education")
    if sections["Education"]:
        for line in sections["Education"]:
            markdown_parts.append(line)
    else:
        markdown_parts.append("Refer to raw CV for educational history.")
    markdown_parts.append("")
    
    markdown_parts.append("## Key Skills")
    if sections["Key Skills"]:
        for line in sections["Key Skills"]:
            if ',' in line and len(line) > 15:
                for skill in line.split(','):
                    if skill.strip():
                        markdown_parts.append(f"- {skill.strip()}")
            else:
                markdown_parts.append(line if (line.startswith("-") or line.startswith("*")) else f"- {line}")
    else:
        markdown_parts.append("Refer to raw CV for full list of skills.")
    markdown_parts.append("")
    
    markdown_parts.append("## Certifications & Licenses")
    if sections["Certifications & Licenses"]:
        for line in sections["Certifications & Licenses"]:
            markdown_parts.append(line if (line.startswith("-") or line.startswith("*")) else f"- {line}")
    else:
        markdown_parts.append("Refer to raw CV for certifications.")
        
    return "\n".join(markdown_parts)

def generate_summary_offline(raw_text, years_exp):
    titles = ["Nurse", "Engineer", "Developer", "Accountant", "Manager", "Teacher", "Administrator", "Technician", "Analyst", "Therapist", "Specialist"]
    detected_title = "Professional"
    lower_text = raw_text.lower()
    
    for title in titles:
        if title.lower() in lower_text:
            detected_title = title
            break
            
    s1 = f"Results-oriented {detected_title} with over {years_exp} years of dedicated work experience."
    s2 = f"Demonstrates a proven track record of handling key responsibilities, optimization, and collaboration within professional settings."
    s3 = "Committed to delivering high-quality service and maintaining operational excellence in demanding corporate environments."
    
    return f"{s1} {s2} {s3}"

def extract_medical_details(raw_text):
    lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
    text = raw_text
    d = {
        "dob": " ", "nationality": " ", "gender": " ", "maritalStatus": " ",
        "religion": " ", "height": " ", "weight": " ", "bmi": " ",
        "passportNo": " ", "placeOfIssue": " ", "dateOfIssue": " ", "dateOfExpiry": " ",
        "education": [], "experience": [],
        "prcRegNo": " ", "prcRegDate": " ", "prcValidity": " ",
        "scfhsRegNo": " ", "scfhsRegDate": " ", "scfhsValidity": " ",
    }

    def field(pats, key, case_sensitive=False):
        for pat in pats:
            raw = re.search(pat, text, re.IGNORECASE if not case_sensitive else 0)
            if raw:
                value = (raw.group(1) or "").strip()
                value = re.split(r'\s{2,}', value)[0].strip()
                d[key] = value if value else " "
                return
    field([r'(?:Date of Birth|DOB|Birthdate)\s*:?\s*([^\n]{1,40})'], "dob")
    field([r'Nationality\s*:?\s*([^\n]{1,40})'], "nationality")
    field([r'Gender\s*:?\s*([^\n]{1,60})', r'\bSex\s*:?\s*([^\n]{1,60})'], "gender")
    field([r'Marital Status\s*:?\s*([^\n]{1,60})', r'Civil Status\s*:?\s*([^\n]{1,60})'], "maritalStatus")
    field([r'Religion\s*:?\s*([^\n]{1,60})'], "religion")
    field([r'Height\s*:?\s*([^\n,;]{1,20})'], "height")
    field([r'Weight\s*:?\s*([^\n,;]{1,20})'], "weight")
    field([r'BMI\s*:?\s*([^\n,;]{1,20})'], "bmi")
    field([r'Passport No\.?\s*:?\s*([^\n,;]{1,40})', r'Passport Number\s*:?\s*(P\d*[A-Z0-9]*)', r'Passport\s*:?\s*([A-Z0-9]{6,10})'], "passportNo")
    field([r'Place of Issue\s*:?\s*([^\n,;]{1,40})', r'Issued at\s*:?\s*([^\n,;]{1,40})'], "placeOfIssue")
    field([r'Date of Issue\s*:?\s*([^\n,;]{1,40})'], "dateOfIssue")
    field([r'Date of Expiry\s*:?\s*([^\n,;]{1,40})', r'Expiry\s*:?\s*([^\n,;]{1,40})'], "dateOfExpiry")
    field([r'(?:Registration\s*)?Registration\s*No\.?\s*:?\s*([^\n,;]{1,40})'], "prcRegNo")
    field([r'Registration\s*Date\s*:?\s*([^\n,;]{1,40})'], "prcRegDate")
    field([r'Validity\s*Date\s*:?\s*([^\n,;]{1,40})'], "prcValidity")

    # Education (school + degree)
    edu_pats = [
        r'(Bachelor(?: of)?[^\n]{2,60})',
        r'(BS\s+Nursing[^\n]{0,40})',
    ]
    matched_edu = []
    for pat in edu_pats:
        for m in re.finditer(pat, text):
            val = m.group(1).strip()
            if val and val not in matched_edu:
                matched_edu.append(val)
            if len(matched_edu) >= 2:
                break
        if len(matched_edu) >= 2:
            break
    school = " "
    for l in lines:
        if re.search(r'(University|College)\b', l, re.IGNORECASE) and len(l) < 80 and "http" not in l.lower():
            school = re.sub(r'^[\s\u2022*\-\u2023\u25aa\u25cf]+', '', l).strip()
            break
    if matched_edu:
        d["education"].append({"school": school, "qualification": matched_edu[0]})
    elif school != " ":
        d["education"].append({"school": school, "qualification": " "})

    # Work experience blocks (template-style: Name of Employer / Bed Capacity /
    # Area of Exposure / Position / Duration / Duties and Responsibilities)
    exp_headers = ("EXPERIENCE SUMMARY", "WORK EXPERIENCE", "EMPLOYMENT HISTORY",
                   "PROFESSIONAL EXPERIENCE", "PERSONAL DETAILS", "PASSPORT DETAILS",
                   "EDUCATION", "EDUCATIONAL", "REGISTRATION", "REFEREES", "SKILLS",
                   "CERTIFICATION", "LICENSES", "TRAINING", "SEMINAR", "ADDITIONAL")
    section_re = re.compile(
        r'^\s*(?:[\u2022*\-\u2023\u25aa\u25cf])?\s*'
        r'(EXPERIENCE SUMMARY|WORK EXPERIENCE|EMPLOYMENT HISTORY|PROFESSIONAL EXPERIENCE|'
        r'PERSONAL DETAILS|PASSPORT DETAILS|EDUCATIONAL BACKGROUND|EDUCATION|REGISTRATION DETAILS|'
        r'REFERENCES|REFEREES|KEY SKILLS|SKILLS|CERTIFICATIONS\s*[&/]?\s*LICENSES|LICENSES|'
        r'ADDITIONAL INFORMATION|TRAININGS?|SEMINARS?)\s*:?\s*$', re.IGNORECASE)
    exp_year = re.compile(r'(?:19\d{2}|20\d{2})\s*(?:-|–|—|to)\s*(?:(?:19\d{2}|20\d{2})|present|current)', re.IGNORECASE)
    exp_key = re.compile(
        r'(?i)^(Name of Employer|Bed Capacity|Area of Exposure|Position Held?|Duration|'
        r'Duties and Responsibilities|Job Description|General)\s*:?\s*(.*)$')
    exp_label = re.compile(r'^[\s\u2022*\-\u2023\u25aa\u25cf]+')

    blocks = []
    in_exp = False
    cur = None
    duty_on = False
    duties = []
    for l in lines:
        s = exp_label.sub('', l).strip()
        up = s.upper()
        if section_re.match(s) or (len(s) < 45 and up.isupper() and any(h in up for h in exp_headers)):
            if cur is not None and (cur.get("duration") or cur.get("employer")):
                cur["duties"] = duties[:10]
                blocks.append(cur)
            cur = None
            duties = []
            duty_on = False
            in_exp = bool(re.search(r'(EXPERIENCE|EMPLOYMENT|PROFESSIONAL EXPERIENCE)', up))
            continue
        if not in_exp:
            continue
        km = exp_key.match(s)
        if km:
            k, v = (km.group(1).strip(), km.group(2).strip())
            kl = k.lower().replace(" ", "")
            if kl.startswith("nameofemployer"):
                if cur is not None and (cur.get("employer") or cur.get("duration")):
                    cur["duties"] = duties[:10]
                    blocks.append(cur)
                cur = {"employer": v or " "}
                duty_on = False
                duties = []
            elif kl.startswith("bedcapacity"):
                cur["bedCapacity"] = v or " "
            elif kl.startswith("areaofexposure"):
                cur["areaExposure"] = v or " "
            elif kl.startswith("position"):
                cur["position"] = v or " "
            elif kl.startswith("duration"):
                cur["duration"] = v or " "
            elif kl.startswith("dutiesandresponsibilities") or kl.startswith("jobdescription") or kl.startswith("general"):
                duty_on = True
            continue
        if cur is None:
            cur = {}
        if duty_on and s and len(s) > 2:
            duties.append(s)
    if cur is not None and (cur.get("duration") or cur.get("employer")):
        cur["duties"] = duties[:10]
        blocks.append(cur)
    if blocks:
        d["experience"] = blocks
    return d

def offline_parse_cv(raw_text):
    full_name, last_name, first_name = extract_name_offline(raw_text)
    years_exp = calculate_years_exp_offline(raw_text)
    exec_summary = generate_summary_offline(raw_text, years_exp)
    cv_markdown = structure_cv_offline(raw_text)
    details = extract_medical_details(raw_text)
    
    return {
        "fullName": full_name,
        "lastName": last_name,
        "firstName": first_name,
        "yearsOfExperience": years_exp,
        "executiveSummary": exec_summary,
        "auditedCvMarkdown": cv_markdown,
        "details": details
    }


@st.cache_data(show_spinner=False)
def build_export_zip(items: tuple) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in items:
            zf.writestr(name, data)
    return buf.getvalue()

# =========================================================
# VIEW DISPATCH
# =========================================================
_VIEW = st.session_state.view

if _VIEW == "Dashboard":
    # --- STATS ROW (full width) ---
    _total = st.session_state.processed_count
    summary_exists = os.path.exists(SUMMARY_DIR)
    gcc_exists = os.path.exists(GCC_TEMPLATE_PATH)
    _dot2 = "" if gcc_exists else " inactive"
    st.markdown(f"""<div class='stats-row'>
      <div class='stats-card'>
        <div class='label'>Processed This Session</div>
        <div class='number'>{_total} <span>CVs</span></div>
      </div>
      <div class='status-card'>
        <div class='status-item'><span class='dot'></span><span class='label'>Local Folder Status</span><span class='value'>Candidate CV Summary <span class='highlight'>{'active' if summary_exists else 'missing'}</span></span></div>
        <div class='status-item'><span class='dot{_dot2}'></span><span class='label'>CV Format</span><span class='value'><span class='highlight'>GCC_CV_FORMAT.doc</span> {'loaded' if gcc_exists else 'missing'}</span></div>
      </div>
    </div>""", unsafe_allow_html=True)

    # Uploads now live in the sidebar (the control panel)
    uploaded_files = st.session_state.get(cv_uploader_key)
    uploaded_photo = None
    if uploaded_files and len(uploaded_files) == 1:
        uploaded_photo = st.session_state.get(f"photo_uploader_{st.session_state.uploader_key}")

    # --- 70/30 MULTI-PANEL: content vs tips/stats ---
    _main_col, _side_col = st.columns([7, 3], gap="large")
    _last = st.session_state.get("last_parse", "—")

    with _side_col:
        st.markdown("""<div class='tip-panel'>
          <div class='tp-title'><span class='m'>tips_and_updates</span> How it works</div>
          <div class='tip-item'><span class='m'>upload_file</span><div>Upload raw CVs in the sidebar — PDF, DOCX, DOC or TXT.</div></div>
          <div class='tip-item'><span class='m'>auto_fix_high</span><div>Hit <strong>Standardize</strong> to parse and convert them to the GCC standard.</div></div>
          <div class='tip-item'><span class='m'>manage_search</span><div>Inspect the structured profile side-by-side with the raw text.</div></div>
          <div class='tip-item'><span class='m'>file_download</span><div>Export each candidate as DOCX and PDF, or bundle them as a ZIP.</div></div>
        </div>""", unsafe_allow_html=True)

        st.markdown(f"""<div class='tip-panel'>
          <div class='tp-title'><span class='m'>insights</span> Session stats</div>
          <div style='font-size:15px;font-weight:700;color:var(--text-dark);'>Last parse: {_last}</div>
          <div style='font-size:13px;color:var(--text-muted);margin-top:6px;'>{len(uploaded_files or [])} file(s) currently in queue.</div>
          <div style='margin-top:10px;'><span class='pill ok'><span class='m' style='font-size:0.8em'>verified_user</span> 100% offline</span></div>
        </div>""", unsafe_allow_html=True)

        st.markdown("""<div class='tip-panel'>
          <div class='tp-title'><span class='m'>bolt</span> Pro tip</div>
          <div style='font-size:13px;color:var(--text-muted);line-height:1.55;'>
            For a single CV, optionally upload a clean profile photo in the sidebar so the standardized document uses it directly.
          </div>
        </div>""", unsafe_allow_html=True)

    with _main_col:
        # --- UPLOAD QUEUE / PREVIEW (still in sidebar, but queue previewed here) ---
        if uploaded_files:
            st.markdown("<div class='page-section'><span class='m m-soft'>upload_file</span> Upload Queue</div>", unsafe_allow_html=True)
            st.markdown("<div class='workflow'>"
                        "<span class='wf-label'>Pipeline</span>"
                        "<span class='wf-step done'>Upload</span><span class='wf-arrow'>→</span>"
                        "<span class='wf-step done'>Parse</span><span class='wf-arrow'>→</span>"
                        "<span class='wf-step'>Validate</span><span class='wf-arrow'>→</span>"
                        "<span class='wf-step'>Standardize</span><span class='wf-arrow'>→</span>"
                        "<span class='wf-step'>Generate</span><span class='wf-arrow'>→</span>"
                        "<span class='wf-step'>Export</span>"
                        "</div>", unsafe_allow_html=True)
            st.markdown(f"<div style='font-size:0.9rem;font-weight:600;color:var(--text-dark);margin-bottom:0.6rem;'><span style='color:var(--brand);'>{len(uploaded_files)}</span> file(s) ready to process</div>", unsafe_allow_html=True)

            ficon_map = {"pdf": "picture_as_pdf", "docx": "description", "doc": "article", "txt": "text_snippet"}
            for _uf in uploaded_files:
                _ext = _uf.name.split(".")[-1].lower()
                _bcls = {"pdf": "badge-pdf", "docx": "badge-docx", "doc": "badge-doc", "txt": "badge-txt"}.get(_ext, "badge-txt")
                _size_kb = _uf.size / 1024 if _uf.size else 0
                _size_str = f"{_size_kb:.1f} KB" if _size_kb < 1024 else f"{_size_kb/1024:.2f} MB"
                _fic = ficon_map.get(_ext, "description")
                st.markdown(
                    f"<div class='file-row fade-in'>"
                    f"<div class='f-icon'><span class='m'>{_fic}</span></div>"
                    f"<div style='flex:1;'><div class='f-name'>{_uf.name}</div><div class='f-size'>{_ext.upper()} · {_size_str}</div></div>"
                    f"<div class='f-status'><span class='format-badge {_bcls}'>{_ext.upper()}</span></div>"
                    f"</div>", unsafe_allow_html=True)

            st.markdown("<div class='privacy-note' style='margin:0.9rem 0;'><span class='m'>verified_user</span> Local parser is used for all CVs. No external AI or internet connection required.</div>", unsafe_allow_html=True)

        # --- PROCESSING via st.status ---
        if uploaded_files and process_btn:
            st.session_state.batch_results = {}
            fname_badge = {"pdf": "badge-pdf", "docx": "badge-docx", "doc": "badge-doc", "txt": "badge-txt"}
            with st.status("Standardizing CVs...", expanded=True) as status:
                for idx, file in enumerate(uploaded_files):
                    file_name = file.name
                    file_ext = file_name.split(".")[-1].lower()
                    file_bytes = file.read()
                    file.seek(0)
                    _bc = fname_badge.get(file_ext, "badge-txt")

                    try:
                        if file_ext == "docx":
                            raw_text = parse_docx(file_bytes)
                        elif file_ext == "doc":
                            raw_text, doc_docx_bytes = parse_doc(file_bytes, file_name)
                            st.session_state.doc_docx_bytes = doc_docx_bytes
                        elif file_ext == "pdf":
                            raw_text = parse_pdf(file_bytes)
                        else:
                            raw_text = parse_txt(file_bytes)

                        profile_data = offline_parse_cv(raw_text)
                        photo_bytes = None
                        if len(uploaded_files) == 1 and uploaded_photo is not None:
                            photo_bytes = uploaded_photo.read()
                        else:
                            if file_ext == "doc" and 'doc_docx_bytes' in locals():
                                photo_bytes = extract_picture(doc_docx_bytes, "docx")
                            else:
                                photo_bytes = extract_picture(file_bytes, file_ext)

                        last_name = profile_data.get("lastName", "Lastname")
                        first_name = profile_data.get("firstName", "Firstname")

                        import tempfile
                        docx_bytes = None
                        with tempfile.TemporaryDirectory() as temp_dir:
                            temp_docx_out = os.path.join(temp_dir, f"{last_name}_{first_name}_Standardized.docx")
                            generate_docx_document(profile_data, temp_docx_out, photo_bytes=photo_bytes)
                            with open(temp_docx_out, "rb") as doc_f:
                                docx_bytes = doc_f.read()

                        years_exp = profile_data.get("yearsOfExperience", 0)
                        exec_summary = profile_data.get("executiveSummary", "")
                        profile_filename = f"{last_name}, {first_name}.txt"
                        profile_filepath = os.path.join(SUMMARY_DIR, profile_filename)
                        profile_content = (
                            f"Candidate: {last_name}, {first_name}\n"
                            f"Total Experience: {years_exp} Years\n"
                            f"Summary: {exec_summary}\n"
                        )
                        with open(profile_filepath, "w", encoding="utf-8") as f:
                            f.write(profile_content)

                        pdf_bytes = None
                        st.write(f"Generating PDF for `{file_name}`...")
                        try:
                            import tempfile as _tf
                            with _tf.TemporaryDirectory() as _dir:
                                _pdf_path = os.path.join(_dir, f"{last_name}_{first_name}_Standardized.pdf")
                                generate_pdf_direct(profile_data, _pdf_path, photo_bytes=photo_bytes)
                                with open(_pdf_path, "rb") as pf:
                                    pdf_bytes = pf.read()
                        except Exception as pdf_err:
                            print(f"PDF generation error: {pdf_err}")
                            pdf_bytes = None

                        st.session_state.batch_results[file_name] = {
                            "profile_data": profile_data,
                            "raw_text": raw_text,
                            "photo_bytes": photo_bytes,
                            "docx_bytes": docx_bytes,
                            "pdf_bytes": pdf_bytes,
                            "status": "Success"
                        }
                    except Exception as e:
                        st.session_state.batch_results[file_name] = {
                            "status": "Error",
                            "error_message": str(e)
                        }

                st.session_state.processed_count += len(uploaded_files)
                _okn = sum(1 for r in st.session_state.batch_results.values() if r.get("status") == "Success")
                _errn = sum(1 for r in st.session_state.batch_results.values() if r.get("status") == "Error")
                if _errn:
                    status.update(label=f"{_okn} CV(s) processed. {_errn} failed.", state="error", expanded=False)
                    st.toast(f"{_okn} CV(s) processed. {_errn} failed.", icon="⚠️")
                else:
                    status.update(label=f"{_okn} CV(s) processed successfully.", state="complete", expanded=False)
                    st.toast(f"{_okn} CV(s) processed successfully.", icon="✅")
            st.session_state.last_parse = datetime.now().strftime("%H:%M")
            st.rerun()

        # --- RESULTS / PREVIEWS ---
        if "batch_results" in st.session_state and len(st.session_state.batch_results) > 0:
            # Filter successful candidates
            success_candidates = {
                filename: f"{res['profile_data'].get('lastName', 'Lastname')}, {res['profile_data'].get('firstName', 'Firstname')} ({filename})"
                for filename, res in st.session_state.batch_results.items()
                if res["status"] == "Success"
            }

            if success_candidates:
                is_batch = len(success_candidates) > 1

                # Batch summary & exports (only shown if multiple files processed)
                if is_batch:
                    success_count = len(success_candidates)
                    error_count = sum(1 for res in st.session_state.batch_results.values() if res["status"] == "Error")

                    m_col1, m_col2, m_col3 = st.columns(3)
                    with m_col1:
                        st.markdown(f"<div class='stat-card'><div class='stat-number'>{len(st.session_state.batch_results)}</div><div class='stat-label-sm'>Total Uploaded</div></div>", unsafe_allow_html=True)
                    with m_col2:
                        st.markdown(f"<div class='stat-card' style='color:var(--green)'><div class='stat-number' style='color:var(--green)'>{success_count}</div><div class='stat-label-sm'>Successfully Processed</div></div>", unsafe_allow_html=True)
                    with m_col3:
                        st.markdown(f"<div class='stat-card' style='color:var(--red)'><div class='stat-number' style='color:var(--red)'>{error_count}</div><div class='stat-label-sm'>Errors / Failures</div></div>", unsafe_allow_html=True)

                    # Zip DOCX files
                    docx_zip_items = tuple(
                        (f"{res['profile_data'].get('lastName', 'Lastname')}, {res['profile_data'].get('firstName', 'Firstname')} - CV.docx",
                         res["docx_bytes"])
                        for filename, res in st.session_state.batch_results.items()
                        if res["status"] == "Success"
                    )
                    docx_zip_bytes = build_export_zip(docx_zip_items) if docx_zip_items else None

                    # Zip PDF files
                    pdf_zip_bytes = None
                    pdf_zip_items = tuple(
                        (f"{res['profile_data'].get('lastName', 'Lastname')}, {res['profile_data'].get('firstName', 'Firstname')} - CV.pdf",
                         res["pdf_bytes"])
                        for filename, res in st.session_state.batch_results.items()
                        if res["status"] == "Success" and res["pdf_bytes"] is not None
                    )
                    if pdf_zip_items:
                        pdf_zip_bytes = build_export_zip(pdf_zip_items)
                    compiled_pdf_count = len(pdf_zip_items)

                    st.markdown("<div class='page-section'><span class='m m-soft'>inventory_2</span> Batch Exports</div>", unsafe_allow_html=True)
                    b_col1, b_col2 = st.columns(2)
                    with b_col1:
                        st.download_button(
                            label=f"Download All Standardized DOCX ({success_count} files as ZIP)",
                            data=docx_zip_bytes,
                            file_name="GCC_Standardized_CVs_Word.zip",
                            mime="application/zip",
                            key="dl_docx_zip",
                            use_container_width=True,
                        )
                    with b_col2:
                        if compiled_pdf_count > 0:
                            st.download_button(
                                label=f"Download All Standardized PDFs ({compiled_pdf_count} files as ZIP)",
                                data=pdf_zip_bytes,
                                file_name="GCC_Standardized_CVs_PDF.zip",
                                mime="application/zip",
                                key="dl_pdf_zip",
                                use_container_width=True,
                            )
                        else:
                            st.info("Compile PDFs for individual candidates below, then download them as a batch ZIP here.")

                    st.markdown("<div class='page-section'><span class='m m-soft'>manage_search</span> Candidate Inspection & Single Export</div>", unsafe_allow_html=True)
                    if "selected_candidate_idx" not in st.session_state:
                        st.session_state.selected_candidate_idx = 0
                    cand_keys = list(success_candidates.keys())
                    sel_idx_holder = st.session_state.selected_candidate_idx
                    if sel_idx_holder >= len(cand_keys):
                        sel_idx_holder = 0
                    cc_cols = st.columns(min(len(cand_keys), 3))
                    for i, cf in enumerate(cand_keys):
                        _res = st.session_state.batch_results[cf]
                        _pd = _res.get("profile_data", {})
                        _ln = _pd.get("lastName", "Lastname")
                        _fn = _pd.get("firstName", "Firstname")
                        _ye = _pd.get("yearsOfExperience", 0)
                        _init = (_fn[0] if _fn else "?") + (_ln[0] if _ln else "?")
                        _active = " active" if i == sel_idx_holder else ""
                        _c = cc_cols[i % len(cc_cols)]
                        with _c:
                            st.button(
                                f"👤 {_ln}, {_fn} — {_ye} yrs",
                                key=f"cand_{i}",
                                on_click=lambda j=i: setattr(st.session_state, "selected_candidate_idx", j),
                                use_container_width=True,
                                help=cf
                            )
                            st.markdown(
                                f"<div class='candidate-card{_active}'>"
                                f"<div class='candidate-avatar'>{_init}</div>"
                                f"<div class='candidate-info'>"
                                f"<div class='candidate-name'>{_ln}, {_fn}</div>"
                                f"<div class='candidate-meta'>🗂 {cf}</div>"
                                f"</div>"
                                f"</div>", unsafe_allow_html=True)
                    selected_filename = cand_keys[st.session_state.selected_candidate_idx]
                else:
                    selected_filename = list(success_candidates.keys())[0]
                    st.session_state.selected_candidate_idx = 0

                # Unpack details for selected candidate
                candidate_result = st.session_state.batch_results[selected_filename]
                data = candidate_result["profile_data"]
                raw_text = candidate_result["raw_text"]
                photo_bytes = candidate_result["photo_bytes"]
                docx_bytes = candidate_result["docx_bytes"]
                pdf_bytes = candidate_result["pdf_bytes"]

                last_name = data.get("lastName", "Lastname")
                first_name = data.get("firstName", "Firstname")
                full_name = data.get("fullName", f"{last_name.upper()}, {first_name.upper()}")
                years_exp = data.get("yearsOfExperience", 0)
                exec_summary = data.get("executiveSummary", "")
                cv_markdown = data.get("auditedCvMarkdown", "")

                profile_filename = f"{last_name}, {first_name}.txt"
                st.info(f"Dynamic Profile saved locally: `{os.path.join('Candidate CV Summary', profile_filename)}`")

                details = data.get("details", {})

                chip_items = []
                for _k in ["nationality", "gender", "religion", "maritalStatus", "height", "weight"]:
                    _v = details.get(_k)
                    if _v and str(_v).strip():
                        chip_items.append((_k.replace("maritalStatus", "Marital Status").replace("_", " ").title(), str(_v).strip()))
                if details.get("passportNo"):
                    chip_items.append(("Passport No.", str(details["passportNo"]).strip()))

                chip_html = "".join(
                    f"<span class='chip chip-primary'>{label}: {val}</span>" for label, val in chip_items
                ) if chip_items else "<span class='chip chip-secondary'>No personal details detected</span>"

                st.markdown("<div class='fade-in'>", unsafe_allow_html=True)
                st.markdown(f"""<div class='profile-hero'>
                  <div style='display:flex;align-items:center;gap:1.5rem;flex-wrap:wrap;'>
                    <div class='candidate-avatar' style='width:80px;height:80px;font-size:2rem;'>{first_name[0] if first_name else '?'}{last_name[0] if last_name else '?'}</div>
                    <div style='flex:1;min-width:220px;'>
                      <div class='profile-hero-name'>{full_name}</div>
                      <div class='profile-hero-sub'><span class='m m-soft' style='vertical-align:text-bottom;'>workspace_premium</span> {years_exp} Years of Experience</div>
                      <div class='profile-hero-summary'>{exec_summary}</div>
                      <div class='chip-group'>{chip_html}</div>
                    </div>
                  </div>
                </div>""", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

                # --- SPLIT SCREEN LAYOUT ---
                col1, col2 = st.columns(2)

                with col1:
                    st.markdown("<div class='page-section'><span class='m m-soft'>subject</span> Raw Ingested Text</div>", unsafe_allow_html=True)
                    st.text_area("Original CV Content", raw_text, height=450, disabled=True, key=f"raw_{selected_filename}")

                with col2:
                    st.markdown("<div class='page-section'><span class='m m-soft'>task_alt</span> Structured CV Profile (Local Parser)</div>", unsafe_allow_html=True)
                    if photo_bytes:
                        st.image(photo_bytes, width=120)
                    st.markdown(f"**Name:** {full_name}  \n**Total Experience:** {years_exp} Years")
                    st.markdown(f"**Executive Summary:** {exec_summary}")
                    st.markdown("---")
                    st.markdown(cv_markdown)

                # --- DUAL EXPORT SYSTEM CARD ---
                st.markdown("<div class='page-section'><span class='m m-soft'>file_download</span> Export Options</div>", unsafe_allow_html=True)
                st.markdown("<div class='export-panel'>", unsafe_allow_html=True)
                st.markdown("<div class='export-panel-title'>Download standardized documents for this candidate</div>", unsafe_allow_html=True)

                exp_col1, exp_col2 = st.columns(2)

                with exp_col1:
                    st.download_button(
                        label="Download Standardized Word (DOCX)",
                        data=docx_bytes,
                        file_name=f"{last_name}, {first_name} - CV.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_docx_{selected_filename}",
                        use_container_width=True,
                    )

                with exp_col2:
                    if pdf_bytes is not None:
                        st.download_button(
                            label="Download Standardized PDF",
                            data=pdf_bytes,
                            file_name=f"{last_name}, {first_name} - CV.pdf",
                            mime="application/pdf",
                            key=f"dl_pdf_{selected_filename}",
                            use_container_width=True,
                        )
                    else:
                        if st.button("Generate PDF Export", key=f"gen_pdf_{selected_filename}", help="Generate and download the CV as a PDF", use_container_width=True):
                            with st.status("Generating PDF...", expanded=True) as st_pdf:
                                try:
                                    import tempfile as _tf
                                    with _tf.TemporaryDirectory() as _dir:
                                        _pdf_path = os.path.join(_dir, f"{last_name}_{first_name}_Standardized.pdf")
                                        _st.write(f"Compiling PDF for `{full_name}`...")
                                        generate_pdf_direct(data, _pdf_path, photo_bytes=photo_bytes)
                                        with open(_pdf_path, "rb") as pf:
                                            compiled_pdf = pf.read()
                                    if compiled_pdf:
                                        st.session_state.batch_results[selected_filename]["pdf_bytes"] = compiled_pdf
                                        st_pdf.update(label="PDF ready", state="complete", expanded=False)
                                        st.rerun()
                                    else:
                                        st_pdf.update(label="Failed to generate PDF", state="error")
                                        st.error("Failed to generate PDF document.")
                                except Exception as e:
                                    st_pdf.update(label="PDF generation error", state="error")
                                    st.error(f"Error generating PDF: {e}")
                st.markdown("</div>", unsafe_allow_html=True)

            # Display failures if any
            errors = [
                (filename, res["error_message"])
                for filename, res in st.session_state.batch_results.items()
                if res["status"] == "Error"
            ]
            if errors:
                st.markdown("<div class='page-section'><span class='m' style='color:var(--red)'>warning</span> Processing Failures</div>", unsafe_allow_html=True)
                for filename, err in errors:
                    st.error(f"**{filename}**: {err}")
        elif not uploaded_files:
            # --- DASHBOARD EMPTY STATE (nothing uploaded yet) ---
            st.markdown("<div class='state-block fade-in' style='margin-top:1.5rem;'>"
                        "<div class='state-icon'><span class='m'>upload_file</span></div>"
                        "<div class='state-title'>Ready to Process</div>"
                        "<div class='state-desc'>Upload one or more candidate CVs in the sidebar to begin standardized GCC CV generation.</div>"
                        "</div>", unsafe_allow_html=True)

elif _VIEW == "Processing History":
    # --- PROCESSING HISTORY ---
    st.markdown("<div class='page-heading fade-in'>"
                "<div class='page-title'>Processing <span class='accent'>History</span></div>"
                "<div class='page-sub'>Review the results of your current session.</div>"
                "</div>", unsafe_allow_html=True)

    _br = st.session_state.get("batch_results", {})
    _total_done = st.session_state.processed_count
    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f"<div class='stat-card'><div class='stat-number'>{_total_done}</div><div class='stat-label-sm'>Processed (Session)</div></div>", unsafe_allow_html=True)
    with m2:
        _okc = sum(1 for r in _br.values() if r.get("status") == "Success")
        st.markdown(f"<div class='stat-card' style='color:var(--green)'><div class='stat-number' style='color:var(--green)'>{_okc}</div><div class='stat-label-sm'>Succeeded</div></div>", unsafe_allow_html=True)
    with m3:
        _errc = sum(1 for r in _br.values() if r.get("status") == "Error")
        st.markdown(f"<div class='stat-card' style='color:var(--red)'><div class='stat-number' style='color:var(--red)'>{_errc}</div><div class='stat-label-sm'>Failed</div></div>", unsafe_allow_html=True)

    if _br:
        st.markdown("<div class='page-section'><span class='m m-soft'>receipt_long</span> Session Records</div>", unsafe_allow_html=True)
        for _fname, _res in _br.items():
            _st = _res.get("status", "Error")
            _stcls = "status-success" if _st == "Success" else "status-error"
            _sticon = "check_circle" if _st == "Success" else "cancel"
            if _st == "Success":
                _pd = _res.get("profile_data", {})
                _meta = f"{_pd.get('lastName', 'Lastname')}, {_pd.get('firstName', 'Firstname')} · {_pd.get('yearsOfExperience', 0)} yrs exp"
                _ext = _fname.split(".")[-1].upper()
            else:
                _meta = _res.get("error_message", "Unknown error")
                _ext = _fname.split(".")[-1].upper()
            st.markdown(
                f"<div class='record-row fade-in'>"
                f"<div class='f-icon'><span class='m'>description</span></div>"
                f"<div style='flex:1;'><div class='f-name'>{_fname}</div><div class='f-size'>{_meta}</div></div>"
                f"<div class='f-status'><span class='status-pill {_stcls}'><span class='m' style='font-size:0.9em'>{_sticon}</span> {_st}</span></div>"
                f"</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='empty-state fade-in'><span class='m' style='font-size:2rem;color:var(--brand);'>inbox</span><div class='empty-title'>No processing records yet</div><div class='empty-desc'>Upload CVs on the Dashboard and run a batch to see records here.</div></div>", unsafe_allow_html=True)

elif _VIEW == "Templates":
    # --- TEMPLATES ---
    st.markdown("<div class='page-heading fade-in'>"
                "<div class='page-title'>CV <span class='accent'>Templates</span></div>"
                "<div class='page-sub'>The GCC standard template and your local summary folder.</div>"
                "</div>", unsafe_allow_html=True)

    gcc_exists = os.path.exists(GCC_TEMPLATE_PATH)
    st.markdown("<div class='page-section'><span class='m m-soft'>description</span> GCC Standard Template</div>", unsafe_allow_html=True)
    if gcc_exists:
        try:
            with open(GCC_TEMPLATE_PATH, "rb") as _gf:
                _gdata = _gf.read()
            st.markdown("<div class='export-panel'><div class='export-panel-title'>GCC_CV_FORMAT.doc — the official template used for all standardized exports.</div>", unsafe_allow_html=True)
            st.download_button(
                "Download Template (GCC_CV_FORMAT.doc)",
                data=_gdata,
                file_name=GCC_TEMPLATE_NAME,
                mime="application/msword",
                key="dl_template"
            )
            st.markdown("</div>", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Could not read template: {e}")
    else:
        st.markdown("<div class='status-pill status-warn'><span class='m'>warning</span> Template missing</div>", unsafe_allow_html=True)
        st.info(f"Expected at `{GCC_TEMPLATE_PATH}`. Restore GCC_CV_FORMAT.doc to this location.")

    st.markdown("<div class='page-section'><span class='m m-soft'>folder</span> Candidate CV Summary Folder</div>", unsafe_allow_html=True)
    try:
        _summary_files = sorted(os.listdir(SUMMARY_DIR))
        _summary_files = [f for f in _summary_files if not f.startswith(".")]
    except Exception:
        _summary_files = []
    if _summary_files:
        st.markdown(f"<div style='font-size:0.85rem;color:var(--text-muted);margin-bottom:0.5rem;'>{len(_summary_files)} profile file(s) generated</div>", unsafe_allow_html=True)
        for _sf in _summary_files:
            st.markdown(f"<div class='record-row fade-in'><div class='f-icon'><span class='m'>article</span></div><div style='flex:1;'><div class='f-name'>{_sf}</div></div><div class='f-status'><span class='status-pill status-success'><span class='m' style='font-size:0.9em'>check</span> Saved</span></div></div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='empty-state fade-in'><span class='m' style='font-size:2rem;color:var(--brand);'>folder_open</span><div class='empty-title'>No profile files yet</div><div class='empty-desc'>Processed candidate profiles are saved to 'Candidate CV Summary/'.</div></div>", unsafe_allow_html=True)

elif _VIEW == "Settings":
    # --- SETTINGS ---
    st.markdown("<div class='page-heading fade-in'>"
                "<div class='page-title'>App <span class='accent'>Settings</span></div>"
                "<div class='page-sub'>Configure GCC CV Studio behaviour.</div>"
                "</div>", unsafe_allow_html=True)

    st.markdown("<div class='page-section'><span class='m m-soft'>tune</span> Preferences</div>", unsafe_allow_html=True)
    _auto = st.session_state.pref_auto_pdf
    _save = st.session_state.pref_save_profiles
    st.markdown(f"<div class='record-row'><div class='f-icon'><span class='m'>picture_as_pdf</span></div><div style='flex:1;'><div class='f-name'>Auto-generate PDF per candidate</div><div class='f-size'>{'Enabled' if _auto else 'Disabled'}</div></div></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='record-row'><div class='f-icon'><span class='m'>description</span></div><div style='flex:1;'><div class='f-name'>Save local profile .txt files</div><div class='f-size'>{'Enabled' if _save else 'Disabled'}</div></div></div>", unsafe_allow_html=True)
    st.caption("Toggles for these preferences live in the sidebar control panel and persist for the current session.")

    st.markdown("<div class='page-section'><span class='m m-soft'>storage</span> Paths</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='record-row'><div class='f-icon'><span class='m'>folder</span></div><div style='flex:1;'><div class='f-name'>Summary Folder</div><div class='f-size'>{SUMMARY_DIR}</div></div></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='record-row'><div class='f-icon'><span class='m'>description</span></div><div style='flex:1;'><div class='f-name'>GCC Template</div><div class='f-size'>{GCC_TEMPLATE_PATH}</div></div></div>", unsafe_allow_html=True)

    st.markdown("<div class='page-section'><span class='m m-soft'>construction</span> Maintenance</div>", unsafe_allow_html=True)
    if st.button(":material/restart_alt:" " Reset Session State", key="reset_session_btn"):
        for _k in list(st.session_state.keys()):
            if _k in ("batch_results", "selected_candidate_idx", "processed_count"):
                st.session_state.pop(_k, None)
        st.session_state.processed_count = 0
        st.rerun()

elif _VIEW == "Help & Guide":
    # --- HELP & GUIDE ---
    st.markdown("<div class='page-heading fade-in'>"
                "<div class='page-title'>Help &amp; <span class='accent'>Guide</span></div>"
                "<div class='page-sub'>Everything you need to get the most out of GCC CV Studio.</div>"
                "</div>", unsafe_allow_html=True)

    st.markdown("<div class='page-section'><span class='m m-soft'>play_circle</span> Quick Start</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='export-panel'>"
        "<div class='export-panel-title'>Step-by-step</div>"
        "<ol style='margin:0;padding-left:1.2rem;line-height:1.9;color:var(--text-body);'>"
        "<li><strong>Upload</strong> one or more raw CVs (PDF, DOCX, DOC, TXT) on the Dashboard.</li>"
        "<li>Click <strong>Standardize &amp; Export</strong> to parse and convert them to the GCC standard.</li>"
        "<li>Inspect each candidate's structured profile and raw text side-by-side.</li>"
        "<li><strong>Download</strong> standardized DOCX and PDF documents — individually or as a ZIP bundle.</li>"
        "</ol></div>", unsafe_allow_html=True)

    st.markdown("<div class='page-section'><span class='m m-soft'>support</span> Supported Formats</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='category-view'>"
        "<div class='feature-grid'>"
        "<div class='feature-card'><div class='feature-icon'><span class='m'>picture_as_pdf</span></div><div class='feature-title'>PDF</div><div class='feature-desc'>Text-based PDF CVs are parsed directly.</div></div>"
        "<div class='feature-card'><div class='feature-icon'><span class='m'>description</span></div><div class='feature-title'>DOCX</div><div class='feature-desc'>Microsoft Word files parsed using python-docx.</div></div>"
        "<div class='feature-card'><div class='feature-icon'><span class='m'>article</span></div><div class='feature-title'>DOC</div><div class='feature-desc'>Legacy Word files converted via LibreOffice.</div></div>"
        "<div class='feature-card'><div class='feature-icon'><span class='m'>text_snippet</span></div><div class='feature-title'>TXT</div><div class='feature-desc'>Plain-text CVs read directly.</div></div>"
        "</div>"
        "</div>", unsafe_allow_html=True)

    st.markdown("<div class='page-section'><span class='m m-soft'>privacy_tip</span> Privacy &amp; Security</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='privacy-note' style='margin:0.75rem 0;'><span class='m'>verified_user</span> All parsing happens locally on this device. Your candidate data never leaves your machine — no external AI or internet connection is required.</div>", unsafe_allow_html=True)

    st.markdown("<div class='page-section'><span class='m m-soft'>help</span> Troubleshooting</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='export-panel'>"
        "<div class='export-panel-title'>Common issues</div>"
        "<ul style='margin:0;padding-left:1.2rem;line-height:1.9;color:var(--text-body);'>"
        "<li><strong>.DOC files fail</strong> — LibreOffice (soffice) must be installed and available on the PATH.</li>"
        "<li><strong>No photo extracted</strong> — some raw CVs embed photos in non-standard ways; upload a custom photo when processing a single CV.</li>"
        "<li><strong>PDF export fails</strong> — click 'Generate PDF Export' to compile it on demand.</li>"
        "</ul></div>", unsafe_allow_html=True)

# --- FLOATING MANAGE APP BUTTON (routes to Settings) ---
st.markdown("<div class='floating-manage'>", unsafe_allow_html=True)
if st.button(":material/settings:" " Manage App", key="manage_app_btn"):
    st.session_state.view = "Settings"
    st.rerun()
st.markdown("</div>", unsafe_allow_html=True)
