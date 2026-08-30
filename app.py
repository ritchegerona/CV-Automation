#!/usr/bin/env python3
"""MSR CV Studio - CV Processing & Standardization Dashboard"""

import streamlit as st
from PIL import Image

st.set_page_config(
    page_title="MSR CV Studio",
    page_icon=":clipboard:",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
<style>
    /* === Global Reset & Base === */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }

    /* Hide Streamlit default elements */
    #MainMenu, header, footer, [data-testid="stSidebarNav"] { display: none; }
    [data-testid="stSidebar"] { padding-top: 2rem; }
    [data-testid="stHeader"] { display: none; }
    [data-testid="stToolbar"] { display: none; }

    /* Clean body background */
    .block-container {
        padding: 1.5rem 2rem;
        max-width: 1400px;
    }

    body {
        background: #f8fafc;
    }

    /* === Sidebar Styling === */
    .css-1d391kg, .css-1v5fmjr, [data-testid="stSidebar"] > div:first-child {
        background: #ffffff !important;
        border-right: 1px solid #e2e8f0 !important;
        padding: 1.25rem 1rem !important;
    }

    /* Sidebar logo area */
    .sidebar-logo {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        margin-bottom: 0.5rem;
    }
    .sidebar-logo-icon {
        width: 32px;
        height: 32px;
        background: linear-gradient(135deg, #0d9488, #14b8a6);
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-size: 0.85rem;
    }
    .sidebar-logo-text {
        font-size: 0.6rem;
        font-weight: 700;
        color: #0d9488;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }
    .sidebar-title {
        font-size: 1.15rem;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 0.15rem;
    }
    .sidebar-subtitle {
        font-size: 0.65rem;
        color: #94a3b8;
        margin-bottom: 0.75rem;
        line-height: 1.4;
    }
    .sidebar-dev {
        font-size: 0.6rem;
        color: #64748b;
        margin-bottom: 1.25rem;
        padding: 0.2rem 0.5rem;
        background: #f1f5f9;
        border-radius: 6px;
        display: inline-block;
    }

    /* Nav items */
    .nav-item {
        display: flex;
        align-items: center;
        gap: 0.6rem;
        padding: 0.6rem 0.75rem;
        border-radius: 10px;
        font-size: 0.82rem;
        color: #64748b;
        cursor: pointer;
        margin-bottom: 0.15rem;
        transition: all 0.2s;
        text-decoration: none;
    }
    .nav-item:hover { background: #f0fdfa; color: #0d9488; }
    .nav-item.active { background: #f0fdfa; color: #0d9488; font-weight: 600; }
    .nav-icon {
        font-size: 0.9rem;
        width: 20px;
        text-align: center;
    }

    /* Stats section */
    .stats-section {
        margin-top: 1.5rem;
        padding-top: 1.5rem;
        border-top: 1px solid #e2e8f0;
    }
    .stats-title {
        font-size: 0.6rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #94a3b8;
        margin-bottom: 0.5rem;
    }
    .stat-number {
        font-size: 2rem;
        font-weight: 800;
        color: #0d9488;
        line-height: 1;
    }
    .stat-label {
        font-size: 0.6rem;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }

    /* Status cards */
    .status-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 0.7rem 0.85rem;
        margin-bottom: 0.4rem;
        font-size: 0.72rem;
    }
    .status-card-header {
        display: flex;
        align-items: center;
        gap: 0.4rem;
        margin-bottom: 0.15rem;
    }
    .status-dot {
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: #10b981;
        display: inline-block;
        flex-shrink: 0;
    }
    .status-card-title {
        font-weight: 600;
        color: #334155;
        font-size: 0.75rem;
    }
    .status-card-desc {
        color: #64748b;
        font-size: 0.65rem;
        margin-left: 1rem;
        line-height: 1.4;
    }

    /* === Main Content === */
    h1, h2, h3 { color: #0f172a; font-weight: 700; }

    /* Upload card */
    .upload-card {
        background: linear-gradient(135deg, #f0fdfa 0%, #ecfeff 50%, #f0f9ff 100%);
        border: 1px solid #99f6e4;
        border-radius: 16px;
        padding: 1.75rem 2rem;
        margin-bottom: 2rem;
        position: relative;
        overflow: hidden;
        box-shadow: 0 2px 12px rgba(14, 165, 149, 0.06);
    }
    .upload-card-content {
        display: flex;
        align-items: center;
        gap: 2rem;
        justify-content: space-between;
    }
    .upload-left {
        flex: 1;
    }
    .upload-title {
        font-size: 0.95rem;
        font-weight: 600;
        color: #334155;
        margin-bottom: 0.75rem;
    }
    .upload-or {
        font-size: 0.75rem;
        color: #94a3b8;
        margin-bottom: 0.5rem;
    }
    .upload-right {
        display: flex;
        align-items: center;
        gap: 1.5rem;
    }
    .upload-icon-circle {
        width: 64px;
        height: 64px;
        background: white;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 2px 8px rgba(14, 165, 149, 0.12);
    }
    .upload-icon-circle i {
        font-size: 1.5rem;
        color: #0d9488;
    }
    .upload-cv-icon {
        font-size: 3.5rem;
        color: #0d9488;
    }

    /* Format badges */
    .format-badges-col {
        display: flex;
        flex-direction: column;
        gap: 0.35rem;
    }
    .format-badge {
        padding: 0.2rem 0.65rem;
        border-radius: 20px;
        font-size: 0.65rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }
    .format-pdf { background: #fee2e2; color: #dc2626; }
    .format-docx { background: #dbeafe; color: #2563eb; }
    .format-doc { background: #ffedd5; color: #ea580c; }
    .format-txt { background: #d1fae5; color: #059669; }

    .format-badges-row {
        display: flex;
        justify-content: center;
        gap: 0.4rem;
        margin-top: 0.75rem;
    }
    .format-badges-row .format-badge {
        padding: 0.2rem 0.6rem;
        font-size: 0.6rem;
    }

    /* CTA button */
    .cta-button {
        background: linear-gradient(135deg, #0d9488, #0891b2);
        color: white;
        border: none;
        padding: 0.6rem 1.5rem;
        border-radius: 10px;
        font-size: 0.88rem;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s;
        box-shadow: 0 2px 10px rgba(13, 148, 136, 0.25);
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
    }
    .cta-button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 14px rgba(13, 148, 136, 0.35);
    }
    .cta-button i {
        font-size: 0.85rem;
    }

    /* Section headings */
    .section-heading {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        margin-bottom: 0.5rem;
    }
    .section-icon {
        width: 44px;
        height: 44px;
        background: linear-gradient(135deg, #f0fdfa, #ccfbf1);
        border: 1px solid #99f6e4;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.1rem;
        color: #0d9488;
        flex-shrink: 0;
    }
    .section-icon i {
        font-size: 1.15rem;
    }
    .section-title {
        font-size: 1.45rem;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.2;
    }
    .section-title span {
        color: #0d9488;
    }
    .section-desc {
        font-size: 0.82rem;
        color: #64748b;
        margin-top: 0.2rem;
        margin-left: 0;
    }

    /* === Feature Cards (3 main) === */
    .feature-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1.25rem;
        margin-bottom: 2rem;
        margin-top: 1.5rem;
    }
    .feature-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 1.5rem;
        transition: all 0.2s;
        position: relative;
        overflow: hidden;
    }
    .feature-card:hover {
        border-color: #99f6e4;
        box-shadow: 0 4px 20px rgba(14, 165, 149, 0.1);
        transform: translateY(-2px);
    }
    .feature-card::after {
        content: "";
        position: absolute;
        bottom: 0;
        left: 0;
        right: 0;
        height: 3px;
        background: linear-gradient(90deg, #14b8a6, #06b6d4);
        border-radius: 0 0 14px 14px;
    }
    .feature-card-inner {
        display: flex;
        align-items: flex-start;
        gap: 1rem;
    }
    .feature-icon {
        width: 52px;
        height: 52px;
        background: linear-gradient(135deg, #f0fdfa, #ccfbf1);
        border: 1px solid #99f6e4;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        color: #0d9488;
        flex-shrink: 0;
    }
    .feature-icon i {
        font-size: 1.35rem;
    }
    .feature-content {
        flex: 1;
    }
    .feature-title {
        font-size: 0.95rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.25rem;
    }
    .feature-desc {
        font-size: 0.75rem;
        color: #64748b;
        line-height: 1.5;
        margin-bottom: 0.75rem;
    }
    .feature-tag {
        display: inline-flex;
        align-items: center;
        gap: 0.3rem;
        font-size: 0.68rem;
        font-weight: 600;
        color: #0d9488;
        background: #f0fdfa;
        border: 1px solid #ccfbf1;
        padding: 0.2rem 0.5rem;
        border-radius: 6px;
    }
    .feature-tag i {
        font-size: 0.6rem;
    }

    /* Bottom feature row */
    .bottom-features {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1rem;
        margin-top: 1rem;
    }
    .bottom-feature {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.1rem;
        text-align: left;
        display: flex;
        align-items: flex-start;
        gap: 0.75rem;
    }
    .bottom-feature-icon {
        width: 40px;
        height: 40px;
        background: linear-gradient(135deg, #f0fdfa, #ccfbf1);
        border: 1px solid #99f6e4;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1rem;
        color: #0d9488;
        flex-shrink: 0;
    }
    .bottom-feature-icon i {
        font-size: 1rem;
    }
    .bottom-feature-content {
        flex: 1;
    }
    .bottom-feature-title {
        font-size: 0.78rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.15rem;
    }
    .bottom-feature-desc {
        font-size: 0.68rem;
        color: #64748b;
        line-height: 1.4;
    }

    /* Manage app floating button */
    .manage-fab {
        position: fixed;
        bottom: 1.5rem;
        right: 1.5rem;
        background: linear-gradient(135deg, #0d9488, #0891b2);
        color: white;
        border: none;
        padding: 0.65rem 1.1rem;
        border-radius: 50px;
        font-size: 0.82rem;
        font-weight: 600;
        cursor: pointer;
        box-shadow: 0 4px 20px rgba(13, 148, 136, 0.4);
        z-index: 999;
        display: flex;
        align-items: center;
        gap: 0.4rem;
        transition: all 0.2s;
    }
    .manage-fab:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 24px rgba(13, 148, 136, 0.5);
    }
    .manage-fab i {
        font-size: 0.85rem;
    }

    /* Footer */
    .app-footer {
        text-align: center;
        padding: 1rem;
        font-size: 0.62rem;
        color: #94a3b8;
        margin-top: 1rem;
    }

    /* Top bar */
    .top-bar {
        display: flex;
        justify-content: flex-end;
        align-items: center;
        gap: 0.5rem;
        padding: 0.25rem 0;
        margin-bottom: 0.75rem;
    }
    .top-bar-btn {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 0.45rem 0.75rem;
        font-size: 0.8rem;
        color: #475569;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 0.35rem;
        transition: all 0.2s;
    }
    .top-bar-btn:hover {
        border-color: #99f6e4;
        color: #0d9488;
    }
    .top-bar-btn i {
        font-size: 0.85rem;
    }
    .top-bar-btn.share-btn {
        background: #f8fafc;
    }
    .avatar {
        width: 32px;
        height: 32px;
        background: linear-gradient(135deg, #0d9488, #0891b2);
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-size: 0.7rem;
        font-weight: 700;
        margin-left: 0.25rem;
    }

    /* Upload button override */
    .stButton > button {
        background: linear-gradient(135deg, #0d9488, #0891b2) !important;
        color: white !important;
        border: none !important;
        padding: 0.55rem 1.25rem !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
        box-shadow: 0 2px 10px rgba(13,148,136,0.25) !important;
        transition: all 0.2s !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 14px rgba(13,148,136,0.35) !important;
    }

    /* File uploader override */
    .stFileUploader > div,
    [data-testid="stFileUploader"] > div {
        background: transparent !important;
        border: none !important;
        padding: 0 !important;
    }
    [data-testid="stFileUploader"] > div:first-child {
        min-height: 0 !important;
    }

    /* Make the custom upload card fully clickable: cover it with the invisible
       native picker so clicking anywhere on the card opens the file browser.
       The uploader's clickable surface is its dropzone, so the dropzone is
       stretched across the whole card (the uploader div alone would leave large
       dead zones). Only the container block that directly holds the card is
       targeted, so nested stVerticalBlocks elsewhere are unaffected. */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .upload-card) {
        position: relative;
    }
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .upload-card) > [data-testid="stElementContainer"] {
        position: static !important;
    }
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .upload-card) [data-testid="stFileUploader"] {
        position: absolute !important;
        inset: 0 !important;
        opacity: 0 !important;
        cursor: pointer !important;
        z-index: 50 !important;
        padding: 0 !important;
        margin: 0 !important;
    }
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .upload-card) [data-testid="stFileUploaderDropzone"] {
        position: absolute !important;
        inset: 0 !important;
        min-height: 0 !important;
        border: none !important;
        background: transparent !important;
        cursor: pointer !important;
    }
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .upload-card) [data-testid="stFileUploader"] * {
        cursor: pointer !important;
    }
</style>
""", unsafe_allow_html=True)

# ─── TOP BAR ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="top-bar">
    <div class="top-bar-btn" title="Toggle theme"><i class="fa-regular fa-moon"></i></div>
    <div class="top-bar-btn share-btn" title="Share"><i class="fa-solid fa-share-nodes"></i> Share</div>
    <div class="top-bar-btn" title="Notifications"><i class="fa-regular fa-bell"></i></div>
    <div class="avatar">RG</div>
</div>
""", unsafe_allow_html=True)

# ─── UPLOAD CARD ───────────────────────────────────────────────────────────────
# Wrap the custom card and the native picker in one container so the invisible
# uploader can overlay the card and open the file browser when clicked anywhere.
upload_zone = st.container()
with upload_zone:
    st.markdown("""
<div class="upload-card">
    <div class="upload-card-content">
        <div class="upload-left">
            <div class="upload-title"><i class="fa-solid fa-cloud-arrow-up" style="color:#0d9488;margin-right:0.4rem;"></i>Drag & Drop raw CVs here (PDF, DOCX, DOC, or TXT)</div>
            <div class="upload-or">or</div>
            <button class="cta-button"><i class="fa-solid fa-upload"></i> Upload Files</button>
            <div style="font-size:0.68rem;color:#94a3b8;margin-top:0.4rem;">Max 200MB per file</div>
        </div>
        <div class="upload-right">
            <div class="upload-icon-circle">
                <i class="fa-solid fa-cloud-arrow-up"></i>
            </div>
            <div class="format-badges-col">
                <span class="format-badge format-pdf">PDF</span>
                <span class="format-badge format-docx">DOCX</span>
                <span class="format-badge format-doc">DOC</span>
                <span class="format-badge format-txt">TXT</span>
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

    # Streamlit file uploader (functional) – covers the custom card via CSS overlay
    uploaded_files = st.file_uploader(
        "Upload Files",
        type=["pdf", "docx", "doc", "txt"],
        accept_multiple_files=True,
        help="Drag & drop files here or click to browse",
    )
    if uploaded_files:
        st.success(f"{len(uploaded_files)} file(s) uploaded.")
        for uf in uploaded_files:
            st.write(uf.name)

# ─── SECTION HEADING ──────────────────────────────────────────────────────────
st.markdown("""
<div class="section-heading">
    <div class="section-icon"><i class="fa-solid fa-file-lines"></i></div>
    <div>
        <div class="section-title">CV Processing & Standardization <span>Studio</span></div>
        <div class="section-desc">Upload raw candidate CVs and transform them into polished, corporate-aligned GCC-standard resumes.</div>
    </div>
</div>
<div class="format-badges-row" style="margin-bottom:1.5rem;margin-top:0.75rem;">
    <span class="format-badge format-pdf">PDF</span>
    <span class="format-badge format-docx">DOCX</span>
    <span class="format-badge format-doc">DOC</span>
    <span class="format-badge format-txt">TXT</span>
</div>
""", unsafe_allow_html=True)

# ─── FEATURE CARDS (3 main) ───────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-card-inner">
            <div class="feature-icon"><i class="fa-solid fa-microchip"></i></div>
            <div class="feature-content">
                <div class="feature-title">Local Parser</div>
                <div class="feature-desc">Fully offline parsing — no external AI or internet required.</div>
                <span class="feature-tag"><i class="fa-solid fa-check"></i> Secure & Private</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-card-inner">
            <div class="feature-icon"><i class="fa-solid fa-bullseye"></i></div>
            <div class="feature-content">
                <div class="feature-title">GCC Standard</div>
                <div class="feature-desc">Outputs structured, corporate-aligned CVs matching the official template.</div>
                <span class="feature-tag"><i class="fa-solid fa-check"></i> 100% Compliant</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-card-inner">
            <div class="feature-icon"><i class="fa-solid fa-box-open"></i></div>
            <div class="feature-content">
                <div class="feature-title">Batch Export</div>
                <div class="feature-desc">Generate DOCX & PDF exports, individually or as a ZIP bundle.</div>
                <span class="feature-tag"><i class="fa-solid fa-check"></i> Fast & Reliable</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ─── BOTTOM FEATURES (4) ──────────────────────────────────────────────────────
st.markdown('<div style="margin-top:1.5rem;"></div>', unsafe_allow_html=True)
bf1, bf2, bf3, bf4 = st.columns(4)

with bf1:
    st.markdown("""
    <div class="bottom-feature">
        <div class="bottom-feature-icon"><i class="fa-solid fa-wand-magic-sparkles"></i></div>
        <div class="bottom-feature-content">
            <div class="bottom-feature-title">Smart Parsing</div>
            <div class="bottom-feature-desc">Extracts data accurately from any CV format.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with bf2:
    st.markdown("""
    <div class="bottom-feature">
        <div class="bottom-feature-icon"><i class="fa-solid fa-layer-group"></i></div>
        <div class="bottom-feature-content">
            <div class="bottom-feature-title">Template Driven</div>
            <div class="bottom-feature-desc">Uses GCC standard template for consistency.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with bf3:
    st.markdown("""
    <div class="bottom-feature">
        <div class="bottom-feature-icon"><i class="fa-solid fa-shield-halved"></i></div>
        <div class="bottom-feature-content">
            <div class="bottom-feature-title">Data Security</div>
            <div class="bottom-feature-desc">Your data stays local. Always.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with bf4:
    st.markdown("""
    <div class="bottom-feature">
        <div class="bottom-feature-icon"><i class="fa-solid fa-users"></i></div>
        <div class="bottom-feature-content">
            <div class="bottom-feature-title">Bulk Processing</div>
            <div class="bottom-feature-desc">Process multiple CVs in one go.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ─── FLOATING ACTION BUTTON ───────────────────────────────────────────────────
st.markdown("""
<button class="manage-fab"><i class="fa-solid fa-gear"></i> Manage app</button>
""", unsafe_allow_html=True)

# ─── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown('<div class="app-footer">© 2025 MSR CV Studio &nbsp;·&nbsp; All rights reserved.</div>', unsafe_allow_html=True)

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────

st.sidebar.markdown("""
<div style="padding:0.25rem 0.5rem;margin-bottom:0.75rem;">
    <div class="sidebar-logo">
        <div class="sidebar-logo-icon"><i class="fa-solid fa-file-medical"></i></div>
        <div>
            <div class="sidebar-logo-text">Medical Staffing</div>
        </div>
    </div>
    <div class="sidebar-title">MSR CV Studio</div>
    <div class="sidebar-subtitle">Enterprise CV Parser & Standardizer</div>
    <div class="sidebar-dev"><i class="fa-solid fa-code"></i> Developed by Ritchie Gerona</div>
</div>
""", unsafe_allow_html=True)

# Nav items
st.sidebar.markdown("""
<div class="nav-item active"><span class="nav-icon"><i class="fa-solid fa-house"></i></span> Dashboard</div>
<div class="nav-item"><span class="nav-icon"><i class="fa-regular fa-clock"></i></span> Processing History</div>
<div class="nav-item"><span class="nav-icon"><i class="fa-regular fa-file-lines"></i></span> Templates</div>
<div class="nav-item"><span class="nav-icon"><i class="fa-solid fa-gear"></i></span> Settings</div>
<div class="nav-item"><span class="nav-icon"><i class="fa-regular fa-circle-question"></i></span> Help & Guide</div>
""", unsafe_allow_html=True)

# Stats
st.sidebar.markdown("""
<div class="stats-section">
    <div class="stats-title">STATS</div>
    <div style="display:flex;align-items:baseline;gap:0.5rem;margin-bottom:0.15rem;">
        <span class="stat-number">0</span>
    </div>
    <div class="stat-label">PROCESSED THIS SESSION</div>
</div>
""", unsafe_allow_html=True)

# System status
st.sidebar.markdown("""
<div class="stats-section">
    <div class="stats-title">SYSTEM STATUS</div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown("""
<div class="status-card">
    <div class="status-card-header">
        <span class="status-dot"></span>
        <span class="status-card-title">Local Folder Status</span>
    </div>
    <div class="status-card-desc">Candidate CV Summary is active</div>
</div>
<div class="status-card">
    <div class="status-card-header">
        <span class="status-dot"></span>
        <span class="status-card-title">CV Format</span>
    </div>
    <div class="status-card-desc">GCC_CV_FORMAT.doc is loaded as the template</div>
</div>
""", unsafe_allow_html=True)

# Sidebar footer
st.sidebar.markdown("""
<div style="margin-top:1.5rem;padding:0.5rem;font-size:0.58rem;color:#94a3b8;line-height:1.6;border-top:1px solid #e2e8f0;">
    <i class="fa-regular fa-copyright"></i> 2025 MSR CV Studio<br>All rights reserved.
</div>
""", unsafe_allow_html=True)
