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
<style>
    /* === Global Reset & Base === */
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

    /* === Sidebar Styling === */
    .css-1d391kg, .css-1v5fmjr, [data-testid="stSidebar"] > div:first-child {
        background: #f0f8ff !important;
        padding: 1.5rem 1rem !important;
    }

    /* Sidebar title */
    .sidebar-title {
        font-size: 1.25rem;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 0.25rem;
    }
    .sidebar-subtitle {
        font-size: 0.7rem;
        color: #64748b;
        margin-bottom: 1.5rem;
        line-height: 1.4;
    }
    .sidebar-dev {
        font-size: 0.65rem;
        color: #94a3b8;
        margin-bottom: 1.5rem;
        padding: 0.25rem 0.5rem;
        background: #f1f5f9;
        border-radius: 6px;
        display: inline-block;
    }

    /* Nav items */
    .nav-item {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.6rem 0.75rem;
        border-radius: 8px;
        font-size: 0.85rem;
        color: #475569;
        cursor: pointer;
        margin-bottom: 0.25rem;
        transition: all 0.2s;
    }
    .nav-item:hover { background: #e0f2fe; color: #0369a1; }
    .nav-item.active { background: #e0f2fe; color: #0369a1; font-weight: 600; }
    .nav-icon { font-size: 1rem; }

    /* Stats section */
    .stats-section {
        margin-top: 1.5rem;
        padding-top: 1.5rem;
        border-top: 1px solid #e2e8f0;
    }
    .stats-title {
        font-size: 0.65rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #94a3b8;
        margin-bottom: 0.75rem;
    }
    .stat-number {
        font-size: 2rem;
        font-weight: 800;
        color: #0d9488;
        line-height: 1;
    }
    .stat-label {
        font-size: 0.65rem;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }

    /* Status cards */
    .status-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 0.75rem;
        margin-bottom: 0.5rem;
        font-size: 0.75rem;
    }
    .status-card-header {
        display: flex;
        align-items: center;
        gap: 0.4rem;
        margin-bottom: 0.2rem;
    }
    .status-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: #10b981;
        display: inline-block;
    }
    .status-card-title {
        font-weight: 600;
        color: #334155;
    }
    .status-card-desc {
        color: #64748b;
        font-size: 0.7rem;
        margin-left: 1rem;
    }

    /* === Main Content === */
    h1, h2, h3 { color: #0f172a; font-weight: 700; }

    /* Upload card */
    .upload-card {
        background: linear-gradient(135deg, #f0fdfa 0%, #ecfeff 50%, #f0f9ff 100%);
        border: 2px dashed #99f6e4;
        border-radius: 16px;
        padding: 2rem;
        text-align: center;
        margin-bottom: 2rem;
        position: relative;
        overflow: hidden;
    }
    .upload-card::before {
        content: "";
        position: absolute;
        top: -50px;
        right: -50px;
        width: 200px;
        height: 200px;
        background: radial-gradient(circle, rgba(20,184,166,0.08) 0%, transparent 70%);
        border-radius: 50%;
    }
    .upload-icon {
        font-size: 3rem;
        margin-bottom: 0.5rem;
    }
    .upload-title {
        font-size: 1rem;
        font-weight: 600;
        color: #334155;
        margin-bottom: 0.5rem;
    }
    .upload-subtitle {
        font-size: 0.8rem;
        color: #64748b;
        margin-bottom: 1rem;
    }
    .upload-formats {
        display: flex;
        justify-content: center;
        gap: 0.5rem;
        margin-top: 1rem;
    }
    .format-badge {
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.7rem;
        font-weight: 700;
        text-transform: uppercase;
    }
    .format-pdf { background: #fee2e2; color: #dc2626; }
    .format-docx { background: #dbeafe; color: #2563eb; }
    .format-doc { background: #ffedd5; color: #ea580c; }
    .format-txt { background: #d1fae5; color: #059669; }

    /* Feature cards */
    .feature-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1.25rem;
        margin-bottom: 2rem;
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
    .feature-icon {
        width: 56px;
        height: 56px;
        background: linear-gradient(135deg, #f0fdfa, #ccfbf1);
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.5rem;
        margin-bottom: 1rem;
    }
    .feature-title {
        font-size: 1rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.35rem;
    }
    .feature-desc {
        font-size: 0.78rem;
        color: #64748b;
        line-height: 1.5;
        margin-bottom: 0.75rem;
    }
    .feature-tag {
        display: inline-flex;
        align-items: center;
        gap: 0.3rem;
        font-size: 0.7rem;
        font-weight: 600;
        color: #0d9488;
        background: #f0fdfa;
        padding: 0.2rem 0.5rem;
        border-radius: 6px;
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
        padding: 1.25rem;
        text-align: center;
    }
    .bottom-feature-icon {
        font-size: 1.5rem;
        margin-bottom: 0.5rem;
    }
    .bottom-feature-title {
        font-size: 0.8rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.25rem;
    }
    .bottom-feature-desc {
        font-size: 0.7rem;
        color: #64748b;
        line-height: 1.4;
    }

    /* CTA button */
    .cta-button {
        background: linear-gradient(135deg, #0d9488, #0891b2);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 10px;
        font-size: 0.95rem;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s;
        box-shadow: 0 4px 14px rgba(13, 148, 136, 0.3);
    }
    .cta-button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 20px rgba(13, 148, 136, 0.4);
    }

    /* Upload button override */
    .stButton > button {
        background: linear-gradient(135deg, #0d9488, #0891b2) !important;
        color: white !important;
        border: none !important;
        padding: 0.6rem 1.5rem !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        box-shadow: 0 4px 14px rgba(13,148,136,0.3) !important;
        transition: all 0.2s !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(13,148,136,0.4) !important;
    }

    /* File uploader override */
    .stFileUploader > div {
        background: transparent !important;
        border: none !important;
        padding: 0 !important;
    }

    /* Section headings */
    .section-heading {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        margin-bottom: 1.25rem;
    }
    .section-icon {
        width: 44px;
        height: 44px;
        background: linear-gradient(135deg, #f0fdfa, #ccfbf1);
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.25rem;
    }
    .section-title {
        font-size: 1.5rem;
        font-weight: 800;
        color: #0f172a;
    }
    .section-title span {
        color: #0d9488;
    }
    .section-desc {
        font-size: 0.85rem;
        color: #64748b;
        margin-top: 0.2rem;
    }

    /* Divider */
    .divider {
        border: none;
        border-top: 1px solid #e2e8f0;
        margin: 1.5rem 0;
    }

    /* Manage app floating button */
    .manage-fab {
        position: fixed;
        bottom: 2rem;
        right: 2rem;
        background: linear-gradient(135deg, #0d9488, #0891b2);
        color: white;
        border: none;
        padding: 0.75rem 1.25rem;
        border-radius: 50px;
        font-size: 0.85rem;
        font-weight: 600;
        cursor: pointer;
        box-shadow: 0 4px 20px rgba(13, 148, 136, 0.4);
        z-index: 999;
        display: flex;
        align-items: center;
        gap: 0.5rem;
        transition: all 0.2s;
    }
    .manage-fab:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 24px rgba(13, 148, 136, 0.5);
    }

    /* Footer */
    .app-footer {
        text-align: center;
        padding: 1rem;
        font-size: 0.65rem;
        color: #94a3b8;
        margin-top: 2rem;
    }

    /* Top bar simulation */
    .top-bar {
        display: flex;
        justify-content: flex-end;
        gap: 0.75rem;
        padding: 0.5rem 0;
        margin-bottom: 1rem;
    }
    .top-bar-btn {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 0.5rem 1rem;
        font-size: 0.8rem;
        color: #475569;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 0.4rem;
        transition: all 0.2s;
    }
    .top-bar-btn:hover {
        border-color: #99f6e4;
        color: #0d9488;
    }
    .avatar {
        width: 28px;
        height: 28px;
        background: linear-gradient(135deg, #0d9488, #0891b2);
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-size: 0.75rem;
        font-weight: 700;
    }
</style>
""", unsafe_allow_html=True)

# ─── Page Layout ───────────────────────────────────────────────────────────────

# TOP BAR
st.markdown("""
<div class="top-bar">
    <div class="top-bar-btn">🌙</div>
    <div class="top-bar-btn">📤 Share</div>
    <div class="top-bar-btn">🔔</div>
    <div class="avatar">RG</div>
</div>
""", unsafe_allow_html=True)

# ─── MAIN CONTENT AREA ────────────────────────────────────────────────────────

# UPLOAD CARD
st.markdown("""
<div class="upload-card">
    <div style="display:flex;align-items:center;gap:2rem;justify-content:center;flex-wrap:wrap;">
        <div style="text-align:left;">
            <div class="upload-title">📤 Drag & Drop raw CVs here (PDF, DOCX, DOC, or TXT)</div>
            <div class="upload-subtitle">or</div>
            <button class="cta-button" style="margin-top:0.5rem;">⬆️ &nbsp;Upload Files</button>
            <div style="font-size:0.7rem;color:#94a3b8;margin-top:0.5rem;">Max 200MB per file</div>
        </div>
        <div style="text-align:center;">
            <div style="font-size:4rem;margin-bottom:0.5rem;">📁</div>
            <div style="font-size:1.2rem;font-weight:800;color:#0d9488;letter-spacing:0.1em;">CV</div>
        </div>
        <div style="display:flex;flex-direction:column;gap:0.4rem;">
            <span class="format-badge format-pdf">PDF</span>
            <span class="format-badge format-docx">DOCX</span>
            <span class="format-badge format-doc">DOC</span>
            <span class="format-badge format-txt">TXT</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Section heading
st.markdown("""
<div class="section-heading">
    <div class="section-icon">📄</div>
    <div>
        <div class="section-title">CV Processing & Standardization <span>Studio</span></div>
        <div class="section-desc">Upload raw candidate CVs and transform them into polished, corporate-aligned GCC-standard resumes.</div>
    </div>
</div>
<div class="upload-formats" style="margin-bottom:2rem;">
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
        <div class="feature-icon">⚙️</div>
        <div class="feature-title">Local Parser</div>
        <div class="feature-desc">Fully offline parsing — no external AI or internet required.</div>
        <div class="feature-tag">✅ Secure & Private</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon">🎯</div>
        <div class="feature-title">GCC Standard</div>
        <div class="feature-desc">Outputs structured, corporate-aligned CVs matching the official template.</div>
        <div class="feature-tag">✅ 100% Compliant</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon">📦</div>
        <div class="feature-title">Batch Export</div>
        <div class="feature-desc">Generate DOCX & PDF exports, individually or as a ZIP bundle.</div>
        <div class="feature-tag">✅ Fast & Reliable</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ─── BOTTOM FEATURES (4) ──────────────────────────────────────────────────────
bf1, bf2, bf3, bf4 = st.columns(4)

with bf1:
    st.markdown("""
    <div class="bottom-feature">
        <div class="bottom-feature-icon">✨</div>
        <div class="bottom-feature-title">Smart Parsing</div>
        <div class="bottom-feature-desc">Extracts data accurately from any CV format.</div>
    </div>
    """, unsafe_allow_html=True)

with bf2:
    st.markdown("""
    <div class="bottom-feature">
        <div class="bottom-feature-icon">📋</div>
        <div class="bottom-feature-title">Template Driven</div>
        <div class="bottom-feature-desc">Uses GCC standard template for consistency.</div>
    </div>
    """, unsafe_allow_html=True)

with bf3:
    st.markdown("""
    <div class="bottom-feature">
        <div class="bottom-feature-icon">🛡️</div>
        <div class="bottom-feature-title">Data Security</div>
        <div class="bottom-feature-desc">Your data stays local. Always.</div>
    </div>
    """, unsafe_allow_html=True)

with bf4:
    st.markdown("""
    <div class="bottom-feature">
        <div class="bottom-feature-icon">👥</div>
        <div class="bottom-feature-title">Bulk Processing</div>
        <div class="bottom-feature-desc">Process multiple CVs in one go.</div>
    </div>
    """, unsafe_allow_html=True)

# Floating action button
st.markdown("""
<button class="manage-fab">⚙️ &nbsp;Manage app</button>
""", unsafe_allow_html=True)

# Footer
st.markdown('<div class="app-footer">© 2025 MSR CV Studio &nbsp;·&nbsp; All rights reserved.</div>', unsafe_allow_html=True)

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────

st.sidebar.markdown("""
<div style="padding:0.5rem 0.75rem;margin-bottom:1rem;">
    <div style="font-size:0.6rem;font-weight:700;color:#0d9488;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:0.25rem;">MEDICAL STAFFING</div>
    <div class="sidebar-title">MSR CV Studio</div>
    <div class="sidebar-subtitle">Enterprise CV Parser & Standardizer</div>
    <div class="sidebar-dev">⌨ Developed by Ritchie Gerona</div>
</div>
""", unsafe_allow_html=True)

# Nav items
st.sidebar.markdown("""
<div class="nav-item active"><span class="nav-icon">🏠</span> Dashboard</div>
<div class="nav-item"><span class="nav-icon">🕐</span> Processing History</div>
<div class="nav-item"><span class="nav-icon">📑</span> Templates</div>
<div class="nav-item"><span class="nav-icon">⚙️</span> Settings</div>
<div class="nav-item"><span class="nav-icon">❓</span> Help & Guide</div>
""", unsafe_allow_html=True)

# Stats
st.sidebar.markdown("""
<div class="stats-section">
    <div class="stats-title">STATS</div>
    <div style="display:flex;align-items:baseline;gap:0.5rem;margin-bottom:0.25rem;">
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
<div style="margin-top:2rem;padding:0.75rem;font-size:0.6rem;color:#94a3b8;line-height:1.6;">
    © 2025 MSR CV Studio<br>All rights reserved.
</div>
""", unsafe_allow_html=True)
