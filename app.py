#!/usr/bin/env python3
"""MSR CV Studio - CV Processing & Standardization Dashboard"""

import html
import io

import streamlit as st
from PIL import Image

try:
    from docx import Document
    from pypdf import PdfReader
except ImportError:
    Document = None
    PdfReader = None

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

    /* When the sidebar is collapsed, re-show only the expand button
       (right-pointing arrow) inside the otherwise hidden app header/toolbar. */
    [data-testid="stHeader"]:has([data-testid="stExpandSidebarButton"]) {
        display: flex !important;
        pointer-events: none;
        background: transparent !important;
    }
    [data-testid="stToolbar"]:has([data-testid="stExpandSidebarButton"]) {
        display: flex !important;
        pointer-events: none;
    }
    [data-testid="stToolbarActions"],
    [data-testid="stAppDeployButton"],
    [data-testid="stMainMenu"] {
        display: none !important;
    }
    [data-testid="stExpandSidebarButton"] {
        pointer-events: auto !important;
        z-index: 130 !important;
    }

    /* Clean body background */
    .block-container {
        padding: 1.5rem 2rem;
        max-width: 1400px;
    }

    body {
        background: #f8fafc;
    }

    /* === Sidebar Styling === */
    /* Collapsible sidebar; the hide/reveal controls stay always visible */
    [data-testid="stSidebarCollapseButton"] {
        visibility: visible !important;
        z-index: 120 !important;
        opacity: 1 !important;
    }
    [data-testid="stSidebarCollapseButton"] button {
        background: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.12) !important;
    }
    [data-testid="stSidebarCollapseButton"]:hover button {
        border-color: #0d9488 !important;
        box-shadow: 0 2px 8px rgba(13, 148, 136, 0.2) !important;
    }
    [data-testid="stSidebarCollapseButton"] [data-testid="stIconMaterial"] {
        color: #0d9488 !important;
    }
    [data-testid="stExpandSidebarButton"] {
        pointer-events: auto !important;
        z-index: 130 !important;
    }
    [data-testid="stExpandSidebarButton"] button,
    [data-testid="stExpandSidebarButton"] {
        background: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.18) !important;
    }
    [data-testid="stExpandSidebarButton"]:hover button,
    [data-testid="stExpandSidebarButton"]:hover {
        border-color: #0d9488 !important;
        box-shadow: 0 4px 12px rgba(13, 148, 136, 0.25) !important;
    }
    [data-testid="stExpandSidebarButton"] [data-testid="stIconMaterial"] {
        color: #0d9488 !important;
    }
    body.cv-dark [data-testid="stSidebarCollapseButton"] button,
    body.cv-dark [data-testid="stExpandSidebarButton"] button,
    body.cv-dark [data-testid="stExpandSidebarButton"] {
        background: #0f1b2d !important;
        border-color: #334155 !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.4) !important;
    }
    body.cv-dark [data-testid="stSidebarCollapseButton"] [data-testid="stIconMaterial"],
    body.cv-dark [data-testid="stExpandSidebarButton"] [data-testid="stIconMaterial"] {
        color: #5eead4 !important;
    }
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
        position: relative;
        overflow: hidden;
        transition: transform 0.15s ease, box-shadow 0.15s ease, filter 0.15s ease;
        box-shadow: 0 2px 10px rgba(13, 148, 136, 0.25);
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
    }
    .cta-button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 14px rgba(13, 148, 136, 0.35);
    }
    /* Hover is also intercepted by the covering dropzone, so drive the lift
       from the dropzone's :hover to get feedback anywhere on the card. */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .upload-card):has([data-testid="stFileUploaderDropzone"]:hover) .cta-button {
        transform: translateY(-1px);
        box-shadow: 0 4px 14px rgba(13, 148, 136, 0.35);
    }
    .cta-button i {
        font-size: 0.85rem;
    }
    /* Click ripple: white circle that expands across the button. Hidden by
       default; triggered while the covering dropzone is held down. */
    .cta-button::after {
        content: "";
        position: absolute;
        inset: 0;
        margin: auto;
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: rgba(255, 255, 255, 0.85);
        opacity: 0;
        pointer-events: none;
    }
    @keyframes upload-btn-ripple {
        0% { transform: scale(1); opacity: 0.7; }
        100% { transform: scale(22); opacity: 0; }
    }
    /* Click-press animation. The invisible dropzone covers the card, so it
       receives the pointer events; :active on the dropzone drives this. */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .upload-card):has([data-testid="stFileUploaderDropzone"]:active) .cta-button {
        transform: scale(0.92) translateY(1px);
        box-shadow: 0 1px 4px rgba(13, 148, 136, 0.4), inset 0 2px 6px rgba(2, 44, 41, 0.25);
        filter: brightness(1.1);
    }
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .upload-card):has([data-testid="stFileUploaderDropzone"]:active) .cta-button::after {
        animation: upload-btn-ripple 0.5s ease-out;
    }

    /* Decorative 3D teal folder, anchored right beside the upload card */
    .upload-folder {
        position: absolute;
        top: 0;
        left: calc(100% + 24px);
        width: 246px;
        z-index: 60;
    }
    .upload-folder svg {
        width: 100%;
        height: auto;
    }
    @media (max-width: 960px) {
        .upload-folder {
            position: static;
            margin: 0.5rem 0 1.5rem;
            width: 200px;
        }
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
       targeted, so nested stVerticalBlocks elsewhere are unaffected. The
       container is capped in width and left-aligned (the decorative folder
       sits in the blank space to its right). */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .upload-card) {
        position: relative;
        max-width: 521px;
        margin-right: auto;
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

    /* Post-upload summary + processing results */
    .format-file { background: #f1f5f9; color: #334155; border-color: #e2e8f0; }
    .upload-summary {
        margin-top: 1rem;
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 1rem 1.1rem;
        box-shadow: 0 2px 12px rgba(15, 23, 42, 0.06);
        max-width: 521px;
    }
    .upload-summary-title {
        font-size: 0.8rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.65rem;
        display: flex;
        align-items: center;
        gap: 0.4rem;
    }
    .upload-summary-title i { color: #0d9488; }
    .upload-file-row {
        display: flex;
        align-items: center;
        gap: 0.65rem;
        padding: 0.5rem 0.55rem;
        border-radius: 10px;
        background: #f8fafc;
        border: 1px solid #eef2f7;
        margin-bottom: 0.45rem;
    }
    .upload-file-row:last-child { margin-bottom: 0; }
    .upload-file-row .file-name {
        flex: 1;
        font-size: 0.8rem;
        font-weight: 600;
        color: #0f172a;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }
    .upload-file-row .file-size { font-size: 0.7rem; color: #94a3b8; white-space: nowrap; }
    .upload-status-pill {
        font-size: 0.68rem;
        font-weight: 700;
        padding: 0.15rem 0.55rem;
        border-radius: 999px;
        display: inline-flex;
        align-items: center;
        gap: 0.3rem;
        white-space: nowrap;
    }
    .upload-status-pill.status-ready { background: #ecfdf5; color: #047857; }
    .upload-status-pill.status-ok { background: #ecfdf5; color: #047857; }
    .upload-status-pill.status-skip { background: #fffbeb; color: #b45309; }
    .upload-status-pill.status-error { background: #fef2f2; color: #b91c1c; }
    .process-button-row { margin-top: 0.4rem; }
    .results-note { font-size: 0.72rem; color: #94a3b8; margin-top: 0.5rem; }

    /* === Dark Mode === */
    body.cv-dark { background: #0b1220; color: #dbe4f0; }
    body.cv-dark [data-testid="stAppViewContainer"] { background: #0b1220; }
    body.cv-dark [data-testid="stMainViewContainer"] > [data-testid="stAppViewContainer"] { background: #0b1220; }
    body.cv-dark .block-container { background: #0b1220; }
    body.cv-dark [data-testid="stSidebar"] > div:first-child,
    body.cv-dark [data-testid="stSidebar"] [data-testid="stSidebarContent"],
    body.cv-dark .css-1d391kg, body.cv-dark .css-1v5fmjr {
        background: #0d1729 !important;
        border-right-color: #1e293b !important;
    }
    body.cv-dark h1, body.cv-dark h2, body.cv-dark h3,
    body.cv-dark .sidebar-title, body.cv-dark .section-title,
    body.cv-dark .feature-title, body.cv-dark .bottom-feature-title,
    body.cv-dark .status-card-title { color: #e2e8f0; }
    body.cv-dark .sidebar-subtitle, body.cv-dark .sidebar-dev,
    body.cv-dark .section-desc, body.cv-dark .feature-desc,
    body.cv-dark .bottom-feature-desc, body.cv-dark .stats-title,
    body.cv-dark .stat-label, body.cv-dark .upload-or { color: #94a3b8; }
    body.cv-dark .sidebar-dev { background: #0f1b2d; }
    body.cv-dark .stats-section { border-top-color: #1e293b; }
    body.cv-dark .nav-item { color: #94a3b8; }
    body.cv-dark .nav-item:hover, body.cv-dark .nav-item.active { background: #0e2b33; color: #5eead4; }
    body.cv-dark .top-bar-btn { background: #0f1b2d; color: #94a3b8; }

    body.cv-dark .upload-card {
        background: linear-gradient(135deg, #073228 0%, #08263a 60%, #0a1f33 100%);
        border-color: #134e4a;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.4);
    }
    body.cv-dark .upload-title { color: #cbd5e1; }
    body.cv-dark .upload-icon-circle { background: #0f2740; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.4); }
    body.cv-dark .feature-card, body.cv-dark .bottom-feature,
    body.cv-dark .status-card, body.cv-dark .app-footer {
        background: #0f1b2d;
        border-color: #1e293b;
    }
    body.cv-dark .feature-icon, body.cv-dark .bottom-feature-icon,
    body.cv-dark .section-icon {
        background: linear-gradient(135deg, #0e2b33, #123042);
        border-color: #155e5a;
    }
    body.cv-dark .feature-tag { background: #0e2330; border-color: #155e5a; color: #5eead4; }
    body.cv-dark .upload-summary {
        background: #0f1b2d;
        border-color: #1e293b;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.3);
    }
    body.cv-dark .upload-summary-title { color: #e2e8f0; }
    body.cv-dark .upload-summary-title i { color: #2dd4bf; }
    body.cv-dark .upload-file-row { background: #0e2330; border-color: #1e293b; }
    body.cv-dark .upload-file-row .file-name { color: #e2e8f0; }
    body.cv-dark .format-file { background: #1e293b; color: #94a3b8; border-color: #334155; }

    /* Toast + profile menu */
    .cv-toast {
        position: fixed;
        left: 50%;
        bottom: 1.75rem;
        transform: translateX(-50%) translateY(20px);
        background: #0f172a;
        color: #e2e8f0;
        padding: 0.6rem 1.1rem;
        border-radius: 10px;
        font-size: 0.78rem;
        font-weight: 600;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.25);
        border: 1px solid #134e4a;
        z-index: 99999;
        opacity: 0;
        pointer-events: none;
        transition: opacity 0.2s ease, transform 0.2s ease;
        max-width: 80vw;
    }
    .cv-toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }
    .cv-toast::before {
        content: "\f05a";
        font-family: "Font Awesome 6 Free";
        font-weight: 900;
        color: #2dd4bf;
        margin-right: 0.4rem;
    }
    .cv-profile-menu {
        position: fixed;
        right: 1rem;
        top: 3.4rem;
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        box-shadow: 0 10px 34px rgba(15, 23, 42, 0.16);
        width: 190px;
        z-index: 9999;
        opacity: 0;
        transform: translateY(-6px);
        transition: all 0.18s ease;
        pointer-events: none;
        overflow: hidden;
    }
    .cv-profile-menu.open { opacity: 1; transform: translateY(0); pointer-events: auto; }
    .cv-profile-menu .pm-header {
        padding: 0.85rem 1rem;
        border-bottom: 1px solid #eef2f7;
        background: linear-gradient(135deg, #f0fdfa, #ecfeff);
    }
    .cv-profile-menu .pm-header b { color: #0f172a; font-size: 0.8rem; display: block; }
    .cv-profile-menu .pm-header span { color: #64748b; font-size: 0.68rem; }
    .cv-profile-menu .pm-item {
        padding: 0.6rem 1rem;
        font-size: 0.78rem;
        color: #334155;
        cursor: pointer;
        display: flex;
        gap: 0.5rem;
        align-items: center;
        transition: background 0.15s ease, color 0.15s ease;
    }
    .cv-profile-menu .pm-item:hover { background: #f0fdfa; color: #0d9488; }
    .cv-profile-menu .pm-item i { width: 16px; text-align: center; }
    body.cv-dark .cv-profile-menu { background: #0f1b2d; border-color: #1e293b; }
    body.cv-dark .cv-profile-menu .pm-header {
        background: linear-gradient(135deg, #0e2b33, #123042);
        border-color: #1e293b;
    }
    body.cv-dark .cv-profile-menu .pm-header b { color: #e2e8f0; }
    body.cv-dark .cv-profile-menu .pm-header span { color: #94a3b8; }
    body.cv-dark .cv-profile-menu .pm-item { color: #cbd5e1; }
    body.cv-dark .cv-profile-menu .pm-item:hover { background: #0e2b33; color: #5eead4; }
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

# ─── UI INTERACTIVITY ─────────────────────────────────────────────────────────
# Wires up the top-bar buttons (darkmode, share, notifications, profile) and the
# sidebar nav items. Scripts run in the app document via st.html, so they reach
# the whole page DOM including the sidebar.
st.html("""
<div class="cv-toast" id="cvToast"></div>
<div class="cv-profile-menu" id="cvProfileMenu">
    <div class="pm-header"><b>Ritchie Gerona</b><span>Administrator</span></div>
    <div class="pm-item" data-action="profile"><i class="fa-regular fa-user"></i>My Profile</div>
    <div class="pm-item" data-action="logout"><i class="fa-solid fa-arrow-right-from-bracket"></i>Logout</div>
</div>
<script>
(function () {
    var tries = 0;
    function toast(msg) {
        var t = document.getElementById('cvToast');
        if (!t) return;
        t.textContent = msg;
        t.classList.add('show');
        clearTimeout(t._h);
        t._h = setTimeout(function () { t.classList.remove('show'); }, 2200);
    }
    function mount() {
        if (tries++ > 40) return;
        var topBar = document.querySelector('.top-bar');
        var nav = document.querySelectorAll('.nav-item').length;
        if (!topBar || nav === 0) { setTimeout(mount, 200); return; }

        var moon = topBar.querySelector('[title="Toggle theme"]');
        if (moon && !moon.dataset.wired) {
            moon.dataset.wired = 1;
            moon.addEventListener('click', function () {
                var dark = document.body.classList.toggle('cv-dark');
                try { localStorage.setItem('cv-theme', dark ? 'dark' : 'light'); } catch (e) {}
                toast(dark ? 'Dark mode enabled' : 'Dark mode disabled');
            });
        }
        var share = topBar.querySelector('.share-btn');
        if (share && !share.dataset.wired) {
            share.dataset.wired = 1;
            share.addEventListener('click', function () {
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    navigator.clipboard.writeText(location.href).then(function () {
                        toast('Link copied to clipboard');
                    }).catch(function () { toast(location.href); });
                } else { toast(location.href); }
            });
        }
        var bell = topBar.querySelector('[title="Notifications"]');
        if (bell && !bell.dataset.wired) {
            bell.dataset.wired = 1;
            bell.addEventListener('click', function () { toast("You're all caught up"); });
        }
        var avatar = topBar.querySelector('.avatar');
        var menu = document.getElementById('cvProfileMenu');
        if (avatar && menu && !avatar.dataset.wired) {
            avatar.dataset.wired = 1;
            avatar.addEventListener('click', function (e) {
                e.stopPropagation();
                menu.classList.toggle('open');
            });
            menu.querySelectorAll('.pm-item').forEach(function (item) {
                item.addEventListener('click', function () {
                    menu.classList.remove('open');
                    toast(item.dataset.action === 'logout' ? 'Logged out (demo)' : 'My Profile - coming soon');
                });
            });
            document.addEventListener('click', function () { menu.classList.remove('open'); });
        }
        document.querySelectorAll('.nav-item').forEach(function (n) {
            if (n.dataset.wired) return;
            n.dataset.wired = 1;
            n.addEventListener('click', function () {
                document.querySelectorAll('.nav-item').forEach(function (x) { x.classList.remove('active'); });
                n.classList.add('active');
                var label = (n.textContent || '').trim();
                if (label === 'Dashboard') {
                    toast('Dashboard');
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                } else {
                    toast(label + ' - coming soon');
                }
            });
        });
        var fab = document.querySelector('.manage-fab');
        if (fab && !fab.dataset.wired) {
            fab.dataset.wired = 1;
            fab.addEventListener('click', function () { toast('Manage app - coming soon'); });
        }
    }
    function init() {
        var saved = null;
        try { saved = localStorage.getItem('cv-theme'); } catch (e) {}
        if (saved === 'dark' ||
            (saved === null && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
            document.body.classList.add('cv-dark');
        }
        mount();
    }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
</script>
""", unsafe_allow_javascript=True)

# ─── SECTION HEADING ──────────────────────────────────────────────────────────
st.markdown("""
<div class="section-heading">
    <div class="section-icon"><i class="fa-solid fa-file-lines"></i></div>
    <div>
        <div class="section-title">CV Processing & Standardization <span>Studio</span></div>
        <div class="section-desc">Upload raw candidate CVs and transform them into polished, corporate-aligned GCC-standard resumes.</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ─── UPLOAD CARD ───────────────────────────────────────────────────────────────
# Card sits on the left; the 3D teal folder is anchored right beside it.

def _human_size(num):
    """Format a byte count as a short human-readable string."""
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            return f"{int(num)} {unit}" if unit == "B" else f"{num:.1f} {unit}"
        num /= 1024
    return f"{num:.1f} GB"


def parse_cv(uploaded_file):
    """Extract plain text from an uploaded CV.

    Returns a (status, meta) tuple; meta carries the extracted text (or None)
    and a human-readable message. Status is one of ok / skip / error.
    """
    name = uploaded_file.name
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if ext == "docx":
        try:
            doc = Document(io.BytesIO(uploaded_file.getvalue()))
            text = "\n".join(p.text for p in doc.paragraphs)
        except Exception as exc:
            return "error", {"text": None, "message": f"Could not read DOCX: {exc}"}
        return "ok", {"text": text, "message": f"{len(text):,} characters extracted"}
    if ext == "pdf":
        try:
            reader = PdfReader(io.BytesIO(uploaded_file.getvalue()))
            text = "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception as exc:
            return "error", {"text": None, "message": f"Could not read PDF: {exc}"}
        return "ok", {"text": text, "message": f"{len(text):,} characters extracted"}
    if ext == "txt":
        text = uploaded_file.getvalue().decode("utf-8", errors="replace")
        return "ok", {"text": text, "message": f"{len(text):,} characters extracted"}
    return "skip", {"text": None, "message": f".{ext} uploads need LibreOffice installed (only DOCX, PDF, TXT parse on-device)"}


def _format_badge(ext):
    cls = f"format-{ext}" if ext in ("pdf", "docx", "doc", "txt") else "format-file"
    return f'<span class="format-badge {cls}">{ext.upper()}</span>'


def _draw_upload_summary(files):
    """Styled list of loaded files shown once the user has picked CVs."""
    rows = []
    for uf in files:
        ext = uf.name.rsplit(".", 1)[-1].lower() if "." in uf.name else "file"
        rows.append(
            '<div class="upload-file-row">'
            + _format_badge(ext)
            + f'<span class="file-name">{html.escape(uf.name)}</span>'
            + f'<span class="file-size">{_human_size(uf.size)}</span>'
            + '<span class="upload-status-pill status-ready">Ready</span>'
            + "</div>"
        )
    st.markdown(
        '<div class="upload-summary">'
        '<div class="upload-summary-title"><i class="fa-solid fa-circle-check"></i>'
        f'{len(files)} CV(s) loaded &nbsp;&mdash;&nbsp; review, then process</div>'
        + "".join(rows)
        + "</div>",
        unsafe_allow_html=True,
    )


def _draw_results(results):
    """Per-file parse status cards plus a raw-text preview per parsed file."""
    cards = []
    icons = {"ok": "fa-circle-check", "skip": "fa-triangle-exclamation", "error": "fa-circle-xmark"}
    labels = {"ok": "Parsed", "skip": "Skipped", "error": "Error"}
    pills = {"ok": "status-ok", "skip": "status-skip", "error": "status-error"}
    for uf, (status, meta) in results:
        ext = uf.name.rsplit(".", 1)[-1].lower() if "." in uf.name else "file"
        cards.append(
            '<div class="upload-file-row">'
            + _format_badge(ext)
            + f'<span class="file-name">{html.escape(uf.name)}</span>'
            + f'<span class="file-size">{html.escape(meta["message"])}</span>'
            + f'<span class="upload-status-pill {pills[status]}">'
            + f'<i class="fa-solid {icons[status]}"></i>{labels[status]}</span>'
            + "</div>"
        )
    st.markdown(
        '<div class="upload-summary">'
        '<div class="upload-summary-title"><i class="fa-solid fa-wand-magic-sparkles"></i>'
        "Processing complete &mdash; review below</div>"
        + "".join(cards)
        + "</div>",
        unsafe_allow_html=True,
    )
    for uf, (status, meta) in results:
        if status == "ok" and meta["text"]:
            preview = meta["text"].strip()
            if len(preview) > 5000:
                preview = preview[:5000] + "\n… (truncated)"
            with st.expander(f"Raw text preview — {uf.name}"):
                st.text(preview or "(no extractable text found)")


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
    if not uploaded_files:
        st.session_state.pop("cv_results", None)

    # Decorative 3D teal folder anchored beside the card
    st.markdown("""
<div class="upload-folder" aria-hidden="true">
    <svg viewBox="0 0 320 240" fill="none" xmlns="http://www.w3.org/2000/svg">
        <defs>
            <linearGradient id="uflFront" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0" stop-color="#2dd4bf"/>
                <stop offset="1" stop-color="#0d9488"/>
            </linearGradient>
            <linearGradient id="uflGloss" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0" stop-color="#ffffff" stop-opacity="0.35"/>
                <stop offset="1" stop-color="#ffffff" stop-opacity="0.05"/>
            </linearGradient>
        </defs>
        <ellipse cx="188" cy="216" rx="118" ry="12" fill="rgba(13,148,136,0.16)"/>
        <g>
            <rect x="104" y="50" width="34" height="60" rx="3" fill="#ffffff"/>
            <rect x="140" y="50" width="6" height="60" rx="2" fill="#cbd5e1"/>
            <rect x="112" y="62" width="18" height="3" rx="1.5" fill="#99f6e4"/>
            <rect x="112" y="70" width="18" height="3" rx="1.5" fill="#99f6e4"/>
            <rect x="122" y="42" width="34" height="64" rx="3" fill="#ffffff"/>
            <rect x="158" y="42" width="6" height="64" rx="2" fill="#cbd5e1"/>
            <rect x="130" y="56" width="18" height="3" rx="1.5" fill="#5eead4"/>
            <rect x="130" y="64" width="18" height="3" rx="1.5" fill="#5eead4"/>
            <rect x="140" y="58" width="34" height="58" rx="3" fill="#f0fdfa"/>
            <rect x="176" y="58" width="6" height="58" rx="2" fill="#99f6e4"/>
            <rect x="148" y="72" width="18" height="3" rx="1.5" fill="#14b8a6"/>
            <rect x="148" y="80" width="18" height="3" rx="1.5" fill="#14b8a6"/>
        </g>
        <rect x="84" y="76" width="48" height="32" rx="6" fill="#5eead4" fill-opacity="0.95"/>
        <rect x="84" y="76" width="48" height="32" rx="6" fill="url(#uflGloss)"/>
        <path d="M84 104h48" stroke="#0f766e" stroke-width="2"/>
        <path d="M72 104h150a10 10 0 0 1 10 10v66a10 10 0 0 1-10 10H72a10 10 0 0 1-10-10V114a10 10 0 0 1 10-10z" fill="url(#uflFront)"/>
        <path d="M222 104 L222 190 L252 202 L252 116 Z" fill="#0f766e"/>
        <path d="M72 190 L222 190 L252 202 L102 202 Z" fill="#115e59"/>
        <path d="M222 190 L252 202" stroke="#134e4a" stroke-width="2"/>
        <rect x="76" y="110" width="128" height="34" rx="7" fill="url(#uflGloss)"/>
        <path d="M72 104h150" stroke="rgba(255,255,255,0.5)" stroke-width="3"/>
        <path d="M222 104v86" stroke="rgba(255,255,255,0.35)" stroke-width="3"/>
    </svg>
</div>
""", unsafe_allow_html=True)

# Post-upload flow lives OUTSIDE the upload zone: the invisible picker overlay
# covers the whole zone block, so these elements must not sit underneath it.
if not uploaded_files:
    st.session_state.pop("cv_results", None)

if uploaded_files:
    _draw_upload_summary(uploaded_files)
    if st.button("Process CVs", key="process_cvs", type="primary"):
        st.session_state["cv_results"] = [(uf, parse_cv(uf)) for uf in uploaded_files]
    if "cv_results" in st.session_state:
        _draw_results(st.session_state["cv_results"])

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
