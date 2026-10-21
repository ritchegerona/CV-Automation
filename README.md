<div align="center">
  <h1>MSR CV Processing Studio</h1>
  <p><b>Standardize candidate CVs into the official GCC format — fully locally.</b></p>
  <p>
    <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License: MIT"/></a>
    <img src="https://img.shields.io/badge/Python-3.11-3776AB.svg?logo=python&logoColor=white" alt="Python 3.11"/>
    <img src="https://img.shields.io/badge/Streamlit-app-FF4B4B.svg?logo=streamlit&logoColor=white" alt="Streamlit app"/>
    <img src="https://img.shields.io/badge/AI%20API-not%20required-brightgreen.svg" alt="No AI API required"/>
  </p>
</div>

---

## Overview

A Streamlit web app that standardizes candidate CVs into a single, official **GCC CV format** using a fully **local, rule-based parser** — no external AI service, no internet access, and no API keys required. Upload a CV (or a batch), review the extracted profile, and export a clean, standardized document.

## Features

- **Flexible input** — upload CVs as PDF, DOCX, DOC, or TXT (single or batch).
- **Local parsing** — rule-based extraction of name, experience, personal details, passport details, education, and registration info. No API key, no network.
- **Official GCC layout** — renders every CV in the `GCC_CV_FORMAT.doc` structure (EXPERIENCE SUMMARY, PERSONAL DETAILS, PASSPORT DETAILS, EDUCATIONAL DETAILS, REGISTRATION DETAILS for PRC and SCFHS) with the MSR letterhead.
- **Exports** — DOCX download (works everywhere) and PDF download generated in-app with `reportlab`.
- **Candidate summaries** — saves a profile summary text file per CV.
- **Review UI** — side-by-side view of raw ingested text vs. the structured profile.

## Requirements

- **Python 3.11** (pinned in `runtime.txt` for cloud deploys)
- Python packages (see `requirements.txt`): `streamlit`, `python-docx`, `pypdf`, `reportlab`
- **Optional:** LibreOffice — only needed to parse legacy `.doc` uploads. DOCX, PDF, and TXT work without it.

## Project structure

```text
app.py                     Streamlit application (main entry point)
GCC_CV_FORMAT.doc          Official GCC CV format template
GCC_Header.png             MSR letterhead/header graphic
requirements.txt           Python dependencies
runtime.txt                Python version pin for cloud deploys
.streamlit/config.toml     Streamlit configuration
Installer/                 Optional desktop installer packaging (Win/macOS/Linux)
UI_PLAN.md                 UI design notes
```

## Run locally

```bash
# create a virtual environment and install deps
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

# start the app
.venv/bin/streamlit run app.py
```

Open the printed local URL (default <http://localhost:8501>).

### PDF export

PDF is generated in-app with **reportlab**, so it works in every environment (local and cloud) — no LibreOffice required.

> LibreOffice is only needed if you want to upload legacy `.doc` files for parsing; DOCX, PDF, and TXT uploads work without it.

## Deploy as a web app

This repository is configured for **Streamlit Community Cloud**:

1. Go to <https://share.streamlit.io> and sign in with your GitHub account.
2. Click **Create app** → **Deploy a public app from GitHub**.
3. Select repository `ritchegerona/CV-Automation`, branch `main`, main file `app.py`.
4. Click **Deploy**.

You get a live URL such as `https://<your-app>.streamlit.app`. DOCX/PDF export, parsing, and the full review UI all work in the cloud.

## Desktop installer (optional)

The [`Installer/`](Installer/README.md) directory contains PyInstaller/Inno Setup tooling to bundle the app as a native installer for **Windows**, **macOS**, and **Linux**. After launch it copies `GCC_CV_FORMAT.doc` and `GCC_Header.png` into the working directory and creates the `Candidate CV Summary/` folder automatically. No API key is needed.

## License

Distributed under the [MIT License](LICENSE).

© 2025 MSR CV Studio — made by Ritche Gerona.
