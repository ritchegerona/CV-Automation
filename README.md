# MSR CV Processing Studio

A Streamlit web app that standardizes candidate CVs into a single, official
**GCC CV format** using a fully **local parser** (no external AI or API keys
required).

## Features

- Upload CVs as PDF, DOCX, DOC, or TXT (single or batch).
- Local, rule-based parsing extracts name, experience, personal details,
  passport details, education, and registration info — no internet or API key.
- Renders every CV in the official **`GCC_CV_FORMAT.doc`** layout
  (EXPERIENCE SUMMARY, PERSONAL DETAILS, PASSPORT DETAILS, EDUCATIONAL
  DETAILS, REGISTRATION DETAILS for PRC and SCFHS) with the MSR letterhead.
- Exports standardized documents:
  - **DOCX** download (works everywhere).
  - **PDF** download (requires LibreOffice installed locally — see below).
- Saves a candidate profile summary text file per CV.
- Side-by-side view of raw ingested text vs. the structured CV profile.

## Project structure

```text
app.py               Streamlit application (main entry point)
GCC_CV_FORMAT.doc    Official GCC CV format template
GCC_Header.png       MSR letterhead/header graphic
requirements.txt     Python dependencies
runtime.txt          Python version pin for cloud deploys
.streamlit/config.toml
Installer/           Optional desktop installer packaging scripts
```

## Run locally

```bash
# create a virtual environment and install deps
python -m venv .venv
.venv/bin/pip install -r requirements.txt

# start the app
.venv/bin/streamlit run app.py
```

Open the printed local URL (default http://localhost:8501).

### PDF export (local only)

PDF generation relies on **LibreOffice** (`soffice`). Install it so the app
can convert DOCX to PDF:

- macOS: `brew install --cask libreoffice`
- Windows/Linux: install LibreOffice from https://www.libreoffice.org

If LibreOffice is not installed, DOCX export and parsing still work and the
app simply disables the PDF button.

## Deploy as a web app

This repository is configured for **Streamlit Community Cloud**:

1. Go to https://share.streamlit.io and sign in with your GitHub account.
2. Click **Create app** → **Deploy a public app from GitHub**.
3. Select repository `ritchegerona/CV-Automation`, branch `main`,
   main file `app.py`.
4. Click **Deploy**.

You get a live URL such as `https://<your-app>.streamlit.app`.

> Note: Streamlit Community Cloud has **no LibreOffice**, so PDF export is
> disabled in the hosted version. DOCX download, parsing, and the review UI
> all work in the cloud.
