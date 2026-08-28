import os
import io
import re
import zipfile
import subprocess
import shutil
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor

# Set up page config
st.set_page_config(
    page_title="MSR CV Processing Studio",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Slate background & Deep Indigo accents)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    /* Main block padding */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    
    /* Premium Headers */
    .dashboard-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 0.2rem;
    }
    
    .dashboard-subtitle {
        font-size: 1.05rem;
        color: #64748B;
        margin-bottom: 2rem;
    }
    
    /* Card Container */
    .premium-card {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -2px rgba(0, 0, 0, 0.05);
        margin-bottom: 1.5rem;
    }
    
    /* Sidebar metrics styling */
    .sidebar-metric {
        background-color: #F8FAFC;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #E2E8F0;
        margin-bottom: 1rem;
    }
    
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #6366F1;
    }
    
    .metric-label {
        font-size: 0.85rem;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    /* Instruction Box */
    .instruction-box {
        background-color: #EEF2F6;
        border-left: 4px solid #6366F1;
        padding: 1rem;
        border-radius: 0 8px 8px 0;
        margin-bottom: 1.5rem;
    }
    
    /* Section dividers */
    .section-header {
        font-size: 1.25rem;
        font-weight: 600;
        color: #0F172A;
        border-bottom: 2px solid #F1F5F9;
        padding-bottom: 0.5rem;
        margin-top: 1rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Define directories and file targets
WORKSPACE_DIR = os.getcwd()
SUMMARY_DIR = os.path.join(WORKSPACE_DIR, "Candidate CV Summary")
GCC_TEMPLATE_NAME = "GCC_CV_FORMAT.doc"
GCC_TEMPLATE_PATH = os.path.join(WORKSPACE_DIR, GCC_TEMPLATE_NAME)
LOGO_PATH = os.path.join(WORKSPACE_DIR, "GCC_Header.png")

# Ensure the output summary directory exists
os.makedirs(SUMMARY_DIR, exist_ok=True)

# Initialize Session State
if "processed_count" not in st.session_state:
    st.session_state.processed_count = 0
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# --- SIDEBAR & SYSTEM METRICS ---
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/resume.png", width=64)
    st.markdown("### **MSR CV Studio**")
    st.markdown("Enterprise CV Parser & Standardizer")
    st.write("---")
    
    # Session Metrics
    st.markdown("<div class='sidebar-metric'>", unsafe_allow_html=True)
    st.markdown(f"<div class='metric-value'>{st.session_state.processed_count}</div>", unsafe_allow_html=True)
    st.markdown("<div class='metric-label'>Processed This Session</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Path & Status Information
    summary_exists = os.path.exists(SUMMARY_DIR)
    st.markdown("**Local Folder Status:**")
    if summary_exists:
        st.success("📁 `Candidate CV Summary/` is active")
    else:
        st.warning("⚠️ `Candidate CV Summary/` missing")
        
    # Template Checklist
    gcc_exists = os.path.exists(GCC_TEMPLATE_PATH)
    st.markdown("**CV Format:**")
    if gcc_exists:
        st.success("📄 `GCC_CV_FORMAT.doc` is loaded as the template")
    else:
        st.error("❌ `GCC_CV_FORMAT.doc` is missing in root")
        st.info("💡 Please place the official GCC CV format file `GCC_CV_FORMAT.doc` in your project root directory.")

# --- INSTRUCTION BLOCK ---
with st.expander("ℹ️ Local Execution & Project Instructions", expanded=False):
    st.markdown("""
    <div class='instruction-box'>
        <strong>🚀 Local Execution Guide:</strong><br/>
        1. Keep the Streamlit server running in your terminal: <code>.venv/bin/streamlit run app.py</code>.<br/>
        2. Verify that <code>GCC_CV_FORMAT.doc</code> is placed in your project root folder.<br/>
        3. Upload a CV file (Word <code>.docx</code>, <code>.doc</code>, PDF, or plain <code>.txt</code>) in the uploader below.<br/>
        4. Once standardized, view the split comparison, check the dynamically saved text profile, and download your export documents.
    </div>
    """, unsafe_allow_html=True)

# --- HEADER TITLE ---
st.markdown("<div class='dashboard-title'>CV Processing & Standardization Studio</div>", unsafe_allow_html=True)
st.markdown("<div class='dashboard-subtitle'>Transform raw, unformatted CVs into corporate-aligned, polished resumes instantly.</div>", unsafe_allow_html=True)

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

    section_style = ParagraphStyle('sec', fontName='Helvetica-Bold', fontSize=12,
                                   spaceBefore=10, spaceAfter=4, textColor=colors.HexColor('#003333'))
    field_style = ParagraphStyle('fld', fontName='Helvetica', fontSize=10.5,
                                 spaceBefore=3, spaceAfter=1)

    story.append(Paragraph(title, section_style))
    for k, v in pairs:
        story.append(Paragraph(f"<b>{k}:</b>&nbsp;&nbsp;{(v or ' ')}", field_style))
        story.append(HRFlowable(width='100%', thickness=0.7, color=colors.black,
                                dash=(1, 2), spaceBefore=1, spaceAfter=4))
    story.append(Spacer(1, 4))


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
                c.drawImage(logo_path, margin, h - margin - ch,
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
        c.setStrokeColor(colors.HexColor('#009999'))
        c.setLineWidth(1.2)
        c.setDash(1, 3)
        c.line(margin, h - margin - ch - 8, w - margin, h - margin - ch - 8)
        c.restoreState()

    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            leftMargin=margin, rightMargin=margin,
                            topMargin=header_pad, bottomMargin=margin,
                            title="Curriculum Vitae", author="MSR CV Processing Studio")

    story = []
    last_name = profile_data.get("lastName", "LAST NAME")
    first_name = profile_data.get("firstName", "FIRST NAME")
    details = profile_data.get("details", {})

    name_style = ParagraphStyle('nm', fontName='Helvetica-Bold', fontSize=13,
                                alignment=TA_CENTER, spaceAfter=8, textColor=colors.black)
    story.append(Paragraph(
        f"<u>{last_name.upper()}, {first_name.upper()}, MIDDLE NAME</u>", name_style))

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
            story.append(Paragraph(f"\u2022 {duty}", ParagraphStyle(
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
    
    # Check first 8 lines for a candidate name
    for line in lines[:8]:
        lower_line = line.lower()
        # Skip lines that look like contact info, email, phone, web, headers
        if '@' in line or 'mailto' in line or 'http' in line or 'www' in line:
            continue
        if any(char.isdigit() for char in line) and len(line) < 16: # Phone numbers or zip codes
            continue
        if len(line) > 40:
            continue
        # Skip common section headers if they appear early
        if lower_line in ["cv", "resume", "curriculum vitae", "summary", "contact", "about me", "experience"]:
            continue
        # Skip address/location lines
        if any(kw in lower_line for kw in ["city", "philippines", "street", "road", "province", "zip", "address", "brgy", "barangay", "block", "lot", "zone"]):
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
        if re.search(r'(University|College|School|Institute)\b', l) and len(l) < 80:
            school = l
            break
    if matched_edu:
        d["education"].append({"school": school, "qualification": matched_edu[0]})
    elif school != " ":
        d["education"].append({"school": school, "qualification": " "})

    # Work experience blocks (employer, position, duration, duties)
    exp_text_end = text
    m = re.search(r'(?i)(EXPERIENCE SUMMARY|WORK EXPERIENCE|EMPLOYMENT HISTORY|PROFESSIONAL EXPERIENCE)', text)
    if m:
        exp_text_end = text[m.start():]
    exp_year = r'(?:19\d{2}|20\d{2})\s*(?:-|–|—|to)\s*(?:(?:19\d{2}|20\d{2})|present|current)'
    blocks = []
    current_block = {}
    duty_lines = []
    for l in lines:
        if re.search(exp_year, l, re.IGNORECASE) and len(l) < 90:
            if current_block:
                current_block["duties"] = duty_lines[:10]
                blocks.append(current_block)
            pos = " "
            pm = re.search(r'(Staff Nurse|Nurse|RN|Registered Nurse)', l, re.IGNORECASE)
            if pm:
                pos = pm.group(1)
            current_block = {"employer": l, "position": pos, "duration": l}
            duty_lines = []
        elif "Area of Exposure" in l or "Bed Capacity" in l or "Name of Employer" in l:
            continue
        elif current_block and l:
            duty_lines.append(l)
    if current_block:
        current_block["duties"] = duty_lines[:10]
        blocks.append(current_block)
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

# --- FILE UPLOADER ---
uploaded_files = st.file_uploader(
    "Drag & Drop raw CVs here (PDF, DOCX, DOC, or TXT)", 
    type=["pdf", "docx", "doc", "txt"],
    accept_multiple_files=True,
    help="Supported files: PDF (.pdf), Microsoft Word (.docx, .doc), and plain text (.txt)",
    key=f"cv_uploader_{st.session_state.uploader_key}"
)

if uploaded_files:
    col_file, col_clear = st.columns([5, 1])
    with col_file:
        st.success(f"📂 Loaded: **{len(uploaded_files)} files**")
    with col_clear:
        if st.button("✖ Clear All Files", key="clear_files_btn", help="Clear all uploaded files"):
            st.session_state.uploader_key += 1
            for key in ["batch_results", "selected_candidate_idx"]:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()

    # Photo Uploader (only visible if a single CV is uploaded)
    uploaded_photo = None
    if len(uploaded_files) == 1:
        uploaded_photo = st.file_uploader(
            "📷 Optionally upload a custom profile photo (JPEG or PNG)",
            type=["jpg", "jpeg", "png"],
            help="If not provided, the studio will automatically attempt to extract the candidate photo from the raw CV.",
            key=f"photo_uploader_{st.session_state.uploader_key}"
        )
    
    # Local parsing notice
    st.info("🔌 Local parser is used for all CVs. No external AI or internet connection required.")
    
    # Action button to trigger processing
    col_btn_space, col_btn_act = st.columns([5, 1])
    with col_btn_act:
        process_btn = st.button("PROCEED ➔", type="primary", use_container_width=True)
        
    if process_btn:
        st.session_state.batch_results = {}
        progress_bar = st.progress(0.0)
        status_text = st.empty()
        
        for idx, file in enumerate(uploaded_files):
            file_name = file.name
            file_ext = file_name.split(".")[-1].lower()
            file_bytes = file.read()
            file.seek(0)
            
            status_text.markdown(f"⏳ **Processing file {idx+1} of {len(uploaded_files)}**: `{file_name}`...")
            progress_bar.progress(idx / len(uploaded_files))
            
            try:
                # 1. Parse raw text
                if file_ext == "docx":
                    raw_text = parse_docx(file_bytes)
                elif file_ext == "doc":
                    raw_text, doc_docx_bytes = parse_doc(file_bytes, file_name)
                    st.session_state.doc_docx_bytes = doc_docx_bytes
                elif file_ext == "pdf":
                    raw_text = parse_pdf(file_bytes)
                else:
                    raw_text = parse_txt(file_bytes)
                    
                # 2. Extract profile data (local parser)
                profile_data = offline_parse_cv(raw_text)
                # 3. Retrieve or extract profile picture
                photo_bytes = None
                if len(uploaded_files) == 1 and uploaded_photo is not None:
                    photo_bytes = uploaded_photo.read()
                else:
                    if file_ext == "doc" and 'doc_docx_bytes' in locals():
                        photo_bytes = extract_picture(doc_docx_bytes, "docx")
                    else:
                        photo_bytes = extract_picture(file_bytes, file_ext)
                        
                # 4. Pre-generate DOCX using the template
                last_name = profile_data.get("lastName", "Lastname")
                first_name = profile_data.get("firstName", "Firstname")
                
                import tempfile
                docx_bytes = None
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_docx_out = os.path.join(temp_dir, f"{last_name}_{first_name}_Standardized.docx")
                    generate_docx_document(profile_data, temp_docx_out, photo_bytes=photo_bytes)
                    with open(temp_docx_out, "rb") as doc_f:
                        docx_bytes = doc_f.read()
                        
                # 5. Save candidate profile text file locally
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
                    
                # 6. Pre-generate PDF for candidate
                pdf_bytes = None
                status_text.markdown(f"⏳ **Processing file {idx+1} of {len(uploaded_files)}**: Generating PDF for `{file_name}`...")
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
                    
                # 7. Cache result
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
                
        progress_bar.progress(1.0)
        status_text.markdown(f"✅ **Batch processing completed!** Processed {len(uploaded_files)} files.")
        st.session_state.processed_count += len(uploaded_files)
        st.rerun()

    # Check if we have processed data in state
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
                    st.markdown(f"<div class='sidebar-metric'><div class='metric-value'>{len(st.session_state.batch_results)}</div><div class='metric-label'>Total Uploaded</div></div>", unsafe_allow_html=True)
                with m_col2:
                    st.markdown(f"<div class='sidebar-metric'><div class='metric-value'>{success_count}</div><div class='metric-label'>Successfully Processed</div></div>", unsafe_allow_html=True)
                with m_col3:
                    st.markdown(f"<div class='sidebar-metric'><div class='metric-value'>{error_count}</div><div class='metric-label'>Errors/Failures</div></div>", unsafe_allow_html=True)
                
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
                
                st.markdown("<div class='section-header'>Batch Exports</div>", unsafe_allow_html=True)
                b_col1, b_col2 = st.columns(2)
                with b_col1:
                    st.download_button(
                        label=f"📦 Download All Standardized DOCX ({success_count} files as ZIP)",
                        data=docx_zip_bytes,
                        file_name="MSR_Standardized_CVs_Word.zip",
                        mime="application/zip",
                        key="dl_docx_zip"
                    )
                with b_col2:
                    if compiled_pdf_count > 0:
                        st.download_button(
                            label=f"📦 Download All Standardized PDFs ({compiled_pdf_count} files as ZIP)",
                            data=pdf_zip_bytes,
                            file_name="MSR_Standardized_CVs_PDF.zip",
                            mime="application/zip",
                            key="dl_pdf_zip"
                        )
                    else:
                        st.info("💡 Compile PDFs for individual candidates below, then download them as a batch ZIP here.")
                
                st.markdown("<div class='section-header'>Candidate Inspection & Single Export</div>", unsafe_allow_html=True)
                selected_filename = st.selectbox(
                    "🔍 Select Candidate Profile to Inspect:",
                    options=list(success_candidates.keys()),
                    format_func=lambda x: success_candidates[x]
                )
            else:
                selected_filename = list(success_candidates.keys())[0]
                
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
            st.info(f"💾 Dynamic Profile saved locally: `{os.path.join('Candidate CV Summary', profile_filename)}`")
            
            # --- SPLIT SCREEN LAYOUT ---
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("<div class='section-header'>Raw Ingested Text</div>", unsafe_allow_html=True)
                st.text_area("Original CV Content", raw_text, height=450, disabled=True, key=f"raw_{selected_filename}")
                
            with col2:
                header_text = "Structured CV Profile (Local Parser)"
                st.markdown(f"<div class='section-header'>{header_text}</div>", unsafe_allow_html=True)
                if photo_bytes:
                    st.image(photo_bytes, width=120)
                st.markdown(f"**Name:** {full_name}  \n**Total Experience:** {years_exp} Years")
                st.markdown(f"**Executive Summary:** {exec_summary}")
                st.markdown("---")
                st.markdown(cv_markdown)
                
            # --- DUAL EXPORT SYSTEM CARD ---
            st.markdown("<div class='section-header'>Single Export Options</div>", unsafe_allow_html=True)
            
            exp_col1, exp_col2 = st.columns(2)
            
            with exp_col1:
                st.download_button(
                    label="💾 Download Standardized Word (DOCX)",
                    data=docx_bytes,
                    file_name=f"{last_name}, {first_name} - CV.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_docx_{selected_filename}"
                )
                
            with exp_col2:
                if pdf_bytes is not None:
                    st.download_button(
                        label="📄 Download Standardized PDF",
                        data=pdf_bytes,
                        file_name=f"{last_name}, {first_name} - CV.pdf",
                        mime="application/pdf",
                        key=f"dl_pdf_{selected_filename}"
                    )
                else:
                    if st.button("📄 Generate PDF Export", key=f"gen_pdf_{selected_filename}", help="Generate and download the CV as a PDF"):
                        with st.spinner("Generating PDF..."):
                            try:
                                import tempfile as _tf
                                with _tf.TemporaryDirectory() as _dir:
                                    _pdf_path = os.path.join(_dir, f"{last_name}_{first_name}_Standardized.pdf")
                                    generate_pdf_direct(data, _pdf_path, photo_bytes=photo_bytes)
                                    with open(_pdf_path, "rb") as pf:
                                        compiled_pdf = pf.read()
                                if compiled_pdf:
                                    st.session_state.batch_results[selected_filename]["pdf_bytes"] = compiled_pdf
                                    st.rerun()
                                else:
                                    st.error("Failed to generate PDF document.")
                            except Exception as e:
                                st.error(f"Error generating PDF: {e}")
                                    
        # Display failures if any
        errors = [
            (filename, res["error_message"])
            for filename, res in st.session_state.batch_results.items()
            if res["status"] == "Error"
        ]
        if errors:
            st.markdown("<div class='section-header'>⚠️ Processing Failures</div>", unsafe_allow_html=True)
            for filename, err in errors:
                st.error(f"**{filename}**: {err}")
else:
    st.info("📂 Please upload one or more candidate CVs in PDF, DOCX, DOC, or TXT format above to begin.")
