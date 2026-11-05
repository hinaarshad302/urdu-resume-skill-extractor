
import re
import io
import os
import html
import streamlit as st
import pdfplumber
import pandas as pd

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    URDU_SHAPING_AVAILABLE = True
except ImportError:
    URDU_SHAPING_AVAILABLE = False


# =====================================================
# PAGE CONFIGURATION
# =====================================================
st.set_page_config(
    page_title="Urdu Resume Skill Extractor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =====================================================
# PROFESSIONAL DESIGN
# =====================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {font-family:'Inter',sans-serif;}
.stApp {background:#F5F7FB;color:#172B4D;}
.block-container {max-width:1250px;padding-top:2rem;padding-bottom:3rem;}
[data-testid="stSidebar"] {background:linear-gradient(180deg,#102A43 0%,#163B59 100%);}
[data-testid="stSidebar"] * {color:#F4F8FC !important;}
[data-testid="stSidebar"] .stCaption {color:#B9CBD9 !important;}
.sidebar-brand {padding:12px 0 22px 0;}
.sidebar-logo {font-size:34px;margin-bottom:8px;}
.sidebar-title {font-size:19px;font-weight:800;color:white;line-height:1.4;}
.sidebar-subtitle {font-size:12px;color:#B9CBD9;margin-top:5px;}
.sidebar-section {color:#8FCFC8;font-size:11px;font-weight:700;letter-spacing:1.2px;margin-top:22px;margin-bottom:9px;}
.sidebar-item {padding:9px 11px;background:rgba(255,255,255,0.07);border:1px solid rgba(255,255,255,0.08);border-radius:9px;margin-bottom:7px;font-size:12px;}
.sidebar-credit {border-top:1px solid rgba(255,255,255,0.2);margin-top:22px;padding-top:15px;text-align:center;color:#A8E2D8;font-size:12px;font-weight:700;}
.hero {background:linear-gradient(120deg,#102A43 0%,#176B70 100%);padding:30px 32px;border-radius:18px;color:white;margin-bottom:24px;box-shadow:0 8px 25px rgba(16,42,67,0.12);}
.hero h1 {color:white !important;font-size:31px !important;font-weight:800 !important;margin:0 0 10px 0;}
.hero p {color:#E1F1F1;font-size:14px;margin:0;line-height:1.7;}
.eyebrow {color:#A8E2D8;font-size:11px;font-weight:700;letter-spacing:1.5px;margin-bottom:10px;}
.section-heading {font-size:19px;font-weight:750;color:#17324D;margin:22px 0 13px 0;}
.info-card {background:white;border:1px solid #E5EAF1;border-radius:13px;padding:18px;min-height:112px;box-shadow:0 3px 12px rgba(20,45,75,0.035);}
.card-label {color:#718096;font-size:12px;font-weight:600;margin-bottom:10px;}
.card-value {color:#17324D;font-size:22px;font-weight:800;}
.content-card {background:white;border:1px solid #E5EAF1;border-radius:14px;padding:20px;margin-bottom:15px;}
.content-card h3 {color:#17324D;font-size:16px;margin-top:0;}
.skill-tag {display:inline-block;background:#E8F5F3;color:#176B70;border:1px solid #CDE9E4;padding:6px 11px;border-radius:20px;margin:4px 5px 4px 0;font-size:12px;font-weight:600;}
.upload-note {background:#EAF2FA;border:1px solid #D6E5F3;padding:13px 16px;border-radius:10px;color:#285477;font-size:13px;margin-bottom:15px;}
.stButton > button,.stDownloadButton > button {background:#176B70;color:white;border:none;border-radius:9px;font-weight:650;padding:9px 18px;}
.stButton > button:hover,.stDownloadButton > button:hover {background:#12565A;color:white;border:none;}
[data-testid="stFileUploader"] {background:white;border:1px dashed #9AB8C8;border-radius:13px;padding:14px;}
.footer {border-top:1px solid #E1E7EF;padding-top:20px;margin-top:40px;color:#7A8999;font-size:11px;text-align:center;line-height:1.8;}
</style>
""", unsafe_allow_html=True)


# =====================================================
# SKILL DATABASE
# =====================================================
SKILLS = [
    "retrieval augmented generation", "sentence transformers",
    "natural language processing", "artificial intelligence",
    "machine learning", "deep learning", "computer vision",
    "project management", "problem solving", "time management",
    "google gemini api", "google gemini", "generative ai",
    "agentic ai", "ai agents", "ai automation",
    "javascript", "typescript", "python", "java", "c++", "c#",
    "html", "css", "react", "node.js", "sql", "mysql",
    "mongodb", "data science", "data analysis", "streamlit",
    "fastapi", "django", "flask", "pandas", "numpy",
    "tensorflow", "pytorch", "github", "git", "gemini api",
    "prompt engineering", "rag", "faiss", "langchain",
    "langgraph", "crewai", "n8n", "make.com", "automation",
    "api", "rest api", "pypdf", "pdfplumber", "gradio",
    "communication", "teamwork", "ms word", "ms excel",
    "ms office", "excel", "documentation"
]

SKILL_DISPLAY = {
    "javascript": "JavaScript", "typescript": "TypeScript",
    "html": "HTML", "css": "CSS", "fastapi": "FastAPI",
    "node.js": "Node.js", "mysql": "MySQL",
    "mongodb": "MongoDB", "github": "GitHub",
    "python": "Python", "c++": "C++", "c#": "C#",
    "sql": "SQL", "api": "API", "rest api": "REST API",
    "ms office": "MS Office", "ms word": "MS Word",
    "ms excel": "MS Excel", "google gemini api": "Google Gemini API",
    "google gemini": "Google Gemini", "gemini api": "Gemini API",
    "generative ai": "Generative AI", "agentic ai": "Agentic AI",
    "ai agents": "AI Agents", "rag": "RAG", "faiss": "FAISS",
    "pypdf": "PyPDF", "n8n": "n8n", "make.com": "Make.com",
    "langchain": "LangChain", "langgraph": "LangGraph",
    "crewai": "CrewAI", "gradio": "Gradio",
    "sentence transformers": "Sentence Transformers",
    "machine learning": "Machine Learning",
    "deep learning": "Deep Learning",
    "artificial intelligence": "Artificial Intelligence",
    "natural language processing": "Natural Language Processing",
    "computer vision": "Computer Vision",
    "ai automation": "AI Automation",
    "problem solving": "Problem Solving",
    "time management": "Time Management",
    "project management": "Project Management"
}


# =====================================================
# TEXT EXTRACTION
# =====================================================
def extract_pdf(file):
    text = ""
    with pdfplumber.open(io.BytesIO(file.read())) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text


def extract_docx(file):
    document = Document(io.BytesIO(file.read()))
    paragraphs = [p.text for p in document.paragraphs if p.text.strip()]

    for table in document.tables:
        for row in table.rows:
            paragraphs.append(" | ".join(cell.text for cell in row.cells))

    return "\n".join(paragraphs)


def extract_txt(file):
    return file.read().decode("utf-8", errors="replace")


# =====================================================
# TEXT CLEANING
# =====================================================
def clean_extracted_line(line):
    if not line:
        return ""

    line = str(line)
    line = re.sub(r"\(cid:\d+\)", " ", line, flags=re.I)
    line = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", " ", line)
    line = re.sub(r"\s+", " ", line)
    return line.strip(" \t|•·-–—:;,")


def get_clean_lines(text):
    return [
        clean_extracted_line(line)
        for line in text.splitlines()
        if clean_extracted_line(line)
    ]


# =====================================================
# BASIC INFORMATION DETECTION
# =====================================================
def extract_email(text):
    match = re.search(
        r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
        text
    )
    return match.group(0) if match else ""


def extract_phone(text):
    match = re.search(
        r"(?<!\d)(?:\+92|0092|0)?[\s-]?\d{3}[\s-]?\d{7}(?!\d)",
        text
    )
    return match.group(0).strip() if match else ""


def extract_name(text):
    ignored = [
        "resume", "curriculum vitae", "personal information",
        "contact information", "professional summary",
        "career objective", "profile", "projects", "education",
        "experience", "technical skills", "skills", "phone",
        "email", "address", "linkedin", "github"
    ]

    for line in get_clean_lines(text):
        value = re.sub(
            r"\s+(projects?|resume|curriculum vitae)$",
            "", line, flags=re.I
        ).strip()

        if not value or len(value) > 65:
            continue
        if re.search(r"@|https?://|www\.|\d{4,}|[|]", value, re.I):
            continue
        if any(x in value.lower() for x in ignored):
            continue

        words = value.split()
        if 2 <= len(words) <= 4 and all(
            re.fullmatch(r"[A-Za-z][A-Za-z.'-]*", word)
            for word in words
        ):
            return value.title()

    return ""


# =====================================================
# SECTION DETECTION
# =====================================================
SECTION_ALIASES = {
    "projects": [
        "projects", "project", "personal projects",
        "academic projects", "key projects", "my projects"
    ],
    "experience": [
        "experience", "work experience", "professional experience",
        "employment history", "work history", "career history"
    ],
    "education": [
        "education", "academic background", "qualification",
        "qualifications", "educational background"
    ],
    "skills": [
        "skills", "technical skills", "core skills",
        "key skills", "technical expertise", "competencies"
    ],
    "summary": [
        "summary", "professional summary", "profile",
        "career objective", "objective", "about me"
    ],
    "languages": ["languages", "language"],
    "certifications": ["certifications", "certificates"],
    "achievements": ["achievements", "awards", "honors"]
}


def normalize_heading(line):
    value = line.lower().strip()
    value = re.sub(r"^[\d.)\s]+", "", value)
    value = re.sub(r"[:：|]+$", "", value)
    return value.strip()


def detect_section_heading(line):
    normalized = normalize_heading(line)

    for section, aliases in SECTION_ALIASES.items():
        if normalized in aliases:
            return section

    return None


def split_resume_sections(text):
    sections = {key: [] for key in SECTION_ALIASES}
    current_section = None

    for line in get_clean_lines(text):
        heading = detect_section_heading(line)

        if heading:
            current_section = heading
            continue

        if current_section:
            sections[current_section].append(line)

    return sections


# =====================================================
# SKILL DETECTION
# =====================================================
def extract_skills(text):
    lower_text = text.lower()
    found = []

    for skill in sorted(SKILLS, key=len, reverse=True):
        pattern = (
            r"(?<![a-z0-9])"
            + re.escape(skill)
            + r"(?![a-z0-9])"
        )

        if re.search(pattern, lower_text):
            display = SKILL_DISPLAY.get(skill, skill.title())
            if display not in found:
                found.append(display)

    return found


# =====================================================
# EDUCATION EXTRACTION
# =====================================================
def extract_education(text):
    sections = split_resume_sections(text)
    education_lines = sections["education"]

    keywords = [
        "phd", "ph.d", "master", "bachelor", "bs computer science",
        "bs cs", "bscs", "bba", "mba", "b.com", "m.com",
        "associate degree", "intermediate", "matric",
        "secondary school", "higher secondary", "diploma",
        "degree", "computer science", "software engineering",
        "information technology", "university", "college"
    ]

    results = []

    for line in education_lines:
        if any(word in line.lower() for word in keywords):
            if line not in results:
                results.append(line)

    # Fallback if no Education heading was detected.
    if not results:
        for line in get_clean_lines(text):
            lower = line.lower()
            if any(word in lower for word in keywords):
                if line not in results:
                    results.append(line)

    return results[:10]


# =====================================================
# PROJECT EXTRACTION
# =====================================================
def extract_projects(text):
    sections = split_resume_sections(text)
    project_lines = sections["projects"]

    projects = []
    current_project = []

    stop_words = [
        "experience", "work experience", "education",
        "skills", "technical skills", "languages",
        "certifications", "achievements", "references"
    ]

    def save_project():
        nonlocal current_project

        cleaned = []
        for item in current_project:
            item = clean_extracted_line(item)
            if item and item not in cleaned:
                cleaned.append(item)

        if cleaned:
            projects.append("\n".join(cleaned))

        current_project = []

    for line in project_lines:
        lower = line.lower().strip()

        # Ignore contact information accidentally placed in this section.
        if re.search(r"@|https?://|www\.", line, re.I):
            continue
        if re.search(r"\+?\d[\d\s()-]{8,}", line):
            continue

        if lower in stop_words:
            save_project()
            break

        numbered = re.match(
            r"^\s*(?:project\s*)?(\d{1,2})[.)\-:]\s*(.{3,100})$",
            line, re.I
        )

        if numbered:
            save_project()
            title = numbered.group(2).strip()
            if title:
                current_project = [title]
            continue

        # A short title-like line may start a new project.
        if (
            len(line) <= 75
            and not line.endswith((".", ",", ";"))
            and re.search(
                r"\b(app|system|website|platform|assistant|generator|"
                r"dashboard|automation|management|chatbot|application)\b",
                lower
            )
        ):
            save_project()
            current_project = [line]
            continue

        if current_project:
            current_project.append(line)
        else:
            current_project = [line]

    save_project()

    # Remove duplicates and overly short entries.
    unique = []
    for project in projects:
        if len(project.strip()) >= 5 and project not in unique:
            unique.append(project)

    return unique[:12]


# =====================================================
# EXPERIENCE EXTRACTION
# =====================================================
def extract_experience(text):
    sections = split_resume_sections(text)
    experience_lines = sections["experience"]

    if not experience_lines:
        return []

    results = []
    current_entry = []

    role_keywords = [
        "teacher", "developer", "engineer", "manager",
        "designer", "coordinator", "analyst", "accountant",
        "intern", "consultant", "officer", "administrator",
        "supervisor", "technician", "operator", "lecturer",
        "instructor", "sales", "marketing", "data entry",
        "customer service", "assistant", "specialist"
    ]

    date_pattern = re.compile(
        r"\b(19|20)\d{2}\b|present|current|"
        r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b",
        re.I
    )

    action_keywords = [
        "managed", "supported", "handled", "responsible",
        "maintained", "prepared", "assisted", "organized",
        "provided", "conducted", "created", "trained",
        "taught", "documented", "performed", "coordinated",
        "developed", "implemented", "designed", "worked"
    ]

    def save_entry():
        nonlocal current_entry
        cleaned = [clean_extracted_line(x) for x in current_entry]
        cleaned = [x for x in cleaned if x]

        if cleaned:
            item = "\n".join(dict.fromkeys(cleaned))
            if item not in results:
                results.append(item)

        current_entry = []

    for line in experience_lines:
        lower = line.lower()

        if re.search(r"@|https?://|www\.", line, re.I):
            continue

        looks_like_role = any(word in lower for word in role_keywords)
        has_date = bool(date_pattern.search(line))

        if looks_like_role and len(line) < 120:
            save_entry()
            current_entry = [line]
        elif current_entry:
            if has_date or any(word in lower for word in action_keywords):
                current_entry.append(line)
            elif len(line) > 3:
                current_entry.append(line)

    save_entry()

    # If a section exists but no clear role was detected,
    # retain its text as one entry instead of losing information.
    if not results and experience_lines:
        results = ["\n".join(experience_lines[:12])]

    return results[:10]


# =====================================================
# LANGUAGE LABELS
# =====================================================
LABELS = {
    "English": {
        "name": "Full Name",
        "job": "Target Job Title",
        "email": "Email",
        "phone": "Phone",
        "summary": "Professional Summary",
        "skills": "Skills",
        "education": "Education",
        "experience": "Experience",
        "projects": "Projects",
        "languages": "Languages"
    },
    "Urdu": {
        "name": "پورا نام",
        "job": "مطلوبہ عہدہ",
        "email": "ای میل",
        "phone": "فون نمبر",
        "summary": "پیشہ ورانہ تعارف",
        "skills": "مہارتیں",
        "education": "تعلیم",
        "experience": "تجربہ",
        "projects": "منصوبے",
        "languages": "زبانیں"
    }
}

SECTION_KEYS = [
    "summary", "skills", "education",
    "experience", "projects", "languages"
]


# =====================================================
# WORD EXPORT HELPERS
# =====================================================
def set_run_font(run, font_name="Arial", size=10, bold=None):
    run.font.name = font_name
    run.font.size = Pt(size)

    if bold is not None:
        run.bold = bold

    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts

    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)

    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:eastAsia"), font_name)
    r_fonts.set(qn("w:cs"), font_name)


def set_paragraph_rtl(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_pr = paragraph._p.get_or_add_pPr()

    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    p_pr.append(bidi)


def add_docx_section(doc, heading, value, rtl=False):
    if not str(value or "").strip():
        return

    h = doc.add_heading("", level=1)
    h_run = h.add_run(heading)
    set_run_font(h_run, "Noto Nastaliq Urdu" if rtl else "Arial", 13, True)

    p = doc.add_paragraph()
    p_run = p.add_run(str(value))
    set_run_font(p_run, "Noto Nastaliq Urdu" if rtl else "Arial", 11)

    if rtl:
        set_paragraph_rtl(h)
        set_paragraph_rtl(p)


def build_docx(profile, language, urdu_profile=None):
    urdu_profile = urdu_profile or {}
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)

    title = doc.add_heading("", 0)
    title_run = title.add_run(profile["name"] or "Candidate CV")
    set_run_font(title_run, "Arial", 22, True)

    if profile["job"]:
        p = doc.add_paragraph()
        run = p.add_run(profile["job"])
        set_run_font(run, "Arial", 11)

    contact = " | ".join(
        x for x in [profile["email"], profile["phone"]] if x
    )

    if contact:
        p = doc.add_paragraph()
        run = p.add_run(contact)
        set_run_font(run, "Arial", 9)

    if language in ["English", "Bilingual"]:
        for key in SECTION_KEYS:
            add_docx_section(
                doc,
                LABELS["English"][key],
                profile.get(key, "")
            )

    if language in ["Urdu", "Bilingual"]:
        for key in SECTION_KEYS:
            add_docx_section(
                doc,
                LABELS["Urdu"][key],
                urdu_profile.get(key, ""),
                rtl=True
            )

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer.add_run(
        "Prepared with Urdu Resume Skill Extractor | Built by Hina Arshad"
    )
    set_run_font(footer_run, "Arial", 8)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


# =====================================================
# EXCEL EXPORT
# =====================================================
def build_xlsx(profile, language, urdu_profile=None):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    urdu_profile = urdu_profile or {}
    rows = []

    if language == "Urdu":
        rows.extend([
            [LABELS["Urdu"]["name"], profile["name"]],
            [LABELS["Urdu"]["job"], profile["job"]],
            [LABELS["Urdu"]["email"], profile["email"]],
            [LABELS["Urdu"]["phone"], profile["phone"]]
        ])
    else:
        rows.extend([
            [LABELS["English"]["name"], profile["name"]],
            [LABELS["English"]["job"], profile["job"]],
            [LABELS["English"]["email"], profile["email"]],
            [LABELS["English"]["phone"], profile["phone"]]
        ])

    for key in SECTION_KEYS:
        if language == "Urdu":
            rows.append([
                LABELS["Urdu"][key],
                urdu_profile.get(key, "")
            ])
        elif language == "Bilingual":
            rows.append([
                LABELS["English"][key] + " / " + LABELS["Urdu"][key],
                "English:\n" + profile.get(key, "") +
                "\n\nاردو:\n" + urdu_profile.get(key, "")
            ])
        else:
            rows.append([
                LABELS["English"][key],
                profile.get(key, "")
            ])

    if language == "Urdu":
        headers = ["فیلڈ", "امیدوار کی معلومات"]
    elif language == "Bilingual":
        headers = ["Field / فیلڈ", "Candidate Information / امیدوار کی معلومات"]
    else:
        headers = ["Field", "Candidate Information"]

    df = pd.DataFrame(rows, columns=headers)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Candidate Profile")
        ws = writer.sheets["Candidate Profile"]

        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 75
        ws.sheet_view.rightToLeft = language in ["Urdu", "Bilingual"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        header_fill = PatternFill("solid", fgColor="176B70")
        thin_gray = Side(style="thin", color="E5EAF1")

        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = header_fill
            cell.alignment = Alignment(
                horizontal="right" if language == "Urdu" else "left",
                vertical="center",
                wrap_text=True
            )

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(
                    horizontal="right" if language in ["Urdu", "Bilingual"] else "left",
                    vertical="top",
                    wrap_text=True
                )
                cell.border = Border(bottom=thin_gray)

        ws.row_dimensions[1].height = 30

    return output.getvalue()


# =====================================================
# PDF EXPORT
# Urdu font: assets/NotoNastaliqUrdu-Regular.ttf
# =====================================================
def build_pdf(profile, language, urdu_profile=None):
    urdu_profile = urdu_profile or {}
    output = io.BytesIO()

    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=42,
        leftMargin=42,
        topMargin=42,
        bottomMargin=55
    )

    has_urdu = language in ["Urdu", "Bilingual"]
    font_name = "Helvetica"

    if has_urdu:
        if not URDU_SHAPING_AVAILABLE:
            raise ValueError(
                "Urdu PDF dependencies missing. requirements.txt mein "
                "arabic-reshaper aur python-bidi check karein."
            )

        font_path = os.path.join(
            "assets", "NotoNastaliqUrdu-Regular.ttf"
        )

        if not os.path.exists(font_path):
            raise ValueError(
                "Noto Nastaliq Urdu font assets folder mein nahi mila."
            )

        if "UrduFont" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("UrduFont", font_path))

        font_name = "UrduFont"

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "CVTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=19,
        leading=28,
        textColor=colors.HexColor("#102A43"),
        alignment=TA_LEFT,
        spaceAfter=10
    )

    english_heading = ParagraphStyle(
        "EnglishHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=18,
        textColor=colors.HexColor("#176B70"),
        alignment=TA_LEFT,
        spaceBefore=12,
        spaceAfter=5
    )

    urdu_heading = ParagraphStyle(
        "UrduHeading",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=13,
        leading=28,
        textColor=colors.HexColor("#176B70"),
        alignment=TA_RIGHT,
        spaceBefore=12,
        spaceAfter=8
    )

    english_body = ParagraphStyle(
        "EnglishBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=15,
        alignment=TA_LEFT
    )

    urdu_body = ParagraphStyle(
        "UrduBody",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=11,
        leading=27,
        alignment=TA_RIGHT
    )

    def safe_text(value):
        return html.escape(str(value or "")).replace("\n", "<br/>")

    def shape_urdu(value):
        if not value:
            return ""

        shaped_lines = []
        for line in str(value).splitlines():
            if line.strip():
                escaped = html.escape(line)
                shaped = arabic_reshaper.reshape(escaped)
                shaped_lines.append(get_display(shaped))
            else:
                shaped_lines.append("&nbsp;")

        return "<br/>".join(shaped_lines)

    story = [
        Paragraph(
            safe_text(profile["name"] or "Candidate CV"),
            title_style
        )
    ]

    if profile["job"]:
        story.append(Paragraph(safe_text(profile["job"]), english_body))

    contact = " | ".join(
        x for x in [profile["email"], profile["phone"]] if x
    )
    if contact:
        story.append(Paragraph(safe_text(contact), english_body))

    for key in SECTION_KEYS:
        english_value = profile.get(key, "")
        urdu_value = urdu_profile.get(key, "")

        if language in ["English", "Bilingual"] and english_value.strip():
            story.append(Paragraph(
                safe_text(LABELS["English"][key]),
                english_heading
            ))
            story.append(Paragraph(
                safe_text(english_value),
                english_body
            ))

        if language in ["Urdu", "Bilingual"] and urdu_value.strip():
            story.append(Paragraph(
                shape_urdu(LABELS["Urdu"][key]),
                urdu_heading
            ))
            story.append(Paragraph(
                shape_urdu(urdu_value),
                urdu_body
            ))

    def add_pdf_footer(canvas, document):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#DCE5EC"))
        canvas.line(42, 38, A4[0] - 42, 38)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#718096"))
        canvas.drawCentredString(
            A4[0] / 2,
            24,
            "Prepared with Urdu Resume Skill Extractor | Built by Hina Arshad"
        )
        canvas.restoreState()

    doc.build(
        story,
        onFirstPage=add_pdf_footer,
        onLaterPages=add_pdf_footer
    )

    return output.getvalue()


# =====================================================
# SIDEBAR
# =====================================================
with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="sidebar-logo">📄</div>
        <div class="sidebar-title">Resume Intelligence</div>
        <div class="sidebar-subtitle">Urdu Resume Skill Extractor</div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    st.markdown(
        '<div class="sidebar-section">PROJECT OVERVIEW</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">🏠 &nbsp; Resume Dashboard</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">📊 &nbsp; Candidate Analysis</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">📥 &nbsp; Profile Export</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sidebar-section">SUPPORTED FORMATS</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">📕 &nbsp; PDF Documents</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">📝 &nbsp; DOCX Documents</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">📃 &nbsp; TXT Documents</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sidebar-section">CURRENT CAPABILITIES</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">✓ &nbsp; Text extraction</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">✓ &nbsp; Email and phone detection</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">✓ &nbsp; Skills and project detection</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sidebar-item">✓ &nbsp; Word, Excel and PDF export</div>',
        unsafe_allow_html=True
    )

    st.divider()
    st.caption("ACADEMIC PROJECT")
    st.markdown("**FYP 2026–2027**")
    st.caption("Developed with Python and Streamlit")
    st.caption("Prototype Version 1.3")

    st.markdown(
        '<div class="sidebar-credit">Built by Hina Arshad</div>',
        unsafe_allow_html=True
    )


# =====================================================
# MAIN HEADER
# =====================================================
st.markdown("""
<div class="hero">
    <div class="eyebrow">INTELLIGENT DOCUMENT ANALYSIS PLATFORM</div>
    <h1>Urdu Resume Skill Extractor</h1>
    <p>
        Transform resume documents into structured candidate information.
        Extract contact details, education, experience, projects and
        technical skills through a simple and organized workflow.
    </p>
</div>
""", unsafe_allow_html=True)


# =====================================================
# WORKFLOW OVERVIEW
# =====================================================
c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("""
    <div class="info-card">
        <div class="card-label">STEP 01</div>
        <div class="card-value">📤 Upload</div>
        <div style="color:#718096;font-size:12px;">Select a resume document</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div class="info-card">
        <div class="card-label">STEP 02</div>
        <div class="card-value">🔍 Analyze</div>
        <div style="color:#718096;font-size:12px;">Detect candidate information</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown("""
    <div class="info-card">
        <div class="card-label">STEP 03</div>
        <div class="card-value">📥 Export</div>
        <div style="color:#718096;font-size:12px;">Download structured profile</div>
    </div>
    """, unsafe_allow_html=True)


# =====================================================
# RESUME UPLOAD
# =====================================================
st.markdown(
    '<div class="section-heading">01 &nbsp; Upload Resume</div>',
    unsafe_allow_html=True
)

st.markdown("""
<div class="upload-note">
    Upload a digital PDF, DOCX or TXT resume.
    Maximum file size is controlled by the application settings.
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Choose a candidate resume",
    type=["pdf", "docx", "txt"],
    help="Supported formats: PDF, DOCX and TXT"
)


# =====================================================
# PROCESS RESUME
# =====================================================
if uploaded_file:
    try:
        file_name = uploaded_file.name.lower()

        if file_name.endswith(".pdf"):
            resume_text = extract_pdf(uploaded_file)
        elif file_name.endswith(".docx"):
            resume_text = extract_docx(uploaded_file)
        else:
            resume_text = extract_txt(uploaded_file)

        st.success("Resume uploaded successfully.")

        if not resume_text.strip():
            st.warning(
                "No readable text was found. This may be a scanned PDF. "
                "OCR support will be introduced in a future development phase."
            )

        else:
            email = extract_email(resume_text)
            phone = extract_phone(resume_text)
            skills = extract_skills(resume_text)
            education = extract_education(resume_text)
            experience = extract_experience(resume_text)
            projects = extract_projects(resume_text)
            detected_name = extract_name(resume_text)

            # -----------------------------------------
            # 02 CANDIDATE OVERVIEW
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">02 &nbsp; Candidate Overview</div>',
                unsafe_allow_html=True
            )

            m1, m2, m3, m4 = st.columns(4)

            metrics = [
                ("DETECTED SKILLS", len(skills)),
                ("EDUCATION ENTRIES", len(education)),
                ("EXPERIENCE ENTRIES", len(experience)),
                ("PROJECTS DETECTED", len(projects))
            ]

            for column, (label, value) in zip([m1, m2, m3, m4], metrics):
                with column:
                    st.markdown(
                        f'<div class="info-card"><div class="card-label">{label}</div><div class="card-value">{value}</div></div>',
                        unsafe_allow_html=True
                    )

            # -----------------------------------------
            # 03 CONTACT INFORMATION
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">03 &nbsp; Contact Information</div>',
                unsafe_allow_html=True
            )

            contact1, contact2 = st.columns(2)

            with contact1:
                st.markdown(
                    '<div class="content-card"><h3>✉️ Email Address</h3>',
                    unsafe_allow_html=True
                )
                st.write(email or "Not detected")
                st.markdown("</div>", unsafe_allow_html=True)

            with contact2:
                st.markdown(
                    '<div class="content-card"><h3>📞 Phone Number</h3>',
                    unsafe_allow_html=True
                )
                st.write(phone or "Not detected")
                st.markdown("</div>", unsafe_allow_html=True)

            # -----------------------------------------
            # 04 TECHNICAL SKILLS
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">04 &nbsp; Technical Skills</div>',
                unsafe_allow_html=True
            )

            st.markdown('<div class="content-card">', unsafe_allow_html=True)

            if skills:
                skill_html = "".join(
                    f'<span class="skill-tag">{html.escape(skill)}</span>'
                    for skill in skills
                )
                st.markdown(skill_html, unsafe_allow_html=True)
            else:
                st.info("No matching skills detected.")

            st.markdown("</div>", unsafe_allow_html=True)

            # -----------------------------------------
            # 05 EDUCATION AND 06 EXPERIENCE
            # -----------------------------------------
            left, right = st.columns(2)

            with left:
                st.markdown(
                    '<div class="section-heading">05 &nbsp; Education</div>',
                    unsafe_allow_html=True
                )
                st.markdown(
                    '<div class="content-card"><h3>🎓 Academic Background</h3>',
                    unsafe_allow_html=True
                )

                if education:
                    for item in education:
                        st.write("• " + item)
                else:
                    st.write("No education information detected.")

                st.markdown("</div>", unsafe_allow_html=True)

            with right:
                st.markdown(
                    '<div class="section-heading">06 &nbsp; Experience</div>',
                    unsafe_allow_html=True
                )
                st.markdown(
                    '<div class="content-card"><h3>💼 Work Background</h3>',
                    unsafe_allow_html=True
                )

                if experience:
                    for item in experience:
                        st.write(item)
                else:
                    st.write("No experience details detected.")

                st.markdown("</div>", unsafe_allow_html=True)

            # -----------------------------------------
            # 07 PROJECTS
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">07 &nbsp; Projects</div>',
                unsafe_allow_html=True
            )

            st.markdown(
                '<div class="content-card"><h3>🚀 Candidate Projects</h3>',
                unsafe_allow_html=True
            )

            if projects:
                for index, project in enumerate(projects, start=1):
                    st.markdown(f"**Project {index}**")
                    st.write(project)
                    if index < len(projects):
                        st.divider()
            else:
                st.info(
                    "No separate project entries detected. "
                    "You can add project details manually below."
                )

            st.markdown("</div>", unsafe_allow_html=True)

            # -----------------------------------------
            # ORIGINAL EXTRACTED TEXT
            # -----------------------------------------
            with st.expander("📄 View Original Extracted Resume Text"):
                st.text_area(
                    "Extracted document content",
                    resume_text,
                    height=250,
                    key="original_resume_text"
                )

            # -----------------------------------------
            # 08 EDITABLE CANDIDATE PROFILE
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">08 &nbsp; Edit Candidate Profile</div>',
                unsafe_allow_html=True
            )

            st.caption(
                "Extracted information ko edit karein. "
                "Missing details manually enter kar sakte hain."
            )

            col_a, col_b = st.columns(2)

            with col_a:
                candidate_name = st.text_input(
                    "Candidate Full Name",
                    value=detected_name,
                    key="candidate_name"
                )

                candidate_email = st.text_input(
                    "Email Address",
                    value=email,
                    key="candidate_email"
                )

                candidate_phone = st.text_input(
                    "Phone Number",
                    value=phone,
                    key="candidate_phone"
                )

                target_job = st.text_input(
                    "Target Job Title",
                    value="",
                    key="target_job"
                )

            with col_b:
                summary = st.text_area(
                    "Professional Summary",
                    value="",
                    height=100,
                    key="candidate_summary"
                )

                language = st.selectbox(
                    "Select CV Language / CV ki zaban",
                    ["English", "Urdu", "Bilingual"],
                    key="cv_language"
                )

            skills_text = st.text_area(
                "Skills (comma-separated or one per line)",
                value=", ".join(skills),
                height=90,
                key="candidate_skills"
            )

            education_text = st.text_area(
                "Education",
                value="\n".join(education),
                height=100,
                key="candidate_education"
            )

            experience_text = st.text_area(
                "Experience",
                value="\n\n".join(experience),
                height=120,
                key="candidate_experience"
            )

            projects_text = st.text_area(
                "Projects",
                value="\n\n".join(projects),
                height=160,
                key="candidate_projects",
                help="Projects ko edit karein ya manually naye projects add karein."
            )

            languages_text = st.text_input(
                "Languages",
                value="",
                key="candidate_languages"
            )

            profile = {
                "name": candidate_name.strip(),
                "job": target_job.strip(),
                "email": candidate_email.strip(),
                "phone": candidate_phone.strip(),
                "summary": summary.strip(),
                "skills": skills_text.strip(),
                "education": education_text.strip(),
                "experience": experience_text.strip(),
                "projects": projects_text.strip(),
                "languages": languages_text.strip()
            }

            # -----------------------------------------
            # 09 MANUAL URDU CONTENT
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">09 &nbsp; Urdu CV Content</div>',
                unsafe_allow_html=True
            )

            st.caption(
                "Urdu ya Bilingual CV ke liye neeche Urdu content enter karein. "
                "Name, email aur phone original profile se liye jayenge."
            )

            urdu_profile = {}

            with st.expander(
                "✍️ Enter Urdu CV Details",
                expanded=(language != "English")
            ):
                urdu_profile["summary"] = st.text_area(
                    "پیشہ ورانہ تعارف",
                    key="urdu_summary",
                    height=100,
                    placeholder="اپنا پیشہ ورانہ تعارف اردو میں لکھیں"
                )

                urdu_profile["skills"] = st.text_area(
                    "مہارتیں",
                    key="urdu_skills",
                    height=90,
                    placeholder="اپنی مہارتیں اردو میں لکھیں"
                )

                urdu_profile["education"] = st.text_area(
                    "تعلیم",
                    key="urdu_education",
                    height=100,
                    placeholder="اپنی تعلیمی معلومات اردو میں لکھیں"
                )

                urdu_profile["experience"] = st.text_area(
                    "تجربہ",
                    key="urdu_experience",
                    height=100,
                    placeholder="اپنا تجربہ اردو میں لکھیں"
                )

                urdu_profile["projects"] = st.text_area(
                    "منصوبے",
                    key="urdu_projects",
                    height=100,
                    placeholder="اپنے منصوبوں کی تفصیل اردو میں لکھیں"
                )

                urdu_profile["languages"] = st.text_area(
                    "زبانیں",
                    key="urdu_languages",
                    height=70,
                    placeholder="مثلاً اردو، انگریزی"
                )

            # -----------------------------------------
            # 10 EXPORT PROFESSIONAL CV
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">10 &nbsp; Export Professional CV</div>',
                unsafe_allow_html=True
            )

            st.markdown('<div class="content-card">', unsafe_allow_html=True)
            st.write(
                "Review your details and download your selected CV format."
            )

            if language in ["Urdu", "Bilingual"]:
                missing_urdu = [
                    LABELS["Urdu"][key]
                    for key in SECTION_KEYS
                    if not urdu_profile.get(key, "").strip()
                ]

                if missing_urdu:
                    st.warning(
                        "Urdu fields incomplete hain: " +
                        ", ".join(missing_urdu) +
                        ". Empty Urdu fields export mein include nahi honge."
                    )
                else:
                    st.success("All Urdu CV sections have content.")

                if not URDU_SHAPING_AVAILABLE:
                    st.error(
                        "Urdu PDF support ke liye requirements.txt mein "
                        "arabic-reshaper aur python-bidi required hain."
                    )

            e1, e2, e3 = st.columns(3)

            with e1:
                try:
                    word_file = build_docx(profile, language, urdu_profile)
                    st.download_button(
                        "⬇ Download Word CV",
                        word_file,
                        "candidate_cv.docx",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                except Exception as err:
                    st.error(f"Word export error: {err}")

            with e2:
                try:
                    excel_file = build_xlsx(profile, language, urdu_profile)
                    st.download_button(
                        "⬇ Download Excel",
                        excel_file,
                        "candidate_profile.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as err:
                    st.error(f"Excel export error: {err}")

            with e3:
                try:
                    pdf_file = build_pdf(profile, language, urdu_profile)
                    st.download_button(
                        "⬇ Download PDF CV",
                        pdf_file,
                        "candidate_cv.pdf",
                        "application/pdf",
                        use_container_width=True
                    )
                except Exception as err:
                    st.warning(f"PDF export issue: {err}")

            st.markdown("</div>", unsafe_allow_html=True)

            st.info(
                "Important: Urdu content is entered manually. This app does "
                "not automatically translate English into Urdu. Review all "
                "candidate information before using the CV. The parser uses "
                "rule-based detection and does not yet support OCR for scanned "
                "resumes. Complex PDF layouts may affect section detection."
            )

    except Exception as error:
        st.error("Unable to process this resume.")
        st.exception(error)

else:
    st.info("Upload a resume above to start candidate analysis.")


# =====================================================
# PROFESSIONAL FOOTER
# =====================================================
st.markdown("""
<div class="footer">
    <b style="color:#17324D;font-size:13px;">
        Urdu Resume Skill Extractor
    </b><br>
    Intelligent Resume Analysis & Candidate Information Extraction<br>
    Academic Final Year Project | 2026–2027<br>
    Python • Streamlit • PDF • Word • Excel<br><br>
    <b style="color:#176B70;">Designed & Developed by Hina Arshad</b><br>
    <span style="font-size:10px;">
        This tool provides automated assistance. Please verify extracted
        information before making recruitment decisions.
    </span>
</div>
""", unsafe_allow_html=True)
