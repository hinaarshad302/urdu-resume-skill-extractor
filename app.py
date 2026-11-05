
import re
import io
import os
import html
import streamlit as st
import pdfplumber
import pandas as pd

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.platypus import SimpleDocTemplate, Paragraph
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
    "python", "java", "javascript", "c++", "c#",
    "html", "css", "react", "node.js", "sql",
    "mysql", "mongodb", "machine learning",
    "deep learning", "artificial intelligence",
    "data science", "streamlit", "fastapi",
    "django", "flask", "pandas", "numpy",
    "tensorflow", "pytorch", "git", "github",
    "communication", "project management",
    "problem solving", "ms office"
]

SKILL_DISPLAY = {
    "javascript": "JavaScript",
    "html": "HTML",
    "css": "CSS",
    "fastapi": "FastAPI",
    "node.js": "Node.js",
    "mysql": "MySQL",
    "mongodb": "MongoDB",
    "github": "GitHub",
    "python": "Python",
    "c++": "C++",
    "c#": "C#",
    "sql": "SQL",
    "api": "API",
    "ms office": "MS Office"
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
    return "\n".join(p.text for p in document.paragraphs)


def extract_txt(file):
    return file.read().decode("utf-8", errors="replace")


# =====================================================
# TEXT CLEANING
# =====================================================
def clean_extracted_line(line):
    if not line:
        return ""

    line = str(line)

    # Remove PDF character encoding artifacts
    line = re.sub(r"\(cid:\d+\)", " ", line, flags=re.I)

    # Remove control characters
    line = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", " ", line)

    # Normalize spaces
    line = re.sub(r"\s+", " ", line)

    # Remove unnecessary separators
    line = line.strip(" \t|•·-–—:;,")

    return line.strip()


def get_clean_lines(text):
    lines = []

    for line in text.splitlines():
        cleaned = clean_extracted_line(line)

        if cleaned:
            lines.append(cleaned)

    return lines


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


# =====================================================
# IMPROVED CANDIDATE NAME DETECTION
# =====================================================
def extract_name(text):

    ignored_phrases = [
        "resume",
        "curriculum vitae",
        "personal information",
        "contact information",
        "professional summary",
        "career objective",
        "profile",
        "projects",
        "education",
        "experience",
        "technical skills",
        "skills",
        "phone",
        "email",
        "address",
        "linkedin",
        "github"
    ]

    for line in get_clean_lines(text):

        value = line.strip()

        if not value or len(value) > 65:
            continue

        # Remove common heading words at the end
        value = re.sub(
            r"\s+(projects?|resume|curriculum vitae)$",
            "",
            value,
            flags=re.I
        ).strip()

        if not value:
            continue

        if re.search(
            r"@|https?://|www\.|\d{4,}|[|]",
            value,
            re.I
        ):
            continue

        if any(
            phrase in value.lower()
            for phrase in ignored_phrases
        ):
            continue

        # Candidate name should contain alphabetic words
        words = value.split()

        if not 2 <= len(words) <= 4:
            continue

        if not all(
            re.fullmatch(r"[A-Za-z][A-Za-z.'-]*", word)
            for word in words
        ):
            continue

        # Avoid common section labels
        if len(value) < 4:
            continue

        return value.title()

    return ""


# =====================================================
# SKILL DETECTION
# =====================================================
def extract_skills(text):
    lower_text = text.lower()
    found = []

    for skill in SKILLS:
        pattern = (
            r"(?<![a-z0-9])"
            + re.escape(skill)
            + r"(?![a-z0-9])"
        )

        if re.search(pattern, lower_text):
            display_name = SKILL_DISPLAY.get(
                skill,
                skill.title()
            )
            found.append(display_name)

    return found


# =====================================================
# IMPROVED EDUCATION EXTRACTION
# =====================================================
def extract_education(text):

    education_keywords = [
        "phd",
        "ph.d",
        "master",
        "bachelor",
        "bs computer science",
        "bs cs",
        "bscs",
        "bba",
        "mba",
        "b.com",
        "m.com",
        "associate degree",
        "intermediate",
        "matric",
        "secondary school",
        "higher secondary",
        "diploma",
        "degree",
        "computer science",
        "software engineering",
        "information technology",
        "university",
        "college"
    ]

    ignored_patterns = [
        r"^english\s*[-:|]",
        r"^urdu\s*[-:|]",
        r"^language\s*[-:|]",
        r"^languages\s*[-:|]",
        r"^skills?\s*[-:|]",
        r"^experience\s*[-:|]"
    ]

    results = []

    for line in get_clean_lines(text):

        lower_line = line.lower()

        if any(
            re.search(pattern, lower_line)
            for pattern in ignored_patterns
        ):
            continue

        if not any(
            keyword in lower_line
            for keyword in education_keywords
        ):
            continue

        # Ignore lines that only mention a language
        if re.fullmatch(
            r"(english|urdu|arabic|punjabi)\s*.*",
            lower_line
        ):
            continue

        if len(line) < 4:
            continue

        if line not in results:
            results.append(line)

    return results[:10]


# =====================================================
# IMPROVED EXPERIENCE EXTRACTION
# =====================================================
def extract_experience(text):

    lines = get_clean_lines(text)

    section_headings = [
        "experience",
        "work experience",
        "professional experience",
        "employment history",
        "work history",
        "career history",
        "internship experience",
        "internships",
        "employment"
    ]

    stop_headings = [
        "education",
        "academic background",
        "qualification",
        "qualifications",
        "technical skills",
        "skills",
        "projects",
        "certifications",
        "certificates",
        "languages",
        "references",
        "achievements",
        "awards",
        "personal information",
        "contact information",
        "professional summary",
        "career objective"
    ]

    action_keywords = [
        "worked",
        "working",
        "developed",
        "designed",
        "managed",
        "created",
        "built",
        "implemented",
        "maintained",
        "responsible",
        "responsibility",
        "intern",
        "internship",
        "freelance",
        "volunteer",
        "experience",
        "employment",
        "developer",
        "engineer",
        "assistant",
        "teacher",
        "designer",
        "coordinator",
        "trained",
        "led",
        "served",
        "operated"
    ]

    results = []
    collecting = False

    for line in lines:

        lower_line = line.lower().strip()

        # Detect start of an experience section
        is_heading = lower_line.rstrip(":") in section_headings

        if is_heading:
            collecting = True
            continue

        # Stop when another major section begins
        if collecting and lower_line.rstrip(":") in stop_headings:
            collecting = False
            continue

        if collecting:

            # Ignore very short lines
            if len(line) < 3:
                continue

            if line not in results:
                results.append(line)

    # Fallback when a separate experience heading is not present
    if not results:

        for line in lines:

            lower_line = line.lower()

            # Do not treat section labels alone as experience
            if lower_line.rstrip(":") in section_headings:
                continue

            if any(
                keyword in lower_line
                for keyword in action_keywords
            ):
                if len(line) >= 5 and line not in results:
                    results.append(line)

    return results[:15]


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
        "languages": "زبانیں"
    }
}

SECTION_KEYS = [
    "summary",
    "skills",
    "education",
    "experience",
    "languages"
]


# =====================================================
# WORD EXPORT HELPERS
# =====================================================
def set_paragraph_rtl(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_pr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    p_pr.append(bidi)


def add_docx_section(doc, heading, value, rtl=False):
    if not str(value or "").strip():
        return

    h = doc.add_heading(heading, level=1)
    p = doc.add_paragraph(str(value))

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

    if language == "Urdu":

        title = doc.add_heading(profile["name"] or "CV", 0)
        set_paragraph_rtl(title)

        if profile["job"]:
            p = doc.add_paragraph(profile["job"])
            set_paragraph_rtl(p)

        contact = " | ".join(
            x for x in [profile["email"], profile["phone"]] if x
        )

        if contact:
            doc.add_paragraph(contact)

        for key in SECTION_KEYS:
            add_docx_section(
                doc,
                LABELS["Urdu"][key],
                urdu_profile.get(key, ""),
                rtl=True
            )

    elif language == "Bilingual":

        doc.add_heading(
            profile["name"] or "Candidate CV",
            0
        )

        if profile["job"]:
            doc.add_paragraph(profile["job"])

        contact = " | ".join(
            x for x in [profile["email"], profile["phone"]] if x
        )

        if contact:
            doc.add_paragraph(contact)

        for key in SECTION_KEYS:
            add_docx_section(
                doc,
                LABELS["English"][key],
                profile.get(key, "")
            )

            add_docx_section(
                doc,
                LABELS["Urdu"][key],
                urdu_profile.get(key, ""),
                rtl=True
            )

    else:

        doc.add_heading(
            profile["name"] or "Candidate CV",
            0
        )

        if profile["job"]:
            doc.add_paragraph(profile["job"])

        contact = " | ".join(
            x for x in [profile["email"], profile["phone"]] if x
        )

        if contact:
            doc.add_paragraph(contact)

        for key in SECTION_KEYS:
            add_docx_section(
                doc,
                LABELS["English"][key],
                profile.get(key, "")
            )

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run(
        "Prepared with Urdu Resume Skill Extractor | Built by Hina Arshad"
    )

    output = io.BytesIO()
    doc.save(output)

    return output.getvalue()


# =====================================================
# EXCEL EXPORT
# =====================================================
def build_xlsx(profile, language, urdu_profile=None):

    from openpyxl.styles import (
        Font,
        PatternFill,
        Alignment,
        Border,
        Side
    )

    urdu_profile = urdu_profile or {}
    rows = []

    if language == "Urdu":

        rows.extend([
            [LABELS["Urdu"]["name"], profile["name"]],
            [LABELS["Urdu"]["job"], profile["job"]],
            [LABELS["Urdu"]["email"], profile["email"]],
            [LABELS["Urdu"]["phone"], profile["phone"]]
        ])

        for key in SECTION_KEYS:
            rows.append([
                LABELS["Urdu"][key],
                urdu_profile.get(key, "")
            ])

    elif language == "Bilingual":

        rows.extend([
            ["Full Name / پورا نام", profile["name"]],
            ["Target Job / مطلوبہ عہدہ", profile["job"]],
            ["Email / ای میل", profile["email"]],
            ["Phone / فون نمبر", profile["phone"]]
        ])

        for key in SECTION_KEYS:
            rows.append([
                LABELS["English"][key] + " / " + LABELS["Urdu"][key],
                profile.get(key, "") + "\n\n" + urdu_profile.get(key, "")
            ])

    else:

        rows.extend([
            [LABELS["English"]["name"], profile["name"]],
            [LABELS["English"]["job"], profile["job"]],
            [LABELS["English"]["email"], profile["email"]],
            [LABELS["English"]["phone"], profile["phone"]]
        ])

        for key in SECTION_KEYS:
            rows.append([
                LABELS["English"][key],
                profile.get(key, "")
            ])

    df = pd.DataFrame(
        rows,
        columns=["Field", "Candidate Information"]
    )

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Candidate Profile"
        )

        ws = writer.sheets["Candidate Profile"]

        ws.column_dimensions["A"].width = 34
        ws.column_dimensions["B"].width = 68

        ws.sheet_view.rightToLeft = language in [
            "Urdu",
            "Bilingual"
        ]

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        header_fill = PatternFill(
            "solid",
            fgColor="176B70"
        )

        thin_gray = Side(
            style="thin",
            color="E5EAF1"
        )

        for cell in ws[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF",
                size=11
            )

            cell.fill = header_fill

            cell.alignment = Alignment(
                horizontal="right" if language == "Urdu" else "left",
                vertical="center"
            )

        for row in ws.iter_rows(min_row=2):

            for cell in row:

                cell.alignment = Alignment(
                    horizontal="right" if language in [
                        "Urdu",
                        "Bilingual"
                    ] else "left",
                    vertical="top",
                    wrap_text=True
                )

                cell.border = Border(
                    bottom=thin_gray
                )

        ws.row_dimensions[1].height = 25

    return output.getvalue()


# =====================================================
# PDF EXPORT
# Urdu font required:
# assets/NotoNastaliqUrdu-Regular.ttf
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
                "Urdu PDF ke liye arabic-reshaper aur python-bidi "
                "requirements.txt mein add karein."
            )

        font_path = os.path.join(
            "assets",
            "NotoNastaliqUrdu-Regular.ttf"
        )

        if not os.path.exists(font_path):
            raise ValueError(
                "Urdu PDF font missing hai. GitHub ke assets folder mein "
                "NotoNastaliqUrdu-Regular.ttf upload karein."
            )

        if "UrduFont" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(
                TTFont("UrduFont", font_path)
            )

        font_name = "UrduFont"

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "CVTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold" if language == "English" else font_name,
        fontSize=19,
        leading=28,
        textColor=colors.HexColor("#102A43"),
        alignment=TA_RIGHT if language == "Urdu" else TA_LEFT,
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
        fontSize=12,
        leading=24,
        textColor=colors.HexColor("#176B70"),
        alignment=TA_RIGHT,
        spaceBefore=12,
        spaceAfter=7
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
        fontSize=10,
        leading=24,
        alignment=TA_RIGHT
    )

    def safe_text(value):
        return html.escape(
            str(value or "")
        ).replace("\n", "<br/>")

    def shape_urdu(value):

        if not value:
            return ""

        if not URDU_SHAPING_AVAILABLE:
            return safe_text(value)

        shaped_lines = []

        for line in str(value).splitlines():

            escaped = html.escape(line)

            if escaped.strip():

                shaped_lines.append(
                    get_display(
                        arabic_reshaper.reshape(escaped)
                    )
                )

            else:
                shaped_lines.append("&nbsp;")

        return "<br/>".join(shaped_lines)

    story = []

    if language == "Urdu":

        story.append(
            Paragraph(
                safe_text(profile["name"] or "CV"),
                title_style
            )
        )

        if profile["job"]:
            story.append(
                Paragraph(
                    safe_text(profile["job"]),
                    urdu_body
                )
            )

        contact = " | ".join(
            x for x in [profile["email"], profile["phone"]] if x
        )

        if contact:
            story.append(
                Paragraph(
                    safe_text(contact),
                    english_body
                )
            )

        for key in SECTION_KEYS:

            value = urdu_profile.get(key, "")

            if value.strip():

                story.append(
                    Paragraph(
                        shape_urdu(LABELS["Urdu"][key]),
                        urdu_heading
                    )
                )

                story.append(
                    Paragraph(
                        shape_urdu(value),
                        urdu_body
                    )
                )

    elif language == "Bilingual":

        story.append(
            Paragraph(
                safe_text(profile["name"] or "Candidate CV"),
                title_style
            )
        )

        if profile["job"]:
            story.append(
                Paragraph(
                    safe_text(profile["job"]),
                    english_body
                )
            )

        contact = " | ".join(
            x for x in [profile["email"], profile["phone"]] if x
        )

        if contact:
            story.append(
                Paragraph(
                    safe_text(contact),
                    english_body
                )
            )

        for key in SECTION_KEYS:

            english_value = profile.get(key, "")
            urdu_value = urdu_profile.get(key, "")

            if english_value.strip():

                story.append(
                    Paragraph(
                        safe_text(LABELS["English"][key]),
                        english_heading
                    )
                )

                story.append(
                    Paragraph(
                        safe_text(english_value),
                        english_body
                    )
                )

            if urdu_value.strip():

                story.append(
                    Paragraph(
                        shape_urdu(LABELS["Urdu"][key]),
                        urdu_heading
                    )
                )

                story.append(
                    Paragraph(
                        shape_urdu(urdu_value),
                        urdu_body
                    )
                )

    else:

        story.append(
            Paragraph(
                safe_text(profile["name"] or "Candidate CV"),
                title_style
            )
        )

        if profile["job"]:
            story.append(
                Paragraph(
                    safe_text(profile["job"]),
                    english_body
                )
            )

        contact = " | ".join(
            x for x in [profile["email"], profile["phone"]] if x
        )

        if contact:
            story.append(
                Paragraph(
                    safe_text(contact),
                    english_body
                )
            )

        for key in SECTION_KEYS:

            value = profile.get(key, "")

            if value.strip():

                story.append(
                    Paragraph(
                        safe_text(LABELS["English"][key]),
                        english_heading
                    )
                )

                story.append(
                    Paragraph(
                        safe_text(value),
                        english_body
                    )
                )

    def add_pdf_footer(canvas, document):

        canvas.saveState()

        canvas.setStrokeColor(
            colors.HexColor("#DCE5EC")
        )

        canvas.line(
            42,
            38,
            A4[0] - 42,
            38
        )

        canvas.setFont("Helvetica", 8)

        canvas.setFillColor(
            colors.HexColor("#718096")
        )

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
# SIDEBAR - ORIGINAL DESIGN + CREDIT
# =====================================================
with st.sidebar:

    st.markdown("""
    <div class="sidebar-brand">
        <div class="sidebar-logo">📄</div>
        <div class="sidebar-title">Resume Intelligence</div>
        <div class="sidebar-subtitle">
            Urdu Resume Skill Extractor
        </div>
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
        '<div class="sidebar-item">✓ &nbsp; Basic skill detection</div>',
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
    st.caption("Prototype Version 1.1")

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
        Extract contact details, education, experience and technical skills
        through a simple and organized workflow.
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
        <div style="color:#718096;font-size:12px;">
            Select a resume document
        </div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div class="info-card">
        <div class="card-label">STEP 02</div>
        <div class="card-value">🔍 Analyze</div>
        <div style="color:#718096;font-size:12px;">
            Detect candidate information
        </div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown("""
    <div class="info-card">
        <div class="card-label">STEP 03</div>
        <div class="card-value">📥 Export</div>
        <div style="color:#718096;font-size:12px;">
            Download structured profile
        </div>
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
            detected_name = extract_name(resume_text)

            # -----------------------------------------
            # 02 CANDIDATE OVERVIEW
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">02 &nbsp; Candidate Overview</div>',
                unsafe_allow_html=True
            )

            m1, m2, m3 = st.columns(3)

            with m1:
                st.markdown(
                    f'<div class="info-card"><div class="card-label">DETECTED SKILLS</div><div class="card-value">{len(skills)}</div></div>',
                    unsafe_allow_html=True
                )

            with m2:
                st.markdown(
                    f'<div class="info-card"><div class="card-label">EDUCATION ENTRIES</div><div class="card-value">{len(education)}</div></div>',
                    unsafe_allow_html=True
                )

            with m3:
                st.markdown(
                    f'<div class="info-card"><div class="card-label">EXPERIENCE KEYWORDS</div><div class="card-value">{len(experience)}</div></div>',
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

            st.markdown(
                '<div class="content-card">',
                unsafe_allow_html=True
            )

            if skills:

                skill_html = "".join(
                    f'<span class="skill-tag">{html.escape(skill)}</span>'
                    for skill in skills
                )

                st.markdown(
                    skill_html,
                    unsafe_allow_html=True
                )

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
                        st.write("• " + item)

                else:
                    st.write("No experience details detected.")

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
            # 07 EDITABLE CANDIDATE PROFILE
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">07 &nbsp; Edit Candidate Profile</div>',
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
                value="\n".join(experience),
                height=100,
                key="candidate_experience"
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
                "languages": languages_text.strip()
            }

            # -----------------------------------------
            # 08 MANUAL URDU CONTENT
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">08 &nbsp; Urdu CV Content</div>',
                unsafe_allow_html=True
            )

            st.caption(
                "Agar Urdu ya Bilingual CV banana hai to neeche Urdu "
                "content manually enter karein. Name, email aur phone "
                "aapke original profile se liye jayenge."
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

                urdu_profile["languages"] = st.text_area(
                    "زبانیں",
                    key="urdu_languages",
                    height=70,
                    placeholder="مثلاً اردو، انگریزی"
                )

            # -----------------------------------------
            # 09 EXPORT PROFESSIONAL CV
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">09 &nbsp; Export Professional CV</div>',
                unsafe_allow_html=True
            )

            st.markdown(
                '<div class="content-card">',
                unsafe_allow_html=True
            )

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
                        "Urdu fields abhi incomplete hain. Behtar result "
                        "ke liye Urdu content enter karein. Empty fields "
                        "CV mein include nahi honge."
                    )

            e1, e2, e3 = st.columns(3)

            with e1:

                try:

                    word_file = build_docx(
                        profile,
                        language,
                        urdu_profile
                    )

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

                    excel_file = build_xlsx(
                        profile,
                        language,
                        urdu_profile
                    )

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

                    pdf_file = build_pdf(
                        profile,
                        language,
                        urdu_profile
                    )

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
                "Important: Urdu content is entered manually. This app "
                "does not automatically translate English into Urdu. "
                "Review all candidate information before using the CV. "
                "The current parser uses rule-based detection and does "
                "not yet support OCR for scanned resumes."
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
