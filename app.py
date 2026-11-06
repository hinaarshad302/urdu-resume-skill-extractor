

import re
import io
import os
import html
import hashlib
import json
import time
import streamlit as st
import pdfplumber
from pypdf import PdfReader
import pandas as pd

from PIL import Image as PILImage, ImageDraw, ImageFont, features

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Flowable
from reportlab.lib.utils import ImageReader
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
    initial_sidebar_state="collapsed"
)

# =====================================================
# GEMINI AI CONFIGURATION
# =====================================================
from google import genai
from google.genai import types


def get_gemini_client():
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        raise ValueError(
            "Gemini API Key missing. Add GEMINI_API_KEY to Streamlit Cloud Secrets."
        )

    return genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(timeout=45000)
    )


def generate_professional_profiles(profile, resume_text):
    """Generate factual English and Urdu CV sections from supplied information."""
    client = get_gemini_client()

    prompt = f"""
You are a professional resume writer and Urdu translator.

Create a polished English CV profile and a natural professional Urdu
translation using ONLY the supplied candidate profile and resume text.

STRICT FACTUAL RULES:
- Never invent education, institutions, degrees, dates, employers, job titles,
  experience, skills, projects, certifications, achievements, or languages.
- Do not infer a skill merely because it is related to a project.
- Do not create experience when the source has none.
- Preserve candidate name, email, and phone exactly as supplied in the profile.
- Improve grammar and organization without changing facts.
- The English object must be a separate, professional English CV.
- The Urdu object must be a separate CV written in Urdu script for all natural-language content.
- In Urdu summary, education, experience, project descriptions, and languages, do not copy full English sentences from the English object.
- In BOTH CVs, divide the skills field into two clearly labelled groups using newline characters:
  English: "Technical Skills: ..." then "Soft Skills: ..."
  Urdu: "تکنیکی مہارتیں: ..." then "ذاتی مہارتیں: ..."
- English technical skills must contain only technical tools, technologies, frameworks, languages, and AI concepts. English soft skills must contain only genuine soft skills.
- In the Urdu skills field, translate every ordinary/soft skill into natural Urdu ONLY. Do not add English translations in brackets or parentheses. For example: Communication = ابلاغی مہارت; Problem Solving = مسئلہ حل کرنے کی مہارت; Teamwork = ٹیم ورک; Time Management = وقت کا انتظام.
- In all Urdu natural-language content, use fluent Urdu script. Avoid English verbs or mixed phrases such as "plan, execute, verify, report", "AI powered", or "learning roadmap" when a natural Urdu equivalent exists. Use منصوبہ بندی، کام انجام دینا، تصدیق کرنا، رپورٹ پیش کرنا، مصنوعی ذہانت پر مبنی، and تعلیمی منصوبہ بندی کا خاکہ as appropriate.
- Keep ONLY official technology names, acronyms, programming languages, frameworks, tool names, and official project/product names in English (for example AI, RAG, FastAPI, Streamlit, Gemini, FAISS, AURA, Hospital RAG Assistant).
- Do not translate official project names, tools, email, phone, candidate name, or official institution names.
- Keep project descriptions in Urdu; do not leave English explanatory sentences inside the Urdu CV.
- If information is absent, return an empty string for that field.
- Return valid JSON only. No Markdown fences and no explanation.

Return this exact JSON structure:
{{
  "english": {{
    "name": "", "job": "", "email": "", "phone": "",
    "summary": "", "skills": "", "education": "",
    "experience": "", "projects": "", "languages": ""
  }},
  "urdu": {{
    "name": "", "job": "", "email": "", "phone": "",
    "summary": "", "skills": "", "education": "",
    "experience": "", "projects": "", "languages": ""
  }}
}}

USER-EDITED CANDIDATE PROFILE:
{json.dumps(profile, ensure_ascii=False)}

ORIGINAL RESUME TEXT:
{resume_text[:26000]}
"""

    # Use lightweight models first for resume extraction and translation.
    # Keep a Flash model as the final fallback.
    response = None
    last_api_error = None
    model_options = (
        "gemini-3.5-flash-lite",
        "gemini-3.1-flash-lite",
        "gemini-3.6-flash",
    )

    for model_name in model_options:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    max_output_tokens=8192,
                    thinking_config=types.ThinkingConfig(
                        thinking_level="low"
                    )
                )
            )
            break
        except Exception as api_error:
            last_api_error = api_error
            error_text = str(api_error).lower()
            temporary_error = any(
                marker in error_text
                for marker in (
                    "503", "429", "500", "502", "504",
                    "unavailable", "resource_exhausted",
                    "high demand", "overloaded", "timeout",
                    "rate limit", "quota"
                )
            )
            if not temporary_error:
                raise

    if response is None:
        detail = str(last_api_error) if last_api_error else "Unknown Gemini API error"
        raise RuntimeError(
            "Gemini API ne tamam available models par request reject ki. "
            "Asal error: " + detail[:1200]
        ) from last_api_error

    raw = (response.text or "").strip()
    raw = re.sub(r"^\s*```(?:json)?\s*", "", raw, flags=re.I)
    raw = re.sub(r"\s*```\s*$", "", raw)

    first_brace = raw.find("{")
    last_brace = raw.rfind("}")
    if first_brace < 0 or last_brace <= first_brace:
        raise ValueError("Gemini ne valid JSON response return nahi kiya.")

    result = json.loads(raw[first_brace:last_brace + 1])

    for language_key in ("english", "urdu"):
        if not isinstance(result.get(language_key), dict):
            raise ValueError("Gemini response mein English/Urdu profile missing hai.")
        for field in (
            "name", "job", "email", "phone", "summary", "skills",
            "education", "experience", "projects", "languages"
        ):
            result[language_key].setdefault(field, "")
            if result[language_key][field] is None:
                result[language_key][field] = ""

    for language_key in ("english", "urdu"):
        result[language_key]["name"] = profile.get("name", "")
        result[language_key]["email"] = profile.get("email", "")
        result[language_key]["phone"] = profile.get("phone", "")

    if profile.get("job", "").strip():
        result["english"]["job"] = profile["job"].strip()

    return result


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
.sidebar-title {font-size:21px;font-weight:800;color:white;line-height:1.4;}
.sidebar-subtitle {font-size:14px;color:#B9CBD9;margin-top:6px;line-height:1.5;}
.sidebar-section {color:#8FCFC8;font-size:12px;font-weight:700;letter-spacing:1.2px;margin-top:22px;margin-bottom:10px;}
.sidebar-item {padding:12px 13px;background:rgba(255,255,255,0.07);border:1px solid rgba(255,255,255,0.08);border-radius:10px;margin-bottom:8px;font-size:14px;line-height:1.45;}
.sidebar-credit {border-top:1px solid rgba(255,255,255,0.2);margin-top:22px;padding-top:15px;text-align:center;color:#A8E2D8;font-size:13px;font-weight:700;}
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
textarea[aria-label*="پیشہ ورانہ"],
textarea[aria-label*="مہارتیں"],
textarea[aria-label*="تعلیم"],
textarea[aria-label*="تجربہ"],
textarea[aria-label*="منصوبے"],
textarea[aria-label*="زبانیں"] {
    direction:rtl !important;
    text-align:right !important;
}
.stButton > button,.stDownloadButton > button {background:#176B70;color:white;border:none;border-radius:9px;font-weight:650;padding:9px 18px;}
.stButton > button:hover,.stDownloadButton > button:hover {background:#12565A;color:white;border:none;}
[data-testid="stFileUploader"] {background:white;border:1px dashed #9AB8C8;border-radius:13px;padding:14px;}
.footer {border-top:1px solid #E1E7EF;padding:18px 12px 8px;margin-top:32px;color:#7A8999;font-size:11px;text-align:center;line-height:1.65;overflow-wrap:anywhere;}
@media (max-width: 640px) {
    /* Keep the main page clear of Streamlit's mobile top toolbar. */
    .block-container {padding:2.35rem 0.85rem 2rem !important;}

    /* Let Streamlit manage the native mobile navigation drawer width. */
    [data-testid="stSidebar"] {
        min-width: 0 !important;
    }

    /* Stack dashboard cards and form columns instead of squeezing them. */
    [data-testid="stHorizontalBlock"] {
        flex-wrap: wrap !important;
        gap: 0.75rem !important;
    }
    [data-testid="stHorizontalBlock"] > [data-testid="column"] {
        flex: 1 1 100% !important;
        width: 100% !important;
        min-width: 100% !important;
    }

    .block-container {max-width:100% !important;}
    .section-heading {font-size:20px;line-height:1.4;}
    .card-label {font-size:13px;}
    .skill-tag {font-size:13px;padding:7px 12px;}

    .hero {padding:21px 18px;border-radius:14px;margin-bottom:18px;}
    .hero h1 {font-size:27px !important;line-height:1.25 !important;}
    .hero p {font-size:15px;line-height:1.6;}
    .eyebrow {font-size:11px;letter-spacing:1px;line-height:1.5;}
    .content-card {padding:16px;}
    .info-card {padding:16px;min-height:auto;}
    .card-value {font-size:21px;}
    .stButton > button,.stDownloadButton > button {width:100%;white-space:normal;padding:10px 12px;}
    .footer {font-size:12px;line-height:1.6;margin-top:26px;padding:15px 8px 6px;}
}
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
    "generative ai", "agentic ai", "ai agents",
    "natural language processing", "computer vision",
    "data science", "data analysis",
    "streamlit", "fastapi", "django", "flask",
    "pandas", "numpy", "tensorflow", "pytorch",
    "git", "github", "google gemini",
    "google gemini api", "gemini api",
    "prompt engineering", "rag",
    "retrieval augmented generation",
    "faiss", "sentence transformers",
    "langchain", "langgraph", "crewai",
    "n8n", "make.com", "automation",
    "ai automation", "api", "rest api",
    "pypdf", "pdfplumber", "gradio",
    "communication", "teamwork",
    "time management", "problem solving",
    "project management", "ms word",
    "ms excel", "ms office", "excel",
    "word", "documentation"
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
    "rest api": "REST API",
    "ms office": "MS Office",
    "ms word": "MS Word",
    "ms excel": "MS Excel",
    "google gemini api": "Google Gemini API",
    "google gemini": "Google Gemini",
    "gemini api": "Gemini API",
    "generative ai": "Generative AI",
    "agentic ai": "Agentic AI",
    "ai agents": "AI Agents",
    "rag": "RAG",
    "faiss": "FAISS",
    "pypdf": "PyPDF",
    "n8n": "n8n",
    "make.com": "Make.com",
    "langchain": "LangChain",
    "langgraph": "LangGraph",
    "crewai": "CrewAI",
    "gradio": "Gradio",
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
    """
    Extract selectable text from a PDF.
    Try pdfplumber first, then use PyPDF as a fallback.
    This does not perform OCR on image-only scanned pages.
    """
    pdf_bytes = file.getvalue() if hasattr(file, "getvalue") else file.read()
    if not pdf_bytes:
        return ""

    extracted = []

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text(
                    layout=False,
                    x_tolerance=2,
                    y_tolerance=3
                ) or ""
                if page_text.strip():
                    extracted.append(page_text)
    except Exception:
        extracted = []

    text = "\n".join(extracted).strip()
    if text:
        return text

    # Fallback for PDFs whose text layer is not handled by pdfplumber.
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)
        fallback_pages = []
        for page in reader.pages:
            page_text = page.extract_text() or ""
            if page_text.strip():
                fallback_pages.append(page_text)
        return "\n".join(fallback_pages).strip()
    except Exception:
        return ""


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
    line = re.sub(r"\(cid:\d+\)", " ", line, flags=re.I)
    line = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", " ", line)
    line = re.sub(r"\s+", " ", line)
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
# CANDIDATE NAME DETECTION
# =====================================================
def extract_name(text):

    ignored_phrases = [
        "resume", "curriculum vitae",
        "personal information", "contact information",
        "professional summary", "career objective",
        "profile", "projects", "education",
        "experience", "technical skills", "skills",
        "phone", "email", "address", "linkedin", "github"
    ]

    for line in get_clean_lines(text):

        value = line.strip()

        if not value or len(value) > 65:
            continue

        value = re.sub(
            r"\s+(projects?|resume|curriculum vitae)$",
            "",
            value,
            flags=re.I
        ).strip()

        if not value:
            continue

        if re.search(r"@|https?://|www\.|\d{4,}|[|]", value, re.I):
            continue

        if any(phrase in value.lower() for phrase in ignored_phrases):
            continue

        words = value.split()

        if not 2 <= len(words) <= 4:
            continue

        if not all(
            re.fullmatch(r"[A-Za-z][A-Za-z.'-]*", word)
            for word in words
        ):
            continue

        if len(value) >= 4:
            return value.title()

    return ""


# =====================================================
# SKILL DETECTION
# =====================================================
SOFT_SKILLS = {
    "communication", "teamwork", "time management", "problem solving",
    "project management", "leadership", "adaptability", "critical thinking",
    "documentation"
}


def extract_skills(text):

    lower_text = text.lower()
    found = []

    # Longer phrases are checked first to avoid duplicate labels.
    for skill in sorted(SKILLS, key=len, reverse=True):

        # Soft skills are displayed in their own category, not as technical skills.
        if skill in SOFT_SKILLS:
            continue

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

            if display_name not in found:
                found.append(display_name)

    # Remove generic duplicate labels when a more specific skill is present.
    specific_to_generic = {
        "Excel": ["MS Excel"],
        "Word": ["MS Word"],
        "API": ["REST API", "Google Gemini API", "Gemini API"]
    }
    # Keep the most descriptive Gemini skill label only.
    if "Google Gemini API" in found:
        found = [x for x in found if x not in ("Google Gemini", "Gemini API")]
    elif "Gemini API" in found and "Google Gemini" in found:
        found.remove("Google Gemini")

    for generic, specific_items in specific_to_generic.items():
        if generic in found and any(item in found for item in specific_items):
            found.remove(generic)

    return found


def extract_soft_skills(text):
    lower_text = text.lower()
    found = []
    for skill in sorted(SOFT_SKILLS, key=len, reverse=True):
        pattern = r"(?<![a-z0-9])" + re.escape(skill) + r"(?![a-z0-9])"
        if re.search(pattern, lower_text):
            label = SKILL_DISPLAY.get(skill, skill.title())
            if label not in found:
                found.append(label)
    return found


# =====================================================
# SECTION-AWARE RESUME EXTRACTION
# =====================================================
def normalize_heading(line):
    value = re.sub(r"^\s*\d{1,2}[.)\-:]?\s*", "", str(line)).strip().lower()
    value = re.sub(r"[:|]+$", "", value).strip()
    return value


SECTION_ALIASES = {
    "projects": {"project", "projects", "my projects", "personal projects", "academic projects"},
    "experience": {"experience", "work experience", "professional experience", "employment history", "work history"},
    "education": {"education", "academic background", "academic qualification", "qualifications"},
    "skills": {"skills", "technical skills", "key skills", "core competencies", "additional skills"},
    "summary": {"profile", "professional summary", "career objective", "summary", "about me"},
    "languages": {"languages", "language"},
}

ALL_SECTION_HEADINGS = set().union(*SECTION_ALIASES.values()) | {
    "certifications", "certificates", "achievements", "awards", "references",
    "interests", "hobbies", "contact", "contact information"
}


def find_heading_in_line(line):
    """Recognize standalone headings and headings appended to a name/header."""
    normalized = normalize_heading(line)
    for heading in sorted(ALL_SECTION_HEADINGS, key=len, reverse=True):
        if normalized == heading:
            return heading
        if len(normalized) <= 65 and normalized.endswith(" " + heading):
            return heading
    return None


def is_section_heading(line):
    return find_heading_in_line(line) is not None


def get_section_lines(text, section_name):
    lines = get_clean_lines(text)
    aliases = SECTION_ALIASES.get(section_name, {section_name})
    active = False
    collected = []
    for line in lines:
        heading = find_heading_in_line(line)
        if heading in aliases:
            active = True
            continue
        if active and heading:
            break
        if active:
            collected.append(line)
    return collected


def remove_contact_noise(value):
    value = clean_extracted_line(value)
    if not value:
        return ""
    value = re.sub(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", "", value, flags=re.I)
    value = re.sub(r"(?:\+92|0092|0)?[\s-]?\d{3}[\s-]?\d{7}", "", value)
    value = re.sub(r"https?://\S+|www\.\S+|linkedin\.com/\S*|github\.com/\S*", "", value, flags=re.I)
    if re.match(r"^(address|location|email|phone|mobile|contact)\s*[:|-]", value.strip(), re.I):
        return ""
    if re.search(r"rahim yar khan|punjab|pakistan", value, re.I):
        return ""
    value = re.sub(r"\s+", " ", value).strip(" |,-")
    return value


def extract_education(text):
    keywords = (
        "phd", "ph.d", "master", "bachelor", "bs computer science",
        "bs cs", "bscs", "bba", "mba", "b.com", "m.com",
        "associate degree", "intermediate", "matric", "secondary school",
        "higher secondary", "diploma", "degree", "computer science",
        "software engineering", "information technology", "university", "college"
    )
    source_lines = get_section_lines(text, "education")
    if not source_lines:
        source_lines = get_clean_lines(text)
    results, seen = [], set()
    for line in source_lines:
        value = remove_contact_noise(line)
        lower = value.lower()
        if not value or not any(k in lower for k in keywords):
            continue
        key = re.sub(r"[^a-z0-9]", "", lower)
        # Merge common duplicate degree-name variants (e.g. BSCS / Bachelor of Science in CS).
        if "computer science" in lower and ("bachelor" in lower or "bscs" in lower or "bs cs" in lower or "bs computer science" in lower):
            key = "bachelorcomputerscience"
        if key and key not in seen:
            seen.add(key)
            results.append(value)
    return results[:10]


def extract_projects(text):
    lines = get_section_lines(text, "projects")
    if not lines:
        return []

    projects = []
    current = []
    title_pattern = re.compile(
        r"^(?:\d{1,2}[.)]?\s+)?(.{3,100})$"
    )
    known_project_words = (
        "aura", "skillgap", "roadmap generator", "complaint sathi",
        "hospital rag", "rag assistant", "generator", "automation system"
    )

    def save_project():
        nonlocal current
        cleaned = []
        for raw in current:
            value = remove_contact_noise(raw)
            if value and value not in cleaned:
                cleaned.append(value)
        if cleaned:
            item = "\n".join(cleaned)
            if item not in projects:
                projects.append(item)
        current = []

    for line in lines:
        value = remove_contact_noise(line)
        if not value:
            continue
        # Ignore education fragments accidentally placed inside a PDF Projects column.
        if re.search(
            r"\b(?:BS|BSc|Bachelor|Master|Degree)\b.*\bComputer Science\b|"
            r"^BS Computer Science$",
            value,
            re.I
        ) and not any(word in value.lower() for word in known_project_words):
            continue
        # Ignore the two-column PDF header fragment above the first project.
        if value.endswith("|") and len(value) < 70:
            continue
        numbered = re.match(r"^\s*\d{1,2}(?:[.)\-:]\s*|\s+)(.+)$", value)
        looks_like_title = bool(numbered) or (
            len(value) <= 100 and
            any(word in value.lower() for word in known_project_words)
        )
        if looks_like_title:
            save_project()
            title = numbered.group(1).strip() if numbered else value
            if not re.fullmatch(r"\d{4}", title):
                current = [title]
        elif current:
            current.append(value)
        elif len(value) >= 8:
            # Some CVs use unnumbered project descriptions; keep the first
            # line as a title only when the Projects section explicitly exists.
            current = [value]

    save_project()
    return projects[:12]


def extract_experience(text):
    lines = get_section_lines(text, "experience")
    if not lines:
        # Conservative fallback: do not scan every line for "assistant",
        # because project names such as "Hospital RAG Assistant" are not jobs.
        return []

    role_pattern = re.compile(
        r"\b(teacher|teaching|data entry|customer service|developer|engineer|"
        r"manager|designer|coordinator|analyst|accountant|intern|consultant|"
        r"officer|administrator|supervisor|technician|operator|lecturer|"
        r"instructor|sales|marketing)\b", re.I
    )
    duration_pattern = re.compile(
        r"\b\d+\s*(?:years?|months?|weeks?)\b", re.I
    )
    action_pattern = re.compile(
        r"\b(managed|supported|handled|responsible|maintained|prepared|"
        r"assisted|organized|provided|conducted|created|trained|taught|"
        r"documented|performed|coordinated|entered|updated)\b", re.I
    )

    results, current = [], []
    started = False
    for line in lines:
        value = remove_contact_noise(line)
        if not value:
            continue

        role_match = role_pattern.search(value)
        if role_match:
            # PDF column extraction can prepend the end of the Profile section
            # to the first job title. Keep only text from the actual role onward.
            value = value[role_match.start():].strip(" |,-")
            # Also remove trailing fragments pulled from the neighboring PDF column.
            value = re.split(
                r"\b(?:projects?|skilled in|strong interest|structured learning)\b",
                value,
                maxsplit=1,
                flags=re.I
            )[0].strip(" |,-")
            started = True
            if current:
                results.append("\n".join(current))
            current = [value]
            continue

        # Never treat summary/profile fragments before the first job title as experience.
        if not started:
            continue
        elif current and (
            duration_pattern.search(value) or action_pattern.search(value)
            or value.startswith(("•", "-", "▪"))
        ):
            current.append(value)
        elif current and len(current) == 1 and len(value) < 140:
            current.append(value)

    if current:
        results.append("\n".join(dict.fromkeys(current)))

    unique = []
    seen = set()
    for item in results:
        key = re.sub(r"\s+", " ", item.lower()).strip()
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique[:10]

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
    "summary",
    "skills",
    "education",
    "experience",
    "projects",
    "languages"
]


# =====================================================
# WORD EXPORT HELPERS - ENGLISH + URDU RTL
# =====================================================
def set_run_font(run, font_name="Arial"):
    """Set both normal and complex-script font properties for Urdu text."""
    run.font.name = font_name
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        r_fonts.set(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + attr,
            font_name
        )
    rtl = r_pr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rtl")
    if rtl is None:
        rtl = OxmlElement("w:rtl")
        r_pr.append(rtl)
    lang = r_pr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lang")
    if lang is None:
        lang = OxmlElement("w:lang")
        r_pr.append(lang)
    lang.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "ur-PK")
    lang.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bidi", "ur-PK")


def set_paragraph_rtl(paragraph):
    """Apply paragraph-level right-to-left direction and Urdu run settings."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_pr = paragraph._p.get_or_add_pPr()
    bidi = p_pr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bidi")
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        p_pr.append(bidi)
    for run in paragraph.runs:
        set_run_font(run)


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

    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(10)

    title = doc.add_heading(profile.get("name") or "Candidate CV", 0)
    if language == "Urdu":
        set_paragraph_rtl(title)

    job_value = (
        urdu_profile.get("job", "")
        if language == "Urdu"
        else profile.get("job", "")
    )
    if job_value:
        job_paragraph = doc.add_paragraph(job_value)
        if language == "Urdu":
            set_paragraph_rtl(job_paragraph)

    contact = " | ".join(
        value for value in [profile.get("email", ""), profile.get("phone", "")]
        if value
    )
    if contact:
        doc.add_paragraph(contact)

    for key in SECTION_KEYS:
        english_value = str(profile.get(key, "") or "").strip()
        urdu_value = str(urdu_profile.get(key, "") or "").strip()

        if language in ("English", "Bilingual") and english_value:
            add_docx_section(doc, LABELS["English"][key], english_value)

        if language in ("Urdu", "Bilingual") and urdu_value:
            add_docx_section(doc, LABELS["Urdu"][key], urdu_value, rtl=True)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("Prepared with Urdu Resume Skill Extractor | Built by Hina Arshad")

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

    contact_language = "Urdu" if language == "Urdu" else "English"
    job_value = (
        urdu_profile.get("job", "")
        if language == "Urdu"
        else profile.get("job", "")
    )
    rows.extend([
        [LABELS[contact_language]["name"], profile.get("name", "")],
        [LABELS[contact_language]["job"], job_value],
        [LABELS[contact_language]["email"], profile.get("email", "")],
        [LABELS[contact_language]["phone"], profile.get("phone", "")]
    ])

    for key in SECTION_KEYS:
        en_value = str(profile.get(key, "") or "").strip()
        ur_value = str(urdu_profile.get(key, "") or "").strip()

        if language == "Urdu":
            if ur_value:
                rows.append([LABELS["Urdu"][key], ur_value])
        elif language == "Bilingual":
            values = [value for value in [en_value, ur_value] if value]
            if values:
                rows.append([
                    LABELS["English"][key] + " / " + LABELS["Urdu"][key],
                    "\n\n".join(values)
                ])
        elif en_value:
            rows.append([LABELS["English"][key], en_value])

    df = pd.DataFrame(rows, columns=["Field", "Candidate Information"])
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Candidate Profile")
        ws = writer.sheets["Candidate Profile"]

        ws.column_dimensions["A"].width = 34
        ws.column_dimensions["B"].width = 68
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
                vertical="center"
            )

        urdu_pattern = re.compile(r"[\\u0600-\\u06FF]")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                contains_urdu = bool(urdu_pattern.search(str(cell.value or "")))
                cell.alignment = Alignment(
                    horizontal="right" if contains_urdu or language in ["Urdu", "Bilingual"] else "left",
                    vertical="top",
                    wrap_text=True,
                    readingOrder=2 if contains_urdu else 0
                )
                if contains_urdu:
                    cell.font = Font(
                        name="Arial",
                        size=11,
                        color="17324D"
                    )
                cell.border = Border(bottom=thin_gray)

        ws.row_dimensions[1].height = 25

    return output.getvalue()

# =====================================================
# PDF EXPORT - IMPROVED URDU RTL + NASTALIQ RENDERING
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

    has_urdu = language in ("Urdu", "Bilingual")
    font_path = os.path.join(
        "assets", "NotoNastaliqUrdu-Regular.ttf"
    )

    if has_urdu:
        if not URDU_SHAPING_AVAILABLE:
            raise ValueError(
                "Urdu packages missing. Check requirements.txt."
            )

        if not os.path.exists(font_path):
            raise ValueError(
                "Noto Nastaliq Urdu font assets folder mein nahi mila."
            )

    styles = getSampleStyleSheet()
    story = []

    # -------------------------------------------------
    # URDU TEXT RENDERING
    # Prefer Pillow RAQM for proper Urdu shaping. If RAQM
    # is unavailable, use arabic_reshaper + python-bidi.
    # -------------------------------------------------

    # Use the stable reshaper + bidi path for Urdu PDF rendering.
    # This avoids device/build-specific RAQM rendering differences.
    URDU_RAQM_AVAILABLE = False

    def make_urdu_font(font_size, scale=3):
        layout = (
            ImageFont.Layout.RAQM
            if URDU_RAQM_AVAILABLE
            else ImageFont.Layout.BASIC
        )
        return ImageFont.truetype(
            font_path,
            int(font_size * scale),
            layout_engine=layout
        )

    def prepare_urdu_text(value):
        value = str(value)
        if URDU_RAQM_AVAILABLE:
            return value
        return get_display(
            arabic_reshaper.reshape(value),
            base_dir="R"
        )

    def urdu_text_bbox(draw, value, font):
        if URDU_RAQM_AVAILABLE:
            return draw.textbbox(
                (0, 0), value, font=font,
                direction="rtl", language="ur"
            )
        return draw.textbbox((0, 0), value, font=font)

    def draw_urdu_text(draw, position, value, font, fill):
        if URDU_RAQM_AVAILABLE:
            draw.text(
                position, value, font=font, fill=fill,
                direction="rtl", language="ur"
            )
        else:
            draw.text(position, value, font=font, fill=fill)

    class UrduLine(Flowable):
        def __init__(
            self, text, font_size=10, color="#222222",
            space_after=5, right_align=True
        ):
            Flowable.__init__(self)
            self.text = str(text)
            self.visual = prepare_urdu_text(self.text)
            self.font_size = font_size
            self.color = color
            self.space_after = space_after
            self.right_align = right_align
            self.scale = 3
            self.font = make_urdu_font(font_size, self.scale)

            dummy = ImageDraw.Draw(PILImage.new("RGB", (10, 10)))
            bbox = urdu_text_bbox(dummy, self.visual, self.font)
            self.text_width = bbox[2] - bbox[0]
            self.text_height = bbox[3] - bbox[1]
            self.text_top = bbox[1]
            self.line_height = max(
                40, (self.text_height + 30) / self.scale
            )
            self.height = self.line_height + self.space_after

        def wrap(self, availWidth, availHeight):
            self.width = availWidth
            return availWidth, self.height

        def draw(self):
            image_width = max(1, int(self.width * self.scale))
            image_height = max(1, int(self.line_height * self.scale))
            image = PILImage.new(
                "RGBA", (image_width, image_height),
                (255, 255, 255, 0)
            )
            draw = ImageDraw.Draw(image)

            if self.right_align:
                x = image_width - self.text_width - 12
            else:
                x = 12
            x = max(0, x)
            y = (image_height - self.text_height) // 2 - self.text_top

            draw_urdu_text(
                draw, (x, y), self.visual, self.font, self.color
            )

            image_buffer = io.BytesIO()
            image.save(image_buffer, format="PNG")
            image_buffer.seek(0)
            self.canv.drawImage(
                ImageReader(image_buffer), 0, self.space_after,
                width=self.width, height=self.line_height, mask="auto"
            )

    # -------------------------------------------------
    # URDU TEXT WRAPPING
    # -------------------------------------------------

    def add_urdu_text(value, font_size=10, color="#222222", space_after=5):
        if value is None or not str(value).strip():
            return

        max_width = doc.width
        scale = 3
        font = make_urdu_font(font_size, scale)
        dummy = ImageDraw.Draw(PILImage.new("RGB", (10, 10)))

        for paragraph in str(value).splitlines():
            paragraph = paragraph.strip()
            if not paragraph:
                continue

            words = paragraph.split()
            current_line = ""

            for word in words:
                test_line = (current_line + " " + word).strip()
                visual = prepare_urdu_text(test_line)
                bbox = urdu_text_bbox(dummy, visual, font)
                text_width = bbox[2] - bbox[0]

                if text_width > max_width * scale and current_line:
                    story.append(
                        UrduLine(current_line, font_size, color, 3)
                    )
                    current_line = word
                else:
                    current_line = test_line

            if current_line:
                story.append(
                    UrduLine(current_line, font_size, color, space_after)
                )

    # -------------------------------------------------
    # ENGLISH PDF STYLES
    # -------------------------------------------------

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
        spaceBefore=12,
        spaceAfter=5
    )

    english_body = ParagraphStyle(
        "EnglishBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=15,
        spaceAfter=4
    )

    # -------------------------------------------------
    # CV HEADER
    # -------------------------------------------------

    story.append(
        Paragraph(
            html.escape(
                str(profile.get("name") or "Candidate CV")
            ),
            title_style
        )
    )

    job_value = (
        urdu_profile.get("job", "")
        if language == "Urdu"
        else profile.get("job", "")
    )

    if job_value:
        if language == "Urdu":
            add_urdu_text(
                job_value,
                font_size=10,
                color="#222222",
                space_after=5
            )
        else:
            story.append(
                Paragraph(
                    html.escape(str(job_value)),
                    english_body
                )
            )

    contact = " | ".join(
        str(x)
        for x in [
            profile.get("email", ""),
            profile.get("phone", "")
        ]
        if x
    )

    if contact:

        story.append(
            Paragraph(
                html.escape(contact),
                english_body
            )
        )

    # -------------------------------------------------
    # CV SECTIONS
    # -------------------------------------------------

    for key in SECTION_KEYS:

        en_value = str(profile.get(key) or "")
        ur_value = str(urdu_profile.get(key) or "")

        if (
            language in ("English", "Bilingual")
            and en_value.strip()
        ):

            story.append(
                Paragraph(
                    html.escape(
                        str(LABELS["English"][key])
                    ),
                    english_heading
                )
            )

            story.append(
                Paragraph(
                    html.escape(en_value).replace(
                        "\n", "<br/>"
                    ),
                    english_body
                )
            )

        if (
            language in ("Urdu", "Bilingual")
            and ur_value.strip()
        ):

            add_urdu_text(
                LABELS["Urdu"][key],
                font_size=12,
                color="#176B70",
                space_after=5
            )

            add_urdu_text(
                ur_value,
                font_size=10,
                color="#222222",
                space_after=8
            )

    # -------------------------------------------------
    # PROFESSIONAL FOOTER
    # -------------------------------------------------

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

        canvas.setFont(
            "Helvetica",
            8
        )

        canvas.setFillColor(
            colors.HexColor("#718096")
        )

        canvas.drawCentredString(
            A4[0] / 2,
            24,
            "Prepared with Urdu Resume Skill Extractor | Built by Hina Arshad"
        )

        canvas.restoreState()

    # -------------------------------------------------
    # BUILD PDF
    # -------------------------------------------------

    doc.build(
        story,
        onFirstPage=add_pdf_footer,
        onLaterPages=add_pdf_footer
    )

    return output.getvalue()


# =====================================================
# SIDEBAR - ORIGINAL DESIGN
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
    st.caption("Prototype Version 1.4")

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

    file_signature = hashlib.sha256(uploaded_file.getvalue()).hexdigest()
    if st.session_state.get("active_resume_signature") != file_signature:
        st.session_state["active_resume_signature"] = file_signature
        st.session_state["ai_generated_profiles"] = None
        for old_widget_key in (
            "candidate_name", "candidate_email", "candidate_phone",
            "target_job", "candidate_summary", "cv_language",
            "candidate_skills", "candidate_soft_skills", "candidate_education",
            "candidate_experience", "candidate_projects",
            "candidate_languages"
        ):
            st.session_state.pop(old_widget_key, None)

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
                "No selectable text could be extracted from this PDF. "
                "It may be image-based, protected, or use an unsupported text layout. "
                "OCR is not available yet. If you can copy the resume text, "
                "save it as a TXT or DOCX file and upload that instead."
            )

        else:

            email = extract_email(resume_text)
            phone = extract_phone(resume_text)
            skills = extract_skills(resume_text)
            soft_skills = extract_soft_skills(resume_text)
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
                    f'<div class="info-card"><div class="card-label">EXPERIENCE ENTRIES</div><div class="card-value">{len(experience)}</div></div>',
                    unsafe_allow_html=True
                )

            with m4:
                st.markdown(
                    f'<div class="info-card"><div class="card-label">PROJECTS DETECTED</div><div class="card-value">{len(projects)}</div></div>',
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

            st.markdown("**Additional / Soft Skills**")
            if soft_skills:
                st.write(" • ".join(soft_skills))
            else:
                st.write("No additional skills detected.")

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

            skills_text = st.text_area(
                "Technical Skills (comma-separated or one per line)",
                value=", ".join(skills),
                height=90,
                key="candidate_skills"
            )

            soft_skills_text = st.text_input(
                "Additional / Soft Skills",
                value=", ".join(soft_skills),
                key="candidate_soft_skills"
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
                "skills": ", ".join(
                    value for value in [skills_text.strip(), soft_skills_text.strip()]
                    if value
                ),
                "education": education_text.strip(),
                "experience": experience_text.strip(),
                "projects": projects_text.strip(),
                "languages": languages_text.strip()
            }

            # -----------------------------------------
            # 09 GEMINI AI: ENGLISH + URDU CV GENERATION
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">09 &nbsp; AI CV Generation</div>',
                unsafe_allow_html=True
            )

            st.caption(
                "Gemini AI uploaded resume aur edited profile ki madad se "
                "professional English aur Urdu CV prepare karega. Missing facts "
                "invent nahi kiye jayenge."
            )

            st.warning(
                "Generate button press karne par resume ka extracted text Google "
                "Gemini API ko process karne ke liye send hoga. Sensitive data "
                "upload karne se pehle privacy ka khayal rakhein."
            )

            if "ai_generated_profiles" not in st.session_state:
                st.session_state["ai_generated_profiles"] = None

            if st.button(
                "✨ Generate English & Urdu CV with Gemini AI",
                type="primary",
                use_container_width=True,
                key="generate_ai_cv"
            ):
                try:
                    with st.spinner(
                        "Gemini AI aapki English aur Urdu CV bana raha hai..."
                    ):
                        st.session_state["ai_generated_profiles"] = (
                            generate_professional_profiles(profile, resume_text)
                        )
                    st.success("English aur Urdu CV generate ho gayi!")
                except Exception as ai_error:
                    st.session_state["ai_generated_profiles"] = None
                    st.error(f"Gemini AI error: {ai_error}")

            ai_profiles = st.session_state.get("ai_generated_profiles")
            urdu_profile = (
                ai_profiles.get("urdu", {})
                if isinstance(ai_profiles, dict)
                else {}
            )
            english_ai_profile = (
                ai_profiles.get("english", {})
                if isinstance(ai_profiles, dict)
                else {}
            )

            if ai_profiles:
                st.markdown("### 👁️ Professional CV Previews")

                # Prepare independent export profiles.
                english_export_profile = english_ai_profile.copy()
                english_export_profile["name"] = profile.get("name", "")
                english_export_profile["email"] = profile.get("email", "")
                english_export_profile["phone"] = profile.get("phone", "")
                if not english_export_profile.get("job"):
                    english_export_profile["job"] = profile.get("job", "")

                st.markdown("## 🇬🇧 English CV Preview")
                st.markdown(
                    f"### {html.escape(str(english_ai_profile.get('name') or profile.get('name') or ''))}"
                )
                if english_ai_profile.get("job"):
                    st.write(english_ai_profile["job"])
                for key in SECTION_KEYS:
                    value = str(english_ai_profile.get(key, "") or "").strip()
                    if value:
                        st.markdown(f"**{LABELS['English'][key]}**")
                        st.write(value)

                st.markdown("### Download English CV")
                st.caption("English preview ke neeche sirf English CV files download karein.")
                en1, en2, en3 = st.columns(3)
                with en1:
                    try:
                        en_docx = build_docx(english_export_profile, "English", {})
                        st.download_button(
                            "⬇ English Word", en_docx,
                            "Hina_Arshad_English_CV.docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True, key="download_english_word"
                        )
                    except Exception as err:
                        st.error(f"English Word export error: {err}")
                with en2:
                    try:
                        en_xlsx = build_xlsx(english_export_profile, "English", {})
                        st.download_button(
                            "⬇ English Excel", en_xlsx,
                            "Hina_Arshad_English_CV.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True, key="download_english_excel"
                        )
                    except Exception as err:
                        st.error(f"English Excel export error: {err}")
                with en3:
                    try:
                        en_pdf = build_pdf(english_export_profile, "English", {})
                        st.download_button(
                            "⬇ English PDF", en_pdf,
                            "Hina_Arshad_English_CV.pdf",
                            "application/pdf",
                            use_container_width=True, key="download_english_pdf"
                        )
                    except Exception as err:
                        st.error(f"English PDF export error: {err}")

                st.divider()
                st.markdown("## اردو سی وی کا پیش نظارہ")
                st.markdown(
                    '<div dir="rtl" style="text-align:right;">',
                    unsafe_allow_html=True
                )
                st.markdown(
                    f"### {html.escape(str(urdu_profile.get('name') or profile.get('name') or ''))}"
                )
                if urdu_profile.get("job"):
                    st.markdown(
                        '<div dir="rtl" style="text-align:right;">'
                        + html.escape(str(urdu_profile["job"]))
                        + '</div>',
                        unsafe_allow_html=True
                    )
                for key in SECTION_KEYS:
                    value = str(urdu_profile.get(key, "") or "").strip()
                    if value:
                        st.markdown(
                            f"**{LABELS['Urdu'][key]}**"
                        )
                        st.markdown(
                            '<div dir="rtl" style="text-align:right;white-space:pre-wrap;">'
                            + html.escape(value)
                            + '</div>',
                            unsafe_allow_html=True
                        )
                st.markdown("</div>", unsafe_allow_html=True)

                st.markdown("### اردو سی وی ڈاؤن لوڈ کریں")
                st.caption(
                    "Neeche wali files mein Urdu CV ka content hoga. "
                    "Naam, email, phone aur official technology/project names apni asal spelling mein rahenge."
                )
                ur1, ur2, ur3 = st.columns(3)
                with ur1:
                    try:
                        ur_docx = build_docx(profile, "Urdu", urdu_profile)
                        st.download_button(
                            "⬇ Urdu Word", ur_docx,
                            "Hina_Arshad_Urdu_CV.docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True, key="download_urdu_word"
                        )
                    except Exception as err:
                        st.error(f"Urdu Word export error: {err}")
                with ur2:
                    try:
                        ur_xlsx = build_xlsx(profile, "Urdu", urdu_profile)
                        st.download_button(
                            "⬇ Urdu Excel", ur_xlsx,
                            "Hina_Arshad_Urdu_CV.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True, key="download_urdu_excel"
                        )
                    except Exception as err:
                        st.error(f"Urdu Excel export error: {err}")
                with ur3:
                    try:
                        ur_pdf = build_pdf(profile, "Urdu", urdu_profile)
                        st.download_button(
                            "⬇ Urdu PDF", ur_pdf,
                            "Hina_Arshad_Urdu_CV.pdf",
                            "application/pdf",
                            use_container_width=True, key="download_urdu_pdf"
                        )
                    except Exception as err:
                        st.error(f"Urdu PDF export error: {err}")
            else:
                st.info(
                    "Pehle 'Generate English & Urdu CV with Gemini AI' button press karein. "
                    "Uske baad English preview, Urdu preview aur unki separate download files nazar aayengi."
                )

            st.info(
                "Important: Download se pehle tamam details verify karein. "
                "AI missing education, experience, dates ya achievements invent nahi karna chahiye. "
                "Scanned PDFs ke liye OCR abhi available nahi; PDF column layouts extraction ko affect kar sakte hain."
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
    <b style="color:#17324D;font-size:13px;">Urdu Resume Skill Extractor</b><br>
    AI-Powered Resume Analysis & Skill Extraction<br>
    Final Year Project | 2026–2027<br>
    Python · Streamlit · PDF · Word · Excel<br>
    <b style="color:#176B70;">Developed by Hina Arshad</b><br>
    <span style="font-size:10px;">
        Verify all details before downloading. Scanned PDFs (OCR) are not supported;
        complex PDF layouts may affect extraction.
    </span>
</div>
""", unsafe_allow_html=True)
