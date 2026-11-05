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
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Optional PDF/Urdu libraries. The app still runs if they are not installed.
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_RIGHT
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

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
# PROFESSIONAL DESIGN SYSTEM — ORIGINAL DESIGN PRESERVED
# =====================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #F5F7FB; color: #172B4D; }
.block-container { max-width: 1250px; padding-top: 2rem; padding-bottom: 3rem; }
[data-testid="stSidebar"] { background: linear-gradient(180deg, #102A43 0%, #163B59 100%); }
[data-testid="stSidebar"] * { color: #F4F8FC !important; }
[data-testid="stSidebar"] .stCaption { color: #B9CBD9 !important; }
.sidebar-brand { padding: 12px 0 22px 0; }
.sidebar-logo { font-size: 34px; margin-bottom: 8px; }
.sidebar-title { font-size: 19px; font-weight: 800; color: white; line-height: 1.4; }
.sidebar-subtitle { font-size: 12px; color: #B9CBD9; margin-top: 5px; }
.sidebar-section { color: #8FCFC8; font-size: 11px; font-weight: 700; letter-spacing: 1.2px; margin-top: 22px; margin-bottom: 9px; }
.sidebar-item { padding: 9px 11px; background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.08); border-radius: 9px; margin-bottom: 7px; font-size: 12px; }
.hero { background: linear-gradient(120deg, #102A43 0%, #176B70 100%); padding: 30px 32px; border-radius: 18px; color: white; margin-bottom: 24px; box-shadow: 0 8px 25px rgba(16,42,67,0.12); }
.hero h1 { color: white !important; font-size: 31px !important; font-weight: 800 !important; margin: 0 0 10px 0; }
.hero p { color: #E1F1F1; font-size: 14px; margin: 0; line-height: 1.7; }
.eyebrow { color: #A8E2D8; font-size: 11px; font-weight: 700; letter-spacing: 1.5px; margin-bottom: 10px; }
.section-heading { font-size: 19px; font-weight: 750; color: #17324D; margin: 22px 0 13px 0; }
.info-card { background: white; border: 1px solid #E5EAF1; border-radius: 13px; padding: 18px; min-height: 112px; box-shadow: 0 3px 12px rgba(20,45,75,0.035); }
.card-label { color: #718096; font-size: 12px; font-weight: 600; margin-bottom: 10px; }
.card-value { color: #17324D; font-size: 22px; font-weight: 800; }
.content-card { background: white; border: 1px solid #E5EAF1; border-radius: 14px; padding: 20px; margin-bottom: 15px; }
.content-card h3 { color: #17324D; font-size: 16px; margin-top: 0; }
.skill-tag { display: inline-block; background: #E8F5F3; color: #176B70; border: 1px solid #CDE9E4; padding: 6px 11px; border-radius: 20px; margin: 4px 5px 4px 0; font-size: 12px; font-weight: 600; }
.upload-note { background: #EAF2FA; border: 1px solid #D6E5F3; padding: 13px 16px; border-radius: 10px; color: #285477; font-size: 13px; margin-bottom: 15px; }
.stButton > button, .stDownloadButton > button { background: #176B70; color: white; border: none; border-radius: 9px; font-weight: 650; padding: 9px 18px; }
.stButton > button:hover, .stDownloadButton > button:hover { background: #12565A; color: white; border: none; }
[data-testid="stFileUploader"] { background: white; border: 1px dashed #9AB8C8; border-radius: 13px; padding: 14px; }
.footer { border-top: 1px solid #E1E7EF; padding-top: 18px; margin-top: 40px; color: #7A8999; font-size: 11px; text-align: center; }
</style>
""", unsafe_allow_html=True)


# =====================================================
# SKILL DATABASE
# =====================================================
SKILLS = [
    "python", "java", "javascript", "c++", "c#", "html", "css",
    "react", "node.js", "sql", "mysql", "mongodb", "machine learning",
    "deep learning", "artificial intelligence", "data science",
    "streamlit", "fastapi", "django", "flask", "pandas", "numpy",
    "tensorflow", "pytorch", "git", "github", "communication",
    "project management", "problem solving", "ms office"
]
