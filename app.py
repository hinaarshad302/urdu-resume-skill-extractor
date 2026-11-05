
import re
import io
import streamlit as st
import pdfplumber
import pandas as pd
from docx import Document

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
# PROFESSIONAL DESIGN SYSTEM
# =====================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background: #F5F7FB;
    color: #172B4D;
}

.block-container {
    max-width: 1250px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #102A43 0%, #163B59 100%);
}

[data-testid="stSidebar"] * {
    color: #F4F8FC !important;
}

[data-testid="stSidebar"] .stCaption {
    color: #B9CBD9 !important;
}

.sidebar-brand {
    padding: 12px 0 22px 0;
}

.sidebar-logo {
    font-size: 34px;
    margin-bottom: 8px;
}

.sidebar-title {
    font-size: 19px;
    font-weight: 800;
    color: white;
    line-height: 1.4;
}

.sidebar-subtitle {
    font-size: 12px;
    color: #B9CBD9;
    margin-top: 5px;
}

.sidebar-section {
    color: #8FCFC8;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1.2px;
    margin-top: 22px;
    margin-bottom: 9px;
}

.sidebar-item {
    padding: 9px 11px;
    background: rgba(255,255,255,0.07);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 9px;
    margin-bottom: 7px;
    font-size: 12px;
}

.hero {
    background: linear-gradient(120deg, #102A43 0%, #176B70 100%);
    padding: 30px 32px;
    border-radius: 18px;
    color: white;
    margin-bottom: 24px;
    box-shadow: 0 8px 25px rgba(16,42,67,0.12);
}

.hero h1 {
    color: white !important;
    font-size: 31px !important;
    font-weight: 800 !important;
    margin: 0 0 10px 0;
}

.hero p {
    color: #E1F1F1;
    font-size: 14px;
    margin: 0;
    line-height: 1.7;
}

.eyebrow {
    color: #A8E2D8;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1.5px;
    margin-bottom: 10px;
}

.section-heading {
    font-size: 19px;
    font-weight: 750;
    color: #17324D;
    margin: 22px 0 13px 0;
}

.info-card {
    background: white;
    border: 1px solid #E5EAF1;
    border-radius: 13px;
    padding: 18px;
    min-height: 112px;
    box-shadow: 0 3px 12px rgba(20,45,75,0.035);
}

.card-label {
    color: #718096;
    font-size: 12px;
    font-weight: 600;
    margin-bottom: 10px;
}

.card-value {
    color: #17324D;
    font-size: 22px;
    font-weight: 800;
}

.content-card {
    background: white;
    border: 1px solid #E5EAF1;
    border-radius: 14px;
    padding: 20px;
    margin-bottom: 15px;
}

.content-card h3 {
    color: #17324D;
    font-size: 16px;
    margin-top: 0;
}

.skill-tag {
    display: inline-block;
    background: #E8F5F3;
    color: #176B70;
    border: 1px solid #CDE9E4;
    padding: 6px 11px;
    border-radius: 20px;
    margin: 4px 5px 4px 0;
    font-size: 12px;
    font-weight: 600;
}

.upload-note {
    background: #EAF2FA;
    border: 1px solid #D6E5F3;
    padding: 13px 16px;
    border-radius: 10px;
    color: #285477;
    font-size: 13px;
    margin-bottom: 15px;
}

.stButton > button,
.stDownloadButton > button {
    background: #176B70;
    color: white;
    border: none;
    border-radius: 9px;
    font-weight: 650;
    padding: 9px 18px;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background: #12565A;
    color: white;
    border: none;
}

[data-testid="stFileUploader"] {
    background: white;
    border: 1px dashed #9AB8C8;
    border-radius: 13px;
    padding: 14px;
}

.footer {
    border-top: 1px solid #E1E7EF;
    padding-top: 18px;
    margin-top: 40px;
    color: #7A8999;
    font-size: 11px;
    text-align: center;
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
    "data science", "streamlit", "fastapi",
    "django", "flask", "pandas", "numpy",
    "tensorflow", "pytorch", "git", "github",
    "communication", "project management",
    "problem solving", "ms office"
]


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
    return "\n".join(
        paragraph.text for paragraph in document.paragraphs
    )


def extract_txt(file):
    return file.read().decode("utf-8", errors="replace")


# =====================================================
# INFORMATION DETECTION
# =====================================================
def extract_email(text):
    match = re.search(
        r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
        text
    )
    return match.group(0) if match else "Not detected"


def extract_phone(text):
    match = re.search(
        r"(?<!\d)(?:\+92|0092|0)?[\s-]?\d{3}[\s-]?\d{7}(?!\d)",
        text
    )
    return match.group(0) if match else "Not detected"


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
            found.append(skill.title())

    return found


def extract_education(text):
    keywords = [
        "phd", "master", "bachelor", "bs computer science",
        "bs cs", "bscs", "bba", "mba", "intermediate",
        "matric", "associate degree", "diploma"
    ]

    results = []
    for line in text.splitlines():
        if any(word in line.lower() for word in keywords):
            results.append(line.strip())

    return results[:10]


def extract_experience(text):
    keywords = [
        "experience", "employment", "internship",
        "work history", "professional experience"
    ]

    return [
        line.strip()
        for line in text.splitlines()
        if any(word in line.lower() for word in keywords)
    ][:10]


# =====================================================
# SIDEBAR
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
        '<div class="sidebar-item">✓ &nbsp; CSV profile export</div>',
        unsafe_allow_html=True
    )

    st.divider()

    st.caption("ACADEMIC PROJECT")
    st.markdown("**FYP 2026–2027**")
    st.caption("Developed with Python and Streamlit")
    st.caption("Prototype Version 1.0")


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
                "No readable text was found. "
                "This may be a scanned PDF. OCR support "
                "will be introduced in a future development phase."
            )

        else:

            email = extract_email(resume_text)
            phone = extract_phone(resume_text)
            skills = extract_skills(resume_text)
            education = extract_education(resume_text)
            experience = extract_experience(resume_text)

            # -----------------------------------------
            # SUMMARY METRICS
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">02 &nbsp; Candidate Overview</div>',
                unsafe_allow_html=True
            )

            m1, m2, m3 = st.columns(3)

            with m1:
                st.markdown(f"""
                <div class="info-card">
                    <div class="card-label">DETECTED SKILLS</div>
                    <div class="card-value">{len(skills)}</div>
                </div>
                """, unsafe_allow_html=True)

            with m2:
                st.markdown(f"""
                <div class="info-card">
                    <div class="card-label">EDUCATION ENTRIES</div>
                    <div class="card-value">{len(education)}</div>
                </div>
                """, unsafe_allow_html=True)

            with m3:
                st.markdown(f"""
                <div class="info-card">
                    <div class="card-label">EXPERIENCE KEYWORDS</div>
                    <div class="card-value">{len(experience)}</div>
                </div>
                """, unsafe_allow_html=True)

            # -----------------------------------------
            # CONTACT INFORMATION
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
                st.write(email)
                st.markdown("</div>", unsafe_allow_html=True)

            with contact2:
                st.markdown(
                    '<div class="content-card"><h3>📞 Phone Number</h3>',
                    unsafe_allow_html=True
                )
                st.write(phone)
                st.markdown("</div>", unsafe_allow_html=True)

            # -----------------------------------------
            # SKILLS
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">04 &nbsp; Technical Skills</div>',
                unsafe_allow_html=True
            )

            st.markdown('<div class="content-card">', unsafe_allow_html=True)

            if skills:
                skill_html = "".join(
                    f'<span class="skill-tag">{skill}</span>'
                    for skill in skills
                )
                st.markdown(skill_html, unsafe_allow_html=True)
            else:
                st.info("No matching skills detected.")

            st.markdown("</div>", unsafe_allow_html=True)

            # -----------------------------------------
            # EDUCATION AND EXPERIENCE
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
                    st.write("No experience heading detected.")

                st.markdown("</div>", unsafe_allow_html=True)

            # -----------------------------------------
            # ORIGINAL TEXT
            # -----------------------------------------
            with st.expander("📄 View Original Extracted Resume Text"):
                st.text_area(
                    "Extracted document content",
                    resume_text,
                    height=250
                )

            # -----------------------------------------
            # CSV EXPORT
            # -----------------------------------------
            st.markdown(
                '<div class="section-heading">07 &nbsp; Export Candidate Profile</div>',
                unsafe_allow_html=True
            )

            profile = {
                "File Name": uploaded_file.name,
                "Email": email,
                "Phone": phone,
                "Skills": ", ".join(skills),
                "Education": " | ".join(education),
                "Experience": " | ".join(experience)
            }

            df = pd.DataFrame([profile])

            st.markdown('<div class="content-card">', unsafe_allow_html=True)

            st.write(
                "Download the extracted candidate information "
                "as a CSV file for further review."
            )

            st.download_button(
                label="⬇ Download Candidate Profile (CSV)",
                data=df.to_csv(index=False).encode("utf-8-sig"),
                file_name="candidate_profile.csv",
                mime="text/csv"
            )

            st.markdown("</div>", unsafe_allow_html=True)

            st.info(
                "Review all extracted information manually. "
                "This version uses rule-based detection and is not "
                "yet a fully AI-powered or OCR-enabled parser."
            )

    except Exception as error:
        st.error("Unable to process this resume.")
        st.exception(error)

else:
    st.info("Upload a resume above to start candidate analysis.")


# =====================================================
# FOOTER
# =====================================================
st.markdown("""
<div class="footer">
    <b>Urdu Resume Skill Extractor</b><br>
    Academic Final Year Project | 2026–2027<br>
    Python • Streamlit • Document Information Extraction<br><br>
    Prototype output should be reviewed before any recruitment decision.
</div>
""", unsafe_allow_html=True)
