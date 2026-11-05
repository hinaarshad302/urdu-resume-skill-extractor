
import re
import io
import streamlit as st
import pdfplumber
import pandas as pd
from docx import Document

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Urdu Resume Skill Extractor",
    page_icon="📄",
    layout="wide"
)

# ---------------- CUSTOM STYLE ----------------
st.markdown("""
<style>
.main {
    background-color: #f7f9fc;
}
.block-container {
    padding-top: 2rem;
}
h1 {
    color: #16345c;
}
</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------
st.title("📄 Urdu Resume Skill Extractor")
st.markdown(
    "### AI-Powered Resume Information Extraction System"
)
st.write(
    "Upload a resume to extract candidate information, "
    "education, experience and technical skills."
)

st.info(
    "Initial prototype: PDF, DOCX and TXT text extraction "
    "with basic information detection."
)

# ---------------- SKILL DATABASE ----------------
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

# ---------------- TEXT EXTRACTION ----------------
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


# ---------------- BASIC INFORMATION DETECTION ----------------
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
        pattern = r"(?<![a-z0-9])" + re.escape(skill) + r"(?![a-z0-9])"
        if re.search(pattern, lower_text):
            found.append(skill.title())

    return found


def extract_education(text):
    keywords = [
        "phd", "master", "bachelor", "bs computer science",
        "bs cs", "bscs", "bba", "mba", "intermediate",
        "matric", "associate degree", "diploma"
    ]

    lines = text.splitlines()
    results = []

    for line in lines:
        if any(word in line.lower() for word in keywords):
            results.append(line.strip())

    return results[:10]


def extract_experience(text):
    lines = text.splitlines()
    keywords = [
        "experience", "employment", "internship",
        "work history", "professional experience"
    ]

    results = [
        line.strip() for line in lines
        if any(word in line.lower() for word in keywords)
    ]

    return results[:10]


# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.header("About Project")
    st.write(
        "An academic prototype designed to extract "
        "structured information from Urdu, English "
        "and mixed-language resumes."
    )

    st.divider()
    st.caption("FYP Project | 2026–2027")
    st.caption("Python + Streamlit")


# ---------------- RESUME UPLOAD ----------------
st.subheader("1. Upload Your Resume")

uploaded_file = st.file_uploader(
    "Choose a resume file",
    type=["pdf", "docx", "txt"],
    help="Supported formats: PDF, DOCX and TXT"
)

if uploaded_file:

    try:
        file_name = uploaded_file.name.lower()

        if file_name.endswith(".pdf"):
            resume_text = extract_pdf(uploaded_file)

        elif file_name.endswith(".docx"):
            resume_text = extract_docx(uploaded_file)

        else:
            resume_text = extract_txt(uploaded_file)

        st.success("Resume uploaded successfully!")

        if not resume_text.strip():
            st.warning(
                "No readable text was found. "
                "This may be a scanned PDF. OCR support "
                "will be added in a later phase."
            )

        else:
            st.subheader("2. Extracted Candidate Information")

            email = extract_email(resume_text)
            phone = extract_phone(resume_text)
            skills = extract_skills(resume_text)
            education = extract_education(resume_text)
            experience = extract_experience(resume_text)

            col1, col2 = st.columns(2)

            with col1:
                st.metric("Detected Skills", len(skills))
                st.write("**Email Address**")
                st.write(email)

            with col2:
                st.metric("Education Entries", len(education))
                st.write("**Phone Number**")
                st.write(phone)

            st.divider()

            st.subheader("Technical Skills")

            if skills:
                st.write(", ".join(skills))
            else:
                st.write("No skills detected.")

            st.subheader("Education")

            if education:
                for item in education:
                    st.write("• " + item)
            else:
                st.write("No education information detected.")

            st.subheader("Experience")

            if experience:
                for item in experience:
                    st.write("• " + item)
            else:
                st.write("No experience heading detected.")

            st.divider()

            with st.expander("View Original Extracted Text"):
                st.text_area(
                    "Resume Text",
                    resume_text,
                    height=300
                )

            # ---------------- EXPORT ----------------
            st.subheader("3. Export Candidate Profile")

            profile = {
                "Email": email,
                "Phone": phone,
                "Skills": ", ".join(skills),
                "Education": " | ".join(education),
                "Experience": " | ".join(experience)
            }

            df = pd.DataFrame([profile])

            st.download_button(
                label="Download Extracted Profile (CSV)",
                data=df.to_csv(index=False).encode("utf-8-sig"),
                file_name="candidate_profile.csv",
                mime="text/csv"
            )

            st.caption(
                "Note: This is a preliminary rule-based prototype. "
                "Detected information must be manually verified."
            )

    except Exception as error:
        st.error("Unable to process this resume.")
        st.exception(error)

else:
    st.info("Please upload a resume to begin.")
