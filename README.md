# Urdu Resume Skill Extractor

**An intelligent resume analysis and candidate information extraction tool with English, Urdu, and bilingual CV support.**

## Project Overview

Urdu Resume Skill Extractor is a web-based application developed using Python and Streamlit. It helps users extract important information from resumes, identify candidate skills, and organize professional profiles.

The application supports English and Urdu CV content, allowing users to create and export professional resumes in multiple formats.

This project aims to make resume analysis and candidate information management easier, especially for users working with Urdu and English resumes.

## Key Features

### 1. Resume Upload
- Upload PDF documents.
- Upload DOCX documents.
- Upload TXT documents.
- Extract readable text from uploaded resumes.

### 2. Resume Information Extraction
- Candidate name detection.
- Email address detection.
- Phone number detection.
- Basic skill identification.
- Education information extraction.
- Experience keyword extraction.

### 3. Editable Candidate Profile
- Edit candidate name and contact details.
- Add target job position.
- Update professional summary.
- Edit skills, education, and experience.
- Add language information.

### 4. English and Urdu CV Support
- English CV content.
- Urdu CV content.
- Bilingual CV content.
- Manual Urdu content entry.
- Right-to-Left (RTL) formatting support.

### 5. Professional CV Export

Export candidate profiles into three formats:

| Format | Supported |
|---|---|
| Microsoft Word (.docx) | Yes |
| PDF (.pdf) | Yes |
| Excel (.xlsx) | Yes |

## Application Screenshots

Add your application screenshots here:

```text
assets/
├── dashboard.png
├── candidate-analysis.png
└── profile-export.png
```

## Technology Stack

- **Python** – Core programming language
- **Streamlit** – Web application framework
- **Pandas** – Data handling
- **pdfplumber** – PDF text extraction
- **pypdf** – PDF processing
- **python-docx** – Word document generation
- **ReportLab** – PDF generation
- **openpyxl** – Excel file generation
- **arabic-reshaper** – Urdu/Arabic text shaping
- **python-bidi** – Bidirectional text handling
- **Pillow** – Image processing support

## Project Structure

```text
urdu-resume-skill-extractor/
│
├── app.py
├── requirements.txt
├── README.md
│
└── assets/
    └── NotoNastaliqUrdu-Regular.ttf
```

*Note: The structure above shows the main application files. Additional folders and assets may be present in the repository.*

## Installation and Setup

### Step 1: Clone the Repository

```bash
git clone https://github.com/hinaarshad302/urdu-resume-skill-extractor.git
```

### Step 2: Open Project Directory

```bash
cd urdu-resume-skill-extractor
```

### Step 3: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 4: Run the Application

```bash
streamlit run app.py
```

The application will open in your web browser.

## How to Use

1. Open the Urdu Resume Skill Extractor application.
2. Upload a resume in PDF, DOCX, or TXT format.
3. Allow the application to extract available information.
4. Review the candidate overview and detected skills.
5. Edit candidate information in the profile section.
6. Enter Urdu CV content manually when required.
7. Select English, Urdu, or Bilingual CV.
8. Export the professional CV as Word, PDF, or Excel.

## Live Demo

**Try the Application:**

https://urdu-resume-skill-extractor-3vjtecyvnp6rgfuqgf7t8w.streamlit.app/

## GitHub Repository

https://github.com/hinaarshad302/urdu-resume-skill-extractor

## Current Limitations

- Information extraction uses basic rule-based techniques.
- Skill detection depends on predefined keywords.
- Scanned PDFs may not provide readable text without OCR.
- Urdu CV content is entered manually; automatic translation is not included.
- Extracted candidate information should be reviewed and corrected before exporting.
- Results may vary depending on resume formatting and document quality.

## Future Improvements

- AI-powered resume analysis.
- Advanced skill extraction using NLP.
- OCR support for scanned resumes.
- Improved candidate name and education detection.
- Job description and resume matching.
- Resume scoring and recommendations.
- Support for additional languages.
- Advanced Urdu document formatting.

## Project Information

| Detail | Information |
|---|---|
| Project Name | Urdu Resume Skill Extractor |
| Project Type | Academic Final Year Project |
| Academic Session | 2026–2027 |
| Framework | Streamlit |
| Programming Language | Python |
| Developer | Hina Arshad |
| Version | 1.1 |

## Developer

**Designed and Developed by Hina Arshad**

GitHub: [@hinaarshad302](https://github.com/hinaarshad302)

---

*Making resume analysis and professional CV creation easier for English and Urdu users.*

**© 2026 Hina Arshad. All Rights Reserved.**# urdu-resume-skill-extractor
AI-powered resume extraction system for Urdu, English, and mixed-language CVs using Python and Streamlit.
