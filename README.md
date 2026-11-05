# Urdu Resume Skill Extractor

**Intelligent Resume Analysis & Candidate Information Extraction Platform**

A Streamlit-based application that extracts and organizes candidate information from PDF, DOCX, and TXT resumes. It supports English, Urdu, and bilingual CV content with professional Word, Excel, and PDF exports.

## Live Application

[Launch Urdu Resume Skill Extractor](https://urdu-resume-skill-extractor-3vjtecyvnp6rgfuqgf7t8w.streamlit.app/)

## GitHub Repository

[View Source Code](https://github.com/hinaarshad302/urdu-resume-skill-extractor)

---

## 1. Project Overview

Urdu Resume Skill Extractor is an academic project designed to simplify resume analysis and candidate information management.

The application extracts useful information from uploaded resumes and presents it in an editable candidate profile. Users can review extracted details, manually add Urdu content, and export professional CV documents.

## 2. Key Features

### Resume Upload

- PDF document support
- DOCX document support
- TXT document support
- Maximum upload size: 200 MB per file

### Candidate Information Extraction

- Candidate name detection
- Email address detection
- Phone number detection
- Basic skill detection
- Education information extraction
- Experience keyword extraction

### Candidate Profile Management

- Editable candidate information
- Professional summary
- Target job position
- Skills management
- Education details
- Experience details
- Language information

### Multilingual CV Support

- English CV
- Urdu CV
- Bilingual CV
- Manual Urdu content entry
- Right-to-Left (RTL) formatting support

### Professional Export System

| Export Format  | File Extension | Status    |
| -------------- | -------------- | --------- |
| Microsoft Word | `.docx`        | Supported |
| PDF Document   | `.pdf`         | Supported |
| Excel Workbook | `.xlsx`        | Supported |

---

## 3. Application Workflow

1. **Upload Resume:** Upload a PDF, DOCX, or TXT document.
2. **Extract Information:** Extract available candidate information.
3. **Analyze Candidate:** Review contact details, skills, education, and experience.
4. **Edit Profile:** Correct extracted information and update candidate details.
5. **Add Urdu Content:** Enter Urdu summary, skills, education, experience, and languages manually.
6. **Select CV Language:** Choose English, Urdu, or Bilingual.
7. **Export CV:** Download the candidate profile in Word, PDF, or Excel format.

---

## 4. Application Modules

| Module             | Description                                          |
| ------------------ | ---------------------------------------------------- |
| Resume Dashboard   | Main application dashboard and workflow              |
| Candidate Analysis | Resume information extraction and candidate overview |
| Profile Export     | Editable profile and professional CV exports         |

## 5. Supported File Formats

| Input Format | Purpose                          |
| ------------ | -------------------------------- |
| PDF          | Resume text extraction           |
| DOCX         | Microsoft Word resume extraction |
| TXT          | Plain text resume extraction     |

## 6. Technology Stack

| Technology      | Purpose                      |
| --------------- | ---------------------------- |
| Python          | Core application development |
| Streamlit       | Interactive web interface    |
| Pandas          | Data processing              |
| pdfplumber      | PDF text extraction          |
| pypdf           | PDF processing               |
| python-docx     | Word document generation     |
| ReportLab       | PDF document generation      |
| openpyxl        | Excel workbook generation    |
| arabic-reshaper | Urdu text shaping            |
| python-bidi     | Bidirectional text support   |
| Pillow          | Image processing support     |

---

## 7. Project Structure

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

*Additional project files and folders may be included in the repository.*

---

## 8. Installation Guide

### Prerequisites

- Python 3.10 or compatible version
- Git
- Internet connection
- Code editor (VS Code recommended)

### Step 1: Clone Repository

```bash
git clone https://github.com/hinaarshad302/urdu-resume-skill-extractor.git
```

### Step 2: Navigate to Project Folder

```bash
cd urdu-resume-skill-extractor
```

### Step 3: Create Virtual Environment

```bash
python -m venv venv
```

### Step 4: Activate Virtual Environment

**Windows:**

```bash
venv\Scripts\activate
```

**Linux / macOS:**

```bash
source venv/bin/activate
```

### Step 5: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 6: Run Application

```bash
streamlit run app.py
```

### Step 7: Open Application

Open the local URL displayed in the terminal, usually:

```text
http://localhost:8501
```

---

## 9. Requirements

The application uses the following Python packages:

```text
streamlit
pandas
pdfplumber
pypdf
python-docx
Pillow
openpyxl
reportlab
arabic-reshaper
python-bidi
```

---

## 10. How to Use the Application

1. Open the live application.
2. Upload a digital PDF, DOCX, or TXT resume.
3. Wait for resume processing.
4. Review the candidate overview.
5. Check extracted contact information.
6. Review detected skills.
7. Check education and experience information.
8. Edit the candidate profile where necessary.
9. Select the preferred CV language.
10. Enter Urdu information manually if preparing an Urdu or bilingual CV.
11. Choose the required export format.
12. Download the generated CV.

---

## 11. Urdu CV Content

The application provides manual Urdu content fields for:

- Professional Summary
- Skills
- Education
- Experience
- Languages

Urdu content is formatted with Right-to-Left (RTL) support where applicable.

**Note:** Automatic English-to-Urdu translation is not included.

---

## 12. Current Limitations

- Information extraction uses basic rule-based techniques.
- Skill detection depends on predefined keywords.
- Scanned PDF documents may require OCR for text extraction.
- Candidate information may need manual correction.
- Urdu CV content must be entered manually.
- Extraction accuracy depends on document structure and text quality.
- The application does not make automated hiring decisions.

---

## 13. Future Enhancements

- AI-powered resume analysis
- Advanced NLP-based skill extraction
- OCR support for scanned resumes
- Improved candidate name detection
- Job description and resume matching
- Candidate skill-gap analysis
- Resume scoring and recommendations
- Advanced multilingual support
- Improved Urdu document formatting

---

## 14. Academic Project Information

| Detail               | Information                 |
| -------------------- | --------------------------- |
| Project Name         | Urdu Resume Skill Extractor |
| Project Category     | Resume Intelligence         |
| Project Type         | Academic Final Year Project |
| Academic Session     | 2026–2027                   |
| Version              | Prototype 1.1               |
| Programming Language | Python                      |
| Framework            | Streamlit                   |
| Developer            | Hina Arshad                 |

---

## 15. Developer

**Hina Arshad**

GitHub: [hinaarshad302](https://github.com/hinaarshad302)

**Designed & Developed by Hina Arshad**

---

## 16. Disclaimer

This application provides automated assistance for resume analysis and candidate information extraction. Extracted information should be reviewed and verified before use in recruitment or other professional decisions.

---

**Urdu Resume Skill Extractor**

*Making resume analysis and professional CV creation easier for English and Urdu users.*

© 2026 Hina Arshad. All Rights Reserved.
