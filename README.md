# Enterprise Document & Quality Assurance Auditor

A Python-based digitalisation and workflow automation application designed to streamline corporate document reviews. This app replaces manual proofreading by scanning multi-format files (.docx, .pdf, .pptx) for grammatical errors, stylistic tone, and brand formatting consistency.

## Key Features
- **Automated QA Audit:** Scans for grammar, punctuation, and professional tone using cloud-based NLP engines.
- **Brand Consistency Checker:** Identifies font families and sizing hierarchies across reports to ensure corporate brand compliance.
- **Technical Terminology Whitelist:** Pre-configured to ignore industry-specific jargon, chemical terms, and operational codes during spelling audits.
- **Academic/Business Metrics:** Tracks total word counts, sentence structures, and readability scores (Flesch-Kincaid).

## Tech Stack
- **Frontend/Backend:** Python, Streamlit
- **NLP & Grammar:** language-tool-python, pyspellchecker
- **Document Parsing:** pypdf, python-docx, python-pptx

## Local Setup
1. Clone the repository:
   ```bash
   git clone [https://github.com/mursaleennaveedkhan1088-oss/enterprise-document-auditor.git](https://github.com/mursaleennaveedkhan1088-oss/enterprise-document-auditor.git)
   pip install -r requirements.txt
   streamlit run pro_auditor.py
