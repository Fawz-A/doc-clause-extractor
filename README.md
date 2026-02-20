# doc-clause-extractor

OCR-powered PDF and image clause extraction tool that converts structured document clauses (e.g. 6.1.2) into formatted Excel output.

---

## Features

- 📄 PDF text extraction (pdfplumber)
- 🖼 Image OCR support (Tesseract)
- 🔢 Clause parsing (e.g. 6.1, 6.1.2, 3.2.4)
- 📊 Structured Excel export
- 📁 Multi-file upload (Streamlit UI)
- 🧾 Wrapped text formatting
- 🧊 Frozen headers + clean column sizing

---

## Requirements

This project requires **Tesseract OCR** installed on your system.

### macOS
```bash
brew install tesseract
```

### Ubuntu / Debian
```bash
sudo apt install tesseract-ocr
```

---

## Setup (Recommended)

Create and activate virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Run Locally (Frontend - Streamlit App)

```bash
streamlit run streamlit_app.py
```

App will open at:
```
http://localhost:8501
```

---

## Run Locally (Backend Script Only)

```bash
python convert.py
```

This runs the standalone conversion script without the UI.

---

## Docker (Optional)

Build image:

```bash
docker build -t doc-clause-extractor .
```

Run container:

```bash
docker run -p 8501:8501 doc-clause-extractor
```

---

## Project Structure

```
doc-clause-extractor/
│
├── streamlit_app.py
├── convert.py
├── requirements.txt
├── Dockerfile
└── README.md
```

---

## Notes

- Maximum recommended upload size: 20MB per file (for demo).
- OCR accuracy depends on image quality.
- Designed as a demo deployment; production hardening can be added later.