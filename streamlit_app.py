import streamlit as st
import pdfplumber
from PIL import Image
import pytesseract
import pandas as pd
import re
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

st.set_page_config(page_title="Document Clause Converter")

st.title("📄 PDF / Image to Structured Excel Converter")


# -----------------------------
# TEXT EXTRACTION
# -----------------------------

def extract_text_from_pdf(pdf_path: Path):
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text_content = page.extract_text() or ""
            pages_text.append(text_content.strip())
    return pages_text


def extract_text_from_image(image_path: Path):
    img = Image.open(image_path)
    img = img.convert("L")
    text_content = pytesseract.image_to_string(img, config="--psm 6")
    return [text_content.strip()]


# -----------------------------
# CLAUSE PROCESSING
# -----------------------------

def process_text_pages(text_pages):
    rows = []

    clause_pattern = re.compile(r"^\s*(\d+(?:\.\d+)*)\s+(.*)")

    current_clause = None
    current_text = ""

    for page_content in text_pages:
        lines = page_content.splitlines()

        for line in lines:
            line = line.strip()
            if not line:
                continue

            match = clause_pattern.match(line)

            if match:
                if current_clause:
                    rows.append({
                        "clause_number": current_clause,
                        "content": current_text.strip(),
                    })

                current_clause = match.group(1)
                current_text = match.group(2)

            else:
                if current_clause:
                    current_text += " " + line

    if current_clause:
        rows.append({
            "clause_number": current_clause,
            "content": current_text.strip(),
        })

    df = pd.DataFrame(rows)

    # ✅ Add description column
    df["description"] = ""

    return df


# -----------------------------
# STREAMLIT UI
# -----------------------------

uploaded_file = st.file_uploader(
    "Upload PDF or Image",
    type=["pdf", "png", "jpg", "jpeg"]
)

if uploaded_file:

    with st.spinner("Processing document..."):

        suffix = Path(uploaded_file.name).suffix

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_path = Path(tmp_file.name)

        ext = tmp_path.suffix.lower()

        if ext == ".pdf":
            text_pages = extract_text_from_pdf(tmp_path)
        elif ext in [".png", ".jpg", ".jpeg"]:
            text_pages = extract_text_from_image(tmp_path)
        else:
            st.error("Unsupported file type")
            st.stop()

        df = process_text_pages(text_pages)

    if df.empty:
        st.warning("No numbered clauses detected.")
    else:
        st.success(f"Extracted {len(df)} clauses.")
        st.dataframe(df)

        original_name = Path(uploaded_file.name).stem
        output_filename = f"converted_{original_name}.xlsx"

        # Save Excel temporarily
        output_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        df.to_excel(output_file.name, index=False)

        # -----------------------------
        # APPLY FORMATTING
        # -----------------------------
        wb = load_workbook(output_file.name)
        ws = wb.active

        # Freeze header
        ws.freeze_panes = "A2"

        # Bold header
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # Wrap text for all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        # Auto column width (limit 60)
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))

            ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

        wb.save(output_file.name)

        # -----------------------------
        # DOWNLOAD BUTTON
        # -----------------------------
        with open(output_file.name, "rb") as f:
            st.download_button(
                label="⬇ Download Excel File",
                data=f,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )