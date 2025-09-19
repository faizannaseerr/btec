import os
import io
import zipfile
import streamlit as st
from typing import List

from script import convert_xlsx_to_csv, generate_documents_from_csv


def list_generated_docs(output_dir: str) -> List[str]:
    if not os.path.exists(output_dir):
        return []
    return [
        os.path.join(output_dir, name)
        for name in os.listdir(output_dir)
        if name.lower().endswith(".docx")
    ]


def zip_files(file_paths: List[str]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file_path in file_paths:
            arcname = os.path.basename(file_path)
            zipf.write(file_path, arcname=arcname)
    buffer.seek(0)
    return buffer.read()


SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SOURCE_DIR, "template.docx")
OUTPUT_DIR = os.path.join(SOURCE_DIR, "output")

os.makedirs(OUTPUT_DIR, exist_ok=True)

st.set_page_config(page_title="BTEC Doc Generator", page_icon="📄", layout="centered")
st.title("BTEC Assessment Document Generator")
st.write(
    "Upload your Excel (.xlsx) file. The app will convert it to CSV and generate Word documents using the local template.docx."
)

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"Template not found at {TEMPLATE_PATH}. Place template.docx in this folder.")

uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"], accept_multiple_files=False)

with st.sidebar:
    st.markdown("**Paths**")
    st.code(f"Template: {TEMPLATE_PATH}")
    st.code(f"Output:   {OUTPUT_DIR}")

generate_clicked = st.button("Generate Documents", type="primary", disabled=uploaded is None)

if generate_clicked and uploaded is not None:
    try:
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save uploaded XLSX to a temporary file
            temp_xlsx_path = os.path.join(
                tmpdir,
                uploaded.name if isinstance(getattr(uploaded, "name", None), str) and uploaded.name.lower().endswith(".xlsx") else "uploaded.xlsx",
            )
            with open(temp_xlsx_path, "wb") as f:
                f.write(uploaded.getbuffer())

            # Convert to a temporary CSV
            temp_csv_path = os.path.join(tmpdir, "uploaded.csv")
            convert_xlsx_to_csv(temp_xlsx_path, temp_csv_path)

            # Generate documents into OUTPUT_DIR
            generate_documents_from_csv(temp_csv_path, TEMPLATE_PATH, OUTPUT_DIR)

        generated = list_generated_docs(OUTPUT_DIR)
        if generated:
            st.success(f"Generated {len(generated)} document(s) in {OUTPUT_DIR}.")
            with st.expander("View generated files"):
                for path in generated:
                    st.write(os.path.basename(path))

            # Offer a zip download for convenience
            zip_bytes = zip_files(generated)
            st.download_button(
                label="Download all as ZIP",
                data=zip_bytes,
                file_name="generated_docs.zip",
                mime="application/zip",
            )
        else:
            st.warning("No documents were generated. Check your data headers and placeholders.")
    except Exception as e:
        st.error(f"Error: {e}")

