import os
import io
import zipfile
import multiprocessing
import streamlit as st
from typing import List, Tuple

from script import convert_xlsx_to_csv, generate_documents_from_csv


def create_zip_from_docs(docs: List[Tuple[str, bytes]]) -> bytes:
    """Create a ZIP file in memory from a list of (filename, document_bytes) tuples."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for filename, doc_bytes in docs:
            zipf.writestr(filename, doc_bytes)
    buffer.seek(0)
    return buffer.read()


SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SOURCE_DIR, "template.docx")

st.set_page_config(page_title="BTEC Doc Generator", page_icon="📄", layout="centered")
st.title("BTEC Assessment Document Generator")
st.write(
    "Upload your Excel (.xlsx) file. The app will convert it to CSV and generate Word documents using the local template.docx."
)

with st.expander("What will happen?", expanded=False):
    st.markdown(
        "- Convert the uploaded Excel to CSV\n"
        "- Read each row as one learner/record\n"
        "- Fill placeholders in `template.docx` using the row values in parallel\n"
        "- Save one `.docx` per row into the output folder"
    )

# Add CPU cores selector
cpu_count = multiprocessing.cpu_count()
max_workers = st.slider(
    "Number of CPU cores to use",
    min_value=1,
    max_value=max(1, cpu_count),
    value=max(1, cpu_count - 1),
    help="More cores = faster processing, but leave at least 1 core free for system tasks"
)


if not os.path.exists(TEMPLATE_PATH):
    st.error(f"Template not found at {TEMPLATE_PATH}. Place template.docx in this folder.")

uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"], accept_multiple_files=False)

with st.sidebar:
    st.markdown("**Paths**")
    st.code(f"Template: {TEMPLATE_PATH}")
    
    st.markdown("---")
    st.markdown("**Downloads 📥**")
    
    # Add download button for the Excel template
    excel_path = os.path.join(SOURCE_DIR, "btec_data_template.xlsx")
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as file:
            st.download_button(
                label="Download Marksheet Template",
                data=file,
                file_name="btec_data_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Excel template file not found")

generate_clicked = st.button("Generate Documents", type="primary", disabled=uploaded is None)

if generate_clicked and uploaded is not None:
    try:
        import tempfile
        status_placeholder = st.empty()
        progress_bar = st.progress(0)
        log_placeholder = st.empty()
        logs = []

        def append_log(line: str) -> None:
            logs.append(line)
            # Render as a growing log
            log_placeholder.code("\n".join(logs))

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
            status_placeholder.write("Converting Excel to CSV…")
            convert_xlsx_to_csv(temp_xlsx_path, temp_csv_path)
            status_placeholder.write("CSV ready. Starting document generation…")

            # Generate documents in memory
            total_rows_box = {"value": 0}
            done_count = {"value": 0}

            def on_progress(event: str, payload: dict) -> None:
                if event == "start":
                    total = int(payload.get("total_rows", 0) or 0)
                    total_rows_box["value"] = total
                    progress_bar.progress(0)
                    status_placeholder.write(f"Processing {total} row(s)…")
                elif event == "row_start":
                    idx = int(payload.get("index", 0)) + 1
                    row = payload.get("row", {}) or {}
                    learner = (row.get("Learner Name") or "").strip() or "(no name)"
                    append_log(f"Row {idx}: generating for {learner}…")
                elif event == "row_done":
                    done_count["value"] += 1
                    filename = payload.get("filename", "")
                    append_log(f"✅ Generated: {filename}")
                    total = total_rows_box["value"] or 0
                    if total > 0:
                        percent = int(min(100, round((done_count["value"] / total) * 100)))
                        progress_bar.progress(percent)
                elif event == "row_error":
                    done_count["value"] += 1
                    err = payload.get("error", "Unknown error")
                    idx = int(payload.get("index", 0)) + 1
                    append_log(f"❌ Row {idx} error: {err}")
                    total = total_rows_box["value"] or 0
                    if total > 0:
                        percent = int(min(100, round((done_count["value"] / total) * 100)))
                        progress_bar.progress(percent)
                elif event == "complete":
                    gen = int(payload.get("generated", 0) or 0)
                    total = int(payload.get("total_rows", 0) or 0)
                    status_placeholder.write(f"Completed. Generated {gen} of {total} row(s).")

            generated_docs = generate_documents_from_csv(temp_csv_path, TEMPLATE_PATH, progress=on_progress, max_workers=max_workers)

        if generated_docs:
            st.success(f"Generated {len(generated_docs)} document(s).")
            doc_filenames = [doc[0] for doc in generated_docs]
            with st.expander("View generated files"):
                for doc_name in doc_filenames:
                    st.write(doc_name)

            zip_bytes = create_zip_from_docs(generated_docs)
            st.download_button(
                label="Download all as ZIP",
                data=zip_bytes,
                file_name="generated_docs.zip",
                mime="application/zip",
            )

            # Minimize progress UI after ZIP is generated; provide collapsible details
            try:
                progress_bar.empty()
                status_placeholder.empty()
                log_placeholder.empty()
            except Exception:
                pass

            with st.expander("View progress details", expanded=False):
                if logs:
                    st.code("\n".join(logs))
                else:
                    st.write("No progress logs available.")
        else:
            st.warning("No documents were generated. Check your data headers and placeholders.")
    except Exception as e:
        st.error(f"Error: {e}")

