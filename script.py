from docx import Document
import csv
import os
from typing import Dict, List, Set, Optional, Callable
import tempfile
from openpyxl import load_workbook
from datetime import date, datetime

def replace_text_in_paragraph(paragraph, placeholder: str, replacement: str) -> None:
    full_text = ''.join(run.text for run in paragraph.runs)
    if placeholder in full_text:
        new_text = full_text.replace(placeholder, replacement)
        # Clear existing runs
        for run in paragraph.runs:
            run.text = ''
        # Set the new text as a single run (handles case where there are no runs)
        if paragraph.runs:
            paragraph.runs[0].text = new_text
        else:
            paragraph.add_run(new_text)

def replace_placeholders(doc: Document, placeholder: str, replacement: str) -> None:
    # Replace in paragraphs outside tables
    # for paragraph in doc.paragraphs:
    #     replace_text_in_paragraph(paragraph, placeholder, replacement)

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_paragraph(paragraph, placeholder, replacement)

 


def generate_placeholder_variants(placeholder: str) -> Set[str]:
    """Generate reasonable variants for a placeholder to improve matching robustness."""
    variants: Set[str] = set()
    raw = placeholder
    variants.add(raw)

    # Ensure closing bracket if missing
    if raw.startswith('[') and not raw.endswith(']'):
        variants.add(raw + ']')

    # Normalize en dash and hyphen both ways
    if '–' in raw:
        variants.add(raw.replace('–', '-'))
    if '-' in raw:
        variants.add(raw.replace('-', '–'))

    return variants


DECLARED_PLACEHOLDERS: List[str] = [
    "[Programme Title]",
    "[Learner Registration Number]",
    "[Learner Name]",
    "[Assignment Title]",
    "[Assessor Name]",
    "[Unit/Component Number and Title]",
    "[Targeted Learning Aims/Assessment Criteria (Initial)]",
    "[First Submission - Deadline]",
    "[First Submission - Date Submitted]",
    "[Extension Approved (Y/N)]",
    "[Initial - General Comments]",
    "[Initial - Learner Signature (Name or File Path)]",
    "[Initial - Learner Declaration Date]",
    "[Initial - Assessor Signature (Name or File Path)]",
    "[Initial - Assessor Declaration Date]",
    "[Initial - Date of Feedback to Learner]",
    "[Resubmission - Authorised by Lead Internal Verifier (Name)]",
    "[Resubmission - Authorisation Date]",
    "[Resubmission - Deadline]",
    "[Resubmission - Date Submitted]",
    "[Resubmission - General Comments]",
    "[Resubmission - Learner Signature (Name or File Path)]",
    "[Resubmission - Learner Declaration Date]",
    "[Resubmission - Assessor Signature (Name or File Path)]",
    "[Resubmission - Assessor Declaration Date]",
    "[Resubmission - Date of Feedback to Learner]",
    "[Retake - Deadline]",
    "[Retake - Date Submitted]",
    "[Retake - General Comments]",
    "[Retake - Learner Signature (Name or File Path)]",
    "[Retake - Learner Declaration Date]",
    "[Retake - Assessor Signature (Name or File Path)]",
    "[Retake - Assessor Declaration Date]",
    "[Retake - Date of Feedback to Learner]",
]


def process_criteria(targeted: str, achieved: str, max_criteria: int = 3) -> tuple[list[str], list[str]]:
    """Process targeted and achieved criteria into lists and Y/N markers.
    
    Args:
        targeted: Comma-separated string of targeted criteria
        achieved: Comma-separated string of achieved criteria
        max_criteria: Maximum number of criteria to process (default 3)
    """
    # Split and clean targeted criteria
    targeted_list = [c.strip() for c in targeted.split(',') if c.strip()]
    # Split and clean achieved criteria
    achieved_list = [c.strip() for c in achieved.split(',') if c.strip()]
    
    # Generate Y/N list based on whether each targeted criteria was achieved
    achieved_yn = ['Y' if t in achieved_list else 'N' for t in targeted_list]
    
    # Pad both lists to specified length with empty strings
    targeted_list.extend([''] * (max_criteria - len(targeted_list)))
    achieved_yn.extend([''] * (max_criteria - len(achieved_yn)))
    
    return targeted_list[:max_criteria], achieved_yn[:max_criteria]

def replace_all_placeholders(doc: Document, row: Dict[str, str]) -> None:
    """Replace placeholders in doc using both declared list and dynamic [Header] placeholders."""
    # Build mapping: placeholder variant -> replacement value
    replacement_map: Dict[str, str] = {}

    # Handle Initial criteria placeholders (up to 3)
    initial_targeted = row.get('Initial - Targeted Criteria', '').strip()
    initial_achieved = row.get('Initial - Criteria Achieved', '').strip()
    initial_targeted_list, initial_achieved_yn = process_criteria(initial_targeted, initial_achieved, max_criteria=3)
    
    # Add Initial criteria mappings
    for i, (target, achieved) in enumerate(zip(initial_targeted_list, initial_achieved_yn), 1):
        replacement_map[f'[ITC{i}]'] = target
        replacement_map[f'[ICA{i}]'] = achieved

    # Handle Resubmission criteria placeholders (up to 5)
    resub_targeted = row.get('Resubmission - Targeted Criteria', '').strip()
    resub_achieved = row.get('Resubmission - Criteria Achieved', '').strip()
    resub_targeted_list, resub_achieved_yn = process_criteria(resub_targeted, resub_achieved, max_criteria=5)
    
    # Add Resubmission criteria mappings
    for i, (target, achieved) in enumerate(zip(resub_targeted_list, resub_achieved_yn), 1):
        replacement_map[f'[RTC{i}]'] = target
        replacement_map[f'[RCA{i}]'] = achieved

    # Declared placeholders from specification
    for placeholder in DECLARED_PLACEHOLDERS:
        column_name = placeholder.strip().lstrip('[').rstrip(']')
        value = (row.get(column_name) or '').strip()
        for variant in generate_placeholder_variants(placeholder):
            replacement_map[variant] = value

    # Dynamic placeholders for each CSV column (covers additional fields if any)
    # for header in csv_headers:
    #     placeholder = f"[{header}]"
    #     value = (row.get(header) or '').strip()
    #     for variant in generate_placeholder_variants(placeholder):
    #         replacement_map[variant] = value

    # Perform replacements
    for placeholder, replacement in replacement_map.items():
        replace_placeholders(doc, placeholder, replacement)


def generate_documents_from_csv(
    csv_path: str,
    template_path: str,
    output_dir: str,
    progress: Optional[Callable[[str, Dict[str, object]], None]] = None,
) -> None:
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    os.makedirs(output_dir, exist_ok=True)

    # Pre-count total data rows for progress reporting
    try:
        with open(csv_path, mode='r', encoding='utf-8-sig', newline='') as f_count:
            total_rows = max(0, sum(1 for _ in csv.reader(f_count)) - 1)
    except Exception:
        total_rows = 0

    if progress:
        progress('start', {'total_rows': total_rows})

    generated_count = 0

    with open(csv_path, mode='r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)

        for index, row in enumerate(reader):
            try:
                if progress:
                    progress('row_start', {'index': index, 'row': row})

                # Create a fresh document from the template for each row
                print(f"Creating document for row {index}")
                doc = Document(template_path)

                replace_all_placeholders(doc, row)

                name = (row.get('Learner Name') or '').strip()
                reg = (row.get('Learner Registration Number') or '').strip()
                base_name = f"{name} {reg}".strip()
                if f"{base_name}.docx" in os.listdir(output_dir):
                    base_name = f"{base_name}_{index + 1}"
                if not base_name:
                    base_name = f"output_{index + 1}"
                out_path = os.path.join(output_dir, f"{base_name}.docx")

                doc.save(out_path)
                print(f"Saved: {out_path}")

                generated_count += 1
                if progress:
                    progress('row_done', {'index': index, 'out_path': out_path})
            except Exception as e:
                if progress:
                    progress('row_error', {'index': index, 'error': str(e)})
                # Continue with next row
                continue

    if progress:
        progress('complete', {'generated': generated_count, 'total_rows': total_rows})

def convert_xlsx_to_csv(xlsx_path: str, csv_path: str) -> None:
    """Convert the first worksheet of an .xlsx file to a UTF-8 CSV file."""
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"XLSX not found: {xlsx_path}")

    wb = load_workbook(filename=xlsx_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    with open(csv_path, mode='w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        for row in ws.iter_rows(values_only=True):
            def format_cell_for_csv(value):
                if isinstance(value, datetime):
                    return value.date().isoformat()
                if isinstance(value, date):
                    return value.isoformat()
                return "" if value is None else value

            writer.writerow([format_cell_for_csv(v) for v in row])

    wb.close()


if __name__ == "__main__":
    source_dir = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_PATH = os.path.join(source_dir, "template.docx")
    XLSX_PATH = os.path.join(source_dir, "dummy_data.xlsx")
    OUTPUT_DIR = os.path.join(source_dir, "output")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Convert XLSX to a temp CSV file
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_csv_path = os.path.join(tmpdir, "dummy.csv")
        convert_xlsx_to_csv(XLSX_PATH, temp_csv_path)
        generate_documents_from_csv(temp_csv_path, TEMPLATE_PATH, OUTPUT_DIR)

