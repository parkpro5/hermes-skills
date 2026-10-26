import os
import xlrd
import json
import re
import zipfile
import xml.etree.ElementTree as ET
import sys
from pathlib import Path
import fitz  # PyMuPDF
import openpyxl

def extract_pdf_text(path):
    try:
        doc = fitz.open(path)
        text = ""
        for page in doc:
            text += page.get_text()
        return text.strip()
    except Exception as e:
        return f"[PDF Error: {e}]"

def extract_xlsx_text(path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        text = ""
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                row_text = " ".join([str(cell) for cell in row if cell is not None])
                text += row_text + "\\n"
        return text.strip()
    except Exception as e:
        return f"[XLSX Error: {e}]"

def extract_hwpx_text(path):
    try:
        with zipfile.ZipFile(path, 'r') as z:
            content_files = [f for f in z.namelist() if f.startswith('Contents/section')]
            text = ""
            for cf in content_files:
                with z.open(cf) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    for node in root.iter():
                        if node.text:
                            text += node.text + " "
            return text.strip()
    except Exception as e:
        return f"[HWPX Error: {e}]"

def run_prepare(base_dir, out_dir):
    """
    KORUS 교육훈련 증빙 텍스트 추출 (Prepare)
    base_dir: 'IN' 폴더 경로
    out_dir: 결과 저장 폴더 경로
    """
    os.makedirs(out_dir, exist_ok=True)
    try:
        files_in_dir = os.listdir(base_dir)
    except Exception as e:
        print(f"Error accessing directory {base_dir}: {e}")
        return

    excel_files = [f for f in files_in_dir if f.endswith('.xls')]
    if not excel_files:
        print("Excel file not found")
        return
    
    excel_path = os.path.join(base_dir, excel_files[0])
    evidence_dir = os.path.join(base_dir, "증빙")
    
    wb = xlrd.open_workbook(excel_path)
    sheet = wb.sheet_by_index(0)
    header = [sheet.cell_value(0, col) for col in range(sheet.ncols)]
    
    # Column indices
    col_seq = header.index("순번")
    col_file = header.index("수료증빙파일")
    col_status = header.index("교육훈련승인")
    col_name = header.index("성명")
    col_course = header.index("과정명")
    col_hr_real = header.index("실적시간")
    col_hr_rec = header.index("인정시간")
    
    prepared_data = []
    evidence_rows = []
    for r in range(1, sheet.nrows):
        if sheet.cell_value(r, col_seq) == "계": continue
        if str(sheet.cell_value(r, col_file)).strip():
            evidence_rows.append(r)
            
    row_to_k = {row_idx: k for k, row_idx in enumerate(evidence_rows)}
    
    for r in range(1, sheet.nrows):
        if sheet.cell_value(r, col_seq) == "계": continue
        status = sheet.cell_value(r, col_status)
        if status != "대기중": continue
        
        file_val = str(sheet.cell_value(r, col_file)).strip()
        row_info = {
            "seq": sheet.cell_value(r, col_seq),
            "name": sheet.cell_value(r, col_name),
            "course": sheet.cell_value(r, col_course),
            "hr_real": sheet.cell_value(r, col_hr_real),
            "hr_real_min": sheet.cell_value(r, col_hr_real + 1),
            "hr_rec": sheet.cell_value(r, col_hr_rec),
            "hr_rec_min": sheet.cell_value(r, col_hr_rec + 1),
            "excel_file_val": file_val,
            "evidence_files": [],
            "extracted_texts": []
        }
        
        if file_val:
            k = row_to_k.get(r)
            if k is not None:
                folder_files = os.listdir(evidence_dir)
                matched_files = [f for f in folder_files if f.startswith(f"{k}_")]
                row_info["evidence_files"] = matched_files
                for f in matched_files:
                    f_path = os.path.join(evidence_dir, f)
                    ext = Path(f).suffix.lower()
                    if ext == '.pdf': text = extract_pdf_text(f_path)
                    elif ext == '.xlsx': text = extract_xlsx_text(f_path)
                    elif ext == '.hwpx': text = extract_hwpx_text(f_path)
                    else: text = f"[Unsupported format: {ext}]"
                    row_info["extracted_texts"].append({"file": f, "text": text})
        
        prepared_data.append(row_info)
        
    with open(os.path.join(out_dir, "evidence_data.json"), "w", encoding="utf-8") as f:
        json.dump(prepared_data, f, ensure_ascii=False, indent=2)
    
    print(f"Successfully prepared {len(prepared_data)} rows. Saved to {os.path.join(out_dir, 'evidence_data.json')}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python verify_prepare.py <base_dir> <out_dir>")
    else:
        run_prepare(sys.argv[1], sys.argv[2])
