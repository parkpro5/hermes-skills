import os
import xlrd
import re
import sys
from collections import Counter

def run_preflight(base_dir):
    """
    KORUS 교육훈련 증빙 정합성 점검 (Preflight)
    base_dir: 'IN' 폴더의 경로 (엑셀 파일과 '증빙' 폴더가 포함된 경로)
    """
    # Use base_dir for all paths
    try:
        files_in_dir = os.listdir(base_dir)
    except Exception as e:
        print(f"Error accessing directory {base_dir}: {e}")
        return False

    excel_files = [f for f in files_in_dir if f.endswith('.xls')]
    if not excel_files:
        print("No .xls files found in the specified folder")
        return False
    
    excel_path = os.path.join(base_dir, excel_files[0])
    evidence_dir = os.path.join(base_dir, "증빙")
    
    if not os.path.exists(evidence_dir):
        print(f"Evidence directory not found at: {evidence_dir}")
        return False

    # 1. Read Excel
    wb = xlrd.open_workbook(excel_path)
    sheet = wb.sheet_by_index(0)
    
    # Get header
    header = [sheet.cell_value(0, col) for col in range(sheet.ncols)]
    
    # Find columns by name
    def get_col_idx(name):
        try:
            return header.index(name)
        except ValueError:
            return -1

    col_seq = get_col_idx("순번")
    col_file = get_col_idx("수료증빙파일")
    col_status = get_col_idx("교육훈련승인")
    
    if col_seq == -1 or col_file == -1 or col_status == -1:
        print(f"Required columns missing. Header: {header}")
        return False

    # 2. Process data rows (Exclude header and sum row)
    data_rows = []
    for row_idx in range(1, sheet.nrows):
        seq_val = sheet.cell_value(row_idx, col_seq)
        if seq_val == "계":
            continue
        data_rows.append(row_idx)
    
    total_data_count = len(data_rows)
    
    # 3. Evidence matching (M3, M4)
    evidence_rows = [] # Rows that have evidence files
    waiting_rows = []   # '대기중' rows
    waiting_no_evidence = 0
    
    for row_idx in data_rows:
        status = sheet.cell_value(row_idx, col_status)
        file_val = sheet.cell_value(row_idx, col_file)
        
        is_evidence_present = bool(str(file_val).strip())
        
        if status == "대기중":
            waiting_rows.append(row_idx)
            if not is_evidence_present:
                waiting_no_evidence += 1
        
        if is_evidence_present:
            evidence_rows.append(row_idx)

    # M3: Assign index k to evidence rows
    row_to_k = {row_idx: k for k, row_idx in enumerate(evidence_rows)}
    
    # 4. Scan evidence folder
    all_files = os.listdir(evidence_dir)
    
    # Extract k from filename "k_filename"
    file_k_map = {} # k -> [filenames]
    for f in all_files:
        match = re.match(r'^(\d+)_', f)
        if match:
            k = int(match.group(1))
            file_k_map.setdefault(k, []).append(f)
    
    # P1: Evidence rows count == unique k's count
    p1_success = len(evidence_rows) == len(file_k_map)
    
    # P2: k is continuous from 0 to len-1
    k_values = sorted(file_k_map.keys())
    p2_success = k_values == list(range(len(evidence_rows))) if k_values else True
    
    # P3: Total files == Sum(1 + N)
    total_expected_files = 0
    for row_idx in evidence_rows:
        file_val = sheet.cell_value(row_idx, col_file)
        n_match = re.search(r' 외 (\d+)건', str(file_val))
        n = int(n_match.group(1)) if n_match else 0
        total_expected_files += (1 + n)
    
    p3_success = len(all_files) == total_expected_files
    
    # P4: Each evidence row passes M4 (cross-check)
    m4_failures = []
    for row_idx in evidence_rows:
        k = row_to_k[row_idx]
        file_val = str(sheet.cell_value(row_idx, col_file))
        base_filename = re.sub(r' 외 \d+건$', '', file_val).strip()
        row_files = file_k_map.get(k, [])
        match_count = sum(1 for f in row_files if f.endswith(base_filename))
        n_match = re.search(r' 외 (\d+)건', file_val)
        n = int(n_match.group(1)) if n_match else 0
        count_ok = len(row_files) == (1 + n)
        if match_count != 1 or not count_ok:
            m4_failures.append((row_idx, f"match_count={match_count}, count_ok={count_ok}, filename={base_filename}"))

    p4_success = len(m4_failures) == 0
    
    print(f"--- Preflight Result ---")
    print(f"P1: Evidence rows ({len(evidence_rows)}) == Unique k ({len(file_k_map)}) : {'OK' if p1_success else 'FAIL'}")
    print(f"P2: Continuous k (0 to {len(evidence_rows)-1}) : {'OK' if p2_success else 'FAIL'}")
    print(f"P3: Total files ({len(all_files)}) == Expected ({total_expected_files}) : {'OK' if p3_success else 'FAIL'}")
    print(f"P4: M4 Cross-check (Failures: {len(m4_failures)}) : {'OK' if p4_success else 'FAIL'}")
    if m4_failures:
        print(f"   M4 Failures: {m4_failures[:5]}...")
    
    print(f"P5: Stats")
    print(f"   - Total data rows: {total_data_count}")
    print(f"   - Waiting rows: {len(waiting_rows)}")
    print(f"   - Waiting no evidence: {waiting_no_evidence}")
    
    if not all([p1_success, p2_success, p3_success, p4_success]):
        print("\n[!] PREFLIGHT FAILED. STOPPING.")
        return False
    else:
        print("\n[+] PREFLIGHT PASSED. READY FOR VERIFICATION.")
        return True

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python preflight.py <base_dir>")
    else:
        run_preflight(sys.argv[1])
