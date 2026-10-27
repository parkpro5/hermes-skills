import json
import os
import re
import sys
import pandas as pd
from pathlib import Path

def parse_time_to_min(h, m):
    try:
        total = 0
        if h is not None:
            total += float(h) * 60
        if m is not None:
            total += float(m)
        return total
    except:
        return None

def normalize_text(text):
    if not text: return ""
    text = re.sub(r'\(\d+기\)', '', text)
    text = re.sub(r'제\d+기', '', text)
    text = re.sub(r'\d+기', '', text)
    text = re.sub(r'[^a-zA-Z0-9가-힣]', '', text)
    text = text.replace('제', '').replace('기', '')
    return text

def verify_entry(excel_row, evidence_text):
    if not evidence_text:
        return "증빙 없음", "Low", "증빙 텍스트 추출 실패"

    reasons = []
    
    # --- 1. Name Verification ---
    name_excel = str(excel_row.get('성명', '')).strip().replace(' ', '')
    name_evidence = re.sub(r'\s+', '', evidence_text)
    name_match = name_excel and (name_excel in name_evidence)
    if not name_match:
        reasons.append(f"성명 불일치(엑셀:{name_excel}/증빙:{name_evidence[:10]}...)")

    # --- 2. Course Verification ---
    raw_course_excel = str(excel_row.get('과정명', '')).strip()
    course_excel_norm = normalize_text(raw_course_excel)
    evidence_norm = normalize_text(evidence_text)
    
    course_match = False
    if course_excel_norm and evidence_norm:
        if course_excel_norm in evidence_norm or evidence_norm in course_excel_norm:
            course_match = True
        else:
            words = [w for w in re.split(r'[^a-zA-Z0-9가-힣]', raw_course_excel) if len(w) > 1]
            if words:
                matches = [w for w in words if normalize_text(w) in evidence_norm]
                if len(matches) / len(words) >= 0.6:
                    course_match = True
    
    if not course_match:
        snippet = evidence_text[:60].replace('\n', ' ')
        reasons.append(f"과정명 불일치(엑셀:{raw_course_excel}/증빙:{snippet}...)")

    # --- 3. Time Verification ---
    time_match = False
    try:
        h_excel = excel_row.get('교육시간(시간)', 0)
        m_excel = excel_row.get('교육시간(분)', 0)
        target_min = parse_time_to_min(h_excel, m_excel)
        
        time_matches = re.findall(r'(\d+(?:\.\d+)?)\s*시간\s*(?:(\d+(?:\.\d+)?)\s*분)?', evidence_text)
        if time_matches:
            m_list = []
            for h, m in time_matches:
                m_val = float(h) * 60 + (float(m) if m else 0)
                m_list.append(m_val)
            
            sum_evid_min = sum(m_list)
            max_evid_min = max(m_list)
            if abs(max_evid_min - target_min) <= 1 or abs(sum_evid_min - target_min) <= 1 or max_evid_min >= target_min - 1:
                time_match = True
            else:
                closest = max_evid_min if abs(max_evid_min - target_min) < abs(sum_evid_min - target_min) else sum_evid_min
                reasons.append(f"시간 부족(엑셀:{int(target_min)}분/증빙:{int(closest)}분)")
        else:
            reasons.append(f"시간 확인 불가(엑셀:{int(target_min)}분)")
    except Exception as e:
        reasons.append(f"시간 계산 오류: {str(e)}")

    # --- 4. Org and Form Verification ---
    trusted_orgs = ['인프런', 'Inflearn', 'Udemy', 'Coursera', '국가공무원인재개발원', '대학교', '총장', '원장', 'CEO', 'Representative', '재단']
    found_orgs = [org for org in trusted_orgs if org.lower() in evidence_text.lower()]
    is_trusted = len(found_orgs) > 0
    
    cert_keywords = ['CERTIFICATE', 'COMPLETION', 'AWARD', '수료', '이수', '증명', '확인증']
    is_cert_form = False
    found_kw = []
    evidence_upper = evidence_text.upper()
    for kw in cert_keywords:
        if kw in evidence_upper:
            is_cert_form = True
            found_kw.append(kw)
    
    org_match = is_trusted and is_cert_form
    
    if not org_match:
        org_detail = []
        if not is_trusted: org_detail.append("신뢰기관명 미검출")
        else: org_detail.append(f"기관확인됨({', '.join(found_orgs)})")
        if not is_cert_form: org_detail.append("표준서식키워드 미검출")
        else: org_detail.append(f"서식확인됨({', '.join(found_kw)})")
        org_reason = f"발급주체/형식 확인 필요({', '.join(org_detail)})"
    else:
        org_reason = ""

    # --- FINAL DECISION ---
    if name_match and course_match and time_match:
        final_decision = "승인 추천"
        if is_trusted:
            confidence = "High"
            final_reason = "모든 핵심 항목 일치"
        elif is_cert_form:
            confidence = "Medium"
            final_reason = f"핵심 항목 일치(단, {org_reason})"
        else:
            confidence = "Medium"
            final_reason = f"핵심 항목 일치(단, {org_reason})"
    else:
        critical_failures = [r for r in reasons if "발급주체" not in r]
        if len(critical_failures) == 1:
            final_decision = "검토 필요"
            confidence = "Medium"
        else:
            final_decision = "수동 확인"
            confidence = "Low"
        final_reasons = []
        final_reasons.extend(critical_failures)
        if not org_match:
            final_reasons.append(org_reason)
        final_reason = ". ".join(final_reasons) if final_reasons else "알 수 없는 불일치"

    return final_decision, confidence, final_reason

def main():
    if len(sys.argv) < 3:
        print("Usage: python verify_final.py <json_path> <out_dir>")
        return

    json_path = sys.argv[1]
    out_dir = sys.argv[2]
    os.makedirs(out_dir, exist_ok=True)

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"Error loading JSON: {e}")
        return
    
    results = []
    for entry in data:
        seq = entry['seq']
        name = entry['name']
        course = entry['course']
        excel_row = {
            '성명': name,
            '과정명': course,
            '교육시간(시간)': entry.get('hr_real'),
            '교육시간(분)': entry.get('hr_real_min')
        }
        combined_text = " ".join([ext['text'] for ext in entry.get('extracted_texts', [])])
        decision, confidence, reason = verify_entry(excel_row, combined_text)
        results.append({
            '순번': seq, '성명': name, '과정명': course,
            '판정': decision, '확신도': confidence, '사유': reason,
            '증빙파일': entry.get('evidence_files', [])
        })

    df = pd.DataFrame(results)
    df.to_excel(os.path.join(out_dir, 'final_results.xlsx'), index=False)

    summary = df['판정'].value_counts().to_dict()
    total = len(df)
    html_content = f"<html><head><style>body{{font-family:sans-serif; padding:20px;}} table{{border-collapse:collapse; width:100%;}} td,th{{border:1px solid #ddd; padding:8px;}} th{{background-color:#f2f2f2;}} .approve{{background-color:lightgreen;}} .review{{background-color:white;}} .manual{{background-color:lightcoral;}}</style></head><body><h1 style='text-align:center;'>KORUS 교육훈련 실적 검증 리포트 (v19 - 신뢰기관 우선 반영 버전)</h1><p><b>요약:</b> 총 {total}건 = 승인추천 {summary.get('승인 추천', 0)} + 검토필요 {summary.get('검토 필요', 0)} + 수동확인 {summary.get('수동 확인', 0)} + 증빙없음 {summary.get('증빙 없음', 0)}</p><table><tr style='background-color:#f2f2f2;'><th>순번</th><th>성명</th><th>과정명</th><th>판정</th><th>확신도</th><th>사유</th><th>증빙파일</th></tr>"
    for _, row in df.iterrows():
        cls = "approve" if row['판정'] == "승인 추천" else "review" if row['판정'] == "검토 필요" else "manual"
        html_content += f"<tr class='{cls}'><td align='center'>{row['순번']}</td><td>{row['성명']}</td><td>{row['과정명']}</td><td align='center'>{row['판정']}</td><td align='center'>{row['확신도']}</td><td>{row['사유']}</td><td>{row['증빙파일']}</td></tr>"
    html_content += "</table></body></html>"
    
    with open(os.path.join(out_dir, 'final_report.html'), 'w', encoding='utf-8') as f:
        f.write(html_content)
    with open(os.path.join(out_dir, 'checkbox_list.txt'), 'w', encoding='utf-8') as f:
        approved_seqs = df[df['판정'] == '승인 추천']['순번'].tolist()
        f.write("\n".join(approved_seqs))
    print("Verification complete. Results saved to " + out_dir)

if __name__ == "__main__":
    main()
