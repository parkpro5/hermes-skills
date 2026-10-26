import json
import os
import re
import sys
import pandas as pd
from pathlib import Path

def parse_time(h, m):
    try:
        return int(float(h)) * 60 + int(float(m))
    except:
        return None

def extract_time_from_text(text):
    match = re.search(r'(\d+(?:\.\d+)?)\s*시간\s*(\d+)\s*분', text)
    if match:
        return int(float(match.group(1)) * 60) + int(match.group(2))
    match = re.search(r'(\d+(?:\.\d+)?)\s*시간', text)
    if match:
        return int(float(match.group(1)) * 60)
    match = re.search(r'\(?(\d+)\s*분\)?', text)
    if match:
        return int(match.group(1))
    return None

def normalize(text):
    if text is None: return ""
    return "".join(str(text).split())

def verify_row(row):
    seq = row['seq']
    name = row['name']
    course = row['course']
    real_min = parse_time(row['hr_real'], row['hr_real_min'])
    rec_min = parse_time(row['hr_rec'], row['hr_rec_min'])
    target_min = max(real_min, rec_min) if real_min is not None and rec_min is not None else (real_min or rec_min)

    if not row['evidence_files']:
        return {
            "seq": seq, "name": name, "course": course,
            "decision": "증빙없음(제외)", "confidence": "Low",
            "reason": "수료증빙파일 항목이 비어 있음", "evidence_files": [],
            "excel_val": row['excel_file_val'], "evid_time": None, "target_time": target_min
        }

    merged_text = " ".join([t['text'] for t in row['extracted_texts']])
    norm_merged_text = normalize(merged_text)
    name_ok = normalize(name) in norm_merged_text
    org_ok = normalize("서울과학기술대학교") in norm_merged_text
    course_keywords = re.sub(r'\(.*?\) |제\s*\d+\s*기', '', course)
    norm_course_kw = normalize(course_keywords)
    course_ok = norm_course_kw in norm_merged_text if norm_course_kw else False
    if not course_ok:
        words = [w for w in course.split() if len(w) > 1]
        course_ok = any(normalize(w) in norm_merged_text for w in words)

    evid_min = extract_time_from_text(merged_text)
    if evid_min is None:
        time_ok = None
        time_reason = "증빙에서 시간 정보 확인 불가"
    elif target_min is None:
        time_ok = True
        time_reason = "엑셀 시간 확인 불가(증빙 시간 존재)"
    elif evid_min >= target_min:
        time_ok = True
        time_reason = ""
    else:
        time_ok = False
        time_reason = f"증빙시간({evid_min}분) < 엑셀시간({target_min}분)"

    format_ok = any(normalize(kw) in norm_merged_text for kw in ["수료증", "이수증", "참석확인증"])
    issuer_titles = ["총장", "원장", "이사장", "회장", "처장", "센터장"]
    confirmation_phrases = ["확인합니다", "증명합니다", "수여합니다", "인증"]
    has_title = any(normalize(title) in norm_merged_text for title in issuer_titles)
    has_confirm = any(normalize(phrase) in norm_merged_text for phrase in confirmation_phrases)
    issuer_ok = has_title or has_confirm

    reasons = []
    if not name_ok: reasons.append("성명 불일치")
    if not org_ok: reasons.append("기관 불일치")
    if not course_ok: reasons.append("과정명 불일치")
    if time_ok == False: reasons.append(time_reason)
    if not format_ok: reasons.append("부적절한 문서 형식")
    if not issuer_ok: reasons.append("발급 주체 확인 불가")
    
    if time_ok is None:
        decision, confidence, reason = "수동 확인", "Low", "시간 판단 불가"
    elif not reasons:
        decision, confidence, reason = "승인 추천", "High", "모든 항목 일치"
    else:
        decision, confidence, reason = "검토 필요", "Low", ", ".join(reasons)
        
    return {
        "seq": seq, "name": name, "course": course,
        "decision": decision, "confidence": confidence, "reason": reason,
        "evidence_files": row['evidence_files'], "excel_val": row['excel_file_val'],
        "evid_time": evid_min, "target_time": target_min
    }

def main(in_data_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    try:
        with open(in_data_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"Error reading {in_data_path}: {e}")
        return
        
    results = [verify_row(row) for row in data]
    
    df = pd.DataFrame(results)
    cols = ["seq", "name", "decision", "confidence", "reason", "evidence_files"]
    df_final = df[cols].copy()
    df_final.columns = ["순번", "성명", "판정", "확신도", "사유", "증빙파일"]
    df_final["체크대상(Y/N)"] = df_final.apply(lambda r: "Y" if r["판정"] == "승인 추천" and r["확신도"] in ["High", "Medium"] else "N", axis=1)
    df_final.to_excel(os.path.join(out_dir, "final_results.xlsx"), index=False)
    
    checkbox_list = sorted([str(r["seq"]) for r in results if r["decision"] == "승인 추천" and r["confidence"] in ["High", "Medium"]])
    with open(os.path.join(out_dir, "checkbox_list.txt"), "w", encoding="utf-8") as f:
        f.write(", ".join(checkbox_list))
        
    html = "<html><head><style>body{font-family:sans-serif;} table{border-collapse:collapse;width:100%;} td,th{border:1px solid #ddd;padding:8px;} th{background-color:#f2f2f2;}</style></head><body>"
    html += "<h1 style='text-align:center;'>KORUS 교육훈련 실적 검증 리포트</h1>"
    html += f"<p><b>요약:</b> 대기중 {len(results)}건 = 승인추천 {len([r for r in results if r['decision'] == '승인 추천'])} + 검토필요 {len([r for r in results if r['decision'] == '검토 필요'])} + 증빙없음 {len([r for r in results if r['decision'] == '증빙없음(제외)'])} + 수동확인 {len([r for r in results if r['decision'] == '수동 확인'])}</p>"
    html += "<table><tr><th>순번</th><th>성명</th><th>과정명</th><th>판정</th><th>확신도</th><th>사유</th><th>증빙파일</th></tr>"
    for r in results:
        color = "lightgreen" if r["decision"] == "승인 추천" else "white"
        if r["decision"] == "수동 확인": color = "lightyellow"
        html += f"<tr style='background-color:{color}'><td align='center'>{r['seq']}</td><td>{r['name']}</td><td>{r['course']}</td><td align='center'>{r['decision']}</td><td align='center'>{r['confidence']}</td><td>{r['reason']}</td><td>{r['evidence_files']}</td></tr>"
    html += "</table></body></html>"
    
    with open(os.path.join(out_dir, "final_report.html"), "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Verification complete. Results saved to {out_dir}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python verify_final.py <in_data_path> <out_dir>")
    else:
        main(sys.argv[1], sys.argv[2])
