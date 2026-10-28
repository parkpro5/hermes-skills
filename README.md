# 🎓 KORUS 교육훈련 실적 검증 스킬 (korus-verification)

KORUS 시스템에서 내려받은 교육훈련 실적 엑셀 파일과 증빙 서류(PDF 등)를 자동으로 대조하여, 승인 추천 여부를 판정하는 검증 스킬입니다.

## 🌟 Support & Usage

이 프로젝트가 마음에 드신다면 오른쪽 상단의 **Star (⭐)**를 눌러 응원해 주세요! 지속적인 업데이트와 개발에 큰 힘이 됩니다.

### 📌 사용 및 인용 안내 (Credit)
본 프로젝트를 설치하여 사용하시거나, 포크(Fork), 혹은 강의·발표 등에서 활용하실 경우 **개발자(출처)를 반드시 언급**해 주시기 바랍니다. 
> *예시: "본 강의/프로젝트는 [parkpro5]님의 hermes-skills/korus-verification을 바탕으로 제작/활용되었습니다."*

## ✨ 주요 특징

- **엄격한 정합성 점검 (Preflight)**: 검증 시작 전, 증빙 파일의 개수, 번호의 연속성, 파일명 일치 여부를 먼저 확인하여 데이터 누락을 방지합니다.
- **지능형 유연 매칭**:
  - **정규화 대조**: 성명, 기관명, 과정명 대조 시 모든 공백과 특수문자를 제거하여 `천 선 주` $\to$ `천선주`와 같은 표기 차이를 자동으로 해결합니다.
  - **핵심 과정명 추출**: `(기획력) 일잘러...`와 같이 과정명 앞뒤에 붙은 분류 태그를 자동으로 제거하고 핵심 키워드만으로 증빙 내 포함 여부를 판단합니다.
  - **PDF 줄바꿈 대응**: PDF 텍스트 추출 시 발생하는 줄바꿈 문자를 제거하여 문장이 끊겨 있어도 정확하게 인식합니다.
- **정밀한 시간 검증**: 증빙서류의 이수 시간을 추출하여 `증빙 시간 ≥ max(실적시간, 인정시간)` 조건이 충족되는지 검증합니다.
- **실무 최적화 결과물**:
  - `final_report.html`: 불일치 사유(목록 값 vs 증빙 값)가 상세히 명시된 웹 리포트.
  - `final_results.xlsx`: 판정, 확신도, 사유가 포함된 상세 결과 파일.
  - `checkbox_list.txt`: KORUS 화면에서 즉시 체크하여 반영할 수 있는 승인 추천 순번 목록.

## 🛠️ 작동 파이프라인

1. **Guidance**: 사용자에게 입력(`IN`) 및 출력(`OUT`) 폴더 준비사항 안내.
2. **Preflight**: 데이터 정합성(P1~P5) 점검.
3. **Matching**: 증빙 파일 인덱스($k$) 기반의 행-파일 매칭.
4. **Extraction**: PDF/XLSX 등에서 텍스트 및 시간 데이터 추출.
5. **Verification**: 6대 검증 항목(성명, 기관, 과정명, 시간, 형식, 발급주체) 대조.
6. **Reporting**: HTML/XLSX 리포트 및 체크리스트 생성.

## 📂 폴더 구조 및 설정

- **입력 폴더**: `D:\dev\hermes\pr01\IN\`
  - `*.xls` (KORUS 실적 엑셀 파일 1개)
  - `증빙\` (증빙 파일들이 저장된 폴더)
- **출력 폴더**: `D:\dev\hermes\pr01\OUT\`
  - `final_report.html`, `final_results.xlsx`, `checkbox_list.txt`

## 🚀 사용 방법

Hermes Agent에게 다음과 같이 요청하세요:
> "KORUS 교육훈련 실적 검증 시작해줘"

---
*This skill is designed for high-precision administrative verification and focuses on reducing manual cross-checking effort.*
