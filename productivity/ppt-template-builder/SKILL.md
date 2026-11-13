---
name: ppt-template-builder
description: PPT 서식(첨부 서식 또는 예전에 쓴 PPT)의 디자인으로 지정 폰트의 PPT를 만들고 스스로 검증·수정한다. "이 파일을 서식으로 만들어줘/템플릿화"하면 플레이스홀더 서식(파일 하나로 공유 가능)으로 변환하고, 복제 vs 플레이스홀더 방식을 추천해 결정받는다. 등록된 서식이 여러 개면 주제·목적·청중에 맞는 서식을 추천한다. 번들 서식은 서울과학기술대학교 AX 서식.
---

# PPT 서식 기반 생성 + 자동 서식화 + 서식 추천 + 자가 검증

서식(또는 실제 사용했던 PPT)의 디자인으로 슬라이드를 만들고, 결과를 검사 스크립트와 렌더링으로 스스로 검수해 고친다.
셸과 Python이 있으면 어떤 에이전트(Claude, Hermes 등)에서도 동작한다.

## 패키지 구성
```
ppt-template-builder/
├─ SKILL.md
├─ scripts/
│  ├─ inspect_template.py    서식 분석 + 방식 추천(복제 vs 플레이스홀더) 보고서
│  ├─ auto_convert.py        서식·실사용 자료 → 서식 pptx 자동 생성(roles 내장, --demo로 미리보기)
│  ├─ recommend_template.py  주제·목적·청중에 맞는 등록 서식 추천
│  ├─ add_template.py        서식을 스킬 목록(registry.json)에 등록(사용자가 요청할 때만)
│  ├─ build_deck.py          spec.json → PPT 생성
│  ├─ verify_deck.py         검사 + 렌더링(FAIL이 있으면 종료 코드 1)
│  ├─ tpl_analysis.py, pptx_lib.py   분석·공통 모듈
│  ├─ analyze_template.py    도형 단위 구조 덤프(수동 확인용)
│  └─ make_placeholder_template.py  번들 서식 전용 변환기(참고용)
├─ templates/
│  ├─ registry.json          등록된 서식 목록(용도, 청중, 분위기, 주제, 특성)
│  ├─ <id>.pptx (+ .roles.json), previews/<id>.png
└─ examples/ spec.example.json, roles.duplicate-mode.example.json
```
서식 pptx에는 roles(레이아웃·플레이스홀더 설정)가 안에 내장되어 있어 **pptx 한 파일만 공유해도** 다른 사람이 이 스킬로 그 서식을 쓸 수 있다.

## 요청 유형별 흐름
| 사용자 요청 | 흐름 |
|---|---|
| "이 파일을 서식으로 만들어줘" (서식화만) | **A. 서식화 전용** |
| "이 서식으로 PPT 만들어줘" (서식 첨부 또는 등록 서식 지정) | **B. 서식으로 PPT 만들기** |
| 서식을 정하지 않고 PPT를 요청, 또는 "어떤 서식이 좋아?" | **C. 서식 추천** 후 B |
| "이 서식을 스킬에 반영해줘" | **D. 스킬에 등록** |

## 준비
- 필요: Python 3, `python-pptx`(`pip install python-pptx`), LibreOffice(`soffice`), poppler(`pdftoppm`). Pillow가 있으면 `contact_sheet.png`도 만든다.
- soffice, pdftoppm이 없으면 렌더링 검사는 건너뛰고 그 사실을 보고한다.
- 한글 폰트(Pretendard, 맑은 고딕)가 없으면 렌더링 이미지는 대체 폰트로 보인다. 파일 안의 폰트 지정이 맞으면 통과로 본다.
- 서식의 일부 글자(표지 날짜 등)는 그라디언트·투명도 서식이라 LibreOffice에서 안 보일 수 있다(원본도 같음).

## 입력
- 서식: 번들 서식을 쓰거나, 사용자가 첨부한 `.pptx`(샘플 서식 또는 **예전에 실제로 만든 PPT**)를 쓴다.
- 내용: 주제, 원문, 구성. 없으면 합리적으로 정하고 가정을 보고한다.
- 지정 폰트: 있으면 모든 글자(AGENDA, PART 같은 라벨 포함)에 한글(`ea`)과 영문(`latin`)으로 적용한다.

## 작업 순서

### A. 서식화 전용 (실제 문서 → 서식)
1. 사용 의도(once/repeat/edit)를 한 번 묻는다. 모르면 생략한다.
2. 분석·추천: `python3 scripts/inspect_template.py 첨부.pptx --intent repeat --outdir insp`
   슬라이드를 표지/목차/구분/본문/마무리로 분류하고 같은 디자인끼리 묶어(예: 본문 12장, 본문 2단 2장) 바뀌는 칸과 고정 문구를 찾는다. 이어서 플레이스홀더 변환 적합도 점수와 근거, 두 방식의 장단점, 실행 명령이 나온다. `insp/source_contact_sheet.png`도 직접 열어 본다.
3. 결정받기: "결정 필요" 블록을 보여 주고 선택받는다(AskUserQuestion이 있으면 추천 방식을 첫 선택지로). 답이 없으면 추천 방식으로 진행하고 그렇게 했다고 보고한다.
   - 플레이스홀더 권장: 같은 서식을 여러 번 쓰거나 사람이 이어서 편집할 때, 목차 항목 수·글 길이가 달라질 때, 유형 분류가 명확할 때.
   - 복제 권장: 한 번만 쓸 때, 그룹 안 글자·SmartArt·회전 글자가 많거나 장식이 복잡해 변환 시 모양이 틀어질 위험이 클 때, 슬라이드가 1~2장뿐일 때.
4. 변환 + 미리보기: `python3 scripts/auto_convert.py 첨부.pptx --mode placeholder --out 서식이름.pptx --demo` (또는 `--mode duplicate`)
   - 서식 pptx(roles 내장), `.roles.json`, 서식의 모든 유형을 보여 주는 `서식이름_demo.pptx`, 검증 결과, 미리보기 이미지(`서식이름_demo_check/contact_sheet.png`)가 생긴다.
   - 같은 디자인은 대표 1장만 샘플로 남는다. 목차는 항목 수별(1~N개) 레이아웃도 만든다. 표·차트·반복되지 않는 그림은 레이아웃이 아니라 슬라이드 내용으로 둔다.
5. 미리보기 이미지를 눈으로 확인한다. 모양이 원본과 다르면 `--mode duplicate`로 바꾸거나 inspect의 "확신이 낮은 판단"을 근거로 roles.json을 손본다.
6. 사용자에게 전달: **서식 pptx 한 파일**(roles 내장, 이것만 공유하면 된다), 미리보기 이미지, 사용법 한 줄("이 서식을 첨부하고 '이 서식으로 PPT 만들어줘'라고 요청"). 스킬에 영구 반영은 사용자가 원할 때만 D로 한다.

### B. 서식으로 PPT 만들기
- 서식은 사용자가 첨부한 pptx, 등록된 서식 id, 번들 서식 중 하나다. 첨부했으면 그 서식을 쓴다. roles가 없는 pptx(일반 PPT)면 먼저 A로 서식화한다.
1. 내용 설계 → `spec.json`(아래 표).
2. 생성: `python3 scripts/build_deck.py --spec spec.json --out result.pptx --template <파일 또는 id> [--font "Malgun Gothic"]`
3. 자가 검증(최대 3회): `python3 scripts/verify_deck.py result.pptx --spec spec.json --template <파일 또는 id> [--font ...] --outdir verify_out`
4. 이미지 육안 확인 → 보고(결과 파일, 구성 요약, 검증 결과, 가정, 사용한 서식과 이유).

spec 종류 (전체 예시는 `examples/spec.example.json`)

| type | 필드 | 비고 |
|---|---|---|
| `cover` | `title`(문자열 목록, 2줄), `subtitle`, `date_org`, `date` | 둘째 줄 첫 단어에 강조색 |
| `agenda` | `title`, `items`[{`title`, `org`}] | 한 장 최대 항목 수(서식에서 감지) 초과 시 자동 분할 |
| `section` | `title`, `subtitle`, (`part`) | `PART 01 / 05`는 자동 계산 |
| `content` | `title`, `bullets`, `table`{`header`, `rows`} | `bullets` 항목은 `"문장"` 또는 `["문장", 1]`(들여쓰기). 넘치면 쪽 나눔 |
| `closing` | `text`, `date` | 기본 "감사합니다" |
| `custom` | `layout`, `fields`{필드명: 문자열 또는 목록} | 서식의 다른 레이아웃(예: `본문 2단`)을 직접 지정. 필드명은 roles의 `layout_fields` 참고(title, body, body2 등) |

작성 원칙: 제목은 한 줄(roles의 `title_max_chars` 이하). 본문은 핵심 3~6개. 목차 항목과 구분 제목은 같은 순서·이름. 표지 → 목차 → (구분 → 본문…) → 마무리. 서식에 없는 유형은 spec에 쓰면 오류가 나므로 빼거나 `custom`을 쓴다.

검증 항목: 폰트 일치, 샘플 문구 잔존, 빈 플레이스홀더, 경계 이탈, 넘침 추정, 제목 길이, 금지 영역 겹침, 구성(표지·마무리, 목차↔구분 제목, PART 총계, 쪽수).
수정 규칙: 제목 길이→제목을 줄임 / 넘침→내용을 줄이거나 슬라이드를 나눔(글자 크기는 줄이지 않음) / 폰트→`--font`로 재생성 / 샘플 문구·빈 칸→spec에 필드를 넣거나 정리 / 목차·구분 불일치→이름·순서를 맞춤. 3회 후에도 남은 문제는 숨기지 말고 보고한다.

### C. 서식 추천 (서식을 정하지 않았을 때)
1. 만들 PPT의 주제, 목적, 청중, 필요한 기능(2단, 표 등)을 파악한다(부족하면 한 번만 묻는다).
2. `python3 scripts/recommend_template.py --topic "..." --purpose "..." --audience "..." [--content "핵심 내용"] [--need 2단]`
   등록된 서식별 점수, 근거, 구성, 미리보기 경로가 나온다. 점수는 키워드 일치 기반의 참고값이므로 후보 설명과 미리보기를 보고 **의미를 직접 판단**한다.
3. 상위 2~3개를 미리보기와 함께 보여 주고 선택받는다(AskUserQuestion이 있으면 추천을 첫 선택지로). 후보가 하나뿐이거나 답이 없으면 추천(또는 기본) 서식으로 진행하고 이유를 보고한다.
4. 점수가 모두 낮으면(일치하는 서식 없음) 기본 서식으로 진행할지, 서식 파일(또는 예전에 쓴 PPT)을 첨부할지 묻는다. 그다음 B로 간다.

### D. 스킬에 등록 (사용자가 "이 서식을 스킬에 반영해줘"라고 할 때만)
서식을 만들었다고 자동으로 등록하지 않는다. 스킬 폴더는 읽기 전용 캐시이므로 작업 폴더로 복사해서 수정하고 zip으로 전달한다.
1. 메타데이터를 정한다: id(영문 소문자·하이픈), 이름, 한 줄 설명, 용도(purposes), 청중(audiences), 분위기(tone), 주제 키워드(tags), 어울리지 않는 용도(avoid). 서식의 내용과 용도에서 합리적으로 정하되 사용자에게 한 번 확인받는다. 좋은 메타데이터가 추천 품질을 좌우한다.
2. `cp -r <스킬 폴더> ./skill_work && python3 skill_work/scripts/add_template.py 서식.pptx --id ... --name ... --description ... --purposes ... --audiences ... --tone ... --tags ... --avoid ... --skill-dir ./skill_work`
   서식·roles 복사, 특성(레이아웃, 목차 최대 항목, 2단 여부, 표지·본문 밝기)과 미리보기 자동 추출, registry.json 갱신이 이루어진다. 같은 id가 있으면 `--replace`.
3. 등록된 서식으로 간단한 spec을 빌드·검증해 보고, `recommend_template.py`가 기대한 주제에서 이 서식을 상위로 내는지 확인한다.
4. `zip -r ppt-template-builder.zip skill_work`(폴더 이름은 `ppt-template-builder`로 맞춘다)를 전달하고, 계정의 스킬을 이 zip으로 교체해 올려야 반영된다고 알린다.

## 자동 분석이 하는 일과 한계
- 유형 분류는 위치(첫/마지막 슬라이드), 문구(감사, PART, 목차), 번호 격자, 글자 크기로 추정한다. 같은 유형 슬라이드가 여러 장이면 문구가 바뀌는 칸은 가변, 항상 같은 칸은 고정(라벨·로고 문구)으로 본다. 장이 하나뿐이면 샘플 글자는 가변으로 본다.
- 그룹 안의 글자는 플레이스홀더로 바꾸지 못하고 고정 요소가 된다(추천 점수에서 감점).
- 번호 격자 목차는 항목 수별 레이아웃을 만들지만, 격자가 아닌 목차는 일반 본문으로 처리된다.
- 변환 결과는 LibreOffice와 python-pptx로만 확인했다. 중요한 서식은 PowerPoint에서도 열어 확인한다.

## 주의
- 결과 pptx와 서식 pptx는 저장할 때 `scripts/fix_dangling.py`가 끊어진 참조(관계 파일에 없는 hlinkClick 등)를 자동 제거한다. 이게 남아 있으면 PowerPoint가 "복구" 메시지를 띄운다. (ai-science-gradient의 '본문' 레이아웃에서 발견되어 수정함) 직접 점검하려면 `python3 scripts/fix_dangling.py 입력.pptx 출력.pptx`.
- 서식 원본 파일을 덮어쓰지 않는다. 로고·푸터·페이지 번호는 서식에 이미 있으므로 중복 추가하지 않는다(본문 페이지 번호는 빌더가 자동 삽입).
- 번들 서식의 폰트: 제목·본문 `Malgun Gothic`, 영문 라벨 `Courier New`, 날짜·마무리 `Pretendard`.
- 서식 pptx에 내장된 roles는 PowerPoint에서 다시 저장해도 대부분 유지되지만(customXml), 사라졌다면 `.roles.json`을 같이 두거나 `auto_convert.py`로 다시 서식화한다. PowerPoint 실제 환경에서의 동작은 아직 확인하지 못했다.
