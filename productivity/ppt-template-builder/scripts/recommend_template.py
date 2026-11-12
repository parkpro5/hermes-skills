#!/usr/bin/env python3
"""만들 PPT의 주제·목적·청중·내용에 맞는 서식을 등록된 서식(templates/registry.json) 중에서 추천한다.

사용:
  python3 recommend_template.py --topic "온프레미스 AI 플랫폼 구축 현황" --purpose "경영진 보고" \\
      [--audience "총장단"] [--content "원문 요약이나 핵심 문장"] [--need 2단 --need 표] [--json]

점수는 키워드 일치로 만든 참고값이다. 결과(후보 설명, 미리보기)를 보고 의미를 직접 판단해
사용자에게 상위 2~3개를 보여 주고 선택받는다(점수가 모두 낮으면 새 서식 첨부를 권한다).
"""
import argparse
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from pptx_lib import TEMPLATES_DIR, load_registry  # noqa: E402


def hits(words, text):
    return [w for w in words if w and w.lower() in text]


def tokens(s):
    return {t for t in re.split(r'[\s,.·/()\[\]\-_:;]+', s.lower()) if len(t) >= 2}


def score(entry, text, needs):
    why, s = [], 0.0
    for key, w, label in (('purposes', 3, '용도'), ('tags', 2, '주제'), ('audiences', 2, '청중'), ('tone', 1, '분위기')):
        h = hits(entry.get(key, []), text)
        if h:
            s += w * len(h)
            why.append(f'+{w * len(h):g} {label} 일치: {", ".join(h)}')
    h = hits(entry.get('avoid', []), text)
    if h:
        s -= 4 * len(h)
        why.append(f'-{4 * len(h)} 어울리지 않는 용도: {", ".join(h)}')
    common = tokens(entry.get('description', '')) & tokens(text)
    if common:
        add = min(3, 0.5 * len(common))
        s += add
        why.append(f'+{add:g} 설명 단어 일치: {", ".join(sorted(common)[:5])}')
    f = entry.get('features', {})
    for n in needs:
        if '2단' in n or '두 열' in n or '비교' in n:
            if f.get('has_two_col'):
                s += 3
                why.append('+3 2단 레이아웃 있음')
            else:
                s -= 2
                why.append('-2 2단 레이아웃 없음(일반 본문으로 대체)')
    return round(s, 1), why


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--topic', default='')
    ap.add_argument('--purpose', default='')
    ap.add_argument('--audience', default='')
    ap.add_argument('--content', default='')
    ap.add_argument('--need', action='append', default=[], help='필요한 기능(예: 2단, 표)')
    ap.add_argument('--json', action='store_true')
    a = ap.parse_args()

    reg = load_registry()
    ts = reg['templates']
    if not ts:
        raise SystemExit('등록된 서식이 없습니다. 서식 파일을 첨부받아 auto_convert.py로 서식화하세요.')
    text = ' '.join([a.topic, a.purpose, a.audience, a.content] + a.need).lower()
    rows = []
    for e in ts:
        s, why = score(e, text, a.need)
        rows.append(dict(id=e['id'], name=e['name'], score=s, why=why, description=e.get('description', ''),
                         features=e.get('features', {}),
                         preview=(os.path.normpath(os.path.join(TEMPLATES_DIR, e['preview'])) if e.get('preview') else None),
                         default=(e['id'] == reg.get('default'))))
    rows.sort(key=lambda r: (-r['score'], not r['default']))
    best = rows[0]['score']
    verdict = ('적합한 서식 있음' if best >= 5 else '약한 후보만 있음' if best >= 2 else '일치하는 서식 없음')
    if a.json:
        print(json.dumps(dict(verdict=verdict, candidates=rows), ensure_ascii=False, indent=2))
        return
    print(f'# 서식 추천 ({verdict}) — 등록된 서식 {len(ts)}개')
    for i, r in enumerate(rows, 1):
        f = r['features']
        tone = ', '.join(f'{k.split("_")[0]} {v}' for k, v in f.items() if k.endswith('_tone'))
        print(f'\n{i}. {r["name"]} [{r["id"]}] 점수 {r["score"]}' + (' (기본 서식)' if r['default'] else ''))
        print(f'   {r["description"]}')
        print(f'   구성: {", ".join(f.get("layouts", []))}' + (f' | 밝기: {tone}' if tone else '')
              + (f' | 목차 최대 {f["max_agenda_items"]}개' if f.get('max_agenda_items') else ''))
        for w in r['why']:
            print('   ' + w)
        if r['preview']:
            print(f'   미리보기: {r["preview"]}')
    print('\n다음 단계: 상위 후보 2~3개와 미리보기를 사용자에게 보여 주고 선택받는다.')
    if best < 2:
        print('※ 점수가 모두 낮습니다. 사용자에게 서식 파일(또는 예전에 쓴 PPT) 첨부를 권하거나 기본 서식으로 진행할지 묻는다.')


if __name__ == '__main__':
    main()
