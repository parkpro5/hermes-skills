#!/usr/bin/env python3
"""첨부된 서식(또는 실제 사용했던 자료)을 분석하고, 복제 방식과 플레이스홀더 방식 중 무엇이 나은지 추천한다.

사용:  python3 inspect_template.py 파일.pptx [--intent once|repeat|edit] [--outdir insp] [--json out.json]
  --intent  once   : 이번 한 번만 쓸 서식
            repeat : 앞으로 같은 서식으로 여러 번 만들 예정
            edit   : 만든 뒤 PowerPoint에서 사람이 직접 고칠 예정
            (생략하면 unknown. 사용자에게 물어본 뒤 다시 실행하면 추천이 정교해진다)
출력은 한국어 보고서이며, 마지막의 '결정 필요' 블록을 사용자에게 보여 주고 선택을 받는다.
"""
import argparse
import json
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import tpl_analysis as T  # noqa: E402

TYPE_KO = {'cover': '표지', 'agenda': '목차', 'section': '구분', 'content': '본문', 'closing': '마무리'}
FIELD_KO = {'title': '제목', 'subtitle': '부제', 'body': '본문', 'body2': '본문2', 'date': '날짜', 'date_org': '날짜+소속',
            'part': 'PART 번호', 'text': '문구'}


def render(path, outdir):
    if not (shutil.which('soffice') and shutil.which('pdftoppm')):
        return None
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', outdir, path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + '.pdf')
    if not os.path.exists(pdf):
        return None
    subprocess.run(['pdftoppm', '-r', '40', '-png', pdf, os.path.join(outdir, 'src')])
    try:
        from PIL import Image
        files = sorted(f for f in os.listdir(outdir) if f.startswith('src-') and f.endswith('.png'))
        ims = [Image.open(os.path.join(outdir, f)) for f in files]
        if not ims:
            return None
        w, h = ims[0].size
        cols = 4
        rows = (len(ims) + cols - 1) // cols
        sheet = Image.new('RGB', (w * cols, h * rows), 'white')
        for i, im in enumerate(ims):
            sheet.paste(im, ((i % cols) * w, (i // cols) * h))
        out = os.path.join(outdir, 'source_contact_sheet.png')
        sheet.save(out)
        return out
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('--intent', choices=['once', 'repeat', 'edit', 'unknown'], default='unknown')
    ap.add_argument('--outdir', default='insp')
    ap.add_argument('--json', default=None)
    a_ = ap.parse_args()

    a = T.analyze(a_.pptx)
    rec = T.recommend(a, a_.intent)
    m = rec['metrics']
    print(f'# 서식 분석: {os.path.basename(a_.pptx)}')
    kind = '실제 사용했던 자료로 보임(슬라이드 8장 이상)' if a['deck_like'] else '샘플 서식으로 보임'
    print(f'- 슬라이드 {a["n"]}장 → 서로 다른 디자인 {m["n_clusters"]}종, {kind}')
    print('- 슬라이드별 유형: ' + ' / '.join(f'{s["i"] + 1}:{TYPE_KO[s["type"]]}' for s in a['slides']))
    print()
    print('## 디자인 묶음(같은 모양의 슬라이드끼리)')
    for c in a['clusters']:
        extra = ''
        if c['type'] == 'agenda' and c['agenda']:
            extra = f', 항목 {len(c["agenda"]["items"])}개 격자'
        two = ' 2단' if c['key'][1] == 'two-col' else ''
        names = ', '.join(f'{FIELD_KO.get(f["name"], f["name"])}' for f in c['fields'])
        fixed = [s['text'][:12] for s in a['slides'][c['rep']]['shapes']
                 if s['id'] in c['decisions'] and c['decisions'][s['id']]['fixed'] and s['text']
                 and not (c['decisions'][s['id']].get('role') or '').startswith('keep')]
        print(f'- {TYPE_KO[c["type"]]}{two}: {len(c["slides"])}장(대표 {c["rep"] + 1}번{extra}) '
              f'| 바뀌는 칸: {names or "-"}' + (f' | 고정 문구: {", ".join(fixed)}' if fixed else ''))
    unsure = [(c['type'], f['name'], c['decisions'][f['s']['id']]['why']) for c in a['clusters'] for f in c['fields']
              if c['decisions'][f['s']['id']]['conf'] < 0.7]
    if unsure:
        print('\n※ 확신이 낮은 판단(변환 뒤 눈으로 확인하세요):')
        for t, n, why in unsure[:8]:
            print(f'  · {TYPE_KO[t]}/{FIELD_KO.get(n, n)}: {why}')
    print()
    print('## 추천')
    if rec['mode'] == 'as-is':
        print('- 이미 플레이스홀더 서식입니다. 변환 없이 `auto_convert.py --mode placeholder`로 roles.json만 만들면 됩니다.')
    else:
        label = '플레이스홀더 변환' if rec['mode'] == 'placeholder' else '복제 방식'
        print(f'- **{label} 권장** (플레이스홀더 적합도 {rec["score"]}/100, 확신 {rec.get("level", "-")}, 사용 의도: {rec["intent"]})')
    for x in rec['plus']:
        print('  ' + x)
    for x in rec['minus']:
        print('  ' + x)
    if rec['intent'] == 'unknown':
        print('  ※ 사용 의도를 알면 정확해집니다: 한 번만 쓸지(once), 계속 쓸지(repeat), 만든 뒤 직접 편집할지(edit).')

    print('\n## 결정 필요 (사용자에게 선택받기)')
    rp = rec['mode'] == 'placeholder'
    opts = [
        ('플레이스홀더 변환' + (' (추천)' if rp else ''),
         '레이아웃·플레이스홀더로 바꿔 새 서식을 만듭니다. 항목 수·글 길이 변화에 강하고 PowerPoint에서 "새 슬라이드"로 이어서 편집할 수 있습니다. 변환 때 모양이 달라질 수 있어 결과를 확인해야 합니다.',
         f'python3 scripts/auto_convert.py {a_.pptx} --mode placeholder --out templates/<이름>.pptx'),
        ('복제 방식' + (' (추천)' if rec['mode'] == 'duplicate' else ''),
         '원본 슬라이드를 그대로 두고 복제해 글자만 바꿉니다. 설정이 가볍고 원본 모양이 그대로지만, 목차 항목 수가 달라지면 빈 도형 정리가 필요하고 PowerPoint에서 새 슬라이드로 서식을 이어 쓸 수 없습니다.',
         f'python3 scripts/auto_convert.py {a_.pptx} --mode duplicate --out templates/<이름>.pptx'),
    ]
    if not rp:
        opts.reverse()
    for i, (t, d, cmd) in enumerate(opts, 1):
        print(f'{i}) {t}\n   {d}\n   실행: {cmd}')
    if a['n'] > 0:
        print('※ 먼저 복제 방식으로 바로 쓰고, 반복해서 쓰게 되면 나중에 변환하는 것도 가능합니다.')

    sheet = render(a_.pptx, a_.outdir)
    if sheet:
        print(f'\n원본 렌더링: {sheet}')
    if a_.json:
        slim = dict(recommendation={k: v for k, v in rec.items() if k != 'metrics'}, metrics=m,
                    clusters=[dict(type=c['type'], slides=[k + 1 for k in c['slides']], rep=c['rep'] + 1,
                                   fields=[f['name'] for f in c['fields']]) for c in a['clusters']])
        with open(a_.json, 'w', encoding='utf8') as f:
            json.dump(slim, f, ensure_ascii=False, indent=2)


if __name__ == '__main__':
    main()
