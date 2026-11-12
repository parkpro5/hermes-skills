#!/usr/bin/env python3
"""서식화된 pptx를 스킬의 서식 목록(templates/registry.json)에 등록한다.

사용자가 "이 서식을 스킬에 반영해줘"라고 요청했을 때만 쓴다. 스킬 폴더는 읽기 전용 캐시이므로
스킬 폴더를 작업 폴더로 복사한 뒤 --skill-dir로 가리키고, 끝나면 폴더를 zip으로 만들어 전달한다.

사용:
  python3 add_template.py 서식.pptx --id my-report --name "내 보고서" \\
      --description "한 줄 설명" --purposes "업무 보고,사업 발표" --audiences "경영진,교직원" \\
      --tone "공식,신뢰" --tags "AI,데이터" --avoid "캐주얼" [--skill-dir 스킬폴더] [--default] [--replace]
서식 파일은 auto_convert.py로 만든 것(roles 내장 또는 .roles.json 동반)이어야 한다.
특성(레이아웃, 목차 최대 항목 수, 2단 여부, 표지·본문 밝기)과 미리보기 이미지는 자동으로 채운다.
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from pptx import Presentation  # noqa: E402

from pptx_lib import embed_roles, load_roles  # noqa: E402


def csv(v):
    return [x.strip() for x in (v or '').split(',') if x.strip()]


def render_preview(pptx, out_png):
    """샘플 슬라이드를 렌더링해 미리보기 이미지(contact sheet)와 슬라이드별 밝기를 만든다. 실패하면 (None, {})."""
    if not (shutil.which('soffice') and shutil.which('pdftoppm')):
        return None, {}
    tmp = tempfile.mkdtemp()
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, pptx],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
        pdf = os.path.join(tmp, os.path.splitext(os.path.basename(pptx))[0] + '.pdf')
        if not os.path.exists(pdf):
            return None, {}
        subprocess.run(['pdftoppm', '-r', '40', '-png', pdf, os.path.join(tmp, 's')])
        from PIL import Image, ImageStat
        files = sorted(f for f in os.listdir(tmp) if f.startswith('s-') and f.endswith('.png'))
        ims = [Image.open(os.path.join(tmp, f)).convert('RGB') for f in files]
        lum = {i: round(ImageStat.Stat(im.convert('L')).mean[0]) for i, im in enumerate(ims)}
        w, h = ims[0].size
        cols = 3
        rows = (len(ims) + cols - 1) // cols
        sheet = Image.new('RGB', (w * cols, h * rows), 'white')
        for i, im in enumerate(ims):
            sheet.paste(im, ((i % cols) * w, (i // cols) * h))
        os.makedirs(os.path.dirname(out_png), exist_ok=True)
        sheet.save(out_png)
        return out_png, lum
    except Exception as e:      # 미리보기는 선택 사항
        print('미리보기 생성 실패:', e)
        return None, {}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def features_of(pptx, roles, lum):
    prs = Presentation(pptx)
    layouts = [l.name for m in prs.slide_masters for l in m.slide_layouts if l.name in roles.get('layout_fields', {})]
    if not layouts:                                  # 복제 방식 서식: 레이아웃이 없고 샘플 슬라이드가 곧 유형
        layouts = [t for t in ('cover', 'agenda', 'section', 'content', 'closing') if t in roles]
    f = dict(mode=roles.get('mode'), types=[t for t in ('cover', 'agenda', 'section', 'content', 'closing')
                                            if t in roles.get('layouts', {}) or t in roles],
             layouts=layouts, max_agenda_items=(roles.get('agenda') or {}).get('max_items'),
             has_two_col=any('2단' in n for n in layouts), n_samples=len(prs.slides))
    ss = roles.get('sample_slides', {})
    for key, t in (('cover_tone', 'cover'), ('content_tone', 'content')):
        i = ss.get(t) if isinstance(ss.get(t), int) else (roles.get(t) or {}).get('slide')
        if i is not None and i in lum:
            f[key] = '어두움' if lum[i] < 110 else '밝음'
    return f


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('--id', required=True, help='영문 소문자·숫자·하이픈 (예: seoultech-ax)')
    ap.add_argument('--name', required=True)
    ap.add_argument('--description', default='')
    ap.add_argument('--purposes', default='', help='어떤 발표·문서에 쓰는지(쉼표로 구분)')
    ap.add_argument('--audiences', default='', help='주 청중')
    ap.add_argument('--tone', default='', help='분위기(공식, 신뢰, 캐주얼 등)')
    ap.add_argument('--tags', default='', help='주제 키워드')
    ap.add_argument('--avoid', default='', help='어울리지 않는 용도')
    ap.add_argument('--skill-dir', default=os.path.join(HERE, '..'))
    ap.add_argument('--default', action='store_true', help='기본 서식으로 지정')
    ap.add_argument('--replace', action='store_true', help='같은 id가 있으면 덮어쓰기')
    a = ap.parse_args()

    if not re.fullmatch(r'[a-z0-9][a-z0-9-]*', a.id):
        raise SystemExit('id는 영문 소문자, 숫자, 하이픈만 쓸 수 있습니다.')
    roles = load_roles(a.pptx)
    tdir = os.path.join(a.skill_dir, 'templates')
    os.makedirs(tdir, exist_ok=True)
    reg_p = os.path.join(tdir, 'registry.json')
    reg = json.load(open(reg_p, encoding='utf8')) if os.path.exists(reg_p) else dict(default=None, templates=[])
    old = next((e for e in reg['templates'] if e['id'] == a.id), None)
    if old and not a.replace:
        raise SystemExit(f"id '{a.id}'가 이미 등록되어 있습니다. 덮어쓰려면 --replace를 쓰세요.")

    fname = os.path.basename(a.pptx) if os.path.abspath(os.path.dirname(a.pptx)) == os.path.abspath(tdir) else f'{a.id}.pptx'
    dst = os.path.join(tdir, fname)
    if os.path.abspath(a.pptx) != os.path.abspath(dst):
        shutil.copyfile(a.pptx, dst)
    roles['name'] = a.id
    embed_roles(dst, roles)
    with open(os.path.splitext(dst)[0] + '.roles.json', 'w', encoding='utf8') as f:
        json.dump(roles, f, ensure_ascii=False, indent=2)

    prev, lum = render_preview(dst, os.path.join(tdir, 'previews', f'{a.id}.png'))
    entry = dict(id=a.id, name=a.name, file=fname, description=a.description,
                 purposes=csv(a.purposes), audiences=csv(a.audiences), tone=csv(a.tone), tags=csv(a.tags),
                 avoid=csv(a.avoid), features=features_of(dst, roles, lum),
                 preview=(f'previews/{a.id}.png' if prev else None))
    reg['templates'] = [e for e in reg['templates'] if e['id'] != a.id] + [entry]
    if a.default or not reg.get('default'):
        reg['default'] = a.id
    with open(reg_p, 'w', encoding='utf8') as f:
        json.dump(reg, f, ensure_ascii=False, indent=2)
    print(f"등록 완료: {a.id} ({a.name}) → {dst}")
    print('특성:', json.dumps(entry['features'], ensure_ascii=False))
    if prev:
        print('미리보기:', os.path.join(tdir, entry['preview']))
    print(f"기본 서식: {reg['default']} / 등록된 서식 {len(reg['templates'])}개")


if __name__ == '__main__':
    main()
