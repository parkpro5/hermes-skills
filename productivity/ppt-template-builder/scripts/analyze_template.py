#!/usr/bin/env python3
"""서식 구조 분석: 레이아웃/플레이스홀더, 슬라이드별 도형(이름, 위치, 폰트)을 출력하고 PNG로 렌더링한다.

사용:  python3 analyze_template.py template.pptx [--outdir tpl_png]
"""
import argparse
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from pptx import Presentation  # noqa: E402

from pptx_lib import detect_role, iter_shapes  # noqa: E402


def walk(shapes, depth=0):
    for sh in shapes:
        info = ''
        if sh.has_text_frame and sh.text_frame.text.strip():
            fonts = {(r.font.name, r.font.size.pt if r.font.size else None, r.font.bold)
                     for p in sh.text_frame.paragraphs for r in p.runs}
            info = f'{sh.text_frame.text[:25]!r} {sorted(map(str, fonts))}'
        print('  ' * depth, sh.shape_type, repr(sh.name), (sh.left, sh.top, sh.width, sh.height), info)
        if sh.shape_type == 6:
            walk(sh.shapes, depth + 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('template')
    ap.add_argument('--outdir', default='tpl_png')
    a = ap.parse_args()
    prs = Presentation(a.template)
    print('슬라이드 크기', prs.slide_width, prs.slide_height, '/ 슬라이드', len(prs.slides), '장')
    n_ph = 0
    for i, lay in enumerate(prs.slide_layouts):
        phs = [(ph.placeholder_format.idx, str(ph.placeholder_format.type)) for ph in lay.placeholders]
        n_ph += len(phs)
        print(f'[layout {i}] {lay.name} 플레이스홀더: {phs}')
    print()
    if n_ph == 0:
        print('>> 레이아웃에 플레이스홀더가 없습니다. 슬라이드 복제 방식(build_deck.py)을 사용하세요.\n')
    for n, s in enumerate(prs.slides, 1):
        print(f'--- 슬라이드 {n} (추정 역할: {detect_role(s)}, 레이아웃: {s.slide_layout.name})')
        walk(s.shapes)
    if shutil.which('soffice'):
        os.makedirs(a.outdir, exist_ok=True)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', a.outdir, a.template],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdf = os.path.join(a.outdir, os.path.splitext(os.path.basename(a.template))[0] + '.pdf')
        if os.path.exists(pdf) and shutil.which('pdftoppm'):
            subprocess.run(['pdftoppm', '-r', '50', '-png', pdf, os.path.join(a.outdir, 'tpl')])
            print(f'\n렌더링 이미지: {a.outdir}/tpl-N.png (눈으로 확인하세요)')


if __name__ == '__main__':
    main()
