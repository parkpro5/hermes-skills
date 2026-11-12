#!/usr/bin/env python3
"""생성된 PPT를 검사한다(서식 준수, 배치, 넘침, 구성) + 렌더링.

사용:  python3 verify_deck.py result.pptx [--template T.pptx] [--font "Malgun Gothic"]
                                [--spec spec.json] [--outdir out]
종료 코드: FAIL이 있으면 1, 없으면 0
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from pptx import Presentation  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

from pptx_lib import (EMU_PT, detect_role, est_lines, inherited_size_pt,  # noqa: E402
                      iter_shapes, load_roles, resolve_template, ph_idx, rect_intersects, shape_text,
                      text_width_pt)

results = []


def add(level, slide_no, check, msg):
    results.append((level, slide_no, check, msg))


def run_fonts(run):
    rPr = run._r.find(qn('a:rPr'))
    out = {}
    if rPr is not None:
        for tag in ('a:latin', 'a:ea'):
            el = rPr.find(qn(tag))
            out[tag] = el.get('typeface') if el is not None else None
    return out


def all_runs(slide):
    for sh in iter_shapes(slide.shapes):
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    yield sh, r
        if getattr(sh, 'has_table', False) and sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    for p in c.text_frame.paragraphs:
                        for r in p.runs:
                            yield sh, r


def run_size_pt(sh, run):
    if run.font.size:
        return run.font.size.pt
    s = inherited_size_pt(sh)
    if s:
        return s
    for el in sh._element.iter(qn('a:defRPr'), qn('a:rPr')):
        sz = el.get('sz')
        if sz:
            return int(sz) / 100
    return 18.0


def wrap_is_none(sh):
    """줄바꿈 없음(wrap=none) 여부. 플레이스홀더는 슬라이드에 없으면 레이아웃 값을 상속한다."""
    bp = sh.text_frame._txBody.find(qn('a:bodyPr'))
    if bp is not None and bp.get('wrap'):
        return bp.get('wrap') == 'none'
    if getattr(sh, 'is_placeholder', False):
        try:
            layout = sh.part.slide.slide_layout
        except Exception:
            return False
        for lph in layout.placeholders:
            if lph.placeholder_format.idx == sh.placeholder_format.idx:
                lbp = lph.text_frame._txBody.find(qn('a:bodyPr'))
                return lbp is not None and lbp.get('wrap') == 'none'
    return False


def is_body(sh, roles):
    if sh.name == 'BODY_TABLE' or (getattr(sh, 'has_table', False) and sh.has_table):
        return True
    if sh.name.startswith('BODY_'):
        return True
    return ph_idx(sh) == roles['content'].get('body', -1)


def title_shape(sl, roles, role):
    t = roles[role].get('title')
    if isinstance(t, int):
        for sh in sl.placeholders:
            if sh.placeholder_format.idx == t:
                return sh
        return None
    if isinstance(t, str) and t.startswith('#'):
        return next((sh for sh in sl.shapes if str(sh.shape_id) == t[1:]), None)
    return next((sh for sh in sl.shapes if sh.name == t), None)


def check_fonts(prs, tpl_fonts, font):
    used = set()
    for n, sl in enumerate(prs.slides, 1):
        for sh, r in all_runs(sl):
            if not r.text.strip():
                continue
            f = run_fonts(r)
            used.add((f.get('a:latin') or r.font.name or '(서식 상속)', f.get('a:ea') or '(서식 상속)'))
            if font:
                if f.get('a:latin') != font or f.get('a:ea') != font:
                    add('FAIL', n, '폰트', f"'{r.text[:20]}' 폰트가 지정({font})과 다름: latin={f.get('a:latin')}, ea={f.get('a:ea')}")
            else:
                names = {x for x in (f.get('a:latin'), f.get('a:ea'), r.font.name) if x}
                if names and not names & tpl_fonts and not any(x.startswith('+') for x in names):
                    add('WARN', n, '폰트', f"서식에 없는 폰트 사용: {sorted(names)} ('{r.text[:20]}')")
    return used


def check_samples(prs, roles, allowed_text):
    samples = roles.get('sample_strings', [])
    for n, sl in enumerate(prs.slides, 1):
        for sh in iter_shapes(sl.shapes):
            t = shape_text(sh)
            if not t.strip():
                continue
            if re.search(r'O{4,}', t):
                add('FAIL', n, '샘플 문구', f"'OOOO' 자리표시자가 남아 있음: '{t[:30]}'")
            for s in samples:
                if s in t and s not in allowed_text:
                    add('FAIL', n, '샘플 문구', f"서식 샘플 문구가 남아 있음: '{s}'")


def check_placeholders(prs):
    """비어 있는 플레이스홀더(안내 문구만 보이는 칸)가 남았는지."""
    for n, sl in enumerate(prs.slides, 1):
        for sh in sl.placeholders:
            if 'SLIDE_NUMBER' in str(sh.placeholder_format.type):
                continue
            if sh.has_text_frame and not sh.text_frame.text.strip():
                add('FAIL', n, '빈 플레이스홀더', f"'{sh.name}'(idx {sh.placeholder_format.idx})이 비어 있음")


def check_bounds_and_fit(prs, roles, seq):
    W, H = prs.slide_width, prs.slide_height
    tol = int(0.01 * W)
    for n, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            has_table = getattr(sh, 'has_table', False) and sh.has_table
            has_text = (sh.has_text_frame and sh.text_frame.text.strip()) or has_table
            if not has_text or sh.left is None:
                continue
            if sh.left < -tol or sh.top < -tol or sh.left + sh.width > W + tol or sh.top + sh.height > H + tol:
                add('FAIL', n, '경계', f"'{sh.name}'이 슬라이드 밖으로 나감")
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            tf = sh.text_frame
            wrap_none = wrap_is_none(sh)
            width_pt = sh.width / EMU_PT - 14.4
            need = 0.0
            for p in tf.paragraphs:
                txt = ''.join(r.text for r in p.runs)
                if not p.runs or not txt:
                    continue
                size = max(run_size_pt(sh, r) for r in p.runs)
                if wrap_none:
                    right = sh.left + (text_width_pt(txt, size) + 14.4) * EMU_PT
                    if right > W:
                        add('FAIL', n, '넘침', f"'{sh.name}' 한 줄 텍스트가 슬라이드 오른쪽을 넘김: '{txt[:20]}'")
                    continue
                need += est_lines(txt, size, width_pt) * size * 1.2
            if not wrap_none and need > (sh.height / EMU_PT) * 1.15 and not is_body(sh, roles):
                add('FAIL', n, '넘침', f"'{sh.name}' 텍스트가 상자 높이를 넘을 것으로 추정(필요 {need:.0f}pt > {sh.height / EMU_PT:.0f}pt): '{shape_text(sh)[:20]}'")
    c = roles['content']
    bx, by, bw, bh = c['body_area']
    for n, sl in enumerate(prs.slides, 1):
        if seq[n - 1] != 'content':
            continue
        for sh in sl.shapes:
            if not is_body(sh, roles):
                continue
            rect = (sh.left, sh.top, sh.width, sh.height)
            if sh.left < bx - 20000 or sh.top < by - 20000 or sh.left + sh.width > bx + bw + 20000 \
                    or sh.top + sh.height > by + bh + 20000:
                add('FAIL', n, '본문 영역', f"'{sh.name}'이 본문 영역을 벗어남")
            for ng in c.get('no_go', []):
                if rect_intersects(rect, tuple(ng), tol=20000):
                    add('FAIL', n, '금지 영역', f"'{sh.name}'이 제목/로고/장식/페이지 번호 영역과 겹침")
            if sh.has_text_frame and not (getattr(sh, 'has_table', False) and sh.has_table):
                wp = sh.width / EMU_PT - 14.4
                need = 0.0
                for p in sh.text_frame.paragraphs:
                    txt = ''.join(r.text for r in p.runs)
                    lvl = p.level
                    size = 20 if lvl == 0 else 18
                    need += est_lines(txt, size, wp - 22.5 * (lvl + 1)) * size * 1.25 + 8
                if need > bh / EMU_PT:
                    add('FAIL', n, '넘침', f'본문 텍스트가 본문 영역보다 길 것으로 추정({need:.0f}pt > {bh / EMU_PT:.0f}pt)')
                elif need > bh / EMU_PT * 0.92:
                    add('WARN', n, '넘침', f'본문이 영역을 거의 채움({need:.0f}pt / {bh / EMU_PT:.0f}pt)')
        title = title_shape(sl, roles, 'content')
        if title is not None and len(title.text_frame.text) > c.get('title_max_chars', 36):
            add('FAIL', n, '제목 길이', f"제목이 한 줄({c.get('title_max_chars', 36)}자)을 넘음: '{title.text_frame.text}'")


def agenda_titles(sl, roles):
    out = []
    if roles.get('mode') == 'placeholder':
        for it in roles['agenda']['items']:
            for sh in sl.placeholders:
                if sh.placeholder_format.idx == it['title'] and sh.text_frame.text.strip():
                    out.append(sh.text_frame.text.strip())
        return out
    r_items = roles['agenda'].get('items', [])
    if r_items and str(r_items[0].get('title', '')).startswith('#'):      # 자동 생성 roles(도형 id)
        for it in r_items:
            sh = next((x for x in sl.shapes if str(x.shape_id) == str(it['title'])[1:]), None)
            if sh is not None and sh.has_text_frame and sh.text_frame.text.strip():
                out.append(sh.text_frame.text.strip())
        return out
    items = sorted((sh for sh in sl.shapes if sh.has_text_frame and re.fullmatch(r'\d\d', sh.text_frame.text.strip())),
                   key=lambda s: s.text_frame.text)
    for num in items:
        near = [sh for sh in sl.shapes if sh.has_text_frame and sh.text_frame.text.strip()
                and sh.left > num.left and abs((sh.top + sh.height / 2) - (num.top + num.height / 2)) < 400000
                and sh.height < 700000]
        if near:
            out.append(sorted(near, key=lambda s: s.top)[0].text_frame.text.strip())
    return out


def check_structure(prs, roles):
    seq = [detect_role(s, roles) for s in prs.slides]
    if not seq:
        add('FAIL', 0, '구성', '슬라이드가 없음')
        return seq
    if seq[0] != 'cover':
        add('WARN', 1, '구성', '첫 슬라이드가 표지가 아님')
    if seq[-1] != 'closing':
        add('WARN', len(seq), '구성', '마지막 슬라이드가 마무리가 아님')
    if 'content' not in seq:
        add('FAIL', 0, '구성', '본문 슬라이드가 없음')
    for n, sl in enumerate(prs.slides, 1):
        if seq[n - 1] == 'content' and not any(is_body(sh, roles) for sh in sl.shapes):
            add('FAIL', n, '구성', '본문 내용이 비어 있음')
    ag, sec = [], []
    for n, sl in enumerate(prs.slides, 1):
        if seq[n - 1] == 'agenda':
            ag += agenda_titles(sl, roles)
        if seq[n - 1] == 'section':
            t = title_shape(sl, roles, 'section')
            if t is not None:
                sec.append(t.text_frame.text.strip())
    if sec and ag and ag != sec:
        add('WARN', 0, '구성', f'목차 항목과 구분 슬라이드 제목이 다름\n   목차: {ag}\n   구분: {sec}')
    for i, sl in enumerate(prs.slides, 1):
        if seq[i - 1] == 'section':
            t = [shape_text(sh) for sh in sl.shapes if shape_text(sh).startswith('PART ')]
            m = re.match(r'PART (\d+) / (\d+)', t[0]) if t else None
            if m and int(m.group(2)) != len(sec):
                add('FAIL', i, '구성', f'구분 번호 총계({m.group(2)})가 실제 구분 슬라이드 수({len(sec)})와 다름')
    return seq


def render(pptx, outdir):
    os.makedirs(outdir, exist_ok=True)
    for f in os.listdir(outdir):                      # 이전 검증의 이미지가 섞이지 않게 비운다
        if f.startswith('slide-') or f in ('contact_sheet.png',) or f.endswith('.pdf'):
            os.remove(os.path.join(outdir, f))
    if not shutil.which('soffice'):
        add('WARN', 0, '렌더링', 'soffice가 없어 렌더링 검사를 건너뜀(기하 검사만 수행)')
        return None
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', outdir, pptx],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + '.pdf')
    if not os.path.exists(pdf):
        add('WARN', 0, '렌더링', 'PDF 변환 실패')
        return None
    if shutil.which('pdftoppm'):
        subprocess.run(['pdftoppm', '-r', '60', '-png', pdf, os.path.join(outdir, 'slide')])
    try:
        from PIL import Image
        files = sorted(f for f in os.listdir(outdir) if f.startswith('slide') and f.endswith('.png'))
        if files:
            ims = [Image.open(os.path.join(outdir, f)) for f in files]
            w, h = ims[0].size
            cols = 3
            rows = (len(ims) + cols - 1) // cols
            sheet = Image.new('RGB', (w * cols, h * rows), 'white')
            for i, im in enumerate(ims):
                sheet.paste(im, ((i % cols) * w, (i // cols) * h))
            sheet.save(os.path.join(outdir, 'contact_sheet.png'))
    except Exception:
        pass
    return pdf


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('--template', default=None, help='서식 파일 경로 또는 등록된 서식 id(생략 시 기본 서식)')
    ap.add_argument('--font', default=None)
    ap.add_argument('--spec', default=None)
    ap.add_argument('--outdir', default='verify_out')
    a = ap.parse_args()

    a.template = resolve_template(a.template)
    roles = load_roles(a.template)
    prs = Presentation(a.pptx)
    tpl = Presentation(a.template)
    tpl_fonts = set()
    for sl in tpl.slides:
        for sh, r in all_runs(sl):
            f = run_fonts(r)
            tpl_fonts |= {x for x in (r.font.name, f.get('a:latin'), f.get('a:ea')) if x}
    tpl_fonts |= {'Pretendard', 'Pretendard ExtraBold', 'Malgun Gothic', 'Courier New', 'Arial'}

    allowed = ''
    if a.spec:
        with open(a.spec, encoding='utf8') as f:
            allowed = json.dumps(json.load(f), ensure_ascii=False)

    used = check_fonts(prs, tpl_fonts, a.font)
    check_samples(prs, roles, allowed)
    if roles.get('mode') == 'placeholder':
        check_placeholders(prs)
    seq = check_structure(prs, roles)
    check_bounds_and_fit(prs, roles, seq)
    pdf = render(a.pptx, a.outdir)
    if pdf and shutil.which('pdfinfo'):
        info = subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout
        m = re.search(r'Pages:\s+(\d+)', info)
        if m and int(m.group(1)) != len(prs.slides):
            add('FAIL', 0, '렌더링', f'PDF 쪽수({m.group(1)})와 슬라이드 수({len(prs.slides)})가 다름')

    fails = [r for r in results if r[0] == 'FAIL']
    warns = [r for r in results if r[0] == 'WARN']
    print(f'슬라이드 {len(prs.slides)}장: {" → ".join(seq)}')
    print(f'사용 폰트(latin, ea): {sorted(map(str, used))}')
    if pdf:
        print(f'렌더링: {a.outdir}/ (slide-N.png, contact_sheet.png)  ※ 폰트가 시스템에 없으면 대체 폰트로 보임')
    print()
    if not results:
        print('검사 결과: 이상 없음')
    else:
        print('| 수준 | 슬라이드 | 항목 | 내용 |')
        print('|---|---|---|---|')
        for lv, n, ck, msg in results:
            print(f'| {lv} | {n or "-"} | {ck} | {msg} |')
    print(f'\nFAIL {len(fails)}건, WARN {len(warns)}건')
    sys.exit(1 if fails else 0)


if __name__ == '__main__':
    main()
