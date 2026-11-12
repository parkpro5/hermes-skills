#!/usr/bin/env python3
"""spec.json → 서식이 적용된 PPT 생성.

사용:  python3 build_deck.py --spec spec.json --out result.pptx [--template T.pptx] [--font "Malgun Gothic"]
"""
import argparse
import copy
import datetime
import json
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from lxml import etree  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.text import MSO_ANCHOR  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Emu, Pt  # noqa: E402

from pptx_lib import (apply_font, drop_slide, duplicate_slide, est_lines,  # noqa: E402
                      find_shape, load_roles, resolve_template, remove_shape, set_font,
                      set_lines, set_text, EMU_PT)

TEXT_COLOR = RGBColor(0x26, 0x26, 0x26)
ACCENT = '00387F'      # 서식 파란색
HEAD_FILL = RGBColor(0x00, 0x38, 0x7F)
BAND_FILL = RGBColor(0xF2, 0xF5, 0xFA)
SIZES = {0: 20, 1: 18}


# ---------------- 본문 쪽 나누기 ----------------
def norm_bullets(items):
    out = []
    for it in items or []:
        if isinstance(it, str):
            out.append((it, 0))
        elif isinstance(it, (list, tuple)):
            out.append((it[0], int(it[1])))
        else:
            out.append((it['text'], int(it.get('level', 0))))
    return out


def bullet_height_pt(text, level, width_pt):
    size = SIZES.get(level, 18)
    avail = width_pt - 14.4 - 22.5 * (level + 1)
    return est_lines(text, size, avail) * size * 1.25 + 8


def paginate_bullets(items, width_pt, height_pt):
    pages, cur, used = [], [], 0.0
    for text, level in items:
        h = bullet_height_pt(text, level, width_pt)
        if cur and used + h > height_pt * 0.92:
            pages.append(cur)
            cur, used = [], 0.0
        cur.append((text, level))
        used += h
    if cur:
        pages.append(cur)
    return pages or [[]]


def table_row_height_pt(row, col_w_pt, size=16):
    lines = max(est_lines(str(c), size, col_w_pt - 14.4) for c in row)
    return lines * size * 1.25 + 12


def paginate_table(header, rows, width_pt, height_pt):
    col_w = width_pt / len(header)
    head_h = table_row_height_pt(header, col_w)
    pages, cur, used = [], [], head_h
    for r in rows:
        h = table_row_height_pt(r, col_w)
        if cur and used + h > height_pt * 0.95:
            pages.append(cur)
            cur, used = [], head_h
        cur.append(r)
        used += h
    if cur:
        pages.append(cur)
    return pages or [[]]


# ---------------- 본문 요소 추가 ----------------
def _ppr(level):
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    pPr = etree.Element(f'{{{ns}}}pPr')
    marL = 285750 + 285750 * level
    pPr.set('marL', str(marL))
    pPr.set('indent', '-285750')
    spc = etree.SubElement(pPr, f'{{{ns}}}spcBef')
    etree.SubElement(spc, f'{{{ns}}}spcPts').set('val', '600' if level else '900')
    buClr = etree.SubElement(pPr, f'{{{ns}}}buClr')
    etree.SubElement(buClr, f'{{{ns}}}srgbClr').set('val', ACCENT)
    etree.SubElement(pPr, f'{{{ns}}}buFont').set('typeface', 'Arial')
    etree.SubElement(pPr, f'{{{ns}}}buChar').set('char', '•' if level == 0 else '–')
    return pPr


def add_bullets(slide, area, items, font):
    x, y, w, h = area
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tb.name = 'BODY_TEXT'
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    for i, (text, level) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = text
        r.font.size = Pt(SIZES.get(level, 18))
        r.font.bold = False
        r.font.color.rgb = TEXT_COLOR
        set_font(r, font)
        old = p._p.find(qn('a:pPr'))
        if old is not None:
            p._p.remove(old)
        p._p.insert(0, _ppr(level))
    return tb


def add_table(slide, area, header, rows, font, top_offset=0):
    x, y, w, h = area
    y += top_offset
    n_rows, n_cols = len(rows) + 1, len(header)
    col_w = (w / EMU_PT) / n_cols
    gs = slide.shapes.add_table(n_rows, n_cols, Emu(x), Emu(y), Emu(w), Emu(Pt(30) * n_rows))
    gs.name = 'BODY_TABLE'
    tbl = gs.table
    heights = [table_row_height_pt(header, col_w)] + [table_row_height_pt(r, col_w) for r in rows]
    for i, hh in enumerate(heights):
        tbl.rows[i].height = Emu(int(hh * EMU_PT))
    gs.height = Emu(int(sum(heights) * EMU_PT))
    for ci in range(n_cols):
        tbl.columns[ci].width = Emu(int(w / n_cols))
    for ri in range(n_rows):
        for ci in range(n_cols):
            cell = tbl.cell(ri, ci)
            text = header[ci] if ri == 0 else str(rows[ri - 1][ci])
            cell.text = text
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = HEAD_FILL if ri == 0 else (BAND_FILL if ri % 2 == 0 else RGBColor(255, 255, 255))
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(16)
                    r.font.bold = ri == 0
                    r.font.color.rgb = RGBColor(255, 255, 255) if ri == 0 else TEXT_COLOR
                    set_font(r, font)
    return gs


# ---------------- 역할별 슬라이드 ----------------
class Builder:
    def __init__(self, template, roles, font, date):
        self.prs = Presentation(template)
        self.roles = roles
        self.font = font or roles.get('default_font')
        self.body_font = font or roles.get('body_font')
        self.date = date
        self.samples = list(self.prs.slides)
        self.notes = []

    def sample(self, role):
        return self.samples[self.roles[role]['slide']]

    def new(self, role):
        if role not in self.roles:
            raise SystemExit(f"이 서식에는 '{role}' 유형이 없습니다(샘플 자료에 해당 슬라이드가 없었음). spec에서 빼세요.")
        sl = duplicate_slide(self.prs, self.sample(role))
        sl._element.find(qn('p:cSld')).set('name', role)
        for ref in self.roles[role].get('clear', []):      # 샘플의 내용 개체(표, 그림, 본문 글상자)는 지운다
            sh = find_shape(sl, name=ref)
            if sh is not None:
                remove_shape(sh)
        return sl

    def put(self, sl, ref, text, lines=False):
        """roles의 도형 참조(이름 또는 #id)에 텍스트를 넣는다. 참조가 없거나 도형이 없으면 건너뛴다."""
        if not ref:
            return None
        sh = find_shape(sl, name=ref) or find_shape(sl, text=ref)
        if sh is None:
            return None
        (set_lines if lines else set_text)(sh, text, self.font)
        return sh

    def drop(self, sl, ref):
        sh = find_shape(sl, name=ref) if ref else None
        if sh is not None:
            remove_shape(sh)

    def cover(self, s):
        r = self.roles['cover']
        sl = self.new('cover')
        self.put(sl, r.get('title'), s['title'] if isinstance(s['title'], list) else [s['title']], lines=True)
        if s.get('subtitle'):
            self.put(sl, r.get('subtitle'), s['subtitle'])
        else:
            self.drop(sl, r.get('subtitle'))
        self.put(sl, r.get('date_org'), s.get('date_org', f"{s.get('date', self.date)} 서울과학기술대학교"))
        self.put(sl, r.get('date'), s.get('date', self.date))

    def agenda(self, s):
        r = self.roles['agenda']
        mx = r['max_items']
        items = s['items']
        chunks = [items[i:i + mx] for i in range(0, len(items), mx)] or [[]]
        for ci, chunk in enumerate(chunks):
            sl = self.new('agenda')
            title = s.get('title', '목차')
            if len(chunks) > 1:
                title = f'{title} ({ci + 1}/{len(chunks)})'
            self.put(sl, r.get('title'), title)
            # 먼저 모든 항목의 도형을 찾아 둔다(텍스트를 바꾸면 같은 텍스트 검색이 엇갈리므로)
            found = []
            for sample in r['items']:
                if str(sample.get('num', '')).startswith('#'):       # 자동 생성 roles: 도형 id로 지정
                    refs = [sample.get('num'), sample.get('title'), sample.get('org')] + list(sample.get('deco', []))
                    shp = [find_shape(sl, name=x) if x else None for x in refs]
                    found.append((shp[0], shp[1], shp[2], shp[3:]))
                    continue
                num_sh = find_shape(sl, text=sample['num'])
                title_sh = find_shape(sl, text=sample['title'])
                org_sh = find_shape(sl, text=sample['org'])
                circle = None
                for sh in sl.shapes:
                    if (not sh.has_text_frame or not sh.text_frame.text.strip()) and sh.shape_type == 1 \
                            and abs(sh.left - num_sh.left) < 60000 and abs(sh.top - num_sh.top) < 60000 \
                            and sh.width == num_sh.width:
                        circle = sh
                        break
                found.append((num_sh, title_sh, org_sh, [circle]))
            for k, (num_sh, title_sh, org_sh, decos) in enumerate(found):
                if k < len(chunk):
                    it = chunk[k]
                    if num_sh is not None:
                        set_text(num_sh, f'{ci * mx + k + 1:02d}', self.font)
                    if title_sh is not None:
                        set_text(title_sh, it['title'], self.font)
                    if org_sh is not None:
                        set_text(org_sh, it.get('org', ''), self.font)
                else:
                    for sh in [num_sh, title_sh, org_sh] + decos:
                        if sh is not None:
                            remove_shape(sh)

    def section(self, s, idx, total):
        r = self.roles['section']
        sl = self.new('section')
        self.put(sl, r.get('part'), s.get('part', f'PART {idx:02d} / {total:02d}'))
        self.put(sl, r.get('title'), s['title'])
        if s.get('subtitle'):
            self.put(sl, r.get('subtitle'), s['subtitle'])
        else:
            self.drop(sl, r.get('subtitle'))

    def content(self, s):
        r = self.roles['content']
        x, y, w, h = r['body_area']
        wp, hp = w / EMU_PT - 14.4, h / EMU_PT
        bullets = norm_bullets(s.get('bullets'))
        table = s.get('table')
        jobs = []  # (bullets, table_rows)
        if bullets:
            bpages = paginate_bullets(bullets, wp, hp)
            if table and len(bpages) == 1:
                used = sum(bullet_height_pt(t, l, wp) for t, l in bpages[0])
                if used + 40 > hp * 0.5:      # 표를 함께 놓기엔 위쪽 글이 많음
                    jobs.append((bpages[0], None))
                else:
                    jobs.append((bpages[0], 'with'))
            else:
                for pg in bpages:
                    jobs.append((pg, None))
        if table and not (jobs and jobs[-1][1] == 'with'):
            avail = hp
            for pg in paginate_table(table['header'], table['rows'], w / EMU_PT, avail):
                jobs.append((None, pg))
        total = len(jobs)
        for i, (bl, tb) in enumerate(jobs):
            sl = self.new('content')
            title = s['title'] + (f' ({i + 1}/{total})' if total > 1 else '')
            if len(title) > r.get('title_max_chars', 36):
                self.notes.append(f"제목이 길어 한 줄을 넘습니다({len(title)}자): '{title}' → 제목을 줄이세요")
            self.put(sl, r.get('title'), title)
            top_offset = 0
            if bl:
                used = sum(bullet_height_pt(t, l, wp) for t, l in bl)
                bh = int(min(h, (used + 12) * EMU_PT))
                add_bullets(sl, (x, y, w, bh), bl, self.body_font)
                top_offset = bh + 120000
            if tb == 'with':
                add_table(sl, (x, y, w, h), table['header'], table['rows'], self.body_font, top_offset)
            elif tb is not None:
                add_table(sl, (x, y, w, h), table['header'], tb, self.body_font)

    def closing(self, s):
        r = self.roles['closing']
        sl = self.new('closing')
        self.put(sl, r.get('text'), s.get('text', '감사합니다'))
        self.put(sl, r.get('date'), s.get('date', self.date))

    def finish(self, out):
        n = len(self.samples)
        if self.font:   # 폰트 지정 시 라벨(AGENDA, PART 등)까지 모두 같은 폰트로 통일
            from pptx_lib import iter_shapes
            for sl in list(self.prs.slides)[n:]:
                for sh in iter_shapes(sl.shapes):
                    if sh.has_text_frame:
                        apply_font(sh, self.font)
        for _ in range(n):
            drop_slide(self.prs, 0)
        self.prs.save(out)


# ---------------- 플레이스홀더 서식용 빌더 ----------------
class PHBuilder:
    """레이아웃에 플레이스홀더가 있는 서식: 레이아웃으로 슬라이드를 추가하고 idx로 채운다."""

    def __init__(self, template, roles, font, date):
        self.prs = Presentation(template)
        self.roles = roles
        self.font = font
        self.date = date
        self.samples = list(self.prs.slides)
        self.layouts = {l.name: l for m in self.prs.slide_masters for l in m.slide_layouts}
        self.notes = []

    # --- 공통 ---
    def sample_slide(self, role):
        i = self.roles['sample_slides'].get(role)
        return self.samples[i] if i is not None else None

    def need(self, role):
        if role not in self.roles.get('layouts', {}):
            raise SystemExit(f"이 서식에는 '{role}' 유형이 없습니다(샘플 자료에 해당 슬라이드가 없었음). "
                             f"spec에서 빼거나 type=custom으로 다른 레이아웃을 지정하세요.")

    def has_ph(self, slide, idx):
        return idx is not None and any(p.placeholder_format.idx == idx for p in slide.placeholders)

    def new(self, role, layout_name=None):
        if layout_name is None:
            self.need(role)
        name = layout_name or self.roles['layouts'][role]
        if name not in self.layouts:
            raise SystemExit(f"레이아웃 '{name}'이 서식에 없습니다.")
        return self.prs.slides.add_slide(self.layouts[name])

    def fill(self, slide, role, idx, lines):
        """플레이스홀더 idx를 lines로 채운다. run 서식은 서식의 샘플 슬라이드에서 가져온다."""
        if not self.has_ph(slide, idx):
            self.notes.append(f"레이아웃에 플레이스홀더 idx={idx}가 없어 '{lines if isinstance(lines, str) else lines[0]}'를 넣지 못했습니다")
            return
        ph = slide.placeholders[idx]
        sample = None
        sample_slide = self.sample_slide(role)
        for sh in (sample_slide.placeholders if sample_slide is not None else []):
            if sh.placeholder_format.idx == idx:
                sample = sh
        if isinstance(lines, str):
            lines = [lines]
        tf = ph.text_frame
        for p in list(tf.paragraphs)[1:]:
            p._p.getparent().remove(p._p)
        p0 = tf.paragraphs[0]
        for r in list(p0._p):
            if r.tag != qn('a:pPr'):
                p0._p.remove(r)
        s_paras = list(sample.text_frame.paragraphs) if sample is not None else []
        for i, line in enumerate(lines):
            level = 0
            if isinstance(line, (list, tuple)):
                line, level = line[0], int(line[1])
            p = p0 if i == 0 else tf.add_paragraph()
            if level:
                p.level = level
            tp = s_paras[min(i, len(s_paras) - 1)] if s_paras else None
            truns = tp.runs if tp is not None else []
            if len(truns) >= 2 and ' ' in line:
                first, rest = line.split(' ', 1)
                pieces = [(first, truns[0]), (' ' + rest, truns[1])]
            else:
                pieces = [(line, truns[0] if truns else None)]
            for text, tr in pieces:
                r = p.add_run()
                r.text = text
                if tr is not None:
                    src = tr._r.find(qn('a:rPr'))
                    if src is not None:
                        old = r._r.find(qn('a:rPr'))
                        if old is not None:
                            r._r.remove(old)
                        r._r.insert(0, copy.deepcopy(src))
        apply_font(tf, self.font)

    def drop_empty(self, slide):
        for sh in list(slide.placeholders):
            if sh.placeholder_format.type is not None and 'SLIDE_NUMBER' in str(sh.placeholder_format.type):
                continue
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                remove_shape(sh)

    # --- 역할별 ---
    def cover(self, s):
        r = {'title': 0, 'subtitle': 1, 'date_org': 10, 'date': 11, **self.roles.get('cover', {})}
        sl = self.new('cover')
        self.fill(sl, 'cover', r['title'], s['title'] if isinstance(s['title'], list) else [s['title']])
        if s.get('subtitle') and 'subtitle' in r:
            self.fill(sl, 'cover', r['subtitle'], s['subtitle'])
        if 'date_org' in r:
            self.fill(sl, 'cover', r['date_org'], s.get('date_org', f"{s.get('date', self.date)} 서울과학기술대학교"))
        if 'date' in r:
            self.fill(sl, 'cover', r['date'], s.get('date', self.date))
        for k, v in (s.get('fields') or {}).items():
            if k in r:
                self.fill(sl, 'cover', r[k], v)
        self.drop_empty(sl)

    def agenda(self, s):
        r = self.roles['agenda']
        mx = r['max_items']
        items = s['items']
        chunks = [items[i:i + mx] for i in range(0, len(items), mx)] or [[]]
        for ci, chunk in enumerate(chunks):
            sl = self.new('agenda', f"{self.roles['layouts']['agenda']} {max(1, len(chunk))}개")
            title = s.get('title', '목차')
            if len(chunks) > 1:
                title = f'{title} ({ci + 1}/{len(chunks)})'
            self.fill(sl, 'agenda', r['title'], title)
            for k, it in enumerate(chunk):
                row = r['items'][k]
                if 'title' in row:
                    self.fill(sl, 'agenda', row['title'], it['title'])
                if 'org' in row:
                    self.fill(sl, 'agenda', row['org'], it.get('org', '') or ' ')
            self.drop_empty(sl)

    def section(self, s, idx, total):
        self.need('section')
        r = self.roles['section']
        sl = self.new('section')
        if 'part' in r:
            self.fill(sl, 'section', r['part'], s.get('part', f'PART {idx:02d} / {total:02d}'))
        self.fill(sl, 'section', r['title'], s['title'])
        if s.get('subtitle') and 'subtitle' in r:
            self.fill(sl, 'section', r['subtitle'], s['subtitle'])
        self.drop_empty(sl)

    def _add_slide_number(self, sl, role='content'):
        sample = self.sample_slide(role)
        for sh in (sample.placeholders if sample is not None else []):
            if 'SLIDE_NUMBER' in str(sh.placeholder_format.type):
                el = copy.deepcopy(sh._element)
                ids = [int(e.get('id')) for e in sl.shapes._spTree.iter(qn('p:cNvPr')) if e.get('id', '').isdigit()]
                el.find('.//' + qn('p:cNvPr')).set('id', str(max(ids + [1]) + 1))
                sl.shapes._spTree.append(el)
                return

    def content(self, s):
        self.need('content')
        r = self.roles['content']
        x, y, w, h = r['body_area']
        wp, hp = w / EMU_PT - 14.4, h / EMU_PT
        bullets = norm_bullets(s.get('bullets'))
        table = s.get('table')
        jobs = []
        if bullets:
            bpages = paginate_bullets(bullets, wp, hp)
            if table and len(bpages) == 1:
                used = sum(bullet_height_pt(t, l, wp) for t, l in bpages[0])
                jobs.append((bpages[0], 'with' if used + 40 <= hp * 0.5 else None))
            else:
                for pg in bpages:
                    jobs.append((pg, None))
        if table and not (jobs and jobs[-1][1] == 'with'):
            for pg in paginate_table(table['header'], table['rows'], w / EMU_PT, hp):
                jobs.append((None, pg))
        total = len(jobs)
        for i, (bl, tb) in enumerate(jobs):
            sl = self.new('content')
            title = s['title'] + (f' ({i + 1}/{total})' if total > 1 else '')
            if len(title) > r.get('title_max_chars', 36):
                self.notes.append(f"제목이 길어 한 줄을 넘습니다({len(title)}자): '{title}' → 제목을 줄이세요")
            self.fill(sl, 'content', r['title'], title)
            body = sl.placeholders[r['body']]
            top_offset = 0
            if bl:
                tf = body.text_frame
                for k, (text, level) in enumerate(bl):
                    p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                    for el in list(p._p):
                        if el.tag != qn('a:pPr'):
                            p._p.remove(el)
                    p.level = level
                    run = p.add_run()
                    run.text = text
                apply_font(tf, self.font)
                used = sum(bullet_height_pt(t, l, wp) for t, l in bl)
                bh = int(min(h, (used + 12) * EMU_PT))
                body.left, body.top, body.width, body.height = Emu(x), Emu(y), Emu(w), Emu(bh)
                top_offset = bh + 120000
            else:
                remove_shape(body)
            tfont = self.font or r.get('body_font')
            if tb == 'with':
                add_table(sl, (x, y, w, h), table['header'], table['rows'], tfont, top_offset)
            elif tb is not None:
                add_table(sl, (x, y, w, h), table['header'], tb, tfont)
            self._add_slide_number(sl)

    def closing(self, s):
        self.need('closing')
        r = self.roles['closing']
        sl = self.new('closing')
        self.fill(sl, 'closing', r['text'], s.get('text', '감사합니다'))
        if 'date' in r:
            self.fill(sl, 'closing', r['date'], s.get('date', self.date))
        self.drop_empty(sl)

    def custom(self, s):
        """레이아웃 이름과 필드(이름→내용)를 직접 지정하는 슬라이드. 필드 이름은 roles.json의 layout_fields 참고."""
        lay = s['layout']
        fields = self.roles.get('layout_fields', {}).get(lay)
        if fields is None:
            raise SystemExit(f"레이아웃 '{lay}' 정보가 roles.json에 없습니다. 사용 가능: {sorted(self.roles.get('layout_fields', {}))}")
        sl = self.new(None, lay)
        for name, val in s.get('fields', {}).items():
            if name not in fields:
                raise SystemExit(f"레이아웃 '{lay}'에 필드 '{name}'이 없습니다. 사용 가능: {sorted(fields)}")
            self.fill(sl, lay, fields[name], val)
        self.drop_empty(sl)
        if 'sldNum' in fields:
            self._add_slide_number(sl, lay)

    def finish(self, out):
        for _ in range(len(self.samples)):
            drop_slide(self.prs, 0)
        self.prs.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--spec', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--template', default=None, help='서식 파일 경로 또는 등록된 서식 id(생략 시 기본 서식)')
    ap.add_argument('--font', default=None, help='지정 폰트(없으면 서식 폰트 유지)')
    a = ap.parse_args()

    with open(a.spec, encoding='utf8') as f:
        spec = json.load(f)
    a.template = resolve_template(a.template)
    roles = load_roles(a.template)
    font = a.font or spec.get('font')
    date = spec.get('date') or datetime.date.today().strftime('%Y.%m.%d.')
    b = (PHBuilder if roles.get('mode') == 'placeholder' else Builder)(a.template, roles, font, date)

    sections = [s for s in spec['slides'] if s['type'] == 'section']
    sec_i = 0
    for s in spec['slides']:
        t = s['type']
        if t == 'cover':
            b.cover(s)
        elif t == 'agenda':
            b.agenda(s)
        elif t == 'section':
            sec_i += 1
            b.section(s, sec_i, len(sections))
        elif t == 'content':
            b.content(s)
        elif t == 'closing':
            b.closing(s)
        elif t == 'custom' and hasattr(b, 'custom'):
            b.custom(s)
        else:
            raise SystemExit(f'알 수 없는 슬라이드 종류: {t}')
    b.finish(a.out)
    print(f'생성 완료: {a.out} (슬라이드 {len(b.prs.slides)}장, 폰트: {font or "서식 기본"})')
    for n in b.notes:
        print('참고:', n)


if __name__ == '__main__':
    main()
