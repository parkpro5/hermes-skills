#!/usr/bin/env python3
"""서식 또는 실제 사용했던 자료를 분석해 빌더가 쓸 수 있는 서식 + roles.json으로 자동 변환한다.

  --mode placeholder : 슬라이드 유형(표지/목차/구분/본문/마무리/2단 ...)마다 레이아웃과 플레이스홀더를 만든다.
                       같은 디자인이 여러 장이면 대표 1장만 샘플로 남긴다. 목차는 항목 수별(1~N개) 레이아웃도 만든다.
  --mode duplicate   : 슬라이드는 그대로 두고(대표 슬라이드만 남김) 도형 id로 역할을 적은 roles.json만 만든다.

사용:  python3 auto_convert.py 원본.pptx --out 결과.pptx --mode placeholder
       → 결과.pptx, 결과.roles.json 생성 (같은 폴더에 두면 build_deck.py가 자동으로 찾는다)
"""
import argparse
import copy
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from lxml import etree  # noqa: E402

import tpl_analysis as T  # noqa: E402
from pptx_lib import embed_roles  # noqa: E402
from make_placeholder_template import (CT_LAYOUT, NS, R, T_LAYOUT, lvl1_from_paragraph, local,  # noqa: E402
                                       parse, q, read_rels, to_placeholder_layout,
                                       to_placeholder_slide, write, write_rels)

CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
SKIP_PH = T.SKIP_PH
PROMPTS = {'title': '제목', 'subtitle': '부제', 'body': '본문 내용을 입력하세요', 'body2': '본문 2',
           'date': '날짜', 'date_org': '날짜와 소속', 'part': 'PART 번호', 'text': '마무리 문구'}
EMU_PT = 12700


# ---------------- 계획 만들기 ----------------
def layout_names(a):
    names, count = {}, defaultdict(int)
    base = {'cover': '표지', 'agenda': '목차', 'section': '구분', 'content': '본문', 'closing': '마무리'}
    for ci, c in enumerate(a['clusters']):
        t = c['type']
        nm = base[t]
        if t in ('content', 'section') and c['key'][1] == 'two-col':
            nm += ' 2단'
        count[nm] += 1
        if count[nm] > 1 and t != 'agenda':
            nm = f'{nm} {count[nm]}'
        names[ci] = nm
    return names


def body_geometry(a, c):
    """본문 영역(x, y, w, h)과 금지 영역(no_go)을 대표 슬라이드에서 계산한다."""
    W, H = a['W'], a['H']
    rep = a['slides'][c['rep']]
    title = next((f['s'] for f in c['fields'] if f['name'] == 'title'), None)
    body = next((f['s'] for f in c['fields'] if f['name'] == 'body'), None)
    if body:
        x, y, w = body['x'], body['y'], body['w']
        bottom = 0.93 * H
    else:
        x = title['x'] if title else int(0.045 * W)
        w = W - 2 * x
        y = (title['y'] + title['h'] + int(0.02 * H)) if title else int(0.2 * H)
        bottom = 0.93 * H
    # 아래쪽에서 가로로 겹치는 고정 도형이 있으면 그 위까지만 쓴다
    for s in rep['shapes']:
        if s['id'] in c['decor_ids'] and s['y'] > y + 0.2 * H and s['x'] < x + w and s['x'] + s['w'] > x \
                and s['w'] * s['h'] < 0.3 * W * H:
            bottom = min(bottom, s['y'] - 0.01 * H)
    h = max(int(0.2 * H), int(bottom - y))
    area = [int(x), int(y), int(w), int(h)]
    no_go = []
    if title:
        no_go.append([title['x'], title['y'], title['w'], title['h']])
    for s in rep['shapes']:
        if s['id'] in c['decor_ids'] and s['kind'] in ('picture', 'group') and s['w'] * s['h'] < 0.3 * W * H:
            if not T_rect_intersects((s['x'], s['y'], s['w'], s['h']), area):
                no_go.append([s['x'], s['y'], s['w'], s['h']])
        if c['decisions'].get(s['id'], {}).get('role') == 'sldNum':
            no_go.append([s['x'], s['y'], s['w'], s['h']])
    return area, no_go


def T_rect_intersects(a, b):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return not (ax + aw <= bx or bx + bw <= ax or ay + ah <= by or by + bh <= ay)


def make_specs(a):
    names = layout_names(a)
    specs = []
    for ci, c in enumerate(a['clusters']):
        t = c['type']
        rep = a['slides'][c['rep']]
        sid_shape = {s['id']: s for s in rep['shapes']}
        taken = {s['ph_idx'] for s in rep['shapes'] if s['ph']}
        ph, keep, fields_idx = {}, set(), {}

        def alloc(pref=None):
            if pref is not None and pref not in taken:
                taken.add(pref)
                return pref
            n = 10
            while n in taken:
                n += 1
            taken.add(n)
            return n

        for s in rep['shapes']:           # 이미 플레이스홀더인 도형은 그대로 둔다
            if s['ph']:
                keep.add(s['id'])
        for f in c['fields']:
            s, name = f['s'], f['name']
            if s['ph']:
                fields_idx[name] = s['ph_idx']
                continue
            if name == 'title':
                ptype, idx = ('ctrTitle' if t in ('cover', 'closing') else 'title'), alloc(0)
            elif name == 'text' and t == 'closing':
                ptype, idx = 'ctrTitle', alloc(0)
            elif name == 'subtitle' and t == 'cover':
                ptype, idx = 'subTitle', alloc(1)
            elif name == 'body':
                ptype, idx = 'body', alloc(1)
            elif name == 'body2':
                ptype, idx = 'body', alloc(2)
            else:
                ptype, idx = 'body', alloc(None)
            ph[s['id']] = (ptype, idx, PROMPTS.get(name, name))
            fields_idx[name] = idx
        # 페이지 번호
        for sid, d in c['decisions'].items():
            if d.get('role') == 'sldNum' and sid not in keep and sid in sid_shape:
                idx = alloc(12)
                ph[sid] = ('sldNum', idx, '‹#›')
                fields_idx['sldNum'] = idx
        # 목차 항목
        variants, items_idx = {}, []
        if t == 'agenda' and c['agenda']:
            ag = c['agenda']
            for k, it in enumerate(ag['items']):
                row = {}
                if it['title']:
                    row['title'] = alloc(10 + 2 * k)
                    ph[it['title']['id']] = ('body', row['title'], f'항목 {k + 1} 제목')
                if it['org']:
                    row['org'] = alloc(11 + 2 * k)
                    ph[it['org']['id']] = ('body', row['org'], f'항목 {k + 1} 담당')
                items_idx.append(row)
            n_items = len(ag['items'])
            for n in range(n_items - 1, 0, -1):
                drop = []
                for it in ag['items'][n:]:
                    drop += [x['id'] for x in (it['num'], it['title'], it['org']) if x] + [d['id'] for d in it['deco']]
                variants[n] = drop
        # 슬라이드에 남는 개체(표/차트/반복되지 않는 그림): 레이아웃으로 옮기지 않는다
        content_ids = {s['id'] for s in rep['shapes']
                       if s['id'] not in ph and s['id'] not in keep and not s['text']
                       and s['id'] not in set(c['decor_ids'])}
        name = f'목차 {len(c["agenda"]["items"])}개' if (t == 'agenda' and c['agenda']) else names[ci]
        spec = dict(type=t, cluster=ci, name=name, slide=c['rep'], ph=ph, keep=keep, content_ids=content_ids,
                    variants=variants, fields=fields_idx, items=items_idx)
        if t == 'content':
            area, no_go = body_geometry(a, c)
            spec['area'], spec['no_go'] = area, no_go
            spec['body_default'] = 'body' not in fields_idx
            if spec['body_default']:
                taken.add(1)
                spec['fields']['body'] = 1
        specs.append(spec)
    return specs


# ---------------- 패키지 변환 ----------------
def default_body_layout(area, shape_id, accent, font):
    x, y, w, h = area
    fnt = (f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/><a:cs typeface="{font}"/>' if font else '')
    lv = lambda n, marL, ch, sz, bef: (  # noqa: E731
        f'<a:lvl{n}pPr marL="{marL}" indent="-285750" algn="l"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>'
        f'<a:spcBef><a:spcPts val="{bef}"/></a:spcBef><a:buClr><a:srgbClr val="{accent}"/></a:buClr>'
        f'<a:buFont typeface="Arial"/><a:buChar char="{ch}"/><a:defRPr sz="{sz}" b="0">'
        f'<a:solidFill><a:srgbClr val="262626"/></a:solidFill>{fnt}</a:defRPr></a:lvl{n}pPr>')
    xml = f'''<p:sp xmlns:a="{NS['a']}" xmlns:p="{NS['p']}">
<p:nvSpPr><p:cNvPr id="{shape_id}" name="본문 개체 틀"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:normAutofit/></a:bodyPr>
<a:lstStyle>{lv(1, 285750, '&#8226;', 2000, 900)}{lv(2, 571500, '&#8211;', 1800, 600)}{lv(3, 857250, '&#8211;', 1600, 400)}</a:lstStyle>
<a:p><a:r><a:rPr lang="ko-KR"/><a:t>본문 내용을 입력하세요</a:t></a:r></a:p>
<a:p><a:pPr lvl="1"/><a:r><a:rPr lang="ko-KR"/><a:t>하위 항목</a:t></a:r></a:p>
</p:txBody></p:sp>'''
    return etree.fromstring(xml)


def default_body_slide(shape_id):
    xml = f'''<p:sp xmlns:a="{NS['a']}" xmlns:p="{NS['p']}">
<p:nvSpPr><p:cNvPr id="{shape_id}" name="본문 개체 틀"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:r><a:rPr lang="ko-KR"/><a:t>본문 내용을 입력하세요</a:t></a:r></a:p>
<a:p><a:pPr lvl="1"/><a:r><a:rPr lang="ko-KR"/><a:t>하위 항목은 Tab 키로 들여쓰기</a:t></a:r></a:p>
</p:txBody></p:sp>'''
    return etree.fromstring(xml)


def multi_level_lst(sp_orig):
    """원본 도형의 문단들(들여쓰기가 다른 것)로 lvl1~lvl3 서식을 만든다. 없으면 None."""
    tx = sp_orig.find(q('p:txBody'))
    seen = {}
    for p in tx.findall(q('a:p')):
        pPr = p.find(q('a:pPr'))
        key = (int(pPr.get('lvl', '0')) if pPr is not None else 0,
               int(pPr.get('marL', '0')) if pPr is not None else 0)
        seen.setdefault(key, p)
    if len(seen) < 2:
        return None
    lst = etree.Element(q('a:lstStyle'))
    for n, key in enumerate(sorted(seen)[:3], 1):
        one = lvl1_from_paragraph(seen[key])
        lv = one[0]
        lv.tag = q(f'a:lvl{n}pPr')
        if 'lvl' in lv.attrib:
            del lv.attrib['lvl']
        lst.append(lv)
    return lst


def ensure_levels(lay_sp, n_levels=3):
    """본문 플레이스홀더에 2~3단계 글머리표 서식이 없으면 1단계에서 만들어 준다(마스터의 큰 기본값 방지)."""
    tx = lay_sp.find(q('p:txBody'))
    lst = tx.find(q('a:lstStyle'))
    lv1 = lst.find(q('a:lvl1pPr'))
    if lv1 is None:
        return
    for n in range(2, n_levels + 1):
        if lst.find(q(f'a:lvl{n}pPr')) is not None:
            continue
        prev = lst.find(q(f'a:lvl{n - 1}pPr'))
        new = copy.deepcopy(prev)
        new.tag = q(f'a:lvl{n}pPr')
        new.set('marL', str(int(prev.get('marL', '0')) + 285750))
        d = new.find(q('a:defRPr'))
        if d is not None and d.get('sz'):
            d.set('sz', str(max(1200, int(d.get('sz')) - 200)))
        bu = new.find(q('a:buChar'))
        if bu is not None:
            bu.set('char', '–')
        lst.append(new)
    bp = tx.find(q('a:bodyPr'))
    bp.set('wrap', 'square')
    for ch in list(bp):
        bp.remove(ch)
    etree.SubElement(bp, q('a:normAutofit'))


class Converter:
    def __init__(self, src):
        self.work = tempfile.mkdtemp()
        with zipfile.ZipFile(src) as z:
            z.extractall(self.work)
            self.order = [i.filename for i in z.infolist()]
        self.ppt = os.path.join(self.work, 'ppt')
        self.ct_path = os.path.join(self.work, '[Content_Types].xml')
        self.ct = parse(self.ct_path)
        pres = parse(os.path.join(self.ppt, 'presentation.xml'))
        prels = {r['Id']: r['Target'] for r in read_rels(os.path.join(self.ppt, '_rels', 'presentation.xml.rels'))}
        self.slide_files = [prels[e.get(R + 'id')] for e in pres.getroot().find(q('p:sldIdLst'))]
        self.masters = {}
        mdir = os.path.join(self.ppt, 'slideMasters')
        max_id = 2147483647
        for e in pres.getroot().find(q('p:sldMasterIdLst')):
            max_id = max(max_id, int(e.get('id')))
        for f in sorted(os.listdir(mdir)):
            if f.endswith('.xml'):
                mp = os.path.join(mdir, f)
                tree = parse(mp)
                rp = os.path.join(mdir, '_rels', f + '.rels')
                self.masters[f] = dict(path=mp, tree=tree, rels_path=rp, rels=read_rels(rp))
                for e in tree.getroot().find(q('p:sldLayoutIdLst')):
                    max_id = max(max_id, int(e.get('id')))
        self.next_lid = max_id + 1
        nums = [int(re.findall(r'\d+', f)[0]) for f in os.listdir(os.path.join(self.ppt, 'slideLayouts')) if f.endswith('.xml')]
        self.next_layout = max(nums) + 1
        self.layout_files = {}      # name -> file

    def save_layout(self, tree, name, rels, master_file):
        lname = f'slideLayout{self.next_layout}.xml'
        self.next_layout += 1
        tree.getroot().find(q('p:cSld')).set('name', name)
        write(tree, os.path.join(self.ppt, 'slideLayouts', lname))
        write_rels(os.path.join(self.ppt, 'slideLayouts', '_rels', lname + '.rels'), rels)
        etree.SubElement(self.ct.getroot(), f'{{{CT_NS}}}Override', PartName=f'/ppt/slideLayouts/{lname}', ContentType=CT_LAYOUT)
        m = self.masters[master_file]
        used = {r['Id'] for r in m['rels']}
        n = 1
        while f'rId{n}' in used:
            n += 1
        rid = f'rId{n}'
        m['rels'].append(dict(Id=rid, Type=T_LAYOUT, Target=f'../slideLayouts/{lname}'))
        el = etree.SubElement(m['tree'].getroot().find(q('p:sldLayoutIdLst')), q('p:sldLayoutId'))
        el.set('id', str(self.next_lid))
        self.next_lid += 1
        el.set(R + 'id', rid)
        self.layout_files[name] = lname
        return lname

    def convert(self, spec, accent, body_font):
        sfile = self.slide_files[spec['slide']]
        slide_path = os.path.join(self.ppt, sfile)
        rels_path = os.path.join(os.path.dirname(slide_path), '_rels', os.path.basename(slide_path) + '.rels')
        stree, srels = parse(slide_path), read_rels(rels_path)
        sp_tree = stree.getroot().find(q('p:cSld')).find(q('p:spTree'))
        base_rel = next(r for r in srels if r['Type'].endswith('/slideLayout'))
        base_file = os.path.basename(base_rel['Target'])
        base_path = os.path.join(self.ppt, 'slideLayouts', base_file)
        lrels = read_rels(os.path.join(self.ppt, 'slideLayouts', '_rels', base_file + '.rels'))
        master_file = os.path.basename(next(r for r in lrels if r['Type'].endswith('/slideMaster'))['Target'])
        ltree = parse(base_path)
        lroot = ltree.getroot()
        if 'type' in lroot.attrib:
            del lroot.attrib['type']
        lroot.set('preserve', '1')
        lroot.set('userDrawn', '1')
        l_sptree = lroot.find(q('p:cSld')).find(q('p:spTree'))
        # 기본 레이아웃의 플레이스홀더는 이 슬라이드가 쓰는 것(keep)과 번호·날짜·바닥글만 남긴다
        keep_idx = {int(sp.find('.//' + q('p:ph')).get('idx', '0')) for sp in sp_tree.findall(q('p:sp'))
                    if sp.find('.//' + q('p:ph')) is not None}
        for sp in l_sptree.findall(q('p:sp')):
            ph = sp.find('.//' + q('p:ph'))
            if ph is not None and ph.get('type') not in ('sldNum', 'dt', 'ftr') and int(ph.get('idx', '0')) not in keep_idx:
                l_sptree.remove(sp)
        # 슬라이드 자체 배경은 레이아웃으로 옮긴다
        bg = stree.getroot().find(q('p:cSld')).find(q('p:bg'))
        if bg is not None:
            old = lroot.find(q('p:cSld')).find(q('p:bg'))
            if old is not None:
                lroot.find(q('p:cSld')).remove(old)
            lroot.find(q('p:cSld')).insert(0, copy.deepcopy(bg))
            bg.getparent().remove(bg)
        # 관계(그림 등) 대응표
        used = {r['Id'] for r in lrels}
        rmap, k = {}, 0
        for r in srels:
            if r['Type'].endswith('/slideLayout') or r['Type'].endswith('/notesSlide'):
                continue
            hit = next((x for x in lrels if x['Target'] == r['Target'] and x['Type'] == r['Type']), None)
            if hit:
                rmap[r['Id']] = hit['Id']
                continue
            while f'rId{100 + k}' in used:
                k += 1
            nid = f'rId{100 + k}'
            used.add(nid)
            lrels.append(dict(Id=nid, Type=r['Type'], Target=r['Target'], Mode=r.get('Mode')))
            rmap[r['Id']] = nid

        next_id = max([int(e.get('id')) for e in l_sptree.iter(q('p:cNvPr')) if e.get('id', '').isdigit()] + [1]) + 1
        slide_children, id_map = [], {}
        for child in list(sp_tree):
            if local(child) in ('nvGrpSpPr', 'grpSpPr'):
                continue
            cid_el = child.find('.//' + q('p:cNvPr'))
            sid = int(cid_el.get('id')) if cid_el is not None else None
            if sid in spec['ph'] and local(child) == 'sp':
                ptype, idx, prompt = spec['ph'][sid]
                lay_sp = to_placeholder_layout(child, ptype, idx, prompt)
                ml = multi_level_lst(child) if ptype == 'body' and idx == spec['fields'].get('body') else None
                if ml is not None:
                    lay_sp.find(q('p:txBody')).replace(lay_sp.find(q('p:txBody')).find(q('a:lstStyle')), ml)
                if ptype == 'body' and idx == spec['fields'].get('body'):
                    ensure_levels(lay_sp)
                slide_sp = to_placeholder_slide(child, ptype, idx)
                lay_sp.find('.//' + q('p:cNvPr')).set('id', str(next_id))
                id_map[sid] = next_id
                next_id += 1
                l_sptree.append(lay_sp)
                slide_sp.find('.//' + q('p:cNvPr')).set('id', str(len(slide_children) + 2))
                slide_children.append(slide_sp)
            elif sid in spec['keep'] or sid in spec['content_ids']:
                slide_children.append(copy.deepcopy(child))
            else:
                c = copy.deepcopy(child)
                for node in c.iter():
                    for ak, av in list(node.attrib.items()):
                        if ak.startswith(R) and av in rmap:
                            node.set(ak, rmap[av])
                first = True
                for e in c.iter(q('p:cNvPr')):
                    e.set('id', str(next_id))
                    if first and sid is not None:
                        id_map[sid] = next_id
                    first = False
                    next_id += 1
                l_sptree.append(c)
        if spec.get('body_default'):
            l_sptree.append(default_body_layout(spec['area'], next_id, accent, body_font))
            next_id += 1
            slide_children.append(default_body_slide(len(slide_children) + 2))

        self.save_layout(ltree, spec['name'], lrels, master_file)
        for n, drop in spec['variants'].items():
            vt = copy.deepcopy(ltree)
            vtree = vt.getroot().find(q('p:cSld')).find(q('p:spTree'))
            drop_new = {str(id_map[x]) for x in drop if x in id_map}
            for ch in list(vtree):
                e = ch.find('.//' + q('p:cNvPr'))
                if e is not None and e.get('id') in drop_new:
                    vtree.remove(ch)
            self.save_layout(vt, f'목차 {n}개', lrels, master_file)

        # 샘플 슬라이드 다시 쓰기
        for ch in list(sp_tree):
            if local(ch) not in ('nvGrpSpPr', 'grpSpPr'):
                sp_tree.remove(ch)
        refd = set()
        for c in slide_children:
            sp_tree.append(c)
            for node in c.iter():
                for ak, av in node.attrib.items():
                    if ak.startswith(R):
                        refd.add(av)
        write(stree, slide_path)
        new_rels = []
        for r in srels:
            if r['Type'].endswith('/slideLayout'):
                r['Target'] = f'../slideLayouts/{self.layout_files[spec["name"]]}'
                new_rels.append(r)
            elif r['Type'].endswith('/notesSlide') or r['Id'] in refd:
                new_rels.append(r)
        write_rels(rels_path, new_rels)

    def finish(self, dst):
        write(self.ct, self.ct_path)
        for m in self.masters.values():
            write(m['tree'], m['path'])
            write_rels(m['rels_path'], m['rels'])
        files = []
        for root, _, fs in os.walk(self.work):
            for f in fs:
                files.append(os.path.relpath(os.path.join(root, f), self.work).replace(os.sep, '/'))
        files.sort(key=lambda f: (f != '[Content_Types].xml', self.order.index(f) if f in self.order else 10 ** 6, f))
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
            for f in files:
                z.write(os.path.join(self.work, f), f)
        shutil.rmtree(self.work)


# ---------------- 마무리: 슬라이드 정리 + roles.json ----------------
def keep_only(path, keep_idx_in_order):
    """대표 슬라이드만 남기고 지정한 순서로 정렬한다. 반환: 원래 인덱스 → 새 인덱스"""
    from pptx import Presentation
    from pptx.oxml.ns import qn
    prs = Presentation(path)
    lst = prs.slides._sldIdLst
    ids = list(lst)
    keep = [ids[i] for i in keep_idx_in_order]
    for e in ids:
        if e not in keep:
            prs.part.drop_rel(e.get(qn('r:id')))
        lst.remove(e)
    for e in keep:
        lst.append(e)
    prs.save(path)
    return {old: new for new, old in enumerate(keep_idx_in_order)}


def accent_of(a):
    """본문 글머리표 강조색: 대표 슬라이드에서 가장 눈에 띄는 단색(제목 외 도형 채움)을 못 찾으면 서식 파란색 계열 대신 회색."""
    from pptx import Presentation
    return '404040'


def sample_strings(a):
    out = []
    for c in a['clusters']:
        rep = a['slides'][c['rep']]
        for f in c['fields']:
            if c['type'] == 'closing' and f['name'] == 'text':
                continue
            for line in re.split(r'[\n\v]+', f['s']['text']):
                line = line.strip()
                if 2 <= len(line) <= 40 and not re.fullmatch(r'[\d\s.]+', line) and not T.DATE_RE.search(line) \
                        and line not in ('목차',) and not T.PART_RE.match(line):
                    out.append(line)
        if c['agenda']:
            for it in c['agenda']['items']:
                for k in ('title', 'org'):
                    if it[k]:
                        out.append(it[k]['text'].strip())
    out += ['필요 시']
    seen, res = set(), []
    for s in out:
        if s not in seen:
            seen.add(s)
            res.append(s)
    return res


def build_roles_placeholder(a, specs, order_map, name):
    names = {sp['cluster']: sp['name'] for sp in specs}
    first = {}
    for sp in specs:
        first.setdefault(sp['type'], sp)
    roles = dict(name=name, mode='placeholder', slide_size=[a['W'], a['H']], default_font=None,
                 layouts={}, sample_slides={}, layout_fields={})
    for t, sp in first.items():
        base = {'agenda': '목차'}.get(t, sp['name'])
        roles['layouts'][t] = base
        roles['sample_slides'][t] = order_map[sp['slide']]
    for sp in specs:
        roles['sample_slides'][sp['name']] = order_map[sp['slide']]
        roles['layout_fields'][sp['name']] = {k: v for k, v in sp['fields'].items()}
        if sp['type'] == 'agenda':
            for n in sp['variants']:
                roles['layout_fields'][f'목차 {n}개'] = {'title': sp['fields'].get('title', 0)}
    for t, sp in first.items():
        f = sp['fields']
        if t == 'cover':
            roles['cover'] = {k: f[k] for k in ('title', 'subtitle', 'date_org', 'date') if k in f}
        elif t == 'agenda':
            roles['agenda'] = dict(title=f.get('title', 0), max_items=len(sp['items']),
                                   items=[{k: v for k, v in row.items()} for row in sp['items']])
        elif t == 'section':
            roles['section'] = {k: f[k] for k in ('part', 'title', 'subtitle') if k in f}
        elif t == 'closing':
            roles['closing'] = {k: f[k] for k in ('text', 'date') if k in f}
        elif t == 'content':
            c = a['clusters'][sp['cluster']]
            title = next((x['s'] for x in c['fields'] if x['name'] == 'title'), None)
            size = (title['size'] if title and title['size'] else 28)
            maxc = int(max(16, min(60, (title['w'] / EMU_PT) / size))) if title else 36
            body = next((x['s'] for x in c['fields'] if x['name'] == 'body'), None)
            roles['content'] = dict(title=f.get('title', 0), body=f.get('body', 1),
                                    body_font=(body['font'] if body and body['font'] else (title['font'] if title else None)),
                                    title_max_chars=maxc, body_area=sp['area'], no_go=sp['no_go'])
    roles['sample_strings'] = sample_strings(a)
    return roles


def build_roles_duplicate(a, clusters_kept, order_map, name):
    roles = dict(name=name, mode='duplicate', slide_size=[a['W'], a['H']], default_font=None)
    first = {}
    for c in clusters_kept:
        first.setdefault(c['type'], c)
    ref = lambda s: f'#{s["id"]}' if s else None  # noqa: E731
    for t, c in first.items():
        d = {f['name']: ref(f['s']) for f in c['fields']}
        slide = order_map[c['rep']]
        rep = a['slides'][c['rep']]
        clear = [ref(s) for s in rep['shapes'] if not s['text'] and s['id'] not in set(c['decor_ids'])
                 and s['kind'] in ('table', 'chart', 'frame', 'picture', 'group')]
        if t == 'cover':
            roles['cover'] = dict(slide=slide, **d)
        elif t == 'agenda':
            ag = c['agenda']
            items = [dict(num=ref(it['num']), title=ref(it['title']), org=ref(it['org']),
                          deco=[ref(x) for x in it['deco']]) for it in ag['items']]
            roles['agenda'] = dict(slide=slide, title=d.get('title'), max_items=len(items), items=items)
        elif t == 'section':
            roles['section'] = dict(slide=slide, **d)
        elif t == 'closing':
            roles['closing'] = dict(slide=slide, **d)
        elif t == 'content':
            sp_area, no_go = body_geometry(a, c)
            title = next((x['s'] for x in c['fields'] if x['name'] == 'title'), None)
            size = (title['size'] if title and title['size'] else 28)
            maxc = int(max(16, min(60, (title['w'] / EMU_PT) / size))) if title else 36
            body = next((x['s'] for x in c['fields'] if x['name'] == 'body'), None)
            roles['content'] = dict(slide=slide, title=d.get('title'), title_max_chars=maxc, body_area=sp_area, no_go=no_go,
                                    body_font=(body['font'] if body and body['font'] else (title['font'] if title else None)))
            clear += [ref(body)] if body else []
        if t != 'agenda' and t in roles:
            roles[t]['clear'] = sorted({x for x in clear if x})
        # 대표 슬라이드에서 변수 도형 중 매핑되지 않은 것은 지운다(추가 필드)
    roles['sample_strings'] = sample_strings(a)
    return roles


def demo_spec(roles):
    """서식의 모든 유형을 한 장씩 보여 주는 미리보기용 spec."""
    L = roles.get('layouts', {})
    sl = []
    if 'cover' in L:
        sl.append(dict(type='cover', title=['서식 미리보기', '제목 두 줄 예시'], subtitle='부제 문구 예시',
                       date_org='2026.01.01. 소속 이름', date='2026.01.01.'))
    if 'agenda' in L:
        n = min(3, max(1, roles['agenda']['max_items']))
        sl.append(dict(type='agenda', title='목차 예시',
                       items=[dict(title=f'첫 번째 항목'.replace('첫', ['첫', '두', '세', '네', '다섯'][i]), org=f'담당 {i + 1}') for i in range(n)]))
    if 'section' in L:
        sl.append(dict(type='section', title='구분 제목 예시', subtitle='구분 설명 예시'))
    if 'content' in L:
        sl.append(dict(type='content', title='본문 슬라이드 예시',
                       bullets=['첫 번째 핵심 문장', ['하위 항목 예시', 1], '두 번째 핵심 문장', '세 번째 핵심 문장']))
        sl.append(dict(type='content', title='표 슬라이드 예시',
                       table=dict(header=['구분', '내용', '비고'], rows=[['A', '예시 내용', '-'], ['B', '예시 내용', '-']])))
    standard = set(L.values()) | {f'{L.get("agenda", "목차")} {k}개' for k in range(1, 10)}
    for name, fields in roles.get('layout_fields', {}).items():
        if name in standard:
            continue
        f = {}
        for k in fields:
            if k == 'sldNum':
                continue
            f[k] = (f'{name} 예시' if k == 'title' else
                    ['예시 문장', ['하위 항목', 1]] if k.startswith('body') else '예시')
        sl.append(dict(type='custom', layout=name, fields=f))
    if 'closing' in L:
        sl.append(dict(type='closing', text='감사합니다', date='2026.01.01.'))
    return dict(date='2026.01.01.', slides=sl)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src')
    ap.add_argument('--out', required=True)
    ap.add_argument('--mode', choices=['placeholder', 'duplicate'], required=True)
    ap.add_argument('--body-font', default=None, help='본문 글꼴(미지정 시 원본에서 추정)')
    ap.add_argument('--demo', action='store_true', help='변환한 서식의 미리보기 PPT(<out>_demo.pptx)도 만들고 검증한다')
    ap.add_argument('--accent', default=None, help='글머리표 강조색 RGB 16진수(예: 00387F)')
    a_ = ap.parse_args()

    a = T.analyze(a_.src)
    name = os.path.splitext(os.path.basename(a_.out))[0]
    kept_slides = [c['rep'] for c in a['clusters']]
    order_map = {old: new for new, old in enumerate(kept_slides)}
    if a_.mode == 'placeholder':
        specs = make_specs(a)
        conv = Converter(a_.src)
        content_c = next((c for c in a['clusters'] if c['type'] == 'content'), None)
        body_font = a_.body_font
        if not body_font and content_c:
            b = next((f['s'] for f in content_c['fields'] if f['name'] == 'body'), None)
            body_font = b['font'] if b else None
        accent = a_.accent or '00387F'
        for sp in specs:
            conv.convert(sp, accent, body_font)
        conv.finish(a_.out)
        keep_only(a_.out, kept_slides)
        roles = build_roles_placeholder(a, specs, order_map, name)
    else:
        shutil.copyfile(a_.src, a_.out)
        keep_only(a_.out, kept_slides)
        roles = build_roles_duplicate(a, a['clusters'], order_map, name)
    rp = os.path.splitext(a_.out)[0] + '.roles.json'
    with open(rp, 'w', encoding='utf8') as f:
        json.dump(roles, f, ensure_ascii=False, indent=2)
    embed_roles(a_.out, roles)         # 파일 하나로 공유할 수 있게 roles를 pptx 안에도 넣는다
    print(f'변환 완료({a_.mode}): {a_.out}\nroles: {rp} (pptx 안에도 내장됨 → pptx 한 파일만 공유해도 됨)')
    print('샘플 슬라이드:', ', '.join(f'{i + 1}={c["type"]}' for i, c in enumerate(a['clusters'])))
    if a_.demo:
        import subprocess
        base = os.path.splitext(a_.out)[0]
        spec_p, demo_p = base + '_demo.spec.json', base + '_demo.pptx'
        with open(spec_p, 'w', encoding='utf8') as f:
            json.dump(demo_spec(roles), f, ensure_ascii=False, indent=2)
        r = subprocess.run([sys.executable, '-I', os.path.join(HERE, 'build_deck.py'), '--spec', spec_p, '--out', demo_p,
                            '--template', a_.out], capture_output=True, text=True)
        print(r.stdout.strip() or r.stderr.strip())
        if r.returncode == 0:
            v = subprocess.run([sys.executable, '-I', os.path.join(HERE, 'verify_deck.py'), demo_p, '--spec', spec_p,
                                '--template', a_.out, '--outdir', base + '_demo_check'], capture_output=True, text=True)
            print('미리보기 검증:', (v.stdout.strip().splitlines() or ['(결과 없음)'])[-1])
            print(f'미리보기 이미지: {base}_demo_check/contact_sheet.png')


if __name__ == '__main__':
    main()
