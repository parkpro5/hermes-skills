#!/usr/bin/env python3
"""도형 기반 서식(플레이스홀더 없음) → 레이아웃 + 플레이스홀더 서식으로 변환한다.

각 샘플 슬라이드의 디자인(그림, 선, 장식 도형)은 새 슬라이드 레이아웃으로 옮기고,
글자가 들어가는 도형은 플레이스홀더로 바꾼다. 샘플 슬라이드는 새 레이아웃을 쓰는
플레이스홀더 슬라이드로 다시 만든다(내용과 모양은 원본과 같다).

사용:  python3 make_placeholder_template.py 원본.pptx 결과.pptx
※ 도형 이름 매핑(SPEC)은 번들 서식(seoultech_ax) 기준이다. 다른 서식은 SPEC을 고쳐서 쓴다.
"""
import copy
import os
import re
import shutil
import sys
import tempfile
import zipfile

from lxml import etree

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
A, P, R = ('{%s}' % NS[k] for k in ('a', 'p', 'r'))
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
T_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
CT_LAYOUT = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'

# 슬라이드 번호(1부터) → 새 레이아웃 정의
# ph: shape_id → (ph type, idx, 안내 문구)   static: 레이아웃에 그대로 남는 도형
SPEC = {
    1: dict(name='표지', base='slideLayout1.xml', master=1, ph={
        4: ('ctrTitle', 0, '제목'), 13: ('subTitle', 1, '부제'),
        10: ('body', 10, '날짜와 소속'), 8: ('body', 11, '날짜')}),
    2: dict(name='목차 5개', base='slideLayout4.xml', master=2, strip_base_ph=True,
            # 항목 수별 변형 레이아웃: n개 → 지울 항목(번호 원, 번호, 제목, 담당 도형 id)
            variants={4: [38, 36, 31, 35], 3: [38, 36, 31, 35, 44, 47, 48, 49],
                      2: [38, 36, 31, 35, 44, 47, 48, 49, 25, 27, 45, 46],
                      1: [38, 36, 31, 35, 44, 47, 48, 49, 25, 27, 45, 46, 20, 22, 37, 39]},
            ph={
        15: ('title', 0, '목차 제목'),
        32: ('body', 10, '항목 1 제목'), 33: ('body', 11, '항목 1 담당'),
        37: ('body', 12, '항목 2 제목'), 39: ('body', 13, '항목 2 담당'),
        45: ('body', 14, '항목 3 제목'), 46: ('body', 15, '항목 3 담당'),
        48: ('body', 16, '항목 4 제목'), 49: ('body', 17, '항목 4 담당'),
        31: ('body', 18, '항목 5 제목'), 35: ('body', 19, '항목 5 담당')}),
    3: dict(name='구분', base='slideLayout3.xml', master=2, ph={
        2: ('body', 10, 'PART 번호'), 3: ('title', 0, '구분 제목'), 6: ('body', 11, '부서 또는 설명')}),
    4: dict(name='본문', base='slideLayout1.xml', master=1, ph={
        3: ('title', 0, '제목'), 46: ('sldNum', 12, '‹#›')},
        body=(548640, 1450000, 11064240, 4650000)),
    5: dict(name='마무리', base='slideLayout1.xml', master=1, ph={
        5: ('ctrTitle', 0, '마무리 문구'), 12: ('body', 10, '날짜')}),
}


def q(tag):
    pfx, name = tag.split(':')
    return '{%s}%s' % (NS[pfx], name)


def parse(path):
    return etree.parse(path, etree.XMLParser(remove_blank_text=False))


def write(tree, path):
    tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)


def read_rels(path):
    if not os.path.exists(path):
        return []
    t = parse(path)
    return [dict(Id=e.get('Id'), Type=e.get('Type'), Target=e.get('Target'), Mode=e.get('TargetMode'))
            for e in t.getroot()]


def write_rels(path, rels):
    root = etree.Element('{%s}Relationships' % REL_NS, nsmap={None: REL_NS})
    for r in rels:
        el = etree.SubElement(root, '{%s}Relationship' % REL_NS, Id=r['Id'], Type=r['Type'], Target=r['Target'])
        if r.get('Mode'):
            el.set('TargetMode', r['Mode'])
    etree.ElementTree(root).write(path, xml_declaration=True, encoding='UTF-8', standalone=True)


# ---------- 플레이스홀더 변환 ----------
ORDER = ['lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx',
         'buFont', 'buNone', 'buAutoNum', 'buChar', 'buBlip', 'tabLst']


def local(el):
    return etree.QName(el).localname


def lvl1_from_paragraph(p):
    """첫 문단/첫 run의 서식을 레이아웃 lstStyle(lvl1pPr/defRPr)로 만든다."""
    lst = etree.Element(q('a:lstStyle'))
    lvl = etree.SubElement(lst, q('a:lvl1pPr'))
    pPr = p.find(q('a:pPr'))
    kids = {}
    if pPr is not None:
        for k, v in pPr.attrib.items():
            if k in ('algn', 'marL', 'indent', 'lvl', 'rtl'):
                lvl.set(k, v)
        for c in pPr:
            if local(c) != 'defRPr':
                kids[local(c)] = copy.deepcopy(c)
    lvl.set('marL', lvl.get('marL', '0'))
    lvl.set('indent', lvl.get('indent', '0'))
    lvl.set('algn', lvl.get('algn', 'l'))
    if not any(k.startswith('bu') for k in kids):
        kids['buNone'] = etree.Element(q('a:buNone'))
    if 'lnSpc' not in kids:
        e = etree.Element(q('a:lnSpc')); etree.SubElement(e, q('a:spcPct')).set('val', '100000'); kids['lnSpc'] = e
    if 'spcBef' not in kids:
        e = etree.Element(q('a:spcBef')); etree.SubElement(e, q('a:spcPts')).set('val', '0'); kids['spcBef'] = e
    for name in ORDER:
        if name in kids:
            lvl.append(kids[name])
    rPr = None
    r = p.find(q('a:r'))
    if r is None:
        r = p.find(q('a:fld'))
    if r is not None:
        rPr = r.find(q('a:rPr'))
    if rPr is None:
        rPr = p.find(q('a:endParaRPr'))
    d = etree.SubElement(lvl, q('a:defRPr'))
    if rPr is not None:
        for k, v in rPr.attrib.items():
            if k in ('sz', 'b', 'i', 'spc', 'kern', 'cap', 'baseline', 'u', 'strike'):
                d.set(k, v)
        for c in rPr:
            d.append(copy.deepcopy(c))
    return lst


def to_placeholder_layout(sp, ph_type, idx, prompt):
    """슬라이드의 텍스트 도형을 레이아웃 플레이스홀더로 바꾼다(위치/배경은 유지)."""
    sp = copy.deepcopy(sp)
    nv = sp.find(q('p:nvSpPr'))
    cNvSpPr = nv.find(q('p:cNvSpPr'))
    for k in list(cNvSpPr.attrib):
        if k == 'txBox':
            del cNvSpPr.attrib[k]
    if cNvSpPr.find(q('a:spLocks')) is None:
        etree.SubElement(cNvSpPr, q('a:spLocks')).set('noGrp', '1')
    nvPr = nv.find(q('p:nvPr'))
    for c in list(nvPr):
        nvPr.remove(c)
    ph = etree.SubElement(nvPr, q('p:ph'))
    ph.set('type', ph_type)
    ph.set('idx', str(idx))
    if ph_type == 'sldNum':
        ph.set('sz', 'quarter')
    tx = sp.find(q('p:txBody'))
    first_p = tx.find(q('a:p'))
    new_tx = etree.Element(q('p:txBody'))
    body_pr = tx.find(q('a:bodyPr'))
    new_tx.append(copy.deepcopy(body_pr) if body_pr is not None else etree.Element(q('a:bodyPr')))
    new_tx.append(lvl1_from_paragraph(first_p))
    p = etree.SubElement(new_tx, q('a:p'))
    if ph_type == 'sldNum':
        f = etree.SubElement(p, q('a:fld'))
        f.set('id', '{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}')
        f.set('type', 'slidenum')
        etree.SubElement(f, q('a:rPr')).set('lang', 'ko-KR')
        etree.SubElement(f, q('a:t')).text = prompt
    else:
        r = etree.SubElement(p, q('a:r'))
        etree.SubElement(r, q('a:rPr')).set('lang', 'ko-KR')
        etree.SubElement(r, q('a:t')).text = prompt
    sp.replace(tx, new_tx)
    return sp


def to_placeholder_slide(sp, ph_type, idx):
    """슬라이드 쪽 플레이스홀더: 위치·서식은 레이아웃에서 상속, 내용은 원본 문단을 그대로 유지."""
    sp = copy.deepcopy(sp)
    nv = sp.find(q('p:nvSpPr'))
    cNvSpPr = nv.find(q('p:cNvSpPr'))
    for k in list(cNvSpPr.attrib):
        if k == 'txBox':
            del cNvSpPr.attrib[k]
    if cNvSpPr.find(q('a:spLocks')) is None:
        etree.SubElement(cNvSpPr, q('a:spLocks')).set('noGrp', '1')
    nvPr = nv.find(q('p:nvPr'))
    for c in list(nvPr):
        nvPr.remove(c)
    ph = etree.SubElement(nvPr, q('p:ph'))
    ph.set('type', ph_type)
    ph.set('idx', str(idx))
    if ph_type == 'sldNum':
        ph.set('sz', 'quarter')
    spPr = sp.find(q('p:spPr'))
    for c in list(spPr):
        spPr.remove(c)
    tx = sp.find(q('p:txBody'))
    for c in list(tx):
        if local(c) in ('bodyPr', 'lstStyle'):
            tx.remove(c)
    tx.insert(0, etree.Element(q('a:lstStyle')))
    tx.insert(0, etree.Element(q('a:bodyPr')))
    # 문단 서식(pPr)은 레이아웃이 담당하므로 비우되, run 서식은 유지한다
    for p in tx.findall(q('a:p')):
        pPr = p.find(q('a:pPr'))
        if pPr is not None:
            p.remove(pPr)
    return sp


def body_placeholder_layout(area, shape_id):
    x, y, w, h = area
    xml = f'''<p:sp xmlns:a="{NS['a']}" xmlns:p="{NS['p']}">
<p:nvSpPr><p:cNvPr id="{shape_id}" name="본문 개체 틀"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:normAutofit/></a:bodyPr>
<a:lstStyle>
<a:lvl1pPr marL="285750" indent="-285750" algn="l"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="900"/></a:spcBef><a:buClr><a:srgbClr val="00387F"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/><a:defRPr sz="2000" b="0"><a:solidFill><a:srgbClr val="262626"/></a:solidFill><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/><a:cs typeface="Malgun Gothic"/></a:defRPr></a:lvl1pPr>
<a:lvl2pPr marL="571500" indent="-285750" algn="l"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:buClr><a:srgbClr val="00387F"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="&#8211;"/><a:defRPr sz="1800" b="0"><a:solidFill><a:srgbClr val="262626"/></a:solidFill><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/><a:cs typeface="Malgun Gothic"/></a:defRPr></a:lvl2pPr>
</a:lstStyle>
<a:p><a:r><a:rPr lang="ko-KR"/><a:t>본문 내용을 입력하세요</a:t></a:r></a:p>
<a:p><a:pPr lvl="1"/><a:r><a:rPr lang="ko-KR"/><a:t>하위 항목</a:t></a:r></a:p>
</p:txBody></p:sp>'''
    return etree.fromstring(xml)


def body_placeholder_slide(shape_id):
    xml = f'''<p:sp xmlns:a="{NS['a']}" xmlns:p="{NS['p']}">
<p:nvSpPr><p:cNvPr id="{shape_id}" name="본문 개체 틀"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:r><a:rPr lang="ko-KR"/><a:t>본문 내용을 입력하세요</a:t></a:r></a:p>
<a:p><a:pPr lvl="1"/><a:r><a:rPr lang="ko-KR"/><a:t>하위 항목은 Tab 키로 들여쓰기</a:t></a:r></a:p>
</p:txBody></p:sp>'''
    return etree.fromstring(xml)


def main(src, dst):
    work = tempfile.mkdtemp()
    with zipfile.ZipFile(src) as z:
        z.extractall(work)
        order = [i.filename for i in z.infolist()]
    ppt = os.path.join(work, 'ppt')
    ct_path = os.path.join(work, '[Content_Types].xml')
    ct = parse(ct_path)
    masters = {}
    for m in (1, 2):
        mp = os.path.join(ppt, 'slideMasters', f'slideMaster{m}.xml')
        masters[m] = dict(path=mp, tree=parse(mp), rels_path=os.path.join(ppt, 'slideMasters', '_rels', f'slideMaster{m}.xml.rels'))
        masters[m]['rels'] = read_rels(masters[m]['rels_path'])

    next_layout = 5
    next_lid = 2147483700
    for sn, spec in SPEC.items():
        slide_path = os.path.join(ppt, 'slides', f'slide{sn}.xml')
        slide_rels_path = os.path.join(ppt, 'slides', '_rels', f'slide{sn}.xml.rels')
        stree = parse(slide_path)
        srels = read_rels(slide_rels_path)
        sp_tree = stree.getroot().find(q('p:cSld')).find(q('p:spTree'))

        base_path = os.path.join(ppt, 'slideLayouts', spec['base'])
        ltree = parse(base_path)
        lroot = ltree.getroot()
        if 'type' in lroot.attrib:
            del lroot.attrib['type']
        lroot.set('preserve', '1')
        lroot.set('userDrawn', '1')
        lroot.find(q('p:cSld')).set('name', spec['name'])
        l_sptree = lroot.find(q('p:cSld')).find(q('p:spTree'))
        if spec.get('strip_base_ph'):
            for sp in l_sptree.findall(q('p:sp')):
                if sp.find('.//' + q('p:ph')) is not None:
                    l_sptree.remove(sp)

        base_rels_path = os.path.join(ppt, 'slideLayouts', '_rels', spec['base'] + '.rels')
        lrels = read_rels(base_rels_path)
        # 새 레이아웃의 관계: 마스터는 해당 마스터로 고정
        for r in lrels:
            if r['Type'].endswith('/slideMaster'):
                r['Target'] = f"../slideMasters/slideMaster{spec['master']}.xml"
        used = {r['Id'] for r in lrels}
        rmap = {}
        k = 0
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

        max_id = max([int(e.get('id')) for e in l_sptree.iter(q('p:cNvPr')) if e.get('id', '').isdigit()] + [1])
        next_id = max_id + 1
        slide_new_children = []
        id_map = {}

        for child in list(sp_tree):
            if local(child) in ('nvGrpSpPr', 'grpSpPr'):
                continue
            cid_el = child.find('.//' + q('p:cNvPr'))
            sid = int(cid_el.get('id')) if cid_el is not None else None
            if sid in spec['ph'] and local(child) == 'sp':
                ph_type, idx, prompt = spec['ph'][sid]
                lay_sp = to_placeholder_layout(child, ph_type, idx, prompt)
                slide_sp = to_placeholder_slide(child, ph_type, idx)
                for c in lay_sp.iter(q('p:cNvPr')):
                    c.set('id', str(next_id)); id_map[sid] = next_id; next_id += 1
                    break
                l_sptree.append(lay_sp)
                slide_sp.find('.//' + q('p:cNvPr')).set('id', str(len(slide_new_children) + 2))
                slide_new_children.append(slide_sp)
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

        if 'body' in spec:
            l_sptree.append(body_placeholder_layout(spec['body'], next_id)); next_id += 1
            slide_new_children.append(body_placeholder_slide(len(slide_new_children) + 2))

        # 레이아웃 저장(+ 항목 수별 변형)
        def save_layout(tree, name):
            nonlocal next_layout, next_lid
            lname = f'slideLayout{next_layout}.xml'
            tree.getroot().find(q('p:cSld')).set('name', name)
            write(tree, os.path.join(ppt, 'slideLayouts', lname))
            write_rels(os.path.join(ppt, 'slideLayouts', '_rels', lname + '.rels'), lrels)
            etree.SubElement(ct.getroot(), '{http://schemas.openxmlformats.org/package/2006/content-types}Override',
                             PartName=f'/ppt/slideLayouts/{lname}', ContentType=CT_LAYOUT)
            m = masters[spec['master']]
            rids = {r['Id'] for r in m['rels']}
            n = 1
            while f'rId{n}' in rids:
                n += 1
            rid = f'rId{n}'
            m['rels'].append(dict(Id=rid, Type=T_LAYOUT, Target=f'../slideLayouts/{lname}'))
            lst = m['tree'].getroot().find(q('p:sldLayoutIdLst'))
            el = etree.SubElement(lst, q('p:sldLayoutId'))
            el.set('id', str(next_lid)); next_lid += 1
            el.set(R + 'id', rid)
            next_layout += 1
            return lname

        lname = save_layout(ltree, spec['name'])
        for n_items, drop in (spec.get('variants') or {}).items():
            vt = copy.deepcopy(ltree)
            vtree = vt.getroot().find(q('p:cSld')).find(q('p:spTree'))
            drop_new = {str(id_map[x]) for x in drop if x in id_map}
            for ch in list(vtree):
                e = ch.find('.//' + q('p:cNvPr'))
                if e is not None and e.get('id') in drop_new:
                    vtree.remove(ch)
            save_layout(vt, f'목차 {n_items}개')
        # 슬라이드는 항목 수가 가장 많은 첫 레이아웃을 쓴다
        lname_for_slide = lname

        # 슬라이드 다시 쓰기: 플레이스홀더만 남기고 레이아웃을 새 것으로 교체
        for child in list(sp_tree):
            if local(child) not in ('nvGrpSpPr', 'grpSpPr'):
                sp_tree.remove(child)
        for c in slide_new_children:
            sp_tree.append(c)
        write(stree, slide_path)
        new_srels = []
        for r in srels:
            if r['Type'].endswith('/slideLayout'):
                r['Target'] = f'../slideLayouts/{lname_for_slide}'
                new_srels.append(r)
            elif r['Type'].endswith('/notesSlide'):
                new_srels.append(r)
        write_rels(slide_rels_path, new_srels)

    write(ct, ct_path)
    for m in masters.values():
        write(m['tree'], m['path'])
        write_rels(m['rels_path'], m['rels'])

    # 새 파일 목록으로 압축([Content_Types].xml을 맨 앞에)
    files = []
    for root, _, fs in os.walk(work):
        for f in fs:
            files.append(os.path.relpath(os.path.join(root, f), work).replace(os.sep, '/'))
    files.sort(key=lambda f: (f != '[Content_Types].xml', order.index(f) if f in order else 10 ** 6, f))
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(os.path.join(work, f), f)
    shutil.rmtree(work)
    print('변환 완료:', dst)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])
