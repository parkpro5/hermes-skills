"""PPT 서식 기반 생성용 공통 함수 (python-pptx, lxml만 사용)."""
import copy
import json
import math
import os

from lxml import etree
from pptx.oxml.ns import qn

RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
EMU_PT = 12700


# ---------- 서식 설정(roles.json) ----------
PTB_NS = 'urn:ppt-template-builder:roles'
RT_CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates')


def read_embedded_roles(path):
    """pptx 안에 내장된 roles(customXml)를 읽는다. 없으면 None."""
    import zipfile
    from xml.etree import ElementTree as ET
    try:
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if name.startswith('customXml/') and name.endswith('.xml') and '_rels' not in name:
                    blob = z.read(name)
                    if PTB_NS.encode() in blob:
                        return json.loads(ET.fromstring(blob).text)
    except (zipfile.BadZipFile, OSError):
        return None
    return None


def embed_roles(pptx_path, roles, out_path=None):
    """roles를 pptx 안(customXml)에 넣어 파일 하나로 공유·배포할 수 있게 한다."""
    from xml.sax.saxutils import escape
    from pptx import Presentation
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI
    prs = Presentation(pptx_path)
    for rid, rel in list(prs.part.rels.items()):
        if rel.reltype == RT_CUSTOM_XML and not rel.is_external and PTB_NS.encode() in rel.target_part.blob:
            prs.part.drop_rel(rid)
    body = escape(json.dumps(roles, ensure_ascii=False))
    blob = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<ptb:roles xmlns:ptb="{PTB_NS}">{body}</ptb:roles>').encode('utf8')
    part = Part(PackURI('/customXml/ptbRoles.xml'), 'application/xml', prs.part.package, blob)
    prs.part.relate_to(part, RT_CUSTOM_XML)
    prs.save(out_path or pptx_path)


def load_roles(template_path):
    """옆의 .roles.json(수정 가능)을 우선 쓰고, 없으면 pptx에 내장된 roles를 쓴다."""
    base = os.path.splitext(template_path)[0]
    p = base + '.roles.json'
    if os.path.exists(p):
        with open(p, encoding='utf8') as f:
            return json.load(f)
    roles = read_embedded_roles(template_path)
    if roles is not None:
        return roles
    raise SystemExit(f'서식 설정(roles)이 없습니다: {template_path}\n'
                     'auto_convert.py로 서식화한 파일이거나 옆에 .roles.json이 있어야 합니다. '
                     '새 서식이면 inspect_template.py → auto_convert.py 순서로 변환하세요.')


def load_registry():
    p = os.path.join(TEMPLATES_DIR, 'registry.json')
    if not os.path.exists(p):
        return dict(default=None, templates=[])
    with open(p, encoding='utf8') as f:
        return json.load(f)


def resolve_template(arg=None):
    """파일 경로 또는 등록된 서식 id를 파일 경로로 바꾼다. 생략하면 기본 서식."""
    reg = load_registry()
    if arg is None:
        arg = reg.get('default')
        if arg is None:
            raise SystemExit('기본 서식이 없습니다. --template으로 서식 파일이나 등록된 id를 지정하세요.')
    if os.path.exists(arg):
        return arg
    for e in reg['templates']:
        if e['id'] == arg:
            return os.path.normpath(os.path.join(TEMPLATES_DIR, e['file']))
    ids = ', '.join(e['id'] for e in reg['templates']) or '(없음)'
    raise SystemExit(f"서식 '{arg}'을(를) 찾을 수 없습니다. 파일 경로이거나 등록된 id여야 합니다. 등록된 id: {ids}")


# ---------- 도형 탐색 ----------
def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:  # GROUP
            yield from iter_shapes(sh.shapes)


def shape_text(sh):
    return sh.text_frame.text if sh.has_text_frame else ''


def find_shape(slide, name=None, text=None):
    """text가 있으면 텍스트가 같은 도형, 없으면 이름이 같은 첫 도형. name이 '#12'이면 도형 id 12."""
    if name is not None and str(name).startswith('#') and text is None:
        for sh in iter_shapes(slide.shapes):
            if str(sh.shape_id) == str(name)[1:]:
                return sh
        return None
    for sh in iter_shapes(slide.shapes):
        if text is not None and sh.has_text_frame and sh.text_frame.text.strip() == text:
            return sh
    if name is not None and text is None:
        for sh in iter_shapes(slide.shapes):
            if sh.name == name:
                return sh
    return None


def remove_shape(sh):
    sh._element.getparent().remove(sh._element)


# ---------- 슬라이드 복제/삭제 ----------
def duplicate_slide(prs, src):
    """src 슬라이드를 맨 뒤에 복제한다(도형, 그림, 배경 관계 포함)."""
    new = prs.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    rmap = {}
    for rel in src.part.rels.values():
        if 'notesSlide' in rel.reltype or 'slideLayout' in rel.reltype:
            continue
        rmap[rel.rId] = (new.part.rels.get_or_add_ext_rel(rel.reltype, rel._target)
                         if rel.is_external else
                         new.part.rels.get_or_add(rel.reltype, rel._target))

    def remap(el):
        for node in el.iter():
            for k, v in list(node.attrib.items()):
                if k.startswith(RNS) and v in rmap:
                    node.set(k, rmap[v])

    for el in src.shapes._spTree:
        if el.tag.endswith('}nvGrpSpPr') or el.tag.endswith('}grpSpPr'):
            continue
        c = copy.deepcopy(el)
        remap(c)
        new.shapes._spTree.append(c)
    bg = src._element.find(qn('p:cSld')).find(qn('p:bg'))
    if bg is not None:
        nb = copy.deepcopy(bg)
        remap(nb)
        new._element.find(qn('p:cSld')).insert(0, nb)
    return new


def drop_slide(prs, idx):
    lst = prs.slides._sldIdLst
    sld = lst[idx]
    prs.part.drop_rel(sld.get(qn('r:id')))
    lst.remove(sld)


# ---------- 텍스트/폰트 ----------
def set_font(run, name):
    """한글(ea)과 영문(latin), 복합(cs) 폰트를 모두 지정한다. name이 없으면 아무것도 하지 않는다."""
    if not name:
        return
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set('typeface', name)


def apply_font(shape_or_tf, name):
    if not name:
        return
    tf = shape_or_tf.text_frame if hasattr(shape_or_tf, 'text_frame') else shape_or_tf
    for p in tf.paragraphs:
        for r in p.runs:
            set_font(r, name)


def fill_paragraph(p, line):
    """첫 run의 서식을 유지한 채 문단 텍스트를 교체한다.
    run이 2개 이상이고 line에 공백이 있으면 첫 단어는 첫 run(강조색 등), 나머지는 둘째 run에 넣는다."""
    runs = p.runs
    if not runs:
        p.add_run()
        runs = p.runs
    if len(runs) >= 2 and ' ' in line:
        first, rest = line.split(' ', 1)
        runs[0].text = first
        runs[1].text = ' ' + rest
        for r in runs[2:]:
            r._r.getparent().remove(r._r)
    else:
        runs[0].text = line
        for r in runs[1:]:
            r._r.getparent().remove(r._r)


def set_lines(shape, lines, font=None):
    """도형의 문단을 lines로 교체한다. 문단 서식은 서식 문단 순서대로 상속한다."""
    if isinstance(lines, str):
        lines = [lines]
    tf = shape.text_frame
    paras = list(tf.paragraphs)
    for i, line in enumerate(lines):
        if i < len(paras):
            p = paras[i]
        else:
            new_p = copy.deepcopy(paras[-1]._p)
            paras[-1]._p.addnext(new_p)
            paras = list(tf.paragraphs)
            p = paras[i]
        fill_paragraph(p, line)
    for p in list(tf.paragraphs)[len(lines):]:
        p._p.getparent().remove(p._p)
    apply_font(tf, font)


def set_text(shape, text, font=None):
    set_lines(shape, [text], font)


# ---------- 길이 추정 (검증과 쪽 나누기에 공통 사용) ----------
def char_width_pt(ch, size):
    o = ord(ch)
    if o >= 0x2E80:
        return size * 1.0
    if ch == ' ':
        return size * 0.3
    if ch.isupper() or ch.isdigit():
        return size * 0.62
    return size * 0.52


def text_width_pt(s, size):
    return sum(char_width_pt(c, size) for c in s)


def est_lines(s, size, avail_pt):
    if not s:
        return 1
    return max(1, math.ceil(text_width_pt(s, size) * 1.05 / max(avail_pt, 1)))


# ---------- 역할 감지 ----------
def detect_role(slide, roles=None):
    """플레이스홀더 서식: 레이아웃 이름으로, 도형 서식(복제 방식): 도형 이름과 텍스트로 판단한다."""
    tag = slide._element.find(qn('p:cSld')).get('name')       # 복제 방식 빌더가 표시한 역할
    if tag in ('cover', 'agenda', 'section', 'content', 'closing'):
        return tag
    if roles and roles.get('layouts'):
        name = slide.slide_layout.name
        for role, prefix in roles['layouts'].items():
            if name.startswith(prefix):
                return role
    names = {sh.name for sh in iter_shapes(slide.shapes)}
    texts = [shape_text(sh).strip() for sh in iter_shapes(slide.shapes)]
    if 'TextBox 9' in names and 'Text 2' in names:
        return 'cover'
    if any(t == 'AGENDA' for t in texts):
        return 'agenda'
    if any(t.startswith('PART ') for t in texts):
        return 'section'
    if 'TextBox 4' in names:
        return 'closing'
    return 'content'


def ph_idx(sh):
    return sh.placeholder_format.idx if getattr(sh, 'is_placeholder', False) else None


def inherited_size_pt(sh):
    """플레이스홀더가 레이아웃에서 상속하는 글자 크기(pt). 없으면 None."""
    if not getattr(sh, 'is_placeholder', False):
        return None
    try:
        layout = sh.part.slide.slide_layout
    except Exception:
        return None
    for lph in layout.placeholders:
        if lph.placeholder_format.idx == sh.placeholder_format.idx:
            for el in lph._element.iter(qn('a:defRPr')):
                if el.get('sz'):
                    return int(el.get('sz')) / 100
            break
    return None


def rect_intersects(a, b, tol=0):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return not (ax + aw <= bx + tol or bx + bw <= ax + tol or
                ay + ah <= by + tol or by + bh <= ay + tol)
