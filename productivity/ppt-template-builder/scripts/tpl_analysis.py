"""서식·샘플 자료 분석 라이브러리 (python-pptx만 사용).

analyze(path)  : 슬라이드 유형 분류 → 같은 디자인끼리 묶기(클러스터) → 도형 역할(고정/가변) 추정
recommend(a)   : 복제 방식 vs 플레이스홀더 방식 추천(점수와 근거)
"""
import hashlib
import os
import re
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from pptx import Presentation  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE_TYPE as MT  # noqa: E402

from pptx_lib import inherited_size_pt  # noqa: E402

NUM_RE = re.compile(r'^\d{1,2}$')
LABEL_RE = re.compile(r'^[A-Z][A-Z &/._-]{1,24}$')          # AGENDA, STECH AX / DX 같은 영문 대문자 라벨
DATE_RE = re.compile(r'\d{4}\s*[.\-/]\s*\d{1,2}\s*[.\-/]\s*\d{1,2}')
PART_RE = re.compile(r'^\s*(PART|Part|파트|CHAPTER|Chapter|챕터|SECTION|Section|섹션)\s*[\dIVX]+')
KW_CLOSE = ('감사', 'thank', 'q&a', 'q & a', '질의', '문의')
KW_AGENDA = ('목차', 'agenda', 'contents', '차례')
SKIP_PH = ('SLIDE_NUMBER', 'DATE', 'FOOTER')


# ---------------- 도형 정보 ----------------
def _ph_type(sh):
    if not getattr(sh, 'is_placeholder', False):
        return None
    return str(sh.placeholder_format.type).split(' ')[0]


def _runs(sh):
    if not sh.has_text_frame:
        return []
    return [r for p in sh.text_frame.paragraphs for r in p.runs]


def _group_texts(sh):
    out = []
    for c in sh.shapes:
        if c.shape_type == MT.GROUP:
            out += _group_texts(c)
        elif c.has_text_frame and c.text_frame.text.strip():
            out.append(c.text_frame.text.strip())
    return out


def shape_info(sh, z, W, H):
    st = sh.shape_type
    ph = _ph_type(sh)
    has_table = bool(getattr(sh, 'has_table', False))
    has_chart = bool(getattr(sh, 'has_chart', False))
    if ph:
        kind = 'ph'
    elif st == MT.GROUP:
        kind = 'group'
    elif st == MT.PICTURE:
        kind = 'picture'
    elif has_table:
        kind = 'table'
    elif has_chart:
        kind = 'chart'
    elif st in (MT.LINE,):
        kind = 'line'
    elif sh.shape_type in (MT.EMBEDDED_OLE_OBJECT, MT.LINKED_OLE_OBJECT, MT.MEDIA, MT.WEB_VIDEO) or \
            sh._element.tag.endswith('graphicFrame'):
        kind = 'frame'
    else:
        kind = 'shape'
    text = ''
    if kind != 'group' and sh.has_text_frame:
        text = sh.text_frame.text.strip()
    sizes = [r.font.size.pt for r in _runs(sh) if r.font.size]
    size = max(sizes) if sizes else (inherited_size_pt(sh) if ph else None)
    d = dict(
        id=sh.shape_id, name=sh.name, z=z, kind=kind, ph=ph,
        ph_idx=sh.placeholder_format.idx if ph else None,
        x=int(sh.left or 0), y=int(sh.top or 0), w=int(sh.width or 0), h=int(sh.height or 0),
        text=text, size=size, size_known=size is not None,
        bold=any(r.font.bold for r in _runs(sh)),
        font=next((r.font.name for r in _runs(sh) if r.font.name), None),
        n_par=len(sh.text_frame.paragraphs) if (kind != 'group' and sh.has_text_frame) else 0,
        rot=float(getattr(sh, 'rotation', 0) or 0),
        gtext=_group_texts(sh) if kind == 'group' else [],
        img=None,
    )
    if kind == 'picture':
        try:
            d['img'] = hashlib.md5(sh.image.blob).hexdigest()[:10]
        except Exception:
            pass
    d['has_text'] = bool(text) or bool(d['gtext'])
    d['fld_num'] = any(f.get('type', '') == 'slidenum' for f in sh._element.iter(
        '{http://schemas.openxmlformats.org/drawingml/2006/main}fld'))
    return d


def slide_info(slide, i, W, H):
    shapes = [shape_info(sh, z, W, H) for z, sh in enumerate(slide.shapes)]
    has_bg = slide._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld') \
        .find('{http://schemas.openxmlformats.org/presentationml/2006/main}bg') is not None
    return dict(i=i, layout=slide.slide_layout.name, shapes=shapes, has_bg=has_bg)


def texts_of(si):
    return [s for s in si['shapes'] if s['text'] and not (s['ph'] in SKIP_PH)]


def is_pagenum(s, H):
    return bool(NUM_RE.match(s['text'])) and s['y'] > 0.86 * H and (s['size'] or 12) <= 14


def pick_title(shapes):
    cand = [s for s in shapes if s['text'] and s['ph'] not in SKIP_PH and not NUM_RE.match(s['text'])]
    ph_t = [s for s in cand if s['ph'] in ('TITLE', 'CENTER_TITLE')]
    if ph_t:
        return ph_t[0]
    if not cand:
        return None
    return sorted(cand, key=lambda s: (-(s['size'] or 18), s['y'], s['x']))[0]


# ---------------- 슬라이드 유형 분류 ----------------
def classify(si, i, n, W, H):
    shapes = si['shapes']
    tx = texts_of(si)
    all_text = [s['text'].lower() for s in tx]
    title = pick_title(shapes)
    tsize = (title['size'] or 18) if title else 0
    nums = [s for s in tx if NUM_RE.match(s['text']) and not is_pagenum(s, H)]
    has_table = any(s['kind'] == 'table' for s in shapes)
    has_chart = any(s['kind'] == 'chart' for s in shapes)
    ev = []

    def has_kw(kws):
        return any(any(k in t for k in kws) for t in all_text)

    # 닫는 슬라이드
    if has_kw(KW_CLOSE) and len(tx) <= 5 and not has_table:
        return 'closing', (0.9 if i == n - 1 else 0.7), ['마무리 문구 + 텍스트 적음']
    if i == n - 1 and len(tx) <= 3 and tsize >= 28 and not has_table:
        return 'closing', 0.6, ['마지막 슬라이드 + 큰 짧은 문구']
    if i == 0:
        conf = 0.9 if (tsize >= 28 or len(tx) <= 6) else 0.6
        return 'cover', conf, ['첫 슬라이드']
    # 목차: 연속 번호 3개 이상 또는 목차 키워드
    if len(nums) >= 3:
        vals = sorted(int(s['text']) for s in nums)
        if vals == list(range(vals[0], vals[0] + len(vals))):
            return 'agenda', 0.9, [f'연속 번호 {len(nums)}개']
    if has_kw(KW_AGENDA) and (len(nums) >= 2 or len(tx) >= 7):
        return 'agenda', 0.75, ['목차 키워드']
    # 구분
    if any(PART_RE.match(s['text']) for s in tx):
        return 'section', 0.85, ['PART/챕터 번호 문구']
    others = [s for s in shapes if s['kind'] in ('table', 'chart', 'frame')]
    if len(tx) <= 4 and tsize >= 28 and not others and 0 < i < n - 1:
        return 'section', 0.55, ['큰 제목만 있는 슬라이드']
    return 'content', 0.8 if title else 0.55, ev


def content_subkind(si, W, H):
    shapes = si['shapes']
    if any(s['kind'] == 'table' for s in shapes):
        return 'table'
    if any(s['kind'] == 'chart' for s in shapes):
        return 'chart'
    pics = sum(s['w'] * s['h'] for s in shapes if s['kind'] == 'picture')
    if pics > 0.25 * W * H:
        return 'image'
    title = pick_title(shapes)
    body = [s for s in shapes if s['text'] and s is not title and s['ph'] not in SKIP_PH
            and s['w'] >= 0.25 * W and s['h'] >= 0.15 * H and not NUM_RE.match(s['text'])]
    if len(body) >= 2:
        body.sort(key=lambda s: s['x'])
        a, b = body[0], body[1]
        if a['x'] + a['w'] <= b['x'] + 0.05 * W and abs(a['y'] - b['y']) < 0.15 * H:
            return 'two-col'
    decor = sum(1 for s in shapes if s['kind'] in ('shape', 'group') and not s['text'])
    if decor >= 9:
        return 'diagram'
    return 'text'


# ---------------- 반복 도형(고정 디자인) 판정 ----------------
def _gkey(s, W, H):
    return (s['kind'], round(s['x'] / (W * 0.02)), round(s['y'] / (H * 0.02)),
            round(s['w'] / (W * 0.05)), round(s['h'] / (H * 0.05)))


def repeat_map(slides, W, H):
    cnt = defaultdict(set)
    for si in slides:
        for s in si['shapes']:
            k = (_gkey(s, W, H), s['img'], s['text'] if s['kind'] != 'group' else '|'.join(s['gtext']))
            cnt[k].add(si['i'])
    return cnt


def is_decor(s, cnt, deck_like, W, H):
    """슬라이드마다 반복되는(=서식의) 글자 없는 도형인가. 표/차트는 항상 내용."""
    if s['kind'] in ('table', 'chart', 'frame'):
        return False
    if s['text'] or s['ph']:
        return False
    if s['kind'] == 'group' and s['gtext']:
        return False
    if not deck_like:
        return True
    k = (_gkey(s, W, H), s['img'], '|'.join(s['gtext']) if s['kind'] == 'group' else '')
    return len(cnt.get(k, ())) >= 2


def is_decor_in(s, cluster_idxs, slides, cnt, W, H):
    """클러스터 기준 판정: 표/차트는 내용. 슬라이드가 1장뿐인 유형(표지·목차 등)의 도형은 모두 서식.
    여러 장이면 클러스터 안에서 반복되는 도형만 서식, 한 번뿐인 그림/그룹은 내용으로 본다."""
    if s['kind'] in ('table', 'chart', 'frame') or s['text'] or s['ph']:
        return False
    if len(cluster_idxs) == 1:
        return True
    k = (_gkey(s, W, H), s['img'], '|'.join(s['gtext']) if s['kind'] == 'group' else '')
    same = sum(1 for i in cluster_idxs if any(
        (_gkey(t, W, H), t['img'], '|'.join(t['gtext']) if t['kind'] == 'group' else '') == k for t in slides[i]['shapes']))
    if same >= 2:
        return True
    if s['kind'] == 'group' and s['gtext']:
        return False
    return s['kind'] in ('shape', 'line')      # 글자 없는 단순 도형/선은 서식, 그림·그룹은 내용


def decor_sig(si, cnt, deck_like, W, H):
    return tuple(sorted(_gkey(s, W, H) for s in si['shapes'] if is_decor(s, cnt, deck_like, W, H)
                        and s['kind'] in ('line', 'picture', 'group')))


def _match(a, b, W, H):
    return (a['kind'] == b['kind'] and abs(a['x'] - b['x']) < 0.03 * W and abs(a['y'] - b['y']) < 0.06 * H
            and abs(a['w'] - b['w']) < 0.08 * W)


# ---------------- 목차 격자 분석 ----------------
def analyze_agenda(si, W, H):
    shapes = si['shapes']
    nums = [s for s in shapes if s['text'] and NUM_RE.match(s['text']) and not is_pagenum(s, H) and s['ph'] not in SKIP_PH]
    nums.sort(key=lambda s: int(s['text']))
    if len(nums) < 2:
        return None
    rest = [s for s in shapes if s not in nums and s['text'] and s['ph'] not in SKIP_PH]
    heads = [s for s in rest if any(k in s['text'].lower() for k in KW_AGENDA)]
    title = pick_title([s for s in rest if s not in heads]) if rest else None
    texts = [s for s in rest if s is not title and s not in heads]

    def cen(s):
        return (s['x'] + s['w'] / 2, s['y'] + s['h'] / 2)

    items = [dict(num=n, texts=[], deco=[]) for n in nums]
    for t in texts:
        ax, ay = t['x'], t['y'] + t['h'] / 2          # 글자 상자의 왼쪽 가운데
        best = min(items, key=lambda it: ((it['num']['x'] + it['num']['w'] - ax) / W) ** 2 * 1.0
                   + ((it['num']['y'] + it['num']['h'] / 2 - ay) / H) ** 2 * 1.5)
        best['texts'].append(t)
    for it in items:
        it['texts'].sort(key=lambda s: (-(s['size'] or 18), s['y']))     # 큰 글자 = 제목
        it['title'] = it['texts'][0] if it['texts'] else None
        it['org'] = it['texts'][1] if len(it['texts']) > 1 else None
        boxes = [s for s in [it['num'], it['title'], it['org']] if s]
        it['cell'] = (min(s['x'] for s in boxes) - 0.01 * W, min(s['y'] for s in boxes) - 0.02 * H,
                      min(s['x'] for s in boxes) + 0.45 * W, max(s['y'] + s['h'] for s in boxes) + 0.02 * H)
    used = {s['id'] for it in items for s in (it['num'], it['title'], it['org']) if s}
    for s in shapes:
        if s['id'] in used or s['text'] or s['ph'] or s['kind'] in ('table', 'chart', 'frame', 'picture', 'line'):
            continue
        if s['w'] * s['h'] > 0.3 * W * H:
            continue
        cx, cy = cen(s)
        hits = [it for it in items if it['cell'][0] <= cx <= it['cell'][2] and it['cell'][1] <= cy <= it['cell'][3]]
        if hits:
            min(hits, key=lambda it: abs(cen(it['num'])[0] - cx) + abs(cen(it['num'])[1] - cy))['deco'].append(s)
    return dict(title=title, items=items)


# ---------------- 역할 결정 ----------------
def decide_roles(cluster, slides, cnt, deck_like, W, H):
    """대표 슬라이드의 글자 도형마다 고정(fixed)/가변(variable)을 정한다."""
    rep = slides[cluster['rep']]
    others = [slides[k] for k in cluster['slides'] if k != cluster['rep']]
    decisions = {}
    t0 = pick_title(rep['shapes'])
    for s in rep['shapes']:
        if not s['text'] or s['kind'] == 'group':
            continue
        d = dict(id=s['id'], fixed=False, conf=0.9, why='')
        if s['ph'] in SKIP_PH:
            d.update(role='keep_' + s['ph'].lower(), fixed=True, why='서식 기본 플레이스홀더')
        elif is_pagenum(s, H) or s.get('fld_num'):
            d.update(role='sldNum', why='페이지 번호')
        else:
            seen, same = 0, 0
            for o in others:
                m = next((t for t in o['shapes'] if t['text'] and _match(s, t, W, H)), None)
                if m:
                    seen += 1
                    same += int(m['text'] == s['text'])
            deck_rep = len(cnt.get((_gkey(s, W, H), s['img'], s['text']), ())) if deck_like else 0
            sample_like = bool(re.fullmatch(r'(.)\1{2,}(\s+.*)?', s['text'])) or '필요 시' in s['text']
            if DATE_RE.search(s['text']):
                d.update(fixed=False, why='날짜는 항상 가변')
            elif seen >= 1 and same == seen:
                d.update(fixed=True, why=f'같은 유형 {seen + 1}장에서 문구가 같음')
            elif seen >= 1:
                d.update(fixed=False, why=f'같은 유형 {seen + 1}장에서 문구가 다름')
            elif deck_rep >= 2:
                d.update(fixed=True, why=f'전체 {deck_rep}장에서 같은 위치·문구 반복')
            elif s is t0 or sample_like:
                d.update(fixed=False, conf=0.8, why='제목/샘플 문구는 가변')
            elif LABEL_RE.match(s['text']) and not any(ch.isdigit() for ch in s['text']):
                d.update(fixed=True, conf=0.7, why='영문 대문자 라벨')
            elif s['y'] > 0.9 * H and len(s['text']) < 30 and (s['size'] or 12) <= 12:
                d.update(fixed=True, conf=0.6, why='하단 짧은 문구(푸터로 추정)')
            else:
                single = len(cluster['slides']) == 1 and not deck_like
                d.update(fixed=False, conf=0.8 if single else 0.65,
                         why='샘플 슬라이드의 글자는 가변' if single else '비교 대상이 없어 가변으로 추정')
            d['role'] = None
        decisions[s['id']] = d
    return decisions


def name_fields(ctype, rep, decisions, agenda, W, H):
    """가변 도형에 필드 이름(title, subtitle, body ...)을 붙인다. 반환: [{id,name,...}] (읽는 순서)"""
    items_ids = set()
    if agenda:
        for it in agenda['items']:
            for s in (it['num'], it['title'], it['org']):
                if s:
                    items_ids.add(s['id'])
    var = [s for s in rep['shapes'] if s['id'] in decisions and not decisions[s['id']]['fixed']
           and decisions[s['id']].get('role') != 'sldNum' and s['id'] not in items_ids]
    var.sort(key=lambda s: (round(s['y'] / (H * 0.03)), s['x']))
    fields = []
    title = pick_title(var)
    taken = set()
    if title is not None:
        fields.append(dict(s=title, name='title'))
        taken.add(title['id'])
    rest = [s for s in var if s['id'] not in taken]
    n_extra = 0
    for s in rest:
        if ctype == 'cover':
            pure_date = bool(DATE_RE.search(s['text'])) and len(DATE_RE.sub('', s['text']).strip(' .')) == 0
            has_date = bool(DATE_RE.search(s['text']))
            if pure_date and not any(f['name'] == 'date' for f in fields):
                name = 'date'
            elif has_date and not any(f['name'] == 'date_org' for f in fields):
                name = 'date_org'
            elif not has_date and not any(f['name'] == 'subtitle' for f in fields):
                name = 'subtitle'
            else:
                n_extra += 1
                name = f'extra{n_extra}'
        elif ctype == 'section':
            if PART_RE.match(s['text']) or (LABEL_RE.match(s['text'].upper()) and any(c.isdigit() for c in s['text'])):
                name = 'part'
            elif not any(f['name'] == 'subtitle' for f in fields):
                name = 'subtitle'
            else:
                n_extra += 1
                name = f'extra{n_extra}'
        elif ctype == 'closing':
            if DATE_RE.search(s['text']):
                name = 'date'
            else:
                n_extra += 1
                name = f'extra{n_extra}'
        else:   # content / agenda
            big = s['w'] >= 0.25 * W and s['h'] >= 0.12 * H
            if big and not any(f['name'] == 'body' for f in fields) and ctype == 'content':
                name = 'body'
            elif ctype == 'content' and big and not any(f['name'] == 'body2' for f in fields):
                name = 'body2'
            else:
                n_extra += 1
                name = f'extra{n_extra}'
        fields.append(dict(s=s, name=name))
    # 구분 슬라이드: 제목이 PART 라벨로 잘못 잡힌 경우 보정
    if ctype == 'section' and title is not None and PART_RE.match(title['text']):
        names = [f for f in fields if f['name'] == 'part']
        if names:
            fields_by = {f['name']: f for f in fields}
            others = [f for f in fields if f['name'] in ('subtitle',)]
            if others:
                fields_by['title']['name'], others[0]['name'] = 'part', 'title'
    if ctype == 'closing':
        for f in fields:
            if f['name'] == 'title':
                f['name'] = 'text'
    return fields


# ---------------- 전체 분석 ----------------
def analyze(path):
    prs = Presentation(path)
    W, H = int(prs.slide_width), int(prs.slide_height)
    slides = [slide_info(s, i, W, H) for i, s in enumerate(prs.slides)]
    n = len(slides)
    deck_like = n >= 8
    cnt = repeat_map(slides, W, H)

    # 1) 유형 분류
    for si in slides:
        t, conf, ev = classify(si, si['i'], n, W, H)
        si['type'], si['conf'], si['evidence'] = t, conf, ev
        if t == 'agenda' and analyze_agenda(si, W, H) is None:      # 번호 격자가 없는 목차는 일반 본문으로 처리
            t = si['type'] = 'content'
            si['evidence'] = si['evidence'] + ['번호 격자 없음 → 본문으로 처리']
        si['subkind'] = content_subkind(si, W, H) if t == 'content' else ''
        si['decor'] = decor_sig(si, cnt, deck_like, W, H)

    # 2) 클러스터(같은 디자인 묶음)
    groups = defaultdict(list)
    for si in slides:
        t = si['type']
        if t in ('content', 'section'):
            sub = 'two-col' if si['subkind'] == 'two-col' else 'std'
            key = (t, sub, si['decor'])
        else:
            key = (t, '', None)
        groups[key].append(si['i'])
    clusters = []
    for key, idxs in groups.items():
        t = key[0]
        # 대표 슬라이드
        def typ(k):
            si = slides[k]
            body = any(s['text'] and s['w'] >= 0.25 * W and s['h'] >= 0.12 * H and s['ph'] not in SKIP_PH
                       for s in si['shapes'])
            n_text = len(texts_of(si))
            if t == 'agenda':
                ag = analyze_agenda(si, W, H)
                return (len(ag['items']) if ag else 0, n_text, -k)
            return (int(body), n_text, -k)
        rep = max(idxs, key=typ)
        clusters.append(dict(type=t, key=[key[0], key[1]], slides=idxs, rep=rep))
    order = {'cover': 0, 'agenda': 1, 'section': 2, 'content': 3, 'closing': 4}
    clusters.sort(key=lambda c: (order[c['type']], c['rep']))

    # 3) 역할 결정
    for c in clusters:
        rep = slides[c['rep']]
        c['decisions'] = decide_roles(c, slides, cnt, deck_like, W, H)
        c['agenda'] = analyze_agenda(rep, W, H) if c['type'] == 'agenda' else None
        if c['agenda']:
            for it in c['agenda']['items']:
                for key, fixed in (('num', True), ('title', False), ('org', False)):
                    sh = it[key]
                    if sh and sh['id'] in c['decisions']:
                        c['decisions'][sh['id']].update(fixed=fixed, conf=0.85, why='목차 항목 격자')
        c['fields'] = name_fields(c['type'], rep, c['decisions'], c['agenda'], W, H)
        c['conf'] = round(sum(slides[k]['conf'] for k in c['slides']) / len(c['slides']), 2)
        c['decor_ids'] = [s['id'] for s in rep['shapes'] if is_decor_in(s, c['slides'], slides, cnt, W, H)]
    return dict(path=path, W=W, H=H, n=n, deck_like=deck_like, slides=slides, clusters=clusters, cnt=cnt)


# ---------------- 복잡도·추천 ----------------
def metrics(a):
    slides, W, H = a['slides'], a['W'], a['H']
    text_shapes = [s for si in slides for s in si['shapes'] if s['text'] and s['ph'] not in SKIP_PH]
    ph_text = [s for s in text_shapes if s['ph']]
    grp_text = sum(1 for si in slides for s in si['shapes'] if s['kind'] == 'group' and s['gtext'])
    odd = sum(1 for si in slides for s in si['shapes'] if s['kind'] == 'frame' or abs(s['rot']) > 1)
    uncertain = 0
    total_fields = 0
    for c in a['clusters']:
        for f in c['fields']:
            total_fields += 1
            if c['decisions'][f['s']['id']]['conf'] < 0.7:
                uncertain += 1
    rep_content = [a['slides'][c['rep']] for c in a['clusters'] if c['type'] == 'content']
    dense = max([len([s for s in si['shapes'] if not s['text']]) for si in rep_content] or [0])
    agenda = next((c for c in a['clusters'] if c['type'] == 'agenda'), None)
    return dict(
        n_slides=a['n'], n_clusters=len(a['clusters']),
        conf_avg=round(sum(si['conf'] for si in slides) / max(1, len(slides)), 2),
        ph_ratio=round(len(ph_text) / max(1, len(text_shapes)), 2),
        group_text=grp_text, odd_objects=odd, uncertain_fields=uncertain, total_fields=total_fields,
        dense_decor=dense, agenda_grid=bool(agenda and agenda['agenda'] and len(agenda['agenda']['items']) >= 2),
        n_content=sum(len(c['slides']) for c in a['clusters'] if c['type'] == 'content'),
        types=sorted({c['type'] for c in a['clusters']}),
    )


def recommend(a, intent='unknown'):
    """플레이스홀더 변환 적합도 S(0~100)와 근거. intent: once | repeat | edit | unknown"""
    m = metrics(a)
    S = 50
    plus, minus = [], []

    def add(v, msg):
        nonlocal S
        S += v
        (plus if v > 0 else minus).append(f'{v:+d} {msg}')

    if m['ph_ratio'] >= 0.6:
        return dict(mode='as-is', score=95, metrics=m, plus=[f'+45 글자 도형의 {int(m["ph_ratio"] * 100)}%가 이미 플레이스홀더'],
                    minus=[], decision_required=False, intent=intent,
                    summary='이미 플레이스홀더 서식입니다. 변환 없이 roles.json만 만들면 됩니다.')
    if m['agenda_grid']:
        add(10, '목차에 항목 격자가 있어 항목 수별 레이아웃(1~N개)을 만들 수 있음')
    if a['deck_like'] and m['n_content'] >= 3:
        add(10, f'같은 유형 본문 슬라이드가 {m["n_content"]}장이라 레이아웃으로 만들 가치가 큼')
    if m['conf_avg'] >= 0.8:
        add(5, f'슬라이드 유형 분류 신뢰도 높음({m["conf_avg"]})')
    if m['n_clusters'] >= 4:
        add(5, f'서로 다른 디자인 {m["n_clusters"]}종 → 레이아웃 선택 기능 활용')
    if intent in ('repeat', 'edit'):
        add(15, '여러 번 쓰거나 만든 뒤 PowerPoint에서 직접 편집할 예정')
    if intent == 'once':
        add(-25, '한 번만 쓸 예정(변환 비용 대비 이득 적음)')
    if m['conf_avg'] < 0.65:
        add(-15, f'슬라이드 유형 분류 신뢰도 낮음({m["conf_avg"]})')
    if m['total_fields'] and m['uncertain_fields'] / m['total_fields'] > 0.3:
        add(-10, f'고정/가변을 확신하지 못한 글자 도형 {m["uncertain_fields"]}/{m["total_fields"]}개')
    if m['group_text']:
        add(-min(20, 7 * m['group_text']), f'그룹 안에 글자가 든 도형 {m["group_text"]}개(플레이스홀더로 못 바꿈)')
    if m['odd_objects']:
        add(-min(15, 5 * m['odd_objects']), f'회전 텍스트·OLE/SmartArt 등 특수 개체 {m["odd_objects"]}개')
    if m['dense_decor'] >= 14:
        add(-10, f'장식 도형이 매우 많은 슬라이드(도형 {m["dense_decor"]}개) → 변환 시 모양이 틀어질 위험')
    if a['n'] <= 2:
        add(-15, '슬라이드가 1~2장뿐이라 유형 분류 근거가 부족')
    S = max(0, min(100, S))
    if S >= 65:
        mode, lvl = 'placeholder', '높음'
    elif S >= 45:
        mode, lvl = ('placeholder' if S >= 55 else 'duplicate'), '보통'
    else:
        mode, lvl = 'duplicate', '높음' if S < 35 else '보통'
    return dict(mode=mode, score=S, level=lvl, metrics=m, plus=plus, minus=minus, intent=intent,
                decision_required=True,
                summary='플레이스홀더 변환 권장' if mode == 'placeholder' else '복제 방식 권장')
