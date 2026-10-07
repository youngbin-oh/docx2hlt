#!/usr/bin/env python3
"""docx2hlt.py — convert a .docx file to NKEditor .hlt

Usage: python docx2hlt.py input.docx output.hlt

Parses the docx (zip+XML, no dependencies), emits a binary job file,
then drives TerConv.exe (32-bit helper that calls hwplt32.dll) to build
the .hlt document.
"""
import io
import os
import re
import struct
import subprocess
import sys
import time
import tempfile
import zipfile
from xml.etree import ElementTree as ET

try:
    from PIL import Image
    HAVE_PIL = True
except ImportError:
    HAVE_PIL = False

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
V = '{urn:schemas-microsoft-com:vml}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

HERE = os.path.dirname(os.path.abspath(__file__))
TERCONV = os.path.join(HERE, 'TerConv.exe')
NK_DIR = r'C:\KipoNet\NKEditor'

MAX_CHUNK = 290          # bytes per TerInsertLine call (limit is 299)
STYLE_BOLD = 1
STYLE_ITALIC = 2
STYLE_UNDERLINE = 4
STYLE_STRIKE = 8
STYLE_SUB = 0x20         # TE char-style bits (mutually exclusive pair w/ super)
STYLE_SUPER = 0x10

# TE paragraph alignment flag bits (mutually exclusive group in para+0x18)
# 0x800 is the stock default (left). 0x1/0x2/0x400 -> center/right/justify,
# verified empirically below (PALIGN_MAP).
PALIGN_LEFT = 0x800
PALIGN_CENTER = 0x1
PALIGN_RIGHT = 0x2
PALIGN_JUSTIFY = 0x400


def _on(el):
    """True unless the on/off element is explicitly disabled."""
    return el is not None and el.get(W + 'val') not in ('0', 'false', 'off', 'none')


def run_props(run, base=None, styles=None):
    """Return (styleMask, sizeHalfPt, colorBGR) for a w:r run element.
    base: style-inherited (mask, sz, color) tuple to start from.
    A character style (w:rStyle) merges between base and direct rPr."""
    mask, sz, color = (base if base else (0, -1, -1))
    rpr = run.find(W + 'rPr')
    rst = None
    if rpr is not None and styles:
        rs = rpr.find(W + 'rStyle')
        if rs is not None:
            rst = styles.get(rs.get(W + 'val'))
    for src in ((rst[0] if rst else None), rpr):
        if src is None:
            continue
        om, oc, osz, ocol = (src if isinstance(src, tuple)
                             else rpr_props(src))
        mask = (mask & ~oc) | om          # explicit-off bits clear inherited
        if osz != -1:
            sz = osz
        if ocol != -1:
            color = ocol
    return mask, sz, color


def para_props(p, styles):
    """Paragraph attrs for a w:p: returns dict or None.
    Keys: left,right,first (twips; hanging -> negative first), jc (TE flag bits),
    before,after (twips). Merges paragraph-style pPr then direct pPr."""
    ppr = p.find(W + 'pPr')
    st = None
    if ppr is not None:
        ps = ppr.find(W + 'pStyle')
        if ps is not None:
            st = styles.get(ps.get(W + 'val'))
    out = {}
    for src in ((st[1] if st else None), ppr):
        if src is None:
            continue
        ind = src.find(W + 'ind')
        if ind is not None:
            def tw(name):
                v = ind.get(W + name)
                return int(v) if v and v.lstrip('-').isdigit() else None
            v = tw('left') if tw('left') is not None else tw('start')
            if v is not None:
                out['left'] = v
            v = tw('right') if tw('right') is not None else tw('end')
            if v is not None:
                out['right'] = v
            if tw('firstLine') is not None:
                out['first'] = tw('firstLine')
            if tw('hanging') is not None:
                out['first'] = -tw('hanging')
        jc = src.find(W + 'jc')
        if jc is not None:
            v = jc.get(W + 'val')
            out['jc'] = {'center': PALIGN_CENTER, 'right': PALIGN_RIGHT,
                         'end': PALIGN_RIGHT, 'both': PALIGN_JUSTIFY,
                         'distribute': PALIGN_JUSTIFY}.get(v, PALIGN_LEFT)
        sp = src.find(W + 'spacing')
        if sp is not None:
            for k, dst in (('before', 'before'), ('after', 'after')):
                v = sp.get(W + k)
                if v and v.isdigit():
                    out[dst] = int(v)
    return out or None


def _merge_props(base, over):
    """Cascade: `over` overrides `base` per field (-1 = inherit base's).
    Tuples are (mask, clearMask, sz, color) — clearMask removes inherited
    bits (e.g. <w:b w:val="0"/> beats a bold base style)."""
    if not over:
        return base
    mask, clr, sz, color = base
    om, oc, osz, ocol = over
    return (((mask & ~oc) | om), ((clr | oc) & ~om),
            osz if osz != -1 else sz,
            ocol if ocol != -1 else color)


def para_style_fmt(p, styles):
    """style-inherited (mask, sz, color) base for runs of this paragraph.
    docDefaults supplies the lowest-priority values."""
    dd = styles.get('\x00docdef')
    base = dd[0] if dd else (0, 0, -1, -1)
    ppr = p.find(W + 'pPr')
    if ppr is not None:
        ps = ppr.find(W + 'pStyle')
        if ps is not None:
            st = styles.get(ps.get(W + 'val'))
            if st:
                base = _merge_props(base, st[0])
    return (base[0], base[2], base[3])


def rpr_props(rpr):
    """(mask, clearMask, sizeHalfPt, colorBGR) from a w:rPr element
    (-1 = unspecified; clearMask = bits explicitly turned off)."""
    mask = clr = 0
    sz = color = -1
    if rpr is None:
        return mask, clr, sz, color
    for tag, bit in ((W + 'b', STYLE_BOLD), (W + 'i', STYLE_ITALIC),
                     (W + 'strike', STYLE_STRIKE), (W + 'dstrike', STYLE_STRIKE)):
        el = rpr.find(tag)
        if el is not None:
            if _on(el):
                mask |= bit
            else:
                clr |= bit
    u = rpr.find(W + 'u')
    if u is not None:
        if u.get(W + 'val') not in ('none', '0', 'false', 'off'):
            mask |= STYLE_UNDERLINE
        else:
            clr |= STYLE_UNDERLINE
    va = rpr.find(W + 'vertAlign')
    if va is not None:
        v = va.get(W + 'val')
        if v == 'subscript':
            mask, clr = (mask | STYLE_SUB) & ~STYLE_SUPER, clr | STYLE_SUPER
        elif v == 'superscript':
            mask, clr = (mask | STYLE_SUPER) & ~STYLE_SUB, clr | STYLE_SUB
        else:
            clr |= STYLE_SUB | STYLE_SUPER    # 'baseline' clears both
    szel = rpr.find(W + 'sz')
    if szel is not None:
        try:
            sz = int(szel.get(W + 'val'))
        except (TypeError, ValueError):
            pass
    cel = rpr.find(W + 'color')
    if cel is not None:
        cv = cel.get(W + 'val')
        if cv and cv != 'auto':
            try:
                rgb = int(cv, 16)
                color = (rgb & 0xFF) << 16 | (rgb & 0xFF00) | (rgb >> 16) & 0xFF
            except ValueError:
                pass
    return mask, clr, sz, color


def load_styles(zf):
    """styleId -> (runProps tuple, pPr element) from styles.xml."""
    out = {}
    try:
        root = ET.fromstring(zf.read('word/styles.xml'))
    except KeyError:
        return out
    raw = {}  # sid -> (props, basedOnId, pPr)
    for st in root.findall(W + 'style'):
        sid = st.get(W + 'styleId')
        if sid is None:
            continue
        bo = st.find(W + 'basedOn')
        raw[sid] = (rpr_props(st.find(W + 'rPr')),
                    bo.get(W + 'val') if bo is not None else None,
                    st.find(W + 'pPr'))

    # Resolve w:basedOn chains so each style's effective props include inherited
    # values (e.g. a body style that only overrides some attributes).
    def resolved(sid, seen=()):
        if sid in seen or sid not in raw:
            return (0, 0, -1, -1)
        props, boid, _ = raw[sid]
        base = resolved(boid, seen + (sid,)) if boid else (0, 0, -1, -1)
        return _merge_props(base, props)

    for sid, (_, _, ppr) in raw.items():
        out[sid] = (resolved(sid), ppr)
    # docDefaults: the document-wide run-property base (default font size etc.)
    dd = root.find(W + 'docDefaults')
    if dd is not None:
        rpr = dd.find('.//' + W + 'rPrDefault/' + W + 'rPr')
        if rpr is None:
            rpr = dd.find('.//' + W + 'rPr')
        if rpr is not None:
            out['\x00docdef'] = (rpr_props(rpr), None)
    return out


def load_rels(zf):
    """rId -> media target (e.g. 'media/image1.tif') from document rels."""
    out = {}
    try:
        root = ET.fromstring(zf.read('word/_rels/document.xml.rels'))
    except KeyError:
        return out
    for rel in root:
        rid = rel.get('Id')
        tgt = rel.get('Target', '')
        typ = rel.get('Type', '')
        if rid and 'image' in typ.lower():
            if not tgt.startswith('word/'):
                tgt = 'word/' + tgt.lstrip('/')
            out[rid] = tgt
    return out


def img_px_dims(data):
    """Return (width_px, height_px, dpi) from image file bytes, or None."""
    try:
        # PNG
        if data[:8] == b'\x89PNG\r\n\x1a\n':
            w, h = struct.unpack('>II', data[16:24])
            dpi = None
            # pHYs chunk
            i = 8
            while i + 8 < len(data):
                ln, typ = struct.unpack('>I4s', data[i:i + 8])
                if typ == b'pHYs' and ln >= 9:
                    xppm, yppm, unit = struct.unpack('>IIB', data[i + 8:i + 17])
                    if unit == 1 and xppm:
                        dpi = int(xppm * 0.0254)
                    break
                i += 8 + ln + 4
            return w, h, dpi
        # GIF
        if data[:6] in (b'GIF87a', b'GIF89a'):
            w, h = struct.unpack('<HH', data[6:10])
            return w, h, None
        # BMP
        if data[:2] == b'BM':
            w, h = struct.unpack('<ii', data[18:26])
            xppm = struct.unpack('<i', data[38:42])[0]
            dpi = int(xppm * 0.0254) if xppm else None
            return abs(w), abs(h), dpi
        # JPEG — parse SOF
        if data[:2] == b'\xff\xd8':
            i = 2
            while i + 9 < len(data):
                if data[i] == 0xff and data[i + 1] in (0xc0, 0xc1, 0xc2, 0xc5, 0xc6, 0xc7):
                    h, w = struct.unpack('>HH', data[i + 5:i + 9])
                    return w, h, None
                if data[i] == 0xff and data[i + 1] not in (0xff, 0x00) and data[i + 1] != 0xd8:
                    ln = struct.unpack('>H', data[i + 2:i + 4])[0]
                    i += 2 + ln
                else:
                    i += 1
            return None
        # TIFF (II / MM)
        if data[:2] in (b'II', b'MM'):
            fmt = '<' if data[:2] == b'II' else '>'
            ifd = struct.unpack(fmt + 'I', data[4:8])[0]
            n = struct.unpack(fmt + 'H', data[ifd:ifd + 2])[0]
            w = h = None
            xd = yd = None
            for i in range(n):
                e = ifd + 2 + i * 12
                tag, typ, cnt, val = struct.unpack(fmt + 'HHII', data[e:e + 12])
                if tag == 256:
                    w = val
                elif tag == 257:
                    h = val
                elif tag == 282:
                    xd = struct.unpack(fmt + 'II', data[val:val + 8])
                elif tag == 283:
                    yd = struct.unpack(fmt + 'II', data[val:val + 8])
            dpi = None
            if xd and xd[1]:
                dpi = xd[0] // xd[1]
            if w and h:
                return w, h, dpi
    except Exception:
        pass
    return None


# Subtrees whose content is not part of the visible document text:
# tracked deletions / move sources and property blocks.
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
SKIP_SUBTREES = (W + 'del', W + 'moveFrom', W + 'pPr', W + 'rPr',
                 M + 'ctrlPr')
# mc:AlternateContent carries the same content twice (Choice + Fallback);
# reading both duplicates text-box text and images. Per the OOXML
# markup-compatibility rule we take the first Choice whose Requires prefixes
# we understand — DrawingML shapes/groups/canvases (their w:drawing holds the
# picture or text box) — and otherwise the Fallback (e.g. w16se:symEx emoji
# symbols, whose Fallback is the plain w:t character).
MC_UNDERSTOOD = {'wps', 'wpg', 'wpc', 'wpi', 'wp14', 'w14', 'a14'}


def mc_select(ac):
    """The branch of an mc:AlternateContent to read (or None)."""
    for alt in ac:
        if alt.tag == MC + 'Choice':
            req = (alt.get('Requires') or '').split()
            if req and all(r in MC_UNDERSTOOD for r in req):
                return alt
        elif alt.tag == MC + 'Fallback':
            return alt
    return None


def _find_skip_txbx(node, tag):
    """First descendant `tag` of node, not descending into text-box content
    (a text box's own pictures are handled when its paragraphs are walked)."""
    for ch in node:
        if ch.tag == W + 'txbxContent' or ch.tag in SKIP_SUBTREES:
            continue
        if ch.tag == MC + 'AlternateContent':
            ch = mc_select(ch)
            if ch is None:
                continue
        if ch.tag == tag:
            return ch
        hit = _find_skip_txbx(ch, tag)
        if hit is not None:
            return hit
    return None


def _math_text(node):
    """Linear text of an m:oMath, without tracked-deleted runs."""
    out = []
    for ch in node:
        if ch.tag in SKIP_SUBTREES:
            continue
        if ch.tag == M + 't':
            out.append(ch.text or '')
        out.append(_math_text(ch))
    return ''.join(out)


def main_story_omaths(root):
    """m:oMath elements in the order Word's doc.OMaths enumerates them:
    document order, main story only (text boxes and VML fallbacks are other
    stories / duplicates), tracked-deleted equations included. EqRender keys
    its eqNNNN.tif files by this ordinal."""
    out = []

    def walk(node):
        for ch in node:
            if ch.tag == W + 'txbxContent':
                continue
            if ch.tag == MC + 'AlternateContent':
                ch = mc_select(ch)
                if ch is None:
                    continue
            if ch.tag == M + 'oMath':
                out.append(ch)
            else:
                walk(ch)
    walk(root)
    return out


def para_runs(p, styles, rels, media, eqs=None):
    """Extract (runs, images, paraProps) from a w:p paragraph element.
    runs: [(text, mask, sizeHalfPt, colorBGR)];
    images: [(mediaTarget, cx, cy, pos)] — pos = text offset where the
    image sits inside the paragraph (for inline emission).
    OMML equations (m:oMath): `eqs` = (ordinal_by_element, {ordinal:
    (tif, w_mm, h_mm)}); a rendered equation is spliced in at its text
    offset, any other degrades to linear m:t text (submission rules only
    forbid non-image content in 식별항목, which we never emit, so plain text
    stays legal).
    The paragraph is walked as a tree (not p.iter()) so that runs nested in
    hyperlinks/fields/sdt are read exactly once, tracked deletions are
    skipped, and a text box's content is read once (Choice, not Fallback)."""
    base = para_style_fmt(p, styles)
    pp = para_props(p, styles)
    runs = []
    images = []
    nchar = [0]
    eq_ord, eq_img = eqs if eqs else ({}, {})

    def add_text(text, mask, sz, color, ctx=None):
        if text:
            text = sanitize_text(text, CHAR_REPORT,
                                 ctx if ctx is not None else text[:40])
            if text:
                runs.append((text, mask, sz, color))
                nchar[0] += len(text)

    def add_image(node):
        # embedded image rel id and display extent of a w:drawing / w:pict
        rid = None
        cx = cy = None
        ext = _find_skip_txbx(node, WP + 'extent')
        if ext is not None:
            cx = int(ext.get('cx', '0'))
            cy = int(ext.get('cy', '0'))
        blip = _find_skip_txbx(node, A + 'blip')
        if blip is not None:
            rid = blip.get(R + 'embed') or blip.get(R + 'link')
        if rid is None:
            imd = _find_skip_txbx(node, V + 'imagedata')
            if imd is not None:
                rid = imd.get(R + 'id')
        if rid and rid in rels:
            tgt = rels[rid]
            if tgt in media:
                images.append((tgt, cx, cy, nchar[0]))

    def do_math(node):
        n = eq_ord.get(id(node))
        hit = eq_img.get(n) if n is not None else None
        if hit:
            # Word equation rendered by EqRender — registered under a
            # '@eq<N>' pseudo target
            tif, w_mm, h_mm = hit
            key = '@eq%d' % n
            media[key] = tif
            images.append((key, int(round(w_mm * 36000)),
                           int(round(h_mm * 36000)), nchar[0]))
        else:
            text = _math_text(node)
            add_text(text, 0, base[1], base[2], '(수식) ' + text[:37])

    def do_run(r):
        mask, sz, color = run_props(r, base, styles)
        buf = []

        def flush():
            add_text(''.join(buf), mask, sz, color)
            del buf[:]
        run_children(r, buf, flush)
        flush()

    def run_children(r, buf, flush):
        for ch in r:
            tag = ch.tag
            if tag in SKIP_SUBTREES:
                continue
            if tag == W + 't':
                buf.append(ch.text or '')
            elif tag == W + 'tab':
                buf.append('\t')
            elif tag in (W + 'br', W + 'cr'):
                buf.append(' ')
            elif tag == W + 'noBreakHyphen':
                buf.append('-')
            elif tag in (W + 'drawing', W + 'pict'):
                flush()
                add_image(ch)
                walk(ch)                # text-box paragraphs, if any
            elif tag == MC + 'AlternateContent':
                alt = mc_select(ch)
                if alt is not None:
                    run_children(alt, buf, flush)
            elif tag == M + 'oMath':
                flush()
                do_math(ch)

    def walk(node):
        for ch in node:
            tag = ch.tag
            if tag in SKIP_SUBTREES:
                continue
            if tag == W + 'r':
                do_run(ch)
            elif tag == MC + 'AlternateContent':
                alt = mc_select(ch)
                if alt is not None:
                    walk(alt)
            elif tag == M + 'oMath':
                do_math(ch)
            elif tag == W + 'p' and nchar[0]:
                # paragraph inside a text box: keep it word-separated from
                # the surrounding text
                add_text(' ', 0, base[1], base[2])
                walk(ch)
            else:
                walk(ch)

    walk(p)
    return runs, images, pp


def iter_paragraphs(parent, styles, rels, media, eqs=None,
                    tbls=None, tstate=None, depth=0):
    """Yield (run-list, image-list, paraProps) from body.

    A top-level w:tbl becomes an image-only paragraph holding its
    EqRender TIFF (spliced under a '@tbl<N>' pseudo target, like '@eq'),
    preserving the table's visual grid; when no rendered image exists for
    that ordinal — or the table is nested inside a cell — the old
    flatten-to-cell-paragraphs behaviour applies."""
    for child in parent:
        if child.tag == W + 'p':
            yield para_runs(child, styles, rels, media, eqs)
        elif child.tag == W + 'sdt':
            # structured-document-tag is a container — its sdtContent holds
            # block-level w:p/w:tbl at the same depth (also fixes sdt
            # content being silently dropped and table ordinals drifting)
            content = child.find(W + 'sdtContent')
            if content is not None:
                yield from iter_paragraphs(content, styles, rels, media,
                                           eqs, tbls, tstate, depth)
        elif child.tag == W + 'tbl':
            hit = None
            if depth == 0 and tbls is not None and tstate is not None:
                tstate[0] += 1
                hit = tbls.get(tstate[0])
            if hit:
                tif, w_mm, h_mm = hit
                key = '@tbl%d' % tstate[0]
                media[key] = tif
                yield ([], [(key, int(round(w_mm * 36000)),
                             int(round(h_mm * 36000)), 0)], {})
                continue
            for tr in child.findall(W + 'tr'):
                for tc in tr.findall(W + 'tc'):
                    yield from iter_paragraphs(tc, styles, rels, media, eqs,
                                               tbls, tstate, depth + 1)


# --- KIPO structured-document support -------------------------------------
# A real e-filing .hlt is a structured (SGML-tagged) document: the section
# headings are element boundaries that feed the structure tree and the
# 제출파일 export, not just styled text. The element set lives in the .frm
# skeleton (we use Patentd.frm). DOCX patent specs already mark sections with
# 【…】 headings; we map each onto the skeleton element whose marker matches.
#
# Keys below are the docx heading text with 【】 and whitespace stripped;
# values are the exact marker text the skeleton uses (spaces preserved).
ELEM_ALIASES = {
    '요약서': '요약서', '요약': '요약', '초록': '요약',
    '대표도': '대표도', '색인어': '색인어', '명세서': '명세서',
    '발명의명칭': '발명의 명칭', '발명의명세': '발명의 명칭',
    '발명의상세한설명': '발명의 상세한 설명', '발명의설명': '발명의 상세한 설명',
    '기술분야': '기술분야', '기술분야및배경기술': '기술분야',
    '배경기술': '배경기술', '발명의배경이되는기술': '배경기술',
    '배경이되는기술': '배경기술', '종래기술': '배경기술', '종래의기술': '배경기술',
    '종래기술설명': '배경기술', '선행기술': '배경기술',
    '발명의내용': '발명의 내용',
    '발명의목적': '해결하고자 하는 과제', '발명의과제': '해결하고자 하는 과제',
    '발명이해결하고자하는과제': '해결하고자 하는 과제',
    '발명이이루고자하는기술적과제': '해결하고자 하는 과제',
    '해결하고자하는과제': '해결하고자 하는 과제', '해결하려는과제': '해결하고자 하는 과제',
    '해결하고자하는기술적과제': '해결하고자 하는 과제', '기술적과제': '해결하고자 하는 과제',
    '과제의해결수단': '과제 해결 수단', '과제해결수단': '과제 해결 수단',
    '기술적해결방법': '과제 해결 수단', '발명의구성': '과제 해결 수단',
    '발명을달성하기위한수단': '과제 해결 수단',
    '발명의효과': '효과', '효과': '효과', '유리한효과': '효과', '발명의작용효과': '효과',
    '발명을실시하기위한구체적인내용': '발명의 실시를 위한 구체적인 내용',
    '발명의실시를위한구체적인내용': '발명의 실시를 위한 구체적인 내용',
    '발명을실시하기위한형태': '발명의 실시를 위한 구체적인 내용',
    '발명의실시를위한형태': '발명의 실시를 위한 구체적인 내용',
    '발명을실시하기위한방법': '발명의 실시를 위한 구체적인 내용',
    '구체적인실시예': '발명의 실시를 위한 구체적인 내용',
    '도면의간단한설명': '도면의 간단한 설명', '간단한도면설명': '도면의 간단한 설명',
    '도면설명': '도면의 간단한 설명', '부호의설명': '도면의 간단한 설명',
    '특허청구범위': '특허청구범위', '청구범위': '특허청구범위', '청구의범위': '특허청구범위',
    '청구항': '청구항 1',
    '도면': '도면', '서열목록': '서열목록', '서열목록프리텍스트': '서열목록',
}

# Element names whose docx heading carries a number that must be preserved in
# the body (청구항 1..N, 도 1..N): the skeleton has a single slot for them, so
# we keep the heading text as a content line rather than dropping it.
ELEM_KEEP_HEADING = ('청구항', '도 ')


def _norm_tag(s):
    return re.sub(r'\s+', '', s)


def canon_element(text):
    """If `text` is a patent section heading (【tag】 or a bare known element
    name), return (skeleton element display name, keep_heading). Else None."""
    t = text.strip()
    if not t:
        return None
    m = re.match(r'^【\s*(.+?)\s*】\s*$', t)
    if not m:
        return None                      # only 【…】 markers delimit sections
    inner = _norm_tag(m.group(1))
    # 청구항 N / 도 N carry an index -> map to the single skeleton slot but keep
    # the heading line so the number is preserved as visible content.
    mnum = re.match(r'^(청구항|도)\s*(\d+)', m.group(1))
    if mnum:
        base = mnum.group(1)
        tgt = '청구항 1' if base == '청구항' else '도 1'
        return (tgt, True)
    if inner in ELEM_ALIASES:
        return (ELEM_ALIASES[inner], False)
    return None


def clean_text(s):
    # strip chars that are control codes in the hlt text stream
    return re.sub(r'[\x00-\x08\x0b\x0e-\x1f]', '', s)


# --- KIPO-supported character set ------------------------------------------
# https://www.patent.go.kr/.../word.do lists every glyph the KIPO editor
# recognizes (KS X 1001 Hangul + 특수문자 + ASCII). Anything outside the set
# cannot survive the .hlt — it would turn into '?' at CP949 encode time. We
# substitute the closest supported glyph and record every change for review.
KIPO_CHARS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               'kipo_chars.txt')
_KIPO_CHARS = None


def kipo_chars():
    """The set of characters the KIPO editor can represent."""
    global _KIPO_CHARS
    if _KIPO_CHARS is None:
        try:
            with open(KIPO_CHARS_FILE, encoding='utf-8') as f:
                _KIPO_CHARS = set(f.read())
        except OSError:
            _KIPO_CHARS = None
    return _KIPO_CHARS


# Unsupported char -> ordered candidate replacements (first fully-supported
# wins). Multi-char strings are allowed ('‼'->'!!').
CHAR_SUBS = {
    '‖': ['∥'], '≈': ['≒', '≃'], '≅': ['≒'], '∼': ['∽', '~'], 'Ω': ['Ω'],
    '‐': ['-'], '‑': ['-'], '‒': ['-'], '–': ['-'], '—': ['―', '-'],
    '©': ['(C)'], '®': ['Ⓡ', '(R)'], '℗': ['(P)'],
    '°': ['˚'], '•': ['·'], '․': ['.'], '‥': ['..'],
    '‼': ['!!'], '⁉': ['!?'], '⁈': ['?!'], '⁇': ['??'],
    '㏒': ['log'], '㏑': ['ln'],
    '◦': ['ㅇ'], '∘': ['ㅇ'],
    'Ⓐ': ['(A)'], 'Ⓑ': ['(B)'], 'Ⓒ': ['(C)'], 'Ⓓ': ['(D)'],
    'µ': ['μ'],
    '’': ["'"], '‘': ["'"], 'ʼ': ["'"], '´': ["'"], '`': ["'"],
    '“': ['"'], '”': ['"'], '„': ['"'], '″': ['"'],
    '«': ['≪', '〈'], '»': ['≫', '〉'], '‹': ['〈'], '›': ['〉'],
    '〜': ['～'], '˜': ['~'], '﹣': ['-'], '−': ['-'],
    '⟦': ['〔'], '⟧': ['〕'], '⟪': ['≪'], '⟫': ['≫'],
    '∛': ['√'], '∜': ['√'], '∬': ['∫∫'], '∭': ['∫∫∫'], '∮': ['∫'],
    '⊕': ['(+)'], '⊖': ['(-)'], '⊗': ['(x)'], '⊙': ['(.)'], '⊘': ['(/)'],
    '⟹': ['⇒', '→'], '⇐': ['⇐', '←'], '⟸': ['⇐', '←'], '⇔': ['⇔', '↔'],
    '∆': ['△'], '∇': ['▽'],
    '⑪': ['(11)'], '⑫': ['(12)'], '⑬': ['(13)'], '⑭': ['(14)'],
    '⑮': ['(15)'], '⑯': ['(16)'], '⑰': ['(17)'], '⑱': ['(18)'],
    '⑲': ['(19)'], '⑳': ['(20)'],
    '⑴': ['(1)'], '⑵': ['(2)'], '⑶': ['(3)'], '⑷': ['(4)'], '⑸': ['(5)'],
    '⑹': ['(6)'], '⑺': ['(7)'], '⑻': ['(8)'], '⑼': ['(9)'], '⑽': ['(10)'],
    '♂': ['(남)'], '♀': ['(여)'],
    # invisible / zero-width junk that must never reach the editor
    ' ': [' '], '​': [''], '‌': [''], '‍': [''],
    '﻿': [''], '­': [''],
}


def sanitize_text(s, report=None, where=''):
    """Map every char not in the KIPO set to a supported substitute.
    Appends (orig, replacement, context) rows to `report` (or drops the char
    when nothing is mappable). Returns the cleaned string."""
    allowed = kipo_chars()
    if allowed is None:
        return s
    out = []
    for ch in s:
        if ch in allowed or ch in '\t':
            out.append(ch)
            continue
        repl = None
        for cand in CHAR_SUBS.get(ch, ()):
            if all(c in allowed for c in cand):
                repl = cand
                break
        if repl is None:
            repl = ''                       # unrepresentable: drop, but report
        if report is not None:
            report.append((ch, repl, where))
        out.append(repl)
    return ''.join(out)


CHAR_REPORT = []          # (orig, replacement, paragraph snippet) rows


def _w_text(s):
    return ('<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
            % (s.replace('&', '&amp;').replace('<', '&lt;')
               .replace('>', '&gt;')))


def write_charset_docx(path, warnings, rows):
    """Write the charset review report as a minimal Word document —
    a heading, warning lines, then a table of
    코드포인트 | 원본 | 치환 | 문맥."""
    trs = []
    hdr = ('<w:tr><w:trPr><w:tblHeader/></w:trPr>' +
           ''.join('<w:tc><w:p><w:r><w:rPr><w:b/></w:rPr>'
                   '<w:t>%s</w:t></w:r></w:p></w:tc>' % h
                   for h in ('코드포인트', '원본', '치환', '문맥')) +
           '</w:tr>')
    for orig, repl, ctx in rows:
        cells = ('U+%04X' % ord(orig), orig, repl or '(삭제)', ctx)
        trs.append('<w:tr>' +
                   ''.join('<w:tc><w:p>%s</w:p></w:tc>' % _w_text(c)
                           for c in cells) +
                   '</w:tr>')
    warn_xml = ''.join(
        '<w:p><w:r><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr>'
        '<w:t xml:space="preserve">!! %s</w:t></w:r></w:p>' % w
        for w in warnings)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/'
           'wordprocessingml/2006/main"><w:body>'
           '<w:p><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr>'
           '<w:t>KIPO 미지원 문자 — 치환/삭제 내역</w:t></w:r></w:p>'
           + warn_xml +
           '<w:p/>'
           '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
           '<w:tblBorders>'
           '<w:top w:val="single" w:sz="4"/>'
           '<w:left w:val="single" w:sz="4"/>'
           '<w:bottom w:val="single" w:sz="4"/>'
           '<w:right w:val="single" w:sz="4"/>'
           '<w:insideH w:val="single" w:sz="4"/>'
           '<w:insideV w:val="single" w:sz="4"/>'
           '</w:tblBorders></w:tblPr>'
           + hdr + ''.join(trs) +
           '</w:tbl>'
           '<w:sectPr/></w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
          'content-types">'
          '<Default Extension="rels" ContentType="application/vnd.'
          'openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/'
          'vnd.openxmlformats-officedocument.wordprocessingml.document.'
          'main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/'
            'package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/></Relationships>')
    wrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/'
             'package/2006/relationships"/>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', wrels)
        z.writestr('word/document.xml', doc)


def cp949(s):
    return clean_text(s).encode('cp949', errors='replace')


# --- PatentCAFDOC XML emission ---------------------------------------------
# NKEditor natively imports KIPO PatentCAFDOC XML (its own e-filing interchange
# format) and builds a fully-structured patent document — element tree, claims,
# summary, drawings and all. That is far more robust than manufacturing the
# binary .hlt element table by hand, so the converter emits this XML and lets
# NKEditor do the structured build, then saves the result as .hlt.
#
# Schema (from a real KIPO submission package, Epasl/ManageSample/*.mof):
#   <KIPO><PatentCAFDOC>
#     <description> invention-title | technical-field | background-art |
#       summary-of-invention{tech-problem,tech-solution,advantageous-effects} |
#       description-of-drawings | description-of-embodiments </description>
#     <claims><claim num><claim-text>…</claim-text></claim>…</claims>
#     <abstract><summary><p num>…</p></summary>
#       <abstract-figure><p num><figref num/></p></abstract-figure></abstract>
#     <drawings><figure num><img file he wi/></figure>…</drawings>
#   </PatentCAFDOC></KIPO>
# Body text uses <p num="0001">…</p>; inline markup is <sub><sup><u><br/>.

# normalized 【heading】 -> the XML section bucket that receives its content.
HEAD2XML = {
    '발명의설명': '_root', '명세서': '_root',
    '발명의명칭': 'title', '발명의명세': 'title', '발명의이름': 'title',
    '기술분야': 'techfield', '기술분야및배경기술': 'techfield',
    '배경기술': 'background', '발명의배경이되는기술': 'background',
    '배경이되는기술': 'background', '종래기술': 'background',
    '종래의기술': 'background', '종래기술설명': 'background',
    '선행기술': 'background', '배경기술및선행기술': 'background',
    '발명의내용': 'problem',                      # container -> problem bucket
    '해결하고자하는과제': 'problem', '발명의목적': 'problem',
    '발명의과제': 'problem', '기술적과제': 'problem',
    '해결하려는과제': 'problem', '발명이해결하고자하는과제': 'problem',
    '발명이해결하고자하는기술적과제': 'problem',
    '발명이이루고자하는기술적과제': 'problem',
    '해결하고자하는기술적과제': 'problem', '발명의기술적과제': 'problem',
    '과제의해결수단': 'solution', '과제해결수단': 'solution',
    '과제해결수단': 'solution', '기술적해결방법': 'solution',
    '발명의구성': 'solution', '발명을달성하기위한수단': 'solution',
    '발명의해결수단': 'solution', '수단': 'solution',
    '효과': 'effects', '발명의효과': 'effects', '유리한효과': 'effects',
    '발명의작용효과': 'effects', '유리한효과및작용': 'effects',
    '발명에따른효과': 'effects', '기술적효과': 'effects',
    '도면의간단한설명': 'drawdesc', '간단한도면설명': 'drawdesc',
    '도면설명': 'drawdesc', '부호의설명': 'drawdesc', '도면의설명': 'drawdesc',
    '발명을실시하기위한구체적인내용': 'embodiments',
    '발명의실시를위한구체적인내용': 'embodiments',
    '발명을실시하기위한형태': 'embodiments',
    '발명의실시를위한형태': 'embodiments',
    '발명을실시하기위한방법': 'embodiments',
    '구체적인실시예': 'embodiments', '실시예': 'embodiments',
    '발명의상세한설명': 'embodiments', '상세한설명': 'embodiments',
    '발명의실시예': 'embodiments', '바람직한실시예': 'embodiments',
    '특허청구범위': 'claims', '청구범위': 'claims', '청구의범위': 'claims',
    '특허청구의범위': 'claims', '청구항목': 'claims',
    '요약서': 'abstract', '요약': 'summary', '초록': 'summary',
    '요약문': 'summary', '개요': 'summary',
    '대표도': 'absfig', '색인어': 'indexword',
    '도면': 'drawings', '도면들': 'drawings', '첨부도면': 'drawings',
}

# bucket -> XML element, in DTD emission order inside <description>
DESC_ORDER = ['techfield', 'background', 'problem', 'solution', 'effects',
              'drawdesc', 'embodiments']
BUCKET_TAG = {
    'techfield': 'technical-field', 'background': 'background-art',
    'problem': 'tech-problem', 'solution': 'tech-solution',
    'effects': 'advantageous-effects', 'drawdesc': 'description-of-drawings',
    'embodiments': 'description-of-embodiments',
}
SOI = ('problem', 'solution', 'effects')     # live inside <summary-of-invention>


def canon_xml_section(text):
    """If `text` is a patent 【…】 heading, return (bucket, numstr).
    bucket '' means a pure container heading (no content of its own).
    numstr is set for 청구항 N / 도 N.  Returns None for non-headings."""
    t = text.strip()
    if not t:
        return None
    m = re.match(r'^【\s*(.+?)\s*】\s*$', t)
    if not m:
        return None
    inner_raw = m.group(1).strip()
    inner = _norm_tag(inner_raw)
    # 청구항 N  /  도 N  carry an index
    mnum = re.match(r'^(청구항|청구|도)\s*(\d+)([a-zA-Z]?)', inner_raw)
    if mnum:
        base = _norm_tag(mnum.group(1))
        num = mnum.group(2) + mnum.group(3).lower()
        return ('claim' if base in ('청구항', '청구') else 'figure', num)
    if inner in HEAD2XML:
        return (HEAD2XML[inner], None)
    return None


def _esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


def runs_to_xml(runs):
    """Render a paragraph's runs to PatentCAFDOC inline markup.
    Applies <sub>/<sup>/<u>/<b>/<i> per run style; <br/> handled by caller."""
    out = []
    for text, mask, sz, color in runs:
        if not text:
            continue
        t = _esc(text)
        # wrap innermost->outermost; sub/sup mutually exclusive
        if mask & STYLE_SUB:
            t = '<sub>' + t + '</sub>'
        elif mask & STYLE_SUPER:
            t = '<sup>' + t + '</sup>'
        if mask & STYLE_UNDERLINE:
            t = '<u>' + t + '</u>'
        if mask & STYLE_BOLD:
            t = '<b>' + t + '</b>'
        if mask & STYLE_ITALIC:
            t = '<i>' + t + '</i>'
        out.append(t)
    return ''.join(out)


def _emit_figure_image(data, base):
    """Write one figure image for the KIPO XML. Patent drawings are usually
    tiny Group-4 TIFFs — TIFF and JPEG bytes pass through unchanged; other
    formats are re-encoded (bilevel art -> Group-4 TIFF, else JPEG).
    Returns (filename, img-format)."""
    if data[:4] in (b'II*\x00', b'MM\x00*'):
        with open(base + '.tif', 'wb') as fh:
            fh.write(data)
        return os.path.basename(base) + '.tif', 'tif'
    if data[:3] == b'\xff\xd8\xff':
        with open(base + '.jpg', 'wb') as fh:
            fh.write(data)
        return os.path.basename(base) + '.jpg', 'jpg'
    im = Image.open(io.BytesIO(data)); im.load()
    if im.mode != '1':
        g = im.convert('L')
        hist = g.histogram()
        tot = g.size[0] * g.size[1] or 1
        # near-bilevel line art (equations, tables, diagrams): <=3% midtones
        # -> threshold to 1-bit Group-4 TIFF: far smaller than JPEG and keeps
        # edges crisp; only true gray/color content goes to JPEG.
        bilevel = (len(g.getcolors(maxcolors=256) or ()) <= 2 or
                   sum(hist[56:200]) <= 0.03 * tot)
        if bilevel:
            im = g.point(lambda v: 255 if v > 127 else 0, '1')
        elif im.mode != 'RGB':
            im = im.convert('RGB')
    if im.mode == '1':
        im.save(base + '.tif', 'TIFF', compression='group4')
        return os.path.basename(base) + '.tif', 'tif'
    im.save(base + '.jpg', 'JPEG', quality=92)
    return os.path.basename(base) + '.jpg', 'jpg'


def build_patent_xml(src, paras, outdir, imgfiles, xml_name='spec.xml'):
    """Emit a KIPO PatentCAFDOC xml + pat* images into outdir.
    Returns the xml path, or None if the doc isn't patent-sectioned."""
    buckets = {}            # bucket -> list of (runs, images) paragraphs
    title = []
    claims = []             # [{'num', 'paras': [(runs, images)]}]
    figures = []            # [{'num', 'imgs': [(tgt,cx,cy,pos)]}]
    absfig = None
    absfig_img = []         # 대표도 inline images (rare; E-242 allows one)
    cur = 'title'           # default: leading text -> invention title
    cur_claim = None        # current claim dict while in 'claim'
    cur_fig = None          # current figure list while in 'figure'
    seen_heading = False

    def add(bucket, runs, images):
        buckets.setdefault(bucket, []).append((runs, images))

    def unmark(runs):
        """Rewrite stray 식별기호 【】 in body text to plain [ ] — KIPO forbids
        the brackets outside real identification items (E-206) and a bare
        【수학식 1】-style line would otherwise create an item whose content
        can't satisfy its PIC/TABLE rule (S-009)."""
        return [(r[0].replace('【', '[').replace('】', ']'),
                 r[1], r[2], r[3]) for r in runs]

    for runs, images, pp in paras:
        text = ''.join(r[0] for r in runs)
        sec = canon_xml_section(text)
        if sec is not None:
            bucket, num = sec
            seen_heading = True
            cur_claim = None
            cur_fig = None
            if bucket == '_root':
                cur = 'title'
            elif bucket == 'claim':
                cur_claim = {'num': num or str(len(claims) + 1), 'paras': []}
                claims.append(cur_claim)
                cur = 'claim'
            elif bucket == 'figure':
                cur_fig = {'num': num or str(len(figures) + 1), 'imgs': []}
                figures.append(cur_fig)
                cur = 'figure'
            elif bucket == 'absfig':
                cur = 'absfig'
            elif bucket == 'claims' or bucket == 'drawings' or bucket == 'abstract':
                cur = ''          # container heading: no direct content
            else:
                cur = bucket
            continue
        # content paragraph
        if not text.strip() and not images:
            # blank line: keep section spacing but XML ignores empties
            continue
        if cur == 'figure' and cur_fig is not None:
            cur_fig['imgs'].extend(images)
        elif cur == 'claim' and cur_claim is not None:
            cur_claim['paras'].append((unmark(runs), images))
        elif cur == 'absfig':
            # 대표도: capture figure number referenced (e.g. "도 1"), or an
            # embedded representative image (E-242 allows exactly one).
            absfig_img.extend(images)
            mn = re.search(r'(\d+)([a-zA-Z]?)', text)
            if mn:
                absfig = mn.group(1) + mn.group(2).lower()
            elif figures:
                absfig = figures[0]['num']
        elif cur == 'title':
            if text.strip():
                title.append(''.join(r[0] for r in unmark(runs)).strip())
        elif cur:
            add(cur, unmark(runs), images)
        elif images:
            # stray image outside any content section -> own auto figure
            figures.append({'num': str(len(figures) + 1),
                            'imgs': list(images)})

    if not seen_heading:
        return None                     # not a patent-sectioned doc

    # --- write images + build <figure> list -------------------------------
    os.makedirs(outdir, exist_ok=True)
    fig_xml = []
    img_id = [0]
    # media target -> original bytes, for image emission
    media = {}
    with zipfile.ZipFile(src) as zf:
        for name in zf.namelist():
            if name.startswith('word/media/'):
                media[name] = zf.read(name)

    def emu_mm(v):
        # docx extent is EMU; KIPO img he/wi are whole mm (1mm = 36000 EMU)
        return max(1, int(round((v or 0) / 36000.0)))

    emitted_files = {}      # mediaTarget -> (fname, fmt); dedupe repeats
    file_seq = [0]          # pat%05d file counter

    def img_ref(tgt, cx, cy):
        """Write the image file on first use and return its <img> XML."""
        data = media.get(tgt)
        if data is None:
            # rendered-equation pseudo target ('@eqN') — bytes live in imgfiles
            try:
                with open(imgfiles[tgt], 'rb') as fh:
                    data = fh.read()
            except Exception:
                return ''
        if tgt in emitted_files:
            fname, fmt = emitted_files[tgt]
        else:
            file_seq[0] += 1
            base = os.path.join(outdir, 'pat%05d' % file_seq[0])
            try:
                fname, fmt = _emit_figure_image(data, base)
            except Exception:
                fname = 'pat%05d.jpg' % file_seq[0]
                with open(os.path.join(outdir, fname), 'wb') as fh:
                    fh.write(data)
                fmt = 'jpg'
            emitted_files[tgt] = (fname, fmt)
        img_id[0] += 1
        wi, he = emu_mm(cx), emu_mm(cy)
        # KIPO page box is 165x222mm (E-218); clamp oversized extents
        # proportionally — he/wi are display dims, the file stays untouched.
        if wi > 165 or he > 222:
            s = min(165.0 / wi, 222.0 / he)
            wi, he = max(1, int(wi * s)), max(1, int(he * s))
        return ('<img id="i%04d" he="%d" wi="%d" file="%s" img-format="%s"/>'
                % (img_id[0], he, wi, fname, fmt))

    def runs_with_imgs(runs, images):
        """Paragraph body XML with <img> elements spliced at their original
        text offsets (equations/tables arrive as inline images)."""
        if not images:
            return runs_to_xml(runs)
        out, off = [], 0
        pending = sorted(images, key=lambda i: i[3])
        for text, mask, sz, color in runs:
            while pending and pending[0][3] <= off:
                t, cx, cy, _ = pending.pop(0)
                out.append(img_ref(t, cx, cy))
            out.append(runs_to_xml([(text, mask, sz, color)]))
            off += len(text)
        for t, cx, cy, _ in pending:
            out.append(img_ref(t, cx, cy))
        return ''.join(out)

    emitted_fignums = []
    for f in figures:
        if not f['imgs']:
            continue
        imgs = [img_ref(t, cx, cy) for (t, cx, cy, _p) in f['imgs']]
        imgs = [i for i in imgs if i]
        if imgs:
            emitted_fignums.append(f['num'])
            fig_xml.append('<figure num="%s">%s</figure>' %
                           (_esc(f['num']), ''.join(imgs)))

    # --- assemble ---------------------------------------------------------
    def paras_xml(bucket, counter):
        rows = []
        for runs, images in buckets.get(bucket, []):
            body = runs_with_imgs(runs, images)
            if not body.strip():
                continue
            rows.append('<p num="%04d">%s</p>' % (counter(), body))
        return ''.join(rows)

    _ctr = [0]
    def counter():
        _ctr[0] += 1
        return _ctr[0]

    # a (heading*,p+)+ section is only valid if it holds at least one <p>;
    # an empty one breaks NKEditor's submission export, so omit it.
    def sect(tag, bucket):
        body = paras_xml(bucket, counter)
        return '<%s>%s</%s>' % (tag, body, tag) if body else ''

    soi = (sect('tech-problem', 'problem') +
           sect('tech-solution', 'solution') +
           sect('advantageous-effects', 'effects'))

    L = []
    L.append('<?xml version="1.0" encoding="euc-kr" ?>')
    L.append('<KIPO keapsVersion="4.02" editorKind="K" pageCount="1" '
             'xmlns="http://www.kipo.go.kr">')
    L.append('<PatentCAFDOC docflag="1.0" documentID="0">')
    L.append('<description>')
    L.append('<invention-title>%s</invention-title>' %
             _esc(' '.join(title)))
    L.append(sect('technical-field', 'techfield'))
    L.append(sect('background-art', 'background'))
    if soi:
        L.append('<summary-of-invention>%s</summary-of-invention>' % soi)
    L.append(sect('description-of-drawings', 'drawdesc'))
    L.append(sect('description-of-embodiments', 'embodiments'))
    L.append('</description>')
    # claims
    L.append('<claims>')
    for c in claims:
        body = '<br/>'.join(runs_with_imgs(r, i) for r, i in c['paras']
                            if runs_with_imgs(r, i).strip())
        L.append('<claim num="%s"><claim-text>%s</claim-text></claim>' %
                 (_esc(c['num']), body))
    L.append('</claims>')
    # abstract
    L.append('<abstract><summary>%s</summary>' %
             paras_xml('summary', counter))
    # the abstract figure reference must resolve to an emitted <figure num>;
    # docx 대표도 text may name a group ("도 3") that emitted as 3a/3b.
    if absfig and absfig not in emitted_fignums:
        cand = [n for n in emitted_fignums if n.startswith(absfig)]
        absfig = cand[0] if cand else None
    if not absfig and emitted_fignums:
        absfig = emitted_fignums[0]
    if absfig_img:
        # 대표도 carried its own image (E-242 allows exactly one)
        inner = ''.join(img_ref(t, cx, cy) for t, cx, cy, _p in absfig_img[:1])
        L.append('<abstract-figure><p num="%04d">%s</p></abstract-figure>'
                 % (counter(), inner))
    elif absfig:
        L.append('<abstract-figure><p num="%04d"><figref num="%s"/></p>'
                 '</abstract-figure>' % (counter(), _esc(absfig)))
    L.append('</abstract>')
    # drawings
    L.append('<drawings>')
    L.extend(fig_xml)
    L.append('</drawings>')
    L.append('</PatentCAFDOC>')
    L.append('</KIPO>')
    xml = '\n'.join(L)
    xp = os.path.join(outdir, xml_name)
    with open(xp, 'wb') as f:
        f.write(xml.encode('cp949', 'replace'))
    return xp


def split_cp949(data, limit):
    """Split cp949 bytes at char boundaries into <=limit chunks."""
    chunks = []
    start = 0
    n = len(data)
    while start < n:
        end = min(start + limit, n)
        # don't split a 2-byte cp949 char: a lead byte 0x81-0xFE must pair
        # walk back if the byte at 'end-1' is an odd-count lead byte run
        if end < n:
            k = end - 1
            lead = 0
            while k >= start and 0x81 <= data[k] <= 0xFE:
                lead += 1
                k -= 1
            if lead % 2 == 1:
                end -= 1
            if end <= start:           # safety: at least one byte
                end = start + 1
        chunks.append(data[start:end])
        start = end
    return chunks


def build_job(paras, imgdir, imgfiles, skel=None):
    """paras: list of (runs, images, paraProps) -> job bytes.
    imgdir: dir for HLT%N.HIF staging. imgfiles: {mediaTarget: extractedPath}.
    skel: path to a .frm structured skeleton (Patentd.frm); when set the job is
          emitted section-tagged so TerConv fills the skeleton's elements.
    Job format:
      u32 imgDirLen; imgDir (utf8)
      u32 skelLen; skel (utf8)               # empty -> flat (unstructured) doc
      u32 secCount
        per section: u32 tagLen; tag (cp949) # element display name; ''=default
      u32 lineCount
      per line: u32 secIndex; u8 kind
        kind 0 (text): u32 len; bytes; u8 term; u32 nfmt;
                       fmts: u32 s,u32 e,u32 mask,i32 sizeX2,u32 colorBGR
        kind 1 (image): u32 pathLen; path(utf8);
                        i32 dispW_pt; i32 dispH_pt; i32 pxW; i32 pxH
      both end with: u8 hasPara; i32 left,right,first; u32 pflags;
                     i32 spBef, spAft, spBet   (twips)
    """
    lines = []          # (secIndex, lineTuple)
    sections = []       # index -> skeleton element display name ('' = default)
    secmap = {}
    hif = [0]           # next hif index

    def sec_index(name):
        if name not in secmap:
            secmap[name] = len(sections)
            sections.append(name)
        return secmap[name]

    sec_index('')       # section 0: content before the first 【…】 marker
    cur_sec = 0

    def para_spec(pp):
        """(hasPara, left,right,first, pflags, spB,spA,spBt)"""
        if not pp:
            return (0, 0, 0, 0, 0, 0, 0, 0)
        return (1,
                pp.get('left', 0), pp.get('right', 0), pp.get('first', 0),
                pp.get('jc', PALIGN_LEFT),
                pp.get('before', 0), pp.get('after', 0), 160)

    for runs, images, pp in paras:
        ps = para_spec(pp)
        ptext = ''.join(r[0] for r in runs).strip()
        elem = canon_element(ptext)
        if elem:
            name, keep = elem
            cur_sec = sec_index(name)
            if not keep:
                # The skeleton already draws the 【…】 element marker — emitting
                # the docx heading too would show it twice. Numbered elements
                # (청구항/도) keep their heading so the index isn't lost.
                continue
        # text part
        pieces = []
        fmts = []
        pos = 0
        for text, mask, sz, color in runs:
            b = cp949(text)
            if not b:
                continue
            if (mask or sz >= 0 or color >= 0) and len(b):
                fmts.append((pos, pos + len(b), mask, sz, color))
            pieces.append(b)
            pos += len(b)
        data = b''.join(pieces)

        if not data and not images:
            lines.append((cur_sec, ('t', b'', True, [], ps)))
            continue

        if data:
            chunks = split_cp949(data, MAX_CHUNK)
            base = 0
            for ci, ch in enumerate(chunks):
                last = ci == len(chunks) - 1 and not images
                lo, hi = base, base + len(ch)
                lf = []
                for (s, e, mk, sz, color) in fmts:
                    if e <= lo or s >= hi:
                        continue
                    lf.append((max(s, lo) - lo, min(e, hi) - lo, mk, sz, color))
                lines.append((cur_sec, ('t', ch, last, lf, ps)))
                base = hi

        for (tgt, cx, cy, pos) in images:
            path = imgfiles.get(tgt)
            if not path:
                continue
            data_b = open(path, 'rb').read()
            dims = img_px_dims(data_b)
            pxW = pxH = None
            dpi = None
            if dims:
                pxW, pxH, dpi = dims
            # display size in points: prefer docx extent (EMU/12700), else px->pt via dpi
            if cx and cy:
                dw, dh = cx // 12700, cy // 12700
            elif pxW and pxH:
                use_dpi = dpi or 96
                dw, dh = int(pxW * 72 / use_dpi), int(pxH * 72 / use_dpi)
            else:
                dw = dh = 0
            if not pxW or not pxH:
                pxW = int(dw * 96 / 72) if dw else 0
                pxH = int(dh * 96 / 72) if dh else 0
            lines.append((cur_sec, ('i', path, dw, dh, pxW, pxH, ps)))

    d = imgdir.encode('utf-8')
    sb = (skel or '').encode('utf-8')
    out = [struct.pack('<I', len(d)), d]
    out.append(struct.pack('<I', len(sb)))
    out.append(sb)
    out.append(struct.pack('<I', len(sections)))
    for s in sections:
        eb = s.encode('cp949', 'replace')
        out.append(struct.pack('<I', len(eb)))
        out.append(eb)
    out.append(struct.pack('<I', len(lines)))
    for sec, ln in lines:
        out.append(struct.pack('<I', sec))
        if ln[0] == 't':
            _, ch, term, lf, ps = ln
            out.append(b'\x00')
            out.append(struct.pack('<I', len(ch)))
            out.append(ch)
            out.append(b'\x01' if term else b'\x00')
            out.append(struct.pack('<I', len(lf)))
            for (s, e, mk, sz, color) in lf:
                out.append(struct.pack('<IIIiI', s, e, mk, sz, color & 0xFFFFFFFF))
        else:
            _, path, dw, dh, pw, ph, ps = ln
            pb = path.encode('utf-8')
            out.append(b'\x01')
            out.append(struct.pack('<I', len(pb)))
            out.append(pb)
            out.append(struct.pack('<iiii', dw, dh, pw, ph))
        if ps[0]:
            out.append(b'\x01')
            out.append(struct.pack('<iiiIiii', ps[1], ps[2], ps[3], ps[4], ps[5], ps[6], ps[7]))
        else:
            out.append(b'\x00')
    return b''.join(out)


def media_to_tiff(data, out_tif):
    """Convert image bytes to an uncompressed TIFF at out_tif via Pillow.

    The TE save path embeds TIFF verbatim but must ImageGear-decode PNG/JPEG —
    which fails on palette/odd PNGs. Staging a TIFF makes the embed reliable.
    Returns True on success."""
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception:
        return False
    # normalize modes ImageGear's TIFF reader handles cleanly
    if im.mode in ('1',):
        pass                          # keep 1-bit (patent line drawings)
    elif im.mode in ('L', 'RGB'):
        pass
    else:
        im = im.convert('RGB')        # P, RGBA, LA, CMYK, I;16, ...
    dpi = im.info.get('dpi')
    try:
        if dpi:
            im.save(out_tif, 'TIFF', compression='raw', dpi=dpi)
        else:
            im.save(out_tif, 'TIFF', compression='raw')
        return True
    except Exception:
        return False


DOCCMP = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                      'DocCmp.exe')
EQRENDER = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        'EqRender.exe')


WARNINGS = []             # conversion-level problems for the review report


def render_objects(src, imgdir):
    """Render the docx's OMML equations AND top-level tables to trimmed
    300dpi G4 TIFFs via EqRender.exe + Word COM (Range.CopyAsPicture ->
    EMF -> rasterize -> crop).

    Returns (eqs, tbls): eqs = {omath_ordinal: (tif, w_mm, h_mm)} keyed by
    the 1-based position in main_story_omaths() / Word's doc.OMaths,
    tbls = {top_level_ordinal: (tif, w_mm, h_mm)}. Both empty when the
    doc has neither equations nor tables. A missing ordinal (tracked
    deletion, render failure) degrades just that equation to linear text.

    If Word enumerates a different number of equations than the XML walk
    finds, ordinals can't be trusted to line up, so no equation image is
    used at all — a wrong equation in a patent spec is worse than linear
    text.

    Rendering can fail transiently — Word COM refuses automation while the
    user's own WINWORD shows a modal dialog, or a stale instance is still
    shutting down — so retry a few times before degrading to linear text
    (equations) / flattened cells (tables)."""
    empty = ({}, {})
    try:
        with zipfile.ZipFile(src) as zf:
            xml = zf.read('word/document.xml')
        n_math = len(main_story_omaths(ET.fromstring(xml)))
        n_tbl = xml.count(b'<w:tbl>') + xml.count(b'<w:tbl ')
        if n_math == 0 and n_tbl == 0:
            return empty
        if not os.path.exists(EQRENDER):
            if n_math:
                WARNINGS.append('EqRender.exe 없음 — 수식이 선형 텍스트로 변환됨')
            if n_tbl:
                WARNINGS.append('EqRender.exe 없음 — 표가 텍스트 나열로 변환됨')
            return empty

        def read_map(name, prefix):
            out = {}
            path = os.path.join(imgdir, name)
            if os.path.exists(path):
                with open(path) as f:
                    for line in f:
                        idx, w, h = line.rstrip('\n').split('\t')
                        out[int(idx)] = (
                            os.path.join(imgdir,
                                         '%s%04d.tif' % (prefix, int(idx))),
                            float(w), float(h))
            return out

        last_err = ''
        eqs, tbls, e_expect, t_expect = {}, {}, None, None
        for attempt in range(4):
            try:
                r = subprocess.run([EQRENDER, src, imgdir],
                                   capture_output=True, text=True,
                                   timeout=600)
                last_err = (r.stderr or r.stdout or '').strip()[-300:]
                eqs = read_map('eqmap.tsv', 'eq')
                tbls = read_map('tblmap.tsv', 'tbl')
                # expected counts come from EqRender's summary lines (it
                # alone knows deletion/nesting skips)
                e_total = e_expect = t_expect = None
                m = re.search(r'EQSUMMARY total=(\d+) skipped=(\d+) '
                              r'rendered=(\d+)', r.stdout or '')
                if m:
                    e_total = int(m.group(1))
                    e_expect = e_total - int(m.group(2))
                m = re.search(r'TBLSUMMARY toplevel=(\d+) skipped=(\d+) '
                              r'rendered=(\d+)', r.stdout or '')
                if m:
                    t_expect = int(m.group(1)) - int(m.group(2))
                if e_total is not None and e_total != n_math:
                    # deterministic mismatch — retrying won't help
                    WARNINGS.append(
                        '수식 개수 불일치(Word %d개, XML %d개) — 위치를 '
                        '보장할 수 없어 모든 수식을 선형 텍스트로 변환함'
                        % (e_total, n_math))
                    eqs = {}
                    e_expect = 0
                    if t_expect is None or len(tbls) == t_expect:
                        break
                if r.returncode == 0 and e_expect is not None \
                        and len(eqs) == e_expect \
                        and (t_expect is None or len(tbls) == t_expect):
                    break
                last_err = ('partial eq %d/%s tbl %d/%s'
                            % (len(eqs), e_expect, len(tbls), t_expect))
            except subprocess.TimeoutExpired:
                last_err = 'timeout'
            # transient Word COM contention — let a busy/modal WINWORD settle
            time.sleep(3 + attempt * 3)
        else:
            print('WARNING: object render incomplete (%s)' % last_err,
                  file=sys.stderr)
        if eqs:
            print('equations: rendered %d/%d to TIFF' % (len(eqs), n_math))
        if tbls:
            print('tables: rendered %d to TIFF' % len(tbls))
        if n_math and (e_expect is None or len(eqs) < e_expect):
            WARNINGS.append('수식 렌더링 실패(%s) — %d개 수식이 선형 텍스트로 '
                            '변환됨. Word의 대화상자/팝업을 모두 닫고 다시 '
                            '변환하세요.'
                            % (last_err, n_math - len(eqs)))
        if t_expect and len(tbls) < t_expect:
            WARNINGS.append('표 렌더링 실패(%s) — %d개 표가 텍스트 나열로 '
                            '변환됨' % (last_err, t_expect - len(tbls)))
        return eqs, tbls
    except Exception as e:
        WARNINGS.append('수식/표 렌더링 불가(%r) — 텍스트로 폴백됨' % e)
        print('WARNING: object render unavailable (%r)' % e,
              file=sys.stderr)
        return empty


def parse_docx(path, imgdir, eqs=None, tbls=None):
    """Parse docx; extract media to imgdir (as TIFF). Returns (paras, {mediaTarget: path}).
    eqs: {omath_ordinal: (tif, w_mm, h_mm)} from render_objects — spliced
    into paragraphs where each m:oMath appears.
    tbls: {ordinal: (tif, w_mm, h_mm)} rendered top-level tables — each
    becomes an image-only paragraph at the w:tbl's document position."""
    imgfiles = {}
    tstate = [0]                          # top-level w:tbl ordinal counter
    with zipfile.ZipFile(path) as zf:
        styles = load_styles(zf)
        rels = load_rels(zf)
        os.makedirs(imgdir, exist_ok=True)
        for name in zf.namelist():
            if name.startswith('word/media/'):
                base = os.path.basename(name)
                data = zf.read(name)
                # stage as TIFF so the HIF embed path works for every format
                # (append .tif to the full basename to keep names unique)
                tif = os.path.join(imgdir, base + '.tif')
                if HAVE_PIL and media_to_tiff(data, tif):
                    imgfiles[name] = tif
                else:
                    tgt = os.path.join(imgdir, base)
                    with open(tgt, 'wb') as f:
                        f.write(data)
                    imgfiles[name] = tgt
        root = ET.fromstring(zf.read('word/document.xml'))
    body = root.find(W + 'body')
    eqstate = None
    if eqs:
        eq_ord = {id(om): i + 1
                  for i, om in enumerate(main_story_omaths(body))}
        eqstate = (eq_ord, eqs)
    paras = list(iter_paragraphs(body, styles, rels, imgfiles, eqstate,
                                 tbls=tbls, tstate=tstate))
    return paras, imgfiles


# Patent specification skeleton: carries the KIPO element tree (요약서/명세서/
# 특허청구범위/도면 + subsections) so the output is a real structured 특허명세서
# rather than a flat doc. Built by NKEditor's own form engine.
SKEL_PATENT = os.path.join(NK_DIR, 'Epasl', 'FRM', 'Patentd.frm')


def _is_patent_doc(paras):
    return any(canon_xml_section(''.join(r[0] for r in runs).strip())
               for runs, _i, _pp in paras)


def convert_structured(src, dst, paras, imgfiles):
    """docx -> KIPO PatentCAFDOC xml -> NKEditor builds the structured doc ->
    drives Save-As to a real .hlt at `dst`. Returns True on success."""
    import nkwin
    stem = os.path.splitext(os.path.basename(dst))[0]
    scratch = tempfile.mkdtemp(prefix='d2h_xml')
    xmlpath = build_patent_xml(src, paras, scratch, imgfiles,
                               xml_name=stem + '.xml')
    if not xmlpath:
        return False
    print('structured: emitted PatentCAFDOC ->', xmlpath)
    saved = nkwin.nk_import_and_save_hlt(
        xmlpath, stem, out_dir=os.path.dirname(os.path.abspath(dst)))
    if not saved or not os.path.exists(saved):
        print('ERROR: NKEditor save-as did not produce a .hlt', file=sys.stderr)
        return False
    if os.path.abspath(saved) != os.path.abspath(dst):
        import shutil
        shutil.move(saved, dst)   # move — don't litter the dialog's folder
    print('structured hlt ->', dst)
    return True


def fix_rtf(path):
    """NKEditor's RTF export can drop the root group's closing '}' (observed
    on docs with embedded images — the file ends mid-group and Word refuses
    to open it). Append the missing braces. Escaped \\{ \\} literals are
    skipped when counting."""
    data = open(path, 'rb').read()
    depth = 0
    i = 0
    n = len(data)
    while i < n:
        c = data[i:i + 1]
        if c == b'\\':                       # skip escaped char / control word
            i += 2
            continue
        if c == b'{':
            depth += 1
        elif c == b'}':
            depth -= 1
        i += 1
    if depth > 0:
        with open(path, 'ab') as f:
            f.write(b'}' * depth)


def make_review(src, dst):
    """Human-review artifacts after `dst` (.hlt) exists:
      - export the .hlt back to <stem>.rtf through NKEditor (what the
        editor actually ended up holding), and
      - Word-compare  src.docx  vs  that .rtf  into  <stem>.compare.docx
        (tracked-changes document — every dropped/substituted char and
        lost formatting shows up as a revision).
    Failures are reported but never fail the conversion."""
    import nkwin
    import shutil
    base = os.path.splitext(dst)[0]
    rtfdst = base + '.rtf'
    cmpdoc = base + '.compare.docx'
    try:
        rtf = nkwin.nk_export_rtf(dst)
    except Exception as e:
        print('review: rtf export failed (%r)' % e, file=sys.stderr)
        return
    if not rtf or not os.path.exists(rtf):
        print('review: rtf export produced nothing — skipping compare',
              file=sys.stderr)
        return
    if os.path.abspath(rtf) != os.path.abspath(rtfdst):
        # the save-as dialog's last-used dir can redirect the export
        # (e.g. the user's Desktop) — move it, don't leave litter behind
        shutil.move(rtf, rtfdst)
    fix_rtf(rtfdst)
    print('review: rtf ->', rtfdst)
    if not os.path.exists(DOCCMP):
        print('review: DocCmp.exe missing — skipping compare',
              file=sys.stderr)
        return
    try:
        for cand in (cmpdoc, base + '.compare-new.docx'):
            r = subprocess.run([DOCCMP, src, rtfdst, cand],
                               capture_output=True, text=True, timeout=300)
            if r.returncode == 0 and os.path.exists(cand):
                print('review: compare doc ->', cand)
                return
        print('review: compare failed (%s)'
              % (r.stderr or r.stdout or '').strip()[-200:], file=sys.stderr)
    except Exception as e:
        print('review: compare failed (%r)' % e, file=sys.stderr)


def convert_one(src, dst, force_flat=False):
    """Convert a single .docx to .hlt. Returns 0 on success."""
    imgdir = tempfile.mkdtemp(prefix='d2h_img')
    CHAR_REPORT.clear()
    del WARNINGS[:]
    eqs, tbls = render_objects(src, imgdir)
    paras, imgfiles = parse_docx(src, imgdir, eqs, tbls)

    if CHAR_REPORT or WARNINGS:
        # human-review trail: conversion warnings + every unsupported char
        # and what it became — as a Word table (easy to scan) plus a plain
        # text mirror for scripts.
        stem = os.path.splitext(dst)[0]
        rep = stem + '.charset.txt'
        with open(rep, 'w', encoding='utf-8') as f:
            for w in WARNINGS:
                f.write('!! %s\n' % w)
            if WARNINGS:
                f.write('\n')
            f.write('chars not representable in KIPO editor — substituted or '
                    'dropped during conversion\n\n')
            for orig, repl, ctx in CHAR_REPORT:
                f.write('U+%04X %r -> %r   …%s\n' %
                        (ord(orig), orig, repl or '(삭제)', ctx))
        try:
            write_charset_docx(stem + '.charset.docx', WARNINGS, CHAR_REPORT)
        except Exception as e:
            print('review: charset docx write failed (%r)' % e,
                  file=sys.stderr)
        subs = sum(1 for _o, r, _c in CHAR_REPORT if r)
        drops = len(CHAR_REPORT) - subs
        print('review: %d warnings, %d substituted, %d dropped — %s(.docx)'
              % (len(WARNINGS), subs, drops, rep), file=sys.stderr)
    else:
        # clean state: remove reports from a previous run so stale warnings
        # can't be mistaken for problems in THIS output
        for stale in (os.path.splitext(dst)[0] + '.charset.txt',
                      os.path.splitext(dst)[0] + '.charset.docx'):
            try:
                os.remove(stale)
            except OSError:
                pass

    # Patent-sectioned docs go through the real structured path: emit KIPO
    # PatentCAFDOC XML, let NKEditor build the element tree, then save .hlt.
    # Plain docs fall back to the flat TerConv pipeline.
    if not force_flat and _is_patent_doc(paras) and HAVE_PIL:
        try:
            if convert_structured(src, dst, paras, imgfiles):
                make_review(src, dst)
                return 0
        except Exception as e:
            print('structured path failed (%r); falling back to flat' % e,
                  file=sys.stderr)

    job = build_job(paras, imgdir, imgfiles, None)
    nimg = sum(1 for r, i, pp in paras for _ in i)
    print(f'{len(paras)} paragraphs, {nimg} images -> {len(job)} job bytes')

    with tempfile.NamedTemporaryFile(suffix='.job', delete=False) as tf:
        tf.write(job)
        jobfile = tf.name

    # hwplt32's save embeds images by reading GetTempPath()\HLT%N.HIF — it does
    # NOT use the imagepath/imgdir we stage into. Point TerConv's TMP/TEMP at the
    # staging dir so GetTempPath() resolves there and reads OUR fresh HIF files
    # (otherwise it picks up stale HLT*.HIF left in the user temp root).
    env = dict(os.environ)
    env['TMP'] = env['TEMP'] = imgdir
    try:
        r = subprocess.run([TERCONV, jobfile, dst, NK_DIR],
                           capture_output=True, text=True, timeout=240, env=env)
        sys.stdout.write(r.stdout)
        sys.stderr.write(r.stderr)
        if r.returncode != 0:
            return r.returncode
    finally:
        os.unlink(jobfile)

    if not os.path.exists(dst):
        print('ERROR: output not created', file=sys.stderr)
        return 1
    print('OK ->', dst)
    make_review(src, dst)
    return 0


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    force_flat = '--flat' in sys.argv[1:]
    if not args:
        print('usage: docx2hlt.py input.docx [output.hlt] [more.docx ...]')
        print('  output omitted -> <input>.hlt written next to each input')
        return 2
    # explicit output only in the classic 2-arg form; otherwise every arg is
    # a .docx and each writes <same name>.hlt in its own folder. Conversions
    # run sequentially — the structured path drives the NKEditor GUI.
    if len(args) == 2 and not args[1].lower().endswith('.docx'):
        jobs = [(args[0], args[1])]
    else:
        jobs = [(a, os.path.splitext(os.path.abspath(a))[0] + '.hlt')
                for a in args]
    rc = 0
    for src, dst in jobs:
        if len(jobs) > 1:
            print('=== %s' % src)
        try:
            r = convert_one(src, os.path.abspath(dst), force_flat)
        except Exception as e:
            print('FAILED %s: %r' % (src, e), file=sys.stderr)
            r = 1
        if r:
            rc = r
    return rc


if __name__ == '__main__':
    sys.exit(main())
