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


# --- list numbering ----------------------------------------------------------
# Word list labels (bullets, "1.", "a)") are not text in document.xml - they
# come from numbering.xml - so they used to vanish (an English PCT spec: "- startPRB",
# "○ 3", "2."; a domestic spec: 64 "-"/"•" bullets) while the filed .hlt files carry
# them as text. Paragraph-number lists ("[%1]", "[00%1]" -> [0001]) are
# skipped: NKEditor numbers paragraphs itself.
NUMBERING = {}            # per document, filled by parse_docx

BULLET_MAP = {'': '•', '': '■', '': '□', '': '▶',
              '': '√', '': '―', '': '◆', '': '■',
              '': '●', '': '→', '': '⇒', 'o': '○',
              '': '○', '': '•'}


def load_numbering(zf, styles):
    try:
        root = ET.fromstring(zf.read('word/numbering.xml'))
    except KeyError:
        return {}
    absl = {}                                 # abstractNumId -> {ilvl: lvl}
    for a in root.findall(W + 'abstractNum'):
        lv = {}
        for l in a.findall(W + 'lvl'):
            def val(tag, d=None):
                e = l.find(W + tag)
                return e.get(W + 'val') if e is not None else d
            lv[int(l.get(W + 'ilvl', '0'))] = {
                'fmt': val('numFmt', 'decimal'), 'text': val('lvlText', ''),
                'start': int(val('start', '1') or 1),
                'suff': val('suff', 'tab')}
        absl[a.get(W + 'abstractNumId')] = lv
    nums = {}                                 # numId -> (aid, {ilvl: start})
    for n in root.findall(W + 'num'):
        an = n.find(W + 'abstractNumId')
        if an is None:
            continue
        ov = {}
        for o in n.findall(W + 'lvlOverride'):
            so = o.find(W + 'startOverride')
            if so is not None:
                ov[int(o.get(W + 'ilvl', '0'))] = int(so.get(W + 'val'))
        nums[n.get(W + 'numId')] = (an.get(W + 'val'), ov)
    snum = {}                                 # styleId -> (numId, ilvl)
    for sid, (_props, ppr) in styles.items():
        np_ = ppr.find(W + 'numPr') if ppr is not None else None
        if np_ is not None:
            e = np_.find(W + 'numId')
            i = np_.find(W + 'ilvl')
            if e is not None:
                snum[sid] = (e.get(W + 'val'),
                             int(i.get(W + 'val')) if i is not None else 0)
    return {'abs': absl, 'nums': nums, 'snum': snum, 'ctr': {},
            'started': set()}


def _fmt_num(n, fmt):
    if fmt in ('lowerLetter', 'upperLetter'):
        s = ''
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(97 + r) + s
        return s.upper() if fmt == 'upperLetter' else s
    if fmt in ('lowerRoman', 'upperRoman'):
        out = ''
        for v, r in ((1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'),
                     (100, 'c'), (90, 'xc'), (50, 'l'), (40, 'xl'), (10, 'x'),
                     (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')):
            while n >= v:
                out += r
                n -= v
        return out.upper() if fmt == 'upperRoman' else out
    if fmt == 'ganada' and 1 <= n <= 14:
        return '가나다라마바사아자차카타파하'[n - 1]
    if fmt == 'chosung' and 1 <= n <= 14:
        return 'ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ'[n - 1]
    if fmt == 'decimalEnclosedCircle' and 1 <= n <= 20:
        return chr(0x2460 + n - 1)
    if fmt == 'decimalZero':
        return '%02d' % n
    return str(n)


def list_label(p, styles):
    """The visible list label of paragraph `p` ('' if none), advancing the
    document's list counters."""
    if not NUMBERING:
        return ''
    ppr = p.find(W + 'pPr')
    nid, il = None, 0
    if ppr is not None:
        np_ = ppr.find(W + 'numPr')
        if np_ is not None:
            e = np_.find(W + 'numId')
            i = np_.find(W + 'ilvl')
            nid = e.get(W + 'val') if e is not None else None
            il = int(i.get(W + 'val')) if i is not None else 0
        if nid is None:
            ps = ppr.find(W + 'pStyle')
            if ps is not None and ps.get(W + 'val') in NUMBERING['snum']:
                nid, il2 = NUMBERING['snum'][ps.get(W + 'val')]
                il = il if np_ is not None and np_.find(W + 'ilvl') is not None \
                    else il2
    if not nid or nid == '0' or nid not in NUMBERING['nums']:
        return ''
    aid, ov = NUMBERING['nums'][nid]
    lv = NUMBERING['abs'].get(aid, {})
    if il not in lv:
        return ''
    ctr = NUMBERING['ctr'].setdefault(aid, {})
    if nid not in NUMBERING['started']:
        NUMBERING['started'].add(nid)
        for k, v in ov.items():             # "restart numbering" instance
            ctr[k] = v - 1
    ctr[il] = ctr.get(il, lv[il]['start'] - 1) + 1
    for k in [k for k in ctr if k > il]:
        del ctr[k]
    lvl = lv[il]
    text = lvl['text']
    if '[' in text or '【' in text:
        return ''                           # paragraph number: NKEditor's job
    if lvl['fmt'] == 'none':
        return ''
    if lvl['fmt'] == 'bullet':
        label = BULLET_MAP.get(text, text)
    else:
        def sub(m):
            k = int(m.group(1)) - 1
            f = lv.get(k, lvl)['fmt']
            return _fmt_num(ctr.get(k, lv.get(k, lvl)['start']), f)
        label = re.sub(r'%(\d)', sub, text)
    if not label.strip():
        return ''
    return label + ('' if lvl['suff'] == 'nothing' else ' ')


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
            # VML (w:pict, OLE w:object previews such as MathType) carries
            # its size in the v:shape style, not a wp:extent
            shp = _find_skip_txbx(node, V + 'shape')
            if cx is None and shp is not None:
                st = dict(kv.split(':', 1) for kv in
                          (shp.get('style') or '').replace(' ', '').split(';')
                          if ':' in kv)

                def emu(v):
                    m = re.match(r'^([\d.]+)(pt|in|cm|mm|px)?$', v or '')
                    if not m:
                        return None
                    f = {'pt': 12700, 'in': 914400, 'cm': 360000,
                         'mm': 36000, 'px': 9525, None: 12700}[m.group(2)]
                    return int(float(m.group(1)) * f)
                cx, cy = emu(st.get('width')), emu(st.get('height'))
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
            elif tag in (W + 'drawing', W + 'pict', W + 'object'):
                # w:object = embedded OLE (MathType equations etc.): its VML
                # preview image is what Word shows - it used to be dropped
                # without a trace (a domestic spec: 5 inline objects)
                flush()
                add_image(ch)
                if tag != W + 'object':
                    walk(ch)            # text-box paragraphs, if any
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

    label = list_label(p, styles)
    if label and not ''.join(t.text or '' for t in p.iter(W + 't')) \
            .lstrip().startswith('【'):
        add_text(label, 0, base[1], base[2])
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


def iter_blocks(parent, styles, rels, media, eqs=None):
    """Like iter_paragraphs, but a top-level w:tbl stays a table (the
    의견서 path emits it as a native OASIS table, not an image).
    Yields ('p', (runs, images, pp)) and ('tbl', {'grid': [dxa..],
    'rows': [[cell..]..]}); cell = {'paras': [...], 'span': n,
    'vmerge': None|'restart'|'continue'}. Nested tables flatten into their
    cell's paragraphs."""
    for child in parent:
        if child.tag == W + 'p':
            yield 'p', para_runs(child, styles, rels, media, eqs)
        elif child.tag == W + 'sdt':
            content = child.find(W + 'sdtContent')
            if content is not None:
                yield from iter_blocks(content, styles, rels, media, eqs)
        elif child.tag == W + 'tbl':
            grid = []
            tg = child.find(W + 'tblGrid')
            if tg is not None:
                for gc in tg.findall(W + 'gridCol'):
                    v = gc.get(W + 'w') or '0'
                    grid.append(int(v) if v.isdigit() else 0)
            rows = []
            for tr in child.findall(W + 'tr'):
                cells = []
                for tc in tr.findall(W + 'tc'):
                    span, vm = 1, None
                    tcpr = tc.find(W + 'tcPr')
                    if tcpr is not None:
                        gs = tcpr.find(W + 'gridSpan')
                        if gs is not None and (gs.get(W + 'val') or '').isdigit():
                            span = max(1, int(gs.get(W + 'val')))
                        v = tcpr.find(W + 'vMerge')
                        if v is not None:
                            vm = v.get(W + 'val') or 'continue'
                    cells.append({'paras': list(iter_paragraphs(
                        tc, styles, rels, media, eqs, depth=1)),
                        'span': span, 'vmerge': vm})
                rows.append(cells)
            yield 'tbl', {'grid': grid, 'rows': rows}


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


def write_charset_docx(path, warnings, rows, notes=()):
    """Write the charset review report as a minimal Word document —
    a heading, warning lines, image-conversion notes (first line bold,
    the rest indented), then a table of
    코드포인트 | 원본 | 치환 | 문맥."""
    from xml.sax.saxutils import escape as _esc
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
        '<w:t xml:space="preserve">!! %s</w:t></w:r></w:p>' % _esc(w)
        for w in warnings)
    warn_xml += ''.join(
        '<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (
            '<w:pPr><w:ind w:left="400"/></w:pPr>' if k else '',
            '' if k else '<w:rPr><w:b/></w:rPr>', _esc(n))
        for k, n in enumerate(notes))
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

# 국문 PCT (WIPO ST.36 application-body-v1-6.dtd) differs from PatentCAFDOC:
# <disclosure> / <mode-for-invention> instead of <summary-of-invention> /
# <description-of-embodiments>, <abstract> holds <p> directly (no 【요약】,
# no 대표도), and each <figure> holds exactly one <img>. Headings below
# override HEAD2XML in the PCT build only, so domestic output is unchanged.
PCT_HEAD2XML = {
    '발명의내용': 'disclosure', '발명의상세한설명': 'disclosure',
    '요약서': 'summary',
    '발명의실시를위한최선의형태': 'bestmode',
    '산업상이용가능성': 'industrial',
}
PCT_TAG = {
    'drawdesc': 'description-of-drawings', 'bestmode': 'best-mode',
    'embodiments': 'mode-for-invention',
    'industrial': 'industrial-applicability',
}

# 영문 PCT (application-body lang="en", PCT명세서(국제출원) 영문서식): the same
# ST.36 elements under English headings (【Technical Field】, 【Mode for
# Invention】, 【CLAIMS】, 【Claim 1】, 【Figure 4A】 - the markers of
# filed English PCT .hlt files). Keys: heading lowercased, spaces and punctuation removed.
# 【ABSTRACT】 holds the abstract text itself (no 【요약】 sub-item).
HEAD2XML_EN = {
    'description': '_root', 'specification': '_root',
    'inventiontitle': 'title', 'titleofinvention': 'title',
    'titleoftheinvention': 'title',
    'technicalfield': 'techfield', 'fieldoftheinvention': 'techfield',
    'field': 'techfield',
    'backgroundart': 'background', 'background': 'background',
    'backgroundoftheinvention': 'background',
    'disclosure': 'disclosure', 'disclosureofinvention': 'disclosure',
    'summaryofinvention': 'disclosure', 'summaryoftheinvention': 'disclosure',
    'technicalproblem': 'problem',
    'technicalsolution': 'solution', 'solutiontoproblem': 'solution',
    'advantageouseffects': 'effects',
    'advantageouseffectsofinvention': 'effects',
    'descriptionofdrawings': 'drawdesc', 'briefdescriptionofdrawings': 'drawdesc',
    'briefdescriptionofthedrawings': 'drawdesc',
    'bestmode': 'bestmode', 'bestmodeforcarryingouttheinvention': 'bestmode',
    'modeforinvention': 'embodiments', 'modefortheinvention': 'embodiments',
    'descriptionofembodiments': 'embodiments',
    'detaileddescription': 'embodiments',
    'industrialapplicability': 'industrial',
    'claims': 'claims', 'whatisclaimedis': 'claims',
    'abstract': 'summary', 'abstractofthedisclosure': 'summary',
    'drawings': 'drawings', 'figures': 'drawings',
}


def _norm_en(s):
    return re.sub(r'[^a-z]', '', s.lower())


def canon_xml_section(text, en=False):
    """If `text` is a patent 【…】 heading, return (bucket, numstr).
    bucket '' means a pure container heading (no content of its own).
    numstr is set for 청구항 N / 도 N (Claim N / Figure N with en).
    Returns None for non-headings."""
    t = text.strip()
    if not t:
        return None
    m = re.match(r'^【\s*(.+?)\s*】\s*$', t)
    if not m:
        return None
    inner_raw = m.group(1).strip()
    if en:
        mnum = re.match(r'^(claim|figure|fig\.?)\s*(\d+)\s*([a-z]?)$',
                        inner_raw, re.I)
        if mnum:
            num = mnum.group(2) + mnum.group(3).upper()
            return ('claim' if mnum.group(1).lower() == 'claim'
                    else 'figure', num)
        key = _norm_en(inner_raw)
        if key in HEAD2XML_EN:
            return (HEAD2XML_EN[key], None)
        return None
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


def split_glued_headings(paras, en=False):
    """Split paragraphs where a patent 【section】 marker is glued to body
    text. A 【…】 that canonizes to a section is a heading even without the
    line break around it ("...한다.【대표도】" or "【대표도】도 6"): unsplit,
    it degrades to plain […] via unmark() and the section is silently lost —
    a missing 【대표도】 then makes the abstract pick a wrong default figure.
    (runs, images, pp) tuples; an image follows the fragment holding its
    anchor offset."""
    out = []
    for runs, images, pp in paras:
        text = ''.join(r[0] for r in runs)
        pts = [(m.start(), m.end()) for m in re.finditer(r'【[^【】]*】', text)
               if canon_xml_section(m.group(0), en) is not None
               or (not en and canon_element(m.group(0)) is not None)]
        if not pts:
            out.append((runs, images, pp))
            continue
        bounds = sorted({0, len(text)} | {p for ab in pts for p in ab})
        for a, b in zip(bounds, bounds[1:]):
            sub, off = [], 0
            for text2, m, sz, c in runs:
                n = off + len(text2)
                lo, hi = max(off, a), min(n, b)
                if hi > lo:
                    sub.append((text2[lo - off:hi - off], m, sz, c))
                off = n
            imgs = [(t, cx, cy, p - a) for t, cx, cy, p in images
                    if a <= p < b or (b == len(text) and p == len(text))]
            out.append((sub, imgs, pp))
    return out


def _esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


# 제출파일 변환 rejects a sub/superscript that is also bold, italic or
# underlined: "E-001 문자 스타일 중 BOLD, Italic, UnderLine과 아래첨자를 같이
# 사용 할 수 없습니다" (an English PCT spec: an italic N with italic subscript CSI-PORT).
# The script keeps its position; the B/I/U is dropped and reported.
SCRIPT_FIXED = []         # text of the runs whose B/I/U was dropped


def _script_plain(text, mask):
    if mask & (STYLE_SUB | STYLE_SUPER) and \
            mask & (STYLE_BOLD | STYLE_ITALIC | STYLE_UNDERLINE):
        SCRIPT_FIXED.append(text)
        mask &= ~(STYLE_BOLD | STYLE_ITALIC | STYLE_UNDERLINE)
    return mask


def runs_to_xml(runs):
    """Render a paragraph's runs to PatentCAFDOC inline markup.
    Applies <sub>/<sup>/<u>/<b>/<i> per run style; <br/> handled by caller."""
    out = []
    for text, mask, sz, color in runs:
        if not text:
            continue
        mask = _script_plain(text, mask)
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


def _emit_figure_image(data, base, cx=None, cy=None, label=''):
    """Write one figure image for the KIPO XML. Returns (filename,
    img-format). IMAGE_CONVERT on (default): every image becomes a 1-bit
    300dpi Group-4 TIFF (normalize_g4). Off (--no-image-convert, for a
    domestic filing that keeps color or gray images): TIFF and JPEG pass
    through unchanged and other formats get the minimum re-encode the XML
    needs (near-bilevel art -> Group-4 TIFF, else JPEG q92)."""
    if IMAGE_CONVERT:
        tif = normalize_g4(data, cx, cy, label)
        with open(base + '.tif', 'wb') as fh:
            fh.write(tif)
        return os.path.basename(base) + '.tif', 'tif'
    if data[:4] in (b'II*\x00', b'MM\x00*'):
        with open(base + '.tif', 'wb') as fh:
            fh.write(data)
        try:
            if Image.open(io.BytesIO(data)).mode != '1':
                ASIS_NOT_BW.append(label)
        except Exception:
            pass
        return os.path.basename(base) + '.tif', 'tif'
    if data[:3] == b'\xff\xd8\xff':
        with open(base + '.jpg', 'wb') as fh:
            fh.write(data)
        ASIS_NOT_BW.append(label)
        return os.path.basename(base) + '.jpg', 'jpg'
    im = open_media(data)
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
    ASIS_NOT_BW.append(label)
    return os.path.basename(base) + '.jpg', 'jpg'


def build_patent_xml(src, paras, outdir, imgfiles, xml_name='spec.xml',
                     kind='spec'):
    """Emit a KIPO PatentCAFDOC xml (kind 'spec') or a WIPO application-body
    xml (kind 'pct', 국문 PCT) + pat* images into outdir.
    Returns the xml path, or None if the doc isn't patent-sectioned."""
    pct = kind in ('pct', 'pct_en')
    en = kind == 'pct_en'
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

    def strip_lead(runs, n):
        """runs minus their first n characters."""
        out = []
        for r in runs:
            if n >= len(r[0]):
                n -= len(r[0])
                continue
            out.append((r[0][n:],) + tuple(r[1:]))
            n = 0
        return out

    container = None        # 'claims' / 'drawings' / 'abstract' while cur == ''
    dropped = []            # (container, text) body text with no home
    unknown_heads = []      # (【heading】, section it fell into)

    for runs, images, pp in paras:
        text = ''.join(r[0] for r in runs)
        sec = canon_xml_section(text, en)
        if pct and not en:
            mh = re.match(r'^【\s*(.+?)\s*】\s*$', text.strip())
            if mh and _norm_tag(mh.group(1)) in PCT_HEAD2XML:
                sec = (PCT_HEAD2XML[_norm_tag(mh.group(1))], None)
        if sec is not None:
            bucket, num = sec
            seen_heading = True
            cur_claim = None
            cur_fig = None
            container = None
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
                container = bucket
            else:
                cur = bucket
            continue
        # content paragraph
        if not text.strip() and not images:
            # blank line: keep section spacing but XML ignores empties
            continue
        mu = re.match(r'^【\s*([^【】]+?)\s*】$', text.strip())
        if mu and not re.match(r'(수학식|표|화학식|반응식|equation|math|table|'
                               r'formula|chemistry)\s*[\dA-Za-z]*$',
                               mu.group(1), re.I):
            # a 【heading】 the form has no element for: it stays a plain
            # [heading] line inside the current section - say so
            unknown_heads.append((mu.group(0), cur or container or '?'))
        mclaim = (re.match(r'^\s*(\d+)\s*\.\s*(?=\S)', text)
                  if container == 'claims' or (cur == 'claim' and cur_claim
                                               and cur_claim.get('auto'))
                  else None)
        if mclaim and int(mclaim.group(1)) != (
                int(cur_claim['num']) + 1 if cur == 'claim' else 1):
            mclaim = None       # "2." inside a claim body is not claim 2
        if mclaim:
            # claim set without 【청구항 N】/【Claim N】 headings: "1. A method
            # ..." paragraphs under 【CLAIMS】 (a filed English PCT spec) - each number opens
            # a claim, the number itself becomes the claim's num
            cur_claim = {'num': mclaim.group(1), 'paras': [], 'auto': True}
            claims.append(cur_claim)
            cur = 'claim'
            container = None
            runs = strip_lead(runs, mclaim.end())
            text = text[mclaim.end():]
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
        elif text.strip():
            dropped.append((container or cur or '?', text.strip()))

    if not seen_heading:
        return None                     # not a patent-sectioned doc
    sec_name = {
        'title': '발명의 명칭', 'techfield': '기술분야',
        'background': '배경기술', 'disclosure': '발명의 내용',
        'problem': '해결하려는 과제', 'solution': '과제의 해결 수단',
        'effects': '발명의 효과', 'drawdesc': '도면의 간단한 설명',
        'bestmode': '최선의 형태', 'embodiments': '발명의 실시 내용',
        'industrial': '산업상 이용가능성', 'claims': '청구범위',
        'claim': '청구항', 'drawings': '도면', 'figure': '도',
        'abstract': '요약서', 'summary': '요약', 'absfig': '대표도'}
    for where, t in dropped:
        # body text right under a container heading (【청구범위】, 【도면】,
        # 【요약서】) has no XML slot - it used to vanish without a word
        WARNINGS.append('【%s】 바로 아래 문단은 넣을 자리가 없어 빠짐: "%s"'
                        % (sec_name.get(where, where), t[:60]))
    for h, where in unknown_heads:
        WARNINGS.append('지원하지 않는 제목 %s - 항목으로 만들지 않고 [ ] 본문 '
                        '줄로 앞 섹션(【%s】)에 넣음'
                        % (h, sec_name.get(where, where)))

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
            if emitted_files[tgt] is None:
                return ''
            fname, fmt = emitted_files[tgt]
        else:
            file_seq[0] += 1
            base = os.path.join(outdir, 'pat%05d' % file_seq[0])
            try:
                fname, fmt = _emit_figure_image(data, base, cx, cy,
                                                _img_label(tgt))
            except Exception as e:
                # an image Pillow can't open (e.g. a bare SVG) used to be
                # written raw under a .jpg name - leave it out and say so
                WARNINGS.append('그림 %s: 열 수 없는 형식이라 빠짐 (%s)'
                                % (_img_label(tgt), type(e).__name__))
                emitted_files[tgt] = None
                return ''
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
        src_imgs = f['imgs']
        if pct and len(src_imgs) > 1:
            # ST.36 <figure> is (img): one image per 도
            WARNINGS.append('도 %s: 이미지 %d개 중 첫 번째만 넣음 (PCT 도면은 '
                            '도마다 이미지 1개)' % (f['num'], len(src_imgs)))
            src_imgs = src_imgs[:1]
        imgs = [img_ref(t, cx, cy) for (t, cx, cy, _p) in src_imgs]
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

    if pct:
        return _write_pct_xml(outdir, xml_name, title, buckets, claims,
                              fig_xml, absfig or absfig_img, sect, paras_xml,
                              counter, runs_with_imgs,
                              lang='en' if en else 'ko')

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


def _write_pct_xml(outdir, xml_name, title, buckets, claims, fig_xml,
                   had_absfig, sect, paras_xml, counter, runs_with_imgs,
                   lang='ko'):
    """Assemble the 국문/영문 PCT application-body xml from
    build_patent_xml's parsed sections (same buckets/claims/figures as the
    domestic path)."""
    # text right under 【발명의 내용】 before any sub-heading: <disclosure> is
    # either sub-elements or plain p's, not both — fold it into the first
    # sub-section present, else emit it as the disclosure's own p's
    direct = buckets.pop('disclosure', [])
    if direct:
        for b in SOI:
            if buckets.get(b):
                buckets[b] = direct + buckets[b]
                direct = []
                break
    if direct:
        buckets['disclosure'] = direct
    if had_absfig:
        WARNINGS.append('【대표도】 무시함 (PCT 서식에는 대표도 항목이 없음)')

    L = ['<?xml version="1.0" encoding="euc-kr" ?>',
         '<!DOCTYPE application-body SYSTEM "application-body-v1-6.dtd" >',
         '<application-body lang="%s" dtd-version="1.6" country="KR">' % lang,
         '<description>',
         '<invention-title>%s</invention-title>' % _esc(' '.join(title)),
         sect('technical-field', 'techfield'),
         sect('background-art', 'background')]
    soi = (sect('tech-problem', 'problem') +
           sect('tech-solution', 'solution') +
           sect('advantageous-effects', 'effects'))
    if soi:
        L.append('<disclosure>%s</disclosure>' % soi)
    elif direct:
        L.append(sect('disclosure', 'disclosure'))
    for b in ('drawdesc', 'bestmode', 'embodiments', 'industrial'):
        L.append(sect(PCT_TAG[b], b))
    L.append('</description>')
    # the PCT import drops <br/> inside claim-text (claim lines run together:
    # "메모리; 및프로세싱 ..."); one sibling <claim-text> per line keeps them
    L.append('<claims>')
    for c in claims:
        lines = [runs_with_imgs(r, i) for r, i in c['paras']]
        body = ''.join('<claim-text>%s</claim-text>' % s
                       for s in lines if s.strip())
        L.append('<claim num="%s">%s</claim>' %
                 (_esc(c['num']), body or '<claim-text></claim-text>'))
    L.append('</claims>')
    abst = paras_xml('summary', counter)
    if abst:
        L.append('<abstract>%s</abstract>' % abst)
    if fig_xml:
        L.append('<drawings>')
        L.extend(fig_xml)
        L.append('</drawings>')
    L.append('</application-body>')
    xp = os.path.join(outdir, xml_name)
    with open(xp, 'wb') as f:
        f.write('\n'.join(L).encode('cp949', 'replace'))
    return xp


# --- OpinionDOC (의견서) XML emission ---------------------------------------
# Vocabulary from an NKEditor-exported 의견서 and DOCCommon.xsd:
#   <KIPO><OpinionDOC><Opinion> (P | BR)* </Opinion></OpinionDOC></KIPO>
#   <P indent="14|0" align="JUSTIFIED|LEFT|CENTER|RIGHT"> text, <B> <I> <U>
#     <SB>sub <SP>sup, <BR/>, <EMI id height width file/> (mm, opnNNNNN.jpg)
#   tables: <P indent="0" align="LEFT"><table><tgroup xmlns=OASIS cols>
#     <colspec colnum colname colwidth/> (1/1800 inch) <tbody><row>
#     <entry align colname | namest nameend | morerows> ... </entry>
# The Opinion element itself is the 【의견내용】 item, so a docx 【의견내용】
# line is dropped. Indent follows docx2hml's rule (verified against the
# firm's filed 의견서): 【...】 lines, bare [청구항 N] and 인용발명N captions
# 0, all other paragraphs 14.
OPN_TABLE_WIDTH = 11686       # full text width in colwidth units (165mm)
OASIS_NS = 'http://www.oasis-open.org/tables/exchange/1.0'


def runs_to_opn(runs):
    out = []
    for text, mask, sz, color in runs:
        if not text:
            continue
        mask = _script_plain(text, mask)
        t = _esc(text)
        if mask & STYLE_SUB:
            t = '<SB>' + t + '</SB>'
        elif mask & STYLE_SUPER:
            t = '<SP>' + t + '</SP>'
        if mask & STYLE_UNDERLINE:
            t = '<U>' + t + '</U>'
        if mask & STYLE_BOLD:
            t = '<B>' + t + '</B>'
        if mask & STYLE_ITALIC:
            t = '<I>' + t + '</I>'
        out.append(t)
    return ''.join(out)


def build_opinion_xml(src, blocks, outdir, imgfiles, xml_name='opinion.xml'):
    """Emit a KIPO OpinionDOC xml + opn* JPEG images into outdir.
    blocks: parse_docx(..., blocks=True) output. Returns the xml path."""
    os.makedirs(outdir, exist_ok=True)
    media = {}
    with zipfile.ZipFile(src) as zf:
        for name in zf.namelist():
            if name.startswith('word/media/'):
                media[name] = zf.read(name)
    seq = [0]

    def emi(tgt, cx, cy):
        data = media.get(tgt)
        if data is None:                 # rendered equation ('@eqN')
            try:
                with open(imgfiles[tgt], 'rb') as fh:
                    data = fh.read()
            except Exception:
                return ''
        # 의견서 images are not normalized to G4 (user rule 2026-11-05):
        # they stay JPEG as before, color included
        seq[0] += 1
        fname = 'opn%05d.jpg' % seq[0]
        path = os.path.join(outdir, fname)
        if data[:3] == b'\xff\xd8\xff':
            with open(path, 'wb') as fh:
                fh.write(data)
        else:
            im = open_media(data)
            if im.mode in ('RGBA', 'LA', 'P'):
                bg = Image.new('RGB', im.size, 'white')
                im = im.convert('RGBA')
                bg.paste(im, mask=im.split()[-1])
                im = bg
            elif im.mode != 'RGB':
                im = im.convert('RGB')
            im.save(path, 'JPEG', quality=92)
        wi = max(1, int(round((cx or 0) / 36000.0)))
        he = max(1, int(round((cy or 0) / 36000.0)))
        if wi > 165 or he > 222:         # KIPO page box, keep aspect
            s = min(165.0 / wi, 222.0 / he)
            wi, he = max(1, int(wi * s)), max(1, int(he * s))
        return '<EMI id="%d" height="%d" width="%d" file="%s"/>' % (
            seq[0], he, wi, fname)

    def inline(runs, images):
        if not images:
            return runs_to_opn(runs)
        out, off = [], 0
        pending = sorted(images, key=lambda i: i[3])
        for r in runs:
            while pending and pending[0][3] <= off:
                t, cx, cy, _ = pending.pop(0)
                out.append(emi(t, cx, cy))
            out.append(runs_to_opn([r]))
            off += len(r[0])
        out.extend(emi(t, cx, cy) for t, cx, cy, _ in pending)
        return ''.join(out)

    def align(pp, default):
        jc = (pp or {}).get('jc')
        return {PALIGN_CENTER: 'CENTER', PALIGN_RIGHT: 'RIGHT'}.get(jc, default)

    def table(t):
        rows = t['rows']
        ncol = max([len(t['grid'])] +
                   [sum(c['span'] for c in r) for r in rows] or [1])
        grid = (t['grid'] + [0] * ncol)[:ncol]
        if not any(grid):
            grid = [1] * ncol
        grid = [g or max(1, sum(grid) // ncol) for g in grid]
        # dxa (1/1440 in) -> 1/1800 in; shrink a table wider than the page
        widths = [g * 1800 // 1440 for g in grid]
        if sum(widths) > OPN_TABLE_WIDTH:
            widths = [w * OPN_TABLE_WIDTH // sum(widths) for w in widths]
        # vertical merges: a 'restart' cell spans every 'continue' cell
        # below it in the same grid column
        pos = []                          # per row: [(col, cell)]
        for r in rows:
            col, lst = 1, []
            for c in r:
                lst.append((col, c))
                col += c['span']
            pos.append(lst)
        L = ['<P indent="0" align="LEFT">',
             "<table><tgroup xmlns='%s' cols=\"%d\">" % (OASIS_NS, ncol)]
        for i, w in enumerate(widths):
            L.append('<colspec colnum="%d" colname="col%d" colwidth="%d"/>'
                     % (i + 1, i + 1, max(1, w)))
        L.append('<tbody>')
        for ri, lst in enumerate(pos):
            L.append('<row>')
            for col, c in lst:
                if c['vmerge'] == 'continue':
                    continue
                more = 0
                if c['vmerge'] == 'restart':
                    for below in pos[ri + 1:]:
                        if any(bc == col and b['vmerge'] == 'continue'
                               for bc, b in below):
                            more += 1
                        else:
                            break
                parts = [inline(runs, imgs) for runs, imgs, _pp in c['paras']]
                while parts and not parts[-1].strip():
                    parts.pop()
                first_pp = c['paras'][0][2] if c['paras'] else None
                a = {'CENTER': 'center', 'RIGHT': 'right'}.get(
                    align(first_pp, ''), 'justify')
                attrs = 'align="%s"' % a
                if c['span'] > 1:
                    attrs += ' namest="col%d" nameend="col%d"' % (
                        col, col + c['span'] - 1)
                else:
                    attrs += ' colname="col%d"' % col
                if more:
                    attrs += ' morerows="%d"' % more
                L.append('<entry %s>%s</entry>' % (attrs, '<BR/>'.join(parts)))
            L.append('</row>')
        L.append('</tbody></tgroup></table></P>')
        return '\n'.join(L)

    body = []
    for kind, item in blocks:
        if kind == 'tbl':
            if item['rows']:
                body.append(table(item))
            continue
        runs, images, pp = item
        text = ''.join(r[0] for r in runs).strip()
        if text == '【의견내용】' and not images:
            continue
        if not text and not images:
            body.append('<BR/>')
            continue
        flush = (text.startswith('【') or
                 re.match(r'^\[청구항\s*\d+\]$', text) or
                 re.match(r'^인용발명\d', text))
        body.append('<P indent="%d" align="%s">%s</P>' % (
            0 if flush else 14, align(pp, 'JUSTIFIED'), inline(runs, images)))
    # leading blank lines would sit right under the 【의견내용】 marker
    while body and body[0] == '<BR/>':
        body.pop(0)
    xml = '\n'.join(['<?xml version="1.0" encoding="euc-kr" ?>',
                     '<KIPO keapsVersion="4.02" editorKind="K" pageCount="1" '
                     'xmlns="http://www.kipo.go.kr">',
                     '<OpinionDOC>', '<Opinion>'] + body +
                    ['</Opinion>', '</OpinionDOC>', '</KIPO>'])
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


def open_media(data):
    """Pillow image of docx media bytes. WMF/EMF (OLE previews such as
    MathType equations) load at 72 dpi by default - a 26x18 px blot once
    thresholded - so they are rasterized at 600 dpi on white."""
    im = Image.open(io.BytesIO(data))
    if im.format in ('WMF', 'EMF') or data[:4] == b'\xd7\xcd\xc6\x9a' \
            or data[40:44] == b' EMF':
        im.load(dpi=600)
        if im.mode in ('RGBA', 'LA', 'P'):
            bg = Image.new('RGB', im.size, 'white')
            bg.paste(im.convert('RGBA'), mask=im.convert('RGBA').split()[3])
            im = bg
    else:
        im.load()
    return im


G4_DPI = 300                  # KIPO submission images: 1-bit, 300dpi, Group 4
PAGE_MM = (165.0, 222.0)      # KIPO page box (E-218)
# image normalization record of the XML build: {'ok': [label], 'fixed':
# [note]}; the flat path's staging keeps its own until it is the one used
IMG_REPORT = {'ok': [], 'fixed': []}
FLAT_IMG_REPORT = {'ok': [], 'fixed': []}
# 이미지 자동 변환: on -> every spec/PCT/flat image becomes 1-bit 300dpi G4
# TIFF; off (--no-image-convert / GUI checkbox) -> images go in as they are
# (domestic filings may carry color). 의견서 images are never normalized.
IMAGE_CONVERT = True
ASIS_NOT_BW = []              # with conversion off: gray/color images kept


def _img_label(tgt):
    """Short name of a media target for reports ('image3.png', '수식 7')."""
    if tgt.startswith('@eq'):
        return '수식 ' + tgt[3:]
    if tgt.startswith('@tbl'):
        return '표 ' + tgt[4:]
    return os.path.basename(tgt)


def _img_desc(im, data):
    """'PNG RGB 96dpi 400x300' style description of a source image."""
    fmt = im.format or ('TIFF' if data[:4] in (b'II*\x00', b'MM\x00*') else
                        'WMF' if data[:4] == b'\xd7\xcd\xc6\x9a' else
                        'EMF' if data[40:44] == b' EMF' else '?')
    comp = im.info.get('compression')
    dpi = _src_dpi(im)
    bits = {'1': '1비트', 'L': '회색조', 'P': '팔레트', 'RGB': '컬러',
            'RGBA': '컬러(투명)', 'LA': '회색조(투명)', 'CMYK': 'CMYK'
            }.get(im.mode, im.mode)
    parts = [fmt, bits]
    if fmt == 'TIFF':
        parts.append({'group4': 'G4', 'raw': '무압축', 'tiff_lzw': 'LZW',
                      'group3': 'G3', 'packbits': 'PackBits',
                      'tiff_adobe_deflate': 'Deflate', 'jpeg': 'JPEG압축'
                      }.get(comp, comp or '무압축'))
    parts.append('%ddpi' % round(dpi[0]) if dpi and dpi[0] else 'dpi없음')
    parts.append('%dx%d' % im.size)
    return ' '.join(parts)


def _src_dpi(im):
    """(xdpi, ydpi) of a source image, None when absent - Pillow reports a
    TIFF without ResolutionUnit as 1 dpi, which is no resolution at all."""
    dpi = im.info.get('dpi')
    if dpi and len(dpi) == 2 and dpi[0] >= 50 and dpi[1] >= 50:
        return dpi
    return None


def _is_g4_300(im, data):
    dpi = _src_dpi(im)
    return (data[:4] in (b'II*\x00', b'MM\x00*') and im.mode == '1'
            and im.info.get('compression') == 'group4'
            and bool(dpi) and abs(dpi[0] - G4_DPI) < 1
            and abs(dpi[1] - G4_DPI) < 1
            and getattr(im, 'n_frames', 1) == 1)


def normalize_g4(data, cx=None, cy=None, label='', report=None):
    """Return docx image bytes as a 1-bit 300dpi Group-4 TIFF - the one
    image format of a KIPO submission (equations, tables and drawings
    alike). A TIFF that already is single-page 1-bit G4 at 300dpi passes
    through byte for byte; anything else (PNG/JPEG/GIF/BMP/WMF/EMF, and
    TIFFs that are gray, color, uncompressed, LZW, G3 or not 300dpi) is
    rebuilt and listed in the review report:
      - transparency is flattened onto white;
      - size: the display extent (cx, cy EMU, clamped to the 165x222mm page
        box like the XML he/wi) at 300dpi; without an extent the source
        dpi, without a dpi the source pixels. Within 3% of the current size
        only the dpi tag changes (no resampling of finished drawings);
      - line art (<25% midtone pixels: drawings, equations, tables,
        antialiased text) is thresholded on the darkest channel, so thin
        or light-colored lines stay black; downscaling lowers the cutoff
        towards white so 1px lines survive;
      - continuous-tone content (photos, gradients) is Floyd-Steinberg
        dithered from luminance and reported, since it changes look.
    report: dict collecting 'ok' labels / 'fixed' notes (IMG_REPORT)."""
    from PIL import ImageChops
    if report is None:
        report = IMG_REPORT
    im = open_media(data)
    if _is_g4_300(im, data):
        report['ok'].append(label)
        return data
    before = _img_desc(im, data)
    src_dpi = _src_dpi(im)
    if getattr(im, 'n_frames', 1) > 1:
        im.seek(0)
    if im.mode in ('RGBA', 'LA', 'P', 'PA') or 'transparency' in im.info:
        rgba = im.convert('RGBA')
        bg = Image.new('RGB', im.size, 'white')
        bg.paste(rgba, mask=rgba.split()[3])
        im = bg
    elif im.mode == '1':
        im = im.convert('L')
    elif im.mode not in ('L', 'RGB'):
        im = im.convert('RGB')          # CMYK, I;16, ...

    # target pixel size at 300dpi
    w0, h0 = im.size
    tw, th = w0, h0
    if cx and cy and cx > 0 and cy > 0:
        wmm, hmm = cx / 36000.0, cy / 36000.0
        if wmm > PAGE_MM[0] or hmm > PAGE_MM[1]:
            s = min(PAGE_MM[0] / wmm, PAGE_MM[1] / hmm)
            wmm, hmm = wmm * s, hmm * s
        tw = max(1, int(round(wmm / 25.4 * G4_DPI)))
        th = max(1, int(round(hmm / 25.4 * G4_DPI)))
    else:
        if src_dpi and src_dpi[0] and src_dpi[1]:
            tw = max(1, int(round(w0 * G4_DPI / float(src_dpi[0]))))
            th = max(1, int(round(h0 * G4_DPI / float(src_dpi[1]))))
    if abs(tw - w0) <= 0.03 * w0 and abs(th - h0) <= 0.03 * h0:
        tw, th = w0, h0

    lum = im.convert('L')
    hist = lum.histogram()
    tot = w0 * h0 or 1
    photo = sum(hist[56:200]) > 0.25 * tot
    if photo:
        g = lum if (tw, th) == (w0, h0) else lum.resize((tw, th), Image.LANCZOS)
        out = g.convert('1')            # Floyd-Steinberg dither
        how = '디더링(사진/명암 그림)'
    else:
        if im.mode == 'RGB':
            r, gch, b = im.split()
            g = ImageChops.darker(ImageChops.darker(r, gch), b)
        else:
            g = lum
        scale = min(tw / float(w0), th / float(h0))
        if (tw, th) != (w0, h0):
            g = g.resize((tw, th), Image.LANCZOS)
        cut = 128 if scale >= 1 else min(200, int(255 * (1 - 0.5 * scale)))
        out = g.point(lambda v, c=cut: 255 if v >= c else 0, '1')
        how = '흑백 변환'
    buf = io.BytesIO()
    out.save(buf, 'TIFF', compression='group4', dpi=(G4_DPI, G4_DPI))
    report['fixed'].append('그림 %s: %s -> 1비트 G4 300dpi %dx%d (%s)' % (
        label or '?', before, tw, th, how))
    return buf.getvalue()


def image_report_lines(report):
    """Review-report lines for one IMG_REPORT: a count line, then every
    image that had to be converted (repeats of one image listed once)."""
    fixed = list(dict.fromkeys(report['fixed']))
    ok = len(set(report['ok']))
    if not fixed:
        return []
    return (['그림 %d개를 1비트 G4 300dpi TIFF로 바꿈 (이미 규격에 맞는 그림 %d개는 '
             '그대로 둠)' % (len(fixed), ok)] + fixed)


def media_to_tiff(data, out_tif, label=''):
    """Stage image bytes as a TIFF at out_tif for the flat TerConv path (the
    TE save path embeds TIFF verbatim but must ImageGear-decode PNG/JPEG,
    which fails on palette/odd PNGs). IMAGE_CONVERT on: a 1-bit 300dpi
    Group-4 TIFF; off: an uncompressed TIFF in the source's own mode.
    Returns True on success."""
    if IMAGE_CONVERT:
        try:
            tif = normalize_g4(data, label=label, report=FLAT_IMG_REPORT)
            with open(out_tif, 'wb') as fh:
                fh.write(tif)
            return True
        except Exception:
            return False
    try:
        im = open_media(data)
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
REVIEW_ISSUES = []        # why .rtf / .compare.docx were not produced
REVIEW_NOTES = []         # .compare.docx made, but part of it needs a look


def _decode_output(b):
    """Child-process output -> str. DocCmp's .NET console writes UTF-8 or
    CP949 depending on the parent console's code page, so decoding with the
    locale codec alone raises on Korean paths."""
    if not b:
        return ''
    for enc in ('utf-8', 'cp949'):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            pass
    return b.decode('utf-8', errors='replace')


def _clock(sec):
    return '%d분 %02d초' % divmod(int(sec), 60)


class Stage:
    """`with Stage('NKEditor에서 .hlt 저장'):` - prints the step and keeps
    its elapsed time ticking on one line until the block ends, so a long
    NKEditor/Word step does not look like a hang."""

    def __init__(self, label):
        self.label = label

    def __enter__(self):
        import threading
        self.t0 = time.time()
        self.stop = threading.Event()

        def tick():
            while not self.stop.wait(2):
                sys.stdout.write('\r  %s 중... (%s 경과)   '
                                 % (self.label, _clock(time.time() - self.t0)))
                sys.stdout.flush()
        sys.stdout.write('  %s 중...' % self.label)
        sys.stdout.flush()
        self.th = threading.Thread(target=tick, daemon=True)
        self.th.start()
        return self

    def __exit__(self, *exc):
        self.stop.set()
        self.th.join()
        sys.stdout.write('\r  %s %s (%s)%s\n'
                         % (self.label, '실패' if exc[0] else '끝',
                            _clock(time.time() - self.t0), ' ' * 20))
        sys.stdout.flush()
        return False


def run_with_progress(cmd, limit, imgdir, n_math, n_tbl):
    """subprocess.run(cmd, capture_output, text, timeout=limit), but showing
    one live status line meanwhile: EqRender prints nothing until it ends,
    and a long spec (788 equations) looked hung for minutes. The
    counts are the eqNNNN/tblNNNN files written so far."""
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, text=True,
                            errors='replace')
    t0 = time.time()
    shown = False
    while True:
        try:
            out, err = proc.communicate(timeout=2)
            break
        except subprocess.TimeoutExpired:
            if time.time() - t0 > limit:
                proc.kill()
                proc.communicate()
                if shown:
                    print(flush=True)
                raise
            try:
                names = os.listdir(imgdir)
            except OSError:
                names = []
            ne = sum(1 for n in names if n.startswith('eq')
                     and n.endswith('.tif'))
            nt = sum(1 for n in names if n.startswith('tbl')
                     and n.endswith('.tif'))
            parts = []
            if n_math:
                parts.append('수식 %d/%d' % (min(ne, n_math), n_math))
            if n_tbl:
                parts.append('표 %d/%d' % (min(nt, n_tbl), n_tbl))
            sys.stdout.write('\r  그림 변환 중: %s  (%s 경과)   '
                             % (', '.join(parts), _clock(time.time() - t0)))
            sys.stdout.flush()
            shown = True
    if shown:
        sys.stdout.write('\r  그림 변환 끝: %s 걸림%s\n'
                         % (_clock(time.time() - t0), ' ' * 30))
        sys.stdout.flush()
    return subprocess.CompletedProcess(cmd, proc.returncode, out, err)


def render_objects(src, imgdir, tables=True):
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
    (equations) / flattened cells (tables).

    tables=False: the caller emits tables natively (의견서) - only equations
    are wanted, and a doc without equations skips Word entirely."""
    empty = ({}, {})
    try:
        with zipfile.ZipFile(src) as zf:
            xml = zf.read('word/document.xml')
        n_math = len(main_story_omaths(ET.fromstring(xml)))
        n_tbl = (xml.count(b'<w:tbl>') + xml.count(b'<w:tbl ')) if tables else 0
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
        # Word renders ~1 object per 5-10 s on a long spec (e.g. 148
        # equations + 25 tables); a flat 600 s cut it off mid-way, and every
        # retry left another hidden WINWORD behind
        limit = max(600, 15 * (n_math + n_tbl) + 120)
        what = ', '.join(s for s in (
            '수식 %d개' % n_math if n_math else '',
            '표 %d개' % n_tbl if n_tbl else '') if s)
        print('Word로 %s를 그림으로 변환합니다 (개수가 많으면 몇 분 걸림)' % what,
              flush=True)
        for attempt in range(4):
            try:
                r = run_with_progress([EQRENDER, src, imgdir], limit,
                                      imgdir, n_math, n_tbl)
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
                if not tables:
                    tbls, t_expect = {}, None
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
                # slowness is not transient: another run would only time out
                # again (and orphan another Word)
                last_err = 'timeout %ds' % limit
                print('WARNING: object render incomplete (%s)' % last_err,
                      file=sys.stderr)
                break
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


def parse_docx(path, imgdir, eqs=None, tbls=None, blocks=False):
    """Parse docx; extract media to imgdir (as TIFF). Returns (paras, {mediaTarget: path}).
    eqs: {omath_ordinal: (tif, w_mm, h_mm)} from render_objects — spliced
    into paragraphs where each m:oMath appears.
    tbls: {ordinal: (tif, w_mm, h_mm)} rendered top-level tables — each
    becomes an image-only paragraph at the w:tbl's document position.
    blocks: return iter_blocks() output (tables kept as tables) instead."""
    imgfiles = {}
    tstate = [0]                          # top-level w:tbl ordinal counter
    with zipfile.ZipFile(path) as zf:
        styles = load_styles(zf)
        rels = load_rels(zf)
        NUMBERING.clear()
        NUMBERING.update(load_numbering(zf, styles))
        os.makedirs(imgdir, exist_ok=True)
        for name in zf.namelist():
            if name.startswith('word/media/'):
                base = os.path.basename(name)
                data = zf.read(name)
                # stage as TIFF so the HIF embed path works for every format
                # (append .tif to the full basename to keep names unique)
                tif = os.path.join(imgdir, base + '.tif')
                if HAVE_PIL and media_to_tiff(data, tif, base):
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
    if blocks:
        return list(iter_blocks(body, styles, rels, imgfiles, eqstate)), imgfiles
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


# --- document kind ---------------------------------------------------------
# NKEditor builds a different KIPO document from each XML root (verified by
# importing and saving one of each):
#   <KIPO><PatentCAFDOC>        -> 국내 특허 명세서
#   <application-body lang=ko>  -> 국문 PCT 국제출원 명세서 (【청구의 범위】,
#                                  【기술적 해결방법】 - the element set of the
#                                  filed WO0 .hlt files; header kind 03)
#   <KIPO><OpinionDOC>          -> 의견서 (【의견내용】, header kind 07)
# PCTPatentCAFDOC / PCTPatentDOC do NOT give the PCT form (the first builds a
# domestic spec, the second is pasted in as raw text). 보정서 .hlt files are
# made from the .fin, not from a docx, so an amendment docx is only detected
# (so its 【청구항 N】 lines can't pass it off as a spec) and goes flat.
KIND_LABEL = {'spec': '국내 특허 명세서', 'pct': '국문 PCT 명세서',
              'pct_en': '영문 PCT 명세서',
              'opinion': '의견서', 'amend': '보정서', 'flat': '일반 문서'}
# headings only the PCT international-application form uses
PCT_ONLY_HEADS = ('청구의범위', '기술적해결방법', '발명의실시를위한형태',
                  '발명의실시를위한최선의형태', '산업상이용가능성')
AMEND_HEADS = ('보정대상항목', '보정방법', '보정내용')
# case tokens in the path. Only KIPO filings become .hlt - the PCT
# international application (SP00001-WO0, ...-WO_SP.., SP00003-PCT) and the
# domestic application (P0000001-KR0) - so other country codes are not looked
# at. PCT/WO wins over KR: a PCT draft's filename may say "...-KR0 (SP00003-PCT)".
PCT_TOKEN = re.compile(r'(?<![A-Za-z])(?:WO|PCT)\d*(?![A-Za-z])')
KR_TOKEN = re.compile(r'(?<![A-Za-z])KR\d*(?![A-Za-z])')


def docx_lines(src):
    """Plain text of every w:p (body, cells, text boxes) — cheap, for kind
    detection before the full parse. Tracked deletions are w:delText, so
    they don't show up."""
    with zipfile.ZipFile(src) as zf:
        root = ET.fromstring(zf.read('word/document.xml'))
    return [''.join(t.text or '' for t in p.iter(W + 't'))
            for p in root.iter(W + 'p')]


def detect_kind(src, lines):
    """-> (kind, reason). kind in KIND_LABEL."""
    opinion = amend = None
    heads, pct_heads, en_heads = set(), [], []
    for line in lines:
        s = line.strip()
        for m in re.finditer(r'【\s*([^【】]+?)\s*】', s):
            inner = _norm_tag(m.group(1))
            en_sec = canon_xml_section(m.group(0), en=True)
            if en_sec and en_sec[0] not in ('claim', 'figure') \
                    and m.group(0) not in en_heads:
                en_heads.append(m.group(0))
            if inner == '의견내용':
                opinion = opinion or m.group(0)
            elif inner in AMEND_HEADS and s.startswith(m.group(0)):
                amend = amend or m.group(0)
            if inner in PCT_ONLY_HEADS:
                heads.add(inner)
                if m.group(0) not in pct_heads:
                    pct_heads.append(m.group(0))
            elif canon_xml_section(m.group(0)) is not None:
                heads.add(inner)
    if opinion:
        return 'opinion', opinion
    if amend:
        return 'amend', amend
    # a KR 의견서 may quote 【청구항 N】 lines, and a 보정서 is often a bare
    # claim set — either would pass for a spec by its headings
    base = os.path.basename(src)
    if '의견서' in base:
        return 'opinion', '파일명의 "의견서"'
    if '보정서' in base:
        return 'amend', '파일명의 "보정서"'
    # English ST.36 headings (【Technical Field】, 【CLAIMS】, ...): only the
    # English PCT form uses them, whatever the case folder says
    if len(en_heads) >= 3 and len(en_heads) > len(heads):
        return 'pct_en', ', '.join(en_heads[:3])
    if heads:
        # case token first: a KR0 working draft can carry PCT headings
        # (a KR0 "Processing" docx) yet was filed as a domestic spec
        path = os.path.abspath(src)
        tok = PCT_TOKEN.search(path)
        if tok:
            return 'pct', '경로의 사건 토큰 %s' % tok.group(0)
        tok = KR_TOKEN.search(path)
        if tok:
            return 'spec', '경로의 사건 토큰 %s' % tok.group(0)
        if pct_heads:
            return 'pct', ', '.join(pct_heads[:3])
        return 'spec', '특허 섹션 제목'
    return 'flat', '특허 섹션 제목 없음'


def convert_structured(src, dst, paras, imgfiles, kind='spec'):
    """docx -> KIPO xml for `kind` (spec: PatentCAFDOC, pct: application-body,
    opinion: OpinionDOC; `paras` is parse_docx(blocks=True) output for
    opinion) -> NKEditor builds the structured doc -> drives Save-As to a
    real .hlt at `dst`. Returns True on success."""
    import nkwin
    # NKEditor 6.x crashes (heap corruption 0xc0000374) on Save-As when the
    # opened xml has a long Korean name (a WO0 draft); the save-as
    # filename follows the xml stem, so keep both short and move the .hlt to
    # `dst` afterwards
    stem = 'd2h' + os.urandom(3).hex()
    scratch = tempfile.mkdtemp(prefix='d2h_xml')
    del SCRIPT_FIXED[:]
    if kind == 'opinion':
        xmlpath = build_opinion_xml(src, paras, scratch, imgfiles,
                                    xml_name=stem + '.xml')
    else:
        xmlpath = build_patent_xml(src, paras, scratch, imgfiles,
                                   xml_name=stem + '.xml', kind=kind)
    if not xmlpath:
        return False
    if SCRIPT_FIXED:
        seen = []
        for s in SCRIPT_FIXED:
            if s not in seen:
                seen.append(s)
        WARNINGS.append('첨자 %d곳의 굵게/기울임/밑줄을 뺐음 (제출파일 변환 '
                        'E-001: 첨자와 같이 못 씀): %s'
                        % (len(SCRIPT_FIXED), ', '.join(seen[:8])))
    print('structured: emitted %s ->' % {
        'spec': 'PatentCAFDOC', 'pct': 'application-body',
        'pct_en': 'application-body (en)',
        'opinion': 'OpinionDOC'}[kind], xmlpath)
    with Stage('NKEditor에서 .hlt 저장'):
        saved = nkwin.nk_import_and_save_hlt(
            xmlpath, stem, out_dir=os.path.dirname(os.path.abspath(dst)))
    if not saved or not os.path.exists(saved):
        print('ERROR: NKEditor save-as did not produce a .hlt', file=sys.stderr)
        return False
    if os.path.abspath(saved) != os.path.abspath(dst):
        import shutil
        shutil.move(saved, dst)   # move — don't litter the dialog's folder
    if kind in PCT_KIND_TARGET:
        set_pct_intl_kind(dst, kind)
    print('structured hlt ->', dst)
    return True


# .hlt header byte 0x20 is the document kind. NKEditor builds an imported
# application-body as kind 03 (PctPat.frm, PCT번역문 제출서 = 국내진입), whose
# 제출파일 변환 fails with E-259 "특허법 시행규칙에 따라 작성되지 않았습니다".
# A PCT명세서(국제출원) 한글서식 document (PctPatHD.frm, saved from a new doc)
# is kind 0x48 with an otherwise identical header; the same import with only
# this byte set to 0x48 passes 제출파일 변환 (verified on a Korean PCT draft, 2026-10-28).
HLT_KIND_OFFSET = 0x20
HLT_KIND_PCT_NATIONAL = 0x03
HLT_KIND_PCT_INTL_KO = 0x48
# PCT명세서(국제출원) 영문서식: filed English PCT .hlt files are 0x49
HLT_KIND_PCT_INTL_EN = 0x49
PCT_KIND_TARGET = {
    'pct': (HLT_KIND_PCT_INTL_KO, 'PCT명세서(국제출원) 한글서식'),
    'pct_en': (HLT_KIND_PCT_INTL_EN, 'PCT명세서(국제출원) 영문서식'),
}


def set_pct_intl_kind(path, kind='pct'):
    target, label = PCT_KIND_TARGET[kind]
    with open(path, 'r+b') as f:
        head = f.read(HLT_KIND_OFFSET + 1)
        if not head.startswith(b'HAN Lite Document'):
            WARNINGS.append('PCT 문서 종류 설정 실패 - .hlt 머리가 아님')
            return
        k = head[HLT_KIND_OFFSET]
        if k == target:
            return
        if k != HLT_KIND_PCT_NATIONAL:
            # unknown kind: leave it, but say so - 제출파일 변환 may reject it
            WARNINGS.append('PCT 문서 종류값이 예상(03)과 다름(%02x) - %s(%02x)'
                            '으로 바꾸지 않음' % (k, label, target))
            return
        f.seek(HLT_KIND_OFFSET)
        f.write(bytes([target]))
    print('pct: document type 03 -> %02x (%s)' % (target, label))


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
    Failures never fail the conversion; they are printed and recorded in
    REVIEW_ISSUES. If NKEditor crashes under the rtf export, it is retried
    once.

    NKEditor 6.3 (2026.10.22.01) dies with heap corruption (0xc0000374)
    while saving RTF when the .hlt it opened has a long path — reproduced
    at 146+ CP949 bytes (a case-folder Sent path is ~150), never at 44-136.
    So the export opens a byte-identical copy under a short temp path; the
    rtf still reflects the saved .hlt file, not the import session."""
    import nkwin
    import shutil
    base = os.path.splitext(dst)[0]
    rtfdst = base + '.rtf'
    cmpdoc = base + '.compare.docx'

    def issue(msg):
        print('review: ' + msg, file=sys.stderr)
        REVIEW_ISSUES.append(msg)

    # unique stem: a same-named rtf left in the dialog's folder by an
    # earlier crashed run would trip nkwin's foreign-file guard
    tmpdir = tempfile.mkdtemp(prefix='d2hr')
    probe = os.path.join(tmpdir, 'd2h_%s.hlt' % os.urandom(3).hex())
    try:
        shutil.copyfile(dst, probe)
        rtf = None
        for attempt in (1, 2):
            try:
                with Stage('NKEditor에서 검토용 .rtf 내보내기'):
                    rtf = nkwin.nk_export_rtf(probe)
                why = nkwin.LAST_EXPORT_ERROR
                crashed = nkwin.LAST_EXPORT_CRASHED
            except Exception as e:
                rtf, why, crashed = None, repr(e), False
            if rtf and os.path.exists(rtf):
                break
            if attempt == 1 and crashed:
                print('review: rtf export - %s; retrying once' % why,
                      file=sys.stderr)
                time.sleep(3)      # let WER finish with the dead instance
                continue
            issue('rtf 내보내기 실패 - %s%s'
                  % (why or 'rtf 파일이 만들어지지 않음',
                     ' (재시도 후에도 실패)' if attempt == 2 else ''))
            return
        # the save-as dialog's last-used dir decides where NKEditor writes
        # (e.g. the user's Desktop) — move it, don't leave litter behind
        shutil.move(rtf, rtfdst)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    fix_rtf(rtfdst)
    print('review: rtf ->', rtfdst)
    if not os.path.exists(DOCCMP):
        issue('비교 문서 생성 불가 - DocCmp.exe 없음')
        return
    try:
        for cand in (cmpdoc, base + '.compare-new.docx'):
            with Stage('Word로 원본 docx와 비교 문서 만들기'):
                r = subprocess.run([DOCCMP, src, rtfdst, cand],
                                   capture_output=True, timeout=300)
            if r.returncode == 0 and os.path.exists(cand):
                for ln in _decode_output(r.stdout).splitlines():
                    # the compare doc exists: a WARN is a gap in it (e.g. a
                    # table not shown as a tracked deletion), not a failure
                    if ln.startswith('WARN '):
                        print('review: 비교 문서 경고 - ' + ln[5:],
                              file=sys.stderr)
                        REVIEW_NOTES.append('비교 문서 - ' + ln[5:])
                    elif ln.startswith('tables shown'):
                        print('review: ' + ln)
                print('review: compare doc ->', cand)
                return
        issue('비교 문서 생성 실패 - %s'
              % (_decode_output(r.stderr) or _decode_output(r.stdout)
                 ).strip()[-200:])
    except Exception as e:
        issue('비교 문서 생성 실패 - %r' % e)


def convert_one(src, dst, kind=None):
    """Convert a single .docx to .hlt. Returns 0 on success.
    kind: forced document kind (KIND_LABEL key), None = detect."""
    imgdir = tempfile.mkdtemp(prefix='d2h_img')
    CHAR_REPORT.clear()
    del WARNINGS[:]
    del REVIEW_ISSUES[:]
    del REVIEW_NOTES[:]
    for rep_ in (IMG_REPORT, FLAT_IMG_REPORT):
        rep_['ok'][:] = []
        rep_['fixed'][:] = []
    del ASIS_NOT_BW[:]
    img_lines = []              # image_report_lines() of the path used
    if kind:
        why = '--kind 지정'
    else:
        kind, why = detect_kind(src, docx_lines(src))
    print('document type: %s (판별 근거: %s)' % (KIND_LABEL[kind], why))
    if kind == 'amend':
        # 보정서 .hlt is made from the .fin; keep its 【청구항 N】 lines from
        # turning it into a spec and convert it as a plain document
        print('  보정서는 .fin에서 만드는 문서라 구조 변환 대상이 아닙니다 - '
              '일반 문서로 변환합니다.')
        kind = 'flat'
    eqs, tbls = render_objects(src, imgdir, tables=(kind != 'opinion'))
    if kind == 'opinion':
        blocks, imgfiles = parse_docx(src, imgdir, eqs, blocks=True)
        paras = None            # flat fallback re-parses on demand
    else:
        paras, imgfiles = parse_docx(src, imgdir, eqs, tbls)
        paras = split_glued_headings(paras, en=(kind == 'pct_en'))

    def write_report():
        if not (CHAR_REPORT or WARNINGS or img_lines):
            # clean state: remove reports from a previous run so stale
            # warnings can't be mistaken for problems in THIS output
            for stale in (os.path.splitext(dst)[0] + '.charset.txt',
                          os.path.splitext(dst)[0] + '.charset.docx'):
                try:
                    os.remove(stale)
                except OSError:
                    pass
            return
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
            for k, ln in enumerate(img_lines):
                f.write('%s%s\n' % ('   ' if k else '', ln))
            if img_lines:
                f.write('\n')
            f.write('chars not representable in KIPO editor — substituted or '
                    'dropped during conversion\n\n')
            for orig, repl, ctx in CHAR_REPORT:
                f.write('U+%04X %r -> %r   …%s\n' %
                        (ord(orig), orig, repl or '(삭제)', ctx))
        try:
            write_charset_docx(stem + '.charset.docx', WARNINGS, CHAR_REPORT,
                               img_lines)
        except Exception as e:
            print('review: charset docx write failed (%r)' % e,
                  file=sys.stderr)
        subs = sum(1 for _o, r, _c in CHAR_REPORT if r)
        drops = len(CHAR_REPORT) - subs
        print('review: %d warnings, %d substituted, %d dropped — %s(.docx)'
              % (len(WARNINGS), subs, drops, rep), file=sys.stderr)

    write_report()
    n_warn = len(WARNINGS)

    # Patent / PCT / 의견서 docs go through the real structured path: emit the
    # KIPO xml for that kind, let NKEditor build the element tree, then save
    # .hlt. Plain docs (and a failed structured build) use the flat TerConv
    # pipeline.
    if kind != 'flat' and HAVE_PIL:
        try:
            if convert_structured(src, dst,
                                  blocks if kind == 'opinion' else paras,
                                  imgfiles, kind):
                img_lines[:] = image_report_lines(IMG_REPORT)
                if not IMAGE_CONVERT and kind != 'opinion':
                    print('images: 자동 변환 꺼짐 - 그림을 원래 형식대로 넣음')
                    kept = list(dict.fromkeys(ASIS_NOT_BW))
                    if kind in ('pct', 'pct_en') and kept:
                        WARNINGS.append(
                            'PCT 출원인데 이미지 자동 변환이 꺼져 있어 흑백이 '
                            '아닌 그림 %d개가 그대로 들어감 (PCT는 흑백만 허용): '
                            '%s' % (len(kept), ', '.join(kept[:10])))
                if len(WARNINGS) > n_warn or img_lines:
                    # the XML build / PCT kind step can add warnings
                    for w in WARNINGS[n_warn:]:
                        print('WARNING: %s' % w, file=sys.stderr)
                    if img_lines:
                        print('images: ' + img_lines[0])
                    write_report()
                if IMAGE_CONVERT and not img_lines and IMG_REPORT['ok']:
                    print('images: 그림 %d개 모두 1비트 G4 300dpi TIFF 규격에 '
                          '맞음' % len(set(IMG_REPORT['ok'])))
                done = True
            else:
                done = False
        except Exception as e:
            print('structured path failed (%r); falling back to flat' % e,
                  file=sys.stderr)
            done = False
        if done:
            # outside the try: a review-file failure (an rtf still locked by
            # the editor) used to send a finished structured .hlt down the
            # flat path, which overwrote it
            try:
                make_review(src, dst)
            except Exception as e:
                print('review: 검토 파일 생성 실패 (%r)' % e, file=sys.stderr)
                REVIEW_ISSUES.append('검토 파일 생성 실패 - %r' % e)
            return 0
    if paras is None:
        paras, imgfiles = parse_docx(src, imgdir, eqs)

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
                           capture_output=True, text=True, errors='replace',
                           timeout=240, env=env)
        sys.stdout.write(r.stdout or '')
        sys.stderr.write(r.stderr or '')
        if r.returncode != 0:
            return r.returncode
    finally:
        os.unlink(jobfile)

    if not os.path.exists(dst):
        print('ERROR: output not created', file=sys.stderr)
        return 1
    print('OK ->', dst)
    # staging converted every word/media file; report only the ones the
    # body actually uses (unreferenced leftovers are not in the .hlt)
    used = set(_img_label(tgt) for _r, ims, _p in paras for (tgt, *_x) in ims)
    for k in ('ok', 'fixed'):
        FLAT_IMG_REPORT[k][:] = [e for e in FLAT_IMG_REPORT[k]
                                 if (e if k == 'ok' else
                                     e[len('그림 '):].split(': ', 1)[0])
                                 in used]
    img_lines[:] = image_report_lines(FLAT_IMG_REPORT)
    if img_lines:
        print('images: ' + img_lines[0])
        write_report()
    elif IMAGE_CONVERT and FLAT_IMG_REPORT['ok']:
        print('images: 그림 %d개 모두 1비트 G4 300dpi TIFF 규격에 맞음'
              % len(set(FLAT_IMG_REPORT['ok'])))
    make_review(src, dst)
    return 0


def main():
    # console is cp949 on stock Korean Windows — non-representable chars in
    # log messages (em dash, …) must not crash the run mid-conversion
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(errors='replace')
        except Exception:
            pass
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    kind = 'flat' if '--flat' in sys.argv[1:] else None
    for a in sys.argv[1:]:
        if a.startswith('--kind='):
            kind = a.split('=', 1)[1]
            if kind not in KIND_LABEL:
                print('--kind must be one of: %s' % ', '.join(KIND_LABEL))
                return 2
    # the Explorer right-click wrapper passes this: its console closes on its
    # own after a successful conversion, which hid every review failure
    pause_on_review = '--pause-on-review-error' in sys.argv[1:]
    global IMAGE_CONVERT
    IMAGE_CONVERT = '--no-image-convert' not in sys.argv[1:]
    if not args:
        print('usage: docx2hlt.py input.docx [output.hlt] [more.docx ...]')
        print('  output omitted -> <input>.hlt written next to each input')
        print('  --kind=spec|pct|pct_en|opinion|flat  override the detected kind '
              '(--flat = --kind=flat)')
        print('  --no-image-convert  이미지 자동 변환 끄기: 그림을 1비트 300dpi '
              'G4 TIFF로 바꾸지 않고 원래 형식대로 넣음 (국내 출원 컬러 그림)')
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
    review_failed = []
    review_noted = []
    stale_views = []
    for src, dst in jobs:
        if len(jobs) > 1:
            print('=== %s' % src)
        try:
            r = convert_one(src, os.path.abspath(dst), kind)
        except Exception as e:
            print('FAILED %s: %r' % (src, e), file=sys.stderr)
            r = 1
        if r:
            rc = r
        elif REVIEW_ISSUES:
            review_failed.append((src, REVIEW_ISSUES + REVIEW_NOTES))
        elif REVIEW_NOTES:
            review_noted.append((src, list(REVIEW_NOTES)))
        # an editor the user already had open on the output still shows the
        # old document; saving there would overwrite the new .hlt
        if not r:
            try:
                import nkwin
                pids = nkwin.editors_showing(dst)
            except Exception:
                pids = []
            if pids:
                stale_views.append((dst, pids))
    if review_failed:
        print()
        print('*** .hlt는 만들어졌지만 검토 파일(.rtf/.compare.docx)을 '
              '만들지 못했습니다 ***')
        for src, issues in review_failed:
            print('  %s' % os.path.basename(src))
            for msg in issues:
                print('    - %s' % msg)
    if review_noted:
        print()
        print('*** 검토 파일은 만들어졌지만 비교 문서에 확인할 점이 '
              '있습니다 ***')
        for src, notes in review_noted:
            print('  %s' % os.path.basename(src))
            for msg in notes:
                print('    - %s' % msg)
    if stale_views:
        print()
        print('*** 이미 열려 있던 NKEditor 창이 변환 전 .hlt를 보여 주고 '
              '있습니다 ***')
        for dst, pids in stale_views:
            print('  %s (NKEditor PID %s)'
                  % (os.path.basename(dst), ', '.join(map(str, pids))))
        print('  그 창에서 저장하면 새 .hlt를 덮어씁니다. 저장하지 말고 닫은 뒤 '
              '다시 여세요.')
    if review_failed or review_noted or stale_views:
        if pause_on_review:
            try:
                input('\nEnter 키를 누르면 창이 닫힙니다...')
            except EOFError:
                pass
    return rc


if __name__ == '__main__':
    sys.exit(main())
