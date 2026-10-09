# -*- coding: utf-8 -*-
"""
docx_edit_lib.py — 尽调报告docx手术公共函数库 v1.0
沉淀自2026-09-09某贸易客户戊尽调v3→v4实战（9批注/7表改列/3新表/图片替换/格式统一）。
所有函数遵循铁律：保留批注标记、按ECMA-376顺序插入pPr子元素、克隆模板保格式。

用法：
    sys.path.append(r'D:\\Claude\\tools')
    from docx_edit_lib import load_docx, save_docx, find_para, replace_in_para, ...
    parts, root, body = load_docx(src)
    ... 手术 ...
    save_docx(parts, root, dst)

⚠️ pPr子元素顺序（ECMA-376 sequence）：pStyle,keepNext,...,spacing,ind,jc,...,outlineLvl,...,rPr,sectPr
   jc 严禁append到rPr之后——否则Word报"文档已损坏"。本库 set_jc/insert helpers 已按序处理。
"""
import zipfile, copy
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

def q(t):
    return f'{{{W}}}{t}'


PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
             'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
             'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
             'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing',
             'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
             'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
_ORDER_IDX = {q(n): i for i, n in enumerate(PPR_ORDER)}


def load_docx(path):
    """返回 (parts dict, document root, body element)"""
    zin = zipfile.ZipFile(path)
    parts = {n: zin.read(n) for n in zin.namelist()}
    root = etree.fromstring(parts['word/document.xml'])
    return parts, root, root.find(q('body'))


def save_docx(parts, root, dst):
    parts['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)


def ptext(p):
    """段落全文（该段内所有w:t拼接）"""
    return ''.join(t.text or '' for t in p.iter(q('t'))).strip()


def find_para(body_or_root, snip, exact=False, last=False):
    """按文本找段落；exact=True精确匹配（自然跳过TOC条目，因其带\\t页码）"""
    ms = [p for p in body_or_root.iter(q('p'))
          if (ptext(p) == snip if exact else snip in ptext(p))]
    if not ms:
        raise RuntimeError(f'未找到段落: {snip[:40]}')
    return ms[-1] if last else ms[0]


def para_has_outline(p, lvl=None):
    ppr = p.find(q('pPr'))
    if ppr is None:
        return False
    o = ppr.find(q('outlineLvl'))
    return o is not None and (lvl is None or o.get(q('val')) == str(lvl))


def insert_ppr_child_ordered(ppr, el):
    """按ECMA-376顺序把pPr子元素插到正确位置（防文档损坏）"""
    idx = _ORDER_IDX.get(el.tag, 99)
    for child in ppr:
        cidx = _ORDER_IDX.get(child.tag, 99)
        if cidx > idx:
            child.addprevious(el)
            return
    ppr.append(el)


def set_jc(p, val):
    """设置段落对齐（顺序安全）"""
    ppr = p.find(q('pPr'))
    if ppr is None:
        ppr = etree.Element(q('pPr')); p.insert(0, ppr)
    jce = ppr.find(q('jc'))
    if jce is None:
        jce = etree.Element(q('jc'))
        insert_ppr_child_ordered(ppr, jce)
    jce.set(q('val'), val)


def add_keepnext(p):
    ppr = p.find(q('pPr'))
    if ppr is None:
        ppr = etree.Element(q('pPr')); p.insert(0, ppr)
    if ppr.find(q('keepNext')) is None:
        kn = etree.Element(q('keepNext'))
        ps = ppr.find(q('pStyle'))
        if ps is not None:
            ps.addnext(kn)  # keepNext紧跟pStyle，序位2
        else:
            insert_ppr_child_ordered(ppr, kn)


def add_cantsplit(tr):
    trpr = tr.find(q('trPr'))
    if trpr is None:
        trpr = etree.Element(q('trPr')); tr.insert(0, trpr)
    if trpr.find(q('cantSplit')) is None:
        trpr.insert(0, etree.Element(q('cantSplit')))  # CT_TrPr为choice无序


def replace_in_para(p, old, new, must=True):
    """跨run文本替换，保留格式与批注标记（commentRangeStart/End不受影响）"""
    ts = list(p.iter(q('t')))
    full = ''.join(t.text or '' for t in ts)
    idx = full.find(old)
    if idx < 0:
        if must:
            raise RuntimeError(f'未找到文本: {old[:50]}')
        return False
    spans, pos = [], 0
    for t in ts:
        ln = len(t.text or '')
        spans.append((t, pos, pos + ln)); pos += ln
    end = idx + len(old)
    done = False
    for t, s, e in spans:
        if e <= idx or s >= end:
            continue
        pre = (t.text or '')[0:max(0, idx - s):]
        suf = (t.text or '')[max(0, min(end - s, len(t.text or ''))):]
        if not done:
            t.text = pre + new + suf
            done = True
        else:
            t.text = pre + suf if (s <= end <= e) else ''
        if t.text:
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return True


def _clean_run_proto(proto):
    for el in list(proto):
        if el.tag != q('rPr'):
            proto.remove(el)
    return proto


def _first_text_run(p):
    for r in p.findall(q('r')):
        if r.find(q('t')) is not None and r.find(q('commentReference')) is None:
            return r
    return None


def append_text(p, add):
    """段落末尾追加文本（克隆末个文本run格式；跳过批注引用run）"""
    proto = _first_text_run(p)
    if proto is None:
        raise RuntimeError('段落无可克隆run')
    nr = _clean_run_proto(copy.deepcopy(proto))
    t = etree.SubElement(nr, q('t')); t.text = add
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    p.append(nr)


def set_para_text(p, new_text):
    """整段重写为单run。保留pPr与批注标记；含commentReference的run不删。"""
    proto = _first_text_run(p)
    if proto is None:
        raise RuntimeError('无原型run')
    proto = _clean_run_proto(copy.deepcopy(proto))
    for r in list(p.findall(q('r'))):
        if r.find(q('commentReference')) is None:
            p.remove(r)
    t = etree.SubElement(proto, q('t')); t.text = new_text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    p.append(proto)


def clone_para_with_text(template, text):
    """克隆模板段落（样式/pPr/字体），仅保留首个文本run并替换文本；剥离书签。"""
    np = copy.deepcopy(template)
    for tag in ('bookmarkStart', 'bookmarkEnd', 'hyperlink'):
        for el in np.findall(q(tag)):
            np.remove(el)
    proto = _first_text_run(np)
    if proto is None:
        raise RuntimeError('模板段无可克隆run')
    proto = _clean_run_proto(copy.deepcopy(proto))
    for r in list(np.findall(q('r'))):
        np.remove(r)
    t = etree.SubElement(proto, q('t')); t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    np.append(proto)
    return np


def delete_para(p):
    p.getparent().remove(p)


def insert_after(ref, els):
    cur = ref
    for e in els:
        cur.addnext(e); cur = e


def cell_set(tc, text, jc=None, bold=None):
    """表格单元格设文本（保留首段pPr与run字体格式）"""
    ps = tc.findall(q('p'))
    for extra in ps[1:]:
        tc.remove(extra)
    p = ps[0]
    proto = _first_text_run(p)
    if proto is None:
        proto = etree.Element(q('r'))
        rpr = etree.SubElement(proto, q('rPr'))
        f = etree.SubElement(rpr, q('rFonts'))
        f.set(q('ascii'), 'Times New Roman'); f.set(q('hAnsi'), 'Times New Roman')
    else:
        p.remove(proto)
        proto = _clean_run_proto(proto)
    for r in list(p.findall(q('r'))):
        p.remove(r)
    if bold is not None:
        rpr = proto.find(q('rPr'))
        if rpr is None:
            rpr = etree.Element(q('rPr')); proto.insert(0, rpr)
        if bold and rpr.find(q('b')) is None:
            rpr.insert(0, etree.Element(q('b')))
    t = etree.SubElement(proto, q('t')); t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    p.append(proto)
    if jc is not None:
        set_jc(p, jc)


def build_table(tblpr_proto, cell_proto, widths, rows_data, header_bold=True):
    """rows_data: [[(text, jc), ...], ...] 首行为表头（tblHeader+加粗）。三线表样式随tblPr_proto。"""
    tbl = etree.Element(q('tbl'))
    tbl.append(copy.deepcopy(tblpr_proto))
    grid = etree.SubElement(tbl, q('tblGrid'))
    for wd in widths:
        gc = etree.SubElement(grid, q('gridCol')); gc.set(q('w'), str(wd))
    for ri, row in enumerate(rows_data):
        tr = etree.SubElement(tbl, q('tr'))
        trpr = etree.SubElement(tr, q('trPr'))
        etree.SubElement(trpr, q('jc')).set(q('val'), 'center')
        etree.SubElement(trpr, q('cantSplit'))
        if ri == 0:
            etree.SubElement(trpr, q('tblHeader'))
        for ci, (txt, jc) in enumerate(row):
            tc = copy.deepcopy(cell_proto)
            tcw = tc.find(q('tcPr') + '/' + q('tcW'))
            if tcw is not None:
                tcw.set(q('w'), str(widths[ci]))
            cell_set(tc, txt, jc=jc, bold=(ri == 0 and header_bold))
            tr.append(tc)
    return tbl


def tbl_rows(tbl):
    return tbl.findall(q('tr'))


def row_tcs(tr):
    return tr.findall(q('tc'))


def cell_text(tc):
    return ''.join(tc.itertext()).strip()


def grid_set(tbl, widths):
    """重设tblGrid列宽（新增/删列后调用），并同步每行tcW"""
    grid = tbl.find(q('tblGrid'))
    for gc in grid.findall(q('gridCol')):
        grid.remove(gc)
    for wd in widths:
        gc = etree.SubElement(grid, q('gridCol')); gc.set(q('w'), str(wd))
    for tr in tbl.findall(q('tr')):
        for tc, wd in zip(tr.findall(q('tc')), widths):
            tcw = tc.find(q('tcPr') + '/' + q('tcW'))
            if tcw is not None:
                tcw.set(q('w'), str(wd))


def fix_ppr_order(root):
    """修复全文档pPr子元素顺序（jc排在rPr后等会致Word报损坏）。返回修复数。"""
    fixed = 0
    for ppr in root.iter(q('pPr')):
        children = list(ppr)
        known = [(i, _ORDER_IDX.get(c.tag, -1)) for i, c in enumerate(children)]
        known_pairs = [k for k in known if k[1] >= 0]
        if [k[0] for k in known_pairs] != [k[0] for k in sorted(known_pairs, key=lambda x: x[1])]:
            new_children = sorted(children, key=lambda c: _ORDER_IDX.get(c.tag, 99))
            for c in children:
                ppr.remove(c)
            for c in new_children:
                ppr.append(c)
            fixed += 1
    return fixed


def justify_body_paras(root, min_chars=20):
    """正文段落两端对齐（跳过表格内/标题/居中/右对齐/单位行/备注行）。返回修复数。"""
    body = root.find(q('body'))
    fixed = 0
    for p in body.iter(q('p')):
        if p.getparent().tag == q('tbl'):
            continue
        ppr = p.find(q('pPr'))
        if ppr is None or ppr.find(q('outlineLvl')) is not None:
            continue
        jce = ppr.find(q('jc'))
        cur = jce.get(q('val')) if jce is not None else None
        txt = ptext(p)
        if len(txt) < min_chars or txt.startswith(('单位：', '备注：')) or cur in ('center', 'right'):
            continue
        if cur is None or cur == 'left':
            set_jc(p, 'both')
            fixed += 1
    return fixed


# ==========================================================================
# v1.1 追加（2026-09-11 某锂盐客户庚v2复盘沉淀）
#   ① clone_row_after / delete_row —— 表格行增删（保留tcPr格式）
#   ② set_table_autofit            —— 表格按窗口自动调整（终审人标准）
#   ③ ensure_circled_hint          —— 带圈序号字体归属守卫
# ==========================================================================
import re as _re

_CIRCLED = _re.compile(r'[①-⑳㉑-㉟㊱-㊿]')

# tblPr 子元素顺序（ECMA-376）：tblLayout 必须排在 tblCellMar / tblLook 之前，
# 否则 Word 保存时静默丢弃（设置"成功"实为被吞）
_TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
                'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders',
                'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
_TBLPR_IDX = {q(n): i for i, n in enumerate(_TBLPR_ORDER)}


def _insert_tblpr_ordered(tblpr, el):
    idx = _TBLPR_IDX.get(el.tag, 99)
    for child in tblpr:
        if _TBLPR_IDX.get(child.tag, 99) > idx:
            child.addprevious(el)
            return
    tblpr.append(el)


def set_table_autofit(tbl, pct=5000):
    """表格按窗口自动调整：tblW=pct/5000 + tblLayout=autofit（顺序安全）。返回是否新设tblLayout。"""
    tblpr = tbl.find(q('tblPr'))
    if tblpr is None:
        tblpr = etree.Element(q('tblPr'))
        tbl.insert(0, tblpr)
    w = tblpr.find(q('tblW'))
    if w is None:
        w = etree.Element(q('tblW'))
        _insert_tblpr_ordered(tblpr, w)
    w.set(q('type'), 'pct')
    w.set(q('w'), str(pct))
    lay = tblpr.find(q('tblLayout'))
    fresh = lay is None
    if fresh:
        lay = etree.Element(q('tblLayout'))
        _insert_tblpr_ordered(tblpr, lay)
    lay.set(q('type'), 'autofit')
    return fresh


def clone_row_after(tbl, ref_idx, values, tpl_idx=None):
    """克隆模板行（默认 ref_idx 行）插到 ref_idx 之后，逐格写入 values。
    保留 tcPr/pPr/rPr 格式；清理多余段落。返回新 tr。"""
    trs = tbl.findall(q('tr'))
    tpl = trs[tpl_idx if tpl_idx is not None else ref_idx]
    nt = copy.deepcopy(tpl)
    trs[ref_idx].addnext(nt)
    for tc, v in zip(nt.findall(q('tc')), values):
        cell_set(tc, v)
    return nt


def delete_row(tbl, idx, guard_comments=True):
    """删除表格第 idx 行。guard_comments=True 时若该行承载批注/修订锚点则拒绝并返回 False。"""
    trs = tbl.findall(q('tr'))
    if idx < 0 or idx >= len(trs):
        raise IndexError('行号越界: %d（共%d行）' % (idx, len(trs)))
    tr = trs[idx]
    if guard_comments:
        marks = [(etree.QName(e).localname, e.get(q('id')))
                 for e in tr.iter()
                 if etree.QName(e).localname in ('commentRangeStart', 'commentRangeEnd',
                                                 'commentReference', 'ins', 'del')]
        if marks:
            raise RuntimeError('第%d行承载批注/修订锚点 %s —— 删行将连带删除。'
                               '确认已解决后传 guard_comments=False' % (idx, marks))
    tr.getparent().remove(tr)
    return True


def ensure_circled_hint(p_or_r):
    """带圈序号（①-⑳）字体归属守卫：所在 run 的 rFonts 必须带 w:hint='eastAsia'，
    否则被西文字体（TNR）接管渲染，序号显示偏大、与全文不一致。
    传 <w:p> 或 <w:r>；返回修正的 run 数。"""
    runs = [p_or_r] if p_or_r.tag == q('r') else list(p_or_r.findall(q('r')))
    fixed = 0
    for r in runs:
        if not _CIRCLED.search(''.join(t.text or '' for t in r.iter(q('t')))):
            continue
        rpr = r.find(q('rPr'))
        if rpr is None:
            rpr = etree.Element(q('rPr'))
            r.insert(0, rpr)
        rf = rpr.find(q('rFonts'))
        if rf is None:
            rf = etree.Element(q('rFonts'))
            rpr.insert(0, rf)
        if rf.get(q('hint')) != 'eastAsia':
            rf.set(q('hint'), 'eastAsia')
            fixed += 1
    return fixed


def set_para_text_guarded(p, new_text):
    """set_para_text + 带圈序号 hint 守卫（推荐入口）。"""
    set_para_text(p, new_text)
    return ensure_circled_hint(p)


# ==========================================================================
# v1.2 追加（2026-09-21 某锂盐客户庚第5版二轮复盘沉淀）
#   ① guard_comment_anchors      —— 批注锚点复位守卫（整段转删/重写【前】必须调用，
#                                  否则接受修订时区间被挤空、Word 静默丢批注，v6.5铁律③）
#   ② bold_text_tracked          —— 修订模式加粗原语（每轮重扫run快照+已粗字符置NUL掩码；
#                                  旧版"收集span倒序处理+同一份快照"在同run多命中时重复切分
#                                  致正文文字重复错乱，20260921实测返工一次；含最大命中保护）
#   ③ tracked_delete_para_runs   —— 整段转删（前置①守卫 + 直接run转del包装）
# ==========================================================================
import itertools as _it
import time as _time


def _mk_tag(tag, author, date, nid):
    el = etree.Element(q(tag))
    el.set(q('id'), str(nid))
    el.set(q('author'), author)
    el.set(q('date'), date)
    return el


def guard_comment_anchors(p):
    """批注锚点复位守卫：把段内 commentRangeStart 移到段首（pPr 后）、
    commentRangeEnd 移到段末、commentReference 所在 run 移到 RangeEnd 之后。
    任何"整段转删/整段重写/逐run转del"手术【前】调用；纯元素移动、无 id 需求。
    返回移动的锚点元素数（0=段内无锚点）。"""
    moved = 0
    ppr = p.find(q('pPr'))
    for crs in list(p.findall(q('commentRangeStart'))):
        p.remove(crs)
        (ppr.addnext(crs) if ppr is not None else p.insert(0, crs))
        moved += 1
    for cre in list(p.findall(q('commentRangeEnd'))):
        p.remove(cre)
        p.append(cre)
        moved += 1
        # 同段的 reference run 跟到 cre 之后（cre 可能有多个，逐个处理）
        cid = cre.get(q('id'))
        for r in list(p.findall(q('r'))):
            ref = r.find(q('commentReference'))
            if ref is not None and ref.get(q('id')) == cid:
                p.remove(r)
                cre.addnext(r)
                moved += 1
    return moved


def _run_is_bold(r):
    rpr = r.find(q('rPr'))
    return rpr is not None and rpr.find(q('b')) is not None


def _set_run_bold_tracked(r, author, date, nid):
    """run加粗并挂 rPrChange（原格式留档，可在Word中整批拒绝）。已粗返回 False。"""
    rpr = r.find(q('rPr'))
    if rpr is None:
        rpr = etree.Element(q('rPr'))
        r.insert(0, rpr)
    if rpr.find(q('b')) is not None:
        return False
    old_copy = copy.deepcopy(rpr)
    b = etree.Element(q('b'))
    rf = rpr.find(q('rFonts'))
    if rf is not None:
        rf.addnext(b)
    else:
        rpr.insert(0, b)
    chg = _mk_tag('rPrChange', author, date, nid)
    chg.append(old_copy)
    rpr.append(chg)  # rPrChange 按规范排 rPr 末尾
    return True


def _set_run_text(r, s):
    ts = r.findall(q('t'))
    for t in ts[1:]:
        r.remove(t)
    if not ts:
        t = etree.Element(q('t'))
        r.append(t)
        ts = [t]
    t = ts[0]
    t.text = s
    if s != s.strip():
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    else:
        t.attrib.pop('{http://www.w3.org/XML/1998/namespace}space', None)


def bold_text_tracked(p, word, author='Claude', date=None, next_id=None, max_hits=50):
    """修订模式把段内 word 的所有出现处加粗（切run保格式+格式修订留档）。

    ⚡算法铁律（20260921血泪）：每轮重扫 run 快照 + 已粗字符置 NUL 不参与匹配——
    禁止"先收集全部span再倒序处理同一份快照"（同run多命中/子串嵌套如
    "长期借款"⊂"一年内到期长期借款"会重复切分，正文文字重复错乱）。

    next_id: 返回唯一int的 callable（默认自时间戳派生自增序列）；max_hits: 防异常数据死循环。
    返回加粗span数。文本含圈码时自动补 hint=eastAsia。"""
    if date is None:
        date = _time.strftime('%Y-%m-%dT%H:%M:%SZ', _time.gmtime())
    if next_id is None:
        seq = _it.count(int(_time.time()) % 1000000 * 100)
        next_id = lambda: next(seq)
    n = 0
    while n < max_hits:
        runs = [r for r in p.findall(q('r')) if r.findall(q('t'))]
        texts = [''.join(t.text or '' for t in r.findall(q('t'))) for r in runs]
        masks = ['\x00' * len(t) if _run_is_bold(r) else t for t, r in zip(texts, runs)]
        full = ''.join(masks)
        idx = full.find(word)
        if idx < 0:
            break
        end = idx + len(word)
        pos = 0
        for ri, rt in enumerate(texts):
            s, e = pos, pos + len(rt)
            if e > idx and s < end:
                a, b = max(idx, s) - s, min(end, e) - s
                r = runs[ri]
                pre, mid, suf = rt[:a], rt[a:b], rt[b:]
                if suf:
                    r_suf = copy.deepcopy(r)
                    _set_run_text(r_suf, suf)
                    r.addnext(r_suf)
                _set_run_text(r, mid)
                _set_run_bold_tracked(r, author, date, next_id())
                if pre:
                    r_pre = copy.deepcopy(r)
                    rpr = r_pre.find(q('rPr'))
                    if rpr is not None:
                        for tg in ('b', 'bCs', 'rPrChange'):
                            el = rpr.find(q(tg))
                            if el is not None:
                                rpr.remove(el)
                    _set_run_text(r_pre, pre)
                    r.addprevious(r_pre)
            pos = e
        n += 1
    else:
        raise RuntimeError(f'bold_text_tracked 命中达 max_hits={max_hits} 上限，疑似异常数据/死循环，已中断: {word[:30]}')
    ensure_circled_hint(p)
    return n


def tracked_delete_para_runs(p, author='Claude', date=None, next_id=None, skip_ref_runs=True):
    """整段转删：先跑批注锚点复位守卫，再把段落直接子 run 转 w:del（w:t→w:delText）。
    skip_ref_runs=True 时跳过批注引用 run（无 w:t 本就不在转换范围）。返回转删 run 数。"""
    if date is None:
        date = _time.strftime('%Y-%m-%dT%H:%M:%SZ', _time.gmtime())
    if next_id is None:
        seq = _it.count(int(_time.time()) % 1000000 * 100 + 1)
        next_id = lambda: next(seq)
    guard_comment_anchors(p)
    n = 0
    for r in list(p.findall(q('r'))):
        if not (r.findall(q('t')) or r.findall(q('br')) or r.findall(q('tab'))):
            continue
        del_el = _mk_tag('del', author, date, next_id())
        r.addprevious(del_el)
        del_el.append(r)
        for t in r.findall(q('t')):
            t.tag = q('delText')
            if (t.text or '') != (t.text or '').strip():
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        n += 1
    return n


# ==========================================================================
# v1.3 追加（2026-09-28 客户丙四主体ChatGPT意见轮沉淀）
#   ① tracked_replace        —— 段内跨run修订替换（old转del+new以ins插入；唯一命中
#                               断言+改后自验；20260928客户丙轮9处手术实战验证）
#   ② tracked_insert_para_after —— 修订式整段插入（克隆模板格式+段内run包ins）
#   ③ tracked_insert_row_after   —— 修订式插行（trPr/ins+cellIns；默认补cantSplit——
#                               非build_table路径的表行无cantSplit，克隆后会被劈两半
#                               跨页，20260928表24实测渲染抓到）
# ==========================================================================


def _default_next_id():
    import itertools as _it, time as _t
    seq = _it.count(int(_t.time()) % 1000000 * 100)
    return lambda: next(seq)


def tracked_replace(p, old, new, author='运营风控组', date=None, next_id=None, label=''):
    """段内跨run修订替换：old→w:del(delText)，new→w:ins（紧跟首个del）。
    old须段内唯一（命中≠1直接抛错）；改后自验可见文本含new不含old。
    返回del覆盖的run数。跨run文本、修订标记、run格式(克隆)均保留。"""
    if date is None:
        import time as _t
        date = _t.strftime('%Y-%m-%dT%H:%M:%SZ', _t.gmtime())
    if next_id is None:
        next_id = _default_next_id()
    runs = [r for r in p.findall(q('r')) if r.findall(q('t'))]
    texts = [''.join(t.text or '' for t in r.findall(q('t'))) for r in runs]
    full = ''.join(texts)
    cnt = full.count(old)
    if cnt != 1:
        raise RuntimeError(f'[tracked_replace{"/" + label if label else ""}] 命中{cnt}次(须1): {old[:40]}')
    idx = full.find(old); end = idx + len(old)
    spans, pos = [], 0
    for r, t in zip(runs, texts):
        spans.append((r, pos, pos + len(t))); pos += len(t)
    overlap = [(r, s, e, t) for (r, s, e), t in zip(spans, texts) if e > idx and s < end]
    first_del = None
    for r, s, e, t in reversed(overlap):   # 倒序：后面的先处理，前面run定位不受影响
        a, b = max(idx, s) - s, min(end, e) - s
        pre, mid, suf = t[:a], t[a:b], t[b:]
        if suf:
            rs = copy.deepcopy(r); _set_run_text(rs, suf); r.addnext(rs)
        _set_run_text(r, mid)
        del_el = _mk_tag('del', author, date, next_id())
        r.getparent().replace(r, del_el); del_el.append(r)
        for t_el in r.findall(q('t')):
            t_el.tag = q('delText')
            t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        if pre:
            rp = copy.deepcopy(r)
            for t_el in rp.findall(q('delText')):
                t_el.tag = q('t')
            _set_run_text(rp, pre)
            del_el.addprevious(rp)
        first_del = del_el
    ins_el = _mk_tag('ins', author, date, next_id())
    nr = copy.deepcopy(overlap[0][0])       # 克隆首个重叠run（rPr保留）
    for t_el in nr.findall(q('delText')):
        t_el.tag = q('t')
    _set_run_text(nr, new)
    ins_el.append(nr)
    first_del.addnext(ins_el)
    vis = ptext(p)
    assert new in vis and old not in vis, f'[tracked_replace{"/" + label if label else ""}] 改后自验失败'
    return len(overlap)


def tracked_insert_para_after(ref_el, text, tpl_para, author='运营风控组', date=None, next_id=None):
    """在ref_el（段落或表格元素）后插入整段：克隆tpl_para格式，文本置于段内，
    段内文本run包w:ins（修订插入）。返回新段落元素。
    ⚡模板段铁律（20260928客户庚v6轮）：tpl_para 必须是**未被本轮 tracked_replace/转删
    触碰**的段——被改写过的段其 run 已包进 del/ins、不再是段落直接子 run，
    clone_para_with_text 会报"模板段无可克隆run"。用同格式的干净段（如其他备注段）。
    ⚡模板段铁律二（20260930客户庚v11轮）：tpl_para 还必须**不携带批注锚点**——
    clone_para_with_text 深拷贝会把段内/跨段批注的 commentRangeEnd（甚至Start/ref）
    一并克隆进新段，产生孤儿锚点（crs/cre不配对），Word接受修订时可能静默丢批注。
    selfcheck[2]锚点配对检查会抓到（crs=2/cre=3类报错）。稳妥做法：插段后立即
    清理新段内的全部 commentRangeStart/commentRangeEnd/commentReference 元素。"""
    if date is None:
        import time as _t
        date = _t.strftime('%Y-%m-%dT%H:%M:%SZ', _t.gmtime())
    if next_id is None:
        next_id = _default_next_id()
    np_ = clone_para_with_text(tpl_para, text)
    done = 0
    for r in list(np_.findall(q('r'))):
        if r.findall(q('t')):
            ins_el = _mk_tag('ins', author, date, next_id())
            r.getparent().replace(r, ins_el); ins_el.append(r); done += 1
    if done < 1:
        raise RuntimeError('[tracked_insert_para_after] 段内无ins run（模板段无文本run?）')
    ref_el.addnext(np_)
    return np_


def tracked_insert_row_after(tbl, ref_row, values, author='运营风控组', date=None, next_id=None,
                             cant_split=True):
    """修订式插行：克隆ref_row于其后，trPr挂ins、各tc挂cellIns，values填文本。
    ⚡cant_split=True（默认）给新行补<w:cantSplit/>——非build_table生成的表行
    trPr常只有jc，克隆行不补会在修订显示态被劈两半跨页（20260928表24实测）。
    trPr子元素序：cantSplit须在jc/ins之前，故insert(0)。返回新tr元素。
    ⚡克隆源铁律（20260930客户庚表10实测）：ref_row必须是本轮未触碰的干净行——
    若克隆刚被tracked_replace改过的行，其del/ins标记一并被克隆，cell_set清不掉
    w:del/w:ins内的残留，接受修订后新行单元格会叠着旧标签文字。与
    tracked_insert_para_after的"模板段必须未触碰"同族（段落级→行级）。
    需要插在A行后而A行已脏时：用干净行B作ref_row克隆，再 B.addnext的返回值
    换成 A.addnext(new_tr) 移位（lxml addnext对已有元素=移动）。"""
    if date is None:
        import time as _t
        date = _t.strftime('%Y-%m-%dT%H:%M:%SZ', _t.gmtime())
    if next_id is None:
        next_id = _default_next_id()
    new_tr = copy.deepcopy(ref_row)
    trpr = new_tr.find(q('trPr'))
    if trpr is None:
        trpr = etree.Element(q('trPr')); new_tr.insert(0, trpr)
    if cant_split and trpr.find(q('cantSplit')) is None:
        trpr.insert(0, etree.Element(q('cantSplit')))
    trpr.append(_mk_tag('ins', author, date, next_id()))
    tcs = row_tcs(new_tr)
    if len(tcs) != len(values):
        raise RuntimeError(f'[tracked_insert_row_after] values {len(values)}格 ≠ 行 {len(tcs)}格')
    for tc, val in zip(tcs, values):
        tcpr = tc.find(q('tcPr'))
        if tcpr is None:
            tcpr = etree.Element(q('tcPr')); tc.insert(0, tcpr)
        tcpr.append(_mk_tag('cellIns', author, date, next_id()))
        cell_set(tc, val)
    ref_row.addnext(new_tr)
    return new_tr
