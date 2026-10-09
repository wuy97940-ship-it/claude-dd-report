# -*- coding: utf-8 -*-
"""
dd_comment_closure.py — 尽调报告批注闭环器 v1.0（2026-09-11 某锂盐客户庚v2复盘沉淀）
================================================================================
批注驱动修订场景的"读批注 → 改正文 → 清批注"标准工具。

【为什么需要】
  批注锚点（commentRangeStart/End）可能落在**表格单元格内**。一旦删除该行，
  锚点随行消失，Word 里批注数会莫名减少（实测 4→1），看起来像"批注丢了"。
  本工具在删除前显式告警，并把"已解决批注"清理干净。

【清理范围】（漏任何一个 Word 都报"文档已损坏"）
  ① document.xml 中 commentRangeStart / commentRangeEnd / commentReference 三类元素
  ② word/comments.xml、commentsExtended.xml、commentsIds.xml、commentsExtensible.xml、people.xml
  ③ word/_rels/document.xml.rels 中指向上述部件的 5 条 Relationship

用法：
  python dd_comment_closure.py X.docx --dump          # 读：批注全文 + 精确锚定位置
  python dd_comment_closure.py X.docx --close         # 清：清除全部批注（自动备份）
  python dd_comment_closure.py X.docx --close --keep-anchors   # 只清 comments 部件，留锚点
  python dd_comment_closure.py X.docx --row-guard T3 R16       # 删行前：查该表格行是否承载批注
退出码：0=成功 2=失败
"""
import os
import sys
import shutil
import zipfile
import argparse
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
DROP_PARTS = ('word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml',
              'word/commentsExtensible.xml', 'word/people.xml')
ANCHOR_TAGS = ('commentRangeStart', 'commentRangeEnd', 'commentReference')


def _local(e):
    return etree.QName(e).localname


def _paren(el, stop=('tbl', 'body')):
    p = el
    while p is not None and _local(p) not in stop:
        p = p.getparent()
    return p


def _ptext(p):
    return ''.join(t.text or '' for t in p.iter('{%s}t' % W)).strip()


def _anchor_loc(el, body):
    """返回锚点位置描述：正文段序号 / 表格 T#R#C#"""
    p = _paren(el, ('p', 'body'))
    tbl = _paren(el, ('tbl', 'body'))
    if tbl is not None and _local(tbl) == 'tbl':
        ti = list(body.iter('{%s}tbl' % W)).index(tbl)
        for ri, tr in enumerate(tbl.findall('{%s}tr' % W)):
            for ci, tc in enumerate(tr.findall('{%s}tc' % W)):
                if tc is el or el in tc.iter():
                    return '表格 T%d R%d C%d | %s' % (ti, ri, ci, (_ptext(p) if p is not None else ''))
        return '表格 T%d（行定位失败）' % ti
    # 正文段
    paras = [x for x in body.iter('{%s}p' % W) if _paren(x, ('tbl', 'body')) is not None
             and _local(_paren(x, ('tbl', 'body'))) == 'body']
    for i, x in enumerate(paras):
        if x is p:
            return '正文段 #%d | %s' % (i, _ptext(p))
    return '正文 | %s' % (_ptext(p) if p is not None else '')


def dump(path):
    z = zipfile.ZipFile(path)
    if 'word/comments.xml' not in z.namelist():
        print('无 comments.xml —— 本文件无批注')
        return 0
    cx = etree.fromstring(z.read('word/comments.xml'))
    dx = etree.fromstring(z.read('word/document.xml'))
    body = dx.find('{%s}body' % W)

    meta = {}
    for c in cx.findall('{%s}comment' % W):
        cid = c.get('{%s}id' % W)
        meta[cid] = dict(author=c.get('{%s}author' % W), date=c.get('{%s}date' % W),
                         text='\n'.join(''.join(t.text or '' for t in p.iter('{%s}t' % W))
                                        for p in c.findall('{%s}p' % W)))

    locs = {}
    for el in body.iter():
        if _local(el) in ANCHOR_TAGS and _local(el) != 'commentReference':
            cid = el.get('{%s}id' % W)
            L = _anchor_loc(el, body)
            if L not in locs.setdefault(cid, []):      # Start/End 同段 → 去重
                locs[cid].append(L)

    print('=' * 78)
    print('批注清单：共 %d 条' % len(meta))
    print('=' * 78)
    for cid in sorted(meta, key=lambda x: int(x)):
        m = meta[cid]
        print('[id=%s] %s  %s' % (cid, m['author'], m['date']))
        print('  内容: %s' % m['text'])
        for L in locs.get(cid, ['（未找到锚点：可能所在行/段已被删除）']):
            tag = '  ⚠表格内' if L.startswith('表格') else '   正文  '
            print('%s%s' % (tag, L))
        print('-' * 74)
    in_tbl = [c for c in locs if any(L.startswith('表格') for L in locs[c])]
    if in_tbl:
        print()
        print('⚠ 注意：批注 %s 锚定在表格单元格内。删除其所在行会连带删除锚点。'
              % '、'.join(sorted(in_tbl, key=lambda x: int(x))))
        print('  若这些批注已解决 → 属正常闭环；若未解决 → 请先处理再删行。')
    return 0


def close(path, keep_anchors=False, backup=True):
    if backup:
        bak = path.replace('.docx', '_批注清理前备份.docx')
        shutil.copy2(path, bak)
        print('备份 -> %s' % os.path.basename(bak))
    tmp = path + '.tmp'
    zin = zipfile.ZipFile(path)
    n_anchor = 0
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            name = item.filename
            data = zin.read(name)
            if name in DROP_PARTS:
                print('  移除部件: %s' % name)
                continue
            if name == 'word/document.xml' and not keep_anchors:
                root = etree.fromstring(data)
                for el in list(root.iter()):
                    if _local(el) in ANCHOR_TAGS:
                        el.getparent().remove(el)
                        n_anchor += 1
                print('  清除锚点元素: %d 个' % n_anchor)
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            if name == 'word/_rels/document.xml.rels':
                root = etree.fromstring(data)
                R = 'http://schemas.openxmlformats.org/package/2006/relationships'
                for e in list(root.findall('{%s}Relationship' % R)):
                    tgt = e.get('Target') or ''
                    if 'comments' in tgt or tgt.endswith('people.xml'):
                        print('  移除关系: %s' % tgt)
                        root.remove(e)
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            zout.writestr(item, data)
    zin.close()
    os.replace(tmp, path)
    z = zipfile.ZipFile(path)
    left = [n for n in z.namelist() if 'comment' in n.lower()]
    print('完成。剩余 comments 部件: %s' % (left or '无'))
    return 0


def row_guard(path, t_idx, r_idx):
    """删行前检查：该表格行是否承载批注锚点/修订"""
    z = zipfile.ZipFile(path)
    dx = etree.fromstring(z.read('word/document.xml'))
    body = dx.find('{%s}body' % W)
    tbls = list(body.iter('{%s}tbl' % W))
    if t_idx >= len(tbls):
        print('!! 表格 T%d 不存在（共 %d 个）' % (t_idx, len(tbls)))
        return 2
    trs = tbls[t_idx].findall('{%s}tr' % W)
    if r_idx >= len(trs):
        print('!! T%d 只有 %d 行' % (t_idx, len(trs)))
        return 2
    tr = trs[r_idx]
    hits = [(_local(e), e.get('{%s}id' % W)) for e in tr.iter()
            if _local(e) in ANCHOR_TAGS or _local(e) in ('ins', 'del')]
    rowtxt = ' | '.join(''.join(t.text or '' for t in tc.iter('{%s}t' % W))
                        for tc in tr.findall('{%s}tc' % W))
    print('T%d R%d: %s' % (t_idx, r_idx, rowtxt[:120]))
    if hits:
        print('  ⚠ 该行承载 %d 个批注/修订元素：%s' % (len(hits), hits))
        print('  删此行将一并删除它们 —— 确认这些批注/修订已解决后再删。')
        return 1
    print('  ✅ 该行无批注/修订锚点，可安全删除')
    return 0


def main():
    ap = argparse.ArgumentParser(description='尽调报告批注闭环器')
    ap.add_argument('docx')
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument('--dump', action='store_true', help='读取：批注全文 + 精确锚定位置')
    g.add_argument('--close', action='store_true', help='清除全部批注')
    g.add_argument('--row-guard', nargs=2, type=int, metavar=('T', 'R'), help='删行前检查该行是否承载批注')
    ap.add_argument('--keep-anchors', action='store_true', help='--close 时保留锚点，仅删 comments 部件')
    ap.add_argument('--no-backup', action='store_true')
    a = ap.parse_args()
    if not os.path.exists(a.docx):
        print('!! 文件不存在: %s' % a.docx)
        sys.exit(2)
    if a.dump:
        sys.exit(dump(a.docx))
    if a.close:
        sys.exit(close(a.docx, a.keep_anchors, not a.no_backup))
    if a.row_guard:
        sys.exit(row_guard(a.docx, a.row_guard[0], a.row_guard[1]))


if __name__ == '__main__':
    main()
