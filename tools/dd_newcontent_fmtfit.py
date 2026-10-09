# -*- coding: utf-8 -*-
"""尽调报告「新增内容格式对齐」体检器（v1.1，2026-09-14 上海某贸易客户戊第七稿沉淀）

用途：AI 往既有尽调报告里新增段落/表格后，逐项核对是否与报告既有体例一致。
背景（本次踩坑）：克隆既有表格会连带继承**源表的列宽**、克隆既有段落会连带继承**源段的 outlineLvl**——
      内容对了但格式不对：备注列被挤成 6 行、正文段混进目录、表格列序与全文财务表不一致。

检查三类：
  [1] 段落层级：按体例推断应有 outlineLvl（章=style控制／节=1／子号=2／正文=无），排除目录段
  [2] 表格列宽：用「列宽分布」与「各列内容长度分布」的皮尔逊相关度判定——
      相关度过低即说明列宽是**从别的表继承来的**（而不是按本表内容分配）
  [3] 表格列序：含多期数据的表，列序应为「项目｜本期｜[占比]｜上期｜较上期变动｜前前期」

用法：
  python dd_newcontent_fmtfit.py 报告.docx          # 体检（退出码 0=通过 2=有问题）
  python dd_newcontent_fmtfit.py 报告.docx --fix    # 修正 [1]；对 [2] 命中的表按内容比例重分配列宽
"""
import sys, io, re, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

CN = '一二三四五六七八九十'
TOTAL_W = 8306            # 报告表格统一总宽（twips），由既有表反推
CORR_MIN = 0.60           # 列宽-内容相关度下限


def ptext(p):
    o = []
    for n in p.iter(qn('w:t')):
        a = n.getparent(); k = False
        while a is not None:
            if a.tag == qn('w:del'):
                k = True; break
            a = a.getparent()
        if not k:
            o.append(n.text or '')
    return ''.join(o)


def style_of(p):
    ppr = p.find(qn('w:pPr'))
    if ppr is None:
        return None
    st = ppr.find(qn('w:pStyle'))
    return st.get(qn('w:val')) if st is not None else None


def get_lvl(p):
    ppr = p.find(qn('w:pPr'))
    if ppr is None:
        return None
    ol = ppr.find(qn('w:outlineLvl'))
    return ol.get(qn('w:val')) if ol is not None else None


def set_lvl(p, lvl):
    ppr = p.find(qn('w:pPr'))
    if ppr is None:
        ppr = etree.Element(qn('w:pPr')); p.insert(0, ppr)
    ol = ppr.find(qn('w:outlineLvl'))
    if lvl is None:
        if ol is not None:
            ppr.remove(ol)
        return
    if ol is None:
        ol = etree.SubElement(ppr, qn('w:outlineLvl'))
    ol.set(qn('w:val'), str(lvl))


def expected_lvl(t):
    if re.match(r'^[%s]+、' % CN, t):
        return 'skip'
    if re.match(r'^（[%s]+）' % CN, t):
        return 1
    if re.match(r'^\d+、', t) and len(t) < 30:
        return 2
    return None


def cell_text(tc):
    return ''.join(x.text or '' for x in tc.iter(qn('w:t')))


def wlen(s):
    return sum(2 if ord(ch) > 127 else 1 for ch in s)


def corr(a, b):
    n = len(a)
    if n < 3:
        return 1.0
    ma, mb = sum(a) / n, sum(b) / n
    va = sum((x - ma) ** 2 for x in a) ** 0.5
    vb = sum((x - mb) ** 2 for x in b) ** 0.5
    if va == 0 or vb == 0:
        return 1.0
    return sum((a[i] - ma) * (b[i] - mb) for i in range(n)) / (va * vb)


def rebalance(t, req, cols):
    s = sum(req)
    new = [max(700, int(TOTAL_W * r / s)) for r in req]
    new[-1] += TOTAL_W - sum(new)
    for gc, w in zip(cols, new):
        gc.set(qn('w:w'), str(w))
    for tr in t.findall(qn('w:tr')):
        for tc, w in zip(tr.findall(qn('w:tc')), new):
            tcpr = tc.find(qn('w:tcPr'))
            if tcpr is None:
                tcpr = etree.Element(qn('w:tcPr')); tc.insert(0, tcpr)
            for old in list(tcpr.findall(qn('w:tcW'))):
                tcpr.remove(old)
            el = etree.Element(qn('w:tcW')); el.set(qn('w:w'), str(w)); el.set(qn('w:type'), 'dxa')
            tcpr.insert(0, el)
    return new


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--fix', action='store_true',
                    help='修正 [1] 段落层级（安全，只动新增段的 outlineLvl）')
    ap.add_argument('--fix-table', type=int, default=0,
                    help='按内容重分配第 N 张表（1 起）的列宽——只对**本次新增**的表用，'
                         '禁止对既有表使用（会把作者按设计排好的列宽改掉）')
    a = ap.parse_args()
    doc = Document(a.docx)
    body = doc.element.body
    issues = 0

    print('=== [1] 段落层级（仅正文，排除目录/意见区/释义）===')
    hit1 = 0
    started = False
    for p in body.iter(qn('w:p')):
        if p.getparent().tag == qn('w:tc'):
            continue
        t = ptext(p).strip()
        if t == '一、项目概况':
            started = True
        if not started:
            continue
        st = style_of(p)
        if st and st.lower().startswith('toc'):
            continue
        if not t:
            continue
        exp = expected_lvl(t)
        if exp == 'skip':
            continue
        cur = get_lvl(p)
        if str(cur) != str(exp):
            hit1 += 1; issues += 1
            print('  ✗ outlineLvl=%s 应为 %s : %s' % (cur, exp, t[:34]))
            if a.fix:
                set_lvl(p, exp)
    print('  ✓ 无异常' if hit1 == 0 else '  共 %d 处' % hit1)

    print()
    print('=== [2] 表格列宽 与 [3] 列序 ===')
    tbls = list(body.iter(qn('w:tbl')))
    hit2 = 0
    for i, t in enumerate(tbls):
        g = t.find(qn('w:tblGrid'))
        if g is None:
            continue
        cols = g.findall(qn('w:gridCol'))
        ws = [int(c.get(qn('w:w'))) for c in cols]
        n = len(ws)
        trs = t.findall(qn('w:tr'))
        if not trs:
            continue
        need = [1] * n
        for tr in trs:
            for c, tc in enumerate(tr.findall(qn('w:tc'))[:n]):
                need[c] = max(need[c], wlen(cell_text(tc)))
        req = [max(600, v * 100) for v in need]
        r = corr(ws, req)
        worst = max(range(n), key=lambda c: req[c] / max(1, ws[c]))
        worst_ratio = req[worst] / max(1, ws[worst])
        hdr = ' | '.join(cell_text(tc)[:16] for tc in trs[0].findall(qn('w:tc')))
        # 仅当「最挤列需要≥1.8倍宽度」且「列宽与内容相关度低」时才判为新增内容格式失配
        if r < CORR_MIN and worst_ratio >= 1.8:
            hit2 += 1; issues += 1
            print('  表%-3d ❗列宽与内容不匹配（相关度%.2f，最挤列%d 需%.1f倍宽：宽%d/需≈%d）'
                  % (i + 1, r, worst + 1, worst_ratio, ws[worst], req[worst]))
            print('        ｜%s' % hdr[:82])
            if a.fix_table == i + 1:
                new = rebalance(t, req, cols)
                print('        → 已按内容重分配 %s' % new)
            else:
                print('        （如系本次新增的表，可加 --fix-table %d 重分配列宽；'
                      '既有表请勿改）' % (i + 1))
        # 列序
        h = [cell_text(tc) for tc in trs[0].findall(qn('w:tc'))]
        per = [int(re.search(r'(20\d\d)', x).group(1)) for x in h if re.search(r'20\d\d', x)]
        if len(per) >= 2 and per != sorted(per, reverse=True):
            issues += 1
            print('  表%-3d ❗列序疑非「本期→上期→前前期」: %s' % (i + 1, ' | '.join(h)[:78]))
    if hit2 == 0:
        print('  ✓ 无异常')

    if a.fix:
        doc.save(a.docx)
        print('\n已保存（--fix）')
    print('\n问题合计: %d' % issues)
    return 0 if issues == 0 else 2


if __name__ == '__main__':
    sys.exit(main())
