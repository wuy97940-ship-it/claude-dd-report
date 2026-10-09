# -*- coding: utf-8 -*-
"""尽调报告交付前自检门禁（DD Report Self-Check）v1.0  2026-09-13

把历轮反复手工做的检查固化为一条命令。覆盖：
  [1] 带圈序号字体归属（①②③ 的 run 必须带 w:hint="eastAsia"，否则被西文字体接管、显示偏大）
  [2] 批注锚点完好性（commentRangeStart / End / Reference 三件套计数一致；comments.xml 每条都有锚点）
  [3] 标点配对与半角混入（中文引号/括号/书名号/方头括号；中英标点混用）
  [4] 表格结构（逐行列数一致；报告行×列；识别合并单元格）
  [5] 术语口径统一（--rename "旧=新,旧2=新2"）
  [6] 废弃值残留（--stale "值1,值2"）
  [7] 完全重复段落（疑误复制）
  [8] 修订/批注统计与作者

用法：
  python dd_report_selfcheck.py 报告.docx
  python dd_report_selfcheck.py 报告.docx --rename "现金成本=现金生产成本" --stale "35,000,2,288.60"
  python dd_report_selfcheck.py 报告.docx --json out.json
退出码：0=全过（可能有 WARN）；2=有 FAIL。
"""
import sys, os, io, json, zipfile, re, argparse
from collections import Counter
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def q(t): return W + t

CIRC = re.compile(r'[①-⑳㉑-㉟㊱-㊿]')
HALF = re.compile(r'[一-鿿][,;:?!]|[,;:?!][一-鿿]')


def load(path):
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read('word/document.xml'))
    return z, root, root.find(q('body'))


def eff(p):
    """接受修订后的有效文本"""
    out = []
    for n in p.iter():
        if n.tag == q('t'):
            a = n.getparent(); skip = False
            while a is not None and a.tag != q('p'):
                if a.tag == q('del'): skip = True; break
                a = a.getparent()
            if not skip: out.append(n.text or '')
    return ''.join(out)


def raw(p):
    out = []
    for n in p.iter():
        if n.tag == q('t'): out.append(n.text or '')
        elif n.tag == q('delText'): out.append('«D:' + (n.text or '') + '»')
    return ''.join(out)


def collect(body):
    """按文档顺序收集文本单元：('P', 段号, elem, text) / ('T', 表号, elem, text)"""
    units = []
    pi = ti = 0

    def walk(el):
        nonlocal pi, ti
        for c in el:
            if c.tag == q('p'):
                pi += 1
                units.append(('P%04d' % pi, c, eff(c), raw(c)))
            elif c.tag == q('tbl'):
                ti += 1
                for ri, tr in enumerate(c.findall(q('tr'))):
                    for ci, tc in enumerate(tr.findall(q('tc'))):
                        t = ' '.join(eff(p) for p in tc.findall(q('p')))
                        r = ' '.join(raw(p) for p in tc.findall(q('p')))
                        units.append(('T%d-R%dC%d' % (ti, ri + 1, ci + 1), tc, t, r))
            else:
                walk(c)
    walk(body)
    return units


def main():
    ap = argparse.ArgumentParser(description='尽调报告交付前自检门禁')
    ap.add_argument('docx')
    ap.add_argument('--rename', default='', help='术语统一: "旧=新|旧2=新2"（用 | 分隔，避免千分位逗号冲突）')
    ap.add_argument('--stale', default='', help='废弃值: "值1|值2"（用 | 分隔）')
    ap.add_argument('--json', default='', help='结果写 JSON')
    a = ap.parse_args()

    z, root, body = load(a.docx)
    units = collect(body)
    R = {'file': os.path.basename(a.docx), 'fails': [], 'warns': [], 'passes': []}
    def FAIL(m): R['fails'].append(m); print('  ✗ FAIL  %s' % m)
    def WARN(m): R['warns'].append(m); print('  ! WARN  %s' % m)
    def PASS(m): R['passes'].append(m); print('  ✓ %s' % m)

    print('=' * 78)
    print('尽调报告自检：%s' % os.path.basename(a.docx))
    print('  文本单元 %d 个（段落+表格单元格）' % len(units))
    print('=' * 78)

    # ---- [1] 带圈序号字体归属 ----
    print('\n[1] 带圈序号字体归属（hint=eastAsia）')
    bad = []
    for r in root.iter(q('r')):
        if not any(t.text and CIRC.search(t.text) for t in r.findall(q('t'))): continue
        rp = r.find(q('rPr')); rf = rp.find(q('rFonts')) if rp is not None else None
        if rf is None or rf.get(q('hint')) != 'eastAsia':
            bad.append(''.join(t.text or '' for t in r.findall(q('t')))[:16])
    if bad: FAIL('圈码 run 缺 hint=eastAsia：%d 处 → %s' % (len(bad), bad[:8]))
    else:   PASS('全部圈码 run 带 hint=eastAsia')

    # ---- [2] 批注锚点 ----
    print('\n[2] 批注锚点完好性')
    ns = len(list(root.iter(q('commentRangeStart'))))
    ne = len(list(root.iter(q('commentRangeEnd'))))
    nr = len(list(root.iter(q('commentReference'))))
    nc = 0
    if 'word/comments.xml' in z.namelist():
        nc = len(list(etree.fromstring(z.read('word/comments.xml')).iter(q('comment'))))
    if ns == ne == nr:
        PASS('锚点三件套一致 start=%d end=%d ref=%d' % (ns, ne, nr))
    else:
        FAIL('锚点不配对 start=%d end=%d ref=%d' % (ns, ne, nr))
    if nc != ns:
        FAIL('comments.xml 批注数(%d) 与锚点数(%d) 不一致 → 存在悬空批注或孤儿锚点' % (nc, ns))
    else:
        PASS('comments.xml 批注数 = 锚点数 = %d' % nc)

    # ---- [3] 标点 ----
    print('\n[3] 标点配对与半角混入')
    all_t = '\n'.join(u[2] for u in units)
    PAIRS = [('“', '”', '中文双引号'), ('‘', '’', '中文单引号'),
             ('（', '）', '中文括号'), ('《', '》', '书名号'), ('【', '】', '方头括号')]
    unpair = []
    for x, y, nm in PAIRS:
        cx, cy = all_t.count(x), all_t.count(y)
        if cx != cy:
            FAIL('%s 不配对 %s×%d / %s×%d' % (nm, x, cx, y, cy))
            unpair.append(nm)
    if not unpair:
        PASS('引号/括号/书名号/方头括号 全部配对')
    hm = [(u[0], HALF.search(u[2]).group(0)) for u in units if HALF.search(u[2])]
    if hm: FAIL('半角标点混入中文：%d 处 %s' % (len(hm), hm[:6]))
    else:  PASS('无半角标点混入')

    # ---- [4] 表格结构 ----
    print('\n[4] 表格结构')
    tbls = body.findall(q('tbl'))
    print('  表数量 %d' % len(tbls))
    for i, t in enumerate(tbls, 1):
        trs = t.findall(q('tr'))
        widths = []
        for tr in trs:
            w = 0
            for tc in tr.findall(q('tc')):
                tcpr = tc.find(q('tcPr'))
                gs = tcpr.find(q('gridSpan')) if tcpr is not None else None
                w += int(gs.get(q('val'))) if gs is not None else 1
            widths.append(w)
        uniq = sorted(set(widths))
        flag = '✓' if len(uniq) == 1 else '!'
        print('   %s 表%-2d 行=%-3d 列宽=%s' % (flag, i, len(trs),
              uniq[0] if len(uniq) == 1 else '%s(不一致)' % uniq))
        if len(uniq) != 1:
            WARN('表%d 各行逻辑列数不一致：%s' % (i, uniq))

    # ---- [5] 术语口径统一 ----
    print('\n[5] 术语口径统一')
    if a.rename:
        for pair in [p for p in re.split(r'[|;]', a.rename) if '=' in p]:
            old, new = pair.split('=', 1)
            hits = [(u[0], u[2]) for u in units if old in u[2]]
            hits = [(loc, t) for loc, t in hits if not (new in t and old not in t.replace(new, ''))]
            # 排除"旧词是新词子串"的正常情况
            real = []
            for loc, t in hits:
                tt = t.replace(new, '')
                if old in tt: real.append((loc, t))
            if real:
                FAIL('术语未统一「%s」→「%s」：%d 处 %s' % (old, new, len(real), [x[0] for x in real[:8]]))
            else:
                PASS('「%s」已统一为「%s」' % (old, new))
    else:
        print('  （未指定 --rename）')

    # ---- [6] 废弃值残留 ----
    print('\n[6] 废弃值残留')
    if a.stale:
        for v in [x for x in re.split(r'[|;]', a.stale) if x.strip()]:
            hits = [(u[0], u[2]) for u in units if v in u[2]]
            if hits:
                FAIL('废弃值「%s」残留 %d 处 %s' % (v, len(hits), [x[0] for x in hits[:8]]))
            else:
                PASS('无「%s」' % v)
    else:
        print('  （未指定 --stale）')

    # ---- [7] 重复段落 ----
    print('\n[7] 完全重复段落')
    ps = [u[2].strip() for u in units if u[0].startswith('P') and len(u[2].strip()) > 25]
    dup = {k: c for k, c in Counter(ps).items() if c > 1}
    if dup:
        for k, c in list(dup.items())[:5]:
            WARN('重复 %d 次：「%s…」（确认是否误复制；意见区与正文有意重复可忽略）' % (c, k[:40]))
    else:
        PASS('无完全重复长段落')

    # ---- [8] 修订/批注统计 ----
    print('\n[8] 修订 / 批注统计')
    ins = [n for n in root.iter(q('ins')) if n.get(q('author'))]
    dels = [n for n in root.iter(q('del')) if n.get(q('author'))]
    print('  w:ins=%d  w:del=%d  作者=%s' % (len(ins), len(dels),
          dict(Counter(n.get(q('author')) for n in ins))))
    print('  批注=%d 条，作者=%s' % (nc,
          dict(Counter(c.get(q('author')) for c in
                       etree.fromstring(z.read('word/comments.xml')).iter(q('comment'))))
          if nc else '{}'))

    # ---- [9] 渲染目检需求探测（20260921 新增）----
    # run级切分/格式修订/标插行表的正确性，文本层门禁覆盖有限——
    # 20260921客户庚轮加粗切run致文字重复错乱，selfcheck/consistency全绿，仅PNG渲染目检抓到。
    print('\n[9] 渲染目检需求探测')
    n_rprchg = len(list(root.iter(q('rPrChange'))))
    n_trins = len([t for t in root.iter(q('trPr')) if t.find(q('ins')) is not None])
    n_cellins = len(list(root.iter(q('cellIns'))))
    if n_rprchg or n_trins or n_cellins:
        WARN('本稿含格式修订 rPrChange×%d / 标插行 trPr-ins×%d / cellIns×%d —— run级切分与结构级改动'
             '须渲染验收：交付前导PDF对新增/改动页PNG目检（文字重复、格式突变、表格错位只有渲染能看见）'
             % (n_rprchg, n_trins, n_cellins))
    else:
        PASS('无格式修订/标插结构，常规门禁覆盖充分')

    # ---- [10] 正文段挂目录样式污染扫描（20260924 客户丙之子公司丙1新增）----
    # builder从母版TOC缓存条目抓格式模板→正文段挂TOC样式+左缩进（中文样式名"目录2"令
    # 按名过滤失效，20260924终审人批注抓出）。样式名判断不可靠，改按styleId+文本长度防御。
    print('\n[10] 正文段目录样式污染')
    toc_bad = []
    body_elm = root.find(q('body'))
    for i, p in enumerate(body_elm.findall(q('p'))):
        pPr = p.find(q('pPr'))
        if pPr is None:
            continue
        ps = pPr.find(q('pStyle'))
        if ps is None:
            continue
        sid = (ps.get(q('val')) or '').upper()
        if not sid.startswith('TOC'):
            continue
        txt = ''.join(t.text or '' for t in p.iter(q('t'))).strip()
        tabs = pPr.findall(q('tabs') + '/' + q('tab'))
        dot = any(tb.get(q('leader')) == 'dot' for tb in tabs)
        import re as _re
        page_tail = bool(_re.search(r'\d{1,4}\s*$', txt))
        if len(txt) > 25 and not (dot and page_tail):
            toc_bad.append((i, sid, txt[:18]))
    if toc_bad:
        FAIL('正文段挂TOC样式×%d（意见区/正文被母版目录缓存污染）→ %s' % (len(toc_bad), toc_bad[:4]))
    else:
        PASS('无正文段挂TOC样式')

    # ---- [11] 周转天数类指标自洽（20260928 客户丙轮新增）----
    # 现金周转期＝应收+存货-应付，表内列示值必须可复算；应付口径含票据差异必须有
    # "含票据"字样佐证（行名限定语或表下备注）。20260928客户丙轮：交付版丢失底稿
    # "（含应付票据口径）"行名+表下备注，96/187/387复算不出-187，被外部AI质疑勾稽。
    print('\n[11] 周转天数类指标自洽（现金周转期可复算）')
    import re as _re11
    ccc_checked = 0
    ccc_bad = []
    for i, t in enumerate(body.findall(q('tbl')), 1):
        rows_txt = []
        for tr in t.findall(q('tr')):
            cells = [''.join(x.itertext()).strip() for x in tr.findall(q('tc'))]
            rows_txt.append(cells)
        def _num(cells, ncols=1):
            vals = []
            for c in cells[1:1 + 4]:
                m = _re11.search(r'-?\d[\d,]*\.?\d*', c)
                if m:
                    vals.append(float(m.group(0).replace(',', '')))
            return vals[0] if vals else None
        get = {}
        for cells in rows_txt:
            if not cells:
                continue
            head = cells[0]
            for key in ('应收账款周转天数', '存货周转天数', '应付账款周转天数', '现金周转期'):
                if key in head and key not in get:
                    get[key] = _num(cells)
        if '现金周转期' not in get:
            continue
        missing = [k for k in ('应收账款周转天数', '存货周转天数', '应付账款周转天数')
                   if k not in get or get[k] is None]
        if missing:
            WARN('表%d 含现金周转期但周转天数取不全（缺%s，或该表本无此行），人工复算：%s'
                 % (i, '/'.join(missing), get))
            ccc_checked += 1
            continue
        ccc_checked += 1
        calc = get['应收账款周转天数'] + get['存货周转天数'] - get['应付账款周转天数']
        if abs(calc - get['现金周转期']) <= 1.5:
            continue
        # 不自洽→找"含票据"口径佐证（表内全部文本+表后紧邻2段）
        tbl_all = ' '.join(' '.join(c) for c in rows_txt)
        body_ps = body.findall(q('p'))
        tbl_idx = list(body).index(t)
        ctx = tbl_all
        for p2 in list(body)[tbl_idx + 1:tbl_idx + 3]:
            if p2.tag == q('p'):
                ctx += ' ' + ''.join(x.text or '' for x in p2.iter(q('t')))
        if ('含票据' in ctx) or ('含应付票据' in ctx):
            WARN('表%d 现金周转期%.0f≠应收+存货-应付=%.0f，但有"含票据"口径佐证（复核行名限定语是否完整）'
                 % (i, get['现金周转期'], calc))
        else:
            ccc_bad.append('表%d 现金周转期%.0f ≠ 应收%.0f+存货%.0f-应付%.0f=%.0f，且无"含票据"口径说明'
                           % (i, get['现金周转期'], get['应收账款周转天数'], get['存货周转天数'],
                              get['应付账款周转天数'], calc))
    if ccc_bad:
        for m2 in ccc_bad:
            FAIL(m2)
    elif ccc_checked:
        PASS('现金周转期 %d 处全部可复算（或口径佐证完整）' % ccc_checked)
    else:
        PASS('无现金周转期指标')

    # ---- 汇总 ----
    print('\n' + '=' * 78)
    print('汇总: FAIL=%d  WARN=%d  PASS=%d' % (len(R['fails']), len(R['warns']), len(R['passes'])))
    for m in R['fails']: print('  ❌ %s' % m)
    for m in R['warns']: print('  ⚠  %s' % m)
    print('=' * 78)
    if a.json:
        json.dump(R, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
        print('JSON -> %s' % a.json)
    return 2 if R['fails'] else 0


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.exit(main())
