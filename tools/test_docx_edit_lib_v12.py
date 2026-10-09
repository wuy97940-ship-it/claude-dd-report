# -*- coding: utf-8 -*-
"""docx_edit_lib v1.2 冒烟测试：复现20260921加粗切run bug场景 + 锚点守卫"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r'tools')
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
import docx_edit_lib as lib

W = lib.W
def q(t): return lib.q(t)

ok = fail = 0
def check(name, cond, detail=''):
    global ok, fail
    if cond:
        ok += 1; print(f'[PASS] {name}')
    else:
        fail += 1; print(f'[FAIL] {name} {detail}')


def make_doc(paras_xml_texts):
    """构造测试docx：每项为(文本, 是否含批注标记)"""
    doc = Document()
    body = doc.element.body
    ps = []
    for text, with_cmt in paras_xml_texts:
        p = doc.add_paragraph()._p
        r = p.makeelement(q('r'), {})
        t = r.makeelement(q('t'), {}); t.text = text
        r.append(t); p.append(r)
        if with_cmt:
            ppr = p.find(q('pPr'))
            crs = p.makeelement(q('commentRangeStart'), {}); crs.set(q('id'), '99')
            cre = p.makeelement(q('commentRangeEnd'), {}); cre.set(q('id'), '99')
            # crs插段中、ref run插cre前——模拟"区间夹着正文run"的挤空场景
            p.insert(1 if ppr is None else 2, crs)
            ref_r = p.makeelement(q('r'), {})
            ref = ref_r.makeelement(q('commentReference'), {}); ref.set(q('id'), '99')
            ref_r.append(ref)
            p.append(cre); p.append(ref_r)
        ps.append(p)
    return doc, ps


# ===== 场景1：bug复现——同run双命中 + 子串嵌套 =====
TEXT1 = ('子公司花桥矿业因采矿许可证变更开采矿种计提矿业权出让收益金14,421.93万元计入税金及附加，'
         '导致税金及附加同比激增342.26%；有息负债为短期借款971.00万元+一年内到期长期借款137.01万元+长期借款1,161.22万元。')
doc, ps = make_doc([(TEXT1, False)])
p1 = ps[0]._p if hasattr(ps[0], '_p') else ps[0]

n1 = lib.bold_text_tracked(p1, '税金及附加', next_id=lambda: iter([1, 2, 3, 4, 5, 6, 7, 8]).__next__())
n2 = lib.bold_text_tracked(p1, '一年内到期长期借款', next_id=lambda: iter([11, 12, 13]).__next__())
n3 = lib.bold_text_tracked(p1, '长期借款', next_id=lambda: iter([21, 22, 23, 24, 25]).__next__())

check('同run双命中"税金及附加"加粗2处', n1 == 2, f'实际{n1}')
check('"一年内到期长期借款"加粗1处', n2 == 1, f'实际{n2}')
check('裸"长期借款"仅命中1处(嵌套不重复)', n3 == 1, f'实际{n3}')

# 文本完整性：所有w:t（含新run）拼接 == 原文
new_text = ''.join(t.text or '' for t in p1.iter(q('t')))
check('加粗后正文文本零变化（无重复/丢字）', new_text == TEXT1, f'\n  期望:{TEXT1}\n  实际:{new_text}')

# 加粗正确性：逐字符比对粗/非粗
def char_bold_map(p_el):
    m = []
    for r in p_el.findall(q('r')):
        rpr = r.find(q('rPr'))
        b = rpr is not None and rpr.find(q('b')) is not None
        for ch in ''.join(t.text or '' for t in r.findall(q('t'))):
            m.append((ch, b))
    return m

bm = char_bold_map(p1)
def span_set(phrase):
    s = set()
    start = 0
    while True:
        i = TEXT1.find(phrase, start)
        if i < 0:
            break
        s.update(range(i, i + len(phrase)))
        start = i + 1
    return s
expect_bold = span_set('税金及附加') | span_set('一年内到期长期借款') | span_set('长期借款')
got_bold = {i for i, (ch, b) in enumerate(bm) if b}
check('加粗字符区间精确（不多不少）', got_bold == expect_bold,
      f'多粗:{sorted(got_bold - expect_bold)[:10]} 漏粗:{sorted(expect_bold - got_bold)[:10]}')

# rPrChange留档
n_chg = len(p1.findall('.//' + q('rPrChange')))
check('格式修订rPrChange留档', n_chg >= 4, f'rPrChange×{n_chg}')

# ===== 场景2：锚点守卫——批注区间夹正文run，转删后锚点仍包住全段 =====
TEXT2 = '选取碳酸锂生产商中资源路线相近的A股上市公司作为可比对象，盐湖提锂企业不纳入本次对比。'
doc2, ps2 = make_doc([(TEXT2, True)])
p2 = ps2[0]._p if hasattr(ps2[0], '_p') else ps2[0]

moved = lib.guard_comment_anchors(p2)
check('守卫报告移动3个锚点元素(crs+cre+ref)', moved == 3, f'实际{moved}')
kids = [etree.QName(c).localname for c in p2]
check('crs在段首区(pPr后第一个)', kids[0] in ('commentRangeStart', 'pPr') and kids.index('commentRangeStart') <= 1, str(kids[:4]))
check('cre在段末', kids[-1] in ('commentRangeEnd', 'r') and 'commentRangeEnd' in kids[-2:], str(kids[-3:]))
# ref run 紧跟 cre 之后
cre_pos = kids.index('commentRangeEnd')
check('ref run紧跟cre之后', kids[cre_pos + 1] == 'r' and p2[cre_pos + 1].find(q('commentReference')) is not None)

n_del = lib.tracked_delete_para_runs(p2, next_id=lambda: iter([31, 32, 33, 34, 35]).__next__())
check('转删正文run', n_del >= 1, f'实际{n_del}')
# 转删后：crs...del...cre ref 结构完整（锚点未被挤空、未被删除）
kids2 = [etree.QName(c).localname for c in p2]
check('转删后crs仍在段首', 'commentRangeStart' in kids2[:2])
check('转删后cre与ref仍在段末', 'commentRangeEnd' in kids2[-3:] and kids2[-1] == 'r')
check('转删后del存在且delText含原文', any(len(t.text or '') > 10 for t in p2.iter(q('delText'))))

# ===== 场景3：max_hits 死循环保护 =====
doc3, ps3 = make_doc([('重复重复重复', False)])
p3 = ps3[0]._p if hasattr(ps3[0], '_p') else ps3[0]
try:
    # 已粗文本再跑同词：NUL掩码应命中0次正常返回，不死循环
    lib.bold_text_tracked(p3, '重复', max_hits=5)
    lib.bold_text_tracked(p3, '重复', max_hits=5)  # 第二遍全已粗→0命中
    check('已粗重跑0命中正常返回（无死循环）', True)
except RuntimeError as e:
    check('已粗重跑0命中正常返回（无死循环）', False, str(e))

# 圈码hint联动
doc4, ps4 = make_doc([('①杠杆结构方面有息负债较低', False)])
p4 = ps4[0]._p if hasattr(ps4[0], '_p') else ps4[0]
lib.bold_text_tracked(p4, '①杠杆结构', next_id=lambda: iter([41, 42, 43]).__next__())
rf = p4.find(q('r') + '/' + q('rPr') + '/' + q('rFonts'))
check('圈码run自动补hint=eastAsia', rf is not None and rf.get(q('hint')) == 'eastAsia')

print(f'\n===== 冒烟测试汇总: PASS={ok} FAIL={fail} =====')
sys.exit(1 if fail else 0)
