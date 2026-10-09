# -*- coding: utf-8 -*-
"""解析docx修订(tracked changes ins/del)与批注，输出：段落文本 | +新增 -删除 | 批注"""
import sys, zipfile
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W_ = lambda t: '{%s}%s' % (W, t)

def read_comments(path):
    comments = {}
    try:
        with zipfile.ZipFile(path) as z:
            if 'word/comments.xml' in z.namelist():
                root = ET.fromstring(z.read('word/comments.xml'))
                for c in root.iter(W_('comment')):
                    cid = c.get(W_('id'))
                    author = c.get(W_('author'), '')
                    date = c.get(W_('date'), '')
                    text = ''.join(t.text or '' for t in c.iter(W_('t')))
                    comments[cid] = (author, date, text)
    except Exception as e:
        print(f"[批注读取错误] {e}")
    return comments

def para_rev(p):
    """返回 (稳定文本, ins列表, del列表, comment_ids)"""
    st, ins, dels, cids = [], [], [], []
    for n in p.iter():
        tag = n.tag
        if tag == W_('t'):
            # 找祖先 ins/del
            anc = n.findall('ancestors')  # no-op
            return  # placeholder
    # 用迭代器手动取
    return ('', [], [], [])

def walk(p):
    st, ins, dels, cids = [], [], [], []
    # 按子元素顺序遍历
    for child in p:
        tag = child.tag
        if tag == W_('r') or tag == '{%s}r' % W:
            # run：检查其内部 ins/del
            pass
    return

def render(p):
    """渲染段落：{状态前缀}文本"""
    seg = []
    def rec(el, in_ins=False, in_del=False):
        # 段落根元素自身可能是p，迭代其子
        pass
    # 直接把p当el迭代
    def rec(el, in_ins=False, in_del=False):
        for c in el:
            tag = c.tag
            if tag == W_('t'):
                t = c.text or ''
                if in_del:
                    seg.append('[删除]' + t)
                elif in_ins:
                    seg.append('[新增]' + t)
                else:
                    seg.append(t)
            elif tag == W_('ins'):
                rec(c, in_ins=True, in_del=False)
            elif tag == W_('del'):
                # delText 特殊
                for dt in c.iter(W_('delText')):
                    seg.append('[删除]' + (dt.text or ''))
            elif tag == W_('r'):
                # run 内部可能有 ins/del 包裹
                if c.find(W_('ins')) is not None or c.find(W_('del')) is not None:
                    rec(c, in_ins, in_del)
                else:
                    for t in c.iter(W_('t')):
                        if in_del:
                            seg.append('[删除]' + (t.text or ''))
                        elif in_ins:
                            seg.append('[新增]' + (t.text or ''))
                        else:
                            seg.append(t.text or '')
            elif tag in (W_('hyperlink'), W_('smartTag'), W_('proofErr'), W_('bookmarkStart'), W_('bookmarkEnd'), W_('commentRangeStart'), W_('commentRangeEnd'), W_('commentReference')):
                rec(c, in_ins, in_del)
            else:
                rec(c, in_ins, in_del)
    rec(p)
    return ''.join(seg)

def parse(path):
    comments = read_comments(path)
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
        body = root.find(W_('body'))
        # 建立 commentRangeStart id→段落号 映射
        for i, child in enumerate(body):
            tag = child.tag.replace('{%s}' % W, '')
            if tag != 'p':
                continue
            text = render(child)
            cids = [m.get(W_('id')) for m in child.iter(W_('commentRangeStart'))]
            has_rev = ('[新增]' in text) or ('[删除]' in text)
            if has_rev or cids:
                print(f"\n[p{i}] {text[:500]}")
                for cid in cids:
                    if cid in comments:
                        a, d, ct = comments[cid]
                        print(f"    ↳批注[{d} {a}]: {ct}")

if __name__ == '__main__':
    parse(sys.argv[1])