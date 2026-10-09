# -*- coding: utf-8 -*-
"""读取尽调报告docx：正文结构（标题层级/段落+批注）+ 批注全文。用法: python dd_read_structure.py <docx> <outline|full|comments|body>"""
import sys, zipfile
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W_ = lambda t: '{%s}%s' % (W, t)

def para_text(p):
    return ''.join(t.text or '' for t in p.iter(W_('t')))

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

def parse(path, mode):
    comments = read_comments(path)
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
        body = root.find(W_('body'))
        if body is None:
            print("无body"); return

        idx = 0
        for child in body:
            tag = child.tag.replace('{%s}' % W, '')
            if tag == 'p':
                p = child
                cids = [m.get(W_('id')) for m in p.iter(W_('commentRangeStart'))]
                text = para_text(p)
                style = ''
                outline = ''
                pPr = p.find(W_('pPr'))
                if pPr is not None:
                    ps = pPr.find(W_('pStyle'))
                    if ps is not None:
                        style = ps.get(W_('val'), '')
                    ol = pPr.find(W_('outlineLvl'))
                    if ol is not None:
                        outline = ol.get(W_('val'), '')
                is_head = (style.startswith('Heading') or style in ('1', '2', '3')
                           or outline in ('0', '1', '2', '3', '4', '5'))
                if mode == 'outline':
                    if is_head or cids:
                        marker = '  ' * (int(outline) if outline.isdigit() else 0)
                        seg = f"[p{idx} s={style} ol={outline}] {text[:70]}"
                        if cids:
                            for cid in cids:
                                if cid in comments:
                                    a, d, ctext = comments[cid]
                                    seg += f"\n     ↳批注[{d} {a}]: {ctext}"
                        print(marker + seg)
                elif mode == 'full':
                    if text.strip() or cids:
                        print(f"[p{idx} s={style}] {text}")
                        for cid in cids:
                            if cid in comments:
                                a, d, ctext = comments[cid]
                                print(f"    ↳批注[{d} {a}]: {ctext}")
                elif mode == 'comments':
                    for cid in cids:
                        if cid in comments:
                            a, d, ctext = comments[cid]
                            print(f"[p{idx} s={style}] {text[:50]}  →批注[{d} {a}]: {ctext}")
                idx += 1
            elif tag == 'tbl':
                if mode == 'full':
                    print(f"<表格 @p{idx}>")
                    for row in child.iter(W_('tr')):
                        cells = []
                        for tc in list(row.iter(W_('tc')))[:10]:
                            t = ''.join(x.text or '' for x in tc.iter(W_('t')))
                            cells.append(t[:35])
                        print('  | ' + ' | '.join(cells))
                    print('</表格>')
                idx += 1

if __name__ == '__main__':
    mode = sys.argv[2] if len(sys.argv) > 2 else 'outline'
    parse(sys.argv[1], mode)