# -*- coding: utf-8 -*-
# dd_extract_docx.py <input.docx> <output.txt>
# 通用docx提取：按文档顺序导出全部段落（含样式名）与表格内容。
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

def iter_block_items(parent):
    from docx.oxml.ns import qn
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)

def main():
    src, dst = sys.argv[1], sys.argv[2]
    doc = Document(src)
    lines = []
    t_idx = 0
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            txt = block.text.strip()
            style = block.style.name if block.style else ''
            if txt:
                lines.append('[P|%s] %s' % (style, txt))
        else:
            t_idx += 1
            lines.append('=== TABLE %d (rows=%d, cols=%d) ===' % (t_idx, len(block.rows), len(block.columns)))
            for r in block.rows:
                cells = [c.text.strip().replace('\n', ' / ') for c in r.cells]
                out = []
                prev = object()
                for c in cells:
                    if c != prev:
                        out.append(c)
                    prev = c
                lines.append(' | '.join(out))
            lines.append('=== END TABLE %d ===' % t_idx)
    with open(dst, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print('paragraph/table dump done: %d lines' % len(lines))

if __name__ == '__main__':
    main()
