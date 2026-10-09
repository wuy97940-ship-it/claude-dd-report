# -*- coding: utf-8 -*-
"""
dd_format_probe.py — 尽调报告格式规范卡提取器 (v1.0, 20260910)
用法: python dd_format_probe.py <docx> [输出.md]
用途: 改docx前先跑本工具，提取该报告的格式规范(正文/备注/单位/标题/编号标题/表格)，
      并扫描"裸格式段落"(无pPr或run无rPr)——改前知规范、改后验格式，杜绝格式返工。
沉淀自: 某贸易客户戊尽调报告v5格式返工教训(dd-report skill v5.3)。
"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn


def get_ppr_sig(p):
    """提取段落格式签名"""
    ppr = p._p.pPr
    if ppr is None:
        return "无pPr"
    parts = []
    sp = ppr.find(qn('w:spacing'))
    if sp is not None:
        parts.append("line=%s/%s" % (sp.get(qn('w:line')), sp.get(qn('w:lineRule'))))
    ind = ppr.find(qn('w:ind'))
    if ind is not None:
        parts.append("firstLine=%s(%s)" % (ind.get(qn('w:firstLine')), ind.get(qn('w:firstLineChars'))))
    jc = ppr.find(qn('w:jc'))
    if jc is not None:
        parts.append("jc=%s" % jc.get(qn('w:val')))
    ol = ppr.find(qn('w:outlineLvl'))
    if ol is not None:
        parts.append("outlineLvl=%s" % ol.get(qn('w:val')))
    if ppr.find(qn('w:keepNext')) is not None:
        parts.append("keepNext")
    return ", ".join(parts) if parts else "pPr空"


def get_rpr_sig(p):
    """取段落第一个有格式run的rPr签名"""
    for r in p.runs:
        rpr = r._r.rPr
        if rpr is not None:
            rf = rpr.find(qn('w:rFonts'))
            fonts = []
            if rf is not None:
                for a in ('ascii', 'hAnsi', 'eastAsia', 'hint'):
                    v = rf.get(qn('w:' + a))
                    if v:
                        fonts.append("%s=%s" % (a, v))
            sz = rpr.find(qn('w:sz'))
            if sz is not None:
                fonts.append("sz=%s(半磅)" % sz.get(qn('w:val')))
            if rpr.find(qn('w:b')) is not None:
                fonts.append("b=加粗")
            return ", ".join(fonts) if fonts else "rPr空"
    return "run无rPr"


def main():
    if len(sys.argv) < 2:
        print("用法: python dd_format_probe.py <docx> [输出.md]")
        sys.exit(1)
    path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else path + "_格式规范卡.md"
    doc = Document(path)
    lines = ["# 格式规范卡: " + path.split("\\")[-1], ""]

    # 1. Normal样式
    st = doc.styles['Normal']
    sz = st.font.size.pt if st.font.size else "默认"
    lines += ["## 1. Normal样式", "- 字体: %s | 字号: %s pt" % (st.font.name, sz), ""]

    # 2. 分类取样
    paras = doc.paragraphs
    samples = {}
    bare = []
    for p in paras:
        t = p.text.strip()
        if not t:
            continue
        has_rpr = any(r._r.rPr is not None for r in p.runs) if p.runs else False
        if not has_rpr and len(t) > 20 and p.style.name == 'Normal':
            bare.append(t[:40])
        if "备注" in t[:3] and "备注段" not in samples:
            samples["备注段(五号?)"] = p
        if t.startswith("单位：") and "单位段" not in samples:
            samples["单位段(右对齐?)"] = p
        if re.match(r"^（[一二三四五六七八九十]+）", t) and p._p.pPr is not None \
                and p._p.pPr.find(qn('w:outlineLvl')) is not None \
                and p._p.pPr.find(qn('w:outlineLvl')).get(qn('w:val')) == '1' \
                and "二级标题" not in samples:
            samples["二级标题(outlineLvl=1)"] = p
        if re.match(r"^\d+、", t) and p._p.pPr is not None \
                and p._p.pPr.find(qn('w:outlineLvl')) is not None \
                and p._p.pPr.find(qn('w:outlineLvl')).get(qn('w:val')) == '2' \
                and "编号小标题" not in samples:
            samples["编号小标题(outlineLvl=2)"] = p
        if len(t) > 80 and "正文长段" not in samples and p.style.name == 'Normal' \
                and not t.startswith(("备注", "单位")):
            samples["正文长段"] = p
    lines += ["## 2. 分类格式样本（改docx新增段落时逐类克隆此格式）", ""]
    for label, p in samples.items():
        block = "### %s\n```\n文本: %s\npPr:  %s\nrPr:  %s\n```" % (
            label, p.text[:50], get_ppr_sig(p), get_rpr_sig(p))
        lines += [block, ""]

    # 3. 表格规范
    lines += ["## 3. 表格规范", "- 表格总数: %d" % len(doc.tables)]
    if doc.tables:
        t0 = doc.tables[0]
        tblPr = t0._tbl.tblPr
        tblW = tblPr.find(qn('w:tblW')) if tblPr is not None else None
        layout = tblPr.find(qn('w:tblLayout')) if tblPr is not None else None
        lines += ["- 首表tblW: %s/%s | tblLayout: %s" % (
            tblW.get(qn('w:type')) if tblW is not None else "未设",
            tblW.get(qn('w:w')) if tblW is not None else "",
            layout.get(qn('w:type')) if layout is not None else "未设")]
        autofit_n = 0
        for t in doc.tables:
            tp = t._tbl.tblPr
            if tp is not None:
                w = tp.find(qn('w:tblW'))
                ly = tp.find(qn('w:tblLayout'))
                if w is not None and w.get(qn('w:type')) == 'pct' and ly is not None and ly.get(qn('w:type')) == 'autofit':
                    autofit_n += 1
        lines += ["- 已设pct+autofit的表格: %d/%d" % (autofit_n, len(doc.tables))]
    lines.append("")

    # 4. 裸格式段落警告
    lines += ["## 4. 裸格式段落警告（run无rPr=格式割裂，改前须清零）", ""]
    if bare:
        for b in bare[:20]:
            lines.append("- ⚠ %s..." % b)
        if len(bare) > 20:
            lines.append("- ...共%d处" % len(bare))
    else:
        lines.append("✅ 无裸格式段落")
    lines.append("")

    # 5. 带圈序号hint检查
    hint_miss = []
    CIRCLES = "①②③④⑤⑥⑦⑧⑨⑩"
    for p in paras:
        for r in p.runs:
            if r.text and any(ch in r.text for ch in CIRCLES):
                rpr = r._r.rPr
                rf = rpr.find(qn('w:rFonts')) if rpr is not None else None
                if rf is None or rf.get(qn('w:hint')) != 'eastAsia':
                    hint_miss.append(r.text[:20])
    lines += ["## 5. 带圈序号hint检查（缺hint=eastAsia会被西文字体渲染变大）", ""]
    if hint_miss:
        for h in hint_miss[:10]:
            lines.append("- ⚠ run文本: %s" % h)
        lines.append("- 共%d处缺hint" % len(hint_miss))
    else:
        lines.append("✅ 全部带圈序号均带hint=eastAsia")

    out = "\n".join(lines)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(out)
    print(out)
    print("\n>>> 规范卡已保存: %s" % out_path)


if __name__ == '__main__':
    main()
