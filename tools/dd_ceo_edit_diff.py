# -*- coding: utf-8 -*-
"""终审人手改回收器（DD CEO-Edit Diff）v1.0  2026-09-28
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
批注/修订轮交付后，终审人手改定稿——本工具一条命令找全手改点，供"手改仅学习"
提炼偏好。源自20260928客户丙四主体ChatGPT意见轮实战（tmp_learn_full_diff.py固化）。

流程：
  ① [--surgery op脚本] 对修订前备份重放本轮手术（自动换SRC/DST、剥stdout包装）
  ② 对重放结果做"接受修订模拟"（ins解包/del删除）→ 理论交付版
  ③ 理论版 vs 手改后版 全文dump diff
  ④ 过滤目录页码漂移类噪声 → 输出手改点清单（终端+--out md）

用法：
  D:\\Python312\\python.exe tools\\dd_ceo_edit_diff.py <修订前备份.docx> <手改后.docx> ^
      [--surgery _workN\\op_xx_surgery.py] [--out 手改清单.md]

铁律：只读。绝不写两个输入docx。
"""
import sys, io, os, re, copy, difflib, zipfile, argparse, tempfile
sys.path.insert(0, r"tools")
from lxml import etree
import docx_edit_lib as L


def dump_docx(path):
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(path)
    lines = []
    for child in d.element.body.iterchildren():
        if child.tag.endswith('}p'):
            t = Paragraph(child, d).text.strip()
            if t:
                lines.append("P: " + t)
        elif child.tag.endswith('}tbl'):
            lines.append("===== TABLE =====")
            for r in Table(child, d).rows:
                cells = [c.text.strip().replace("\n", "/") for c in r.cells]
                dedup = []
                for c in cells:
                    if not dedup or dedup[-1] != c:
                        dedup.append(c)
                lines.append(" | ".join(dedup))
            lines.append("===== END TABLE =====")
    return lines


def accept_revisions(docx_path):
    """接受全部修订（ins解包/del删除/trPr-ins与cellIns保留），返回新文件路径。"""
    parts, root, body = L.load_docx(docx_path)
    sim = copy.deepcopy(root)
    for ins_el in list(sim.iter(L.q('ins'))):
        parent = ins_el.getparent()
        if parent is None:
            continue
        if parent.tag == L.q('p'):
            idx = list(parent).index(ins_el)
            for i, ch in enumerate(list(ins_el)):
                parent.insert(idx + i, ch)
            parent.remove(ins_el)
        else:
            parent.remove(ins_el)
    for del_el in list(sim.iter(L.q('del'))):
        parent = del_el.getparent()
        if parent is not None:
            parent.remove(del_el)
    out = docx_path + ".accept_tmp.docx"
    tmp = out + ".new"
    with zipfile.ZipFile(docx_path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            data = zin.read(n)
            if n == 'word/document.xml':
                data = etree.tostring(sim, xml_declaration=True, encoding='UTF-8', standalone=True)
            zout.writestr(n, data)
    os.replace(tmp, out)
    return out


def replay_surgery(surgery_path, backup_path):
    """对备份重放手术脚本：替换SRC=backup、DST=临时文件，剥stdout包装行。"""
    code = open(surgery_path, encoding="utf-8").read()
    code = "\n".join(l for l in code.split("\n") if "TextIOWrapper" not in l)
    # SRC：备份路径（须为已含r""的赋值行）
    m = re.search(r'^SRC\s*=\s*r?"([^"]+)"', code, re.M)
    if not m:
        raise RuntimeError("手术脚本无 SRC = r\"...\" 赋值行")
    code = code[:m.start()] + f'SRC = r"{backup_path}"' + code[m.end():]
    m = re.search(r'^DST\s*=\s*.*$', code, re.M)
    if not m:
        raise RuntimeError("手术脚本无 DST 赋值行")
    dst = os.path.join(tempfile.gettempdir(), "ceo_diff_replay.docx")
    code = code[:m.start()] + f'DST = r"{dst}"' + code[m.end():]
    g = {'__name__': 'replay', '__file__': surgery_path}
    exec(compile(code, surgery_path, "exec"), g)
    if not os.path.exists(dst):
        raise RuntimeError("重放未产出文件")
    return dst


def is_toc_line(ln):
    """目录条目行：P: 开头且以 1-4位数字 结尾且含制表定位（页码引用）。"""
    return ln.startswith("P:") and re.search(r"\t\s*\d{1,4}$", ln)


def toc_key(ln):
    return re.sub(r"\t\s*\d{1,4}$", "", ln)


def main():
    ap = argparse.ArgumentParser(description='终审人手改回收器')
    ap.add_argument('backup', help='修订前备份docx')
    ap.add_argument('after', help='终审人手改后docx')
    ap.add_argument('--surgery', default='', help='本轮手术op脚本（重放生成理论版）')
    ap.add_argument('--out', default='', help='手改清单输出md')
    a = ap.parse_args()

    print('=' * 78)
    print('终审人手改回收器：%s' % os.path.basename(a.after))
    print('=' * 78)

    if a.surgery:
        print('① 重放手术: %s' % os.path.basename(a.surgery))
        theory_docx = replay_surgery(a.surgery, a.backup)
        print('   → 接受修订模拟')
        theory_docx = accept_revisions(theory_docx)
    else:
        print('① 未指定--surgery，直接对备份做接受修订模拟（适用于备份即含修订的版本）')
        theory_docx = accept_revisions(a.backup)

    print('② dump 两版全文')
    th = dump_docx(theory_docx)
    af = dump_docx(a.after)

    print('③ diff（过滤目录页码漂移）')
    # 目录页码行归一：页码不同视为同一条目（合并报"页码漂移N条"）
    th_n, af_n, drift = [], [], 0
    from collections import Counter
    th_cnt = Counter(toc_key(l) for l in th if is_toc_line(l))
    af_cnt = Counter(toc_key(l) for l in af if is_toc_line(l))
    for k in set(th_cnt) | set(af_cnt):
        if th_cnt.get(k, 0) != af_cnt.get(k, 0):
            drift += 1
    th_n = [l for l in th if not is_toc_line(l)]
    af_n = [l for l in af if not is_toc_line(l)]

    diffs = []
    for line in difflib.unified_diff(th_n, af_n, lineterm='', n=0):
        if line.startswith(('---', '+++')):
            continue
        diffs.append(line)

    n_ceo = sum(1 for d in diffs if d[0] in '+-')
    print('\n手改差异行: %d（另有目录页码漂移 %d 条，已归并）' % (n_ceo, drift))
    print('-' * 78)
    report = ["# 终审人手改清单（%s）" % os.path.basename(a.after), ""]
    if drift:
        report.append("· 目录页码漂移 %d 条（内容增删致页码后移，正常）" % drift)
        print('· 目录页码漂移 %d 条（正常）' % drift)
    cur_del = None
    for d in diffs:
        tag, txt = d[0], d[1:].strip()
        if not txt:
            continue
        if tag == '-':
            cur_del = txt
        elif tag == '+' and cur_del is not None:
            print('\n【改】原: %s\n     新: %s' % (cur_del[:150], txt[:150]))
            report.append("## 改\n- 原: %s\n- 新: %s" % (cur_del, txt))
            cur_del = None
        elif tag == '+':
            print('\n【增】%s' % txt[:180])
            report.append("## 增\n- %s" % txt)
        else:
            print('\n【删】%s' % txt[:180])
            report.append("## 删\n- %s" % txt)
    print('\n' + '=' * 78)
    if a.out:
        with open(a.out, 'w', encoding='utf-8') as f:
            f.write("\n".join(report))
        print('清单 -> %s' % a.out)
    print('提醒：手改仅学习不触碰；改后文档为定稿，不再做任何写入。')
    return 0


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.exit(main())
