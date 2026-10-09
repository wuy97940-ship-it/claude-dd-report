# -*- coding: utf-8 -*-
"""
dd_word_finalize.py — 尽调报告 Word 定稿器 v1.0（2026-09-11 某锂盐客户庚v2复盘沉淀）
================================================================================
替代 ps1 + Word COM 手写脚本，根除两大致命坑：
  【坑1】ps1 中文路径无 BOM → PowerShell 5.1 按 ANSI 读 → 路径乱码 → Word 报"文档已损坏"
  【坑2】Word 应用层 Track Changes 开关会跨会话继承 → TOC 刷新被记成修订（实测 86条ins+86条del）

正确的定稿顺序（不可颠倒）：
    打开 → TrackRevisions=False（必须在任何改动之前）→ AcceptAll → 刷新目录/域 → 保存

用法：
  # 仅体检（不改文件）
  python dd_word_finalize.py X.docx --verify-only

  # 定稿：关修订 + 刷新目录 + 保存 + 导出PDF
  python dd_word_finalize.py X.docx --pdf out.pdf

  # 定稿并接受全部修订（把终审人手改/Claude修订转为正式文本）
  python dd_word_finalize.py X.docx --accept --pdf out.pdf

退出码：0=成功 2=失败
"""
import os
import sys
import argparse
import time

WD_FORMAT_PDF = 17
WD_STAT_PAGES = 2
WD_STAT_WORDS = 0
WD_STAT_PARAS = 4


def _open_word():
    import win32com.client as wc
    import pythoncom
    pythoncom.CoInitialize()
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    return word


def finalize(path, pdf=None, accept=False, update_toc=True, verify_only=False):
    path = os.path.abspath(path)
    if not os.path.exists(path):
        print('!! 文件不存在: %s' % path)
        return 2
    # 文件锁前置检查（终审人 Word 可能正开着）
    lock = os.path.join(os.path.dirname(path), '~$' + os.path.basename(path))
    if os.path.exists(lock):
        print('!! 检测到 Word 锁文件 %s —— 文件可能正被打开，请先关闭 Word' % os.path.basename(lock))
        return 2

    word = _open_word()
    doc = None
    try:
        doc = word.Documents.Open(path, False, False)

        def stat():
            return dict(pages=doc.ComputeStatistics(WD_STAT_PAGES),
                        words=doc.ComputeStatistics(WD_STAT_WORDS),
                        paras=doc.Paragraphs.Count,
                        tables=doc.Tables.Count,
                        sections=doc.Sections.Count,
                        comments=doc.Comments.Count,
                        revisions=doc.Revisions.Count,
                        tocs=doc.TablesOfContents.Count)

        st = stat()
        print('打开: %s' % os.path.basename(path))
        print('  BEFORE  页=%s 字=%s 段=%s 表=%s 批注=%s 修订=%s 目录=%s'
              % (st['pages'], st['words'], st['paras'], st['tables'],
                 st['comments'], st['revisions'], st['tocs']))

        if verify_only:
            return 0

        # ---- 关键顺序：先关修订，再动任何内容 ----
        if doc.TrackRevisions:
            print('  · TrackRevisions 原为 ON（跨会话继承），已置 False')
        doc.TrackRevisions = False

        if accept and doc.Revisions.Count > 0:
            n = doc.Revisions.Count
            doc.Revisions.AcceptAll()
            print('  · 已接受全部修订 %d 条（转为正式文本）' % n)

        if update_toc and doc.TablesOfContents.Count:
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
            doc.Fields.Update()
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
            print('  · 目录/域已刷新')

        # 关修订后再确认一次（Fields.Update 可能重新打开修订开关）
        doc.TrackRevisions = False
        st2 = stat()
        if st2['revisions'] > st['revisions']:
            print('  ! 警告：刷新过程新增修订 %d 条，立即清除' % (st2['revisions'] - st['revisions']))
            doc.Revisions.AcceptAll()
            st2 = stat()

        doc.Save()
        print('  SAVED')

        if pdf:
            doc.SaveAs2(os.path.abspath(pdf), WD_FORMAT_PDF)
            print('  PDF -> %s' % pdf)

        st3 = stat()
        print('  AFTER   页=%s 字=%s 段=%s 表=%s 批注=%s 修订=%s'
              % (st3['pages'], st3['words'], st3['paras'], st3['tables'],
                 st3['comments'], st3['revisions']))
        if st3['revisions'] != 0:
            print('  !! 仍有修订残留，请检查')
            return 2
        return 0
    except Exception as ex:
        print('!! Word 操作失败: %s' % ex)
        return 2
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass


def main():
    ap = argparse.ArgumentParser(description='尽调报告 Word 定稿器')
    ap.add_argument('docx')
    ap.add_argument('--pdf', default=None, help='同时导出 PDF 路径')
    ap.add_argument('--accept', action='store_true', help='接受全部修订（转正式文本）')
    ap.add_argument('--no-toc', action='store_true', help='不刷新目录')
    ap.add_argument('--verify-only', action='store_true', help='仅体检，不改文件')
    a = ap.parse_args()
    sys.exit(finalize(a.docx, a.pdf, a.accept, not a.no_toc, a.verify_only))


if __name__ == '__main__':
    main()
