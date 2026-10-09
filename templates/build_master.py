# -*- coding: utf-8 -*-
# dd_report_master.docx 格式母版生成器
# 母版 = 8类格式样本段的载体（cover/blank/h1/h2/br/note/unit/p），
# dd_report_builder.py 从母版克隆段落格式（pPr），清空正文后按 content.json 重建。
# 本母版不含任何真实客户信息；按你的公司格式定稿报告定制时，保留段落结构、改字体字号即可。
import copy
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

EASTASIA = '宋体'
ASCII_FONT = 'Times New Roman'

doc = Document()
# A4 + 常规页边距
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(2.54)
sec.left_margin = sec.right_margin = Cm(3.17)
# Normal 样式：宋体/TNR 小四、1.5倍行距
normal = doc.styles['Normal']
normal.font.name = ASCII_FONT
normal.font.size = Pt(12)
normal._element.rPr.rFonts.set(qn('w:eastAsia'), EASTASIA)
pf = normal.paragraph_format
pf.line_spacing = 1.5
pf.space_before = Pt(0)
pf.space_after = Pt(0)

def add_para(text='', align=None, bold=None, size=None, first_indent_chars=None,
             style=None, outline=None):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    if first_indent_chars:
        pPr = p._element.get_or_add_pPr()
        ind = OxmlElement('w:ind')
        ind.set(qn('w:firstLine'), str(first_indent_chars * 240))
        ind.set(qn('w:firstLineChars'), str(first_indent_chars * 100))
        pPr.append(ind)
    if outline is not None:
        pPr = p._element.get_or_add_pPr()
        ol = OxmlElement('w:outlineLvl')
        ol.set(qn('w:val'), str(outline))
        pPr.append(ol)
    if text:
        r = p.add_run(text)
        rPr = r._element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:ascii'), ASCII_FONT)
        rFonts.set(qn('w:hAnsi'), ASCII_FONT)
        rFonts.set(qn('w:eastAsia'), EASTASIA)
        if size is not None:
            r.font.size = Pt(size)
        if bold is not None:
            r.font.bold = bold
    return p

J, C, R = WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT

# ── 8类格式样本（builder capture_exemplars 依此识别）──
# cover：居中+加粗
add_para('示例公司尽调报告格式母版', align=C, bold=True, size=18)
add_para('（本母版仅承载格式样本，不含真实信息；可按贵司定稿报告字体字号定制）', align=C, size=10.5)
# blank：空段
add_para('')
# h1：标题1样式
add_para('一、格式样本章节', style='Heading 1')
# h2：（一）开头、两端对齐+首行缩进
add_para('（一）节标题样本：本段以（一）开头即被识别为h2格式样本', align=J, first_indent_chars=2)
# p：>30字正文段、首行缩进、非居中
add_para('这是正文格式样本段。尽调报告的正文段落使用两端对齐、首行缩进两字符、'
         '宋体小四、1.5倍行距，数字与西文使用Times New Roman字体。'
         '构建器会克隆本段的段落属性作为全部正文段落的格式来源。', align=J, first_indent_chars=2)
# br：【】开头的意见区条目段
add_para('【项目需求】意见区条目格式样本：本段以【开头，被识别为意见区（br）格式样本。',
         align=J, first_indent_chars=2)
# note：备注开头
add_para('备注：备注段格式样本——数据来源标注、口径说明使用本格式（五号字）。', align=J, first_indent_chars=2)
# unit：单位开头、右对齐
add_para('单位：万元', align=R, first_indent_chars=2)
# 再补一段正文保证p样本稳定
add_para('补充正文样本段：格式母版法的原理是以一份已定稿的尽调报告为格式基线，'
         '构建器自动识别其中八类段落的格式样本，然后清空正文、按结构化内容重建整份报告，'
         '从而百分之百继承既有的字体、缩进、行距与表格样式。', align=J, first_indent_chars=2)

doc.save(r'D:\Claude\projects\claude-dd-report\templates\dd_report_master.docx')
print('母版已生成: templates/dd_report_master.docx')
