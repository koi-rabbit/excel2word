import streamlit as st
from pathlib import Path
import openpyxl
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.shared import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt
from typing import List, Tuple
import re
from datetime import datetime, timedelta
import io

DATE_FMT = '%Y-%m-%d'

def is_empty_row(row):
    """
    整行全是 None 或空格，视为空行
    """
    return all(str(v or '').strip() == '' for v in row)

def is_table_row(row):
    if is_empty_row(row):
        return False
    return sum(1 for c in row if c is not None) >= 2

def set_table_borders(tbl, thick=12, dash=6):
    rows = tbl.rows
    if not rows:
        return

    for cell in rows[0].cells:
        tc_pr = cell._tc.get_or_add_tcPr()
        tc_borders = tc_pr.first_child_found_in('w:tcBorders')
        if tc_borders is None:
            tc_borders = OxmlElement('w:tcBorders')
            tc_pr.append(tc_borders)
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'single')
        top.set(qn('w:sz'), str(thick))
        top.set(qn('w:color'), '000000')
        tc_borders.append(top)

        btm = OxmlElement('w:bottom')
        btm.set(qn('w:val'), 'dotted')
        btm.set(qn('w:sz'), str(dash))
        btm.set(qn('w:color'), '000000')
        tc_borders.append(btm)

    for row in rows[1:-1]:
        for cell in row.cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_borders = tc_pr.first_child_found_in('w:tcBorders')
            if tc_borders is None:
                tc_borders = OxmlElement('w:tcBorders')
                tc_pr.append(tc_borders)
            btm = OxmlElement('w:bottom')
            btm.set(qn('w:val'), 'dotted')
            btm.set(qn('w:sz'), str(dash))
            btm.set(qn('w:color'), '000000')
            tc_borders.append(btm)

    for cell in rows[-1].cells:
        tc_pr = cell._tc.get_or_add_tcPr()
        tc_borders = tc_pr.first_child_found_in('w:tcBorders')
        if tc_borders is None:
            tc_borders = OxmlElement('w:tcBorders')
            tc_pr.append(tc_borders)
        btm = OxmlElement('w:bottom')
        btm.set(qn('w:val'), 'single')
        btm.set(qn('w:sz'), str(thick))
        btm.set(qn('w:color'), '000000')
        tc_borders.append(btm)

    for row in rows:
        for idx, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_borders = tc_pr.first_child_found_in('w:tcBorders')
            if tc_borders is None:
                tc_borders = OxmlElement('w:tcBorders')
                tc_pr.append(tc_borders)

            if idx != len(row.cells) - 1:
                right = OxmlElement('w:right')
                right.set(qn('w:val'), 'dotted')
                right.set(qn('w:sz'), str(dash))
                right.set(qn('w:color'), '000000')
                tc_borders.append(right)

def is_number(s: str) -> bool:
    """纯数字（可带小数点）返回 True"""
    try:
        float(s.replace(",", ""))
        return "." not in s or s.count(".") == 1
    except ValueError:
        return False

def set_cell_vertical_center(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    tcVAlign = OxmlElement('w:vAlign')
    tcVAlign.set(qn('w:val'), 'center')
    tc_pr.append(tcVAlign)

    p = cell.paragraphs[0]
    pfmt = p.paragraph_format
    pfmt.space_before = Pt(5)
    pfmt.space_after = Pt(5)
    pfmt.line_spacing_rule = 1
    pfmt.line_spacing = Pt(12)

def add_formatted_paragraph(doc, text, font_name='宋体', font_size=11):
    """
    在 doc 末尾新增一个段落，并统一设置段前/段后/行距
    """
    p = doc.add_paragraph(text)
    fmt = p.paragraph_format
    fmt.space_before = Pt(6)
    fmt.space_after = Pt(6)
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(18)
    fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT

    for run in p.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)

def strip_trailing_nulls(row: List) -> Tuple[int, List]:
    """
    去掉行尾连续的 None 或空字符串
    返回 (有效列数, 去尾后的新列表)
    """
    tmp = [str(v) if v is not None else '' for v in row]
    i = len(tmp)
    while i > 0 and tmp[i - 1].strip() == '':
        i -= 1
    return i, tmp[:i]

def fmt_date(v) -> str:
    """
    把 openpyxl 的日期序列数字 -> 指定格式字符串
    如果不是日期，原样返回
    """
    if isinstance(v, datetime):
        return v.strftime(DATE_FMT)
    return str(v) if v is not None else ""

def excel_to_docx_bytes(ws):
    """把单个工作表转成 Word 文件，返回 BytesIO"""
    doc = Document()
    in_table, tbl = False, None
    for row in ws.iter_rows(values_only=True):
        if is_empty_row(row):
            if in_table:
                set_table_borders(tbl); in_table=False; tbl=None
            doc.add_paragraph()
            continue
        if is_table_row(row):
            clean = [str(fmt_date(v)) if v is not None else "" for v in row]
            _, clean = strip_trailing_nulls(clean)
            if not in_table:
                tbl = doc.add_table(rows=0, cols=len(clean))
                in_table=True
            cells = tbl.add_row().cells
            for j, txt in enumerate(clean):
                cell = cells[j]
                if is_number(txt):
                    p = cell.paragraphs[0]
                    p.text = f"{float(txt):,.2f}"
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(tbl.rows)==1 or j==0 else WD_ALIGN_PARAGRAPH.RIGHT
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(11)
                else:
                    p = cell.paragraphs[0]
                    p.text = txt
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(tbl.rows)==1 or j==0 else WD_ALIGN_PARAGRAPH.LEFT
                    for run in p.runs:
                        run.font.name = '宋体'
                        run.font.size = Pt(11)
                set_cell_vertical_center(cell)
        else:
            txt = ' '.join(str(v) if v is not None else '' for v in row).strip()
            if txt: 
                add_formatted_paragraph(doc, txt)
            if in_table:
                set_table_borders(tbl); in_table=False; tbl=None
    if in_table: set_table_borders(tbl)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# -------------------- Streamlit 页面 --------------------
st.set_page_config(page_title="Excel→Word 在线转换", layout="centered")
st.title("📄 Excel 转 Word 工具")
st.markdown("上传一个 `.xlsx` 文件，系统自动按你原来的规则生成 Word 表格并下载。")

uploaded = st.file_uploader("选择 Excel 文件", type=["xlsx"])
if uploaded:
    wb = openpyxl.load_workbook(uploaded, data_only=True)
    sheet = wb.worksheets[0]
    doc_io = excel_to_docx_bytes(sheet)
    st.success("转换完成！")
    st.download_button(
        label="⬇ 下载 Word",
        data=doc_io,
        file_name=f"{Path(uploaded.name).stem}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
