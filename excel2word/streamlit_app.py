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
import base64

DATE_FMT = '%Y-%m-%d'

# ---------------- 下面全部是你原来的函数，原封不动 ----------------
def is_table_row(row):
    if is_empty_row(row):
        return False
    return sum(1 for c in row if c is not None) >= 2

def is_empty_row(row):
    """
    整行全是 None 或空格，视为空行
    """
    return all(str(v or '').strip() == '' for v in row)

def set_table_borders(tbl, thick=12, dash=6):   #封装函数
    rows = tbl.rows                             #rows:取行
    if not rows:
        return

    # ---------- 横向 ----------
    # ① 首行：只加 top 粗线
    for cell in rows[0].cells:                  # 取单元格
        tc_pr = cell._tc.get_or_add_tcPr()      # 打开属性
        tc_borders = tc_pr.first_child_found_in('w:tcBorders') #找第一个子节点
        if tc_borders is None:                  #子节点是空的
            tc_borders = OxmlElement('w:tcBorders')  #增加节点
            tc_pr.append(tc_borders)            #挂到父节点下面
        top = OxmlElement('w:top')              #创建一个标签节点
        top.set(qn('w:val'), 'single')          #设置w:val为single
        top.set(qn('w:sz'), str(thick))
        top.set(qn('w:color'), '000000')
        tc_borders.append(top)                  #应用这个设置

        # 追加 bottom 虚线
        btm = OxmlElement('w:bottom')
        btm.set(qn('w:val'), 'dotted')
        btm.set(qn('w:sz'), str(dash))
        btm.set(qn('w:color'), '000000')
        tc_borders.append(btm)
        
    # ② 中间行：只加 bottom 虚线
    for row in rows[1:-1]:   #跳过第一行和最后一行
        for cell in row.cells:    #取行里的单元格
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

    # ③ 末行：只加 bottom 粗线
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

    # ---------- 竖向 ----------
    # ④ 列间虚竖线：除了最右列，其余每列都画 right 虚线
    for row in rows:
        for idx, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_borders = tc_pr.first_child_found_in('w:tcBorders')
            if tc_borders is None:
                tc_borders = OxmlElement('w:tcBorders')
                tc_pr.append(tc_borders)

            # 不是最右列 → 画 right 虚线（列间线）
            if idx != len(row.cells) - 1:
                right = OxmlElement('w:right')
                right.set(qn('w:val'), 'dotted')
                right.set(qn('w:sz'), str(dash))
                right.set(qn('w:color'), '000000')
                tc_borders.append(right)
            # 最左/最左外线：不画 left，保持空

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
    pfmt.space_after  = Pt(5)          # 关键：强制 0 磅
    pfmt.line_spacing_rule = 1         # 固定值
    pfmt.line_spacing = Pt(12)

def add_formatted_paragraph(doc, text,
                            before=6,   # 段前，单位磅
                            after=6,    # 段后，单位磅
                            line_spacing=Pt(18),   # 固定值18磅，可改
                            align=WD_ALIGN_PARAGRAPH.LEFT):
    """
    在 doc 末尾新增一个段落，并统一设置段前/段后/行距
    """
    p = doc.add_paragraph(text)
    fmt = p.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after  = Pt(after)
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY  # 固定值
    fmt.line_spacing = line_spacing
    fmt.alignment = align
    return p

def strip_trailing_nulls(row: List) -> Tuple[int, List]:
    """
    去掉行尾连续的 None 或空字符串
    返回 (有效列数, 去尾后的新列表)
    """
    # 统一转 str，方便判断
    tmp = [str(v) if v is not None else '' for v in row]
    # 从右往左找第一个非空
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

# ---------------- 上面全部是你原来的函数，原封不动 ----------------

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
                else:
                    p = cell.paragraphs[0]
                    p.text = txt
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(tbl.rows)==1 or j==0 else WD_ALIGN_PARAGRAPH.LEFT
                set_cell_vertical_center(cell)
        else:
            txt = ' '.join(str(v) if v is not None else '' for v in row).strip()
            if txt: add_formatted_paragraph(doc, txt)
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
import base64

DATE_FMT = '%Y-%m-%d'

# ---------------- 下面全部是你原来的函数，原封不动 ----------------
def is_empty_row(row):...
def is_table_row(row):...
def set_table_borders(tbl, thick=12, dash=6):...
def is_number(s: str) -> bool:...
def set_cell_vertical_center(cell):...
def add_formatted_paragraph(doc, text, before=6, after=6, line_spacing=Pt(18),
                            align=WD_ALIGN_PARAGRAPH.LEFT):...
def strip_trailing_nulls(row: List) -> Tuple[int, List]:...
def fmt_date(v):...
# ---------------- 上面全部是你原来的函数，原封不动 ----------------

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
                else:
                    p = cell.paragraphs[0]
                    p.text = txt
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(tbl.rows)==1 or j==0 else WD_ALIGN_PARAGRAPH.LEFT
                set_cell_vertical_center(cell)
        else:
            txt = ' '.join(str(v) if v is not None else '' for v in row).strip()
            if txt: add_formatted_paragraph(doc, txt)
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
