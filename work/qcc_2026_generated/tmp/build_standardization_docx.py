from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path("/Users/USER/Documents/Codex/2026-06-21/new-chat-3")
OUT = ROOT / "outputs/qcc_2026_generated"
OUT.mkdir(parents=True, exist_ok=True)
FINAL = OUT / "标准化及防措_55-72黑丝.docx"

BLUE = "1F4D78"
HEADER = "E8EEF5"
BORDER = "B9C7D8"


def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def cell_text(cell, text, bold=False, color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(10)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def set_borders(table):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "6")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), BORDER)


def set_col_widths(table, widths):
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = Inches(width)


def add_table(doc, headers, rows, widths):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_col_widths(table, widths)
    for i, h in enumerate(headers):
        cell_text(table.rows[0].cells[i], h, bold=True, color="FFFFFF")
        shade(table.rows[0].cells[i], BLUE)
    for row in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cell_text(cells[i], value)
    set_borders(table)
    doc.add_paragraph()
    return table


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(16 if level == 1 else 13)
    run.font.color.rgb = RGBColor.from_string(BLUE)
    return p


def add_body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.1
    run = p.add_run(text)
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(11)
    return p


doc = Document()
section = doc.sections[0]
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)
section.left_margin = Inches(1)
section.right_margin = Inches(1)

styles = doc.styles
styles["Normal"].font.name = "Calibri"
styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
styles["Normal"].font.size = Pt(11)

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_after = Pt(6)
run = title.add_run("55/72黑丝标准化及防措")
run.bold = True
run.font.name = "Calibri"
run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
run.font.size = Pt(20)
run.font.color.rgb = RGBColor.from_string(BLUE)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.paragraph_format.space_after = Pt(12)
run = sub.add_run("课题：提高55/72黑丝优等率｜小组：长空之王QC小组｜适用：10#机55/72黑丝")
run.font.name = "Calibri"
run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
run.font.size = Pt(10)

add_heading(doc, "一、标准化目的")
add_body(doc, "将本次QCC活动中验证有效的POY含油控制、上机确认、前纺-加弹联动跟踪和异常防复发要求固化，防止55/72黑丝再次出现断头多、C切多、优等率偏低的问题。")

add_heading(doc, "二、关键标准")
add_table(
    doc,
    ["控制项目", "标准化要求", "验证目标", "责任"],
    [
        ["POY含油率", "55/72黑丝POY含油率按0.60%目标控制；低于目标不得直接扩大生产。", "断头≤90个/日；C切≤50个/日", "前纺工艺"],
        ["上机确认", "记录POY来源线别、油剂、含油率、打样锭位、批量切换时间。", "信息完整，可追溯", "加弹工艺"],
        ["每日跟踪", "每日记录10#机断头、C切、含油阶段及异常说明。", "任一超标当天反馈", "班组/QC小组"],
        ["异常联动", "前纺、加弹、质检共同确认异常项，避免把线别切换、换油剂、DTY端临时调整直接判为主因。", "形成闭环结论", "QC组长"],
    ],
    [1.2, 3.2, 1.2, 0.9],
)

add_heading(doc, "三、防复发措施")
add_table(
    doc,
    ["序号", "防措", "触发条件", "输出记录"],
    [
        ["1", "连续复核POY含油率与张力CV趋势。", "连续3天断头或C切接近目标上限。", "防复发检查表"],
        ["2", "重新核对上机确认信息，不优先扩大DTY端临时试调。", "换线、换批、换油剂后指标波动。", "上机确认单"],
        ["3", "对非要因项保留证据，不在成果报告中写成有效对策。", "出现复盘或评审材料编制时。", "末端因素确认表"],
        ["4", "0.55%试调仅作为过渡验证，不作为最终稳定生产标准。", "新批次55/72黑丝上机。", "含油率确认记录"],
    ],
    [0.5, 2.5, 2.0, 1.5],
)

add_heading(doc, "四、非要因处置口径")
add_body(doc, "5月9日上机时使用前纺CP5 12线POY，5月12日后换到8线生产；5月16日进行竹本油剂试验；5月11日至5月14日有DTY端油轮导丝器调整。上述事项均作为验证和排除证据保留，不作为本课题主要要因，也不写成最终有效对策。")

add_heading(doc, "五、记录表单")
add_table(
    doc,
    ["表单名称", "用途", "保存要求"],
    [
        ["POY含油确认表", "记录55/72黑丝POY含油率目标、实测及批量切换时间。", "随批保存"],
        ["每日目标跟踪表", "记录每日断头、C切、目标达成和备注。", "每日更新"],
        ["异常验证闭环表", "记录异常假设、验证方法、数据、结论和责任人。", "QC项目归档"],
        ["防复发检查表", "记录接近目标上限时的提前复核动作。", "触发后保存"],
    ],
    [1.5, 3.3, 1.7],
)

add_heading(doc, "六、检查与更新")
add_body(doc, "本标准化及防措文件由QC小组负责维护。若后续55/72黑丝连续出现断头或C切超目标，应重新启动末端因素验证，并优先复核POY含油率及前纺-加弹联动执行情况。")

doc.save(FINAL)
print(FINAL)
