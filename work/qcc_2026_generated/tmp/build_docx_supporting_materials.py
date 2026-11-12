from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

ROOT = Path("/Users/USER/Documents/Codex/2026-06-21/new-chat-3")
OUT = ROOT / "outputs/qcc_2026_generated"
OUT.mkdir(parents=True, exist_ok=True)

BLUE = "1F4D78"
NAVY = "0A2742"
HEADER = "E8EEF5"
BORDER = "B9C7D8"
GREEN = "177A51"
RED = "B9403A"

daily = [
    ("5月09日", 151, 103, "0.45%", "上机初期；CP5 12线POY"),
    ("5月10日", 143, 82, "0.45%", "上机初期"),
    ("5月11日", 180, 64, "0.45%", "调整油轮导丝器"),
    ("5月12日", 142, 82, "0.45%", "换至8线生产，仍未稳定"),
    ("5月13日", 190, 64, "0.45%", "调整油轮导丝器"),
    ("5月14日", 141, 89, "0.45%", "调整油轮导丝器"),
    ("5月15日", 243, 114, "0.45%", "张力CV波动明显"),
    ("5月16日", 180, 71, "0.45%", "竹本油剂试验，无明显效果"),
    ("5月17日", 130, 50, "0.45%", "仍高位波动"),
    ("5月18日", 143, 54, "0.45%", "仍高位波动"),
    ("5月19日", 112, 34, "0.55%", "前纺含油0.55%打样"),
    ("5月20日", 98, 21, "0.55%", "有效果但不明显"),
    ("5月21日", 119, 24, "0.55/0.60%", "0.60%打样10锭"),
    ("5月22日", 107, 21, "0.55%", "0.55%整体做到"),
    ("5月23日", 93, 19, "0.55%", "断头仍高于目标"),
    ("5月24日", 97, 10, "0.55%", "断头仍高于目标"),
    ("5月25日", 46, 1, "0.60%", "0.60%整体做到"),
    ("5月26日", 52, 3, "0.60%", "稳定"),
    ("5月27日", 50, 2, "0.60%", "稳定"),
    ("5月28日", 26, 6, "0.60%", "稳定"),
    ("5月29日", 35, 0, "0.60%", "稳定"),
    ("5月30日", 46, 9, "0.60%", "稳定"),
    ("5月31日", 40, 7, "0.60%", "稳定"),
    ("6月01日", 19, 1, "0.60%", "稳定"),
    ("6月02日", 23, 1, "0.60%", "稳定"),
    ("6月03日", 35, 0, "0.60%", "稳定"),
    ("6月04日", 28, 3, "0.60%", "稳定"),
    ("6月05日", 30, 0, "0.60%", "稳定"),
    ("6月06日", 32, 1, "0.60%", "稳定"),
    ("6月07日", 34, 1, "0.60%", "稳定"),
    ("6月08日", 29, 0, "0.60%", "稳定"),
    ("6月09日", 39, 2, "0.60%", "稳定"),
    ("6月10日", 45, 1, "0.60%", "稳定"),
    ("6月11日", 28, 0, "0.60%", "稳定"),
    ("6月12日", 34, 2, "0.60%", "稳定"),
    ("6月13日", 16, 1, "0.60%", "稳定"),
    ("6月14日", 16, 0, "0.60%", "稳定"),
]

target_breaks = 90
target_cuts = 50

factors = [
    ("人", "操作培训不到位", "查培训记录、现场操作一致性", "班组按工艺要求执行，异常主要随含油阶段变化，不随人员班次变化", "非要因", "继续培训与现场抽查"),
    ("人", "上机异常反馈不及时", "查异常发现、反馈、验证闭环记录", "前期0.55%试调有效果但不明显，整体切换至0.60%滞后至5月25日", "要因", "建立前纺-加弹联动验证机制"),
    ("机", "DTY10#机油轮导丝器异常", "5月11-14日停车半台调整后跟踪", "5月15-18日断头仍为243/180/130/143，C切仍有114/71/50/54", "非要因", "作为点检项目，不作为主因"),
    ("机", "张力传感/报警阈值异常", "查看张力CV波动及报警记录", "CV波动是生产不稳定的表现，未见阈值调整带来持续改善", "非要因", "保留张力趋势监控"),
    ("料", "POY原丝含油率偏低", "按0.45%、0.55%、0.60%分阶段对比", "0.45%阶段断头日均164.3、C切77.3；0.60%阶段断头33.5、C切2.0", "要因", "POY含油控制到0.60%"),
    ("料", "前纺线别切换影响", "对比CP5 12线及5月12日后8线生产", "5月12日换8线后仍连续多日高位波动，直至0.60%整体执行才稳定", "非要因", "线别切换作为记录项"),
    ("料", "竹本油剂试验影响", "5月16日换竹本油剂后跟踪", "5月16-18日断头180/130/143，C切71/50/54，未形成明显改善", "非要因", "不作为本课题主因"),
    ("法", "55/72黑丝POY含油控制标准未固化", "查上机初期与试调过程", "0.55%阶段C切下降但断头仍高于目标，0.60%整体后连续21天达标", "要因", "固化55/72黑丝POY含油目标"),
    ("法", "DTY端局部参数调整不足", "核对油轮导丝器、二罗拉、K值、冷却板等临时调整", "DTY端局部调整未改变高断头趋势，改善来自POY含油整体执行", "非要因", "避免把临时试调写成有效对策"),
    ("法", "检验/分级标准不统一", "核对系统记录、日报、趋势图", "断头与C切数据连续记录，阶段差异清晰，非分级口径造成", "非要因", "保留同一口径日报"),
    ("环", "环境温湿度波动", "对比同期间环境与异常走势", "异常改善节点与0.60%整体执行一致，未见环境变化解释阶段改善", "非要因", "保持车间环境巡检"),
    ("环", "上机初期现场状态波动", "观察上机初期与稳定生产期现场情况", "上机初期波动无法解释0.55%与0.60%阶段差异", "非要因", "纳入上机确认清单"),
]

measures = [
    ("POY原丝含油率偏低", "POY含油率目标由0.45/0.55调整并固化至0.60%", "断头≤90个/日，C切≤50个/日", "5月21日打样10锭，5月25日整体执行；每日跟踪断头、C切", "0.60%阶段21天断头日均33.5、C切日均2.0", "有效"),
    ("55/72黑丝POY含油控制标准未固化", "建立55/72黑丝专用含油控制标准", "上机后不再以0.55%作为最终稳定标准", "首批上机确认含油目标、批量切换时间、异常复核节点", "0.55%阶段断头日均104.3仍高于目标；0.60%整体后全部达标", "有效"),
    ("前纺-加弹异常联动验证滞后", "建立异常联动验证与闭环表", "试调结论当天记录，整体调整有责任人与期限", "将打样、线别、油剂、含油率、断头/C切结果纳入一表跟踪", "非要因项被排除，减少DTY端重复试错", "有效"),
]

standards = [
    ("POY含油率", "55/72黑丝POY含油率按0.60%目标控制；低于目标需暂停扩大生产并复核", "每批/每日", "前纺工艺", "POY含油确认表"),
    ("上机确认", "记录POY来源线别、油剂、含油率、打样锭位、批量切换时间", "每次上机/换批", "加弹工艺", "上机确认单"),
    ("过程监控", "每日跟踪10#机断头≤90个、C切≤50个；任一超标当天反馈", "每日", "班组/QC小组", "每日目标跟踪表"),
    ("异常联动", "前纺、加弹、质检三方确认异常项，不把线别/换油剂/DTY临时调整直接判为主因", "异常发生时", "QC组长", "异常验证闭环表"),
    ("防复发", "连续3天断头或C切接近目标上限时，提前复核POY含油与张力CV趋势", "触发式", "工艺组", "防复发检查表"),
]


def avg(rows, idx):
    return round(sum(r[idx] for r in rows) / len(rows), 1)


early = daily[:16]
post = daily[16:]
stage_rows = [
    ("上机初期/含油率0.45%", "5月09日-5月18日", 10, 164.3, 77.3, "断头、C切均高位波动"),
    ("含油率0.55%试调", "5月19日-5月24日", 6, 104.3, 21.5, "C切下降，断头仍高于目标"),
    ("含油率0.60%整体执行", "5月25日-6月14日", 21, 33.5, 2.0, "断头、C切均稳定达标"),
]

trend_path = OUT / "近一个月断头_C切趋势图.png"


def set_doc(doc, landscape=False):
    sec = doc.sections[0]
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width = Inches(11)
        sec.page_height = Inches(8.5)
        sec.left_margin = Inches(0.55)
        sec.right_margin = Inches(0.55)
        sec.top_margin = Inches(0.55)
        sec.bottom_margin = Inches(0.55)
    else:
        sec.left_margin = Inches(0.8)
        sec.right_margin = Inches(0.8)
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
    styles = doc.styles
    styles["Normal"].font.name = "Calibri"
    styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    styles["Normal"].font.size = Pt(12)


def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell(cell, text, bold=False, fill=None, color=None, size=9, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    if align:
        p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if fill:
        shade(cell, fill)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def set_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), BORDER)


def set_widths(table, widths):
    table.autofit = False
    for row in table.rows:
        for i, width in enumerate(widths):
            row.cells[i].width = Inches(width)


def add_title(doc, title, subtitle):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(title)
    r.bold = True
    r.font.name = "Calibri"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    r.font.size = Pt(20)
    r.font.color.rgb = RGBColor.from_string(BLUE)
    sp = doc.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp.paragraph_format.space_after = Pt(10)
    sr = sp.add_run(subtitle)
    sr.font.name = "Calibri"
    sr._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    sr.font.size = Pt(10.5)
    sr.font.color.rgb = RGBColor.from_string("555555")


def heading(doc, text, page_break_before=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    if page_break_before:
        p.paragraph_format.page_break_before = True
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Calibri"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor.from_string(BLUE)


def body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.12
    r = p.add_run(text)
    r.font.name = "Calibri"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    r.font.size = Pt(12)


def add_table(doc, headers, rows, widths=None, font_size=10):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    if widths:
        set_widths(table, widths)
    for i, h in enumerate(headers):
        set_cell(table.rows[0].cells[i], h, bold=True, fill=BLUE, color="FFFFFF", size=font_size, align=WD_ALIGN_PARAGRAPH.CENTER)
    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            set_cell(cells[i], val, size=font_size, align=WD_ALIGN_PARAGRAPH.CENTER if i < 3 else None)
    set_borders(table)
    doc.add_paragraph()
    return table


def save_doc(doc, filename):
    path = OUT / filename
    doc.save(path)
    return path


def build_factor_doc():
    doc = Document()
    set_doc(doc, landscape=True)
    add_title(doc, "55/72黑丝末端因素验证数据", "简洁版｜课题：提高55/72黑丝优等率｜要因3项，非要因按人、机、料、法、环归档")
    heading(doc, "一、分阶段验证数据")
    add_table(doc, ["阶段", "日期范围", "天数", "断头日均", "C切日均", "结论"], stage_rows, [1.75, 1.25, 0.55, 0.9, 0.9, 3.95], 10)
    heading(doc, "二、要因确认表")
    cause_rows = [
        ("要因1", "POY原丝含油率偏低", "0.45%阶段断头164.3、C切77.3；0.60%后断头33.5、C切2.0。", "判定为要因，按0.60%目标控制。"),
        ("要因2", "55/72黑丝POY含油控制标准未固化", "0.55%试调后C切下降，但断头仍高于目标。", "判定为要因，固化专用含油标准。"),
        ("要因3", "前纺-加弹异常联动验证滞后", "打样、整体切换、每日反馈没有形成固定闭环。", "判定为要因，建立异常验证闭环。"),
    ]
    add_table(doc, ["序号", "末端因素", "核心验证过程", "结论/处置"], cause_rows, [0.75, 2.2, 3.55, 3.0], 10)
    heading(doc, "三、非要因确认表")
    non_rows = [
        ("人", "操作培训不到位", "异常随含油阶段变化，不随班次变化。", "非要因，保留培训抽查。"),
        ("机", "DTY端油轮导丝器/局部参数", "5月11-14日调整后仍高位波动。", "非要因，不写成有效对策。"),
        ("料", "CP5 12线/8线切换、竹本油剂试验", "换线、换油剂后未形成持续改善。", "非要因，作为排除证据。"),
        ("法", "检验/分级口径、CV报警阈值", "数据记录连续，CV波动是异常表现。", "非要因，保留监控。"),
        ("环", "环境温湿度及上机现场状态", "改善节点与0.60%整体执行一致。", "非要因，保持巡检。"),
    ]
    add_table(doc, ["维度", "非要因项", "验证过程", "结论"], non_rows, [0.65, 2.3, 3.55, 2.55], 10)
    heading(doc, "四、人、机、料、法、环总结", page_break_before=True)
    summary = [
        ("人", 1, 1, "要因：异常反馈闭环滞后；非要因：操作培训不到位。"),
        ("机", 0, 2, "非要因：油轮导丝器/局部参数、张力报警阈值。"),
        ("料", 1, 2, "要因：POY原丝含油率偏低；非要因：线别切换、竹本油剂试验。"),
        ("法", 1, 2, "要因：55/72黑丝含油控制标准未固化；非要因：分级口径、DTY端临调。"),
        ("环", 0, 2, "非要因：温湿度、上机现场状态。"),
    ]
    add_table(doc, ["维度", "要因数", "非要因数", "结论摘要"], summary, [0.65, 0.8, 0.9, 7.25], 10)
    body(doc, "确认结论：本课题确认3个要因，分别为POY原丝含油率偏低、55/72黑丝POY含油控制标准未固化、前纺-加弹异常联动验证滞后。CP5 12线/8线切换、5月16日竹本油剂试验、DTY端临时调整等均作为非要因归档。")
    return save_doc(doc, "末端因素验证数据_55-72黑丝.docx")


def build_measure_doc():
    doc = Document()
    set_doc(doc, landscape=True)
    add_title(doc, "55/72黑丝措施有效性验证数据", "简洁版｜措施围绕3个要因展开，效果以断头/C切日均和连续达标天数验证")
    heading(doc, "一、措施有效性验证")
    add_table(doc, ["对应要因", "对策", "目标", "具体措施", "验证数据", "判定"], measures, [1.45, 1.85, 1.35, 2.05, 2.35, 0.55], 9.6)
    heading(doc, "二、改善前后效果对比")
    rows = [
        ("断头数", "改善前/试调期日均141.8个", "改善后日均33.5个", "下降108.3个", "下降76.4%", "低于90个/日目标"),
        ("C切数", "改善前/试调期日均56.4个", "改善后日均2.0个", "下降54.4个", "下降96.5%", "低于50个/日目标"),
        ("连续达标", "5月9日-5月24日未稳定", "5月25日-6月14日", "连续21天", "100%达标", "措施有效"),
    ]
    add_table(doc, ["项目", "改善前", "改善后", "改善量", "改善率", "结论"], rows, [1.0, 2.0, 1.75, 1.25, 1.1, 2.5], 10)
    heading(doc, "三、有效性结论")
    body(doc, "0.60%整体执行后，断头与C切同步进入稳定区间，说明围绕POY含油控制标准和异常联动机制制定的措施有效。DTY端局部参数调整、线别切换、竹本油剂试验未形成持续改善，不作为最终有效对策。")
    return save_doc(doc, "措施有效性验证数据_55-72黑丝.docx")


def make_trend_chart():
    rows = daily[-30:]
    labels = [r[0].replace("月", "/").replace("日", "") for r in rows]
    breaks = [r[1] for r in rows]
    cuts = [r[2] for r in rows]
    plt.rcParams["font.sans-serif"] = ["Arial Unicode MS", "PingFang SC", "Heiti TC", "Microsoft YaHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    fig, ax = plt.subplots(figsize=(11.5, 4.2), dpi=180)
    ax.plot(labels, breaks, marker="o", linewidth=2.2, color="#B9403A", label="断头数")
    ax.plot(labels, cuts, marker="o", linewidth=2.2, color="#1E6AA8", label="C切数")
    ax.axhline(target_breaks, color="#B9403A", linestyle="--", linewidth=1.2, alpha=0.65, label="断头目标90")
    ax.axhline(target_cuts, color="#1E6AA8", linestyle="--", linewidth=1.2, alpha=0.65, label="C切目标50")
    ax.axvline(labels[9], color="#177A51", linestyle=":", linewidth=1.6)
    ax.text(labels[9], max(max(breaks), max(cuts)) * 0.92, "0.60%整体执行", color="#177A51", fontsize=11, ha="left")
    ax.set_title("近一个月断头/C切趋势", fontsize=16, fontweight="bold", color="#0A2742", pad=12)
    ax.set_ylabel("数量（个/日）", fontsize=11)
    ax.grid(axis="y", alpha=0.25)
    ax.legend(loc="upper right", ncol=4, fontsize=10, frameon=False)
    ax.set_xticks(range(0, len(labels), 3))
    ax.set_xticklabels([labels[i] for i in range(0, len(labels), 3)], rotation=0, fontsize=9)
    ax.tick_params(axis="y", labelsize=10)
    fig.tight_layout()
    fig.savefig(trend_path, bbox_inches="tight")
    plt.close(fig)
    return trend_path


def build_daily_doc():
    doc = Document()
    set_doc(doc, landscape=True)
    add_title(doc, "近一个月每日目标值跟踪情况", "简洁版｜统计区间：2026年5月16日-2026年6月14日｜目标：断头≤90个/日，C切≤50个/日")
    heading(doc, "一、趋势图")
    chart = make_trend_chart()
    doc.add_picture(str(chart), width=Inches(9.55))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    heading(doc, "二、每日目标跟踪明细", page_break_before=True)
    rows = []
    for date, br, cut, oil, note in daily[-30:]:
        bpass = "达成" if br <= target_breaks else "未达成"
        cpass = "达成" if cut <= target_cuts else "未达成"
        allpass = "达成" if bpass == "达成" and cpass == "达成" else "未达成"
        stage = "对策后" if oil == "0.60%" else "对策前/试调"
        rows.append((date, oil, br, bpass, cut, cpass, allpass, stage if note == "稳定" else note))
    add_table(doc, ["日期", "含油率", "断头", "断头达成", "C切", "C切达成", "综合", "备注"], rows, [0.82, 0.72, 0.55, 0.85, 0.5, 0.85, 0.65, 4.05], 8.8)
    heading(doc, "三、达成情况汇总")
    total_rows = daily[-30:]
    post_rows = [r for r in total_rows if r[3] == "0.60%"]
    before_rows = [r for r in total_rows if r[3] != "0.60%"]
    summary = [
        ("近一个月整体", len(total_rows), sum(1 for r in total_rows if r[1] <= target_breaks and r[2] <= target_cuts), "含对策前/试调和对策后"),
        ("对策前/试调", len(before_rows), sum(1 for r in before_rows if r[1] <= target_breaks and r[2] <= target_cuts), "5月16日至5月24日，整体未稳定"),
        ("0.60%整体执行后", len(post_rows), sum(1 for r in post_rows if r[1] <= target_breaks and r[2] <= target_cuts), "5月25日至6月14日，连续21天达标"),
    ]
    add_table(doc, ["统计项", "天数", "达标天数", "说明"], summary, [1.6, 0.8, 0.9, 6.2], 10)
    return save_doc(doc, "近一个月每日目标值跟踪情况_55-72黑丝.docx")


def build_standard_doc():
    doc = Document()
    set_doc(doc, landscape=True)
    add_title(doc, "55/72黑丝标准化及防措", "简洁版｜课题：提高55/72黑丝优等率｜适用：10#机55/72黑丝")
    heading(doc, "一、标准化目的")
    body(doc, "将本次QCC活动中验证有效的POY含油控制、上机确认、前纺-加弹联动跟踪和异常防复发要求固化，防止55/72黑丝再次出现断头多、C切多、优等率偏低的问题。")
    heading(doc, "二、标准化项目")
    add_table(doc, ["控制项目", "标准/防措", "频次", "责任人", "记录"], standards, [1.1, 4.1, 1.0, 1.0, 2.3], 10)
    heading(doc, "三、防复发触发条件")
    add_table(doc, ["触发条件", "复核动作", "输出"], [
        ("连续3天断头或C切接近目标上限", "提前复核POY含油与张力CV趋势", "防复发检查表"),
        ("换线、换批、换油剂后指标波动", "重新核对上机确认信息，不优先扩大DTY端临时试调", "上机确认单"),
        ("复盘或评审材料编制", "非要因项保留证据，不写成有效对策", "末端因素确认表"),
        ("新批次55/72黑丝上机", "0.55%仅作为过渡验证，0.60%作为稳定生产标准", "含油率确认记录"),
    ], [2.4, 5.0, 2.1], 10)
    heading(doc, "四、非要因归档口径")
    body(doc, "5月9日上机时使用前纺CP5 12线POY，5月12日后换到8线生产；5月16日进行竹本油剂试验；5月11日至5月14日有DTY端油轮导丝器调整。上述事项均作为验证和排除证据保留，不作为本课题主要要因，也不写成最终有效对策。")
    return save_doc(doc, "标准化及防措_55-72黑丝.docx")


paths = [
    build_factor_doc(),
    build_measure_doc(),
    build_daily_doc(),
    build_standard_doc(),
]

for p in paths:
    print(p)
