from datetime import date
from pathlib import Path
from warnings import filterwarnings

import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel
from xlsxwriter.utility import xl_range_abs


ROOT = Path("/Users/USER/Documents/Codex/2026-06-21/new-chat-3")
SOURCE = ROOT / "work/qc_2026_input/2026/断头趋势.xlsx"
OUTPUT = ROOT / "outputs/qcc_source_data_excel/55-72黑丝断头_C切趋势_目标修正版.xlsx"
TARGET_BREAKS = 50
TARGET_CCUT = 20
NODE_DATE_SERIAL = 46167
NODE_LABEL = "整体0.60%含油率"


def read_source():
    filterwarnings("ignore", message="Unable to read chart")
    wb = load_workbook(SOURCE, data_only=True)
    ws = wb["Sheet1"]
    values = []
    for r in range(1, 43):
        values.append([ws.cell(r, c).value for c in range(1, 4)])

    rows = []
    for r in range(6, 43):
        serial = ws.cell(r, 1).value
        day = from_excel(serial).date() if isinstance(serial, (int, float)) else serial
        rows.append(
            {
                "excel_row": r,
                "serial": serial,
                "date": day,
                "label": f"{day.month}/{day.day}",
                "breaks": ws.cell(r, 2).value,
                "ccut": ws.cell(r, 3).value,
            }
        )
    return values, rows


def stage_for(day):
    if day < date(2026, 5, 19):
        return "上机/异常期"
    if day < date(2026, 5, 25):
        return "0.55%试调期"
    return "0.60%执行期"


def write_formula(ws, row, col, formula, fmt=None, value=None):
    ws.write_formula(row, col, formula, fmt, value)


def add_trend_chart(workbook, ws, chart_sheet_name, chart_first_row, anchor, title, first_serial, last_serial):
    chart = workbook.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
    categories = f"='{chart_sheet_name}'!{xl_range_abs(chart_first_row, 0, chart_first_row + 29, 0)}"
    series_specs = [
        ("断头数", 1, "#B33A3A", "solid", 2.25, True),
        ("C切数", 2, "#1F6FAE", "solid", 2.25, True),
        ("目标断头数", 3, "#B99A9A", "dash", 1.25, False),
        ("目标C切数", 4, "#9BB7CC", "dash", 1.25, False),
    ]
    for name, col, color, dash, width, marker_enabled in series_specs:
        opts = {
            "name": name,
            "categories": categories,
            "values": f"='{chart_sheet_name}'!{xl_range_abs(chart_first_row, col, chart_first_row + 29, col)}",
            "line": {"color": color, "width": width},
        }
        if dash != "solid":
            opts["line"]["dash_type"] = dash
        if marker_enabled:
            opts["marker"] = {
                "type": "circle",
                "size": 4,
                "border": {"color": color},
                "fill": {"color": color},
            }
        else:
            opts["marker"] = {"type": "none"}
        chart.add_series(opts)

    chart.add_series(
        {
            "name": NODE_LABEL,
            "categories": f"='{chart_sheet_name}'!$P$5:$P$6",
            "values": f"='{chart_sheet_name}'!$Q$5:$Q$6",
            "line": {"color": "#00A651", "width": 1.5, "dash_type": "dash"},
            "marker": {"type": "none"},
        }
    )

    chart.set_title({"name": title, "name_font": {"size": 13, "bold": False}})
    chart.set_legend({"position": "top", "font": {"size": 8}})
    chart.set_x_axis(
        {
            "min": first_serial,
            "max": last_serial,
            "major_unit": 3,
            "num_format": "m/d",
            "num_font": {"size": 8},
            "major_gridlines": {"visible": True, "line": {"color": "#E5E7EB", "dash_type": "dash"}},
            "line": {"color": "#C9D2DA"},
        }
    )
    chart.set_y_axis(
        {
            "name": "数量（个/日）",
            "min": 0,
            "max": 200,
            "major_unit": 25,
            "num_font": {"size": 8},
            "name_font": {"size": 9},
            "major_gridlines": {"visible": True, "line": {"color": "#E5E7EB"}},
            "line": {"color": "#C9D2DA"},
        }
    )
    chart.set_chartarea({"border": {"color": "#7A7A7A", "width": 1}, "fill": {"color": "#FFFFFF"}})
    chart.set_plotarea({"border": {"color": "#D9D9D9"}, "fill": {"color": "#FFFFFF"}})
    chart.set_size({"width": 760, "height": 335})
    ws.insert_chart(anchor, chart)


def main():
    source_values, source_rows = read_source()
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    workbook = xlsxwriter.Workbook(OUTPUT)
    workbook.set_calc_mode("auto")

    title_fmt = workbook.add_format(
        {
            "bold": True,
            "font_size": 15,
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#0B3A5B",
        }
    )
    note_fmt = workbook.add_format(
        {
            "bold": True,
            "font_color": "#0B3A5B",
            "bg_color": "#EAF4FA",
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
    )
    section_fmt = workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#0B3A5B",
        }
    )
    header_fmt = workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#118AB2",
            "border": 1,
            "border_color": "#D6E4EC",
        }
    )
    cell_fmt = workbook.add_format({"border": 1, "border_color": "#D6E4EC", "valign": "vcenter"})
    int_fmt = workbook.add_format(
        {"border": 1, "border_color": "#D6E4EC", "valign": "vcenter", "num_format": "0"}
    )
    date_fmt = workbook.add_format(
        {"border": 1, "border_color": "#D6E4EC", "valign": "vcenter", "num_format": "m/d"}
    )
    band_fmt = workbook.add_format(
        {"border": 1, "border_color": "#D6E4EC", "valign": "vcenter", "bg_color": "#C7EAF6"}
    )
    band_date_fmt = workbook.add_format(
        {
            "border": 1,
            "border_color": "#D6E4EC",
            "valign": "vcenter",
            "bg_color": "#C7EAF6",
            "num_format": "m/d",
        }
    )
    band_int_fmt = workbook.add_format(
        {
            "border": 1,
            "border_color": "#D6E4EC",
            "valign": "vcenter",
            "bg_color": "#C7EAF6",
            "num_format": "0",
        }
    )
    audit_fmt = workbook.add_format(
        {
            "font_color": "#7A4B00",
            "bg_color": "#FFF7E6",
            "border": 1,
            "border_color": "#E4B860",
            "valign": "vcenter",
            "text_wrap": True,
        }
    )
    plain_fmt = workbook.add_format({"valign": "vcenter"})

    trend = workbook.add_worksheet("趋势图")
    chart_src = workbook.add_worksheet("PPT图表引用表")
    data = workbook.add_worksheet("趋势图数据")
    raw = workbook.add_worksheet("Sheet1")

    for ws in [trend, chart_src, data, raw]:
        ws.hide_gridlines(2)

    raw.set_column("A:C", 12)
    for r, row in enumerate(source_values):
        for c, value in enumerate(row):
            fmt = date_fmt if c == 0 and r >= 5 else plain_fmt
            raw.write(r, c, value, fmt)

    data.freeze_panes(5, 0)
    data.set_column("A:A", 10)
    data.set_column("B:E", 11)
    data.set_column("F:F", 14)
    data.set_column("G:G", 16)
    data.set_column("H:H", 8)
    data.merge_range("A1:H1", "55/72黑丝断头/C切完整原始引用数据", title_fmt)
    data.merge_range("B2:D2", "导入自原始工作簿：断头趋势.xlsx；原始范围：Sheet1!A6:C42", note_fmt)
    data.write("A2", "数据来源", header_fmt)
    data.merge_range("F2:H2", "本页 A6:C42 全部公式引用 Sheet1 原始导入数据。", note_fmt)
    data.write("E2", "说明", header_fmt)
    data.write_row("A4", ["目标断头数", TARGET_BREAKS, "目标C切数", TARGET_CCUT], header_fmt)
    data.write_row("A5", ["日期", "断头数", "C切数", "目标断头数", "目标C切数", "阶段", "原始单元格", "校验"], header_fmt)

    for item in source_rows:
        r = item["excel_row"] - 1
        row_no = item["excel_row"]
        write_formula(data, r, 0, f"='Sheet1'!A{row_no}", date_fmt, item["serial"])
        write_formula(data, r, 1, f"='Sheet1'!B{row_no}", int_fmt, item["breaks"])
        write_formula(data, r, 2, f"='Sheet1'!C{row_no}", int_fmt, item["ccut"])
        write_formula(data, r, 3, "=$B$4", int_fmt, TARGET_BREAKS)
        write_formula(data, r, 4, "=$D$4", int_fmt, TARGET_CCUT)
        write_formula(
            data,
            r,
            5,
            f'=IF(A{row_no}<DATE(2026,5,19),"上机/异常期",IF(A{row_no}<DATE(2026,5,25),"0.55%试调期","0.60%执行期"))',
            cell_fmt,
            stage_for(item["date"]),
        )
        write_formula(data, r, 6, f'="Sheet1!A{row_no}:C{row_no}"', cell_fmt, f"Sheet1!A{row_no}:C{row_no}")
        write_formula(
            data,
            r,
            7,
            f'=IF(AND(ISNUMBER(A{row_no}),ISNUMBER(B{row_no}),ISNUMBER(C{row_no})),"OK","检查")',
            cell_fmt,
            "OK",
        )

    chart_src.freeze_panes(4, 0)
    chart_src.set_column("A:A", 10)
    chart_src.set_column("B:E", 11)
    chart_src.set_column("G:N", 11)
    chart_src.set_column("P:R", 16)
    chart_src.merge_range("A1:E1", "PPT图表实际引用趋势表", title_fmt)
    chart_src.merge_range(
        "A2:E2",
        "本表 A4:E34 是折线图实际数据源；数据行是公式引用，不是手工编造。",
        note_fmt,
    )
    chart_src.write_row("A4", ["日期", "断头数", "C切数", "目标断头数", "目标C切数"], header_fmt)

    visible_rows = [item for item in source_rows if item["date"] >= date(2026, 5, 16)]
    for idx, item in enumerate(visible_rows):
        row = 4 + idx
        data_row = item["excel_row"] - 1
        excel_data_row = item["excel_row"]
        row_fmt = band_fmt if idx % 2 else cell_fmt
        row_int_fmt = band_int_fmt if idx % 2 else int_fmt
        write_formula(chart_src, row, 0, f"='趋势图数据'!A{excel_data_row}", date_fmt if idx % 2 == 0 else band_date_fmt, item["serial"])
        write_formula(chart_src, row, 1, f"='趋势图数据'!B{excel_data_row}", row_int_fmt, item["breaks"])
        write_formula(chart_src, row, 2, f"='趋势图数据'!C{excel_data_row}", row_int_fmt, item["ccut"])
        write_formula(chart_src, row, 3, f"='趋势图数据'!D{excel_data_row}", row_int_fmt, TARGET_BREAKS)
        write_formula(chart_src, row, 4, f"='趋势图数据'!E{excel_data_row}", row_int_fmt, TARGET_CCUT)

    chart_src.add_table(
        "A4:E34",
        {
            "name": "PPTTrendSourceTable",
            "style": "Table Style Medium 2",
            "columns": [{"header": h} for h in ["日期", "断头数", "C切数", "目标断头数", "目标C切数"]],
            "autofilter": False,
        },
    )
    chart_src.write_row("P4", ["节点日期", "节点高度", "节点说明"], header_fmt)
    write_formula(chart_src, 4, 15, "=DATE(2026,5,25)", date_fmt, NODE_DATE_SERIAL)
    chart_src.write_number(4, 16, 0, int_fmt)
    chart_src.write(4, 17, NODE_LABEL, cell_fmt)
    write_formula(chart_src, 5, 15, "=DATE(2026,5,25)", date_fmt, NODE_DATE_SERIAL)
    chart_src.write_number(5, 16, 200, int_fmt)
    chart_src.write(5, 17, NODE_LABEL, cell_fmt)

    first_serial = visible_rows[0]["serial"]
    last_serial = visible_rows[-1]["serial"]
    add_trend_chart(workbook, chart_src, "PPT图表引用表", 4, "G4", "近一个月断头/C切趋势", first_serial, last_serial)

    trend.freeze_panes(2, 0)
    trend.set_column("A:A", 10)
    trend.set_column("B:E", 11)
    trend.set_column("F:I", 11)
    trend.merge_range("A1:I1", "10#机55/72黑丝近一个月断头/C切趋势", title_fmt)
    trend.merge_range(
        "A2:I2",
        "上方为 Excel 内置折线图，数据源绑定 PPT图表引用表!A4:E34；下方为同一数据源的完整引用表。",
        note_fmt,
    )
    add_trend_chart(workbook, trend, "PPT图表引用表", 4, "A4", "近一个月断头/C切趋势", first_serial, last_serial)

    trend.merge_range("A24:E24", "下表为上方线图的完整引用数据", section_fmt)
    trend.write_row("A25", ["日期", "断头数", "C切数", "目标断头数", "目标C切数"], header_fmt)
    for idx, item in enumerate(visible_rows):
        row = 25 + idx
        src_excel_row = 5 + idx
        row_fmt = band_fmt if idx % 2 else cell_fmt
        row_int_fmt = band_int_fmt if idx % 2 else int_fmt
        write_formula(trend, row, 0, f"='PPT图表引用表'!A{src_excel_row}", date_fmt if idx % 2 == 0 else band_date_fmt, item["serial"])
        write_formula(trend, row, 1, f"='PPT图表引用表'!B{src_excel_row}", row_int_fmt, item["breaks"])
        write_formula(trend, row, 2, f"='PPT图表引用表'!C{src_excel_row}", row_int_fmt, item["ccut"])
        write_formula(trend, row, 3, f"='PPT图表引用表'!D{src_excel_row}", row_int_fmt, TARGET_BREAKS)
        write_formula(trend, row, 4, f"='PPT图表引用表'!E{src_excel_row}", row_int_fmt, TARGET_CCUT)

    trend.add_table(
        "A25:E55",
        {
            "name": "TrendVisibleTable",
            "style": "Table Style Medium 2",
            "columns": [{"header": h} for h in ["日期", "断头数", "C切数", "目标断头数", "目标C切数"]],
            "autofilter": False,
        },
    )
    trend.merge_range(
        "A58:I60",
        "审计说明：本工作簿不含图片趋势图。趋势图页的图表为标准 Excel chart，对应数据范围为 PPT图表引用表!A4:E34。",
        audit_fmt,
    )

    workbook.close()
    print(OUTPUT)


if __name__ == "__main__":
    main()
