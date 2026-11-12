import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const sourcePath =
  "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/work/qc_2026_input/2026/断头趋势.xlsx";
const outputDir =
  "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/outputs/qcc_source_data_excel";
const outputPath = path.join(outputDir, "55-72黑丝断头_C切趋势_原生图表版.xlsx");
const qaDir =
  "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/work/spreadsheets/qcc_trend_source_excel/qa";

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(qaDir, { recursive: true });

const sourceBlob = await FileBlob.load(sourcePath);
const sourceWorkbook = await SpreadsheetFile.importXlsx(sourceBlob);
const sourceSheet = sourceWorkbook.worksheets.getItem("Sheet1");
const sourceValues = sourceSheet.getRange("A1:C42").values;

const workbook = Workbook.create();
const trendSheet = workbook.worksheets.add("趋势图");
const chartSourceSheet = workbook.worksheets.add("PPT图表引用表");
const dataSheet = workbook.worksheets.add("趋势图数据");
const rawSheet = workbook.worksheets.add("Sheet1");

for (const sheet of [trendSheet, chartSourceSheet, dataSheet, rawSheet]) {
  sheet.showGridLines = false;
}

function formatTrendChart(chart, title) {
  chart.title = title;
  chart.titleTextStyle.fontSize = 14;
  chart.hasLegend = true;
  chart.legend = { position: "top", textStyle: { fontSize: 9 } };
  chart.xAxis = {
    axisType: "textAxis",
    textStyle: { fontSize: 8 },
    majorGridlines: { fill: "#D9D9D9", style: "dashed", width: 0.75 },
  };
  chart.yAxis = {
    numberFormatCode: "#,##0",
    min: 0,
    max: 200,
    majorUnit: 25,
    title: { text: "数量（个/日）" },
    textStyle: { fontSize: 9 },
    majorGridlines: { fill: "#E5E7EB", style: "solid", width: 0.75 },
  };
  chart.plotAreaFill = { type: "solid", color: "#FFFFFF" };
  chart.chartFill = { type: "solid", color: "#FFFFFF" };
  chart.series.items[0].line = { fill: "#B33A3A", style: "solid", width: 2 };
  chart.series.items[1].line = { fill: "#1F6FAE", style: "solid", width: 2 };
  chart.series.items[2].line = { fill: "#B99A9A", style: "dashed", width: 1.25 };
  chart.series.items[3].line = { fill: "#9BB7CC", style: "dashed", width: 1.25 };
}

rawSheet.getRange("A1:C42").values = sourceValues;
rawSheet.getRange("A6:A42").format.numberFormat = "m/d";
rawSheet.getRange("B6:C42").format.numberFormat = "#,##0";
rawSheet.getRange("A1:C1").format = {
  fill: "#0B3A5B",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
};
rawSheet.getRange("A:C").format.autofitColumns();
rawSheet.getRange("A:C").format.columnWidthPx = 96;

dataSheet.freezePanes.freezeRows(5);
dataSheet.getRange("A1:H1").merge();
dataSheet.getRange("A1").values = [["55/72黑丝断头/C切完整原始引用数据"]];
dataSheet.getRange("A1:H1").format = {
  fill: "#0B3A5B",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
dataSheet.getRange("A1:H1").format.rowHeightPx = 34;

dataSheet.getRange("A2:H3").values = [
  [
    "数据来源",
    "导入自原始工作簿：断头趋势.xlsx；原始范围：Sheet1!A6:C42",
    null,
    null,
    "说明",
    "本页 A6:C42 全部公式引用本工作簿 Sheet1 原始导入数据；PPT图表引用表再引用本页。",
    null,
    null,
  ],
  [
    "关键节点",
    "2026-05-25 0.60%整体执行",
    null,
    null,
    "目标线",
    "断头目标 90；C切目标 50。目标值可在 B4/D4 调整。",
    null,
    null,
  ],
];
dataSheet.getRange("A2:A3").format = {
  fill: "#EAF4FA",
  font: { bold: true, color: "#0B3A5B" },
};
dataSheet.getRange("E2:E3").format = {
  fill: "#EAF4FA",
  font: { bold: true, color: "#0B3A5B" },
};
dataSheet.getRange("B2:D3").merge(true);
dataSheet.getRange("F2:H3").merge(true);
dataSheet.getRange("A2:H3").format = {
  wrapText: true,
  verticalAlignment: "center",
};
dataSheet.getRange("A2:H3").format.borders = {
  preset: "all",
  style: "thin",
  color: "#B7C9D6",
};

dataSheet.getRange("A4:D4").values = [["断头目标", 90, "C切目标", 50]];
dataSheet.getRange("A4:D4").format = {
  fill: "#FFF7E6",
  font: { bold: true, color: "#7A4B00" },
  horizontalAlignment: "center",
};
dataSheet.getRange("B4:D4").format.numberFormat = "#,##0";
dataSheet.getRange("A4:D4").format.borders = {
  preset: "all",
  style: "thin",
  color: "#E4B860",
};

dataSheet.getRange("A5:H5").values = [
  [
    "日期",
    "断头数",
    "C切数",
    "断头目标",
    "C切目标",
    "阶段",
    "原始单元格",
    "校验",
  ],
];
dataSheet.getRange("A5:H5").format = {
  fill: "#118AB2",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
};

const completeDataFormulas = [];
for (let sourceRow = 6; sourceRow <= 42; sourceRow += 1) {
  completeDataFormulas.push([
    `='Sheet1'!A${sourceRow}`,
    `='Sheet1'!B${sourceRow}`,
    `='Sheet1'!C${sourceRow}`,
    "=$B$4",
    "=$D$4",
    `=IF(A${sourceRow}<DATE(2026,5,19),"上机/异常期",IF(A${sourceRow}<DATE(2026,5,25),"0.55%试调期","0.60%执行期"))`,
    `"Sheet1!A${sourceRow}:C${sourceRow}"`,
    `=IF(AND(ISNUMBER(A${sourceRow}),ISNUMBER(B${sourceRow}),ISNUMBER(C${sourceRow})),"OK","检查")`,
  ]);
}
dataSheet.getRange("A6:H42").formulas = completeDataFormulas;
dataSheet.getRange("A6:A42").format.numberFormat = "m/d";
dataSheet.getRange("B6:E42").format.numberFormat = "#,##0";
dataSheet.getRange("A6:H42").format.borders = {
  preset: "inside",
  style: "thin",
  color: "#D6E4EC",
};
dataSheet.getRange("A5:H42").format.borders = {
  preset: "outside",
  style: "thin",
  color: "#7DB5CA",
};
dataSheet.getRange("A:A").format.columnWidthPx = 72;
dataSheet.getRange("B:E").format.columnWidthPx = 86;
dataSheet.getRange("F:F").format.columnWidthPx = 118;
dataSheet.getRange("G:G").format.columnWidthPx = 128;
dataSheet.getRange("H:H").format.columnWidthPx = 64;

chartSourceSheet.freezePanes.freezeRows(4);
chartSourceSheet.getRange("A1:E1").merge();
chartSourceSheet.getRange("A1").values = [["PPT图表实际引用趋势表"]];
chartSourceSheet.getRange("A1:E1").format = {
  fill: "#0B3A5B",
  font: { bold: true, color: "#FFFFFF", size: 15 },
  horizontalAlignment: "center",
};
chartSourceSheet.getRange("A2:E2").merge();
chartSourceSheet.getRange("A2").values = [[
  "本表 A4:E34 是 PPT 折线图的实际数据源，所有数据由公式引用“趋势图数据”页，不手工编造。",
]];
chartSourceSheet.getRange("A2:E2").format = {
  fill: "#EAF4FA",
  font: { bold: true, color: "#0B3A5B" },
  wrapText: true,
};
chartSourceSheet.getRange("A4:E4").values = [
  ["日期", "断头数", "C切数", "断头目标", "C切目标"],
];
chartSourceSheet.getRange("A4:E4").format = {
  fill: "#118AB2",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
};
const chartSourceFormulas = [];
for (let row = 13; row <= 42; row += 1) {
  chartSourceFormulas.push([
    `=TEXT('趋势图数据'!A${row},"m/d")`,
    `='趋势图数据'!B${row}`,
    `='趋势图数据'!C${row}`,
    `='趋势图数据'!D${row}`,
    `='趋势图数据'!E${row}`,
  ]);
}
chartSourceSheet.getRange("A5:E34").formulas = chartSourceFormulas;
chartSourceSheet.getRange("B5:E34").format.numberFormat = "#,##0";
chartSourceSheet.getRange("A4:E34").format.borders = {
  preset: "all",
  style: "thin",
  color: "#D6E4EC",
};
const chartSourceTable = chartSourceSheet.tables.add("A4:E34", true, "PPTTrendSourceTable");
chartSourceTable.style = "TableStyleMedium2";
chartSourceTable.showFilterButton = false;
chartSourceSheet.getRange("A:A").format.columnWidthPx = 72;
chartSourceSheet.getRange("B:E").format.columnWidthPx = 86;

trendSheet.freezePanes.freezeRows(2);
trendSheet.getRange("A1:I1").merge();
trendSheet.getRange("A1").values = [["10#机55/72黑丝近一个月断头/C切趋势"]];
trendSheet.getRange("A1:I1").format = {
  fill: "#0B3A5B",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
trendSheet.getRange("A1:I1").format.rowHeightPx = 38;
trendSheet.getRange("A2:I2").merge();
trendSheet.getRange("A2").values = [[
  "图表数据源：PPT图表引用表!A4:E34；PPT图表引用表由公式引用趋势图数据；趋势图数据由公式引用 Sheet1 原始导入数据。",
]];
trendSheet.getRange("A2:I2").format = {
  fill: "#EAF4FA",
  font: { bold: true, color: "#0B3A5B" },
  horizontalAlignment: "center",
  wrapText: true,
};
trendSheet.getRange("A2:I2").format.rowHeightPx = 30;

const nativeChart = trendSheet.charts.add("line", chartSourceSheet.getRange("A4:E34"));
formatTrendChart(nativeChart, "近一个月断头/C切趋势");
nativeChart.setPosition("A4", "I22");

trendSheet.getRange("A24:E24").merge();
trendSheet.getRange("A24").values = [["下表为上方线图的完整引用数据"]];
trendSheet.getRange("A24:E24").format = {
  fill: "#0B3A5B",
  font: { bold: true, color: "#FFFFFF", size: 13 },
  horizontalAlignment: "center",
};

trendSheet.getRange("A25:E25").values = [
  ["日期", "断头数", "C切数", "断头目标", "C切目标"],
];
const visibleFormulas = [];
for (let row = 5; row <= 34; row += 1) {
  visibleFormulas.push([
    `='PPT图表引用表'!A${row}`,
    `='PPT图表引用表'!B${row}`,
    `='PPT图表引用表'!C${row}`,
    `='PPT图表引用表'!D${row}`,
    `='PPT图表引用表'!E${row}`,
  ]);
}
trendSheet.getRange("A26:E55").formulas = visibleFormulas;
trendSheet.getRange("A25:E25").format = {
  fill: "#118AB2",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
};
trendSheet.getRange("A26:E55").format = { verticalAlignment: "center" };
trendSheet.getRange("B26:E55").format.numberFormat = "#,##0";
trendSheet.getRange("A25:E55").format.borders = {
  preset: "all",
  style: "thin",
  color: "#D6E4EC",
};
const visibleTable = trendSheet.tables.add("A25:E55", true, "TrendVisibleTable");
visibleTable.style = "TableStyleMedium2";
visibleTable.showFilterButton = false;
trendSheet.getRange("A:A").format.columnWidthPx = 72;
trendSheet.getRange("B:E").format.columnWidthPx = 86;
trendSheet.getRange("F:I").format.columnWidthPx = 82;

trendSheet.getRange("A58:I60").merge();
trendSheet.getRange("A58").values = [[
  "审计说明：本页只保留 Excel 原生折线图，不含图片对象。图表使用 PPT图表引用表!A4:E34 作为数据源；该表由公式引用趋势图数据，趋势图数据由公式引用 Sheet1 原始导入数据。",
]];
trendSheet.getRange("A58:I60").format = {
  fill: "#FFF7E6",
  font: { color: "#7A4B00" },
  wrapText: true,
  verticalAlignment: "center",
};
trendSheet.getRange("A58:I60").format.borders = {
  preset: "outside",
  style: "thin",
  color: "#E4B860",
};

const chartSourceNative = chartSourceSheet.charts.add("line", chartSourceSheet.getRange("A4:E34"));
formatTrendChart(chartSourceNative, "PPT图表引用表对应线图");
chartSourceNative.setPosition("G4", "N25");

const keyInspect = await workbook.inspect({
  kind: "table",
  sheetId: "趋势图",
  range: "A1:I55",
  include: "values,formulas",
  tableMaxRows: 40,
  tableMaxCols: 9,
  maxChars: 16000,
});
await fs.writeFile(path.join(qaDir, "trend-redo-inspect.ndjson"), keyInspect.ndjson);

const sourceInspect = await workbook.inspect({
  kind: "table",
  sheetId: "PPT图表引用表",
  range: "A1:E34",
  include: "values,formulas",
  tableMaxRows: 40,
  tableMaxCols: 5,
  maxChars: 12000,
});
await fs.writeFile(path.join(qaDir, "ppt-source-redo-inspect.ndjson"), sourceInspect.ndjson);

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
  maxChars: 4000,
});
await fs.writeFile(path.join(qaDir, "formula-error-scan.ndjson"), errorScan.ndjson);

for (const [sheetName, fileName] of [
  ["趋势图", "趋势图_redo2.png"],
  ["PPT图表引用表", "PPT图表引用表_redo2.png"],
  ["趋势图数据", "趋势图数据_redo2.png"],
]) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(path.join(qaDir, fileName), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
