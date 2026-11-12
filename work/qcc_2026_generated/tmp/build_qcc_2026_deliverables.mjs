import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { Workbook, SpreadsheetFile, Presentation, PresentationFile } from "@oai/artifact-tool";

const execFileAsync = promisify(execFile);

const ROOT = "/Users/USER/Documents/Codex/2026-06-21/new-chat-3";
const OUT = `${ROOT}/outputs/qcc_2026_generated`;
const TMP = `${ROOT}/work/qcc_2026_generated/tmp`;
const ASSET = `${ROOT}/work/presentations/qcc_55_72_heisi/tmp/assets`;
const LOGO_DOCX = "/tmp/codex-remote-attachments/019ee940-a5e2-7413-ad9f-29b56e38089a/7B4A42F2-CB03-432C-BC7A-F1B775261D66/1-logo.docx";
const LOGO_ROOT = `${TMP}/logo_assets`;
const LOGO_MEDIA = `${LOGO_ROOT}/raw/word/media`;
const LOGO1 = `${LOGO_MEDIA}/image1.png`;
const LOGO2 = `${LOGO_MEDIA}/image2.png`;
const LOGO3 = `${LOGO_MEDIA}/image3.png`;

const W = 1280;
const H = 720;
const M = 58;
const FONT = "Microsoft YaHei";

const C = {
  navy: "#0A2742",
  blue: "#1E6AA8",
  teal: "#0E9A9A",
  mint: "#DDF3EC",
  amber: "#D99116",
  amber2: "#FFF3DA",
  red: "#B9403A",
  red2: "#F9E3E0",
  green: "#177A51",
  green2: "#E2F3EA",
  ink: "#102033",
  muted: "#667386",
  line: "#D8DEE8",
  paper: "#F4F6F8",
  white: "#FFFFFF",
  soft: "#EAF0F6",
};

const daily = [
  ["5月09日", 151, 103, "0.45%", "上机初期；CP5 12线POY"],
  ["5月10日", 143, 82, "0.45%", "上机初期"],
  ["5月11日", 180, 64, "0.45%", "调整油轮导丝器"],
  ["5月12日", 142, 82, "0.45%", "换至8线生产，仍未稳定"],
  ["5月13日", 190, 64, "0.45%", "调整油轮导丝器"],
  ["5月14日", 141, 89, "0.45%", "调整油轮导丝器"],
  ["5月15日", 243, 114, "0.45%", "张力CV波动明显"],
  ["5月16日", 180, 71, "0.45%", "竹本油剂试验，无明显效果"],
  ["5月17日", 130, 50, "0.45%", "仍高位波动"],
  ["5月18日", 143, 54, "0.45%", "仍高位波动"],
  ["5月19日", 112, 34, "0.55%", "前纺含油0.55%打样"],
  ["5月20日", 98, 21, "0.55%", "有效果但不明显"],
  ["5月21日", 119, 24, "0.55/0.60%", "0.60%打样10锭"],
  ["5月22日", 107, 21, "0.55%", "0.55%整体做到"],
  ["5月23日", 93, 19, "0.55%", "断头仍高于目标"],
  ["5月24日", 97, 10, "0.55%", "断头仍高于目标"],
  ["5月25日", 46, 1, "0.60%", "0.60%整体做到"],
  ["5月26日", 52, 3, "0.60%", "稳定"],
  ["5月27日", 50, 2, "0.60%", "稳定"],
  ["5月28日", 26, 6, "0.60%", "稳定"],
  ["5月29日", 35, 0, "0.60%", "稳定"],
  ["5月30日", 46, 9, "0.60%", "稳定"],
  ["5月31日", 40, 7, "0.60%", "稳定"],
  ["6月01日", 19, 1, "0.60%", "稳定"],
  ["6月02日", 23, 1, "0.60%", "稳定"],
  ["6月03日", 35, 0, "0.60%", "稳定"],
  ["6月04日", 28, 3, "0.60%", "稳定"],
  ["6月05日", 30, 0, "0.60%", "稳定"],
  ["6月06日", 32, 1, "0.60%", "稳定"],
  ["6月07日", 34, 1, "0.60%", "稳定"],
  ["6月08日", 29, 0, "0.60%", "稳定"],
  ["6月09日", 39, 2, "0.60%", "稳定"],
  ["6月10日", 45, 1, "0.60%", "稳定"],
  ["6月11日", 28, 0, "0.60%", "稳定"],
  ["6月12日", 34, 2, "0.60%", "稳定"],
  ["6月13日", 16, 1, "0.60%", "稳定"],
  ["6月14日", 16, 0, "0.60%", "稳定"],
];

const targets = { breaks: 90, cCuts: 50, excellentRate: 0.85 };

function byStage(rows, label, match) {
  const items = rows.filter((r) => match(r));
  const b = items.reduce((s, r) => s + r[1], 0);
  const c = items.reduce((s, r) => s + r[2], 0);
  return [label, items.length, b, +(b / items.length).toFixed(1), c, +(c / items.length).toFixed(1)];
}

const stages = [
  byStage(daily, "上机初期/含油率0.45%", (r) => r[3] === "0.45%"),
  byStage(daily, "含油率0.55%试调", (r) => r[3].startsWith("0.55")),
  byStage(daily, "含油率0.60%整体执行", (r) => r[3] === "0.60%"),
];

const factors = [
  ["人", 1, "操作培训不到位", "查培训记录、现场操作一致性", "操作按作业指导书执行，无系统性漏项", "班组按工艺要求执行，异常主要随含油阶段变化，不随人员班次变化", "非要因", "继续培训与现场抽查"],
  ["人", 2, "上机异常反馈不及时", "查异常发现、反馈、验证闭环记录", "异常24小时内反馈并形成验证结论", "前期0.55%试调有效果但不明显，整体切换至0.60%滞后至5月25日", "要因", "建立前纺-加弹联动验证机制"],
  ["机", 3, "DTY10#机油轮导丝器异常", "5月11-14日停车半台调整后跟踪", "调整后断头、C切应同步降至目标内", "5月15-18日断头仍为243/180/130/143，C切仍有114/71/50/54", "非要因", "作为点检项目，不作为主因"],
  ["机", 4, "张力传感/报警阈值异常", "查看张力CV波动及报警记录", "异常修正后指标应持续达标", "CV波动是生产不稳定的表现，未见阈值调整带来持续改善", "非要因", "保留张力趋势监控"],
  ["料", 5, "POY原丝含油率偏低", "按0.45%、0.55%、0.60%分阶段对比", "断头≤90个/日，C切≤50个/日", "0.45%阶段断头日均164.3、C切77.3；0.60%阶段断头33.5、C切2.0", "要因", "POY含油控制到0.60%"],
  ["料", 6, "前纺线别切换影响", "对比CP5 12线及5月12日后8线生产", "换线后指标应立即改善", "5月12日换8线后仍连续多日高位波动，直至0.60%整体执行才稳定", "非要因", "线别切换作为记录项"],
  ["料", 7, "竹本油剂试验影响", "5月16日换竹本油剂后跟踪", "换油剂后断头、C切应明显改善", "5月16-18日断头180/130/143，C切71/50/54，未形成明显改善", "非要因", "不作为本课题主因"],
  ["法", 8, "55/72黑丝POY含油控制标准未固化", "查上机初期与试调过程", "标准应能支撑连续达标", "0.55%阶段C切下降但断头仍高于目标，0.60%整体后连续21天达标", "要因", "固化55/72黑丝POY含油目标"],
  ["法", 9, "DTY端局部参数调整不足", "核对油轮导丝器、二罗拉、K值、冷却板等临时调整", "调整后应出现持续稳定改善", "DTY端局部调整未改变高断头趋势，改善来自POY含油整体执行", "非要因", "避免把临时试调写成有效对策"],
  ["法", 10, "检验/分级标准不统一", "核对系统记录、日报、趋势图", "数据口径一致，异常能复核", "断头与C切数据连续记录，阶段差异清晰，非分级口径造成", "非要因", "保留同一口径日报"],
  ["环", 11, "环境温湿度波动", "对比同期间环境与异常走势", "环境变化应与异常同步", "异常改善节点与0.60%整体执行一致，未见环境变化解释阶段改善", "非要因", "保持车间环境巡检"],
  ["环", 12, "上机初期现场状态波动", "观察上机初期与稳定生产期现场情况", "现场状态调整后应独立改善", "上机初期波动无法解释0.55%与0.60%阶段差异", "非要因", "纳入上机确认清单"],
];

const measures = [
  ["POY原丝含油率偏低", "POY含油率目标由0.45/0.55调整并固化至0.60%", "断头≤90个/日，C切≤50个/日", "5月21日打样10锭，5月25日整体执行；每日跟踪断头、C切", "前纺/加弹工艺", "2026-05-25", "0.60%阶段21天断头日均33.5、C切日均2.0", "有效"],
  ["55/72黑丝POY含油控制标准未固化", "建立55/72黑丝专用含油控制标准", "上机后不再以0.55%作为最终稳定标准", "首批上机确认含油目标、批量切换时间、异常复核节点", "工艺组", "2026-06-01", "0.55%阶段断头日均104.3仍高于目标；0.60%整体后全部达标", "有效"],
  ["前纺-加弹联动验证滞后", "建立异常联动验证与闭环表", "试调结论当天记录，整体调整有责任人与期限", "将打样、线别、油剂、含油率、断头/C切结果纳入一表跟踪", "QC小组", "2026-06-05", "非要因项被排除，减少DTY端重复试错", "有效"],
];

const standards = [
  ["控制项目", "标准/防措", "频次", "责任人", "记录"],
  ["POY含油率", "55/72黑丝POY含油率按0.60%目标控制；低于目标需暂停扩大生产并复核", "每批/每日", "前纺工艺", "POY含油确认表"],
  ["上机确认", "记录POY来源线别、油剂、含油率、打样锭位、批量切换时间", "每次上机/换批", "加弹工艺", "上机确认单"],
  ["过程监控", "每日跟踪10#机断头≤90个、C切≤50个；任一超标当天反馈", "每日", "班组/QC小组", "每日目标跟踪表"],
  ["异常联动", "前纺、加弹、质检三方确认异常项，不把线别/换油剂/DTY临时调整直接判为主因", "异常发生时", "QC组长", "异常验证闭环表"],
  ["防复发", "连续3天断头或C切接近目标上限时，提前复核POY含油与张力CV趋势", "触发式", "工艺组", "防复发检查表"],
];

async function ensureDirs() {
  await fs.mkdir(OUT, { recursive: true });
  await fs.mkdir(`${TMP}/preview`, { recursive: true });
  await fs.mkdir(`${TMP}/qa`, { recursive: true });
  await fs.rm(LOGO_ROOT, { recursive: true, force: true });
  await fs.mkdir(LOGO_ROOT, { recursive: true });
  await execFileAsync("unzip", ["-qo", LOGO_DOCX, "word/media/*", "-d", `${LOGO_ROOT}/raw`]);
}

function setTitle(sheet, title, subtitle = "") {
  sheet.showGridLines = false;
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    fill: C.navy,
    font: { bold: true, color: C.white, size: 16 },
    horizontalAlignment: "center",
    verticalAlignment: "middle",
  };
  sheet.getRange("A1").format.rowHeightPx = 34;
  if (subtitle) {
    sheet.getRange("A2:H2").merge();
    sheet.getRange("A2").values = [[subtitle]];
    sheet.getRange("A2").format = {
      fill: C.soft,
      font: { color: C.ink, size: 11 },
      wrapText: true,
      verticalAlignment: "middle",
    };
    sheet.getRange("A2").format.rowHeightPx = 42;
  }
}

function styleHeader(range) {
  range.format = {
    fill: C.blue,
    font: { bold: true, color: C.white },
    horizontalAlignment: "center",
    verticalAlignment: "middle",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#B9C7D8" },
  };
}

function styleBody(range) {
  range.format = {
    verticalAlignment: "middle",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#D8DEE8" },
  };
}

async function saveWorkbook(workbook, filename, previewSheet, previewRange) {
  const outPath = `${OUT}/${filename}`;
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(outPath);
  const png = await workbook.render({ sheetName: previewSheet, range: previewRange, scale: 1, format: "png" });
  await fs.writeFile(`${TMP}/preview/${filename}.png`, new Uint8Array(await png.arrayBuffer()));
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "formula error scan",
  });
  await fs.writeFile(`${TMP}/qa/${filename}.formula-scan.ndjson`, errors.ndjson);
  return outPath;
}

async function buildFactorWorkbook() {
  const wb = Workbook.create();
  const s = wb.worksheets.add("末端因素确认表");
  setTitle(s, "55/72黑丝末端因素验证数据", "要因控制为3项；非要因覆盖人、机、料、法、环五个方面，数据口径来自10#机每日断头/C切记录。");
  const headers = ["维度", "序号", "末端因素", "验证方法", "验证标准", "验证数据/事实", "结论", "处置口径"];
  s.getRange("A4:H4").values = [headers];
  s.getRange(`A5:H${4 + factors.length}`).values = factors;
  styleHeader(s.getRange("A4:H4"));
  styleBody(s.getRange(`A5:H${4 + factors.length}`));
  s.getRange("A:H").format.autofitColumns();
  s.getRange("A:A").format.columnWidthPx = 72;
  s.getRange("B:B").format.columnWidthPx = 52;
  s.getRange("C:C").format.columnWidthPx = 170;
  s.getRange("D:E").format.columnWidthPx = 180;
  s.getRange("F:F").format.columnWidthPx = 330;
  s.getRange("G:G").format.columnWidthPx = 76;
  s.getRange("H:H").format.columnWidthPx = 180;
  s.freezePanes.freezeRows(4);

  const ss = wb.worksheets.add("分阶段数据");
  setTitle(ss, "分阶段验证数据", "目标：断头≤90个/日，C切≤50个/日；0.60%整体执行后连续21天达标。");
  ss.getRange("A4:H4").values = [["阶段", "天数", "断头合计", "断头日均", "C切合计", "C切日均", "断头是否达标", "C切是否达标"]];
  ss.getRange("A5:H7").values = stages.map((r) => [...r, r[3] <= targets.breaks ? "是" : "否", r[5] <= targets.cCuts ? "是" : "否"]);
  styleHeader(ss.getRange("A4:H4"));
  styleBody(ss.getRange("A5:H7"));
  ss.getRange("A:H").format.autofitColumns();
  ss.getRange("D:D").format.numberFormat = "0.0";
  ss.getRange("F:F").format.numberFormat = "0.0";

  const ds = wb.worksheets.add("每日明细");
  setTitle(ds, "每日断头/C切跟踪明细", "含油率阶段、线别/油剂试验、目标达成情况合并记录。");
  ds.getRange("A4:I4").values = [["日期", "断头数", "断头目标", "断头达成", "C切数", "C切目标", "C切达成", "含油率阶段", "备注"]];
  ds.getRange(`A5:I${4 + daily.length}`).values = daily.map((r) => [
    r[0], r[1], targets.breaks, r[1] <= targets.breaks ? "达成" : "未达成",
    r[2], targets.cCuts, r[2] <= targets.cCuts ? "达成" : "未达成", r[3], r[4],
  ]);
  styleHeader(ds.getRange("A4:I4"));
  styleBody(ds.getRange(`A5:I${4 + daily.length}`));
  ds.getRange("A:I").format.autofitColumns();
  ds.freezePanes.freezeRows(4);

  const ms = wb.worksheets.add("人机料法环汇总");
  setTitle(ms, "人、机、料、法、环五方面总结", "料、法和联动反馈为主要突破口；人、机、环及部分料/法项经验证列为非要因。");
  const dims = ["人", "机", "料", "法", "环"].map((dim) => {
    const rows = factors.filter((f) => f[0] === dim);
    return [dim, rows.length, rows.filter((f) => f[6] === "要因").length, rows.filter((f) => f[6] === "非要因").length, rows.map((f) => `${f[2]}（${f[6]}）`).join("；")];
  });
  ms.getRange("A4:E4").values = [["维度", "末端因素数", "要因数", "非要因数", "结论摘要"]];
  ms.getRange("A5:E9").values = dims;
  styleHeader(ms.getRange("A4:E4"));
  styleBody(ms.getRange("A5:E9"));
  ms.getRange("A:E").format.autofitColumns();
  ms.getRange("E:E").format.columnWidthPx = 540;
  return saveWorkbook(wb, "末端因素验证数据_55-72黑丝.xlsx", "末端因素确认表", "A1:H16");
}

async function buildMeasuresWorkbook() {
  const wb = Workbook.create();
  const s = wb.worksheets.add("措施有效性验证");
  setTitle(s, "55/72黑丝措施有效性验证数据", "对策围绕已确认要因展开，效果用断头/C切日均及连续达标天数验证。");
  s.getRange("A4:H4").values = [["对应要因", "对策", "目标", "具体措施", "责任", "完成日期", "验证数据", "判定"]];
  s.getRange("A5:H7").values = measures;
  styleHeader(s.getRange("A4:H4"));
  styleBody(s.getRange("A5:H7"));
  s.getRange("A:H").format.autofitColumns();
  s.getRange("A:A").format.columnWidthPx = 170;
  s.getRange("B:D").format.columnWidthPx = 230;
  s.getRange("G:G").format.columnWidthPx = 260;

  const e = wb.worksheets.add("效果对比");
  setTitle(e, "改善前后效果对比", "改善前取5月9日至5月24日；改善后取5月25日至6月14日。");
  e.getRange("A4:G4").values = [["项目", "改善前日均", "改善后日均", "下降值", "下降率", "目标值", "结论"]];
  e.getRange("A5:G6").values = [
    ["断头数", 141.8, 33.5, 108.3, 0.764, 90, "改善后低于目标"],
    ["C切数", 56.4, 2.0, 54.4, 0.965, 50, "改善后低于目标"],
  ];
  styleHeader(e.getRange("A4:G4"));
  styleBody(e.getRange("A5:G6"));
  e.getRange("B5:F6").format.numberFormat = "0.0";
  e.getRange("E5:E6").format.numberFormat = "0.0%";
  e.getRange("A:G").format.autofitColumns();
  return saveWorkbook(wb, "措施有效性验证数据_55-72黑丝.xlsx", "措施有效性验证", "A1:H8");
}

async function buildDailyTargetWorkbook() {
  const wb = Workbook.create();
  const s = wb.worksheets.add("近一个月目标跟踪");
  setTitle(s, "近一个月每日目标值跟踪情况", "统计区间：2026年5月16日-2026年6月14日，共30天；目标为断头≤90个/日、C切≤50个/日。");
  const rows = daily.slice(-30);
  s.getRange("A4:K4").values = [["日期", "含油率阶段", "断头数", "断头目标", "断头达成", "C切数", "C切目标", "C切达成", "综合达成", "备注", "阶段判定"]];
  s.getRange(`A5:K${4 + rows.length}`).values = rows.map((r) => {
    const bp = r[1] <= targets.breaks;
    const cp = r[2] <= targets.cCuts;
    return [r[0], r[3], r[1], targets.breaks, bp ? "达成" : "未达成", r[2], targets.cCuts, cp ? "达成" : "未达成", bp && cp ? "达成" : "未达成", r[4], r[3] === "0.60%" ? "对策后" : "对策前/试调"];
  });
  styleHeader(s.getRange("A4:K4"));
  styleBody(s.getRange(`A5:K${4 + rows.length}`));
  s.getRange("A:K").format.autofitColumns();
  s.getRange("J:J").format.columnWidthPx = 230;
  s.freezePanes.freezeRows(4);

  const sum = wb.worksheets.add("目标达成汇总");
  setTitle(sum, "目标达成汇总", "对策后5月25日至6月14日连续21天断头和C切均达标。");
  const totalPass = rows.filter((r) => r[1] <= targets.breaks && r[2] <= targets.cCuts).length;
  const post = rows.filter((r) => r[3] === "0.60%");
  const postPass = post.filter((r) => r[1] <= targets.breaks && r[2] <= targets.cCuts).length;
  sum.getRange("A4:D4").values = [["统计项", "天数", "达标天数", "达标率"]];
  sum.getRange("A5:D7").values = [
    ["近一个月整体", rows.length, totalPass, totalPass / rows.length],
    ["对策前/试调", rows.length - post.length, totalPass - postPass, (totalPass - postPass) / (rows.length - post.length)],
    ["0.60%整体执行后", post.length, postPass, postPass / post.length],
  ];
  styleHeader(sum.getRange("A4:D4"));
  styleBody(sum.getRange("A5:D7"));
  sum.getRange("D5:D7").format.numberFormat = "0.0%";
  sum.getRange("A:D").format.autofitColumns();
  return saveWorkbook(wb, "近一个月每日目标值跟踪情况_55-72黑丝.xlsx", "近一个月目标跟踪", "A1:K34");
}

function addText(slide, value, x, y, w, h, o = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = value;
  s.text.style = {
    typeface: FONT,
    fontSize: o.size ?? 18,
    bold: o.bold ?? false,
    color: o.color || C.ink,
    alignment: o.align || "left",
    verticalAlignment: o.valign || "top",
    lineSpacing: o.lineSpacing ?? 1.08,
    autoFit: o.autoFit || "none",
    insets: o.insets || { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return s;
}

function shape(slide, x, y, w, h, o = {}) {
  return slide.shapes.add({
    geometry: o.geometry || "rect",
    position: { left: x, top: y, width: w, height: h },
    fill: o.fill ?? C.white,
    line: { style: "solid", fill: o.stroke ?? o.fill ?? C.line, width: o.lineWidth ?? 1 },
    borderRadius: o.radius,
    shadow: o.shadow,
  });
}

async function addImage(slide, file, x, y, w, h, fit = "cover") {
  const bytes = await fs.readFile(file);
  const ext = path.extname(file).toLowerCase();
  slide.images.add({
    blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    contentType: ext === ".png" ? "image/png" : "image/jpeg",
    position: { left: x, top: y, width: w, height: h },
    fit,
    alt: path.basename(file),
  });
}

function base(slide, index, title) {
  slide.background.fill = C.paper;
  shape(slide, 0, 0, W, 54, { fill: C.navy, stroke: C.navy });
  addText(slide, title, M, 12, 900, 32, { size: 25, bold: true, color: C.white });
  addText(slide, String(index).padStart(2, "0"), W - 148, 15, 54, 24, { size: 17, bold: true, color: "#9FE6E0", align: "right" });
}

function bullets(slide, items, x, y, w, h, o = {}) {
  addText(slide, items.map((v) => `• ${v}`).join("\n"), x, y, w, h, { size: o.size || 18, color: o.color || C.ink, lineSpacing: o.lineSpacing || 1.18 });
}

function miniTable(slide, headers, rows, x, y, w, rowH, widths = null, o = {}) {
  const colW = widths || headers.map(() => w / headers.length);
  const headerSize = o.headerSize || 16;
  const bodySize = o.bodySize || 15;
  const bodyAlign = o.bodyAlign || "center";
  const boldColumns = new Set(o.boldColumns || [0]);
  const boldCells = new Set(o.boldCells || []);
  let cx = x;
  headers.forEach((h, i) => {
    shape(slide, cx, y, colW[i], rowH, { fill: C.blue, stroke: C.white });
    addText(slide, h, cx + 6, y + 8, colW[i] - 12, rowH - 10, { size: headerSize, bold: true, color: C.white, align: "center", valign: "middle" });
    cx += colW[i];
  });
  rows.forEach((row, r) => {
    cx = x;
    row.forEach((v, i) => {
      const fill = r % 2 === 0 ? C.white : "#F8FAFC";
      shape(slide, cx, y + rowH * (r + 1), colW[i], rowH, { fill, stroke: C.line });
      addText(slide, String(v), cx + 8, y + rowH * (r + 1) + 7, colW[i] - 16, rowH - 10, {
        size: bodySize,
        bold: boldColumns.has(i) || boldCells.has(`${r}:${i}`),
        color: C.ink,
        align: o.leftColumns?.includes(i) ? "left" : bodyAlign,
        valign: "middle",
        lineSpacing: 1.08,
      });
      cx += colW[i];
    });
  });
}

function stat(slide, label, value, note, x, y, color = C.blue) {
  shape(slide, x, y, 238, 118, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
  addText(slide, label, x + 18, y + 16, 200, 24, { size: 16, color: C.muted });
  addText(slide, value, x + 18, y + 45, 200, 38, { size: 31, bold: true, color });
  addText(slide, note, x + 18, y + 86, 200, 20, { size: 13, color: C.muted });
}

function dimRows(dim, conclusion) {
  return factors
    .filter((f) => f[0] === dim && f[6] === conclusion)
    .map((f) => [f[2], f[5], f[7]]);
}

function dimSummary(dim) {
  const rows = factors.filter((f) => f[0] === dim);
  const cause = rows.filter((f) => f[6] === "要因").length;
  const nonCause = rows.filter((f) => f[6] === "非要因").length;
  return `${dim}维度共确认${rows.length}项末端因素：要因${cause}项，非要因${nonCause}项。`;
}

function factorSection(slide, title, rows, x, y, w, color, fill) {
  shape(slide, x, y, w, 462, { fill, stroke: color === C.red ? "#E6A09B" : C.line, shadow: "shadow-sm" });
  addText(slide, title, x + 22, y + 20, w - 44, 34, { size: 29, bold: true, color });
  const tableRows = rows.length ? rows : [["无", "本维度未确认导致本课题优等率低的主要原因。", "按标准监控"]];
  miniTable(
    slide,
    ["末端因素", "验证事实", "处置"],
    tableRows,
    x + 22,
    y + 74,
    w - 44,
    tableRows.length > 1 ? 108 : 128,
    [145, w - 44 - 145 - 118, 118],
    { headerSize: 15, bodySize: 14.2, bodyAlign: "left", leftColumns: [0, 1, 2], boldColumns: [0] }
  );
}

function dimensionFactorSlide(slide, dim) {
  const causes = dimRows(dim, "要因");
  const nonCauses = dimRows(dim, "非要因");
  addText(slide, dimSummary(dim), M, 78, W - 2 * M, 28, { size: 18, bold: true, color: C.navy });
  factorSection(slide, "要因", causes, M, 120, 548, C.red, "#FFF1F0");
  factorSection(slide, "非要因", nonCauses, M + 596, 120, 568, "#000000", C.white);
  shape(slide, M, 604, W - 2 * M, 54, { fill: C.mint, stroke: "#B8E2D4" });
  addText(slide, "结论口径", M + 24, 620, 120, 24, { size: 20, bold: true, color: C.green });
  addText(slide, "要因进入后续对策制定；非要因保留验证证据和监控要求，不写成最终有效对策。", M + 150, 618, 880, 28, { size: 18, bold: true, color: C.ink });
}

async function applyBrandLogos(p) {
  const total = p.slides.items.length;
  for (const [idx, slide] of p.slides.items.entries()) {
    if (idx === 0 || idx === total - 1) {
      await addImage(slide, LOGO2, 42, 24, 210, 84, "contain");
    } else {
      await addImage(slide, LOGO1, W - 70, 2, 48, 58, "contain");
    }
    await addImage(slide, LOGO3, (W - 300) / 2, H - 38, 300, 32, "contain");
  }
}

async function buildPptx() {
  const p = Presentation.create({ slideSize: { width: W, height: H } });
  const bg = `${ASSET}/qcc-fiber-bg-imagegen.png`;

  {
    const s = p.slides.add();
    await addImage(s, bg, 0, 0, W, H);
    shape(s, 0, 0, 720, H, { fill: "#061A2BDD", stroke: "#061A2B00" });
    addText(s, "提高55/72黑丝优等率", M, 238, 760, 78, { size: 54, bold: true, color: C.white });
    addText(s, "长空之王QC小组", M, 334, 400, 34, { size: 27, bold: true, color: "#5DD0CE" });
  }

  {
    const s = p.slides.add(); base(s, 2, "小组概况与课题范围");
    miniTable(s, ["项目", "内容"], [
      ["课题名称", "提高55/72黑丝优等率"],
      ["改善对象", "10#机55/72黑丝，重点跟踪断头数与C切数"],
      ["活动周期", "2026年5月9日-2026年6月14日"],
      ["小组定位", "前纺、加弹、质检联动的问题解决型QCC"],
    ], M, 116, 620, 62, [150, 470], { bodyAlign: "center", bodySize: 16, boldColumns: [0, 1] });
    shape(s, 740, 116, 420, 256, { fill: C.navy, stroke: C.navy });
    addText(s, "核心判定", 770, 150, 360, 34, { size: 27, bold: true, color: C.white });
    addText(s, "本课题不把DTY端局部试调写成最终有效对策，主线锁定为POY含油控制不足导致加工稳定性下降。", 770, 210, 340, 112, { size: 21, bold: true, color: C.white, lineSpacing: 1.16 });
  }

  {
    const s = p.slides.add(); base(s, 3, "现状调查：上机后断头、C切高位波动");
    stat(s, "0.45%阶段断头日均", "164.3个", "高于90个目标", M, 110, C.red);
    stat(s, "0.45%阶段C切日均", "77.3个", "高于50个目标", M + 260, 110, C.red);
    stat(s, "现状优等率", "<60%", "目标85%以上", M + 520, 110, C.amber);
    stat(s, "主要表现", "CV波动", "断头/C切同步异常", M + 780, 110, C.blue);
    await addImage(s, `${ASSET}/cv-wave-crop-sheet.jpg`, M, 280, 560, 292);
    await addImage(s, `${ASSET}/photo-initial-machine.jpg`, M + 600, 280, 560, 292);
    addText(s, "证据：刚上机及0.55%前后，张力CV波动与断头/C切异常同步出现。", M, 602, W - 2 * M, 30, { size: 18, bold: true, color: C.navy });
  }

  {
    const s = p.slides.add(); base(s, 4, "目标设定：断头和C切先稳定，优等率再提升");
    miniTable(s, ["指标", "现状/前期", "目标值", "改善后验证口径"], [
      ["断头数", "前期日均141.8个", "≤90个/日", "5月25日后连续21天达标"],
      ["C切数", "前期日均56.4个", "≤50个/日", "5月25日后连续21天达标"],
      ["优等率", "不足60%", "≥85%", "由断头、C切稳定作为过程支撑"],
    ], M, 136, W - 2 * M, 76, [180, 260, 220, 504], { bodyAlign: "center", bodySize: 16, boldColumns: [0, 2] });
    shape(s, M, 420, W - 2 * M, 100, { fill: C.mint, stroke: "#B8E2D4" });
    addText(s, "目标依据", M + 28, 444, 150, 32, { size: 25, bold: true, color: C.green, align: "center" });
    addText(s, "课题目标不是单纯追求某一天低值，而是通过POY含油控制固化，使55/72黑丝连续稳定生产，避免断头和C切反复高位波动。", M + 188, 440, 875, 54, { size: 22, bold: true, color: C.ink, align: "center", lineSpacing: 1.15 });
  }

  {
    const s = p.slides.add(); base(s, 5, "原因分析：按人、机、料、法、环五方面展开");
    const dims = [
      ["人", "异常反馈闭环不够快；操作培训未形成主因证据"],
      ["机", "油轮导丝器、张力报警、10#机状态均未解释持续改善"],
      ["料", "POY原丝含油率偏低是直接影响加工稳定性的关键"],
      ["法", "55/72黑丝含油控制标准未固化，前后道联动验证滞后"],
      ["环", "环境波动未与改善节点同步，判定非主因"],
    ];
    dims.forEach((d, i) => {
      const x = M + (i % 2) * 585;
      const y = 112 + Math.floor(i / 2) * 150;
      shape(s, x, y, 540, 106, { fill: i === 2 || i === 3 ? C.amber2 : C.white, stroke: i === 2 || i === 3 ? "#F0CC89" : C.line, shadow: "shadow-sm" });
      shape(s, x + 18, y + 22, 62, 62, { geometry: "ellipse", fill: i === 2 || i === 3 ? C.amber : C.blue, stroke: C.white, lineWidth: 2 });
      addText(s, d[0], x + 18, y + 35, 62, 30, { size: 25, bold: true, color: C.white, align: "center" });
      addText(s, d[1], x + 98, y + 26, 400, 52, { size: 18, color: C.ink, lineSpacing: 1.16 });
    });
  }

  {
    const s = p.slides.add(); base(s, 6, "末端因素确认：人维度");
    dimensionFactorSlide(s, "人");
  }

  {
    const s = p.slides.add(); base(s, 7, "末端因素确认：机维度");
    dimensionFactorSlide(s, "机");
  }

  {
    const s = p.slides.add(); base(s, 8, "末端因素确认：料维度");
    dimensionFactorSlide(s, "料");
  }

  {
    const s = p.slides.add(); base(s, 9, "末端因素确认：法维度");
    dimensionFactorSlide(s, "法");
  }

  {
    const s = p.slides.add(); base(s, 10, "末端因素确认：环维度");
    dimensionFactorSlide(s, "环");
  }

  {
    const s = p.slides.add(); base(s, 11, "要因确认（1）：POY原丝含油率偏低");
    shape(s, M, 100, 300, 140, { fill: C.green2, stroke: "#A9DDBE", shadow: "shadow-sm" });
    addText(s, "因素", M + 22, 124, 80, 26, { size: 19, bold: true, color: C.green });
    addText(s, "POY原丝含油率偏低", M + 22, 158, 245, 32, { size: 23, bold: true, color: C.ink });
    addText(s, "判定：要因", M + 22, 202, 160, 24, { size: 20, bold: true, color: C.green });
    miniTable(s, ["项目", "验证过程/结论"], [
      ["确认方法", "按0.45%、0.55%、0.60%三个含油阶段，对10#机每日断头和C切进行分阶段对比。"],
      ["验证标准", "断头≤90个/日，C切≤50个/日，且需要连续稳定达标。"],
      ["过程数据", "0.45%阶段断头日均164.3、C切77.3；0.55%阶段C切下降但断头仍为104.3，高于目标。"],
      ["确认结论", "0.60%整体执行后断头日均33.5、C切2.0，连续21天达标，因此判定为要因。"],
    ], M + 335, 100, W - M * 2 - 335, 96, [150, 679], { bodySize: 15, bodyAlign: "left", leftColumns: [1], boldColumns: [0] });
    miniTable(s, ["阶段", "断头日均", "C切日均", "结论"], [
      ["0.45%上机初期", "164.3", "77.3", "未达标"],
      ["0.55%试调", "104.3", "21.5", "断头未达标"],
      ["0.60%整体执行", "33.5", "2.0", "稳定达标"],
    ], M, 520, W - 2 * M, 48, [330, 250, 250, 334], { bodySize: 15, boldColumns: [0, 3] });
  }

  {
    const s = p.slides.add(); base(s, 12, "要因确认（2/3）：含油控制标准及联动验证未固化");
    shape(s, M, 96, 352, 150, { fill: C.amber2, stroke: "#F0CC89", shadow: "shadow-sm" });
    addText(s, "因素2/3", M + 22, 120, 110, 26, { size: 19, bold: true, color: C.amber });
    addText(s, "55/72黑丝专用含油标准与前后道联动机制未固化", M + 22, 154, 290, 56, { size: 21, bold: true, color: C.ink, lineSpacing: 1.08 });
    addText(s, "判定：要因", M + 22, 218, 160, 24, { size: 20, bold: true, color: C.amber });
    miniTable(s, ["时间", "验证过程", "过程结论"], [
      ["5月19日", "前纺含油由0.45%调整至0.55%，先打样10个锭位。", "C切下降，但断头仍高于目标。"],
      ["5月22日", "0.55%整体做到后继续跟踪。", "5月23-24日断头仍为93/97，未稳定。"],
      ["5月21日", "同步启动0.60%打样10锭验证。", "为后续整体执行提供方向。"],
      ["5月25日", "0.60%整体执行。", "断头和C切进入稳定达标区间。"],
    ], M + 385, 96, W - M * 2 - 385, 82, [100, 430, 249], { bodySize: 15, bodyAlign: "left", leftColumns: [1, 2], boldColumns: [0] });
    shape(s, M, 510, W - 2 * M, 92, { fill: C.green2, stroke: "#A9DDBE" });
    addText(s, "确认结论", M + 28, 535, 140, 28, { size: 23, bold: true, color: C.green, align: "center" });
    addText(s, "0.55%试调只能说明方向有效但不充分；0.60%整体执行后才连续稳定达标，说明55/72黑丝专用含油控制标准和前后道联动验证机制必须固化。", M + 188, 526, 850, 45, { size: 20, bold: true, color: C.ink, align: "center", lineSpacing: 1.12 });
  }

  {
    const s = p.slides.add(); base(s, 13, "非要因确认：线别、油剂、DTY端试调均被排除");
    miniTable(s, ["非要因项", "验证过程", "结论"], [
      ["前纺线别切换", "5月9日CP5 12线POY上机，5月12日后换到8线生产；换线后5月12-18日仍高位波动。", "非要因：换线未带来持续改善。"],
      ["竹本油剂试验", "5月16日换竹本油剂试验；5月16-18日断头180/130/143，C切71/50/54。", "非要因：换油剂未形成明显改善。"],
      ["DTY端油轮导丝器/局部参数", "5月11-14日调整油轮导丝器，后续仍出现5月15日断头243、C切114。", "非要因：不能写成最终有效对策。"],
      ["张力CV报警阈值/环境/分级口径", "CV波动作为异常证据，环境和分级口径未与改善节点同步。", "非要因：保留监控，不作为主因。"],
    ], M, 102, W - 2 * M, 88, [230, 620, 314], { bodySize: 15, bodyAlign: "left", leftColumns: [1, 2], boldColumns: [0, 2] });
    addText(s, "口径：非要因不是“不需要管”，而是不能写成造成本次优等率低的主要原因，也不能写成最终有效措施。", M, 610, W - 2 * M, 28, { size: 19, bold: true, color: C.navy, align: "center" });
  }

  {
    const s = p.slides.add(); base(s, 14, "要因确认汇总：先单因素确认，再进入对策");
    miniTable(s, ["类型", "确认因素", "是否要因", "后续处理"], [
      ["要因1", "POY原丝含油率偏低", "是", "按0.60%目标控制并每日跟踪"],
      ["要因2", "55/72黑丝POY含油控制标准未固化", "是", "固化55/72黑丝专用含油标准"],
      ["要因3", "前纺-加弹异常联动验证滞后", "是", "固化打样、整体执行和异常反馈闭环"],
      ["非要因", "CP5 12线/8线切换、竹本油剂试验、DTY端试调、CV报警阈值等", "否", "保留验证证据，不写成有效对策"],
    ], M, 112, W - 2 * M, 74, [150, 460, 150, 404], { bodySize: 14.8, bodyAlign: "center", boldColumns: [0, 2] });
    shape(s, M, 500, W - 2 * M, 88, { fill: C.mint, stroke: "#A9DDBE" });
    addText(s, "总结", M + 28, 526, 110, 30, { size: 24, bold: true, color: C.green, align: "center" });
    addText(s, "要因确认完成后，对策只围绕“POY含油控制”和“前后道联动标准化”展开，避免把试验排除项写成有效措施。", M + 156, 520, 900, 38, { size: 21, bold: true, color: C.ink, align: "center", lineSpacing: 1.15 });
  }

  {
    const s = p.slides.add(); base(s, 15, "对策制定：围绕确认要因建立闭环");
    miniTable(s, ["确认问题", "对策", "目标", "责任"], measures.map((m) => [m[0], m[1], m[2], m[4]]), M, 118, W - 2 * M, 84, [260, 390, 320, 194], { bodySize: 15, bodyAlign: "center", boldColumns: [0, 2] });
    addText(s, "对策原则：只写被数据验证有效的措施；DTY端临时试调、线别切换、换油剂试验归为排除项。", M, 604, W - 2 * M, 28, { size: 19, bold: true, color: C.navy, align: "center" });
  }

  {
    const s = p.slides.add(); base(s, 16, "对策实施：0.60%整体执行并建立联动跟踪");
    const events = [
      ["05/09", "CP5 12线POY上机，断头/C切高位"],
      ["05/12", "切换8线生产，未形成稳定改善"],
      ["05/16", "竹本油剂试验，未见明显效果"],
      ["05/19", "0.55%打样，5/22整体做到"],
      ["05/21", "0.60%打样10锭"],
      ["05/25", "0.60%整体执行，指标稳定"],
    ];
    shape(s, M, 245, W - 2 * M, 4, { fill: C.line, stroke: C.line });
    const gap = (W - 2 * M) / (events.length - 1);
    events.forEach((e, i) => {
      const x = M + i * gap;
      const col = i === events.length - 1 ? C.green : i >= 3 ? C.amber : C.blue;
      shape(s, x - 10, 236, 20, 20, { geometry: "ellipse", fill: col, stroke: C.white, lineWidth: 2 });
      addText(s, e[0], x - 42, 178, 84, 24, { size: 17, bold: true, color: col, align: "center" });
      addText(s, e[1], x - 78, 278, 156, 74, { size: 14.5, color: C.ink, align: "center", lineSpacing: 1.08 });
    });
    shape(s, M, 460, W - 2 * M, 94, { fill: C.white, stroke: C.line });
    bullets(s, ["5月25日起按0.60%整体执行，形成连续21天目标跟踪。", "非要因试验过程纳入记录，避免后续复盘时误写为有效措施。", "每日数据同时看断头与C切，任一超标即触发联动复核。"], M + 24, 482, W - 2 * M - 48, 58, { size: 18 });
  }

  {
    const s = p.slides.add(); base(s, 17, "效果检查：连续21天达成过程目标");
    stat(s, "断头日均下降", "76.4%", "141.8 → 33.5", M, 114, C.green);
    stat(s, "C切日均下降", "96.5%", "56.4 → 2.0", M + 260, 114, C.green);
    stat(s, "连续达标", "21天", "5/25-6/14", M + 520, 114, C.green);
    stat(s, "目标状态", "达成", "断头/C切均低于目标", M + 780, 114, C.green);
    miniTable(s, ["阶段", "断头日均", "C切日均", "状态"], [
      ["改善前/试调期", "141.8", "56.4", "未稳定"],
      ["0.60%整体执行后", "33.5", "2.0", "稳定达标"],
    ], M, 310, W - 2 * M, 76, [360, 240, 240, 324]);
  }

  {
    const s = p.slides.add(); base(s, 18, "标准化及防措");
    miniTable(s, ["控制点", "标准化/防措", "记录"], standards.slice(1).map((r) => [r[0], r[1], r[4]]), M, 96, W - 2 * M, 72, [180, 754, 230]);
    addText(s, "关键防措：当连续3天断头或C切接近目标上限时，提前复核POY含油与张力CV趋势。", M, 628, W - 2 * M, 28, { size: 18, bold: true, color: C.navy });
  }

  {
    const s = p.slides.add(); base(s, 19, "形成文件：数据、报告、标准化闭环");
    miniTable(s, ["文件", "用途"], [
      ["末端因素验证数据.docx", "保留人机料法环验证、要因/非要因结论"],
      ["措施有效性验证数据.docx", "证明每项对策对应要因且有效"],
      ["近一个月每日目标值跟踪情况.docx", "保留每日目标值、达成状态和连续达标证据"],
      ["标准化及防措.docx", "固化POY含油、上机确认、异常联动和防复发要求"],
    ], M, 124, W - 2 * M, 70, [420, 744], { bodySize: 15.2, boldColumns: [0], leftColumns: [1], bodyAlign: "left" });
    shape(s, M, 540, W - 2 * M, 70, { fill: C.mint, stroke: "#A9DDBE" });
    addText(s, "交付口径", M + 24, 560, 150, 26, { size: 22, bold: true, color: C.green });
    addText(s, "本页对应用户要求的五类材料，均围绕同一套数据和同一组要因结论生成。", M + 170, 558, 840, 30, { size: 19, color: C.ink });
  }

  {
    const s = p.slides.add();
    await addImage(s, bg, 0, 0, W, H);
    shape(s, 0, 0, W, H, { fill: "#061A2BDD", stroke: "#061A2B00" });
    addText(s, "总结与下一步", M, 120, 560, 56, { size: 45, bold: true, color: C.white });
    shape(s, M, 204, W - 2 * M, 1, { fill: "#FFFFFF55", stroke: "#FFFFFF55" });
    bullets(s, [
      "确认3个要因：POY含油率偏低、控制标准未固化、前纺-加弹联动验证滞后。",
      "排除线别切换、竹本油剂试验、DTY端局部调整等非要因。",
      "0.60%整体执行后，断头和C切连续21天达成目标。",
      "后续重点是执行标准化、防复发和每日过程跟踪。"
    ], M + 20, 238, 900, 190, { size: 24, color: C.white, lineSpacing: 1.25 });
    addText(s, "THANKS", M + 20, 560, 300, 44, { size: 32, bold: true, color: "#5DD0CE" });
  }

  await applyBrandLogos(p);

  const final = `${OUT}/QCC成果报告_提高55-72黑丝优等率_新增维度页_logo版.pptx`;
  for (const [idx, slide] of p.slides.items.entries()) {
    const stem = `qcc-report-slide-${String(idx + 1).padStart(2, "0")}`;
    const png = await p.export({ slide, format: "png", scale: 1 });
    await fs.writeFile(`${TMP}/preview/${stem}.png`, new Uint8Array(await png.arrayBuffer()));
  }
  const montage = await p.export({ format: "webp", montage: true, scale: 1 });
  await fs.writeFile(`${TMP}/preview/qcc-report-montage.webp`, new Uint8Array(await montage.arrayBuffer()));
  const inspect = await p.inspect({ kind: "slide,textbox,shape,image,table,chart", maxChars: 14000 });
  await fs.writeFile(`${TMP}/qa/qcc-report-inspect.ndjson`, inspect.ndjson);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(final);
  return final;
}

async function main() {
  await ensureDirs();
  const files = [];
  files.push(await buildFactorWorkbook());
  files.push(await buildMeasuresWorkbook());
  files.push(await buildDailyTargetWorkbook());
  files.push(await buildPptx());
  await fs.writeFile(`${OUT}/manifest.json`, JSON.stringify({ generatedAt: new Date().toISOString(), files }, null, 2));
  console.log(files.join("\n"));
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
