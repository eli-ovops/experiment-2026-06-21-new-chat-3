import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/outputs/长空之王QC小组_QCC成果报告_提高55-72黑丝优等率.pptx";
const TMP = "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/work/presentations/qcc_55_72_heisi/tmp";
const PREVIEW = `${TMP}/preview`;
const LAYOUT = `${TMP}/layout`;
const QA = `${TMP}/qa`;

const W = 1280;
const H = 720;
const M = 58;
const FONT = "Microsoft YaHei";

const C = {
  bg: "#F7F9FC",
  white: "#FFFFFF",
  ink: "#102033",
  muted: "#5B677A",
  light: "#E7ECF3",
  lighter: "#F2F5F8",
  navy: "#173B63",
  blue: "#2B6CB0",
  cyan: "#28A6C7",
  green: "#1F8A5B",
  green2: "#DFF4EA",
  orange: "#E68A00",
  orange2: "#FFF1D8",
  red: "#C43D3D",
  red2: "#FBE4E4",
  gray: "#D4DAE3",
  black: "#000000",
};

const daily = [
  ["05/15", 243, 114], ["05/16", 180, 71], ["05/17", 130, 50],
  ["05/18", 143, 54], ["05/19", 112, 34], ["05/20", 98, 21],
  ["05/21", 119, 24], ["05/22", 107, 21], ["05/23", 93, 19],
  ["05/24", 97, 10], ["05/25", 46, 1], ["05/26", 52, 3],
  ["05/27", 50, 2], ["05/28", 26, 6], ["05/29", 35, 0],
  ["05/30", 46, 9], ["05/31", 40, 7], ["06/01", 19, 1],
  ["06/02", 23, 1], ["06/03", 35, 0], ["06/04", 28, 3],
  ["06/05", 30, 0], ["06/06", 32, 1], ["06/07", 34, 1],
  ["06/08", 29, 0], ["06/09", 39, 2], ["06/10", 45, 1],
  ["06/11", 28, 0], ["06/12", 34, 2], ["06/13", 16, 1],
  ["06/14", 16, 0],
];

const stages = [
  ["上机初期/0.45%", "05/09-05/18", 10, 164.3, 77.3, "断头、C切均高位波动"],
  ["过渡试调", "05/19-05/24", 6, 104.3, 21.5, "C切下降，断头仍未稳定达标"],
  ["0.60%整体执行", "05/25-06/14", 21, 33.5, 2.0, "连续21天均达标"],
];

function addText(slide, text, x, y, w, h, opts = {}) {
  const sh = slide.shapes.add({
    geometry: "textbox",
    name: opts.name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  sh.text = text;
  sh.text.style = {
    typeface: opts.typeface || FONT,
    fontSize: opts.size ?? 18,
    bold: opts.bold ?? false,
    color: opts.color || C.ink,
    alignment: opts.align || "left",
    verticalAlignment: opts.valign || "top",
    lineSpacing: opts.lineSpacing ?? 1.12,
    insets: opts.insets || { top: 0, right: 0, bottom: 0, left: 0 },
    autoFit: opts.autoFit || "none",
  };
  return sh;
}

function addBox(slide, x, y, w, h, opts = {}) {
  return slide.shapes.add({
    geometry: opts.geometry || "roundRect",
    name: opts.name,
    position: { left: x, top: y, width: w, height: h },
    fill: opts.fill || C.white,
    line: opts.line || { style: "solid", fill: opts.stroke || C.light, width: opts.lineWidth ?? 1 },
    borderRadius: opts.radius ?? 10,
    shadow: opts.shadow || "shadow-none",
  });
}

function addSlideTitle(slide, no, title, kicker = "QCC成果报告") {
  slide.background.fill = C.bg;
  addText(slide, `0${no}`.slice(-2), M, 30, 46, 28, { size: 16, bold: true, color: C.cyan, align: "left" });
  addText(slide, kicker, M + 50, 30, 180, 28, { size: 14, bold: true, color: C.muted });
  addText(slide, title, M, 62, 850, 48, { size: 36, bold: true, color: C.ink });
  slide.shapes.add({
    geometry: "line",
    position: { left: M, top: 118, width: W - M * 2, height: 0 },
    fill: "none",
    line: { style: "solid", fill: C.gray, width: 1 },
  });
}

function addFooter(slide, no) {
  addText(slide, "长空之王QC小组｜提高55/72黑丝优等率", M, 682, 520, 22, { size: 12, color: C.muted });
  addText(slide, String(no).padStart(2, "0"), W - M - 50, 682, 50, 22, { size: 12, color: C.muted, align: "right" });
}

function addTag(slide, text, x, y, w, color = C.blue, fill = C.white) {
  const box = addBox(slide, x, y, w, 34, { fill, stroke: color, radius: 17 });
  addText(slide, text, x + 12, y + 8, w - 24, 18, { size: 14, bold: true, color, align: "center", valign: "middle" });
  return box;
}

function addMetric(slide, label, value, note, x, y, w, color) {
  addBox(slide, x, y, w, 136, { fill: C.white, stroke: C.light, radius: 12, shadow: "shadow-sm" });
  addText(slide, label, x + 20, y + 18, w - 40, 22, { size: 16, bold: true, color: C.muted });
  addText(slide, value, x + 20, y + 46, w - 40, 50, { size: 40, bold: true, color });
  addText(slide, note, x + 20, y + 98, w - 40, 30, { size: 14, color: C.muted });
}

function addBullets(slide, items, x, y, w, h, opts = {}) {
  const sh = slide.shapes.add({
    geometry: "textbox",
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  sh.text = items.map((item) => ({
    bulletCharacter: opts.marker || "•",
    marginLeft: opts.marginLeft ?? 22,
    indent: opts.indent ?? -12,
    spaceAfter: opts.spaceAfter ?? 7,
    runs: Array.isArray(item) ? item : [item],
  }));
  sh.text.style = {
    typeface: FONT,
    fontSize: opts.size ?? 18,
    color: opts.color || C.ink,
    lineSpacing: opts.lineSpacing ?? 1.12,
    autoFit: "none",
  };
  return sh;
}

function styleTable(table, rows, cols, headerFill = C.navy) {
  table.borders.assign({ style: "solid", fill: C.gray, width: 1 });
  table.cells.block({ row: 0, column: 0, rowCount: 1, columnCount: cols }).assign({
    fill: headerFill,
    textStyle: { color: C.white, bold: true, fontSize: 14, typeface: FONT },
    margins: { left: 8, right: 8, top: 6, bottom: 6 },
  });
  if (rows > 1) {
    table.cells.block({ row: 1, column: 0, rowCount: rows - 1, columnCount: cols }).assign({
      fill: C.white,
      textStyle: { color: C.ink, fontSize: 13, typeface: FONT },
      margins: { left: 8, right: 8, top: 5, bottom: 5 },
      anchor: "middle",
    });
  }
}

function addConclusion(slide, text, x, y, w, color = C.green) {
  addBox(slide, x, y, w, 66, { fill: color === C.green ? C.green2 : C.orange2, stroke: color, radius: 8 });
  addText(slide, "结论", x + 18, y + 14, 54, 24, { size: 18, bold: true, color });
  addText(slide, text, x + 78, y + 13, w - 96, 38, { size: 17, bold: true, color: C.ink, lineSpacing: 1.08 });
}

function addSource(slide, text, x, y, w) {
  addText(slide, text, x, y, w, 18, { size: 10.5, color: C.muted });
}

function addFlow(slide, labels, x, y, w, h) {
  const gap = 14;
  const itemW = (w - gap * (labels.length - 1)) / labels.length;
  let prev = null;
  labels.forEach((label, i) => {
    const bx = x + i * (itemW + gap);
    const box = addBox(slide, bx, y, itemW, h, { fill: i === 0 ? C.red2 : (i === labels.length - 1 ? C.green2 : C.white), stroke: i === 0 ? C.red : (i === labels.length - 1 ? C.green : C.gray), radius: 10 });
    addText(slide, label, bx + 12, y + 14, itemW - 24, h - 26, { size: 15, bold: true, color: C.ink, align: "center", valign: "middle", lineSpacing: 1.05 });
    if (prev) {
      slide.shapes.connect(prev, box, {
        kind: "straight",
        fromSide: "right",
        toSide: "left",
        line: { style: "solid", fill: C.muted, width: 1.5 },
        head: { type: "triangle", width: "sm", length: "sm" },
      });
    }
    prev = box;
  });
}

function addTimeline(slide, events, x, y, w) {
  const gap = w / (events.length - 1);
  slide.shapes.add({
    geometry: "line",
    position: { left: x, top: y + 54, width: w, height: 0 },
    fill: "none",
    line: { style: "solid", fill: C.gray, width: 3 },
  });
  events.forEach((e, i) => {
    const cx = x + gap * i;
    slide.shapes.add({
      geometry: "ellipse",
      position: { left: cx - 12, top: y + 42, width: 24, height: 24 },
      fill: e.good ? C.green : (e.warn ? C.orange : C.blue),
      line: { style: "solid", fill: C.white, width: 2 },
    });
    addText(slide, e.date, cx - 54, y, 108, 22, { size: 15, bold: true, color: e.good ? C.green : (e.warn ? C.orange : C.blue), align: "center" });
    addText(slide, e.text, cx - 78, y + 78, 156, 72, { size: 14, color: C.ink, align: "center", lineSpacing: 1.05 });
  });
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  await fs.mkdir(PREVIEW, { recursive: true });
  await fs.mkdir(LAYOUT, { recursive: true });
  await fs.mkdir(QA, { recursive: true });
  await fs.mkdir("/Users/USER/Documents/Codex/2026-06-21/new-chat-3/outputs", { recursive: true });

  const p = Presentation.create({ slideSize: { width: W, height: H } });

  // Slide 1
  {
    const s = p.slides.add();
    s.background.fill = C.bg;
    addBox(s, 0, 0, W, 720, { geometry: "rect", fill: C.bg, stroke: C.bg, radius: 0 });
    addBox(s, 0, 0, W, 94, { geometry: "rect", fill: C.navy, stroke: C.navy, radius: 0 });
    addText(s, "新凤鸣集团｜问题解决型 QCC", M, 30, 520, 26, { size: 18, bold: true, color: C.white });
    addTag(s, "QCC成果报告", 1000, 28, 164, C.white, "none");
    addText(s, "提高55/72黑丝优等率", M, 188, 860, 78, { size: 54, bold: true, color: C.ink });
    addText(s, "长空之王QC小组", M, 282, 360, 34, { size: 26, bold: true, color: C.blue });
    addText(s, "10#机｜DTY 55/72黑丝｜活动周期：2026年5月9日-6月14日", M, 330, 760, 30, { size: 19, color: C.muted });
    addFlow(s, ["POY油剂控制不足", "POY含油率偏低", "DTY加工稳定性下降", "断头/C切偏多", "优等率低"], M, 438, W - M * 2, 84);
    addConclusion(s, "本报告仅将DTY端调整作为前期排查项；最终要因为POY油剂控制不足导致POY含油率偏低。", M, 584, W - M * 2, C.orange);
  }

  // Slide 2
  {
    const s = p.slides.add(); addSlideTitle(s, 2, "小组与课题概况"); addFooter(s, 2);
    const values = [
      ["项目", "内容", "项目", "内容"],
      ["小组名称", "长空之王QC小组", "QC类型", "问题解决型"],
      ["课题名称", "提高55/72黑丝优等率", "机台/规格", "10#机，DTY 55/72黑丝"],
      ["组长", "褚新良", "活动周期", "2026年5月9日-6月14日"],
      ["组员", "陈杰、曹子建、沈金金、沈佩东、沈小燕、何迪柯、沈忠韬、宋利梁、沈琴玉", "经济效益", "平台填报20万元；按公司财务核算口径确认"],
    ];
    const table = s.tables.add({ rows: values.length, columns: 4, left: M, top: 150, width: W - M * 2, height: 250, values, columnTracks: [{ mode: "fixed", value: 112 }, { mode: "fr", value: 2 }, { mode: "fixed", value: 112 }, { mode: "fr", value: 1.5 }] });
    styleTable(table, values.length, 4);
    table.cells.block({ row: 1, column: 0, rowCount: 4, columnCount: 1 }).assign({ fill: C.lighter, textStyle: { bold: true, color: C.navy, fontSize: 13 } });
    table.cells.block({ row: 1, column: 2, rowCount: 4, columnCount: 1 }).assign({ fill: C.lighter, textStyle: { bold: true, color: C.navy, fontSize: 13 } });
    addBox(s, M, 438, W - M * 2, 138, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "课题主线", M + 24, 460, 160, 28, { size: 22, bold: true, color: C.blue });
    addBullets(s, [
      "55/72黑丝上机后生产不稳定，断头和C切偏多，优等率不足60%。",
      "经车速下调、张力CV报警阈值调整、冷却板调整等DTY端排查后，问题未得到有效解决。",
      "最终确认主要原因为配套POY原丝含油率偏低，需从前道POY油剂控制进行改善。"
    ], M + 196, 456, 900, 100, { size: 17, spaceAfter: 6 });
  }

  // Slide 3
  {
    const s = p.slides.add(); addSlideTitle(s, 3, "课题选择：现状与必要性"); addFooter(s, 3);
    addMetric(s, "上机以来优等率", "不足60%", "课题目标期望：85%以上", M, 148, 250, C.red);
    addMetric(s, "日均断头", "90个", "平台现状描述口径", M + 278, 148, 250, C.orange);
    addMetric(s, "日均C切", "50个", "平台现状描述口径", M + 556, 148, 250, C.orange);
    addMetric(s, "经济效益", "20万元", "按公司财务口径核算", M + 834, 148, 250, C.green);
    addBox(s, M, 332, 542, 190, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "选题理由", M + 24, 356, 150, 28, { size: 24, bold: true, color: C.blue });
    addBullets(s, [
      "断头多、C切多直接影响55/72黑丝质量稳定性和优等率。",
      "问题集中在10#机新规格上机初期，现场急需形成稳定控制方法。",
      "改善结果可沉淀为同类敏感规格的前后道联动经验。"
    ], M + 28, 402, 480, 94, { size: 17 });
    addBox(s, M + 594, 332, 528, 190, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "问题总结", M + 618, 356, 150, 28, { size: 24, bold: true, color: C.red });
    addFlow(s, ["断头多", "张力波动", "C切多", "优等率低"], M + 626, 420, 456, 64);
    addConclusion(s, "课题聚焦“提高55/72黑丝优等率”，目标小、问题明确，符合新凤鸣QC项目选题要求。", M, 572, W - M * 2, C.green);
  }

  // Slide 4
  {
    const s = p.slides.add(); addSlideTitle(s, 4, "目标设定与可行性依据"); addFooter(s, 4);
    addBox(s, M, 148, 530, 374, { fill: C.white, stroke: C.light, radius: 12 });
    addText(s, "SMART目标", M + 28, 174, 220, 32, { size: 25, bold: true, color: C.blue });
    const values = [
      ["指标", "现状/问题", "目标值"],
      ["断头数", "上机初期高位波动", "≤90个/日"],
      ["C切数", "张力波动造成C切偏多", "≤50个/日"],
      ["优等率", "不足60%", "≥85%"],
      ["期限", "2026年5月9日启动", "6月14日前完成验证"],
    ];
    const table = s.tables.add({ rows: values.length, columns: 3, left: M + 28, top: 224, width: 474, height: 238, values, columnTracks: [{ mode: "fixed", value: 110 }, { mode: "fr", value: 1.5 }, { mode: "fr", value: 1 }] });
    styleTable(table, values.length, 3, C.blue);
    table.cells.block({ row: 1, column: 2, rowCount: 4, columnCount: 1 }).assign({ fill: C.green2, textStyle: { bold: true, color: C.green, fontSize: 14 } });
    addSource(s, "依据：XFM03-QC(2022B)-32要求目标符合SMART原则。", M + 30, 478, 440);
    addBox(s, M + 586, 148, 536, 374, { fill: C.white, stroke: C.light, radius: 12 });
    addText(s, "目标可行性", M + 614, 174, 220, 32, { size: 25, bold: true, color: C.green });
    addBullets(s, [
      "关键过程指标每日可采集，断头、C切可直接量化跟踪。",
      "0.55%试调阶段C切已有明显下降，说明POY含油率方向具备改善可能。",
      "0.60%整体执行后连续21天达标，满足至少1个月效果验证的核心要求。",
      "优等率缺少逐日原始数据，本报告不虚构改善后优等率，以过程目标完成作为支撑。"
    ], M + 618, 224, 472, 214, { size: 17, spaceAfter: 7 });
    addConclusion(s, "目标设定采用“断头+C切+优等率”组合，其中断头和C切作为可量化过程控制指标。", M + 614, 456, 480, C.green);
  }

  // Slide 5
  {
    const s = p.slides.add(); addSlideTitle(s, 5, "现状调查：分阶段数据"); addFooter(s, 5);
    const values = [
      ["阶段", "日期范围", "天数", "断头日均", "C切日均", "结论"],
      ...stages.map((r) => [r[0], r[1], r[2], `${r[3]}个`, `${r[4]}个`, r[5]]),
    ];
    const table = s.tables.add({ rows: values.length, columns: 6, left: M, top: 150, width: W - M * 2, height: 188, values, columnTracks: [{ mode: "fr", value: 1.6 }, { mode: "fixed", value: 140 }, { mode: "fixed", value: 70 }, { mode: "fixed", value: 104 }, { mode: "fixed", value: 104 }, { mode: "fr", value: 2 }] });
    styleTable(table, values.length, 6);
    table.cells.block({ row: 3, column: 0, rowCount: 1, columnCount: 6 }).assign({ fill: C.green2, textStyle: { color: C.ink, fontSize: 13, bold: true } });
    s.charts.add("bar", {
      position: { left: M + 16, top: 380, width: 520, height: 230 },
      categories: stages.map((r) => r[0]),
      series: [
        { name: "断头日均", values: stages.map((r) => r[3]), fill: C.blue },
        { name: "C切日均", values: stages.map((r) => r[4]), fill: C.orange },
      ],
      hasLegend: true,
      legend: { position: "bottom", textStyle: { fill: C.muted, fontSize: 12 } },
      barOptions: { direction: "column", grouping: "clustered", gapWidth: 55 },
      yAxis: { min: 0, max: 180, majorUnit: 30, majorGridlines: { style: "solid", fill: C.light, width: 1 }, textStyle: { fill: C.muted, fontSize: 11 } },
      xAxis: { textStyle: { fill: C.muted, fontSize: 10 } },
      dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 11, bold: true } },
      chartFill: C.white,
      plotAreaFill: C.white,
    });
    addBox(s, M + 600, 376, 522, 238, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "现状判断", M + 628, 400, 180, 28, { size: 24, bold: true, color: C.red });
    addBullets(s, [
      "0.45%阶段断头日均164.3个、C切日均77.3个，均超过目标线。",
      "0.55%及0.60%打样过渡期C切下降，但断头仍未稳定达标。",
      "0.60%整体执行后，断头日均33.5个、C切日均2.0个，过程状态趋稳。"
    ], M + 628, 446, 460, 100, { size: 17 });
    addSource(s, "数据来源：10#机55/72黑丝断头、C切数量趋势图；交接资料汇总。", M + 628, 568, 450);
  }

  // Slide 6
  {
    const s = p.slides.add(); addSlideTitle(s, 6, "前期尝试：DTY端排查均未解决"); addFooter(s, 6);
    const values = [
      ["项目", "尝试内容", "结论", "报告定位"],
      ["二罗拉速度", "650m/min降至620m/min", "无明显效果", "前期尝试"],
      ["张力CV报警阈值", "报警阈值调至8；瞬间最高约20", "无明显效果", "异常现象/排查"],
      ["冷却板", "调整部分锭位冷却板", "无明显效果", "前期尝试"],
      ["DTY端油相关排查", "油轮、导丝器等早期排查", "不作为要因", "排除项"],
    ];
    const table = s.tables.add({ rows: values.length, columns: 4, left: M, top: 156, width: W - M * 2, height: 272, values, columnTracks: [{ mode: "fixed", value: 180 }, { mode: "fr", value: 1.9 }, { mode: "fixed", value: 140 }, { mode: "fixed", value: 140 }] });
    styleTable(table, values.length, 4, C.red);
    table.cells.block({ row: 1, column: 2, rowCount: 4, columnCount: 1 }).assign({ fill: C.red2, textStyle: { color: C.red, bold: true, fontSize: 13 } });
    addConclusion(s, "上述内容只作为上机后前期排查和排除项，不作为最终要因，也不作为有效对策。", M, 464, W - M * 2, C.orange);
    addBox(s, M, 558, W - M * 2, 66, { fill: C.white, stroke: C.light, radius: 8 });
    addText(s, "分析转向", M + 22, 576, 110, 24, { size: 20, bold: true, color: C.blue });
    addText(s, "DTY端单项调整未能消除张力CV瞬间冲高、断头和C切异常，课题进一步追溯前道POY油剂控制和原丝含油率。", M + 132, 575, 920, 28, { size: 18, color: C.ink });
  }

  // Slide 7
  {
    const s = p.slides.add(); addSlideTitle(s, 7, "原因分析：锁定POY含油率主线"); addFooter(s, 7);
    addBox(s, M, 146, W - M * 2, 118, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "问题链路", M + 24, 170, 130, 28, { size: 24, bold: true, color: C.blue });
    addFlow(s, ["POY油剂控制不足", "POY原丝含油率偏低", "润滑性/抱合性不足", "张力CV瞬间冲高", "断头/C切偏多"], M + 170, 164, 910, 66);
    addBox(s, M, 310, 540, 216, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "5Why逻辑", M + 24, 334, 160, 28, { size: 24, bold: true, color: C.green });
    addBullets(s, [
      "为什么优等率低？断头、C切偏多。",
      "为什么断头/C切偏多？DTY加工稳定性差，张力CV瞬间冲高。",
      "为什么加工稳定性差？丝束润滑性、抱合性不足。",
      "为什么润滑性不足？配套POY原丝含油率偏低。",
      "为什么含油率偏低？POY油剂控制不足。"
    ], M + 28, 382, 484, 118, { size: 16, spaceAfter: 4 });
    addBox(s, M + 600, 310, 522, 216, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "末端因素", M + 628, 334, 160, 28, { size: 24, bold: true, color: C.orange });
    addText(s, "POY油剂控制不足，导致55/72黑丝配套POY原丝含油率偏低。", M + 628, 386, 450, 58, { size: 24, bold: true, color: C.ink, lineSpacing: 1.08 });
    addText(s, "更完整表述：配套POY原丝含油率偏低，后道DTY加工稳定性下降，造成断头及C切偏多。", M + 628, 464, 450, 40, { size: 16, color: C.muted });
    addConclusion(s, "最终要因必须表述为POY油剂控制不足/POY含油率偏低，不得归因于DTY端油相关排查项。", M, 572, W - M * 2, C.orange);
  }

  // Slide 8
  {
    const s = p.slides.add(); addSlideTitle(s, 8, "要因确认：用阶段数据验证"); addFooter(s, 8);
    s.charts.add("bar", {
      position: { left: M, top: 152, width: 520, height: 340 },
      categories: ["0.45%阶段", "过渡试调", "0.60%整体执行"],
      series: [{ name: "断头日均", values: [164.3, 104.3, 33.5], fill: C.blue }],
      hasLegend: false,
      barOptions: { direction: "column", grouping: "clustered", gapWidth: 55 },
      yAxis: { min: 0, max: 180, majorUnit: 30, majorGridlines: { style: "solid", fill: C.light, width: 1 }, textStyle: { fill: C.muted, fontSize: 12 } },
      xAxis: { textStyle: { fill: C.muted, fontSize: 12 } },
      dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 13, bold: true } },
      chartFill: C.white,
      plotAreaFill: C.white,
    });
    s.charts.add("bar", {
      position: { left: M + 566, top: 152, width: 520, height: 340 },
      categories: ["0.45%阶段", "过渡试调", "0.60%整体执行"],
      series: [{ name: "C切日均", values: [77.3, 21.5, 2.0], fill: C.orange }],
      hasLegend: false,
      barOptions: { direction: "column", grouping: "clustered", gapWidth: 55 },
      yAxis: { min: 0, max: 90, majorUnit: 15, majorGridlines: { style: "solid", fill: C.light, width: 1 }, textStyle: { fill: C.muted, fontSize: 12 } },
      xAxis: { textStyle: { fill: C.muted, fontSize: 12 } },
      dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 13, bold: true } },
      chartFill: C.white,
      plotAreaFill: C.white,
    });
    addTag(s, "断头目标 ≤90个/日", M + 176, 504, 190, C.green, C.green2);
    addTag(s, "C切目标 ≤50个/日", M + 752, 504, 180, C.green, C.green2);
    addConclusion(s, "0.60%整体执行后，断头与C切均明显低于目标线，验证POY含油率控制方向有效。", M, 584, W - M * 2, C.green);
  }

  // Slide 9
  {
    const s = p.slides.add(); addSlideTitle(s, 9, "制定对策：5W1H闭环"); addFooter(s, 9);
    const values = [
      ["5W1H", "本项目对策"],
      ["What", "优化55/72黑丝配套POY油剂控制，提高POY原丝含油率稳定性。"],
      ["Why", "控制断头≤90个/日、C切≤50个/日，为优等率提升至85%以上提供支撑。"],
      ["Who", "前纺工艺、DTY工艺、质量、QC小组成员协同负责。"],
      ["Where", "前道POY工序及10#机DTY 55/72黑丝验证现场。"],
      ["When", "5月19日0.55%打样，5月21日0.60%打样，5月25日0.60%整体执行。"],
      ["How", "取样检测POY含油率；优化油剂添加控制；上机验证断头、C切、张力CV；固化控制标准。"],
    ];
    const table = s.tables.add({ rows: values.length, columns: 2, left: M, top: 150, width: W - M * 2, height: 390, values, columnTracks: [{ mode: "fixed", value: 120 }, { mode: "fr", value: 1 }] });
    styleTable(table, values.length, 2, C.blue);
    table.cells.block({ row: 1, column: 0, rowCount: 6, columnCount: 1 }).assign({ fill: C.lighter, textStyle: { bold: true, color: C.blue, fontSize: 14 } });
    addConclusion(s, "对策仅针对主要要因POY油剂控制制定；DTY端调整不纳入最终有效对策。", M, 582, W - M * 2, C.orange);
  }

  // Slide 10
  {
    const s = p.slides.add(); addSlideTitle(s, 10, "实施对策：含油率调整与上机验证"); addFooter(s, 10);
    addTimeline(s, [
      { date: "05/09", text: "55/72黑丝上机，断头/C切异常，启动排查", warn: true },
      { date: "05/19", text: "POY含油率0.45%调至0.55%，10个锭位打样", warn: true },
      { date: "05/21", text: "继续将POY含油率调至0.60%，10个锭位打样" },
      { date: "05/22", text: "0.55%阶段整体执行，有改善但断头不稳定", warn: true },
      { date: "05/25", text: "0.60%整体执行，断头和C切趋于稳定", good: true },
      { date: "06/14", text: "完成连续21天目标跟踪验证", good: true },
    ], M + 20, 166, W - M * 2 - 40);
    addBox(s, M, 402, 530, 150, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "实施重点", M + 24, 426, 140, 28, { size: 23, bold: true, color: C.blue });
    addBullets(s, [
      "含油率控制从0.45%逐步验证至0.60%。",
      "每次调整后以DTY端断头、C切和张力CV作为效果反馈。",
      "0.55%阶段不作为最终有效措施，0.60%整体执行作为有效对策验证口径。"
    ], M + 28, 470, 470, 58, { size: 16, spaceAfter: 3 });
    addBox(s, M + 594, 402, 528, 150, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "执行边界", M + 618, 426, 140, 28, { size: 23, bold: true, color: C.orange });
    addBullets(s, [
      "0.60%控制值写作“按工艺文件/通知单最终确认值为准”。",
      "换批、换规格或油剂系统调整后需先小批量打样验证。",
      "异常时优先追溯POY含油率和前纺执行情况。"
    ], M + 622, 470, 466, 58, { size: 16, spaceAfter: 3 });
  }

  // Slide 11
  {
    const s = p.slides.add(); addSlideTitle(s, 11, "效果检查：近一个月每日目标跟踪"); addFooter(s, 11);
    const cats = daily.map((d) => d[0]);
    s.charts.add("line", {
      position: { left: M, top: 150, width: W - M * 2, height: 352 },
      categories: cats,
      series: [
        { name: "实际断头", values: daily.map((d) => d[1]), line: { style: "solid", fill: C.blue, width: 2.5 }, marker: { symbol: "circle", size: 4 } },
        { name: "断头目标90", values: daily.map(() => 90), line: { style: "dashed", fill: C.green, width: 2 }, marker: { symbol: "none" } },
        { name: "实际C切", values: daily.map((d) => d[2]), line: { style: "solid", fill: C.orange, width: 2.5 }, marker: { symbol: "circle", size: 4 } },
        { name: "C切目标50", values: daily.map(() => 50), line: { style: "dashed", fill: C.red, width: 2 }, marker: { symbol: "none" } },
      ],
      hasLegend: true,
      legend: { position: "bottom", textStyle: { fill: C.muted, fontSize: 12 } },
      yAxis: { min: 0, max: 260, majorUnit: 50, majorGridlines: { style: "solid", fill: C.light, width: 1 }, textStyle: { fill: C.muted, fontSize: 11 } },
      xAxis: { textStyle: { fill: C.muted, fontSize: 8 } },
      lineOptions: { smooth: false },
      chartFill: C.white,
      plotAreaFill: C.white,
    });
    addMetric(s, "综合达标天数", "21/31天", "近一个月跟踪", M, 536, 250, C.green);
    addMetric(s, "0.60%执行后", "21/21天", "连续全部达标", M + 278, 536, 250, C.green);
    addMetric(s, "最大日断头", "52个", "低于90个目标线", M + 556, 536, 250, C.blue);
    addMetric(s, "最大日C切", "9个", "低于50个目标线", M + 834, 536, 250, C.orange);
  }

  // Slide 12
  {
    const s = p.slides.add(); addSlideTitle(s, 12, "目标完成情况：过程指标显著改善"); addFooter(s, 12);
    addMetric(s, "改善后断头日均", "33.5个", "目标≤90个/日", M, 150, 260, C.green);
    addMetric(s, "改善后C切日均", "2.0个", "目标≤50个/日", M + 296, 150, 260, C.green);
    addMetric(s, "断头下降幅度", "79.6%", "较0.45%上机初期", M + 592, 150, 260, C.blue);
    addMetric(s, "C切下降幅度", "97.5%", "较0.45%上机初期", M + 888, 150, 260, C.blue);
    const values = [
      ["对比口径", "断头日均", "C切日均"],
      ["上机初期0.45%阶段", "164.3个", "77.3个"],
      ["0.60%整体执行后", "33.5个", "2.0个"],
      ["下降幅度", "79.6%", "97.5%"],
    ];
    const table = s.tables.add({ rows: values.length, columns: 3, left: M, top: 340, width: 510, height: 168, values, columnTracks: [{ mode: "fr", value: 1.5 }, { mode: "fixed", value: 140 }, { mode: "fixed", value: 140 }] });
    styleTable(table, values.length, 3, C.green);
    table.cells.block({ row: 3, column: 1, rowCount: 1, columnCount: 2 }).assign({ fill: C.green2, textStyle: { color: C.green, bold: true, fontSize: 15 } });
    addBox(s, M + 588, 340, 534, 168, { fill: C.white, stroke: C.light, radius: 10 });
    addText(s, "目标完成说明", M + 616, 364, 200, 28, { size: 24, bold: true, color: C.blue });
    addBullets(s, [
      "5月25日至6月14日连续21天断头和C切全部达标。",
      "断头、C切两项关键过程指标已完成对策目标。",
      "由于缺少逐日/逐批优等率原始数据，本报告不虚构改善后优等率。"
    ], M + 620, 408, 460, 70, { size: 16, spaceAfter: 4 });
    addConclusion(s, "过程目标完成，为55/72黑丝优等率提升至85%以上提供数据支撑。", M, 568, W - M * 2, C.green);
  }

  // Slide 13
  {
    const s = p.slides.add(); addSlideTitle(s, 13, "标准化与防护措施"); addFooter(s, 13);
    const values = [
      ["序号", "标准化/防护项目", "控制要求", "记录/验证方式"],
      ["1", "POY油剂控制标准固化", "55/72黑丝配套POY含油率控制值按0.60%执行，允许范围以工艺文件/通知单为准。", "POY含油率检测记录、工艺通知单"],
      ["2", "POY含油率首批确认", "上机前确认批号、含油率检测结果和工艺调整状态，未经确认不得批量转入DTY验证。", "首批确认记录、交接记录"],
      ["3", "DTY关键过程指标跟踪", "每日跟踪断头、C切和张力CV异常；断头>90或C切>50时启动异常反馈。", "每日目标跟踪表"],
      ["4", "异常联动处置", "优先追溯配套POY含油率、油剂批次和前纺执行情况；DTY端调整仅作为排查项。", "异常分析单、纠正措施记录"],
      ["5", "换批/换规格防护", "油剂系统或工艺参数调整后，先小批量打样验证，稳定后再整体推行。", "小样验证记录"],
    ];
    const table = s.tables.add({ rows: values.length, columns: 4, left: M, top: 146, width: W - M * 2, height: 370, values, columnTracks: [{ mode: "fixed", value: 56 }, { mode: "fixed", value: 178 }, { mode: "fr", value: 2.2 }, { mode: "fr", value: 1 }] });
    styleTable(table, values.length, 4, C.navy);
    table.cells.block({ row: 1, column: 0, rowCount: 5, columnCount: 1 }).assign({ fill: C.lighter, textStyle: { bold: true, color: C.navy, fontSize: 13 } });
    addConclusion(s, "标准化重点：POY油剂控制标准化 + 上机前确认 + DTY过程日跟踪 + 异常联动处置。", M, 568, W - M * 2, C.green);
  }

  // Slide 14
  {
    const s = p.slides.add(); addSlideTitle(s, 14, "总结与下一步"); addFooter(s, 14);
    addBox(s, M, 150, 345, 300, { fill: C.white, stroke: C.light, radius: 12 });
    addText(s, "本次成果", M + 24, 176, 160, 30, { size: 25, bold: true, color: C.green });
    addBullets(s, [
      "明确最终要因：POY油剂控制不足导致POY含油率偏低。",
      "0.60%整体执行后，连续21天断头和C切全部达标。",
      "形成55/72黑丝配套POY油剂控制经验。"
    ], M + 28, 226, 290, 142, { size: 17, spaceAfter: 7 });
    addBox(s, M + 410, 150, 345, 300, { fill: C.white, stroke: C.light, radius: 12 });
    addText(s, "经验沉淀", M + 434, 176, 160, 30, { size: 25, bold: true, color: C.blue });
    addBullets(s, [
      "前期DTY端尝试不作为要因和最终有效措施。",
      "问题分析需前后道联动，避免局部调整反复试错。",
      "过程目标数据要每日记录，支撑评审闭环。"
    ], M + 438, 226, 290, 142, { size: 17, spaceAfter: 7 });
    addBox(s, M + 820, 150, 302, 300, { fill: C.white, stroke: C.light, radius: 12 });
    addText(s, "下一步", M + 844, 176, 140, 30, { size: 25, bold: true, color: C.orange });
    addBullets(s, [
      "继续跟踪后续批次断头、C切和优等率。",
      "将经验扩展到同类细旦黑丝/高敏感规格。",
      "按公司财务口径确认最终经济效益。"
    ], M + 848, 226, 246, 142, { size: 17, spaceAfter: 7 });
    addBox(s, M, 514, W - M * 2, 88, { fill: C.navy, stroke: C.navy, radius: 10 });
    addText(s, "统一结论", M + 28, 538, 132, 28, { size: 23, bold: true, color: C.white });
    addText(s, "通过优化POY油剂控制、提高配套POY含油率，55/72黑丝DTY加工稳定性明显提升；断头、C切目标完成，具备标准化和推广基础。", M + 170, 532, 890, 44, { size: 19, bold: true, color: C.white, lineSpacing: 1.08 });
  }

  for (const [idx, slide] of p.slides.items.entries()) {
    const stem = `slide-${String(idx + 1).padStart(2, "0")}`;
    await writeBlob(`${PREVIEW}/${stem}.png`, await p.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${LAYOUT}/${stem}.json`, await layout.text());
  }
  await writeBlob(`${PREVIEW}/deck-montage.webp`, await p.export({ format: "webp", montage: true, scale: 1 }));
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(OUT);
  const inspect = await p.inspect({ kind: "slide,textbox,shape,table,chart", maxChars: 20000 });
  await fs.writeFile(`${QA}/inspect.ndjson`, inspect.ndjson);
  console.log(OUT);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
