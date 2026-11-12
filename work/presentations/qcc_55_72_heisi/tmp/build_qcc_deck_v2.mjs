import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const FINAL = "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/outputs/长空之王QC小组_QCC成果报告_提高55-72黑丝优等率_视觉重做版.pptx";
const TMP = "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/work/presentations/qcc_55_72_heisi/tmp";
const BG_PATH = `${TMP}/assets/qcc-fiber-bg-imagegen.png`;
const PHOTO_INITIAL = `${TMP}/assets/photo-initial-machine.jpg`;
const PHOTO_055 = `${TMP}/assets/photo-0.55-oil.jpg`;
const PHOTO_060 = `${TMP}/assets/photo-0.60-stable.jpg`;
const CV_WAVE_01 = `${TMP}/assets/cv-wave-01-crop.jpg`;
const CV_WAVE_02 = `${TMP}/assets/cv-wave-02-crop.jpg`;
const CV_WAVE_03 = `${TMP}/assets/cv-wave-03-crop.jpg`;
const CV_WAVE_04 = `${TMP}/assets/cv-wave-04-crop.jpg`;
const PREVIEW = `${TMP}/preview_v2`;
const LAYOUT = `${TMP}/layout_v2`;
const QA = `${TMP}/qa`;

const W = 1280;
const H = 720;
const M = 62;
const FONT = "Microsoft YaHei";

const C = {
  paper: "#F4F6F8",
  white: "#FFFFFF",
  ink: "#102033",
  muted: "#667386",
  navy: "#0A2742",
  navy2: "#123A5A",
  blue: "#1E6AA8",
  teal: "#0E9A9A",
  mint: "#DDF3EC",
  amber: "#D99116",
  amber2: "#FFF3DA",
  red: "#B9403A",
  red2: "#F9E3E0",
  green: "#177A51",
  line: "#D8DEE8",
  soft: "#EAF0F6",
  darkGlass: "#081B2C",
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
  ["0.45%阶段", "05/09-05/18", 164.3, 77.3, "断头、C切高位波动"],
  ["过渡试调", "05/19-05/24", 104.3, 21.5, "C切下降，断头仍未稳定"],
  ["0.60%整体执行", "05/25-06/14", 33.5, 2.0, "连续21天达标"],
];

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function text(slide, value, x, y, w, h, o = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    name: o.name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = value;
  s.text.style = {
    typeface: o.typeface || FONT,
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

function box(slide, x, y, w, h, o = {}) {
  const geometry = o.geometry || "rect";
  const cfg = {
    geometry: o.geometry || "rect",
    name: o.name,
    position: { left: x, top: y, width: w, height: h },
    fill: o.fill || C.white,
    line: o.line || { style: "solid", fill: o.stroke || C.line, width: o.lineWidth ?? 1 },
    shadow: o.shadow || "shadow-none",
  };
  if (["rect", "textbox", "roundRect"].includes(geometry)) cfg.borderRadius = o.radius ?? 0;
  return slide.shapes.add(cfg);
}

function pill(slide, value, x, y, w, o = {}) {
  box(slide, x, y, w, 30, { geometry: "roundRect", fill: o.fill || C.soft, stroke: o.stroke || C.line, radius: 15 });
  text(slide, value, x + 12, y + 7, w - 24, 17, { size: o.size || 13, bold: true, color: o.color || C.navy, align: "center" });
}

function bullets(slide, items, x, y, w, h, o = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = items.map((item) => ({
    bulletCharacter: o.marker || "•",
    marginLeft: o.marginLeft ?? 22,
    indent: o.indent ?? -12,
    spaceAfter: o.spaceAfter ?? 7,
    runs: Array.isArray(item) ? item : [item],
  }));
  s.text.style = {
    typeface: FONT,
    fontSize: o.size ?? 18,
    color: o.color || C.ink,
    lineSpacing: o.lineSpacing ?? 1.1,
  };
  return s;
}

function addImage(slide, bytes, x, y, w, h, o = {}) {
  return slide.images.add({
    blob: bytes,
    contentType: o.contentType || "image/png",
    alt: o.alt || "ImageGen generated black filament production background",
    fit: o.fit || "cover",
    position: { left: x, top: y, width: w, height: h },
    crop: o.crop,
    geometry: o.geometry,
    borderRadius: o.borderRadius,
  });
}

function base(slide, no, title, section = "QCC成果报告") {
  slide.background.fill = C.paper;
  box(slide, 0, 0, 18, H, { fill: C.navy, stroke: C.navy });
  box(slide, 18, 0, 5, H, { fill: C.teal, stroke: C.teal });
  text(slide, String(no).padStart(2, "0"), M, 36, 44, 22, { size: 14, bold: true, color: C.teal });
  text(slide, section, M + 54, 36, 180, 22, { size: 13, bold: true, color: C.muted });
  text(slide, title, M, 72, 900, 48, { size: 34, bold: true, color: C.ink });
  box(slide, M, 130, W - M * 2, 1, { fill: C.line, stroke: C.line });
  text(slide, "长空之王QC小组｜提高55/72黑丝优等率", M, 686, 520, 18, { size: 11, color: C.muted });
  text(slide, String(no).padStart(2, "0"), W - M - 40, 686, 40, 18, { size: 11, color: C.muted, align: "right" });
}

function callout(slide, label, value, x, y, w, h, o = {}) {
  box(slide, x, y, w, h, { fill: o.fill || C.white, stroke: o.stroke || C.line, radius: o.radius ?? 0, shadow: o.shadow || "shadow-sm" });
  text(slide, label, x + 20, y + 18, w - 40, 22, { size: 15, bold: true, color: o.labelColor || C.muted });
  text(slide, value, x + 20, y + 48, w - 40, 44, { size: o.valueSize || 34, bold: true, color: o.color || C.ink });
  if (o.note) text(slide, o.note, x + 20, y + h - 31, w - 40, 20, { size: 13, color: C.muted });
}

function conclusion(slide, msg, x, y, w, o = {}) {
  box(slide, x, y, w, 58, { fill: o.fill || C.mint, stroke: o.stroke || C.teal, radius: 0 });
  text(slide, "结论", x + 18, y + 17, 56, 22, { size: 18, bold: true, color: o.stroke || C.teal });
  text(slide, msg, x + 84, y + 15, w - 104, 28, { size: 18, bold: true, color: C.ink, lineSpacing: 1.05 });
}

function styleTable(t, rows, cols, headerFill = C.navy) {
  t.borders.assign({ style: "solid", fill: "#BFC9D6", width: 0.8 });
  t.cells.block({ row: 0, column: 0, rowCount: 1, columnCount: cols }).assign({
    fill: headerFill,
    textStyle: { color: C.white, bold: true, fontSize: 15, typeface: FONT },
    margins: { left: 9, right: 9, top: 7, bottom: 7 },
    anchor: "middle",
  });
  if (rows > 1) {
    t.cells.block({ row: 1, column: 0, rowCount: rows - 1, columnCount: cols }).assign({
      fill: C.white,
      textStyle: { color: C.ink, fontSize: 14, typeface: FONT },
      margins: { left: 9, right: 9, top: 6, bottom: 6 },
      anchor: "middle",
    });
  }
}

function stepRibbon(slide, labels, activeIndex, x, y, w) {
  const gap = 5;
  const sw = (w - gap * (labels.length - 1)) / labels.length;
  labels.forEach((l, i) => {
    const active = i === activeIndex;
    box(slide, x + i * (sw + gap), y, sw, 42, {
      fill: active ? C.navy : C.white,
      stroke: active ? C.navy : C.line,
      radius: 0,
    });
    text(slide, l, x + i * (sw + gap) + 5, y + 12, sw - 10, 18, {
      size: 12.5,
      bold: active,
      color: active ? C.white : C.muted,
      align: "center",
    });
  });
}

function processChain(slide, labels, x, y, w, h) {
  const gap = 12;
  const bw = (w - gap * (labels.length - 1)) / labels.length;
  let prev;
  labels.forEach((label, i) => {
    const fill = i === 0 ? C.red2 : i === labels.length - 1 ? C.mint : C.white;
    const stroke = i === 0 ? C.red : i === labels.length - 1 ? C.green : C.line;
    const b = box(slide, x + i * (bw + gap), y, bw, h, { fill, stroke, radius: 0, shadow: "shadow-sm" });
    text(slide, label, x + i * (bw + gap) + 12, y + 14, bw - 24, h - 20, { size: 15, bold: true, color: C.ink, align: "center", valign: "middle" });
    if (prev) {
      slide.shapes.connect(prev, b, {
        kind: "straight",
        fromSide: "right",
        toSide: "left",
        line: { style: "solid", fill: C.muted, width: 1.2 },
        head: { type: "triangle", width: "sm", length: "sm" },
      });
    }
    prev = b;
  });
}

function timeline(slide, events, x, y, w) {
  box(slide, x, y + 55, w, 2, { fill: C.line, stroke: C.line });
  const gap = w / (events.length - 1);
  events.forEach((e, i) => {
    const cx = x + i * gap;
    const col = e.good ? C.green : e.warn ? C.amber : C.blue;
    box(slide, cx - 8, y + 47, 16, 16, { geometry: "ellipse", fill: col, stroke: C.white, lineWidth: 2 });
    text(slide, e.date, cx - 50, y + 8, 100, 20, { size: 15, bold: true, color: col, align: "center" });
    text(slide, e.text, cx - 70, y + 82, 140, 74, { size: 14, color: C.ink, align: "center", lineSpacing: 1.05 });
  });
}

async function main() {
  await fs.mkdir(PREVIEW, { recursive: true });
  await fs.mkdir(LAYOUT, { recursive: true });
  await fs.mkdir(QA, { recursive: true });
  await fs.mkdir("/Users/USER/Documents/Codex/2026-06-21/new-chat-3/outputs", { recursive: true });
  const bg = await fs.readFile(BG_PATH);
  const photoInitial = await fs.readFile(PHOTO_INITIAL);
  const photo055 = await fs.readFile(PHOTO_055);
  const photo060 = await fs.readFile(PHOTO_060);
  const cvWave01 = await fs.readFile(CV_WAVE_01);
  const cvWave02 = await fs.readFile(CV_WAVE_02);
  const cvWave03 = await fs.readFile(CV_WAVE_03);
  const cvWave04 = await fs.readFile(CV_WAVE_04);
  const p = Presentation.create({ slideSize: { width: W, height: H } });

  // 1 Cover
  {
    const s = p.slides.add();
    addImage(s, bg, 0, 0, W, H, { fit: "cover" });
    box(s, 0, 0, 720, H, { fill: "#061A2BCC", stroke: "#061A2B00" });
    box(s, 0, 0, W, H, { fill: "#00000022", stroke: "#00000000" });
    text(s, "提高55/72黑丝优等率", M, 238, 740, 78, { size: 54, bold: true, color: C.white });
    text(s, "长空之王QC小组", M, 334, 360, 34, { size: 27, bold: true, color: "#5DD0CE" });
  }

  // 2 Story map
  {
    const s = p.slides.add(); base(s, 2, "报告路径：按问题解决型QCC闭环展开");
    stepRibbon(s, ["选题", "现状", "目标", "原因", "要因", "对策", "实施", "效果", "标准化", "总结"], -1, M, 156, W - M * 2);
    box(s, M, 236, 350, 270, { fill: C.navy, stroke: C.navy });
    text(s, "核心判断", M + 28, 266, 180, 32, { size: 26, bold: true, color: C.white });
    text(s, "不是DTY端单项调整解决问题，而是前道POY油剂控制改善后，DTY加工稳定性才明显提升。", M + 28, 322, 286, 118, { size: 22, bold: true, color: C.white, lineSpacing: 1.12 });
    box(s, M + 400, 236, 722, 270, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    bullets(s, [
      "模板依据：案例报告顺序采用“小组概况、选题、目标、原因分析、要因确认、对策实施、效果检查、巩固措施、总结”。",
      "制度依据：新凤鸣实施细则要求目标符合SMART原则，对策按5W1H制定，活动效果至少跟踪1个月，措施需标准化。",
      "数据边界：未提供逐日/逐批优等率原始数据，因此不虚构改善后优等率。"
    ], M + 432, 270, 650, 180, { size: 18, spaceAfter: 11 });
    conclusion(s, "本PPT优先保证QCC逻辑和公司口径正确，再做正式发布视觉包装。", M, 572, W - M * 2);
  }

  // 3 Overview
  {
    const s = p.slides.add(); base(s, 3, "小组概况与课题选择");
    const values = [
      ["项目", "内容", "项目", "内容"],
      ["小组名称", "长空之王QC小组", "QC类型", "问题解决型"],
      ["课题名称", "提高55/72黑丝优等率", "机台/规格", "10#机，DTY 55/72黑丝"],
      ["组长", "褚新良", "活动周期", "2026年5月9日-6月14日"],
      ["组员", "陈杰、曹子建、沈金金、沈佩东、沈小燕、何迪柯、沈忠韬、宋利梁、沈琴玉", "经济效益", "平台填报20万元；按公司财务口径确认"],
    ];
    const t = s.tables.add({ rows: values.length, columns: 4, left: M, top: 154, width: W - M * 2, height: 250, values, columnTracks: [{ mode: "fixed", value: 110 }, { mode: "fr", value: 2 }, { mode: "fixed", value: 110 }, { mode: "fr", value: 1.5 }] });
    styleTable(t, values.length, 4);
    t.cells.block({ row: 1, column: 0, rowCount: 4, columnCount: 1 }).assign({ fill: C.soft, textStyle: { bold: true, color: C.navy, fontSize: 14 } });
    t.cells.block({ row: 1, column: 2, rowCount: 4, columnCount: 1 }).assign({ fill: C.soft, textStyle: { bold: true, color: C.navy, fontSize: 14 } });
    callout(s, "上机以来优等率", "不足60%", M, 450, 250, 116, { color: C.red, note: "平台现状描述" });
    callout(s, "日均断头", "90个", M + 294, 450, 250, 116, { color: C.amber, note: "平台现状描述" });
    callout(s, "日均C切", "50个", M + 588, 450, 250, 116, { color: C.amber, note: "平台现状描述" });
    callout(s, "目标优等率", "≥85%", M + 882, 450, 240, 116, { color: C.green, note: "课题目标期望" });
  }

  // 4 Current state
  {
    const s = p.slides.add(); base(s, 4, "现状调查：问题集中在上机初期稳定性");
    const values = [
      ["阶段", "日期范围", "断头日均", "C切日均", "判断"],
      ...stages.map((r) => [r[0], r[1], `${r[2]}个`, `${r[3]}个`, r[4]]),
    ];
    const t = s.tables.add({ rows: values.length, columns: 5, left: M, top: 152, width: W - M * 2, height: 168, values, columnTracks: [{ mode: "fixed", value: 178 }, { mode: "fixed", value: 150 }, { mode: "fixed", value: 130 }, { mode: "fixed", value: 130 }, { mode: "fr", value: 1 }] });
    styleTable(t, values.length, 5);
    t.cells.block({ row: 3, column: 0, rowCount: 1, columnCount: 5 }).assign({ fill: C.mint, textStyle: { color: C.ink, bold: true, fontSize: 14 } });
    box(s, M, 338, W - M * 2, 44, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "调查口径", M + 20, 351, 86, 18, { size: 16, bold: true, color: C.blue });
    text(s, "对象：10#机55/72黑丝；指标：每日断头、C切、张力CV异常；分层：0.45%上机初期、0.55/0.60打样过渡、0.60整体执行。", M + 118, 351, 900, 18, { size: 15, color: C.ink });
    s.charts.add("bar", {
      position: { left: M + 10, top: 410, width: 570, height: 202 },
      categories: stages.map((r) => r[0]),
      series: [
        { name: "断头日均", values: stages.map((r) => r[2]), fill: C.blue },
        { name: "C切日均", values: stages.map((r) => r[3]), fill: C.amber },
      ],
      hasLegend: true,
      legend: { position: "bottom", textStyle: { fill: C.muted, fontSize: 12 } },
      barOptions: { direction: "column", grouping: "clustered", gapWidth: 52 },
      yAxis: { min: 0, max: 180, majorUnit: 30, majorGridlines: { style: "solid", fill: C.line, width: 1 }, textStyle: { fill: C.muted, fontSize: 12 } },
      xAxis: { textStyle: { fill: C.muted, fontSize: 12 } },
      dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 12, bold: true } },
      chartFill: C.white,
      plotAreaFill: C.white,
    });
    box(s, M + 642, 410, 480, 202, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "现状结论", M + 670, 430, 180, 30, { size: 24, bold: true, color: C.red });
    bullets(s, [
      "0.45%阶段断头、C切均超过目标线。",
      "0.55%试调后C切下降，但断头仍未稳定达标。",
      "张力CV尖峰与断头/C切偏多相互印证，问题症结是加工稳定性不足。",
      "0.60%整体执行后，过程指标进入稳定达标区间。"
    ], M + 674, 474, 410, 98, { size: 16.5, spaceAfter: 4 });
  }

  // 5 Target
  {
    const s = p.slides.add(); base(s, 5, "目标设定：可量化、可验证、可闭环");
    stepRibbon(s, ["选题", "现状", "目标", "原因", "要因", "对策", "实施", "效果", "标准化", "总结"], 2, M, 150, W - M * 2);
    callout(s, "断头目标", "≤90个/日", M, 230, 330, 150, { color: C.green, note: "每日跟踪" });
    callout(s, "C切目标", "≤50个/日", M + 410, 230, 330, 150, { color: C.green, note: "每日跟踪" });
    callout(s, "优等率目标", "≥85%", M + 820, 230, 302, 150, { color: C.green, note: "不虚构逐日数据" });
    box(s, M, 430, W - M * 2, 114, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "目标文字", M + 26, 456, 126, 28, { size: 23, bold: true, color: C.blue });
    text(s, "55/72黑丝日均断头数控制在90个以内，日均C切数控制在50个以内，优等率提升至85%以上。", M + 160, 456, 860, 34, { size: 23, bold: true, color: C.ink });
    text(s, "说明：由于缺少优等率逐日/逐批原始数据，报告中以断头、C切过程目标完成作为优等率提升的数据支撑。", M + 160, 504, 860, 24, { size: 16, color: C.muted });
    conclusion(s, "目标符合SMART原则，且断头/C切指标具备每日检查和月度验证条件。", M, 590, W - M * 2);
  }

  // 6 Non-cause exclusions
  {
    const s = p.slides.add(); base(s, 6, "非要因排除：多项调整/试验均未解决");
    const values = [
      ["时间/项目", "尝试或变化内容", "验证结果", "结论定位"],
      ["5/9 POY来源", "前纺CP5 12线POY上机", "上机后即出现断头、C切及张力CV异常", "现状起点，非要因"],
      ["5/12 换线生产", "后续马上切换至前纺8线生产", "换线后问题未得到根本解决", "线别非主要要因"],
      ["5/16 油剂试验", "进行竹本油剂试验", "改善无明显效果", "油剂品种试验非要因"],
      ["二罗拉速度", "650m/min降至620m/min", "无明显效果", "DTY端前期尝试"],
      ["张力CV报警阈值", "报警阈值调至8；瞬间最高约20", "无明显效果", "异常现象/排查"],
      ["冷却板", "调整部分锭位冷却板", "无明显效果", "DTY端前期尝试"],
      ["DTY端油相关排查", "油轮、导丝器等早期排查", "不作为要因", "排除项"],
    ];
    const t = s.tables.add({ rows: values.length, columns: 4, left: M, top: 150, width: W - M * 2, height: 352, values, columnTracks: [{ mode: "fixed", value: 168 }, { mode: "fr", value: 1.75 }, { mode: "fr", value: 1.45 }, { mode: "fixed", value: 158 }] });
    styleTable(t, values.length, 4, C.red);
    t.cells.block({ row: 1, column: 3, rowCount: 7, columnCount: 1 }).assign({ fill: C.red2, textStyle: { color: C.red, bold: true, fontSize: 13 } });
    t.cells.block({ row: 1, column: 0, rowCount: 3, columnCount: 4 }).assign({ fill: C.amber2, textStyle: { color: C.ink, fontSize: 13 } });
    t.cells.block({ row: 1, column: 3, rowCount: 7, columnCount: 1 }).assign({ fill: C.red2, textStyle: { color: C.red, bold: true, fontSize: 13 } });
    box(s, M, 532, W - M * 2, 64, { fill: C.amber2, stroke: C.amber });
    text(s, "边界", M + 24, 553, 70, 22, { size: 20, bold: true, color: C.amber });
    text(s, "CP5 12线、8线切换、竹本油剂试验及DTY端调整均归为非要因；最终转向POY含油率控制不足。", M + 98, 551, 900, 24, { size: 19, bold: true, color: C.ink });
  }

  // 7 Cause chain
  {
    const s = p.slides.add(); base(s, 7, "张力CV波动证据：刚上机与0.55%阶段");
    const cards = [
      { img: cvWave01, title: "Thread 78｜刚上机阶段", meta: "2026/5/16-5/18", note: "蓝色CV曲线多次瞬间冲高，显示张力波动频繁。", color: C.red },
      { img: cvWave02, title: "Thread 80｜刚上机阶段", meta: "2026/5/16-5/18", note: "多个锭位出现尖峰，说明生产稳定性不足。", color: C.red },
      { img: cvWave03, title: "Thread 63｜刚上机阶段", meta: "2026/5/16-5/18", note: "CV曲线持续出现异常尖峰，与断头/C切偏多相互印证。", color: C.red },
      { img: cvWave04, title: "Thread 94｜0.55%试调阶段", meta: "2026/5/19-5/21", note: "0.55%阶段仍有明显波动，改善不稳定，需继续调整至0.60%。", color: C.amber },
    ];
    const cardW = (W - M * 2 - 30) / 2;
    const cardH = 178;
    for (const [i, c] of cards.entries()) {
      const x = M + (i % 2) * (cardW + 30);
      const y = 156 + Math.floor(i / 2) * 204;
      box(s, x, y, cardW, cardH, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
      addImage(s, c.img, x + 14, y + 16, 248, 138, {
        contentType: "image/jpeg",
        alt: c.title,
        fit: "cover",
      });
      box(s, x + 14, y + 16, 80, 24, { fill: c.color, stroke: c.color });
      text(s, i === 3 ? "0.55%" : "上机", x + 28, y + 21, 52, 12, { size: 11, bold: true, color: C.white, align: "center" });
      text(s, c.title, x + 286, y + 22, cardW - 310, 24, { size: 18, bold: true, color: c.color });
      text(s, c.meta, x + 286, y + 56, cardW - 310, 20, { size: 14, bold: true, color: C.ink });
      text(s, c.note, x + 286, y + 90, cardW - 310, 58, { size: 15, color: C.muted, lineSpacing: 1.08 });
    }
    conclusion(s, "张力CV尖峰属于前期异常证据，支撑“刚上机和0.55%试调阶段仍不稳定”的判断。", M, 584, W - M * 2, { fill: C.amber2, stroke: C.amber });
  }

  // 8 Cause chain
  {
    const s = p.slides.add(); base(s, 8, "原因分析：从现象追溯到POY含油率");
    processChain(s, ["POY油剂控制不足", "POY原丝含油率偏低", "润滑性/抱合性不足", "张力CV瞬间冲高", "断头/C切偏多"], M, 158, W - M * 2, 86);
    box(s, M, 308, 500, 230, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "5Why逻辑", M + 24, 334, 160, 28, { size: 25, bold: true, color: C.teal });
    bullets(s, [
      "为什么优等率低？断头、C切偏多。",
      "为什么断头/C切偏多？DTY加工稳定性差。",
      "为什么加工稳定性差？丝束润滑性、抱合性不足。",
      "为什么润滑性不足？配套POY原丝含油率偏低。",
      "为什么含油率偏低？POY油剂控制不足。"
    ], M + 28, 382, 440, 118, { size: 16.5, spaceAfter: 4 });
    box(s, M + 560, 308, 562, 230, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "末端因素", M + 586, 334, 160, 28, { size: 25, bold: true, color: C.amber });
    text(s, "POY油剂控制不足，导致55/72黑丝配套POY原丝含油率偏低。", M + 586, 388, 490, 62, { size: 24, bold: true, color: C.ink, lineSpacing: 1.08 });
    text(s, "本项目最终有效对策必须围绕POY油剂控制和POY含油率稳定性展开。", M + 586, 472, 490, 32, { size: 17, color: C.muted });
    conclusion(s, "最终要因：POY油剂控制不足 -> POY含油率偏低 -> DTY加工稳定性下降。", M, 590, W - M * 2, { fill: C.amber2, stroke: C.amber });
  }

  // 9 Major cause verification
  {
    const s = p.slides.add(); base(s, 9, "要因确认：阶段数据支持含油率方向");
    box(s, M, 150, 340, 58, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "验证方法", M + 18, 164, 86, 20, { size: 16, bold: true, color: C.blue });
    text(s, "按POY含油率阶段分层，对比DTY断头、C切变化。", M + 112, 164, 200, 34, { size: 14.5, color: C.ink, lineSpacing: 1.05 });
    box(s, M + 382, 150, 340, 58, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "判定标准", M + 400, 164, 86, 20, { size: 16, bold: true, color: C.green });
    text(s, "断头≤90个/日，C切≤50个/日，且趋势稳定。", M + 494, 164, 200, 34, { size: 14.5, color: C.ink, lineSpacing: 1.05 });
    box(s, M + 764, 150, 358, 58, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "数据依据", M + 782, 164, 86, 20, { size: 16, bold: true, color: C.amber });
    text(s, "5/25-6/14连续21天跟踪，断头和C切全部达标。", M + 876, 164, 210, 34, { size: 14.5, color: C.ink, lineSpacing: 1.05 });
    s.charts.add("bar", {
      position: { left: M, top: 236, width: 535, height: 250 },
      categories: ["0.45%阶段", "过渡试调", "0.60%整体执行"],
      series: [{ name: "断头日均", values: [164.3, 104.3, 33.5], fill: C.blue }],
      hasLegend: false,
      barOptions: { direction: "column", grouping: "clustered", gapWidth: 56 },
      yAxis: { min: 0, max: 180, majorUnit: 30, majorGridlines: { style: "solid", fill: C.line, width: 1 }, textStyle: { fill: C.muted, fontSize: 12 } },
      xAxis: { textStyle: { fill: C.muted, fontSize: 12 } },
      dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 13, bold: true } },
      chartFill: C.white,
      plotAreaFill: C.white,
    });
    s.charts.add("bar", {
      position: { left: M + 600, top: 236, width: 522, height: 250 },
      categories: ["0.45%阶段", "过渡试调", "0.60%整体执行"],
      series: [{ name: "C切日均", values: [77.3, 21.5, 2.0], fill: C.amber }],
      hasLegend: false,
      barOptions: { direction: "column", grouping: "clustered", gapWidth: 56 },
      yAxis: { min: 0, max: 90, majorUnit: 15, majorGridlines: { style: "solid", fill: C.line, width: 1 }, textStyle: { fill: C.muted, fontSize: 12 } },
      xAxis: { textStyle: { fill: C.muted, fontSize: 12 } },
      dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: C.ink, fontSize: 13, bold: true } },
      chartFill: C.white,
      plotAreaFill: C.white,
    });
    pill(s, "断头目标 ≤90个/日", M + 126, 508, 190, { fill: C.mint, stroke: C.green, color: C.green });
    pill(s, "0.60%后连续21天达标", M + 500, 508, 236, { fill: C.mint, stroke: C.green, color: C.green });
    pill(s, "C切目标 ≤50个/日", M + 850, 508, 180, { fill: C.mint, stroke: C.green, color: C.green });
    conclusion(s, "0.60%整体执行后，断头和C切均低于目标线，验证POY含油率控制方向有效。", M, 590, W - M * 2);
  }

  // 10 Countermeasure
  {
    const s = p.slides.add(); base(s, 10, "制定对策：5W1H计划");
    const values = [
      ["5W1H", "本项目对策"],
      ["What", "优化55/72黑丝配套POY油剂控制，提高POY原丝含油率稳定性。"],
      ["Why", "控制断头≤90个/日、C切≤50个/日，为优等率提升至85%以上提供支撑。"],
      ["Who", "前纺工艺、DTY工艺、质量、QC小组成员协同负责。"],
      ["Where", "前道POY工序及10#机DTY 55/72黑丝验证现场。"],
      ["When", "5月19日0.55%打样，5月21日0.60%打样，5月25日0.60%整体执行。"],
      ["How", "取样检测POY含油率；优化油剂添加控制；上机验证断头、C切、张力CV；固化控制标准。"],
    ];
    const t = s.tables.add({ rows: values.length, columns: 2, left: M, top: 150, width: W - M * 2, height: 390, values, columnTracks: [{ mode: "fixed", value: 125 }, { mode: "fr", value: 1 }] });
    styleTable(t, values.length, 2, C.blue);
    t.cells.block({ row: 1, column: 0, rowCount: 6, columnCount: 1 }).assign({ fill: C.soft, textStyle: { bold: true, color: C.blue, fontSize: 15 } });
    conclusion(s, "对策只针对主要要因POY油剂控制制定，DTY端调整不纳入最终有效对策。", M, 590, W - M * 2, { fill: C.amber2, stroke: C.amber });
  }

  // 11 Implementation
  {
    const s = p.slides.add(); base(s, 11, "实施对策：含油率调整与验证节奏");
    timeline(s, [
      { date: "05/09", text: "CP5 12线POY上机，异常显现", warn: true },
      { date: "05/12", text: "切换前纺8线生产，未根本改善", warn: true },
      { date: "05/16", text: "竹本油剂试验，无明显效果", warn: true },
      { date: "05/19", text: "POY含油率0.45%调至0.55%打样", warn: true },
      { date: "05/21", text: "POY含油率0.60%打样" },
      { date: "05/22", text: "0.55%整体执行，改善不明显", warn: true },
      { date: "05/25", text: "0.60%整体执行，趋于稳定", good: true },
      { date: "06/14", text: "连续21天目标跟踪完成", good: true },
    ], M + 28, 168, W - M * 2 - 56);
    box(s, M, 420, 520, 130, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "实施重点", M + 24, 444, 130, 28, { size: 24, bold: true, color: C.blue });
    bullets(s, [
      "从0.45%逐步试调至0.60%，以断头、C切和张力CV反馈效果。",
      "0.55%阶段不作为最终有效措施，0.60%整体执行作为验证口径。"
    ], M + 28, 492, 456, 44, { size: 16.5, spaceAfter: 2 });
    box(s, M + 600, 420, 522, 130, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "执行边界", M + 624, 444, 130, 28, { size: 24, bold: true, color: C.amber });
    bullets(s, [
      "0.60%控制值写作“按工艺文件/通知单最终确认值为准”。",
      "换批、换规格或油剂系统调整后，先小批量打样验证。"
    ], M + 628, 492, 452, 44, { size: 16.5, spaceAfter: 2 });
  }

  // 12 Effect
  {
    const s = p.slides.add(); base(s, 12, "现场照片：三个阶段对比");
    const cardY = 166;
    const cardW = 352;
    const imgH = 198;
    const gap = 42;
    const x1 = M;
    const x2 = M + cardW + gap;
    const x3 = M + (cardW + gap) * 2;
    const cards = [
      {
        x: x1,
        title: "刚上机阶段（图2）",
        subtitle: "0.45%阶段｜生产不稳定",
        img: photoInitial,
        color: C.red,
        note: "断头、C切偏多，作为改善前现场状态。",
      },
      {
        x: x2,
        title: "0.55%含油阶段（图1）",
        subtitle: "试调阶段｜改善不明显",
        img: photo055,
        color: C.amber,
        note: "C切下降但断头仍未稳定达标，继续调整。",
      },
      {
        x: x3,
        title: "0.60%含油稳定（图3）",
        subtitle: "整体执行后｜生产趋稳",
        img: photo060,
        color: C.green,
        note: "与连续21天断头、C切达标数据相互印证。",
      },
    ];
    for (const c of cards) {
      box(s, c.x, cardY, cardW, 370, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
      addImage(s, c.img, c.x + 14, cardY + 16, cardW - 28, imgH, {
        contentType: "image/jpeg",
        alt: c.title,
        fit: "cover",
      });
      box(s, c.x + 14, cardY + 16, 116, 30, { fill: c.color, stroke: c.color });
      text(s, c.title.includes("刚上机") ? "图2" : c.title.includes("0.55") ? "图1" : "图3", c.x + 42, cardY + 23, 60, 16, { size: 13, bold: true, color: C.white, align: "center" });
      text(s, c.title, c.x + 18, cardY + 232, cardW - 36, 25, { size: 20, bold: true, color: c.color });
      text(s, c.subtitle, c.x + 18, cardY + 266, cardW - 36, 22, { size: 16, bold: true, color: C.ink });
      text(s, c.note, c.x + 18, cardY + 304, cardW - 36, 42, { size: 15, color: C.muted, lineSpacing: 1.08 });
    }
    conclusion(s, "现场照片作为过程佐证：从刚上机异常、0.55%试调，到0.60%整体执行后生产趋稳。", M, 584, W - M * 2);
  }

  // 13 Effect
  {
    const s = p.slides.add(); base(s, 13, "效果检查：近一个月每日目标跟踪");
    box(s, M, 146, W - M * 2, 50, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "验证周期", M + 20, 161, 86, 18, { size: 16, bold: true, color: C.blue });
    text(s, "2026/5/15-6/14，共31天；0.60%整体执行后重点跟踪5/25-6/14。", M + 110, 161, 340, 18, { size: 14.5, color: C.ink });
    text(s, "判定口径", M + 484, 161, 86, 18, { size: 16, bold: true, color: C.amber });
    text(s, "断头≤90且C切≤50，综合判定为达标。", M + 574, 161, 250, 18, { size: 14.5, color: C.ink });
    text(s, "验证结果", M + 856, 161, 86, 18, { size: 16, bold: true, color: C.green });
    text(s, "0.60%执行后21天全部达标。", M + 946, 161, 160, 18, { size: 14.5, color: C.ink });
    s.charts.add("line", {
      position: { left: M, top: 214, width: W - M * 2, height: 286 },
      categories: daily.map((d) => d[0]),
      series: [
        { name: "实际断头", values: daily.map((d) => d[1]), line: { style: "solid", fill: C.blue, width: 2.4 }, marker: { symbol: "circle", size: 4 } },
        { name: "断头目标90", values: daily.map(() => 90), line: { style: "dashed", fill: C.green, width: 2 }, marker: { symbol: "none" } },
        { name: "实际C切", values: daily.map((d) => d[2]), line: { style: "solid", fill: C.amber, width: 2.4 }, marker: { symbol: "circle", size: 4 } },
        { name: "C切目标50", values: daily.map(() => 50), line: { style: "dashed", fill: C.red, width: 2 }, marker: { symbol: "none" } },
      ],
      hasLegend: true,
      legend: { position: "bottom", textStyle: { fill: C.muted, fontSize: 12 } },
      yAxis: { min: 0, max: 260, majorUnit: 50, majorGridlines: { style: "solid", fill: C.line, width: 1 }, textStyle: { fill: C.muted, fontSize: 11 } },
      xAxis: { textStyle: { fill: C.muted, fontSize: 8 } },
      lineOptions: { smooth: false },
      chartFill: C.white,
      plotAreaFill: C.white,
    });
    callout(s, "近一个月综合达标", "21/31天", M, 530, 260, 106, { color: C.green, valueSize: 32 });
    callout(s, "0.60%执行后", "21/21天", M + 298, 530, 260, 106, { color: C.green, valueSize: 32 });
    callout(s, "最大日断头", "52个", M + 596, 530, 250, 106, { color: C.blue, valueSize: 32 });
    callout(s, "最大日C切", "9个", M + 884, 530, 238, 106, { color: C.amber, valueSize: 32 });
  }

  // 14 Results
  {
    const s = p.slides.add(); base(s, 14, "目标完成情况：过程指标显著改善");
    callout(s, "改善后断头日均", "33.5个", M, 154, 260, 120, { color: C.green, note: "目标≤90个/日" });
    callout(s, "改善后C切日均", "2.0个", M + 296, 154, 260, 120, { color: C.green, note: "目标≤50个/日" });
    callout(s, "断头下降幅度", "79.6%", M + 592, 154, 260, 120, { color: C.blue, note: "较0.45%上机初期" });
    callout(s, "C切下降幅度", "97.5%", M + 888, 154, 234, 120, { color: C.blue, note: "较0.45%上机初期" });
    const values = [
      ["对比口径", "断头日均", "C切日均"],
      ["上机初期0.45%阶段", "164.3个", "77.3个"],
      ["0.60%整体执行后", "33.5个", "2.0个"],
      ["下降幅度", "79.6%", "97.5%"],
    ];
    const t = s.tables.add({ rows: values.length, columns: 3, left: M, top: 330, width: 520, height: 176, values, columnTracks: [{ mode: "fr", value: 1.4 }, { mode: "fixed", value: 142 }, { mode: "fixed", value: 142 }] });
    styleTable(t, values.length, 3, C.green);
    t.cells.block({ row: 3, column: 1, rowCount: 1, columnCount: 2 }).assign({ fill: C.mint, textStyle: { color: C.green, bold: true, fontSize: 15 } });
    box(s, M + 602, 330, 520, 176, { fill: C.white, stroke: C.line, shadow: "shadow-sm" });
    text(s, "目标完成说明", M + 628, 356, 190, 28, { size: 24, bold: true, color: C.blue });
    bullets(s, [
      "5月25日至6月14日连续21天断头和C切全部达标。",
      "断头、C切两项关键过程指标已完成对策目标。",
      "不虚构改善后优等率；用过程目标完成支撑优等率提升。"
    ], M + 632, 402, 440, 68, { size: 16.5, spaceAfter: 4 });
    conclusion(s, "过程目标完成，为55/72黑丝优等率提升至85%以上提供数据支撑。", M, 584, W - M * 2);
  }

  // 15 Standardization
  {
    const s = p.slides.add(); base(s, 15, "标准化与防护措施");
    const values = [
      ["序号", "标准化/防护项目", "控制要求", "记录/验证方式"],
      ["1", "POY油剂控制标准固化", "55/72黑丝配套POY含油率控制值按0.60%执行，允许范围以工艺文件/通知单为准。", "POY含油率检测记录、工艺通知单"],
      ["2", "POY含油率首批确认", "上机前确认批号、含油率检测结果和工艺调整状态，未经确认不得批量转入DTY验证。", "首批确认记录、交接记录"],
      ["3", "DTY关键过程指标跟踪", "每日跟踪断头、C切和张力CV异常；断头>90或C切>50时启动异常反馈。", "每日目标跟踪表"],
      ["4", "异常联动处置", "优先追溯配套POY含油率、油剂批次和前纺执行情况；DTY端调整仅作为排查项。", "异常分析单、纠正措施记录"],
      ["5", "换批/换规格防护", "油剂系统或工艺参数调整后，先小批量打样验证，稳定后再整体推行。", "小样验证记录"],
    ];
    const t = s.tables.add({ rows: values.length, columns: 4, left: M, top: 150, width: W - M * 2, height: 356, values, columnTracks: [{ mode: "fixed", value: 58 }, { mode: "fixed", value: 190 }, { mode: "fr", value: 2.2 }, { mode: "fr", value: 1 }] });
    styleTable(t, values.length, 4);
    t.cells.block({ row: 1, column: 0, rowCount: 5, columnCount: 1 }).assign({ fill: C.soft, textStyle: { bold: true, color: C.navy, fontSize: 14 } });
    conclusion(s, "标准化重点：POY油剂控制标准化 + 上机前确认 + DTY过程日跟踪 + 异常联动处置。", M, 584, W - M * 2);
  }

  // 16 Closing
  {
    const s = p.slides.add();
    addImage(s, bg, 0, 0, W, H, { fit: "cover" });
    box(s, 0, 0, W, H, { fill: "#061A2BDD", stroke: "#061A2B00" });
    text(s, "总结与下一步", M, 66, 520, 52, { size: 42, bold: true, color: C.white });
    box(s, M, 160, W - M * 2, 1, { fill: "#FFFFFF55", stroke: "#FFFFFF55" });
    box(s, M, 218, 330, 260, { fill: "#FFFFFF12", stroke: "#FFFFFF35" });
    text(s, "本次成果", M + 24, 246, 150, 28, { size: 24, bold: true, color: "#5DD0CE" });
    bullets(s, ["明确最终要因：POY油剂控制不足导致POY含油率偏低。", "0.60%整体执行后，连续21天断头和C切全部达标。", "形成55/72黑丝配套POY油剂控制经验。"], M + 28, 296, 270, 116, { size: 17, color: C.white });
    box(s, M + 392, 218, 330, 260, { fill: "#FFFFFF12", stroke: "#FFFFFF35" });
    text(s, "经验沉淀", M + 416, 246, 150, 28, { size: 24, bold: true, color: "#5DD0CE" });
    bullets(s, ["前期DTY端尝试不作为要因和最终有效措施。", "问题分析需前后道联动，避免局部反复试错。", "过程目标数据每日记录，支撑评审闭环。"], M + 420, 296, 270, 116, { size: 17, color: C.white });
    box(s, M + 784, 218, 338, 260, { fill: "#FFFFFF12", stroke: "#FFFFFF35" });
    text(s, "下一步", M + 808, 246, 130, 28, { size: 24, bold: true, color: "#F3BE63" });
    bullets(s, ["继续跟踪后续批次断头、C切和优等率。", "将经验扩展到同类细旦黑丝/高敏感规格。", "按公司财务口径确认最终经济效益。"], M + 812, 296, 280, 116, { size: 17, color: C.white });
    box(s, M, 560, W - M * 2, 66, { fill: "#5DD0CE22", stroke: "#5DD0CE88" });
    text(s, "统一结论", M + 26, 581, 112, 22, { size: 20, bold: true, color: "#5DD0CE" });
    text(s, "优化POY油剂控制后，55/72黑丝DTY加工稳定性明显提升；断头、C切目标完成，具备标准化和推广基础。", M + 150, 578, 900, 28, { size: 21, bold: true, color: C.white });
  }

  for (const [idx, slide] of p.slides.items.entries()) {
    const stem = `slide-${String(idx + 1).padStart(2, "0")}`;
    await writeBlob(`${PREVIEW}/${stem}.png`, await p.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${LAYOUT}/${stem}.json`, await layout.text());
  }
  await writeBlob(`${PREVIEW}/deck-montage.webp`, await p.export({ format: "webp", montage: true, scale: 1 }));
  const deck = await PresentationFile.exportPptx(p);
  await deck.save(FINAL);
  const inspect = await p.inspect({ kind: "slide,textbox,shape,table,chart,image", maxChars: 20000 });
  await fs.writeFile(`${QA}/inspect-v2.ndjson`, inspect.ndjson);
  console.log(FINAL);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
