import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const execFileAsync = promisify(execFile);

const ROOT = "/Users/USER/Documents/Codex/2026-06-21/new-chat-3";
const OUT = `${ROOT}/outputs`;
const TMP = `${ROOT}/work/qcc_2026_generated/tmp/real_style`;
const PREVIEW = `${TMP}/preview`;
const QA = `${TMP}/qa`;
const ASSET = `${ROOT}/work/presentations/qcc_55_72_heisi/tmp/assets`;
const TREND = `${ROOT}/outputs/qcc_2026_generated/近一个月断头_C切趋势图.png`;
const LOGO_DOCX = "/tmp/codex-remote-attachments/019ee940-a5e2-7413-ad9f-29b56e38089a/7B4A42F2-CB03-432C-BC7A-F1B775261D66/1-logo.docx";
const LOGO_ROOT = `${TMP}/logo_assets`;
const LOGO_MEDIA = `${LOGO_ROOT}/raw/word/media`;
const LOGO1 = `${LOGO_MEDIA}/image1.png`;
const LOGO2 = `${LOGO_MEDIA}/image2.png`;
const LOGO3 = `${LOGO_MEDIA}/image3.png`;

const W = 1280;
const H = 720;
const FONT = "Microsoft YaHei";
const C = {
  white: "#FFFFFF",
  black: "#000000",
  red: "#D60000",
  green: "#009B52",
  yellow: "#FFF36A",
  paleYellow: "#FFF4A3",
  paleGreen: "#DFF6D5",
  paleBlue: "#DDEFFF",
  lightGray: "#F3F3F3",
  line: "#111111",
  muted: "#555555",
  blue: "#00A6D6",
  orange: "#FF8A2A",
};

const daily = [
  ["5月09日", 151, 103, "0.45%", "CP5 12线POY上机"],
  ["5月10日", 143, 82, "0.45%", "上机初期"],
  ["5月11日", 180, 64, "0.45%", "调整油轮导丝器"],
  ["5月12日", 142, 82, "0.45%", "换至8线生产"],
  ["5月13日", 190, 64, "0.45%", "仍高位波动"],
  ["5月14日", 141, 89, "0.45%", "仍高位波动"],
  ["5月15日", 243, 114, "0.45%", "张力CV波动明显"],
  ["5月16日", 180, 71, "0.45%", "竹本油剂试验"],
  ["5月17日", 130, 50, "0.45%", "仍高位波动"],
  ["5月18日", 143, 54, "0.45%", "仍高位波动"],
  ["5月19日", 112, 34, "0.55%", "前纺含油0.55%打样"],
  ["5月20日", 98, 21, "0.55%", "有效但不明显"],
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

const factors = [
  ["人", "操作培训不到位", "现场检查/记录核对", "非要因", "班组按工艺要求执行，异常主要随含油阶段变化。"],
  ["人", "上机异常反馈不及时", "查异常发现、反馈、验证闭环", "要因", "0.55%有效但不明显，整体切换0.60%滞后至5月25日。"],
  ["机", "DTY10#机油轮导丝器异常", "5月11-14日调整后跟踪", "非要因", "调整后5月15-18日断头仍高位，不能解释最终改善。"],
  ["机", "张力传感/报警阈值异常", "查看CV波动与报警记录", "非要因", "CV波动是异常表现，未见阈值调整带来持续改善。"],
  ["料", "POY原丝含油率偏低", "分阶段对比0.45/0.55/0.60", "要因", "0.60%后断头日均33.5、C切2.0，连续达标。"],
  ["料", "前纺线别切换影响", "对比CP5 12线及8线生产", "非要因", "5月12日换8线后仍高位波动。"],
  ["料", "竹本油剂试验影响", "5月16日换油剂试验", "非要因", "5月16-18日未形成明显改善。"],
  ["法", "55/72黑丝POY含油控制标准未固化", "查上机初期与试调过程", "要因", "0.55%仍不稳定，0.60%整体执行后连续达标。"],
  ["法", "DTY端局部参数调整不足", "核对临时调整记录", "非要因", "局部调整未改变高断头趋势。"],
  ["法", "检验/分级标准不统一", "核对系统记录、日报、趋势图", "非要因", "数据口径连续，阶段差异清晰。"],
  ["环", "环境温湿度波动", "对比环境与异常走势", "非要因", "改善节点与0.60%整体执行一致。"],
  ["环", "上机初期现场状态波动", "观察现场状态", "非要因", "不能解释0.55%与0.60%阶段差异。"],
];

async function ensureDirs() {
  await fs.mkdir(OUT, { recursive: true });
  await fs.mkdir(PREVIEW, { recursive: true });
  await fs.mkdir(QA, { recursive: true });
  await fs.rm(LOGO_ROOT, { recursive: true, force: true });
  await fs.mkdir(LOGO_ROOT, { recursive: true });
  await execFileAsync("unzip", ["-qo", LOGO_DOCX, "word/media/*", "-d", `${LOGO_ROOT}/raw`]);
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
    fontSize: o.size ?? 24,
    bold: o.bold ?? false,
    color: o.color || C.black,
    alignment: o.align || "left",
    verticalAlignment: o.valign || "top",
    lineSpacing: o.lineSpacing ?? 1.15,
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
    line: { style: "solid", fill: o.stroke ?? C.line, width: o.lineWidth ?? 1 },
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

function table(slide, headers, rows, x, y, w, rowH, widths, o = {}) {
  const colW = widths || headers.map(() => w / headers.length);
  let cx = x;
  headers.forEach((h, i) => {
    shape(slide, cx, y, colW[i], rowH, { fill: o.headerFill || C.green, stroke: C.line, lineWidth: 1.2 });
    addText(slide, h, cx + 6, y + 9, colW[i] - 12, rowH - 10, { size: o.headerSize || 22, bold: true, color: o.headerColor || C.white, align: "center", valign: "middle" });
    cx += colW[i];
  });
  rows.forEach((row, r) => {
    cx = x;
    row.forEach((v, i) => {
      const fill = o.bodyFill || (r % 2 ? C.white : "#FAFAFA");
      shape(slide, cx, y + rowH * (r + 1), colW[i], rowH, { fill, stroke: C.line, lineWidth: 1 });
      addText(slide, String(v), cx + 8, y + rowH * (r + 1) + 8, colW[i] - 16, rowH - 12, {
        size: o.bodySize || 18,
        bold: (o.boldColumns || []).includes(i),
        color: o.bodyColor || C.black,
        align: (o.leftColumns || []).includes(i) ? "left" : "center",
        valign: "middle",
        lineSpacing: 1.08,
      });
      cx += colW[i];
    });
  });
}

function frame(slide, title, pageNo) {
  slide.background.fill = C.white;
  addText(slide, title, 300, 36, 680, 54, { size: 41, bold: false, align: "center" });
  shape(slide, 4, 0, 6, H, { fill: "#FFB68F", stroke: "#FFB68F" });
  shape(slide, W - 12, 0, 8, H, { fill: "#FFB68F", stroke: "#FFB68F" });
  shape(slide, W - 66, H - 118, 58, 58, { geometry: "ellipse", fill: C.orange, stroke: C.orange, shadow: "shadow-sm" });
  if (pageNo) addText(slide, String(pageNo), W - 52, H - 104, 32, 20, { size: 14, bold: true, color: C.white, align: "center" });
}

async function applyLogos(slide, pageNo, total) {
  if (pageNo === 1 || pageNo === total) {
    await addImage(slide, LOGO2, 38, 28, 210, 84, "contain");
  } else {
    await addImage(slide, LOGO1, W - 75, 5, 52, 62, "contain");
  }
  await addImage(slide, LOGO3, (W - 300) / 2, H - 54, 300, 32, "contain");
}

function note(slide, text, x, y, w, h, fill = "#FFF7B2") {
  shape(slide, x, y, w, h, { fill, stroke: fill });
  addText(slide, text, x + 18, y + 14, w - 36, h - 20, { size: 22, lineSpacing: 1.18 });
}

function sectionBand(slide, text, x, y, w, fill = C.blue) {
  shape(slide, x, y, w, 44, { geometry: "ellipse", fill, stroke: fill });
  addText(slide, text, x + 18, y + 8, w - 36, 26, { size: 23, bold: true, color: C.white, align: "center" });
}

function fishbone(slide) {
  addText(slide, "55/72黑丝\n断头、C切\n高位波动", 120, 308, 180, 120, { size: 32, lineSpacing: 1.15, align: "center", valign: "middle" });
  const mainY = 378;
  shape(slide, 305, mainY, 770, 3, { fill: C.line, stroke: C.line });
  const groups = [
    ["人", 330, 140, "top", ["操作培训不到位", "上机异常反馈不及时"], [false, true]],
    ["料", 575, 128, "top", ["POY原丝含油率偏低", "线别切换影响", "竹本油剂试验影响"], [true, false, false]],
    ["环", 840, 140, "top", ["环境温湿度波动", "上机现场状态波动"], [false, false]],
    ["机", 330, 448, "bottom", ["油轮导丝器异常", "张力CV报警阈值异常"], [false, false]],
    ["法", 575, 430, "bottom", ["含油控制标准未固化", "DTY端局部参数调整", "分级口径不统一"], [true, false, false]],
  ];
  groups.forEach(([dim, x, y, side, items, keyFlags]) => {
    const boxX = x + 62;
    const dimCenterX = x - 1;
    const dimCenterY = y + 29;
    const connectorTop = side === "top" ? dimCenterY + 25 : mainY;
    const connectorBottom = side === "top" ? mainY : dimCenterY - 25;
    shape(slide, dimCenterX, connectorTop, 2, connectorBottom - connectorTop, { fill: C.line, stroke: C.line });
    shape(slide, dimCenterX, dimCenterY, 62, 2, { fill: C.line, stroke: C.line });
    shape(slide, x - 24, y + 6, 46, 46, { fill: C.white, stroke: C.line });
    addText(slide, dim, x - 24, y + 16, 46, 24, { size: 25, align: "center" });
    items.forEach((it, i) => {
      const yy = y + i * 50;
      shape(slide, boxX, yy, 178, 34, { fill: keyFlags[i] ? C.yellow : C.white, stroke: C.line });
      addText(slide, it, boxX + 4, yy + 7, 170, 22, { size: 17, bold: keyFlags[i], align: "center" });
      shape(slide, x + 22, yy + 17, 40, 2, { fill: C.line, stroke: C.line });
    });
  });
}

function factorConfirmSlide(slide, factorNo, factor, method, fact, conclusion, isCause, extra = "") {
  sectionBand(slide, `因素${factorNo}：${factor}`, 64, 120, 430, isCause ? C.red : C.blue);
  addText(slide, `验证方法：${method}`, 82, 188, 880, 36, { size: 25 });
  note(slide, fact, 80, 250, 780, 160, "#FFFAC4");
  addText(slide, `结论：${factor}，属${isCause ? "要因" : "非要因"}。`, 82, 494, 900, 40, { size: 30, bold: true, color: isCause ? C.red : C.black });
  if (extra) addText(slide, extra, 82, 548, 980, 48, { size: 21, color: C.muted, lineSpacing: 1.15 });
}

async function buildDeck() {
  const p = Presentation.create({ slideSize: { width: W, height: H } });
  const slides = [];
  const pending = [];
  const add = (title, cb) => {
    const s = p.slides.add();
    slides.push(s);
    frame(s, title, slides.length);
    const result = cb?.(s);
    if (result && typeof result.then === "function") pending.push(result);
    return s;
  };

  {
    const s = p.slides.add();
    slides.push(s);
    s.background.fill = C.white;
    await addImage(s, `${ASSET}/qcc-fiber-bg-imagegen.png`, 650, 0, 630, H, "cover");
    shape(s, 0, 0, 720, H, { fill: "#FFFFFFE8", stroke: "#FFFFFF00" });
    addText(s, "提高55/72黑丝优等率", 78, 210, 620, 70, { size: 49, bold: true });
    addText(s, "长空之王QC小组", 84, 316, 360, 38, { size: 30 });
    addText(s, "问题解决型QCC成果报告", 84, 382, 430, 34, { size: 25, color: C.muted });
    shape(s, 4, 0, 6, H, { fill: "#FFB68F", stroke: "#FFB68F" });
  }

  add("一、小组简介", (s) => {
    table(s, ["项目", "内容", "项目", "内容"], [
      ["小组名称", "长空之王QC小组", "课题类型", "问题解决型"],
      ["课题名称", "提高55/72黑丝优等率", "活动周期", "2026.05.09-2026.06.14"],
      ["改善对象", "10#机55/72黑丝", "主攻指标", "断头数、C切数、优等率"],
      ["小组定位", "前纺、加弹、质检联动", "活动口径", "按PDCA步骤推进"],
    ], 80, 150, 1120, 70, [170, 390, 170, 390], { headerFill: C.green, bodySize: 20, leftColumns: [1, 3] });
    note(s, "本课题围绕55/72黑丝上机后断头、C切高位波动问题展开，重点确认POY含油控制与前后道联动验证是否影响加工稳定性。", 110, 520, 1000, 78);
  });

  add("二、计划实施", (s) => {
    table(s, ["阶段", "主要步骤", "时间", "输出"], [
      ["P", "选择课题、现状调查、目标设定、原因分析、要因确认", "5月09日-5月24日", "确认3个要因"],
      ["D", "制定对策、实施0.60%含油整体执行和联动闭环", "5月25日-6月05日", "形成对策与标准"],
      ["C", "每日跟踪断头/C切，验证连续达标情况", "5月25日-6月14日", "连续21天达标"],
      ["A", "标准化、防复发、资料归档", "6月上旬", "防措与成果报告"],
    ], 90, 145, 1100, 86, [130, 500, 220, 250], { headerFill: C.green, bodySize: 20, leftColumns: [1, 3], boldColumns: [0] });
  });

  add("目 录", (s) => {
    const items = [
      ["1", "小组简介", "Group profile"],
      ["2", "选题理由", "Topic selection"],
      ["3", "现状调查", "Current status"],
      ["4", "目标设定", "Set goals"],
      ["5", "原因分析与要因确认", "Cause analysis"],
      ["6", "对策实施与效果检查", "Implementation"],
      ["7", "标准化及总结", "Standardization"],
    ];
    items.forEach((it, i) => {
      const x = 120 + (i % 2) * 520;
      const y = 140 + Math.floor(i / 2) * 110;
      shape(s, x, y, 70, 70, { geometry: "ellipse", fill: i < 5 ? C.green : C.orange, stroke: C.white, lineWidth: 2 });
      addText(s, it[0], x, y + 14, 70, 38, { size: 32, bold: true, color: C.white, align: "center" });
      addText(s, it[1], x + 95, y + 5, 300, 32, { size: 30, bold: true });
      addText(s, it[2], x + 95, y + 43, 300, 24, { size: 17, color: C.muted });
    });
  });

  add("三、选题理由", (s) => {
    const reasons = [
      ["01", "55/72黑丝上机后断头、C切高位波动"],
      ["02", "影响10#机生产稳定性和产品优等率"],
      ["03", "张力CV波动、断头和C切异常相互印证"],
      ["04", "需要形成前纺-加弹-质检联动闭环"],
    ];
    reasons.forEach((r, i) => {
      const x = 100 + i * 270;
      shape(s, x, 160, 210, 210, { fill: C.paleGreen, stroke: C.green, lineWidth: 2 });
      addText(s, r[0], x + 58, 190, 94, 34, { size: 25, bold: true, color: C.green, align: "center" });
      addText(s, r[1], x + 24, 250, 162, 78, { size: 22, bold: true, align: "center", lineSpacing: 1.15 });
    });
    note(s, "课题确定：提高55/72黑丝优等率。通过降低断头与C切高位波动，支撑产品稳定生产。", 160, 470, 920, 86);
  });

  add("四、现状调查", async (s) => {
    addText(s, "刚上机和0.55%含油时期，断头、C切及张力CV波动明显。", 88, 130, 1040, 36, { size: 26 });
    await addImage(s, `${ASSET}/cv-wave-crop-sheet.jpg`, 70, 190, 545, 270, "cover");
    await addImage(s, `${ASSET}/photo-initial-machine.jpg`, 665, 190, 545, 270, "cover");
    note(s, "现场证据表明：异常并非单一班次或单一操作造成，而是随上机阶段和含油率变化呈现阶段性波动。", 130, 500, 980, 72);
  });

  add("四、现状调查", (s) => {
    table(s, ["阶段", "日期范围", "断头日均", "C切日均", "现状判断"], [
      ["上机初期/0.45%", "5月09日-5月18日", "164.3", "77.3", "断头、C切均高位波动"],
      ["0.55%试调", "5月19日-5月24日", "104.3", "21.5", "C切下降，断头仍高于目标"],
      ["0.60%整体执行", "5月25日-6月14日", "33.5", "2.0", "断头、C切稳定达标"],
    ], 95, 150, 1090, 86, [230, 230, 150, 150, 330], { headerFill: C.green, bodySize: 21, boldColumns: [0] });
    note(s, "调查结论：0.55%只能说明方向有效但不足，必须继续确认含油控制标准和前后道联动机制。", 150, 500, 930, 80);
  });

  add("五、目标设定", (s) => {
    shape(s, 230, 190, 210, 85, { fill: C.paleYellow, stroke: C.green, lineWidth: 3 });
    addText(s, "改善前", 230, 205, 210, 30, { size: 28, bold: true, align: "center" });
    addText(s, "断头141.8\nC切56.4", 230, 238, 210, 44, { size: 22, align: "center", lineSpacing: 1.05 });
    shape(s, 535, 285, 210, 85, { fill: C.paleYellow, stroke: C.green, lineWidth: 3 });
    addText(s, "目标", 535, 300, 210, 30, { size: 28, bold: true, align: "center" });
    addText(s, "断头≤90\nC切≤50", 535, 333, 210, 44, { size: 22, align: "center", lineSpacing: 1.05 });
    shape(s, 845, 380, 210, 85, { fill: C.paleGreen, stroke: C.green, lineWidth: 3 });
    addText(s, "改善后", 845, 395, 210, 30, { size: 28, bold: true, align: "center" });
    addText(s, "断头33.5\nC切2.0", 845, 428, 210, 44, { size: 22, align: "center", lineSpacing: 1.05 });
    shape(s, 440, 235, 100, 18, { fill: "#00D81D", stroke: "#00D81D" });
    shape(s, 745, 330, 100, 18, { fill: "#00D81D", stroke: "#00D81D" });
    note(s, "目标：断头≤90个/日，C切≤50个/日；通过过程稳定推动优等率达到≥85%。", 160, 530, 960, 70);
  });

  add("六、原因分析", (s) => fishbone(s));

  add("七、要因分析", (s) => {
    table(s, ["序号", "末端因素", "确认方法", "判断标准", "结论"], factors.map((f, i) => [
      i + 1, f[1], f[2], f[3] === "要因" ? "指标应随改善稳定达标" : "验证后不能解释主波动", f[3],
    ]), 55, 132, 1170, 38, [60, 285, 285, 350, 190], { headerFill: C.paleYellow, headerColor: C.black, bodyFill: C.white, bodySize: 14.5, headerSize: 21, leftColumns: [1, 2, 3], boldColumns: [0, 4] });
  });

  add("八、要因确认", (s) => factorConfirmSlide(s, "一", "上机异常反馈不及时", "查异常发现、试调、反馈及整体切换记录", "前期0.55%试调有效果但不明显，整体切换至0.60%滞后至5月25日；异常验证未形成当天闭环。", "要因", true, "后续需建立前纺-加弹-质检联动验证闭环。"));

  add("八、要因确认", (s) => factorConfirmSlide(s, "二", "POY原丝含油率偏低", "按0.45%、0.55%、0.60%三个阶段对比断头/C切", "0.45%阶段断头日均164.3、C切77.3；0.55%阶段断头仍104.3；0.60%阶段断头33.5、C切2.0。", "要因", true, "数据变化与含油率阶段高度一致。"));

  add("八、要因确认", async (s) => {
    factorConfirmSlide(s, "三", "55/72黑丝POY含油控制标准未固化", "查0.55%试调和0.60%整体执行后的连续达标情况", "0.55%阶段C切下降但断头未稳定；5月25日0.60%整体执行后，断头和C切连续21天达标。", "要因", true, "0.60%应作为稳定生产标准，0.55%仅作过渡验证。");
    await addImage(s, `${ASSET}/photo-0.60-stable.jpg`, 888, 245, 300, 170, "cover");
  });

  add("八、要因确认", (s) => {
    factorConfirmSlide(s, "四", "操作培训不到位、DTY端油轮导丝器异常", "现场检查、调整记录、异常趋势跟踪", "班组按工艺执行；5月11-14日调整油轮导丝器后，5月15-18日断头仍高位。", "非要因", false, "保留培训抽查和设备点检，但不写成最终有效对策。");
  });

  add("八、要因确认", (s) => {
    factorConfirmSlide(s, "五", "前纺线别切换、竹本油剂试验", "对比5月9日CP5 12线、5月12日8线、5月16日竹本油剂试验", "5月12日换8线后仍高位波动；5月16-18日断头180/130/143，C切71/50/54，未形成明显改善。", "非要因", false, "线别与油剂试验作为排除证据保留。");
  });

  add("八、要因确认", (s) => {
    factorConfirmSlide(s, "六", "DTY端局部参数、分级口径、环境因素", "核对临时调整、日报口径、环境及现场状态", "DTY端局部试调未改变趋势；数据记录连续；环境改善节点与0.60%整体执行一致。", "非要因", false, "继续监控，不作为本次课题主要原因。");
  });

  add("九、要因确认汇总", (s) => {
    addText(s, "通过对十二项末端因素进行逐项验证，最终确认要因共三项：", 92, 132, 1000, 34, { size: 27 });
    table(s, ["序号", "要因内容", "确认依据", "后续处理"], [
      ["1", "上机异常反馈不及时", "0.55%有效但不充分，整体切换滞后", "建立联动验证闭环"],
      ["2", "POY原丝含油率偏低", "含油率升至0.60%后指标稳定", "按0.60%目标控制"],
      ["3", "55/72黑丝POY含油控制标准未固化", "0.55%不能稳定，0.60%连续达标", "固化专用含油标准"],
    ], 100, 210, 1080, 94, [100, 310, 400, 270], { headerFill: C.green, bodySize: 21, leftColumns: [1, 2, 3], boldColumns: [0, 1] });
  });

  add("十、对策制定", (s) => {
    table(s, ["要因", "对策", "目标", "措施", "完成时间"], [
      ["POY原丝含油率偏低", "含油率目标调整至0.60%", "断头≤90，C切≤50", "打样10锭后整体执行，每日跟踪", "5月25日"],
      ["含油控制标准未固化", "制定55/72黑丝专用含油标准", "不再以0.55%为稳定标准", "首批上机确认含油目标和切换时间", "6月01日"],
      ["异常联动验证滞后", "建立前纺-加弹-质检闭环", "当天形成试调结论", "打样、线别、油剂、含油率结果一表跟踪", "6月05日"],
    ], 70, 150, 1140, 108, [230, 260, 210, 310, 130], { headerFill: C.green, bodySize: 18.5, leftColumns: [0, 1, 2, 3], boldColumns: [0] });
  });

  add("十一、对策实施", async (s) => {
    addText(s, "1、POY含油率调整及整体执行", 90, 130, 620, 32, { size: 29, bold: true });
    table(s, ["时间", "实施内容", "跟踪结果"], [
      ["5月19日", "前纺含油由0.45%调整至0.55%，先打样", "C切下降，断头仍高于目标"],
      ["5月21日", "同步进行0.60%打样10锭", "为整体执行提供依据"],
      ["5月25日", "0.60%整体做到", "断头/C切进入稳定目标区间"],
    ], 80, 190, 650, 86, [130, 340, 180], { headerFill: C.green, bodySize: 19, leftColumns: [1, 2] });
    await addImage(s, `${ASSET}/photo-0.55-oil.jpg`, 770, 180, 380, 185, "cover");
    await addImage(s, `${ASSET}/photo-0.60-stable.jpg`, 770, 392, 380, 185, "cover");
  });

  add("十一、对策实施", (s) => {
    addText(s, "2、异常联动验证与记录闭环", 90, 130, 650, 32, { size: 29, bold: true });
    table(s, ["联动事项", "执行要求", "输出记录"], [
      ["上机确认", "POY来源线别、油剂、含油率、批量切换时间同步确认", "上机确认单"],
      ["异常验证", "打样结果、断头/C切、张力CV趋势当天反馈", "异常验证闭环表"],
      ["排除项归档", "线别切换、竹本油剂、DTY端临时调整作为证据保留", "末端因素确认表"],
      ["每日跟踪", "10#机断头≤90、C切≤50，任一超标当天复核", "每日目标跟踪表"],
    ], 80, 190, 1080, 82, [220, 600, 260], { headerFill: C.green, bodySize: 20, leftColumns: [0, 1, 2], boldColumns: [0] });
  });

  add("十二、效果检验", async (s) => {
    await addImage(s, TREND, 105, 130, 1070, 400, "contain");
    note(s, "5月25日0.60%整体执行后，断头和C切同步进入稳定区间，连续21天达成过程目标。", 160, 550, 920, 58);
  });

  add("十二、效果检验", (s) => {
    table(s, ["项目", "改善前/试调期", "改善后", "改善量", "改善率", "结论"], [
      ["断头数", "141.8个/日", "33.5个/日", "下降108.3个", "下降76.4%", "低于90目标"],
      ["C切数", "56.4个/日", "2.0个/日", "下降54.4个", "下降96.5%", "低于50目标"],
      ["连续达标", "未稳定", "连续21天", "提升明显", "100%达标", "目标完成"],
    ], 90, 170, 1100, 90, [150, 200, 190, 190, 170, 200], { headerFill: C.green, bodySize: 22, boldColumns: [0, 5] });
    note(s, "实施效果：围绕确认要因实施后，断头、C切两个过程指标均稳定达标，支撑55/72黑丝优等率提升。", 130, 510, 970, 76);
  });

  add("十三、巩固措施", (s) => {
    table(s, ["控制项目", "标准化/防措", "责任", "记录"], [
      ["POY含油率", "55/72黑丝按0.60%目标控制，低于目标暂停扩大生产", "前纺工艺", "POY含油确认表"],
      ["上机确认", "记录线别、油剂、含油率、打样锭位、切换时间", "加弹工艺", "上机确认单"],
      ["过程监控", "每日跟踪断头≤90、C切≤50，超标当天反馈", "班组/QC小组", "每日目标跟踪表"],
      ["异常联动", "前纺、加弹、质检三方确认异常项并闭环", "QC组长", "异常验证闭环表"],
    ], 80, 150, 1080, 88, [190, 520, 170, 200], { headerFill: C.green, bodySize: 20, leftColumns: [0, 1, 2, 3], boldColumns: [0] });
  });

  add("十四、总结及今后打算", (s) => {
    addText(s, "总结：", 110, 150, 160, 40, { size: 32, bold: true });
    addText(s, "1、通过本次QC活动，确认55/72黑丝优等率低的主要原因并形成有效对策。\n2、通过0.60%整体执行和异常联动闭环，断头与C切连续21天达成目标。\n3、非要因作为验证证据保留，避免后续复盘时误写成有效措施。", 150, 210, 940, 150, { size: 25, lineSpacing: 1.22 });
    shape(s, 120, 424, 1000, 4, { fill: "#8FD14F", stroke: "#8FD14F" });
    addText(s, "下一步打算：", 150, 470, 250, 42, { size: 32, bold: true });
    addText(s, "将55/72黑丝POY含油控制、上机确认、每日目标跟踪和防复发检查纳入常态化执行。", 150, 535, 930, 42, { size: 27 });
  });

  {
    const s = p.slides.add();
    slides.push(s);
    s.background.fill = C.white;
    await addImage(s, `${ASSET}/qcc-fiber-bg-imagegen.png`, 650, 0, 630, H, "cover");
    shape(s, 0, 0, 720, H, { fill: "#FFFFFFE8", stroke: "#FFFFFF00" });
    addText(s, "汇报完毕", 120, 250, 420, 62, { size: 58, bold: true, align: "center" });
    addText(s, "感谢观看", 148, 340, 360, 45, { size: 38, align: "center" });
    shape(s, 4, 0, 6, H, { fill: "#FFB68F", stroke: "#FFB68F" });
  }

  await Promise.all(pending);

  const total = p.slides.items.length;
  for (const [idx, slide] of p.slides.items.entries()) {
    await applyLogos(slide, idx + 1, total);
  }

  const final = `${OUT}/QCC成果报告_提高55-72黑丝优等率_真实QCC参考重做版.pptx`;
  for (const [idx, slide] of p.slides.items.entries()) {
    const stem = `real-qcc-slide-${String(idx + 1).padStart(2, "0")}`;
    const png = await p.export({ slide, format: "png", scale: 1 });
    await fs.writeFile(`${PREVIEW}/${stem}.png`, new Uint8Array(await png.arrayBuffer()));
  }
  const montage = await p.export({ format: "webp", montage: true, scale: 1 });
  await fs.writeFile(`${PREVIEW}/real-qcc-montage.webp`, new Uint8Array(await montage.arrayBuffer()));
  const inspect = await p.inspect({ kind: "slide,textbox,shape,image,table,chart", maxChars: 20000 });
  await fs.writeFile(`${QA}/real-qcc-inspect.ndjson`, inspect.ndjson);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(final);
  return final;
}

async function main() {
  await ensureDirs();
  const final = await buildDeck();
  console.log(final);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
