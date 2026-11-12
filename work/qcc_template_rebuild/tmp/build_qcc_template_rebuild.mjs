import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "/Users/USER/Documents/Codex/2026-06-21/new-chat-3";
const TMP = `${ROOT}/work/qcc_template_rebuild/tmp`;
const ASSET = `${ROOT}/work/qcc_template_rebuild/assets`;
const PREVIEW = `${TMP}/preview`;
const QA = `${TMP}/qa`;
const OUT = `${ROOT}/outputs`;

const COVER_BG = `${ASSET}/imagegen_tech_textile_cover.png`;
const BODY_BG = `${ASSET}/ref_media/image3.png`;
const ZHONGWEI_LOGO = `${ASSET}/zhongwei_logo_icon.png`;
const LOGO_ROOT = `${ROOT}/work/qcc_2026_generated/tmp/real_style/logo_assets/raw/word/media`;
const LOGO1 = `${LOGO_ROOT}/image1.png`;
const LOGO2 = `${LOGO_ROOT}/image2.png`;
const LOGO3 = `${LOGO_ROOT}/image3.png`;
const PHOTO_ROOT = `${ROOT}/work/presentations/qcc_55_72_heisi/tmp/assets`;
const TREND = `${ROOT}/outputs/qcc_2026_generated/近一个月断头_C切趋势图.png`;

const W = 1280;
const H = 720;
const FONT = "Microsoft YaHei";

const C = {
  white: "#FFFFFF",
  black: "#000000",
  cyan: "#18DFFF",
  cyan2: "#32B6E8",
  blue: "#0B5A8E",
  blue2: "#083D66",
  deep: "#061B2E",
  panel: "#092A48CC",
  panel2: "#0B4169CC",
  line: "#26C6FF",
  yellow: "#F8D23A",
  red: "#E83434",
  gray: "#D9E6F2",
  muted: "#9CB8CC",
  darkText: "#173047",
  green: "#4DD18B",
};

const daily = [
  ["5月09日", 151, 103, "0.45%", "CP5 12线POY上机"],
  ["5月10日", 143, 82, "0.45%", "上机初期"],
  ["5月11日", 180, 64, "0.45%", "调整油轮导丝器"],
  ["5月12日", 142, 82, "0.45%", "换至8线生产，仍未稳定"],
  ["5月13日", 190, 64, "0.45%", "高位波动"],
  ["5月14日", 141, 89, "0.45%", "高位波动"],
  ["5月15日", 243, 114, "0.45%", "张力CV波动明显"],
  ["5月16日", 180, 71, "0.45%", "换竹本油剂试验"],
  ["5月17日", 130, 50, "0.45%", "仍高位"],
  ["5月18日", 143, 54, "0.45%", "仍高位"],
  ["5月19日", 112, 34, "0.55%", "含油0.55%打样"],
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

const nonCauseRows = [
  ["人", "操作培训不到位", "班组执行口径基本一致，异常随工艺/POY阶段变化", "非要因"],
  ["机", "张力传感/报警阈值异常", "张力CV波动为表现，未见阈值调整带来持续改善", "非要因"],
  ["料", "前纺线别切换影响", "5月12日由CP5 12线换至8线后仍高位波动", "非要因"],
  ["料", "竹本油剂试验单独影响", "5月16日试验后5月16-18日未形成明显改善", "非要因"],
  ["法", "检验/分级口径不一致", "日报、系统记录、趋势图口径连续", "非要因"],
  ["环", "环境温湿度波动", "改善节点与0.60%整体执行一致，未见环境单独驱动", "非要因"],
];

async function ensureDirs() {
  await fs.mkdir(PREVIEW, { recursive: true });
  await fs.mkdir(QA, { recursive: true });
  await fs.mkdir(OUT, { recursive: true });
}

function text(slide, value, x, y, w, h, o = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = value;
  s.text.style = {
    typeface: FONT,
    fontSize: o.size ?? 22,
    bold: o.bold ?? false,
    color: o.color ?? C.white,
    alignment: o.align ?? "left",
    verticalAlignment: o.valign ?? "top",
    lineSpacing: o.lineSpacing ?? 1.12,
    autoFit: o.autoFit ?? "none",
    insets: o.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return s;
}

function shape(slide, x, y, w, h, o = {}) {
  return slide.shapes.add({
    geometry: o.geometry ?? "rect",
    position: { left: x, top: y, width: w, height: h },
    fill: o.fill ?? "none",
    line: { style: "solid", fill: o.stroke ?? "none", width: o.lineWidth ?? 0 },
    rotation: o.rotation,
    borderRadius: o.radius,
    shadow: o.shadow,
  });
}

async function img(slide, file, x, y, w, h, fit = "cover") {
  const bytes = await fs.readFile(file);
  const ext = path.extname(file).toLowerCase();
  const contentType = ext === ".png" ? "image/png" : "image/jpeg";
  slide.images.add({
    blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    contentType,
    position: { left: x, top: y, width: w, height: h },
    fit,
    alt: path.basename(file),
  });
}

async function base(slide, title, sectionNo, total, idx, o = {}) {
  await img(slide, o.cover ? COVER_BG : BODY_BG, 0, 0, W, H, "cover");
  if (!o.cover && !o.end) {
    await img(slide, ZHONGWEI_LOGO, 20, 20, 62, 62, "contain");
    text(slide, "桐乡市中维化纤有限公司", 84, 24, 300, 30, { size: 25, bold: true, color: C.white });
    text(slide, "TONGXIANG ZHONGWEI CHEMICAL FIBER CO.,LTD.", 86, 58, 330, 20, { size: 11.5, bold: true, color: C.white });
    shape(slide, 0, 84, 835, 2, { fill: C.line, stroke: C.line });
    shape(slide, 1168, 84, 112, 2, { fill: C.line, stroke: C.line });
    const headerTitle = sectionNo ? `${sectionNo}、${title}` : title;
    text(slide, headerTitle, 840, 24, 320, 48, { size: 39, bold: true, align: "right" });
  }
  if (o.cover || o.end) {
    await img(slide, LOGO2, 36, 22, 230, 92, "contain");
  }
  if (!o.cover && !o.end) {
    text(slide, "提供优质纤维        创造美好生活", 350, 665, 580, 40, { size: 28, color: C.gray, align: "center" });
  }
}

function panel(slide, x, y, w, h, o = {}) {
  shape(slide, x, y, w, h, {
    fill: o.fill ?? C.panel,
    stroke: o.stroke ?? "#1AA9D8",
    lineWidth: o.lineWidth ?? 1.2,
    radius: o.radius ?? "rounded-lg",
    shadow: o.shadow ?? "shadow-sm",
  });
}

function hex(slide, x, y, w, h, value, fill = C.blue2, size = 26) {
  shape(slide, x, y, w, h, { geometry: "hexagon", fill, stroke: "#42DDFE", lineWidth: 1.5, shadow: "shadow-sm" });
  text(slide, value, x, y + h * 0.28, w, h * 0.45, { size, bold: true, color: C.white, align: "center", valign: "middle" });
}

function pill(slide, label, x, y, w, fill = C.cyan2, color = C.white) {
  shape(slide, x, y, w, 42, { geometry: "roundRect", fill, stroke: "#4CEBFF", lineWidth: 1, radius: "rounded-xl" });
  text(slide, label, x + 8, y + 9, w - 16, 22, { size: 21, bold: true, color, align: "center" });
}

function callout(slide, value, x, y, w, h, fill = "#082946DD", stroke = "#2DDDFD") {
  panel(slide, x, y, w, h, { fill, stroke });
  text(slide, value, x + 18, y + 14, w - 36, h - 28, { size: 22, color: C.gray, lineSpacing: 1.18 });
}

function simpleTable(slide, headers, rows, x, y, w, rowH, widths, o = {}) {
  const colW = widths ?? headers.map(() => w / headers.length);
  const headerFill = o.headerFill ?? "#0B84B5";
  const bodyFill = o.bodyFill ?? "#06223ACC";
  const line = o.line ?? "#25A9D9";
  let cx = x;
  headers.forEach((h, i) => {
    shape(slide, cx, y, colW[i], rowH, { fill: headerFill, stroke: line, lineWidth: 1 });
    text(slide, h, cx + 6, y + 9, colW[i] - 12, rowH - 14, {
      size: o.headerSize ?? 18,
      bold: true,
      color: o.headerColor ?? C.white,
      align: "center",
      valign: "middle",
    });
    cx += colW[i];
  });
  rows.forEach((row, r) => {
    cx = x;
    row.forEach((v, i) => {
      const fill = r % 2 ? bodyFill : "#092D4ACC";
      shape(slide, cx, y + rowH * (r + 1), colW[i], rowH, { fill, stroke: line, lineWidth: 0.8 });
      text(slide, String(v ?? ""), cx + 7, y + rowH * (r + 1) + 8, colW[i] - 14, rowH - 12, {
        size: o.bodySize ?? 17,
        bold: (o.boldColumns ?? []).includes(i),
        color: o.bodyColor ?? C.gray,
        align: (o.leftColumns ?? []).includes(i) ? "left" : "center",
        valign: "middle",
        lineSpacing: 1.05,
      });
      cx += colW[i];
    });
  });
}

function bar(slide, x, y, w, h, label, val, max, color, valText) {
  const bw = Math.max(8, (w * val) / max);
  text(slide, label, x, y - 2, 150, 22, { size: 18, color: C.gray });
  shape(slide, x + 150, y, w, h, { fill: "#143A5A", stroke: "#1D79A8", lineWidth: 0.8 });
  shape(slide, x + 150, y, bw, h, { fill: color, stroke: color });
  text(slide, valText ?? String(val), x + 160 + bw, y - 2, 120, 22, { size: 18, bold: true, color: C.white });
}

function customPolygon(slide, points, fill, stroke, lineWidth = 2) {
  const xs = points.map((p) => p[0]);
  const ys = points.map((p) => p[1]);
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  const width = Math.max(...xs) - left;
  const height = Math.max(...ys) - top;
  slide.shapes.add({
    geometry: "custom",
    position: { left, top, width, height },
    fill,
    line: { style: "solid", fill: stroke, width: lineWidth },
    customPaths: [{
      width,
      height,
      commands: [
        { moveTo: { x: points[0][0] - left, y: points[0][1] - top } },
        ...points.slice(1).map(([x, y]) => ({ lineTo: { x: x - left, y: y - top } })),
        { close: {} },
      ],
    }],
  });
}

function lineBetween(slide, x1, y1, x2, y2, color = "#7CA8C8", width = 1.2) {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const len = Math.sqrt(dx * dx + dy * dy);
  shape(slide, x1, y1, len, width, { fill: color, stroke: color, rotation: Math.atan2(dy, dx) * 180 / Math.PI });
}

function arrowLine(slide, x1, y1, x2, y2, color = C.cyan, width = 3) {
  lineBetween(slide, x1, y1, x2, y2, color, width);
  const angle = Math.atan2(y2 - y1, x2 - x1) * 180 / Math.PI;
  shape(slide, x2 - 12, y2 - 12, 24, 24, {
    geometry: "triangle",
    fill: color,
    stroke: color,
    lineWidth: 0,
    rotation: angle + 90,
  });
}

function radarSummary(slide, cx, cy, radius) {
  const metrics = [
    ["优等率提升", 58, 85],
    ["断头C切控制", 42, 90],
    ["数据分析", 45, 82],
    ["标准化", 40, 80],
    ["QC经验", 45, 78],
    ["团队协作", 55, 86],
  ];
  const angles = metrics.map((_, i) => (-90 + i * 60) * Math.PI / 180);
  for (const r of [0.25, 0.5, 0.75, 1]) {
    const pts = angles.map((a) => [cx + Math.cos(a) * radius * r, cy + Math.sin(a) * radius * r]);
    customPolygon(slide, pts, "none", "#496D8C", r === 1 ? 2 : 1);
  }
  angles.forEach((a) => lineBetween(slide, cx, cy, cx + Math.cos(a) * radius, cy + Math.sin(a) * radius, "#496D8C", 1));
  const beforePts = angles.map((a, i) => [cx + Math.cos(a) * radius * metrics[i][1] / 100, cy + Math.sin(a) * radius * metrics[i][1] / 100]);
  const afterPts = angles.map((a, i) => [cx + Math.cos(a) * radius * metrics[i][2] / 100, cy + Math.sin(a) * radius * metrics[i][2] / 100]);
  customPolygon(slide, beforePts, "#C8D2DE55", "#C8D2DE", 3);
  customPolygon(slide, afterPts, "#00E5D466", "#70FFF0", 4);
  shape(slide, cx - 38, cy - 38, 76, 76, { geometry: "ellipse", fill: "#69FFF0CC", stroke: "#69FFF0", lineWidth: 1.5 });
  text(slide, "能力\n提升", cx - 32, cy - 22, 64, 44, { size: 17, bold: true, color: C.deep, align: "center", valign: "middle", lineSpacing: 1.0 });

  metrics.forEach(([name, before, after], i) => {
    const a = angles[i];
    const lx = cx + Math.cos(a) * (radius + 88);
    const ly = cy + Math.sin(a) * (radius + 64);
    text(slide, `${name}\n${before}→${after}`, lx - 82, ly - 24, 164, 48, {
      size: 17,
      bold: true,
      color: C.white,
      align: "center",
      lineSpacing: 1.0,
    });
  });
  panel(slide, 128, 168, 168, 62, { fill: "#00A996DD", stroke: "#70FFF0" });
  text(slide, "活动后", 128, 180, 168, 34, { size: 30, bold: true, color: C.white, align: "center" });
  panel(slide, 128, 500, 168, 62, { fill: "#546170CC", stroke: "#C8D2DE" });
  text(slide, "活动前", 128, 512, 168, 34, { size: 30, bold: true, color: C.white, align: "center" });
}

function miniTrend(slide, x, y, w, h) {
  const values = daily.map((r) => r[1]);
  const cvals = daily.map((r) => r[2]);
  const max = 260;
  shape(slide, x, y, w, h, { fill: "#061F35CC", stroke: "#1E9CCB", lineWidth: 1 });
  for (let i = 0; i <= 4; i++) {
    shape(slide, x + 42, y + 22 + (h - 55) * i / 4, w - 70, 1, { fill: "#1E4F70", stroke: "#1E4F70" });
  }
  const px = (i) => x + 50 + i * ((w - 92) / (values.length - 1));
  const py = (v) => y + h - 36 - (v / max) * (h - 70);
  values.forEach((v, i) => {
    const cx = px(i);
    const cy = py(v);
    shape(slide, cx - 2, cy - 2, 4, 4, { geometry: "ellipse", fill: C.cyan, stroke: C.cyan });
    if (i > 0) {
      const x0 = px(i - 1);
      const y0 = py(values[i - 1]);
      const dx = cx - x0;
      const dy = cy - y0;
      const len = Math.sqrt(dx * dx + dy * dy);
      shape(slide, x0, y0, len, 3, { fill: C.cyan, stroke: C.cyan, rotation: Math.atan2(dy, dx) * 180 / Math.PI });
    }
  });
  cvals.forEach((v, i) => {
    const cx = px(i);
    const cy = py(v);
    shape(slide, cx - 2, cy - 2, 4, 4, { geometry: "ellipse", fill: C.yellow, stroke: C.yellow });
    if (i > 0) {
      const x0 = px(i - 1);
      const y0 = py(cvals[i - 1]);
      const dx = cx - x0;
      const dy = cy - y0;
      const len = Math.sqrt(dx * dx + dy * dy);
      shape(slide, x0, y0, len, 3, { fill: C.yellow, stroke: C.yellow, rotation: Math.atan2(dy, dx) * 180 / Math.PI });
    }
  });
  text(slide, "断头", x + w - 130, y + 16, 48, 20, { size: 16, bold: true, color: C.cyan });
  text(slide, "C切", x + w - 76, y + 16, 48, 20, { size: 16, bold: true, color: C.yellow });
}

function fishbone(slide) {
  text(slide, "55/72黑丝\n优等率偏低", 105, 310, 150, 78, { size: 27, bold: true, color: C.white, align: "center", lineSpacing: 1.06 });
  shape(slide, 270, 370, 830, 3, { fill: C.cyan, stroke: C.cyan });
  const groups = [
    ["人", 336, 135, ["操作培训不到位", "异常反馈不及时"], [false, false]],
    ["机", 610, 115, ["油轮导丝器磨损", "张力报警阈值异常"], [true, false]],
    ["料", 870, 135, ["POY因素", "线别切换影响"], [true, false]],
    ["法", 470, 455, ["上热箱温度不匹配", "含油控制未固化"], [true, true]],
    ["环", 760, 455, ["环境温湿度波动", "现场状态波动"], [false, false]],
  ];
  groups.forEach(([name, x, y, items, flags]) => {
    const top = y < 370;
    hex(slide, x - 28, y + 18, 56, 50, name, "#1378A9", 25);
    shape(slide, x, top ? y + 70 : 370, 2, top ? 300 - y : y - 370, { fill: C.cyan, stroke: C.cyan });
    items.forEach((it, i) => {
      const bx = x + 52;
      const by = y + i * 54;
      shape(slide, bx, by, 190, 34, { fill: flags[i] ? "#F8D23ACC" : "#0B4169CC", stroke: flags[i] ? C.yellow : C.line, lineWidth: 1 });
      text(slide, it, bx + 8, by + 7, 174, 20, { size: 17, bold: flags[i], color: flags[i] ? C.white : C.gray, align: "center" });
      shape(slide, x + 28, by + 17, 24, 2, { fill: C.cyan, stroke: C.cyan });
    });
  });
}

function factorSlide(slide, num, factor, method, fact, verdict, isCause, percent, extra = "") {
  hex(slide, 54, 120, 112, 96, String(num), isCause ? C.red : "#1378A9", 34);
  text(slide, factor, 190, 127, 410, 38, { size: 30, bold: true, color: C.white });
  shape(slide, 928, 126, 142, 42, {
    fill: isCause ? C.red : C.white,
    stroke: isCause ? "#FF7676" : "#FFFFFF",
    lineWidth: 1.2,
  });
  text(slide, verdict, 940, 134, 118, 24, { size: 23, bold: true, color: isCause ? C.white : C.black, align: "center" });
  text(slide, "影响因素", 190, 172, 100, 24, { size: 18, color: C.cyan, bold: true });
  text(slide, fact, 300, 172, 830, 48, { size: 18, color: C.gray, lineSpacing: 1.15 });
  panel(slide, 80, 250, 270, 240, { fill: "#082946DD" });
  text(slide, `${percent}`, 105, 297, 220, 78, { size: 48, bold: true, color: isCause ? C.red : C.cyan, align: "center", valign: "middle" });
  text(slide, "影响占比/验证权重", 105, 382, 220, 26, { size: 18, color: C.gray, align: "center" });
  text(slide, `确认方法：${method}\n确认人：____\n确认时间：____`, 105, 426, 215, 54, { size: 17, color: C.gray, lineSpacing: 1.15 });
  panel(slide, 392, 250, 720, 240, { fill: "#061F35DD" });
  text(slide, "确认情况：", 420, 274, 130, 28, { size: 21, bold: true, color: C.cyan });
  text(slide, extra || "____", 420, 320, 650, 90, { size: 21, color: C.gray, lineSpacing: 1.2 });
  shape(slide, 930, 405, 138, 138, { geometry: "ellipse", fill: isCause ? "#E83434" : "#F8D23A", stroke: isCause ? "#FF7676" : "#FFF084", lineWidth: 2, shadow: "shadow-sm" });
  text(slide, verdict, 946, 449, 106, 44, { size: 29, bold: true, color: isCause ? C.white : C.black, align: "center", valign: "middle" });
}

async function writePreviews(pres, final) {
  for (const [idx, slide] of pres.slides.items.entries()) {
    const stem = `slide-${String(idx + 1).padStart(2, "0")}`;
    const png = await pres.export({ slide, format: "png", scale: 1 });
    await fs.writeFile(`${PREVIEW}/${stem}.png`, new Uint8Array(await png.arrayBuffer()));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${PREVIEW}/${stem}.layout.json`, await layout.text());
  }
  const montage = await pres.export({ format: "webp", montage: true, scale: 1 });
  await fs.writeFile(`${PREVIEW}/contact-sheet.webp`, new Uint8Array(await montage.arrayBuffer()));
  const inspect = await pres.inspect({ kind: "slide,textbox,shape,image,table,chart", maxChars: 30000 });
  await fs.writeFile(`${QA}/inspect.ndjson`, inspect.ndjson);
  const pptx = await PresentationFile.exportPptx(pres);
  await pptx.save(final);
}

async function main() {
  await ensureDirs();
  const total = 28;
  const p = Presentation.create({ slideSize: { width: W, height: H } });
  const slides = [];
  const add = async (title, sectionNo, cb) => {
    const s = p.slides.add();
    slides.push(s);
    await base(s, title, sectionNo, total, slides.length);
    await cb(s);
  };

  {
    const s = p.slides.add();
    slides.push(s);
    await base(s, "", "", total, slides.length, { cover: true });
    shape(s, 0, 0, W, H, { fill: "#000A18AA", stroke: "none" });
    await img(s, LOGO2, 36, 22, 230, 92, "contain");
    text(s, "提高55/72黑丝优等率", 88, 245, 620, 74, { size: 48, bold: true, color: C.white });
    shape(s, 92, 334, 320, 3, { fill: C.cyan, stroke: C.cyan });
    text(s, "长空之王QC小组", 92, 366, 410, 42, { size: 31, color: C.cyan, bold: true });
  }

  await add("目录", "", (s) => {
    const items = [
      ["1", "小组简介"], ["2", "课题选择"], ["3", "选题理由"], ["4", "现状调查"], ["5", "目标设定"],
      ["6", "原因分析"], ["7", "要因确认"], ["8", "对策实施"], ["9", "实施效果"], ["10", "效益核算"], ["11", "课题总结"],
    ];
    text(s, "目", 105, 300, 40, 40, { size: 30, bold: true, color: C.white, align: "center" });
    text(s, "录", 105, 350, 40, 40, { size: 30, bold: true, color: C.white, align: "center" });
    items.forEach((it, i) => {
      const x = i < 6 ? 260 : 690;
      const y = 112 + (i % 6) * 82;
      hex(s, x, y, 62, 54, it[0], i % 3 === 0 ? "#159BD1" : "#0B5A8E", 21);
      text(s, it[1], x + 82, y + 14, 230, 24, { size: 24, bold: true, color: C.gray });
    });
  });

  await add("小组简介", 1, (s) => {
    shape(s, 155, 168, 126, 126, { geometry: "ellipse", fill: "#48BDD5DD", stroke: "#62EAFD", lineWidth: 1.5 });
    text(s, "QC\n小组", 178, 202, 80, 62, { size: 30, bold: true, color: C.white, align: "center", valign: "middle", lineSpacing: 1.0 });
    lineBetween(s, 280, 205, 432, 160, "#4FAAD0", 1.3);
    lineBetween(s, 280, 232, 432, 232, "#4FAAD0", 1.3);
    lineBetween(s, 280, 260, 432, 300, "#4FAAD0", 1.3);
    text(s, "小组名称", 445, 145, 120, 26, { size: 22, bold: true, color: C.cyan, align: "right" });
    text(s, "长空之王 QC 攻关小组", 585, 139, 420, 36, { size: 31, bold: true, color: C.yellow });
    text(s, "课题名称", 445, 220, 120, 26, { size: 22, bold: true, color: C.cyan, align: "right" });
    text(s, "《提高55/72黑丝优等率》", 585, 212, 420, 38, { size: 29, bold: true, color: C.cyan });
    text(s, "活动时间", 445, 292, 120, 26, { size: 22, bold: true, color: C.cyan, align: "right" });
    text(s, "5月9日~6月17日", 585, 286, 250, 36, { size: 27, color: C.white });
    text(s, "组长", 820, 292, 70, 26, { size: 22, bold: true, color: C.cyan, align: "right" });
    text(s, "褚新良", 905, 286, 180, 36, { size: 27, color: C.white });
    const members = Array.from({ length: 10 }, (_, i) => [i + 1, "____", "____", "____", "____", "____"]);
    simpleTable(s, ["序号", "姓名", "性别", "文化程度", "组内职务", "岗位/技术职称"], members,
      135, 335, 1010, 29, [82, 155, 105, 175, 185, 308],
      { headerFill: "#0A83BD", bodyFill: "#EAF3FFDD", bodyColor: C.black, line: "#041926", bodySize: 18, headerSize: 19, leftColumns: [1, 5] });
  });

  await add("课题选择", 2, (s) => {
    panel(s, 78, 126, 515, 325, { fill: "#061F35D9" });
    text(s, "现状问题", 110, 150, 160, 30, { size: 26, bold: true, color: C.cyan });
    bar(s, 122, 230, 250, 28, "55/72黑丝", 58, 100, C.red, "不到60%");
    bar(s, 122, 310, 250, 28, "75/36黑丝", 89.76, 100, C.green, "89.76%");
    text(s, "在机75/36黑丝5月优等率89.76%。", 122, 386, 420, 30, { size: 21, bold: true, color: C.gray, lineSpacing: 1.08 });
    panel(s, 685, 126, 445, 325, { fill: "#061F35D9" });
    text(s, "课题确定", 720, 150, 160, 30, { size: 26, bold: true, color: C.cyan });
    text(s, "提高55/72黑丝优等率", 735, 244, 330, 48, { size: 34, bold: true, color: C.white, align: "center" });
    callout(s, "选择理由：55/72黑丝刚上机后优等率不到60%，与同类在机黑丝品种存在明显差距，需立项确认原因并提升优等率。", 145, 505, 990, 78);
  });

  await add("选题理由", 3, async (s) => {
    await img(s, `${PHOTO_ROOT}/cv-wave-crop-sheet.jpg`, 70, 132, 500, 270, "cover");
    await img(s, `${PHOTO_ROOT}/photo-initial-machine.jpg`, 640, 132, 500, 270, "cover");
    text(s, "刚上机阶段，张力CV波动、断头、C切异常同步出现。", 105, 430, 1000, 36, { size: 28, bold: true, color: C.white, align: "center" });
    callout(s, "现场照片用于辅证：55/72黑丝上机初期稳定性不足，优等率与同类75/36黑丝存在明显差距，是本次选题的直接现状依据。", 160, 510, 900, 64);
  });

  await add("现状调查", 4, (s) => {
    simpleTable(s, ["调查项目", "调查时间", "现状记录", "判断"], [
      ["品种/机台", "5月9日上机", "10#机55/72黑丝刚上机，生产稳定性不足", "需立项改善"],
      ["优等率", "上机初期", "优等率不到60%，低于在机75/36黑丝5月89.76%", "差距明显"],
      ["张力CV", "刚上机观察", "多锭位张力CV出现尖峰波动", "影响加工稳定"],
      ["断头/C切", "刚上机观察", "断头、C切异常同步出现", "作为现状辅证"],
    ], 76, 140, 1130, 76, [190, 190, 560, 190], { bodySize: 20, boldColumns: [0, 3], leftColumns: [0, 2, 3] });
    callout(s, "结论：刚上机状态下，55/72黑丝优等率低，张力CV波动明显，断头/C切异常同步存在。", 155, 505, 970, 62);
  });

  await add("目标设定", 5, (s) => {
    const xs = [165, 528, 890];
    const titles = ["改善前", "目标", "改善后"];
    const vals = ["断头141.8个/日\nC切56.4个/日", "断头≤90个/日\nC切≤50个/日", "断头33.5个/日\nC切2.0个/日"];
    xs.forEach((x, i) => {
      hex(s, x, 172 + i * 42, 220, 190, titles[i], i === 1 ? "#1279A8" : i === 2 ? "#0DAA79" : "#A94444", 32);
      text(s, vals[i], x + 18, 265 + i * 42, 184, 60, { size: 23, bold: true, color: C.white, align: "center", lineSpacing: 1.12 });
    });
    callout(s, "目标：以断头≤90个/日、C切≤50个/日为过程控制目标，支撑55/72黑丝优等率达到85%以上。", 155, 540, 970, 66);
  });

  await add("鱼刺图分析", 6, (s) => fishbone(s));

  await add("要因确认", 7, (s) => {
    factorSlide(s, 1, "人：操作培训不到位", "班组访谈/现场记录核对", "人员因素从操作一致性、异常反馈及时性两个方面确认。", "非要因", false, "非要因", "同一班组、同一操作口径下异常仍持续出现；异常变化与POY和工艺参数更相关，未发现人员操作差异导致优等率低。");
  });

  await add("要因确认", 7, (s) => {
    factorSlide(s, 2, "机：油轮导丝器磨损", "现场确认/更换验证", "油轮导丝器磨损会增大丝束运行摩擦，造成张力波动并放大断头风险。", "要因", true, "要因", "现场确认存在导丝器磨损。更换旋转导丝器后，张力波动和断头状态有改善，因此判定为设备侧要因。");
  });

  await add("要因确认", 7, (s) => {
    factorSlide(s, 3, "机：张力传感/报警阈值异常", "系统记录核对/报警记录复核", "张力CV异常是现场表现之一，需要确认是否由传感或报警阈值异常造成。", "非要因", false, "非要因", "未发现传感器或报警阈值调整能带来持续改善。张力CV尖峰随生产条件变化出现，判定为异常表现，不作为本次主控要因。");
  });

  await add("要因确认", 7, (s) => {
    factorSlide(s, 4, "料：POY因素", "油嘴、油剂、含油率分项验证", "POY因素对本课题影响最大，其中含油率变化对加工稳定性影响最明显。", "要因", true, "80%", "POY要因内嵌三个小因素：更换油嘴有效但不明显；更换油剂有效但不明显；提高含油率改善最明显，最终以0.60%作为主控方向。");
  });

  await add("要因确认", 7, (s) => {
    text(s, "POY要因内部小因素验证", 88, 124, 430, 36, { size: 32, bold: true, color: C.red });
    simpleTable(s, ["小因素", "验证动作", "验证结果", "结论"], [
      ["油嘴", "更换油嘴", "有改善，但改善幅度不明显", "辅助改善项"],
      ["油剂", "更换油剂/竹本油剂试验", "有改善，但不作为最终主控点", "辅助改善项"],
      ["含油率", "0.45%到0.55%，再到0.60%", "0.60%后断头、C切趋于稳定", "核心改善项"],
      ["线别", "5月9日CP5 12线，5月12日换8线", "换线后仍未稳定", "非主控项"],
    ], 76, 190, 1125, 74, [185, 310, 400, 230], { bodySize: 20, boldColumns: [0, 3], leftColumns: [0, 1, 2, 3] });
    callout(s, "确认口径：POY因素为主要要因，最终主控点为提高并固化含油率。", 155, 575, 970, 52);
  });

  await add("要因确认", 7, (s) => {
    factorSlide(s, 5, "料：前纺线别切换影响", "生产履历核对/切换前后对比", "5月9日上机使用前纺CP5 12线POY，5月12日后切换到8线生产。", "非要因", false, "非要因", "切换到8线后，55/72黑丝仍存在张力CV波动和优等率偏低问题，说明线别切换不是决定性原因。");
  });

  await add("要因确认", 7, (s) => {
    factorSlide(s, 6, "料：竹本油剂试验单独影响", "油剂试验/阶段数据对比", "5月16日进行竹本油剂试验，需要确认油剂变化是否能单独解决异常。", "非要因", false, "非要因", "试验后5月16日至5月18日仍未形成明显改善，油剂变化单独作用有限，作为POY辅助因素保留。");
  });

  await add("要因确认", 7, (s) => {
    factorSlide(s, 7, "法：上热箱温度不匹配", "180/175/170/165/160分档试验", "上热箱温度与55/72黑丝加工稳定性相关，需要通过分档试验确认匹配参数。", "要因", true, "要因", "温度验证：180、175、170、165、160分档对比；综合断头、C切、张力稳定性，最终选择165。具体记录留空补充：____");
  });

  await add("要因确认", 7, (s) => {
    factorSlide(s, 8, "法：检验/分级口径不一致", "日报/系统记录/检验口径复核", "优等率低需要排除检验口径或分级口径变化造成的数据偏差。", "非要因", false, "非要因", "日报、系统记录、趋势图口径连续，未发现因检验或分级口径变化造成优等率异常。");
  });

  await add("要因确认", 7, (s) => {
    factorSlide(s, 9, "环：环境温湿度波动", "环境记录核对/改善节点对比", "环境温湿度可能影响纺丝和加弹稳定性，需要作为环境维度确认。", "非要因", false, "非要因", "未发现温湿度波动与异常改善节点形成一致对应关系，改善节点更贴近POY含油率和工艺参数变化。");
  });

  await add("要因确认", 7, (s) => {
    factorSlide(s, 10, "环：现场状态波动", "现场巡检/异常记录核对", "上机初期现场状态波动需要确认是否为独立环境要因。", "非要因", false, "非要因", "现场状态波动与刚上机阶段同时出现，但不能单独解释后续变化；作为现象背景保留，不作为主控要因。");
  });

  await add("要因确认", 7, (s) => {
    text(s, "要因确认汇总", 92, 128, 360, 36, { size: 32, bold: true, color: C.cyan });
    simpleTable(s, ["维度", "末端因素", "验证结论", "判定"], [
      ["人", "操作培训不到位", "操作一致性未发现决定性差异", "非要因"],
      ["机", "油轮导丝器磨损", "更换后状态改善", "要因"],
      ["机", "张力传感/报警阈值异常", "属于异常表现，未形成独立改善", "非要因"],
      ["料", "POY因素", "含油率改善最明显", "要因"],
      ["料", "前纺线别切换影响", "换8线后仍未稳定", "非要因"],
      ["料", "竹本油剂试验单独影响", "5月16日试验后无明显改善", "非要因"],
      ["法", "上热箱温度不匹配", "分档验证后选择165", "要因"],
      ["法", "检验/分级口径不一致", "记录口径连续一致", "非要因"],
      ["环", "环境温湿度/现场状态", "与改善节点不一致", "非要因"],
    ], 78, 174, 1130, 49, [110, 285, 535, 200], { bodySize: 17, boldColumns: [0, 1, 3], leftColumns: [1, 2] });
    callout(s, "最终判定要因：油轮导丝器磨损、POY因素、上热箱温度不匹配。其中POY因素为主要要因，含油率为主控小因素。", 145, 590, 990, 54);
  });

  await add("对策实施", 8, (s) => {
    simpleTable(s, ["要因", "目标对策", "措施", "负责人", "完成时间"], [
      ["油轮导丝器磨损", "减少摩擦与张力波动", "更换旋转导丝器，纳入点检", "____", "____"],
      ["POY因素", "含油率稳定到有效区间", "油嘴/油剂优化，重点加含油率至0.60%", "____", "____"],
      ["上热箱温度不匹配", "温度参数稳定匹配", "按165°执行并记录断头/C切", "____", "____"],
    ], 70, 150, 1140, 92, [220, 250, 390, 140, 140], { bodySize: 19, boldColumns: [0], leftColumns: [0, 1, 2] });
    callout(s, "对策原则：有效但不明显的措施作为辅助项，改善最明显的含油率调整作为主控项。", 150, 515, 960, 64);
  });

  await add("对策实施", 8, async (s) => {
    text(s, "POY要因措施展开", 86, 125, 310, 34, { size: 31, bold: true, color: C.cyan });
    simpleTable(s, ["时间/阶段", "措施", "效果判断"], [
      ["5月19日", "含油率0.45%调整至0.55%并打样", "有效，但不明显"],
      ["5月21日", "含油率0.60%打样10锭", "改善方向明确"],
      ["5月25日", "0.60%整体做到", "断头、C切趋于稳定"],
    ], 78, 184, 650, 82, [150, 330, 170], { bodySize: 20, boldColumns: [0], leftColumns: [1, 2] });
    await img(s, `${PHOTO_ROOT}/photo-0.55-oil.jpg`, 765, 152, 380, 178, "cover");
    await img(s, `${PHOTO_ROOT}/photo-0.60-stable.jpg`, 765, 372, 380, 178, "cover");
  });

  await add("实施效果", 9, async (s) => {
    await img(s, TREND, 92, 126, 1095, 400, "contain");
    callout(s, "5月25日0.60%整体执行后，断头、C切同步进入稳定区间；后续具体优等率数据可在空白处补录。优等率：____", 145, 548, 990, 64);
  });

  await add("实施效果", 9, (s) => {
    simpleTable(s, ["项目", "改善前/试调期", "改善后", "改善量", "改善率", "结论"], [
      ["断头数", "141.8个/日", "33.5个/日", "下降108.3个", "76.4%", "低于目标"],
      ["C切数", "56.4个/日", "2.0个/日", "下降54.4个", "96.5%", "低于目标"],
      ["过程稳定", "高位波动", "连续21天达标", "明显改善", "____", "攻关有效"],
    ], 82, 170, 1110, 90, [135, 205, 190, 205, 160, 215], { bodySize: 22, boldColumns: [0, 5], leftColumns: [0, 5] });
    shape(s, 995, 545, 116, 116, { geometry: "ellipse", fill: "#183E5DCC", stroke: C.yellow, lineWidth: 2 });
    text(s, "攻关\n成功", 1012, 574, 82, 52, { size: 27, bold: true, color: C.yellow, align: "center", lineSpacing: 1.0 });
  });

  await add("效益核算", 10, (s) => {
    simpleTable(s, ["项目", "计算口径", "金额"], [
      ["一等品转优等品", "1.45吨/日 × 18.0% × 2000元/吨", "522元/日"],
      ["合格品转优等品", "1.45吨/日 × 8.0% × 5550元/吨", "643.8元/日"],
      ["年化效益", "（522+643.8）×365", "约42.55万元"],
    ], 115, 170, 1050, 92, [250, 545, 255], { bodySize: 24, boldColumns: [0, 2], leftColumns: [0, 1] });
    callout(s, "财务签字：____      日期：____      签章：____", 205, 520, 820, 60);
  });

  await add("巩固措施", 10, (s) => {
    const items = [
      ["1", "固化55/72黑丝POY含油率控制标准，按0.60%作为稳定生产主控目标。"],
      ["2", "将旋转导丝器状态纳入日常点检，异常磨损及时更换。"],
      ["3", "上热箱温度按165°执行，变更需同步记录断头/C切验证结果。"],
      ["4", "建立前纺-加弹-质检联动验证表，异常当天形成闭环。"],
    ];
    items.forEach((it, i) => {
      hex(s, 112, 132 + i * 110, 70, 62, it[0], i === 0 ? "#14A8D7" : "#0B5A8E", 24);
      panel(s, 205, 136 + i * 110, 880, 55, { fill: "#082946DD" });
      text(s, it[1], 230, 151 + i * 110, 830, 26, { size: 21, color: C.gray });
    });
  });

  await add("课题总结", 11, (s) => {
    radarSummary(s, 590, 335, 160);
    simpleTable(s, ["项目", "活动前", "活动后", "提升"], [
      ["优等率提升", "58", "85", "+27"],
      ["断头C切控制", "42", "90", "+48"],
      ["数据分析", "45", "82", "+37"],
      ["标准化", "40", "80", "+40"],
      ["QC经验", "45", "78", "+33"],
      ["团队协作", "55", "86", "+31"],
    ], 875, 178, 285, 42, [115, 55, 55, 60], { headerSize: 16, bodySize: 16, bodyColor: C.gray, boldColumns: [0, 3], leftColumns: [0] });
    callout(s, "本次活动后，质量改善、数据分析、标准化和团队协作能力均有提升，形成可复制的55/72黑丝改善方法。", 175, 585, 930, 58);
  });

  {
    const s = p.slides.add();
    slides.push(s);
    await base(s, "", "", total, slides.length, { end: true });
    shape(s, 0, 0, W, H, { fill: "#000A18AA", stroke: "none" });
    await img(s, LOGO2, 36, 22, 230, 92, "contain");
    text(s, "汇报完毕", 420, 300, 440, 56, { size: 56, bold: true, color: C.cyan, align: "center" });
    text(s, "感谢观看", 488, 382, 300, 40, { size: 34, color: C.white, align: "center" });
  }

  const final = `${OUT}/55-72-QCC-封面鱼刺图修改版.pptx`;
  await writePreviews(p, final);
  console.log(final);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
