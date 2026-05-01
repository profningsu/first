const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const outDir = path.resolve('/Users/ning/.hermes/hermes-agent/output');
fs.mkdirSync(outDir, { recursive: true });

const pptxPath = path.join(outDir, 'semiconductor_patent_strategy_deck_zh-TW.pptx');
const previewPath = path.join(outDir, 'semiconductor_patent_strategy_deck_preview.html');

const C = {
  navy: '0B1F33',
  navy2: '132B45',
  slate: '5B6B7F',
  ink: '122033',
  text: '233143',
  muted: '6E7D90',
  line: 'D8E1EA',
  mist: 'F3F7FA',
  white: 'FFFFFF',
  teal: '0F8B8D',
  teal2: '16A6A8',
  aqua: '74D3D4',
  gold: 'E5B85C',
  amber: 'F2C14E',
  rose: 'D96C75',
  green: '4BAE8D',
  blue: '3C78D8',
};

const theme = {
  headFont: 'Aptos Display',
  bodyFont: 'Aptos',
  title: '2021–2025 全球廣義半導體專利：技術與管理意涵',
  subtitle: '基於既有分析報告整理｜用途：高階主管／技術策略／研發管理',
};

const growthData = [
  { label: 'chiplet', start: 1410, end: 8964, color: C.teal },
  { label: 'hybrid bonding', start: 2902, end: 8839, color: C.aqua },
  { label: 'backside power', start: 556, end: 4055, color: C.gold },
  { label: 'HBM', start: 10025, end: 27205, color: C.blue },
  { label: 'EUV', start: 23146, end: 30707, color: C.rose },
];

const coreLeaders = ['Intel', 'TSMC', 'Samsung', 'Adeia', 'Micron', 'Apple', 'Amkor'];
const fabLeaders = ['TSMC', 'ASM', 'Lam Research', 'SanDisk/WD', 'Applied Materials'];
const whitespace = [
  'Chiplet 系統層協同',
  'Hybrid bonding 量產化輔助',
  'Backside power 的 DFT / reliability',
  '先進封裝材料與中介層方案',
  'SiC / GaN 功率器件模組化',
  'Yield-learning / rework / KGD 流程工具鏈',
];

function fmt(n) {
  return Number(n).toLocaleString('en-US');
}
function mul(a, b) {
  return (b / a).toFixed(1) + '×';
}
function shadow() {
  return { type: 'outer', color: '000000', blur: 2, angle: 45, offset: 1, opacity: 0.12 };
}

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';
pptx.author = 'Hermes';
pptx.company = 'Hermes Agent';
pptx.subject = 'Semiconductor patent presentation';
pptx.title = theme.title;
pptx.lang = 'zh-TW';
pptx.theme = {
  headFontFace: theme.headFont,
  bodyFontFace: theme.bodyFont,
  lang: 'zh-TW'
};

function addFooter(slide, page, dark = false) {
  slide.addText(`方法邊界：Google Patents page；用於相對排序、熱點辨識與策略判讀`, {
    x: 0.55, y: 4.92, w: 7.05, h: 0.14,
    fontFace: theme.bodyFont, fontSize: 7.7,
    color: dark ? 'D2DCE6' : C.muted, margin: 0, fit: 'shrink'
  });
  slide.addText(String(page).padStart(2,'0'), {
    x: 8.93, y: 4.86, w: 0.3, h: 0.16, align: 'center', valign: 'mid',
    fontFace: theme.headFont, fontSize: 10.1, bold: true,
    color: dark ? C.white : C.navy, margin: 0
  });
}

function addTitle(slide, title, kicker, dark = false) {
  slide.addText(kicker, {
    x: 0.55, y: 0.34, w: 4.2, h: 0.22,
    fontFace: theme.bodyFont, fontSize: 11.2, bold: true, color: dark ? C.gold : C.teal,
    charSpacing: 1.2, margin: 0
  });
  slide.addText(title, {
    x: 0.55, y: 0.56, w: 8.7, h: 0.58,
    fontFace: theme.headFont, fontSize: 23.5, bold: true,
    color: dark ? C.white : C.ink, margin: 0, fit: 'shrink'
  });
}

function card(slide, x, y, w, h, opts = {}) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    line: { color: opts.line || '000000', transparency: opts.noLine ? 100 : 100, width: 0.5 },
    fill: { color: opts.fill || C.white, transparency: opts.transparency || 0 },
    radius: 0,
    shadow: opts.shadow ? shadow() : undefined,
  });
  if (opts.accent) {
    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: 0.09, h,
      line: { color: opts.accent, transparency: 100 },
      fill: { color: opts.accent }
    });
  }
}

function pill(slide, x, y, w, h, text, fill, color = C.white) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.06,
    line: { color: fill, transparency: 100 },
    fill: { color: fill }
  });
  slide.addText(text, {
    x, y: y + 0.04, w, h: h - 0.03,
    fontFace: theme.bodyFont, fontSize: 10, bold: true, color, align: 'center', margin: 0
  });
}

function addBullets(slide, items, x, y, w, h, color = C.text, size = 13.5, gap = 14) {
  const runs = [];
  items.forEach((t, i) => {
    runs.push({ text: t, options: { bullet: true, breakLine: i !== items.length - 1, color, hanging: 3, paraSpaceAfterPt: gap } });
  });
  slide.addText(runs, {
    x, y, w, h,
    fontFace: theme.bodyFont, fontSize: size,
    color, margin: 0.04,
    breakLine: true,
    valign: 'top'
  });
}

function addMiniStat(slide, x, y, w, label, start, end, color) {
  card(slide, x, y, w, 1.08, { fill: C.white, shadow: true, accent: color });
  slide.addText(label, {
    x: x + 0.18, y: y + 0.12, w: w - 0.26, h: 0.18,
    fontFace: theme.bodyFont, fontSize: 10.8, bold: true, color: C.muted, margin: 0
  });
  slide.addText(mul(start, end), {
    x: x + 0.18, y: y + 0.29, w: w - 0.26, h: 0.26,
    fontFace: theme.headFont, fontSize: 22, bold: true, color, margin: 0
  });
  slide.addText(`${fmt(start)} → ${fmt(end)}`, {
    x: x + 0.18, y: y + 0.68, w: w - 0.26, h: 0.16,
    fontFace: theme.bodyFont, fontSize: 10.4, color: C.text, margin: 0
  });
}

function addSourceNote(slide, x, y, w, text, dark = false) {
  slide.addText(text, {
    x, y, w, h: 0.22,
    fontFace: theme.bodyFont, fontSize: 8.5,
    italic: true, color: dark ? 'D2DCE6' : C.muted, margin: 0
  });
}

function slide1() {
  const slide = pptx.addSlide();
  slide.background = { color: C.navy };
  slide.addShape(pptx.ShapeType.rect, { x: 6.2, y: 0, w: 3.8, h: 5.625, line: { color: C.navy2, transparency: 100 }, fill: { color: C.navy2 } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 4.65, w: 10, h: 0.95, line: { color: C.teal, transparency: 100 }, fill: { color: C.teal, transparency: 65 } });

  slide.addText('EXECUTIVE VIEW', { x: 0.55, y: 0.42, w: 2.2, h: 0.2, fontFace: theme.bodyFont, fontSize: 10.5, bold: true, color: C.gold, charSpacing: 1.6, margin: 0 });
  slide.addText(theme.title, { x: 0.55, y: 0.7, w: 5.3, h: 0.72, fontFace: theme.headFont, fontSize: 26, bold: true, color: C.white, margin: 0 });
  slide.addText('技術競爭主軸已從單點製程，轉向「封裝 × 系統 × 製造整合」', { x: 0.55, y: 1.62, w: 5.1, h: 0.48, fontFace: theme.headFont, fontSize: 21, bold: true, color: 'DDEAF5', margin: 0 });
  slide.addText(theme.subtitle, { x: 0.55, y: 2.18, w: 4.9, h: 0.28, fontFace: theme.bodyFont, fontSize: 11.5, color: 'C8D4E2', margin: 0 });

  addBullets(slide, [
    '最穩定領先群：TSMC、Intel、Samsung',
    '第二梯隊依主題分化：Micron、ASM、Lam Research、Applied Materials、Adeia、Amkor、ASML / Carl Zeiss SMT',
    '高成長主題：chiplet、hybrid bonding、backside power、HBM、EUV'
  ], 0.62, 2.65, 5.0, 1.55, 'EAF2F8', 13, 10);

  addMiniStat(slide, 6.45, 0.78, 1.42, 'chiplet', 1410, 8964, C.teal2);
  addMiniStat(slide, 8.0, 0.78, 1.42, 'hybrid', 2902, 8839, C.aqua);
  addMiniStat(slide, 6.45, 2.02, 1.42, 'backside', 556, 4055, C.gold);
  addMiniStat(slide, 8.0, 2.02, 1.42, 'HBM', 10025, 27205, C.blue);
  addMiniStat(slide, 6.45, 3.26, 2.97, 'EUV', 23146, 30707, C.rose);

  pill(slide, 6.45, 4.42, 1.12, 0.32, 'Design / IC', C.blue);
  pill(slide, 7.62, 4.42, 1.05, 0.32, 'Memory', C.rose);
  pill(slide, 8.72, 4.42, 1.0, 0.32, 'Package', C.teal2);
  pill(slide, 6.45, 4.8, 1.12, 0.28, 'Foundry', C.gold, C.navy);
  pill(slide, 7.62, 4.8, 1.05, 0.28, 'Equipment', '8CA7BF', C.navy);
  pill(slide, 8.72, 4.8, 1.0, 0.28, 'Optics', 'AFC8D6', C.navy);

  addSourceNote(slide, 0.55, 4.34, 4.9, '重點不是單一節點微縮，而是 package / memory / process / equipment 的協同能力。', true);
  addFooter(slide, 1, true);
}

function slide2() {
  const slide = pptx.addSlide();
  slide.background = { color: C.mist };
  addTitle(slide, '產業鏈雙核心：器件/封裝/IC 與製程/設備同步升級', 'COMPETITIVE LANDSCAPE');

  card(slide, 0.55, 1.2, 8.85, 3.42, { fill: C.white, shadow: true });
  slide.addText('用關係圖而不是名單看競爭：需求牽引在左、製造供給在右，TSMC 位於少數可跨鏈整合的樞紐。', {
    x: 0.82, y: 1.34, w: 7.7, h: 0.22,
    fontFace: theme.bodyFont, fontSize: 10.8, color: C.text, margin: 0
  });

  slide.addText('需求 / 系統側', { x: 0.86, y: 1.72, w: 1.15, h: 0.16, fontFace: theme.bodyFont, fontSize: 10.2, bold: true, color: C.teal, margin: 0 });
  slide.addText('製造 / 設備側', { x: 7.32, y: 1.72, w: 1.2, h: 0.16, fontFace: theme.bodyFont, fontSize: 10.2, bold: true, color: C.gold, margin: 0 });

  const leftNodes = [
    ['Intel', 0.95, 2.02, 1.18, 0.44, true],
    ['Samsung', 0.95, 2.62, 1.18, 0.44, true],
    ['Micron', 0.95, 3.22, 1.18, 0.44, false],
    ['Apple', 2.28, 2.02, 1.05, 0.44, false],
    ['Adeia', 2.28, 2.62, 1.05, 0.44, false],
    ['Amkor', 2.28, 3.22, 1.05, 0.44, false],
  ];
  leftNodes.forEach(([label, x, y, w, h, primary]) => {
    card(slide, x, y, w, h, { fill: primary ? 'EAF7F7' : C.mist, accent: primary ? C.teal : C.aqua });
    slide.addText(label, { x: x + 0.14, y: y + 0.14, w: w - 0.2, h: 0.14, fontFace: theme.bodyFont, fontSize: 10.8, bold: primary, color: C.ink, align: 'center', margin: 0 });
  });

  const rightNodes = [
    ['ASM', 7.0, 2.02, 1.02, 0.44, true],
    ['Lam Research', 7.0, 2.62, 1.56, 0.44, true],
    ['Applied Materials', 7.0, 3.22, 1.72, 0.44, true],
    ['ASML / Carl Zeiss', 8.18, 2.02, 0.98, 0.44, false],
    ['SanDisk / WD', 8.18, 2.62, 0.98, 0.44, false],
  ];
  rightNodes.forEach(([label, x, y, w, h, primary]) => {
    card(slide, x, y, w, h, { fill: primary ? 'FFF7E7' : C.mist, accent: primary ? C.gold : C.rose });
    slide.addText(label, { x: x + 0.07, y: y + 0.12, w: w - 0.14, h: 0.17, fontFace: theme.bodyFont, fontSize: primary ? 10.0 : 9.5, bold: primary, color: C.ink, align: 'center', margin: 0, fit: 'shrink' });
  });

  card(slide, 4.05, 2.08, 1.85, 1.68, { fill: C.navy, shadow: true });
  slide.addText('TSMC', { x: 4.33, y: 2.4, w: 1.26, h: 0.28, fontFace: theme.headFont, fontSize: 25, bold: true, color: C.white, align: 'center', margin: 0 });
  slide.addText('device × package × fab 的跨鏈整合樞紐', { x: 4.2, y: 2.88, w: 1.56, h: 0.32, fontFace: theme.bodyFont, fontSize: 9.4, color: 'DCE7F2', align: 'center', margin: 0 });
  pill(slide, 4.28, 3.32, 0.62, 0.24, 'Core', C.teal);
  pill(slide, 5.0, 3.32, 0.62, 0.24, 'Fab', C.gold, C.navy);

  slide.addShape(pptx.ShapeType.chevron, { x: 3.5, y: 2.46, w: 0.42, h: 0.3, line: { color: C.teal, transparency: 100 }, fill: { color: C.teal } });
  slide.addShape(pptx.ShapeType.chevron, { x: 3.5, y: 3.06, w: 0.42, h: 0.3, line: { color: C.teal, transparency: 100 }, fill: { color: C.teal } });
  slide.addShape(pptx.ShapeType.chevron, { x: 6.0, y: 2.46, w: 0.42, h: 0.3, line: { color: C.gold, transparency: 100 }, fill: { color: C.gold } });
  slide.addShape(pptx.ShapeType.chevron, { x: 6.0, y: 3.06, w: 0.42, h: 0.3, line: { color: C.gold, transparency: 100 }, fill: { color: C.gold } });

  card(slide, 0.86, 3.92, 2.35, 0.48, { fill: 'EAF7F7', accent: C.teal });
  slide.addText('需求拉動：AI 平台、記憶體、封裝協同把產品 roadmap 往系統整合推。', { x: 1.02, y: 4.03, w: 2.02, h: 0.22, fontFace: theme.bodyFont, fontSize: 9.1, color: C.text, margin: 0, fit: 'shrink' });
  card(slide, 3.52, 3.92, 2.52, 0.48, { fill: 'EEF3F8', accent: C.blue });
  slide.addText('整合樞紐：少數玩家能同時掌握客戶需求、節點節奏與封裝協同。', { x: 3.7, y: 4.03, w: 2.16, h: 0.22, fontFace: theme.bodyFont, fontSize: 9.1, color: C.text, margin: 0, fit: 'shrink' });
  card(slide, 6.34, 3.92, 2.64, 0.48, { fill: 'FFF7E7', accent: C.gold });
  slide.addText('供給主導：設備 / 材料商正在鎖定未來節點 know-how 與議價權。', { x: 6.52, y: 4.03, w: 2.28, h: 0.22, fontFace: theme.bodyFont, fontSize: 9.1, color: C.text, margin: 0, fit: 'shrink' });

  slide.addText('管理實務意涵', { x: 0.55, y: 4.56, w: 1.1, h: 0.18, fontFace: theme.bodyFont, fontSize: 10.3, bold: true, color: C.teal, margin: 0 });
  addBullets(slide, [
    '競爭者地圖至少拆成「產品 / 封裝 / 設備 / 材料」四層，不要只看總 patent count。',
    '共同開發談判要把設備與材料商的 know-how/IP 邊界提早納入。'
  ], 1.55, 4.5, 7.55, 0.46, C.text, 10.0, 2);
  addFooter(slide, 2, false);
}

function addGrowthChart(slide, x, y, w, h, subset, title) {
  slide.addChart(pptx.ChartType.line, subset.map((s) => ({ name: s.label, labels: ['2021', '2025'], values: [s.start, s.end] })), {
    x, y, w, h,
    showTitle: true,
    title,
    titleFontFace: theme.bodyFont,
    titleFontSize: 12,
    titleColor: C.ink,
    lineSize: 3,
    lineSmooth: false,
    showLegend: true,
    legendPos: 'b',
    legendFontSize: 10,
    chartColors: subset.map((s) => s.color),
    showCatName: false,
    showValAxisTitle: false,
    showCatAxisTitle: false,
    valAxisLabelColor: C.muted,
    catAxisLabelColor: C.muted,
    valGridLine: { color: C.line, size: 0.6 },
    catGridLine: { color: C.line, transparency: 100 },
    showValue: true,
    dataLabelPosition: 't',
    dataLabelColor: C.ink,
    chartArea: { fill: { color: C.white }, border: { color: C.white } },
    showBorder: false,
  });
}

function slide3() {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addTitle(slide, '封裝已從成本中心變成架構創新中心', 'HOTSPOT 01 · ADVANCED PACKAGING');

  card(slide, 0.55, 1.2, 5.45, 3.25, { fill: C.white, shadow: true });
  addGrowthChart(slide, 0.82, 1.45, 4.9, 2.55, growthData.slice(0, 2), 'chiplet 與 hybrid bonding 的加速上升');

  card(slide, 6.2, 1.2, 3.2, 3.25, { fill: C.mist, shadow: true, accent: C.teal });
  slide.addText('Ecosystem stack', { x: 6.45, y: 1.42, w: 1.8, h: 0.18, fontFace: theme.headFont, fontSize: 17, bold: true, color: C.ink, margin: 0 });
  ['logic', 'memory', 'I/O', 'advanced packaging'].forEach((t, i) => pill(slide, 6.48, 1.84 + i * 0.5, 1.9, 0.28, t, [C.blue, C.rose, C.gold, C.teal][i], i === 2 ? C.navy : C.white));
  addBullets(slide, [
    'chiplet 高成長代表效能提升路徑由單晶片擴張為多裸晶異質整合。',
    'hybrid bonding 升溫，表示產業正往更細 pitch、更高頻寬、更低延遲的互連架構推進。',
    'Adeia 等互連/鍵合技術公司前移，顯示價值正從晶片設計轉向結合方式。'
  ], 6.42, 3.1, 2.65, 1.0, C.text, 10.5, 4);

  slide.addText('管理實務意涵', { x: 0.55, y: 4.68, w: 1.2, h: 0.18, fontFace: theme.bodyFont, fontSize: 10.5, bold: true, color: C.teal, margin: 0 });
  addBullets(slide, [
    '建立「封裝架構 PM / 系統封裝架構師」角色，而不是只靠製程或封裝工程師各自優化。',
    '投資評估要把 KGD、對位、rework、封裝內測試列為同級投資項。',
    '產品公司下一輪 roadmap 要從 die roadmap 升級為 package platform roadmap。'
  ], 1.62, 4.62, 7.6, 0.62, C.text, 10.7, 4);
  addFooter(slide, 3, false);
}

function slide4() {
  const slide = pptx.addSlide();
  slide.background = { color: C.mist };
  addTitle(slide, 'HBM 不只是記憶體議題，而是 AI 系統瓶頸的專利映射', 'HOTSPOT 02 · MEMORY × COMPUTE');

  card(slide, 0.55, 1.18, 4.9, 3.15, { fill: C.white, shadow: true });
  slide.addChart(pptx.ChartType.area, [{ name: 'HBM', labels: ['2021', '2025'], values: [10025, 27205] }], {
    x: 0.8, y: 1.45, w: 4.4, h: 2.45,
    showTitle: true, title: 'HBM 年度成長（2021 → 2025）', titleFontFace: theme.bodyFont, titleFontSize: 12, titleColor: C.ink,
    chartColors: [C.blue],
    valAxisLabelColor: C.muted, catAxisLabelColor: C.muted,
    valGridLine: { color: C.line, size: 0.6 }, catGridLine: { color: C.line, transparency: 100 },
    showLegend: false, showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.ink,
    chartArea: { fill: { color: C.white }, border: { color: C.white } }, showBorder: false,
  });

  card(slide, 5.7, 1.18, 3.7, 3.15, { fill: C.white, shadow: true, accent: C.rose });
  slide.addText('Compute–Memory bottleneck', { x: 5.98, y: 1.42, w: 2.7, h: 0.2, fontFace: theme.headFont, fontSize: 17, bold: true, color: C.ink, margin: 0 });
  pill(slide, 6.05, 1.88, 0.95, 0.3, 'Compute', C.navy);
  pill(slide, 7.05, 1.88, 0.9, 0.3, 'HBM', C.blue);
  pill(slide, 8.0, 1.88, 0.95, 0.3, 'Package', C.teal2);
  slide.addShape(pptx.ShapeType.chevron, { x: 6.84, y: 1.93, w: 0.16, h: 0.18, line: { color: C.gold, transparency: 100 }, fill: { color: C.gold } });
  slide.addShape(pptx.ShapeType.chevron, { x: 7.81, y: 1.93, w: 0.16, h: 0.18, line: { color: C.gold, transparency: 100 }, fill: { color: C.gold } });
  addBullets(slide, [
    'HBM 快速上升，與 AI 訓練/推論對高頻寬低延遲資料搬移需求一致。',
    '議題已不再只屬於 memory vendor，而是牽動 GPU、封裝、中介層、散熱、供電的系統性主題。',
    'Core CPC 中 G11C（記憶體）靠前，顯示記憶體重要性已提升到平台決勝層。'
  ], 6.0, 2.45, 2.9, 1.35, C.text, 10.4, 4);

  slide.addText('管理實務意涵', { x: 0.55, y: 4.58, w: 1.2, h: 0.18, fontFace: theme.bodyFont, fontSize: 10.5, bold: true, color: C.rose, margin: 0 });
  addBullets(slide, [
    '把 memory architecture 納入平台級 KPI，而不是只看 compute die 指標。',
    '建立 compute / memory / package 聯合評審機制，避免局部最佳化。',
    '提早鎖定 HBM、interposer、散熱相關合作與授權談判窗口。'
  ], 1.62, 4.52, 7.6, 0.68, C.text, 10.7, 4);
  addFooter(slide, 4, false);
}

function slide5() {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addTitle(slide, '即使封裝升溫，製程設備鏈仍是不可替代的底層戰場', 'HOTSPOT 03 · PROCESS / EQUIPMENT');

  card(slide, 0.55, 1.2, 4.32, 3.2, { fill: C.white, shadow: true });
  slide.addChart(pptx.ChartType.line, [{ name: 'EUV', labels: ['2021', '2025'], values: [23146, 30707] }], {
    x: 0.82, y: 1.44, w: 3.82, h: 2.18,
    showTitle: true, title: 'EUV：高基數下仍持續成長', titleFontFace: theme.bodyFont, titleFontSize: 12, titleColor: C.ink,
    chartColors: [C.rose], lineSize: 4,
    valAxisLabelColor: C.muted, catAxisLabelColor: C.muted,
    valGridLine: { color: C.line, size: 0.6 }, catGridLine: { color: C.line, transparency: 100 },
    showLegend: false, showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: C.ink,
    chartArea: { fill: { color: C.white }, border: { color: C.white } }, showBorder: false,
  });
  card(slide, 0.86, 3.78, 3.7, 0.42, { fill: 'FCEFF1', accent: C.rose });
  slide.addText('訊號重點：即使封裝成為熱點，先進曝光與製程設備專利量仍維持高檔，代表底層 fab 能力沒有被替代。', { x: 1.02, y: 3.9, w: 3.38, h: 0.16, fontFace: theme.bodyFont, fontSize: 9.7, color: C.text, margin: 0 });

  card(slide, 5.05, 1.2, 4.35, 1.62, { fill: C.mist, shadow: true, accent: C.gold });
  slide.addText('AI 晶片量產依賴鏈', { x: 5.3, y: 1.38, w: 2.3, h: 0.18, fontFace: theme.headFont, fontSize: 15.4, bold: true, color: C.ink, margin: 0, fit: 'shrink' });
  const chain = [
    ['Design', C.blue, 5.34],
    ['EUV / Etch', C.rose, 6.2],
    ['HBM / Package', C.teal2, 7.16],
    ['System Yield', C.gold, 8.18],
  ];
  chain.forEach(([label, color, x], i) => {
    card(slide, x, 1.86, i === 3 ? 0.92 : 0.8, 0.48, { fill: C.white, accent: color });
    slide.addText(label, { x: x + 0.08, y: 1.99, w: (i === 3 ? 0.76 : 0.64), h: 0.14, fontFace: theme.bodyFont, fontSize: 9.5, bold: true, color: C.ink, align: 'center', margin: 0 });
    if (i < chain.length - 1) {
      slide.addShape(pptx.ShapeType.chevron, { x: x + (i === 3 ? 0.98 : 0.85), y: 1.96, w: 0.18, h: 0.18, line: { color, transparency: 100 }, fill: { color } });
    }
  });

  card(slide, 5.05, 3.0, 2.08, 1.42, { fill: C.white, shadow: true, accent: C.gold });
  slide.addText('Fab CPC 底盤', { x: 5.28, y: 3.18, w: 1.35, h: 0.16, fontFace: theme.headFont, fontSize: 15.8, bold: true, color: C.ink, margin: 0 });
  ['C23C 沉積/鍍膜', 'G03F 光刻 / 圖案轉移', 'H01J 真空 / 電子束', 'B82Y 奈米製程整合'].forEach((t, i) => {
    slide.addText(`• ${t}`, { x: 5.3, y: 3.48 + i * 0.2, w: 1.6, h: 0.12, fontFace: theme.bodyFont, fontSize: 9.8, color: C.text, margin: 0 });
  });

  card(slide, 7.3, 3.0, 2.1, 1.42, { fill: 'FFF7E7', shadow: true, accent: C.rose });
  slide.addText('管理者雷達', { x: 7.55, y: 3.18, w: 1.0, h: 0.16, fontFace: theme.headFont, fontSize: 15.8, bold: true, color: C.ink, margin: 0 });
  ['節點可得性', '設備 / 材料風險', '共同開發 IP 邊界'].forEach((t, i) => {
    card(slide, 7.52, 3.46 + i * 0.27, 1.42, 0.2, { fill: C.white, accent: [C.gold, C.rose, C.teal][i] });
    slide.addText(t, { x: 7.63, y: 3.52 + i * 0.27, w: 1.12, h: 0.1, fontFace: theme.bodyFont, fontSize: 9.1, color: C.text, margin: 0, align: 'center' });
  });

  slide.addText('管理實務意涵', { x: 0.55, y: 4.62, w: 1.2, h: 0.18, fontFace: theme.bodyFont, fontSize: 10.5, bold: true, color: C.gold, margin: 0 });
  addBullets(slide, [
    '產品策略邊界其實由先進製程可得性決定，不能只偏向設計端思維。',
    '依賴先進代工者，需同步管理曝光、材料、設備與地緣風險。',
    '共同開發應把法務、IP、產品平台一起拉進治理機制。'
  ], 1.62, 4.56, 7.6, 0.7, C.text, 10.8, 3);
  addFooter(slide, 5, false);
}

function slide6() {
  const slide = pptx.addSlide();
  slide.background = { color: C.mist };
  slide.addText('WHITE SPACE & NEXT OPPORTUNITIES', {
    x: 0.55, y: 0.34, w: 4.4, h: 0.22,
    fontFace: theme.bodyFont, fontSize: 10.9, bold: true, color: C.teal,
    charSpacing: 1.2, margin: 0
  });
  slide.addText('白區在量產化輔助與跨層治理', {
    x: 0.55, y: 0.56, w: 7.9, h: 0.42,
    fontFace: theme.headFont, fontSize: 21.2, bold: true,
    color: C.ink, margin: 0, fit: 'shrink'
  });

  card(slide, 0.62, 1.2, 5.28, 3.06, { fill: C.white, shadow: true });
  slide.addText('機會矩陣：先抓高可行 × 高策略價值白區', { x: 0.9, y: 1.34, w: 4.52, h: 0.18, fontFace: theme.bodyFont, fontSize: 10.4, color: C.text, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.line, { x: 1.5, y: 3.42, w: 3.96, h: 0, line: { color: C.line, width: 1.1 } });
  slide.addShape(pptx.ShapeType.line, { x: 3.16, y: 1.86, w: 0, h: 1.56, line: { color: C.line, width: 1.1 } });
  slide.addText('可行性 →', { x: 4.5, y: 3.66, w: 0.82, h: 0.13, fontFace: theme.bodyFont, fontSize: 10.0, color: C.muted, margin: 0 });
  slide.addText('策略價值 ↑', { x: 0.92, y: 1.92, w: 0.76, h: 0.13, fontFace: theme.bodyFont, fontSize: 10.0, color: C.muted, margin: 0, rotate: 270 });
  slide.addText('觀察 / 佈局', { x: 1.18, y: 1.9, w: 0.98, h: 0.13, fontFace: theme.bodyFont, fontSize: 10.0, bold: true, color: C.muted, margin: 0 });
  slide.addText('優先投資', { x: 4.0, y: 1.9, w: 0.92, h: 0.13, fontFace: theme.bodyFont, fontSize: 10.0, bold: true, color: C.teal, margin: 0 });
  slide.addText('戰術切入', { x: 1.18, y: 3.14, w: 0.92, h: 0.13, fontFace: theme.bodyFont, fontSize: 10.0, bold: true, color: C.gold, margin: 0 });
  slide.addText('延後 / 追蹤', { x: 3.98, y: 3.14, w: 1.0, h: 0.13, fontFace: theme.bodyFont, fontSize: 10.0, bold: true, color: C.muted, margin: 0 });

  const wsCards = [
    ['Yield-learning\nrework / KGD', 3.9, 1.94, '8CA7BF', 1.3],
    ['Chiplet\n系統協同', 3.18, 2.12, C.teal, 1.22],
    ['Hybrid bonding\n量產化', 4.02, 2.54, C.blue, 1.22],
    ['封裝材料 /\n中介層', 1.68, 2.28, C.rose, 1.16],
    ['Backside power\nDFT / reliability', 3.28, 2.9, C.gold, 1.48],
    ['SiC / GaN\n模組平台', 1.88, 2.92, C.green, 1.14],
  ];
  wsCards.forEach(([label, x, y, accent, w]) => {
    card(slide, x, y, w, 0.48, { fill: C.white, accent, shadow: false });
    slide.addText(label, { x: x + 0.08, y: y + 0.07, w: w - 0.16, h: 0.26, fontFace: theme.bodyFont, fontSize: 9.1, bold: true, color: C.ink, align: 'center', margin: 0, fit: 'shrink' });
  });

  card(slide, 6.06, 1.2, 3.32, 3.06, { fill: C.white, shadow: true, accent: C.teal });
  slide.addText('佈局節奏', { x: 6.34, y: 1.34, w: 1.08, h: 0.2, fontFace: theme.headFont, fontSize: 16.6, bold: true, color: C.ink, margin: 0 });
  const lanes = [
    ['近程 0–12M', '先做 yield-learning、rework / KGD、封裝材料。', C.gold],
    ['中程 12–24M', '推 hybrid bonding 量產化與 backside power reliability。', C.blue],
    ['長程 24M+', '布局 chiplet 系統協同與 SiC/GaN 模組平台。', C.teal],
  ];
  lanes.forEach(([title, body, accent], i) => {
    const y = 1.7 + i * 0.72;
    card(slide, 6.24, y, 2.9, 0.58, { fill: C.mist, accent });
    slide.addText(title, { x: 6.42, y: y + 0.08, w: 0.92, h: 0.13, fontFace: theme.bodyFont, fontSize: 9.8, bold: true, color: accent, margin: 0, fit: 'shrink' });
    slide.addText(body, { x: 7.18, y: y + 0.07, w: 1.7, h: 0.26, fontFace: theme.bodyFont, fontSize: 8.8, color: C.text, margin: 0, fit: 'shrink' });
  });
  card(slide, 6.24, 3.86, 2.9, 0.3, { fill: 'EAF7F7', accent: C.teal });
  slide.addText('原則：先做能快落地的輔助技術。', { x: 6.42, y: 3.94, w: 2.44, h: 0.12, fontFace: theme.bodyFont, fontSize: 8.9, color: C.text, margin: 0, fit: 'shrink' });

  card(slide, 0.72, 4.4, 8.38, 0.3, { fill: 'EAF0F5', accent: C.teal });
  slide.addText('管理意涵：投資組合可預留 20–30% 給量產化輔助；IP 團隊優先掃 reliability、DFT、rework、yield-learning 空隙。', { x: 0.9, y: 4.48, w: 8.02, h: 0.12, fontFace: theme.bodyFont, fontSize: 9.1, color: C.text, margin: 0, fit: 'shrink' });
  addFooter(slide, 6, false);
}

function slide7() {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addTitle(slide, '從專利情報到技術治理：四個管理動作', 'MANAGEMENT OPERATING SYSTEM');

  card(slide, 0.68, 1.34, 8.7, 2.0, { fill: C.mist, shadow: true });
  slide.addText('Operating loop：把專利雷達直接接到 roadmap、投資組合與 NPI gate。', {
    x: 0.94, y: 1.48, w: 7.86, h: 0.18,
    fontFace: theme.bodyFont, fontSize: 10.9, color: C.text, margin: 0, fit: 'shrink'
  });
  const flow = [
    ['01', 'Sensing', '五大主題雷達', 0.94, C.teal],
    ['02', 'Prioritize', '產品 / 設備 / 材料 map', 3.02, C.blue],
    ['03', 'Invest', '核心創新 + 量產化', 5.1, C.gold],
    ['04', 'Govern', 'NPI gate + FTO', 7.18, C.rose],
  ];
  flow.forEach(([num, title, body, x, accent], i) => {
    card(slide, x, 1.88, 1.62, 0.96, { fill: C.white, accent, shadow: false });
    slide.addText(num, { x: x + 0.1, y: 2.02, w: 0.34, h: 0.16, fontFace: theme.headFont, fontSize: 15.1, bold: true, color: accent, margin: 0, fit: 'shrink' });
    slide.addText(title, { x: x + 0.44, y: 2.0, w: 0.92, h: 0.16, fontFace: theme.headFont, fontSize: 13.0, bold: true, color: C.ink, margin: 0, fit: 'shrink' });
    slide.addText(body, { x: x + 0.12, y: 2.32, w: 1.32, h: 0.2, fontFace: theme.bodyFont, fontSize: 9.3, color: C.text, margin: 0, align: 'center', fit: 'shrink' });
    if (i < flow.length - 1) {
      slide.addShape(pptx.ShapeType.chevron, { x: x + 1.68, y: 2.25, w: 0.2, h: 0.22, line: { color: accent, transparency: 100 }, fill: { color: accent } });
    }
  });

  const cadence = [
    ['季度', '更新熱點與競爭者雷達。', C.teal],
    ['立項前', '檢查 white-space、合作與授權機會。', C.gold],
    ['NPI gate', '納入 FTO、共研 IP、可靠性檢查。', C.blue],
  ];
  cadence.forEach(([title, body, accent], i) => {
    const x = 0.84 + i * 2.86;
    card(slide, x, 3.5, 2.5, 0.82, { fill: C.white, shadow: true, accent });
    slide.addText(title, { x: x + 0.16, y: 3.63, w: 0.9, h: 0.14, fontFace: theme.bodyFont, fontSize: 9.9, bold: true, color: accent, margin: 0, fit: 'shrink' });
    slide.addText(body, { x: x + 0.16, y: 3.86, w: 2.1, h: 0.18, fontFace: theme.bodyFont, fontSize: 9.1, color: C.text, margin: 0, fit: 'shrink' });
  });

  card(slide, 0.72, 4.42, 8.36, 0.34, { fill: 'EEF3F8', accent: C.blue });
  slide.addText('結論：成熟度關鍵不再只是研發效率，而是跨層決策速度。', { x: 0.94, y: 4.51, w: 7.9, h: 0.12, fontFace: theme.bodyFont, fontSize: 10.1, bold: true, color: C.ink, margin: 0, fit: 'shrink' });
  addFooter(slide, 7, false);
}

function slide8() {
  const slide = pptx.addSlide();
  slide.background = { color: C.navy };
  slide.addShape(pptx.ShapeType.rect, { x: 0.45, y: 0.5, w: 9.1, h: 4.35, line: { color: C.teal, transparency: 100 }, fill: { color: '112B44' } });
  slide.addText('CLOSING DECISIONS', { x: 0.8, y: 0.85, w: 2.0, h: 0.18, fontFace: theme.bodyFont, fontSize: 10.5, bold: true, color: C.gold, charSpacing: 1.4, margin: 0 });
  slide.addText('給管理層的三個決策：擴整合、補量產、做平台', { x: 0.8, y: 1.08, w: 6.1, h: 0.36, fontFace: theme.headFont, fontSize: 21.0, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  slide.addText('把專利情報轉成執行板：每個決策都要有 owner、近期行動與 12–24 個月 KPI。', { x: 0.82, y: 1.48, w: 5.8, h: 0.18, fontFace: theme.bodyFont, fontSize: 10.0, color: 'D6E2EE', margin: 0, fit: 'shrink' });

  const decisions = [
    ['擴整合', '重做跨層 roadmap', 'Owner：CTO / platform', 'KPI：聯合審查率', C.teal, 0.82],
    ['補量產', '加碼量產化輔助', 'Owner：R&D / mfg', 'KPI：yield / reliability 立項', C.gold, 2.92],
    ['做平台', '平台合作與 IP 治理', 'Owner：BU / IP / legal', 'KPI：共研 / 授權 / 模組數', C.rose, 5.02],
  ];
  decisions.forEach(([title, act, owner, kpi, accent, x]) => {
    card(slide, x, 2.0, 1.9, 1.7, { fill: C.white, shadow: true, accent });
    slide.addText(title, { x: x + 0.16, y: 2.14, w: 0.98, h: 0.18, fontFace: theme.headFont, fontSize: 17.4, bold: true, color: accent, margin: 0 });
    slide.addText(act, { x: x + 0.16, y: 2.42, w: 1.54, h: 0.16, fontFace: theme.bodyFont, fontSize: 10.7, bold: true, color: C.ink, margin: 0, fit: 'shrink' });
    slide.addText(owner, { x: x + 0.16, y: 2.76, w: 1.58, h: 0.14, fontFace: theme.bodyFont, fontSize: 8.9, color: C.text, margin: 0, fit: 'shrink' });
    card(slide, x + 0.14, 3.08, 1.62, 0.38, { fill: C.mist, accent });
    slide.addText(kpi, { x: x + 0.2, y: 3.2, w: 1.5, h: 0.12, fontFace: theme.bodyFont, fontSize: 8.6, color: C.text, margin: 0, align: 'center', fit: 'shrink' });
  });

  const blocks = [
    ['短期 0–12M', '建五大主題雷達與競爭者 map', C.teal],
    ['中期 12–24M', '把 white-space 轉成研發立項池', C.gold],
    ['長期 24M+', '升級成 system / package / process 共治', C.rose],
  ];
  blocks.forEach((b, i) => {
    const x = 7.14;
    const y = 1.8 + i * 0.88;
    card(slide, x, y, 1.86, 0.8, { fill: 'EAF0F5', shadow: false, accent: b[2] });
    slide.addText(b[0], { x: x + 0.12, y: y + 0.1, w: 1.3, h: 0.12, fontFace: theme.bodyFont, fontSize: 8.9, bold: true, color: b[2], margin: 0, fit: 'shrink' });
    slide.addText(b[1], { x: x + 0.12, y: y + 0.24, w: 1.52, h: 0.24, fontFace: theme.bodyFont, fontSize: 8.1, color: C.text, margin: 0, fit: 'shrink' });
  });

  card(slide, 0.86, 3.9, 6.28, 0.38, { fill: '163550', accent: C.teal });
  slide.addText('一句話結論：下一輪競爭比的不是單點更快，而是跨層整合、量產落地與平台治理。', { x: 1.0, y: 4.01, w: 5.98, h: 0.12, fontFace: theme.bodyFont, fontSize: 9.2, bold: true, color: C.white, margin: 0, align: 'center', fit: 'shrink' });

  addSourceNote(slide, 0.8, 4.44, 7.1, '本簡報著重相對排序與熱點辨識；若要升級董事會版，建議再補多資料源交叉驗證。', true);
  addFooter(slide, 8, true);
}

slide1();
slide2();
slide3();
slide4();
slide5();
slide6();
slide7();
slide8();

function htmlPreview() {
  const cards = growthData.map(d => `
    <div class="mini-card">
      <div class="mini-label">${d.label}</div>
      <div class="mini-value" style="color:#${d.color}">${mul(d.start,d.end)}</div>
      <div class="mini-range">${fmt(d.start)} → ${fmt(d.end)}</div>
    </div>`).join('');
  const whiteCards = whitespace.map((w, i) => `<div class="ws-card"><h4>${w}</h4><p>${i < 3 ? '成熟度中等｜競爭升溫中' : '成熟度中低｜可搶先布局'}</p></div>`).join('');
  return `<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8" />
<title>${theme.title} - Preview</title>
<style>
  :root { --navy:#0B1F33; --navy2:#132B45; --teal:#0F8B8D; --teal2:#16A6A8; --gold:#E5B85C; --mist:#F3F7FA; --text:#233143; --muted:#6E7D90; --blue:#3C78D8; --rose:#D96C75; }
  *{ box-sizing:border-box; }
  body{ margin:0; font-family:-apple-system,BlinkMacSystemFont,"PingFang TC","Microsoft JhengHei",sans-serif; background:#dfe7ee; color:var(--text); }
  .wrap{ padding:24px; display:grid; gap:24px; justify-content:center; }
  .slide{ width:1280px; min-height:720px; background:white; box-shadow:0 12px 32px rgba(11,31,51,.12); position:relative; overflow:hidden; border-radius:10px; }
  .dark{ background:var(--navy); color:white; }
  .title{ position:absolute; left:70px; top:72px; font-size:42px; font-weight:800; width:900px; line-height:1.15; }
  .kicker{ position:absolute; left:70px; top:38px; font-size:16px; font-weight:700; letter-spacing:.12em; color:var(--teal); }
  .dark .kicker{ color:var(--gold); }
  .subtitle{ position:absolute; left:70px; top:210px; width:650px; font-size:18px; line-height:1.45; color:#dce8f4; }
  .metrics{ position:absolute; right:68px; top:72px; width:380px; display:grid; grid-template-columns:repeat(2,1fr); gap:16px; }
  .mini-card{ background:white; color:var(--text); border-left:8px solid var(--teal); border-radius:10px; padding:16px 16px 14px; min-height:116px; box-shadow:0 10px 24px rgba(11,31,51,.08); }
  .mini-card:nth-child(2){ border-left-color:#74D3D4; }
  .mini-card:nth-child(3){ border-left-color:#E5B85C; }
  .mini-card:nth-child(4){ border-left-color:#3C78D8; }
  .mini-card:nth-child(5){ border-left-color:#D96C75; grid-column:span 2; }
  .mini-label{ font-size:15px; color:var(--muted); font-weight:700; }
  .mini-value{ font-size:34px; font-weight:800; margin-top:8px; }
  .mini-range{ font-size:15px; margin-top:6px; }
  .strip{ position:absolute; left:0; right:0; bottom:0; height:108px; background:linear-gradient(90deg, rgba(15,139,141,.20), rgba(15,139,141,.50)); }
  .summary-list{ position:absolute; left:75px; top:295px; width:620px; font-size:25px; line-height:1.5; }
  .summary-list li{ margin:12px 0; }
  .section-grid{ position:absolute; left:70px; right:70px; top:140px; bottom:88px; display:grid; grid-template-columns:1fr 1fr; gap:24px; }
  .panel{ background:white; border-radius:14px; box-shadow:0 10px 24px rgba(11,31,51,.08); padding:24px 26px; }
  .panel h3{ margin:0 0 12px; font-size:28px; }
  .panel ul{ margin:10px 0 0 18px; padding:0; font-size:20px; line-height:1.45; }
  .center-badge{ display:flex; align-items:center; justify-content:center; background:var(--navy); color:#fff; border-radius:16px; font-size:44px; font-weight:800; }
  .two-col{ position:absolute; left:70px; right:70px; top:140px; bottom:88px; display:grid; grid-template-columns:1.2fr .8fr; gap:24px; }
  .chart-box,.stack-box,.info-box{ background:white; border-radius:14px; box-shadow:0 10px 24px rgba(11,31,51,.08); padding:20px 24px; }
  .chart-box h3,.stack-box h3,.info-box h3{ margin:0 0 14px; font-size:28px; }
  .line-chart{ width:100%; height:360px; }
  .stack-pill{ display:inline-block; padding:10px 18px; color:white; border-radius:999px; font-weight:700; margin:8px 10px 8px 0; }
  .stack-box ul{ margin:14px 0 0 18px; font-size:18px; line-height:1.45; }
  .ws-grid{ position:absolute; left:70px; right:70px; top:140px; display:grid; grid-template-columns:repeat(3,1fr); gap:22px; }
  .ws-card{ background:white; border-radius:14px; min-height:164px; padding:20px 22px; box-shadow:0 10px 24px rgba(11,31,51,.08); }
  .ws-card h4{ margin:0 0 10px; font-size:24px; line-height:1.22; }
  .ws-card p{ margin:0; color:var(--muted); font-size:19px; }
  .triple-grid{ position:absolute; left:70px; right:70px; top:150px; display:grid; grid-template-columns:repeat(3,1fr); gap:20px; }
  .mini-panel{ background:white; border-radius:14px; box-shadow:0 10px 24px rgba(11,31,51,.08); padding:20px 22px; min-height:200px; }
  .mini-panel h4{ margin:0 0 10px; font-size:24px; }
  .mini-panel p,.mini-panel li{ font-size:18px; line-height:1.42; }
  .four-grid{ position:absolute; left:70px; right:70px; top:150px; display:grid; grid-template-columns:repeat(2,1fr); gap:22px; }
  .action-card{ background:var(--mist); border-radius:14px; box-shadow:0 10px 24px rgba(11,31,51,.08); padding:20px 22px; min-height:158px; border-left:8px solid var(--teal); }
  .action-card:nth-child(2){ border-left-color:var(--blue); }
  .action-card:nth-child(3){ border-left-color:var(--gold); }
  .action-card:nth-child(4){ border-left-color:var(--rose); }
  .action-card strong{ display:block; font-size:24px; margin-bottom:8px; }
  .action-card p{ font-size:18px; line-height:1.42; margin:0; }
  .timeline{ position:absolute; right:82px; top:160px; width:340px; display:grid; gap:18px; }
  .timeline .mini-panel{ min-height:120px; }
  .relation-map{ position:absolute; left:70px; right:70px; top:150px; bottom:96px; background:white; border-radius:16px; box-shadow:0 10px 24px rgba(11,31,51,.08); padding:26px; }
  .node-grid{ display:grid; grid-template-columns:1fr 220px 1fr; gap:18px; align-items:center; margin-top:22px; }
  .node-stack{ display:grid; grid-template-columns:repeat(2,1fr); gap:12px; }
  .node{ background:var(--mist); border-radius:12px; padding:12px 14px; text-align:center; font-weight:700; font-size:18px; border-left:6px solid var(--teal); }
  .node.gold{ border-left-color:var(--gold); }
  .node.rose{ border-left-color:var(--rose); }
  .bridge-band{ display:grid; grid-template-columns:repeat(3,1fr); gap:14px; margin-top:20px; }
  .bridge-note{ background:#EEF3F8; border-radius:12px; padding:12px 14px; font-size:16px; line-height:1.35; }
  .quad-wrap{ position:absolute; left:70px; right:70px; top:150px; bottom:96px; display:grid; grid-template-columns:1.15fr .85fr; gap:24px; }
  .quad-grid{ background:white; border-radius:16px; box-shadow:0 10px 24px rgba(11,31,51,.08); padding:20px 24px; position:relative; }
  .quad-grid .axis-h,.quad-grid .axis-v{ position:absolute; background:#D8E1EA; }
  .quad-grid .axis-h{ left:76px; right:36px; top:238px; height:2px; }
  .quad-grid .axis-v{ left:270px; top:62px; bottom:46px; width:2px; }
  .quad-card{ position:absolute; width:170px; min-height:66px; background:white; border-radius:12px; box-shadow:0 8px 18px rgba(11,31,51,.08); padding:10px 12px; font-size:15px; font-weight:700; text-align:center; border-left:6px solid var(--teal); }
  .lane-list{ background:white; border-radius:16px; box-shadow:0 10px 24px rgba(11,31,51,.08); padding:20px 24px; }
  .lane{ background:var(--mist); border-radius:12px; padding:14px 16px; margin:14px 0; font-size:16px; line-height:1.35; }
  .lane strong{ display:block; margin-bottom:6px; }
  .flow-row{ position:absolute; left:70px; right:70px; top:160px; bottom:96px; display:grid; grid-template-rows:1.2fr .9fr; gap:22px; }
  .flow-box,.cadence-row{ background:var(--mist); border-radius:16px; box-shadow:0 10px 24px rgba(11,31,51,.08); padding:24px; }
  .step-track{ display:grid; grid-template-columns:repeat(4,1fr); gap:18px; align-items:center; margin-top:22px; }
  .flow-step{ background:white; border-radius:14px; padding:16px; box-shadow:0 8px 18px rgba(11,31,51,.08); font-size:16px; line-height:1.35; border-left:7px solid var(--teal); }
  .cadence-grid{ display:grid; grid-template-columns:repeat(3,1fr); gap:18px; }
  .decision-grid{ position:absolute; left:70px; right:70px; top:170px; bottom:96px; display:grid; grid-template-columns:2.4fr .95fr; gap:22px; }
  .decision-cards{ display:grid; grid-template-columns:repeat(3,1fr); gap:18px; }
  .decision-card{ background:white; color:var(--text); border-radius:14px; padding:18px 18px 16px; box-shadow:0 10px 24px rgba(0,0,0,.18); border-top:8px solid var(--teal); }
  .decision-card.gold{ border-top-color:var(--gold); }
  .decision-card.rose{ border-top-color:var(--rose); }
  .decision-card h4{ margin:0 0 10px; font-size:26px; }
  .decision-card p{ margin:0 0 10px; font-size:17px; line-height:1.35; }
  .kpi-pill{ background:var(--mist); border-radius:999px; padding:10px 12px; font-size:14px; line-height:1.25; }
  .footer-note{ position:absolute; left:70px; right:70px; bottom:24px; font-size:15px; color:var(--muted); }
</style>
</head>
<body>
<div class="wrap">
  <section class="slide dark">
    <div class="kicker">EXECUTIVE VIEW</div>
    <div class="title">${theme.title}</div>
    <div class="subtitle">技術競爭主軸已從單點製程，轉向「封裝 × 系統 × 製造整合」</div>
    <ul class="summary-list">
      <li>最穩定領先群：TSMC、Intel、Samsung</li>
      <li>第二梯隊依主題分化：Micron、ASM、Lam Research、Applied Materials、Adeia、Amkor、ASML / Carl Zeiss SMT</li>
      <li>高成長主題：chiplet、hybrid bonding、backside power、HBM、EUV</li>
    </ul>
    <div class="metrics">${cards}</div>
    <div class="strip"></div>
  </section>
  <section class="slide" style="background:var(--mist)">
    <div class="kicker">COMPETITIVE LANDSCAPE</div>
    <div class="title" style="color:var(--text);font-size:38px;top:70px;width:1140px;">產業鏈雙核心：器件/封裝/IC 與 製程/設備 同步升級</div>
    <div class="relation-map">
      <div style="font-size:18px;line-height:1.4;">用關係圖而不是名單看競爭：需求牽引在左、製造供給在右，TSMC 位於少數可跨鏈整合的樞紐。</div>
      <div class="node-grid">
        <div>
          <div style="font-size:18px;font-weight:800;color:var(--teal);margin-bottom:10px;">需求 / 系統側</div>
          <div class="node-stack">
            <div class="node">Intel</div><div class="node">Apple</div>
            <div class="node">Samsung</div><div class="node">Adeia</div>
            <div class="node rose">Micron</div><div class="node">Amkor</div>
          </div>
        </div>
        <div class="center-badge" style="height:200px; flex-direction:column; gap:8px;">TSMC<div style="font-size:18px;font-weight:600;color:#DCE7F2;">core × fab hub</div></div>
        <div>
          <div style="font-size:18px;font-weight:800;color:var(--gold);margin-bottom:10px;">製造 / 設備側</div>
          <div class="node-stack">
            <div class="node gold">ASM</div><div class="node rose">ASML / Carl Zeiss</div>
            <div class="node gold">Lam Research</div><div class="node rose">SanDisk / WD</div>
            <div class="node gold" style="grid-column:span 2;">Applied Materials</div>
          </div>
        </div>
      </div>
      <div class="bridge-band">
        <div class="bridge-note">需求拉動：AI 平台、HBM、封裝協同把產品 roadmap 推向系統整合。</div>
        <div class="bridge-note">整合樞紐：少數玩家能同時掌握客戶需求、節點節奏與封裝協同。</div>
        <div class="bridge-note">供給主導：設備 / 材料商正鎖定未來節點 know-how 與議價權。</div>
      </div>
    </div>
    <div class="footer-note">管理重點：競爭者地圖至少拆成產品 / 封裝 / 設備 / 材料四層，不要只看總 patent count。</div>
  </section>
  <section class="slide">
    <div class="kicker">HOTSPOT 01 · ADVANCED PACKAGING</div>
    <div class="title" style="color:var(--text);font-size:38px;top:70px;width:1140px;">封裝已從成本中心變成架構創新中心</div>
    <div class="two-col">
      <div class="chart-box">
        <h3>chiplet 與 hybrid bonding 的加速上升</h3>
        <svg class="line-chart" viewBox="0 0 640 360">
          <rect x="0" y="0" width="640" height="360" fill="#fff"/>
          <line x1="70" y1="290" x2="560" y2="290" stroke="#D8E1EA" stroke-width="2"/>
          <line x1="70" y1="60" x2="70" y2="290" stroke="#D8E1EA" stroke-width="2"/>
          <text x="70" y="320" font-size="18" fill="#6E7D90">2021</text>
          <text x="520" y="320" font-size="18" fill="#6E7D90">2025</text>
          <polyline points="100,250 520,90" fill="none" stroke="#16A6A8" stroke-width="5"/>
          <polyline points="100,210 520,130" fill="none" stroke="#74D3D4" stroke-width="5"/>
          <circle cx="100" cy="250" r="6" fill="#16A6A8"/><circle cx="520" cy="90" r="6" fill="#16A6A8"/>
          <circle cx="100" cy="210" r="6" fill="#74D3D4"/><circle cx="520" cy="130" r="6" fill="#74D3D4"/>
          <text x="520" y="82" font-size="16" fill="#16A6A8">8,964</text>
          <text x="520" y="122" font-size="16" fill="#74D3D4">8,839</text>
        </svg>
      </div>
      <div class="stack-box">
        <h3>Ecosystem stack</h3>
        <div class="stack-pill" style="background:#3C78D8">logic</div>
        <div class="stack-pill" style="background:#D96C75">memory</div>
        <div class="stack-pill" style="background:#E5B85C;color:#122033;">I/O</div>
        <div class="stack-pill" style="background:#16A6A8">advanced packaging</div>
        <ul>
          <li>多裸晶異質整合成為效能主路徑</li>
          <li>互連與鍵合方式成為價值核心</li>
          <li>roadmap 要升級為 package platform roadmap</li>
        </ul>
      </div>
    </div>
  </section>
  <section class="slide" style="background:var(--mist)">
    <div class="kicker">HOTSPOT 02 · MEMORY × COMPUTE</div>
    <div class="title" style="color:var(--text);font-size:38px;top:70px;width:1140px;">HBM 不只是記憶體議題，而是 AI 系統瓶頸的專利映射</div>
    <div class="two-col">
      <div class="chart-box">
        <h3>HBM 年度成長（2021 → 2025）</h3>
        <svg class="line-chart" viewBox="0 0 640 360">
          <rect x="0" y="0" width="640" height="360" fill="#fff"/>
          <line x1="70" y1="290" x2="560" y2="290" stroke="#D8E1EA" stroke-width="2"/>
          <line x1="70" y1="70" x2="70" y2="290" stroke="#D8E1EA" stroke-width="2"/>
          <polygon points="100,235 520,120 520,290 100,290" fill="#3C78D8" opacity="0.25"/>
          <polyline points="100,235 520,120" fill="none" stroke="#3C78D8" stroke-width="6"/>
          <text x="85" y="228" font-size="16" fill="#3C78D8">10,025</text>
          <text x="520" y="112" font-size="16" fill="#3C78D8">27,205</text>
        </svg>
      </div>
      <div class="stack-box">
        <h3>Compute–Memory bottleneck</h3>
        <div class="stack-pill" style="background:#0B1F33">Compute</div>
        <div class="stack-pill" style="background:#3C78D8">HBM</div>
        <div class="stack-pill" style="background:#16A6A8">Package</div>
        <ul>
          <li>AI 平台已從 compute 單點優化轉向系統協同</li>
          <li>HBM 牽動封裝、中介層、散熱與供電</li>
          <li>memory architecture 應進入平台級 KPI</li>
        </ul>
      </div>
    </div>
    <div class="footer-note">管理重點：建立 compute / memory / package 聯合評審機制，避免局部最佳化。</div>
  </section>
  <section class="slide">
    <div class="kicker">HOTSPOT 03 · PROCESS / EQUIPMENT</div>
    <div class="title" style="color:var(--text);font-size:38px;top:70px;width:1140px;">即使封裝升溫，製程設備鏈仍是不可替代的底層戰場</div>
    <div class="triple-grid" style="grid-template-columns:1.15fr 1fr 1fr;">
      <div class="mini-panel">
        <h4>EUV：高基數下仍持續成長</h4>
        <p>2021：23,146</p>
        <p>2025：30,707</p>
        <p>訊號重點：先進曝光與製程設備專利量仍維持高檔，代表底層 fab 能力沒有被替代。</p>
      </div>
      <div class="mini-panel">
        <h4>AI 晶片量產依賴鏈</h4>
        <p><span class="stack-pill" style="background:#3C78D8">Design</span><span class="stack-pill" style="background:#D96C75">EUV / Etch</span><span class="stack-pill" style="background:#16A6A8">HBM / Package</span><span class="stack-pill" style="background:#E5B85C;color:#122033;">System Yield</span></p>
        <p>設計、製程、封裝與良率是一條鏈，而非各自獨立優化。</p>
      </div>
      <div class="mini-panel">
        <h4>Fab CPC 底盤 + 管理雷達</h4>
        <ul>
          <li>C23C 沉積 / 鍍膜</li>
          <li>G03F 光刻 / 圖案轉移</li>
          <li>設備 / 材料風險</li>
          <li>共同開發 IP 邊界</li>
        </ul>
      </div>
    </div>
    <div class="footer-note">管理重點：產品策略邊界其實由先進製程可得性決定，不能只偏向設計端思維。</div>
  </section>
  <section class="slide" style="background:var(--mist)">
    <div class="kicker">WHITE SPACE & NEXT OPPORTUNITIES</div>
    <div class="title" style="color:var(--text);font-size:38px;top:70px;width:1140px;">真正的白區在「量產化輔助技術」與「跨層整合治理」</div>
    <div class="quad-wrap">
      <div class="quad-grid">
        <div style="font-size:24px;font-weight:800;">機會矩陣：先抓高可行 × 高策略價值白區</div>
        <div class="axis-h"></div><div class="axis-v"></div>
        <div style="position:absolute;left:80px;top:70px;font-size:15px;color:var(--muted);">觀察 / 佈局</div>
        <div style="position:absolute;left:382px;top:70px;font-size:15px;color:var(--teal);font-weight:800;">優先投資</div>
        <div style="position:absolute;left:80px;top:256px;font-size:15px;color:var(--gold);font-weight:800;">戰術切入</div>
        <div style="position:absolute;left:382px;top:256px;font-size:15px;color:var(--muted);">延後 / 追蹤</div>
        <div class="quad-card" style="left:392px;top:86px;width:164px;border-left-color:#8CA7BF;">Yield-learning / rework / KGD</div>
        <div class="quad-card" style="left:308px;top:110px;width:150px;border-left-color:var(--teal);">Chiplet 系統協同</div>
        <div class="quad-card" style="left:404px;top:166px;width:152px;border-left-color:var(--blue);">Hybrid bonding 量產化</div>
        <div class="quad-card" style="left:132px;top:142px;width:150px;border-left-color:var(--rose);">封裝材料 / 中介層</div>
        <div class="quad-card" style="left:306px;top:258px;width:168px;border-left-color:var(--gold);">Backside power DFT / reliability</div>
        <div class="quad-card" style="left:146px;top:250px;width:146px;border-left-color:#4C9F70;">SiC / GaN 模組平台</div>
      </div>
      <div class="lane-list">
        <div style="font-size:24px;font-weight:800;">佈局節奏</div>
        <div class="lane"><strong style="color:var(--gold)">近程 0–12M</strong>Yield-learning / rework / KGD、封裝材料</div>
        <div class="lane"><strong style="color:var(--blue)">中程 12–24M</strong>Hybrid bonding 量產化、backside power reliability</div>
        <div class="lane"><strong style="color:var(--teal)">長程 24M+</strong>Chiplet 系統協同、SiC/GaN 模組平台</div>
        <div class="bridge-note" style="margin-top:18px;">原則：先做能快落地的輔助技術。</div>
      </div>
    </div>
    <div class="footer-note">管理意涵：投資組合可預留 20–30% 給量產化輔助；IP 團隊優先掃 reliability、DFT、rework、yield-learning 空隙。</div>
  </section>
  <section class="slide">
    <div class="kicker">MANAGEMENT OPERATING SYSTEM</div>
    <div class="title" style="color:var(--text);font-size:38px;top:70px;width:1140px;">從專利情報到技術治理：四個管理動作</div>
    <div class="flow-row">
      <div class="flow-box">
        <div style="font-size:19px;line-height:1.35;">Operating loop：把專利情報直接接到 roadmap、投資組合與 NPI gate。</div>
        <div class="step-track">
          <div class="flow-step" style="border-left-color:var(--teal)"><strong>01 Sensing</strong><br/>五大主題雷達</div>
          <div class="flow-step" style="border-left-color:var(--blue)"><strong>02 Prioritize</strong><br/>產品 / 設備 / 材料 map</div>
          <div class="flow-step" style="border-left-color:var(--gold)"><strong>03 Invest</strong><br/>核心創新 + 量產化</div>
          <div class="flow-step" style="border-left-color:var(--rose)"><strong>04 Govern</strong><br/>NPI gate + FTO</div>
        </div>
      </div>
      <div class="cadence-row">
        <div class="cadence-grid">
          <div class="flow-step" style="border-left-color:var(--teal)"><strong>季度</strong><br/>更新熱點與競爭者雷達</div>
          <div class="flow-step" style="border-left-color:var(--gold)"><strong>立項前</strong><br/>檢查 white-space、合作與授權機會</div>
          <div class="flow-step" style="border-left-color:var(--blue)"><strong>NPI gate</strong><br/>納入 FTO、共研 IP、可靠性檢查</div>
        </div>
      </div>
    </div>
    <div class="footer-note">結論：成熟度關鍵不再只是研發效率，而是跨層決策速度。</div>
  </section>
  <section class="slide dark">
    <div class="kicker">CLOSING DECISIONS</div>
    <div class="title" style="width:650px;">給管理層的三個決策：擴整合、補量產、做平台</div>
    <div style="position:absolute;left:70px;top:150px;width:640px;font-size:20px;line-height:1.4;color:#D6E2EE;">把專利情報轉成執行板：每個決策都要有 owner、近期行動與 12–24 個月 KPI。</div>
    <div class="decision-grid">
      <div class="decision-cards">
        <div class="decision-card"><h4>擴整合</h4><p>重做跨層 roadmap</p><p style="font-size:15px;color:#5C6C7C;">Owner：CTO / platform</p><div class="kpi-pill">KPI：聯合審查率</div></div>
        <div class="decision-card gold"><h4>補量產</h4><p>加碼量產化輔助</p><p style="font-size:15px;color:#5C6C7C;">Owner：R&D / mfg</p><div class="kpi-pill">KPI：yield / reliability 立項</div></div>
        <div class="decision-card rose"><h4>做平台</h4><p>平台合作與 IP 治理</p><p style="font-size:15px;color:#5C6C7C;">Owner：BU / IP / legal</p><div class="kpi-pill">KPI：共研 / 授權 / 模組數</div></div>
      </div>
      <div class="timeline" style="position:static;width:auto;">
        <div class="mini-panel"><h4>短期 0–12M</h4><p>建五大主題雷達與競爭者 map</p></div>
        <div class="mini-panel"><h4>中期 12–24M</h4><p>把 white-space 轉成研發立項池</p></div>
        <div class="mini-panel"><h4>長期 24M+</h4><p>升級成 system / package / process 共治</p></div>
      </div>
    </div>
    <div style="position:absolute;left:70px;top:545px;width:820px;background:#163550;border-radius:14px;padding:14px 18px;font-size:17px;font-weight:700;">一句話結論：下一輪競爭比的不是單點更快，而是跨層整合、量產落地與平台治理。</div>
    <div class="footer-note" style="color:#D2DCE6">附註：本簡報著重相對排序與熱點辨識；若要升級董事會版，建議再補多資料源交叉驗證。</div>
  </section>
</div>
</body>
</html>`;
}

fs.writeFileSync(previewPath, htmlPreview(), 'utf8');

pptx.writeFile({ fileName: pptxPath }).then(() => {
  console.log(JSON.stringify({ pptxPath, previewPath }, null, 2));
}).catch((err) => {
  console.error(err);
  process.exit(1);
});
