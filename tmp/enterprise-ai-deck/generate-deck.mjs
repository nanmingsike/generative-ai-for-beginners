import PptxGenJS from 'pptxgenjs';
import { icon } from '@fortawesome/fontawesome-svg-core';
import {
  faBrain, faUser, faComments, faFileLines, faTable, faImage, faMicrophone,
  faGlobe, faBookOpen, faDatabase, faWrench, faEnvelope, faCalendarDays,
  faMemory, faClockRotateLeft, faListCheck, faDiagramProject, faGear,
  faBullseye, faEye, faArrowsRotate, faLaptop, faCube, faMobileScreenButton,
  faTriangleExclamation, faScaleBalanced, faLightbulb, faShieldHalved,
  faMagnifyingGlassChart, faPenNib, faChartLine, faPeopleGroup, faBuilding,
  faRobot, faArrowTrendUp, faCircleCheck, faCircleXmark, faBriefcase,
  faHeadset, faCartShopping, faUserTie, faMoneyBillTrendUp, faSitemap,
  faLayerGroup, faRoute, faRoad, faFlask, faChartPie, faClipboardCheck,
  faRocket, faPuzzlePiece, faLink, faLock, faGaugeHigh, faMedal,
  faPersonChalkboard, faHandshake, faArrowRight, faArrowDown,
  faBolt, faCompass, faNetworkWired, faCloud, faBoxesStacked,
  faStethoscope, faStore, faFileContract, faCodeBranch, faCircleNodes,
  faUserCheck, faRankingStar, faCheckDouble, faMap, faFlagCheckered,
  faPeopleArrows, faFingerprint, faArrowUpRightDots, faQuestion,
  faBook, faServer, faFileCircleCheck, faRepeat, faFilterCircleDollar,
  faClock, faCircleInfo, faArrowRightArrowLeft, faBarsProgress
} from '@fortawesome/free-solid-svg-icons';

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AIMatrix';
pptx.company = 'Enterprise AI Transformation Workshop';
pptx.subject = '企业AI转型与应用实践';
pptx.title = '企业AI转型与应用实践';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Microsoft YaHei',
  bodyFontFace: 'Microsoft YaHei',
  lang: 'zh-CN'
};
pptx.defineLayout({ name: 'CUSTOM', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM';

const C = {
  bg: '061127', bg2: '081A35', panel: '0C2244', panel2: '0B2B4D', panel3: '101F3C',
  border: '245B8F', blue: '2997FF', cyan: '21D4D8', purple: '8B5CF6',
  green: '29D39B', orange: 'F59E0B', red: 'F87171', yellow: 'F8D34F',
  white: 'F7FAFF', text: 'E8F0FA', muted: '9FB5CB', dim: '5E7897',
  dark: '020817', black: '000000'
};
const FONT = 'Microsoft YaHei';
const SHADOW = { type: 'outer', color: '000000', angle: 45, blur: 2, distance: 2, opacity: 0.22 };
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;
let slideNo = 0;

function faData(def, color = C.white) {
  const svg = icon(def, { styles: { color: `#${color}` } }).html.join('');
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
}

function addBg(slide, accent = C.blue, dense = false) {
  slide.background = { color: C.bg };
  slide.addShape(pptx.ShapeType.ellipse, { x: 9.7, y: -1.0, w: 4.8, h: 4.8, fill: { color: accent, transparency: 87 }, line: { color: accent, transparency: 100 } });
  slide.addShape(pptx.ShapeType.ellipse, { x: -2.2, y: 5.3, w: 5.2, h: 3.2, fill: { color: C.purple, transparency: 92 }, line: { color: C.purple, transparency: 100 } });
  for (let i = 0; i < (dense ? 34 : 20); i++) {
    const x = ((i * 37) % 127) / 10;
    const y = ((i * 53) % 67) / 10 + 0.2;
    const r = i % 4 === 0 ? 0.025 : 0.015;
    slide.addShape(pptx.ShapeType.ellipse, { x, y, w: r, h: r, fill: { color: i % 3 === 0 ? accent : C.muted, transparency: 30 }, line: { color: accent, transparency: 100 } });
  }
  slide.addShape(pptx.ShapeType.line, { x: 0.45, y: 1.37, w: 12.4, h: 0, line: { color: accent, transparency: 72, width: 1 } });
}

function addHeader(slide, title, subtitle, module = '企业AI转型与应用实践', accent = C.blue) {
  slideNo += 1;
  addBg(slide, accent);
  slide.addText(String(slideNo).padStart(2, '0'), { x: 0.32, y: 0.22, w: 0.62, h: 0.36, fontFace: FONT, fontSize: 13, bold: true, color: accent, margin: 0, align: 'center' });
  slide.addText(module, { x: 0.95, y: 0.21, w: 4.7, h: 0.34, fontFace: FONT, fontSize: 10, bold: true, color: C.muted, margin: 0 });
  slide.addText(title, { x: 0.62, y: 0.58, w: 12.1, h: 0.58, fontFace: FONT, fontSize: 28, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  if (subtitle) slide.addText(subtitle, { x: 0.65, y: 1.15, w: 11.8, h: 0.32, fontFace: FONT, fontSize: 13, color: C.muted, margin: 0, fit: 'shrink' });
}

function addFooter(slide, text = '', source = '') {
  if (text) slide.addText(text, { x: 0.65, y: 7.13, w: 7.2, h: 0.18, fontFace: FONT, fontSize: 7.5, color: C.dim, margin: 0, fit: 'shrink' });
  if (source) slide.addText(source, { x: 7.8, y: 7.11, w: 4.9, h: 0.2, fontFace: FONT, fontSize: 6.8, color: C.dim, margin: 0, align: 'right', fit: 'shrink' });
}

function addIcon(slide, def, x, y, w, color = C.white) {
  slide.addImage({ data: faData(def, color), x, y, w, h: w });
}

function card(slide, x, y, w, h, opts = {}) {
  const fill = opts.fill || C.panel;
  const line = opts.line || C.border;
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.08,
    fill: { color: fill, transparency: opts.transparency ?? 3 },
    line: { color: line, transparency: opts.lineTransparency ?? 5, width: opts.lineWidth || 1 },
    shadow: opts.shadow === false ? undefined : SHADOW
  });
  if (opts.icon) {
    const iw = opts.iconSize || 0.34;
    addIcon(slide, opts.icon, x + 0.22, y + 0.18, iw, opts.iconColor || C.cyan);
  }
  if (opts.title) slide.addText(opts.title, { x: x + (opts.icon ? 0.68 : 0.24), y: y + 0.16, w: w - (opts.icon ? 0.9 : 0.48), h: 0.35, fontFace: FONT, fontSize: opts.titleSize || 14.5, bold: true, color: opts.titleColor || C.white, margin: 0, fit: 'shrink' });
  if (opts.body) slide.addText(opts.body, { x: x + 0.24, y: y + (opts.title ? 0.58 : 0.24), w: w - 0.48, h: h - (opts.title ? 0.72 : 0.48), fontFace: FONT, fontSize: opts.bodySize || 10.5, color: opts.bodyColor || C.text, margin: 0.03, breakLine: false, valign: opts.valign || 'top', fit: 'shrink', bullet: opts.bullet });
}

function pill(slide, text, x, y, w, color = C.blue, iconDef = null) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.42, fill: { color, transparency: 77 }, line: { color, transparency: 20, width: 1 } });
  if (iconDef) addIcon(slide, iconDef, x + 0.12, y + 0.09, 0.22, color);
  slide.addText(text, { x: x + (iconDef ? 0.42 : 0.15), y: y + 0.07, w: w - (iconDef ? 0.52 : 0.3), h: 0.25, fontFace: FONT, fontSize: 9.5, bold: true, color: C.white, margin: 0, align: 'center', fit: 'shrink' });
}

function arrow(slide, x, y, w = 0.52, color = C.blue, direction = 'right') {
  const shape = direction === 'down' ? pptx.ShapeType.downArrow : pptx.ShapeType.rightArrow;
  slide.addShape(shape, { x, y, w, h: direction === 'down' ? 0.5 : 0.3, fill: { color, transparency: 5 }, line: { color, transparency: 100 } });
}

function iconCircle(slide, def, x, y, r = 0.56, color = C.blue, label = '') {
  slide.addShape(pptx.ShapeType.ellipse, { x, y, w: r, h: r, fill: { color, transparency: 75 }, line: { color, width: 1.2 } });
  addIcon(slide, def, x + r * 0.25, y + r * 0.25, r * 0.5, color);
  if (label) slide.addText(label, { x: x - 0.25, y: y + r + 0.1, w: r + 0.5, h: 0.28, fontFace: FONT, fontSize: 9, bold: true, color: C.text, margin: 0, align: 'center', fit: 'shrink' });
}

function callout(slide, text, x = 0.72, y = 6.3, w = 11.9, color = C.cyan, iconDef = faCircleInfo) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.56, fill: { color, transparency: 88 }, line: { color, transparency: 42, width: 1 } });
  addIcon(slide, iconDef, x + 0.18, y + 0.14, 0.25, color);
  slide.addText(text, { x: x + 0.55, y: y + 0.11, w: w - 0.73, h: 0.34, fontFace: FONT, fontSize: 11.5, bold: true, color: C.white, margin: 0, fit: 'shrink' });
}

function sectionSlide(num, title, subtitle, chips, accent = C.blue) {
  const slide = pptx.addSlide();
  slideNo += 1;
  slide.background = { color: C.bg };
  slide.addShape(pptx.ShapeType.ellipse, { x: 7.5, y: -1.3, w: 6.6, h: 6.6, fill: { color: accent, transparency: 82 }, line: { color: accent, transparency: 100 } });
  slide.addShape(pptx.ShapeType.arc, { x: 7.9, y: 0.0, w: 4.5, h: 4.5, adjustPoint: 0.5, rotate: 10, fill: { color: accent, transparency: 100 }, line: { color: accent, transparency: 40, width: 2 } });
  slide.addText(String(num).padStart(2, '0'), { x: 0.72, y: 0.72, w: 1.05, h: 0.6, fontFace: FONT, fontSize: 31, bold: true, color: accent, margin: 0 });
  slide.addText(title, { x: 0.72, y: 1.55, w: 8.7, h: 1.1, fontFace: FONT, fontSize: 34, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  slide.addText(subtitle, { x: 0.77, y: 2.84, w: 8.1, h: 0.62, fontFace: FONT, fontSize: 16, color: C.muted, margin: 0, fit: 'shrink' });
  chips.forEach((c, i) => pill(slide, c, 0.78 + i * 2.2, 4.35, 1.85, accent));
  slide.addText('ENTERPRISE AI TRANSFORMATION', { x: 0.76, y: 6.65, w: 4.8, h: 0.3, fontFace: 'Aptos', fontSize: 10, bold: true, color: C.dim, charSpacing: 1.6, margin: 0 });
  slide.addText(String(slideNo).padStart(2, '0'), { x: 12.1, y: 6.7, w: 0.6, h: 0.24, fontFace: FONT, fontSize: 9, bold: true, color: C.dim, margin: 0, align: 'right' });
  return slide;
}

function coverSlide() {
  const slide = pptx.addSlide();
  slideNo += 1;
  slide.background = { color: C.bg };
  addBg(slide, C.blue, true);
  slide.addShape(pptx.ShapeType.arc, { x: 7.55, y: 0.22, w: 5.1, h: 5.1, rotate: 35, fill: { color: C.blue, transparency: 100 }, line: { color: C.blue, transparency: 48, width: 2.4 } });
  slide.addShape(pptx.ShapeType.ellipse, { x: 8.65, y: 1.18, w: 2.9, h: 2.9, fill: { color: C.blue, transparency: 83 }, line: { color: C.cyan, width: 1.2, transparency: 15 } });
  addIcon(slide, faBrain, 9.34, 1.86, 1.55, C.cyan);
  pill(slide, '企业高管与核心骨干 · AI应用实践工作坊', 0.75, 0.7, 3.85, C.blue, faPeopleGroup);
  slide.addText('企业 AI 转型与应用实践', { x: 0.72, y: 1.55, w: 7.9, h: 0.95, fontFace: FONT, fontSize: 38, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  slide.addText('从理解 AI，到发现机会，再到推动企业落地', { x: 0.76, y: 2.72, w: 7.2, h: 0.48, fontFace: FONT, fontSize: 19, color: C.cyan, margin: 0, fit: 'shrink' });
  slide.addText('不是工具培训，而是一场关于工作方式、业务流程与组织能力的重新设计。', { x: 0.78, y: 3.48, w: 7.25, h: 0.75, fontFace: FONT, fontSize: 15, color: C.text, margin: 0, breakLine: false, fit: 'shrink' });
  ['理解 AI', '应用 AI', '落地 AI', '组织进化'].forEach((t, i) => pill(slide, t, 0.78 + i * 1.62, 4.75, 1.35, [C.blue, C.cyan, C.purple, C.green][i]));
  slide.addText('Enterprise AI Transformation & Application Workshop', { x: 0.78, y: 6.58, w: 6.7, h: 0.28, fontFace: 'Aptos', fontSize: 10, color: C.dim, charSpacing: 1.1, margin: 0 });
  slide.addText('Module 0–4 · 完整培训版', { x: 9.4, y: 6.57, w: 2.9, h: 0.28, fontFace: FONT, fontSize: 10, color: C.muted, margin: 0, align: 'right' });
}

function threeEraSlide() {
  const slide = pptx.addSlide();
  addHeader(slide, 'AI 正在进入每一种工作', '从“信息连接”到“人与智能协作”', '开场 · 为什么现在必须重新理解 AI', C.blue);
  const eras = [
    { x: 0.72, title: '互联网时代', sub: '连接信息', icon: faGlobe, color: C.blue, body: '人通过网页与搜索获得信息' },
    { x: 4.52, title: '移动互联网时代', sub: '连接服务', icon: faMobileScreenButton, color: C.purple, body: '人随时调用数字化服务' },
    { x: 8.32, title: 'AI 时代', sub: '连接智能', icon: faBrain, color: C.cyan, body: '智能开始参与理解、判断与执行' }
  ];
  eras.forEach((e, i) => {
    card(slide, e.x, 2.0, 3.25, 3.45, { fill: C.panel, line: e.color, icon: e.icon, iconColor: e.color, iconSize: 0.55, title: e.title, titleSize: 17, body: `${e.sub}\n\n${e.body}`, bodySize: 12.4, bodyColor: C.text });
    if (i < 2) arrow(slide, e.x + 3.43, 3.4, 0.48, e.color);
  });
  callout(slide, '过去的软件帮助人工作；今天的 AI 开始参与完成工作。', 1.5, 6.05, 10.3, C.cyan, faBolt);
  addFooter(slide, '开场认知：AI 不只是一个新工具，而是一种新的生产力基础设施。');
}

function beforeAfterWorkSlide() {
  const slide = pptx.addSlide();
  addHeader(slide, 'AI 不是工具升级，而是工作方式改变', '从“人操作软件”到“人 + AI 协作系统完成目标”', '开场 · 为什么现在必须重新理解 AI', C.cyan);
  card(slide, 0.72, 2.05, 5.45, 3.7, { fill: C.panel3, line: C.dim, icon: faBriefcase, iconColor: C.muted, title: '过去：人使用软件完成任务', titleSize: 17, body: '', shadow: true });
  const old = [
    { icon: faUser, text: '人理解任务' }, { icon: faWrench, text: '人操作工具' }, { icon: faFileCircleCheck, text: '人交付结果' }
  ];
  old.forEach((it, i) => { iconCircle(slide, it.icon, 1.25 + i * 1.55, 3.15, 0.68, C.muted, it.text); if (i < 2) arrow(slide, 2.08 + i * 1.55, 3.35, 0.42, C.dim); });
  card(slide, 7.15, 2.05, 5.45, 3.7, { fill: C.panel2, line: C.cyan, icon: faCircleNodes, iconColor: C.cyan, title: '现在：目标驱动的人机协作系统', titleSize: 17, body: '', shadow: true });
  const now = [
    { icon: faBullseye, text: '目标' }, { icon: faPeopleArrows, text: '人 + AI' }, { icon: faArrowTrendUp, text: '业务结果' }
  ];
  now.forEach((it, i) => { iconCircle(slide, it.icon, 7.7 + i * 1.55, 3.15, 0.68, [C.blue, C.cyan, C.green][i], it.text); if (i < 2) arrow(slide, 8.53 + i * 1.55, 3.35, 0.42, C.cyan); });
  callout(slide, '真正的变化，不是多了一个 AI 工具，而是重新分配“理解、生成、执行与判断”。', 1.15, 6.05, 11.05, C.blue, faArrowRightArrowLeft);
  addFooter(slide);
}

function progressionSlide() {
  const slide = pptx.addSlide();
  addHeader(slide, '生成式 AI 正在从“会回答”走向“能工作”', 'Chatbot → Copilot → Workflow → Agent → AI Native', '开场 · 为什么现在必须重新理解 AI', C.purple);
  const stages = [
    { t: 'Chatbot', s: '回答问题', icon: faComments, c: C.blue },
    { t: 'Copilot', s: '协助工作', icon: faPeopleArrows, c: C.cyan },
    { t: 'Workflow', s: '稳定复用', icon: faDiagramProject, c: C.purple },
    { t: 'Agent', s: '持续执行', icon: faRobot, c: C.orange },
    { t: 'AI Native', s: '重构业务', icon: faBuilding, c: C.green }
  ];
  stages.forEach((st, i) => {
    const x = 0.55 + i * 2.52;
    slide.addShape(pptx.ShapeType.chevron, { x, y: 2.35, w: 2.35, h: 2.15, fill: { color: st.c, transparency: 75 - i * 4 }, line: { color: st.c, width: 1.1 } });
    addIcon(slide, st.icon, x + 0.76, 2.67, 0.54, st.c);
    slide.addText(st.t, { x: x + 0.28, y: 3.35, w: 1.7, h: 0.3, fontFace: FONT, fontSize: 14.5, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(st.s, { x: x + 0.28, y: 3.72, w: 1.7, h: 0.25, fontFace: FONT, fontSize: 10, color: C.text, align: 'center', margin: 0, fit: 'shrink' });
  });
  card(slide, 1.0, 5.08, 11.3, 0.84, { fill: C.panel3, line: C.purple, icon: faLightbulb, iconColor: C.yellow, title: '核心认知', body: '模型是“大脑”；企业真正使用的是“大脑 + 知识 + 工具 + 记忆 + 方法 + 控制”的完整工作系统。', bodySize: 10.4, titleSize: 12.5 });
  addFooter(slide);
}

function whyDifferentSlide() {
  const slide = pptx.addSlide();
  addHeader(slide, '为什么今天的 AI 和过去不同？', '三次跃迁，让 AI 从“规则执行”变成“目标驱动”', '开场 · 为什么现在必须重新理解 AI', C.blue);
  const items = [
    { x: 0.7, title: '过去：规则驱动', icon: faCodeBranch, color: C.dim, body: '如果 A，则 B\n适合固定、可预设的流程' },
    { x: 4.55, title: '现在：生成式 AI', icon: faBrain, color: C.blue, body: '理解、生成、推理\n适合复杂信息与非结构化任务' },
    { x: 8.4, title: '进一步：Agentic AI', icon: faRobot, color: C.cyan, body: '目标、计划、工具、反馈\n适合多步骤、持续推进的任务' }
  ];
  items.forEach((it, i) => { card(slide, it.x, 2.1, 3.35, 3.55, { fill: C.panel, line: it.color, icon: it.icon, iconColor: it.color, title: it.title, titleSize: 16, body: it.body, bodySize: 12.2 }); if (i < 2) arrow(slide, it.x + 3.47, 3.6, 0.43, it.color); });
  callout(slide, '企业的瓶颈正在从“模型够不够聪明”，转向“如何让 AI 可靠地进入真实工作”。', 1.1, 6.08, 11.2, C.cyan, faGaugeHigh);
  addFooter(slide, '', '参考：OpenAI Frontier / Presence 的企业 Agent 运行逻辑（2026）');
}

function rightQuestionSlide() {
  const slide = pptx.addSlide();
  addHeader(slide, '企业真正的问题：不是“要不要 AI”', '而是：让 AI 在什么任务中做到哪一步？', '开场 · 为什么现在必须重新理解 AI', C.orange);
  card(slide, 0.72, 2.05, 5.5, 3.5, { fill: C.panel3, line: C.red, icon: faCircleXmark, iconColor: C.red, title: '错误起点：先看工具，再找场景', titleSize: 17, body: '有什么大模型？\n买哪个 Agent 平台？\n大家学什么 Prompt？', bodySize: 13 });
  arrow(slide, 6.35, 3.45, 0.55, C.orange);
  card(slide, 7.05, 2.05, 5.55, 3.5, { fill: C.panel2, line: C.green, icon: faCircleCheck, iconColor: C.green, title: '正确起点：从业务问题与任务出发', titleSize: 17, body: '业务目标是什么？\n工作如何完成？\nAI 可以介入哪里？\n如何证明业务价值？', bodySize: 13 });
  callout(slide, '技术决定“能不能做”；业务设计决定“值不值得做”。', 2.2, 6.05, 8.9, C.orange, faCompass);
  addFooter(slide);
}

function courseMapSlide() {
  const slide = pptx.addSlide();
  addHeader(slide, '今天，我们共同回答四个问题', '从认知对齐，到场景发现，再到落地行动', '开场 · 课程导航', C.green);
  const items = [
    { n: '01', t: '理解 AI', q: 'AI 到底变成了什么？', c: C.blue, icon: faBrain },
    { n: '02', t: '应用 AI', q: 'AI 可以改变哪些工作？', c: C.cyan, icon: faPuzzlePiece },
    { n: '03', t: '落地 AI', q: '企业应该如何开始？', c: C.purple, icon: faRocket },
    { n: '04', t: '人与组织', q: '未来组织如何进化？', c: C.green, icon: faPeopleGroup }
  ];
  items.forEach((it, i) => {
    const x = 0.65 + i * 3.12;
    card(slide, x, 2.0, 2.75, 3.8, { fill: C.panel, line: it.c, icon: it.icon, iconColor: it.c, title: `${it.n}  ${it.t}`, titleSize: 17, body: it.q, bodySize: 13 });
  });
  callout(slide, '最终目标：每位学员不仅“听懂 AI”，还能够识别一个真实机会，并说明下一步行动。', 1.0, 6.15, 11.3, C.green, faFlagCheckered);
  addFooter(slide);
}

function evolutionSlide({ title, subtitle, leftTitle, leftIcon, leftItems, centerTitle, centerIcon, outputTitle, outputItems, insight, boundary, accent = C.blue, source = '' }) {
  const slide = pptx.addSlide();
  addHeader(slide, title, subtitle, 'Module 1 · 理解 AI：从聊天机器人到数字员工', accent);
  card(slide, 0.72, 1.83, 3.1, 3.72, { fill: C.panel3, line: accent, icon: leftIcon, iconColor: accent, title: leftTitle, titleSize: 15.5, body: '' });
  leftItems.forEach((it, i) => {
    const yy = 2.55 + i * 0.67;
    if (it.icon) addIcon(slide, it.icon, 0.98, yy + 0.02, 0.25, it.color || accent);
    slide.addText(it.text || it, { x: 1.36, y: yy, w: 2.15, h: 0.3, fontFace: FONT, fontSize: 10.2, color: C.text, margin: 0, fit: 'shrink' });
    if (i < leftItems.length - 1) slide.addShape(pptx.ShapeType.line, { x: 0.98, y: yy + 0.43, w: 2.5, h: 0, line: { color: C.border, transparency: 55, width: 0.6 } });
  });
  arrow(slide, 3.96, 3.36, 0.58, accent);
  card(slide, 4.65, 2.12, 3.2, 3.1, { fill: C.panel2, line: accent, title: centerTitle, titleSize: 18.5, body: '' });
  slide.addShape(pptx.ShapeType.ellipse, { x: 5.55, y: 2.78, w: 1.4, h: 1.4, fill: { color: accent, transparency: 75 }, line: { color: accent, width: 1.2 } });
  addIcon(slide, centerIcon, 5.9, 3.13, 0.7, accent);
  arrow(slide, 7.98, 3.36, 0.58, accent);
  card(slide, 8.68, 1.83, 3.92, 3.72, { fill: C.panel3, line: accent, icon: faCircleCheck, iconColor: accent, title: outputTitle, titleSize: 15.5, body: outputItems.join('\n'), bodySize: 11.1 });
  card(slide, 0.72, 5.82, 7.3, 0.88, { fill: C.panel2, line: accent, icon: faArrowTrendUp, iconColor: accent, title: '能力变化', body: insight, bodySize: 9.8, titleSize: 11.8, shadow: false });
  card(slide, 8.25, 5.82, 4.35, 0.88, { fill: boundary ? '2A1B1A' : C.panel3, line: boundary ? C.orange : accent, icon: boundary ? faTriangleExclamation : faLightbulb, iconColor: boundary ? C.orange : accent, title: boundary ? '此时的边界' : '关键认知', body: boundary || insight, bodySize: 9.4, titleSize: 11.6, shadow: false });
  addFooter(slide, '', source);
}

function module1Slides() {
  sectionSlide(1, '理解 AI：从聊天机器人到数字员工', '先理解能力如何生长，才知道企业应该如何设计 AI。', ['大模型', '知识', '工具', '记忆', 'Agent'], C.blue);
  evolutionSlide({
    title: 'AI 最初只有一个“大脑”', subtitle: '会生成，但看不到你的文件、业务与真实现场',
    leftTitle: '输入', leftIcon: faUser, leftItems: [{ icon: faUser, text: '用户问题 / 要求' }, { icon: faFileLines, text: 'System / User Prompt' }],
    centerTitle: '大模型', centerIcon: faBrain, outputTitle: '输出', outputItems: ['回答', '解释', '总结', '内容生成'],
    insight: '大模型基于训练中学到的知识与模式生成结果。', boundary: '看不到企业文件，不了解最新现场，也不能操作任何系统。', accent: C.blue
  });
  evolutionSlide({
    title: 'AI 获得上下文：开始理解当前任务', subtitle: '从一次问答，变成持续对话与共同推进',
    leftTitle: '当前工作台 / Context', leftIcon: faComments, leftItems: [{ icon: faComments, text: '对话历史' }, { icon: faBullseye, text: '当前目标' }, { icon: faBriefcase, text: '项目背景' }, { icon: faUserTie, text: '用户角色与偏好' }],
    centerTitle: '上下文 + 大模型', centerIcon: faBrain, outputTitle: '更贴近任务的结果', outputItems: ['追问澄清', '持续修改', '逐步收敛', '共同完成'],
    insight: '对话成为临时工作台，AI 可以结合已有信息继续工作。', boundary: '仍主要由人推动下一步；上下文过长或缺失会影响质量。', accent: C.blue
  });
  evolutionSlide({
    title: 'AI 获得眼睛：开始看见真实材料', subtitle: '输入不再只有文字，AI 可以直接理解工作资料',
    leftTitle: '多模态输入', leftIcon: faEye, leftItems: [{ icon: faFileLines, text: 'PDF / Word' }, { icon: faTable, text: 'Excel / 表格' }, { icon: faImage, text: '图片 / 截图' }, { icon: faMicrophone, text: '音频 / 视频' }],
    centerTitle: '材料理解 + 大模型', centerIcon: faEye, outputTitle: '理解与转换', outputItems: ['提取', '总结', '比较', '生成草稿'],
    insight: '从“你描述给 AI”升级为“AI 直接阅读真实材料”。', boundary: '复杂表格、模糊图像和超长材料仍可能被误读。', accent: C.cyan
  });
  evolutionSlide({
    title: 'AI 获得知识来源：从猜测到基于事实', subtitle: '企业知识库与联网检索，让回答开始带依据',
    leftTitle: '知识来源', leftIcon: faBookOpen, leftItems: [{ icon: faBuilding, text: '企业知识库' }, { icon: faFileContract, text: '制度 / SOP / 文档' }, { icon: faGlobe, text: '互联网最新信息' }, { icon: faDatabase, text: '行业与业务数据' }],
    centerTitle: '检索 + 大模型', centerIcon: faMagnifyingGlassChart, outputTitle: '基于依据的回答', outputItems: ['结论', '依据', '来源', '可追溯'],
    insight: 'RAG 不是让模型永久记住，而是把相关资料临时拿到工作台。', boundary: '来源质量、检索完整性和资料时效会直接影响结果。', accent: C.cyan
  });
  evolutionSlide({
    title: 'AI 获得手脚：开始执行真实动作', subtitle: '连接工具与系统，从“告诉你怎么做”升级为“替你去做”',
    leftTitle: '可调用的工具与系统', leftIcon: faWrench, leftItems: [{ icon: faBriefcase, text: 'CRM / ERP' }, { icon: faEnvelope, text: '邮件 / 日历' }, { icon: faDatabase, text: '数据库 / API' }, { icon: faGlobe, text: '浏览器 / 应用' }],
    centerTitle: '大模型 + 工具能力', centerIcon: faWrench, outputTitle: '业务动作', outputItems: ['生成文件', '更新记录', '起草邮件', '创建会议'],
    insight: '业务价值第一次从“建议”升级为“状态变化与真实动作”。', boundary: '能力越强，错误成本越高，必须设计权限、确认和回滚。', accent: C.green
  });
  evolutionSlide({
    title: 'AI 获得记忆：从一次合作到长期伙伴', subtitle: '保存偏好、已确认决策与任务状态',
    leftTitle: '记忆与状态', leftIcon: faMemory, leftItems: [{ icon: faMemory, text: '长期记忆' }, { icon: faUserTie, text: '用户偏好' }, { icon: faClockRotateLeft, text: '历史记录' }, { icon: faBarsProgress, text: '任务状态' }],
    centerTitle: '大模型 + Memory', centerIcon: faDatabase, outputTitle: '连续协作', outputItems: ['记住你', '记住决定', '接着做', '更个性化'],
    insight: 'AI 不必每次从零开始，可以跨会话、跨阶段持续协作。', boundary: '正式业务事实应保存在系统中；记忆需要可查看、更正与删除。', accent: C.purple
  });
  evolutionSlide({
    title: 'AI 获得方法：从聪明到稳定可靠', subtitle: '把模板、SOP、工作步骤和最佳实践沉淀为可复用方法',
    leftTitle: '方法库 / Skills', leftIcon: faListCheck, leftItems: [{ icon: faListCheck, text: 'SOP / 检查清单' }, { icon: faMedal, text: '最佳实践' }, { icon: faFileLines, text: '模板 / 提示词' }, { icon: faShieldHalved, text: '业务规则' }],
    centerTitle: '大模型 + 方法', centerIcon: faGear, outputTitle: '稳定输出', outputItems: ['更一致', '可复用', '可评估', '可持续优化'],
    insight: '记忆回答“之前发生了什么”；方法回答“这类事情应该怎么做”。', boundary: '错误方法也会被稳定复制，因此方法需要验证与持续改进。', accent: C.yellow
  });
  const slide = pptx.addSlide();
  addHeader(slide, 'AI 形成执行循环：成为 Agentic 助手', '围绕目标自主规划、行动、观察与纠偏', 'Module 1 · 理解 AI：从聊天机器人到数字员工', C.purple);
  const cx = 6.55, cy = 3.75;
  slide.addShape(pptx.ShapeType.ellipse, { x: 5.55, y: 2.72, w: 2.05, h: 2.05, fill: { color: C.purple, transparency: 80 }, line: { color: C.purple, width: 1.3 } });
  addIcon(slide, faRobot, 6.12, 3.23, 0.82, C.cyan);
  slide.addText('Agent Loop', { x: 5.86, y: 4.1, w: 1.45, h: 0.28, fontFace: FONT, fontSize: 12, bold: true, color: C.white, align: 'center', margin: 0 });
  const loop = [
    { x: 5.55, y: 1.74, icon: faBullseye, t: '1 理解目标', c: C.blue },
    { x: 8.5, y: 2.4, icon: faListCheck, t: '2 制定计划', c: C.purple },
    { x: 8.75, y: 4.65, icon: faWrench, t: '3 调用工具', c: C.orange },
    { x: 3.12, y: 4.65, icon: faEye, t: '4 观察结果', c: C.cyan },
    { x: 2.85, y: 2.4, icon: faClipboardCheck, t: '5 检查达标', c: C.green }
  ];
  loop.forEach(l => card(slide, l.x, l.y, 1.9, 0.82, { fill: C.panel, line: l.c, icon: l.icon, iconColor: l.c, title: l.t, titleSize: 11.8, shadow: false }));
  slide.addShape(pptx.ShapeType.arc, { x: 3.1, y: 1.55, w: 6.2, h: 4.9, rotate: 20, fill: { color: C.purple, transparency: 100 }, line: { color: C.purple, transparency: 35, width: 2.8, beginArrowType: 'none', endArrowType: 'triangle' } });
  callout(slide, 'Agentic 不等于完全自主：AI 在边界内推进，人只在关键节点介入。', 1.15, 6.34, 11.0, C.purple, faShieldHalved);
  addFooter(slide);

  const env = pptx.addSlide();
  addHeader(env, '给 AI 一个“工作环境”：成为数字员工', '持续运行、安全试验、远程管理，让 Agent 真正进入工作', 'Module 1 · 理解 AI：从聊天机器人到数字员工', C.cyan);
  const envs = [
    { x: 0.75, title: '一台电脑', sub: '云端工作空间', icon: faLaptop, body: '持续运行\n保持任务状态\n可 24 小时工作', c: C.blue },
    { x: 4.55, title: '一个试验场地', sub: 'Sandbox', icon: faCube, body: '隔离环境中尝试\n执行、验证、失败\n支持回滚', c: C.purple },
    { x: 8.35, title: '远程控制入口', sub: 'IM / 手机端', icon: faMobileScreenButton, body: '随时发起任务\n查看进度与结果\n关键节点确认', c: C.cyan }
  ];
  envs.forEach(e => card(env, e.x, 2.0, 3.35, 3.75, { fill: C.panel, line: e.c, icon: e.icon, iconColor: e.c, title: e.title, titleSize: 17, body: `${e.sub}\n\n${e.body}`, bodySize: 12.2 }));
  callout(env, '这不是让 AI 更聪明，而是让 AI 拥有完成工作的基础设施。', 1.15, 6.14, 11.0, C.cyan, faCloud);
  addFooter(env);

  const sum = pptx.addSlide();
  addHeader(sum, 'AI 能力进化：模型没有消失，系统不断长大', '同一个“大脑”，被知识、工具、记忆、方法和控制包围', 'Module 1 · 理解 AI：从聊天机器人到数字员工', C.blue);
  const layers = [
    { y: 5.45, w: 8.4, t: '大模型：理解、推理、生成', c: C.blue },
    { y: 4.78, w: 9.1, t: '上下文与多模态：看见当前工作', c: C.cyan },
    { y: 4.11, w: 9.8, t: '知识与数据：获得事实依据', c: C.green },
    { y: 3.44, w: 10.5, t: '工具与业务系统：执行真实动作', c: C.orange },
    { y: 2.77, w: 11.2, t: '记忆与方法：持续、稳定、复用', c: C.purple },
    { y: 2.10, w: 11.9, t: '执行循环与控制机制：成为 Agentic 工作系统', c: C.yellow }
  ];
  layers.forEach((l, i) => {
    const x = (SLIDE_W - l.w) / 2;
    slideShapeRound(sum, x, l.y, l.w, 0.48, l.c, 72 - i * 3, l.t);
  });
  callout(sum, 'Agentic 助手 = 大模型 + 上下文与知识 + 工具与连接 + 记忆与状态 + 方法 + 执行循环 + 控制。', 0.82, 6.36, 11.7, C.blue, faLayerGroup);
  addFooter(sum);

  const boundary = pptx.addSlide();
  addHeader(boundary, '但 AI 仍然有边界', '能力越强，越不能把“流畅”误认为“可靠”', 'Module 1 · 理解 AI：从聊天机器人到数字员工', C.orange);
  const bs = [
    { x: 0.72, t: '知识边界', b: '依赖资料与训练知识\n可能过时、遗漏或编造', icon: faBrain, c: C.blue },
    { x: 3.78, t: '理解边界', b: '只能理解可见上下文\n可能误解意图与隐含背景', icon: faQuestion, c: C.green },
    { x: 6.84, t: '判断边界', b: '缺少价值观与真实责任\n重要判断需要人监督', icon: faScaleBalanced, c: C.orange },
    { x: 9.9, t: '责任边界', b: 'AI 不能独立承担责任\n组织必须明确责任主体', icon: faShieldHalved, c: C.purple }
  ];
  bs.forEach(b => card(boundary, b.x, 2.05, 2.72, 3.7, { fill: C.panel, line: b.c, icon: b.icon, iconColor: b.c, title: b.t, titleSize: 17, body: b.b, bodySize: 11.6 }));
  callout(boundary, 'AI 最大的风险，不是它完全不会做，而是它做得“看起来很像对的”。', 1.2, 6.15, 10.9, C.orange, faTriangleExclamation);
  addFooter(boundary);

  const collab = pptx.addSlide();
  addHeader(collab, '人机协同：优势互补，共同创造价值', '不是 AI 替代人，而是重新设计人与 AI 的分工', 'Module 1 · 理解 AI：从聊天机器人到数字员工', C.green);
  card(collab, 0.75, 1.95, 4.7, 3.9, { fill: C.panel2, line: C.blue, icon: faRobot, iconColor: C.blue, title: 'AI 更擅长', titleSize: 18, body: '快速处理大量信息\n持续监测与发现模式\n内容生成与重复执行\n7×24 不间断工作', bodySize: 12.4 });
  card(collab, 7.88, 1.95, 4.7, 3.9, { fill: C.panel2, line: C.green, icon: faUserTie, iconColor: C.green, title: '人更重要', titleSize: 18, body: '定义目标与价值标准\n复杂关系与情境判断\n创造、共情与领导力\n承担最终责任', bodySize: 12.4 });
  slideShapeRound(collab, 5.63, 2.55, 2.05, 2.05, C.purple, 78, '协同\n\n更好的决策\n更高的价值');
  callout(collab, '企业 AI 成熟不是“最大化自主”，而是在合适任务中设计合适能力与合适控制。', 1.0, 6.18, 11.3, C.green, faHandshake);
  addFooter(collab);

  const trans = pptx.addSlide();
  addHeader(trans, '从理解 AI，到应用 AI', '理解能力进化，才知道如何发现机会、设计方案并持续迭代', 'Module 1 · 总结与过渡', C.cyan);
  const chain = [
    { x: 0.6, t: '理解 AI 能力', s: '知道它能做什么', icon: faBrain, c: C.blue },
    { x: 3.15, t: '发现业务机会', s: '找到高价值场景', icon: faBullseye, c: C.green },
    { x: 5.7, t: '设计解决方案', s: '人机协同的流程与系统', icon: faSitemap, c: C.purple },
    { x: 8.25, t: '落地与迭代', s: '持续优化，创造价值', icon: faArrowUpRightDots, c: C.orange },
    { x: 10.8, t: '企业智能', s: '形成组织能力', icon: faBuilding, c: C.cyan }
  ];
  chain.forEach((c, i) => {
    card(trans, c.x, 2.4, 2.05, 2.45, { fill: C.panel, line: c.c, icon: c.icon, iconColor: c.c, title: c.t, titleSize: 13.2, body: c.s, bodySize: 9.8 });
    if (i < chain.length - 1) arrow(trans, c.x + 2.12, 3.42, 0.38, c.c);
  });
  callout(trans, 'AI 不是终点，而是企业智能化的起点。', 2.15, 5.75, 9.0, C.cyan, faRocket);
  addFooter(trans);
}

function slideShapeRound(slide, x, y, w, h, color, transparency, text) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color, transparency }, line: { color, transparency: 12, width: 1.1 }, shadow: SHADOW });
  slide.addText(text, { x: x + 0.18, y: y + 0.08, w: w - 0.36, h: h - 0.16, fontFace: FONT, fontSize: h > 1 ? 12 : 10.5, bold: true, color: C.white, align: 'center', valign: 'mid', margin: 0.02, fit: 'shrink' });
}

function caseSlide({ title, subtitle, background, traditional, aiFlow, value, human, accent = C.cyan, source = '' }) {
  const slide = pptx.addSlide();
  addHeader(slide, title, subtitle, 'Module 2 · 应用 AI：从个人提效到企业智能化', accent);
  card(slide, 0.68, 1.82, 3.15, 4.58, { fill: C.panel3, line: C.dim, icon: faClockRotateLeft, iconColor: C.muted, title: '过去怎么做？', titleSize: 15.5, body: `${background}\n\n${traditional}`, bodySize: 10.8 });
  card(slide, 4.05, 1.82, 5.0, 4.58, { fill: C.panel2, line: accent, icon: faRobot, iconColor: accent, title: 'AI 如何重新设计？', titleSize: 15.5, body: aiFlow, bodySize: 11.1 });
  card(slide, 9.28, 1.82, 3.35, 2.16, { fill: C.panel, line: C.green, icon: faArrowTrendUp, iconColor: C.green, title: '业务价值', titleSize: 14, body: value, bodySize: 10.5 });
  card(slide, 9.28, 4.18, 3.35, 2.22, { fill: C.panel, line: C.orange, icon: faHandshake, iconColor: C.orange, title: '人机分工', titleSize: 14, body: human, bodySize: 10.3 });
  addFooter(slide, '', source);
}

function workflowSlide(title, subtitle, steps, accent = C.cyan, source = '') {
  const slide = pptx.addSlide();
  addHeader(slide, title, subtitle, 'Module 2 · 应用 AI：从个人提效到企业智能化', accent);
  const n = steps.length;
  const cardW = Math.min(1.65, 10.8 / n);
  const gap = (11.8 - cardW * n) / (n - 1);
  steps.forEach((st, i) => {
    const x = 0.75 + i * (cardW + gap);
    slideShapeRound(slide, x, 2.55, cardW, 1.65, st.color || accent, 78, `${i + 1}\n${st.title}`);
    if (st.icon) addIcon(slide, st.icon, x + cardW / 2 - 0.18, 2.78, 0.36, st.color || accent);
    if (i < n - 1) arrow(slide, x + cardW + 0.05, 3.2, Math.max(0.25, gap - 0.1), accent);
  });
  callout(slide, '稳定工作不是靠一次 Prompt，而是靠目标、步骤、知识、工具与人工控制的组合。', 1.05, 5.45, 11.2, accent, faDiagramProject);
  addFooter(slide, '', source);
}

function module2Slides() {
  sectionSlide(2, '应用 AI：从个人提效到企业智能化', '从“我如何更高效”，逐步走向“企业如何重新设计业务”。', ['个人', '团队', '流程', '企业'], C.cyan);
  const four = pptx.addSlide();
  addHeader(four, '企业 AI 应用的四个层次', '任务越来越复杂，价值越来越深入', 'Module 2 · 应用 AI：从个人提效到企业智能化', C.cyan);
  const levels = [
    { y: 5.35, w: 8.0, t: 'Level 1 · AI Assistant：个人效率', c: C.blue, s: '问 → 做一个' },
    { y: 4.55, w: 8.85, t: 'Level 2 · AI Workflow：团队复用', c: C.cyan, s: '经验 → 可靠流程' },
    { y: 3.75, w: 9.7, t: 'Level 3 · AI Agent：完整任务', c: C.purple, s: '目标 → 持续执行' },
    { y: 2.95, w: 10.55, t: 'Level 4 · AI Native：业务重构', c: C.green, s: '功能 → 结果' }
  ];
  levels.forEach((l, i) => {
    const x = (13.333 - l.w) / 2;
    slide.addShape(pptx.ShapeType.chevron, { x, y: l.y, w: l.w, h: 0.62, fill: { color: l.c, transparency: 60 + i * 3 }, line: { color: l.c, width: 1 } });
    slide.addText(l.t, { x: x + 0.28, y: l.y + 0.12, w: l.w - 2.4, h: 0.28, fontFace: FONT, fontSize: 12.5, bold: true, color: C.white, margin: 0, fit: 'shrink' });
    slide.addText(l.s, { x: x + l.w - 2.2, y: l.y + 0.12, w: 1.65, h: 0.26, fontFace: FONT, fontSize: 9, color: C.text, margin: 0, align: 'right', fit: 'shrink' });
  });
  callout(four, '成熟不是“用了多少 AI 工具”，而是“多少工作被重新设计并形成组织能力”。', 1.0, 6.34, 11.3, C.cyan, faRankingStar);
  addFooter(four);

  const copilot = pptx.addSlide();
  addHeader(copilot, '个人 AI 助手：从“自己做”到“与 AI 一起做”', 'AI 首先改变每个人每天重复的信息工作', 'Module 2 · Part 1 · 个人 AI 助手', C.blue);
  const flow = [
    { x: 0.7, t: '提出目标', icon: faBullseye, c: C.blue },
    { x: 3.25, t: '提供背景', icon: faFileLines, c: C.cyan },
    { x: 5.8, t: 'AI 生成', icon: faBrain, c: C.purple },
    { x: 8.35, t: '人工判断', icon: faUserCheck, c: C.orange },
    { x: 10.9, t: '优化沉淀', icon: faMedal, c: C.green }
  ];
  flow.forEach((f, i) => { card(copilot, f.x, 2.35, 1.9, 2.2, { fill: C.panel, line: f.c, icon: f.icon, iconColor: f.c, title: f.t, titleSize: 13.5, body: '', shadow: true }); if (i < flow.length - 1) arrow(copilot, f.x + 1.97, 3.27, 0.45, f.c); });
  callout(copilot, 'AI 高手不是“一次问得更好”，而是能把目标、资料、判断和复盘组织成工作系统。', 1.05, 5.55, 11.2, C.blue, faLightbulb);
  addFooter(copilot);

  caseSlide({ title: '案例：AI 研究助手', subtitle: '从搜索信息，到生成洞察', background: '市场经理需要准备竞品研究。', traditional: '搜索网页 → 阅读资料 → 复制整理 → 写报告\n大量时间花在信息搬运。', aiFlow: '1. 定义研究目标与比较维度\n2. AI 搜集并筛选公开资料\n3. 整理证据、差异与趋势\n4. 生成分析初稿\n5. 人判断结论与业务含义', value: '更快形成全景\n把时间投入到判断\n报告结构更一致', human: 'AI：搜集、整理、比较\n人：验证、判断、洞察', accent: C.blue });
  caseSlide({ title: '案例：AI 写作与沟通助手', subtitle: '从“写内容”到“共同创造内容”', background: '销售准备客户方案与跟进邮件。', traditional: '翻找资料 → 从零搭结构 → 反复修改措辞\n质量依赖个人表达能力。', aiFlow: '客户背景 + 产品资料 + 历史案例\n→ AI 生成客户画像\n→ 方案结构与核心论点\n→ 邮件 / PPT 初稿\n→ 人补充判断与关系策略', value: '降低第一稿成本\n提高结构与表达质量\n更多时间经营客户', human: 'AI：第一稿与多版本\n人：策略、关系与承诺', accent: C.purple });
  caseSlide({ title: '案例：AI 经营分析助手', subtitle: '从看报表，到发现问题', background: '管理者每月查看销售、成本与运营数据。', traditional: '拉数据 → 做图表 → 查异常 → 写分析\n结论容易停留在“发生了什么”。', aiFlow: '读取 Excel / BI 数据\n→ 识别异常与趋势\n→ 拆分区域、产品、客户\n→ 提出原因假设\n→ 生成经营分析与行动建议', value: '更快发现异常\n统一分析方法\n缩短决策周期', human: 'AI：分析与假设\n人：验证原因、决定行动', accent: C.green });
  caseSlide({ title: '案例：AI 会议助手', subtitle: '把会议从“信息交换”变成“行动推动”', background: '会议多、记录散、行动项经常失踪。', traditional: '会前临时准备\n会中人工记录\n会后重新整理与催办。', aiFlow: '会前：汇总背景与议题\n会中：记录重点、争议与决策\n会后：生成纪要、责任人、截止时间\n持续：跟踪进度与提醒异常', value: '减少信息损失\n提高决策闭环率\n缩短会后整理时间', human: 'AI：记录、整理、跟踪\n人：讨论、判断、承诺', accent: C.cyan });

  const maturity = pptx.addSlide();
  addHeader(maturity, '个人 AI 能力进化：不是会不会用，而是能否设计工作系统', 'Consumer → Tinkerer → Builder → Architect', 'Module 2 · Part 1 · 个人 AI 助手', C.blue);
  const ms = [
    { x: 0.7, t: 'Consumer', zh: '使用者', s: '问 AI\n获得答案', c: C.dim, icon: faComments },
    { x: 3.75, t: 'Tinkerer', zh: '探索者', s: '试工具\n优化提示', c: C.blue, icon: faFlask },
    { x: 6.8, t: 'Builder', zh: '构建者', s: '组合资料、工具与方法', c: C.purple, icon: faPuzzlePiece },
    { x: 9.85, t: 'Architect', zh: '架构者', s: '设计可靠的人机工作系统', c: C.green, icon: faSitemap }
  ];
  ms.forEach((m, i) => { card(maturity, m.x, 2.05 + (3 - i) * 0.28, 2.75, 3.3 + i * 0.28, { fill: C.panel, line: m.c, icon: m.icon, iconColor: m.c, title: `${m.t} · ${m.zh}`, titleSize: 15.5, body: m.s, bodySize: 11.5 }); });
  callout(maturity, '真正的跃迁：从“让 AI 回答”，到“让 AI 按可靠方式持续完成工作”。', 1.15, 6.2, 11.0, C.blue, faArrowUpRightDots);
  addFooter(maturity);

  const personalStack = pptx.addSlide();
  addHeader(personalStack, '个人 AI 高手的秘密：不是 Prompt，而是工作系统', '模型只是入口，价值来自完整组合', 'Module 2 · Part 1 · 个人 AI 助手', C.purple);
  const stack = [
    { t: '模型', s: '通用智能', icon: faBrain, c: C.blue },
    { t: '知识', s: '个人资料与领域经验', icon: faBookOpen, c: C.cyan },
    { t: '工具', s: '文件、搜索、应用', icon: faWrench, c: C.green },
    { t: '方法', s: '模板、流程、检查', icon: faListCheck, c: C.yellow },
    { t: '复盘', s: '反馈、评估、改进', icon: faRepeat, c: C.purple }
  ];
  stack.forEach((st, i) => card(personalStack, 0.66 + i * 2.52, 2.18, 2.25, 3.55, { fill: C.panel, line: st.c, icon: st.icon, iconColor: st.c, title: st.t, titleSize: 16, body: st.s, bodySize: 11.1 }));
  callout(personalStack, 'Prompt 是一次交互；工作系统才是可以反复产生结果的个人能力资产。', 1.0, 6.15, 11.3, C.purple, faLayerGroup);
  addFooter(personalStack);

  const orgExperience = pptx.addSlide();
  addHeader(orgExperience, '企业最大的 AI 机会：复制优秀员工经验', '从个人高手，走向可共享、可评估、可持续优化的组织能力', 'Module 2 · Part 2 · 团队 AI Workflow', C.cyan);
  card(orgExperience, 0.72, 2.0, 4.75, 3.8, { fill: C.panel3, line: C.red, icon: faUserTie, iconColor: C.red, title: '传统：经验隐藏在人脑', titleSize: 17, body: '优秀员工知道怎么做\n但经验难描述、难复制\n人员流动后能力消失', bodySize: 12 });
  arrow(orgExperience, 5.75, 3.55, 0.7, C.cyan);
  card(orgExperience, 6.72, 2.0, 5.9, 3.8, { fill: C.panel2, line: C.cyan, icon: faNetworkWired, iconColor: C.cyan, title: 'AI 时代：经验变成组织资产', titleSize: 17, body: '知识文档化\n方法 Skill 化\n步骤 Workflow 化\n结果可评估、可迭代', bodySize: 12 });
  callout(orgExperience, '个人提效的上限，是个人；团队复用的价值，才会进入企业资产负债表。', 1.05, 6.15, 11.2, C.cyan, faBuilding);
  addFooter(orgExperience);

  caseSlide({ title: '案例：企业知识助手', subtitle: '让组织拥有随时可用的共同记忆', background: '制度、SOP、产品资料和历史经验分散在文档与人脑。', traditional: '员工到处找文件、问同事\n答案不一致，来源不清楚\n新人学习成本高。', aiFlow: '员工用自然语言提问\n→ 检索企业知识库\n→ 基于相关资料生成回答\n→ 引用来源与版本\n→ 无法回答时转人工', value: '减少重复咨询\n缩短新人上手时间\n提高制度执行一致性', human: 'AI：检索与解释\n知识 Owner：更新、审核、治理', accent: C.cyan });
  workflowSlide('从 Prompt 到 Workflow：让 AI 稳定工作', '把一次性技巧变成团队可以复用的流程', [
    { title: '获取客户信息', icon: faDatabase, color: C.blue },
    { title: '分析客户', icon: faMagnifyingGlassChart, color: C.cyan },
    { title: '生成策略', icon: faLightbulb, color: C.purple },
    { title: '生成邮件', icon: faEnvelope, color: C.orange },
    { title: '人工审核', icon: faUserCheck, color: C.green },
    { title: '更新 CRM', icon: faBriefcase, color: C.blue }
  ], C.cyan);

  const anatomy = pptx.addSlide();
  addHeader(anatomy, '一个可靠 AI Workflow 由什么组成？', '不是一串 Prompt，而是一套可运行、可检查的工作设计', 'Module 2 · Part 2 · 团队 AI Workflow', C.cyan);
  const parts = [
    { t: '目标', b: '要完成什么业务结果', icon: faBullseye, c: C.blue },
    { t: '知识', b: '需要哪些资料与规则', icon: faBookOpen, c: C.cyan },
    { t: '步骤', b: '按什么顺序推进', icon: faListCheck, c: C.purple },
    { t: '工具', b: '连接哪些系统', icon: faWrench, c: C.orange },
    { t: '控制', b: '哪里必须人工确认', icon: faShieldHalved, c: C.green }
  ];
  parts.forEach((p, i) => card(anatomy, 0.62 + i * 2.55, 2.1, 2.3, 3.65, { fill: C.panel, line: p.c, icon: p.icon, iconColor: p.c, title: p.t, titleSize: 16, body: p.b, bodySize: 11.1 }));
  callout(anatomy, 'Workflow 的价值：把“偶尔做对”变成“团队稳定做对”。', 1.55, 6.2, 10.2, C.cyan, faCheckDouble);
  addFooter(anatomy);

  workflowSlide('案例：销售机会管理 AI 流程', '让销售把时间投入客户关系与成交，而不是资料整理', [
    { title: '线索进入', icon: faBullseye, color: C.blue },
    { title: 'AI 分析', icon: faMagnifyingGlassChart, color: C.cyan },
    { title: '机会评分', icon: faRankingStar, color: C.purple },
    { title: '推荐行动', icon: faLightbulb, color: C.orange },
    { title: '销售确认', icon: faUserCheck, color: C.green },
    { title: 'CRM 更新', icon: faBriefcase, color: C.blue }
  ], C.cyan);
  caseSlide({ title: '案例：客服 AI 协作', subtitle: '让新人也能快速达到稳定服务水平', background: '客服每天处理大量重复问题与复杂例外。', traditional: '人工理解问题 → 找知识 → 写回复\n新人慢，答案质量依赖经验。', aiFlow: '客户问题进入\n→ AI 识别意图与情绪\n→ 检索知识与历史案例\n→ 推荐回答与下一步\n→ 客服确认或升级', value: '响应更快\n答案更一致\n培训成本更低', human: 'AI：标准问题与信息准备\n人：情绪、例外、关系与承诺', accent: C.cyan });
  caseSlide({ title: '案例：HR 招聘助手', subtitle: '减少低价值筛选，让 HR 聚焦人才判断', background: '岗位多、简历多，招聘团队重复阅读与协调。', traditional: '整理 JD → 人工筛简历 → 写评价 → 约面试\n耗时高且标准不一致。', aiFlow: '理解岗位需求\n→ 结构化候选人信息\n→ 按标准生成匹配分析\n→ 准备面试问题\n→ 协调日程与跟踪反馈', value: '提高筛选效率\n统一评价维度\n缩短招聘周期', human: 'AI：整理与辅助评估\n人：公平性、潜力与录用决策', accent: C.purple });

  const teamMaturity = pptx.addSlide();
  addHeader(teamMaturity, '团队 AI 成熟度：从零散使用到流程嵌入', 'AI 能力只有进入流程，才真正成为组织能力', 'Module 2 · Part 2 · 团队 AI Workflow', C.cyan);
  const tm = [
    { y: 5.45, t: '个人使用', s: '各自探索工具', c: C.dim },
    { y: 4.68, t: '团队共享', s: '模板、案例与知识', c: C.blue },
    { y: 3.91, t: 'Workflow', s: '稳定执行与人工控制', c: C.purple },
    { y: 3.14, t: '流程嵌入', s: '系统连接与业务指标', c: C.cyan },
    { y: 2.37, t: '持续优化', s: '评估、复盘与规模复制', c: C.green }
  ];
  tm.forEach((m, i) => { const w = 6.6 + i * 0.85; const x = (13.333 - w) / 2; slideShapeRound(teamMaturity, x, m.y, w, 0.52, m.c, 64, `${m.t}    ·    ${m.s}`); });
  callout(teamMaturity, '团队级 AI 的标志：不依赖某个高手，也能按照同样方法产生可预测结果。', 1.1, 6.3, 11.1, C.cyan, faPeopleGroup);
  addFooter(teamMaturity);

  const whyAgent = pptx.addSlide();
  addHeader(whyAgent, '为什么企业需要 Agent？', '很多业务问题不是“不会回答”，而是“需要持续完成”', 'Module 2 · Part 3 · AI Agent 数字员工', C.purple);
  card(whyAgent, 0.75, 2.0, 5.3, 3.8, { fill: C.panel3, line: C.dim, icon: faComments, iconColor: C.muted, title: 'Chatbot：一次交互', titleSize: 17, body: '人提出问题\nAI 生成回答\n任务在对话结束时停止', bodySize: 13 });
  arrow(whyAgent, 6.35, 3.55, 0.6, C.purple);
  card(whyAgent, 7.25, 2.0, 5.35, 3.8, { fill: C.panel2, line: C.purple, icon: faRobot, iconColor: C.purple, title: 'Agent：持续任务', titleSize: 17, body: '围绕业务目标\n持续观察环境\n主动行动、检查与调整\n直到完成或请求人工', bodySize: 13 });
  callout(whyAgent, 'Chatbot 解决一个问题；Agent 持续推动一个业务结果。', 2.15, 6.1, 9.05, C.purple, faRobot);
  addFooter(whyAgent);

  const agentDef = pptx.addSlide();
  addHeader(agentDef, 'AI Agent：拥有目标、工具和执行能力的工作单元', '不是一种单独模型，而是完整工作系统', 'Module 2 · Part 3 · AI Agent 数字员工', C.purple);
  const agentParts = [
    { x: 0.75, y: 2.0, t: '目标', icon: faBullseye, c: C.blue },
    { x: 0.75, y: 3.4, t: '知识', icon: faBookOpen, c: C.cyan },
    { x: 0.75, y: 4.8, t: '工具', icon: faWrench, c: C.green },
    { x: 10.25, y: 2.0, t: '记忆', icon: faMemory, c: C.purple },
    { x: 10.25, y: 3.4, t: '流程', icon: faDiagramProject, c: C.orange },
    { x: 10.25, y: 4.8, t: '控制', icon: faShieldHalved, c: C.yellow }
  ];
  agentParts.forEach(p => card(agentDef, p.x, p.y, 2.3, 0.92, { fill: C.panel, line: p.c, icon: p.icon, iconColor: p.c, title: p.t, titleSize: 13.2, shadow: false }));
  card(agentDef, 4.15, 2.05, 5.05, 3.72, { fill: C.panel2, line: C.purple, icon: faRobot, iconColor: C.purple, title: 'Agent Runtime', titleSize: 20, body: '理解任务\n规划下一步\n调用工具执行\n观察环境变化\n检查是否达标\n继续、求助或结束', bodySize: 12.3 });
  callout(agentDef, 'Agent = 目标驱动的执行循环 + 可调用能力 + 可控边界。', 2.0, 6.18, 9.35, C.purple, faCircleNodes);
  addFooter(agentDef);

  const agentFit = pptx.addSlide();
  addHeader(agentFit, '什么任务适合交给 Agent？', '高价值不等于高自主，先判断任务特征', 'Module 2 · Part 3 · AI Agent 数字员工', C.purple);
  const fit = [
    { x: 0.8, y: 2.1, t: '高频', b: '反复发生', icon: faRepeat, c: C.blue },
    { x: 3.2, y: 2.1, t: '信息密集', b: '需要处理大量资料', icon: faDatabase, c: C.cyan },
    { x: 5.6, y: 2.1, t: '规则存在', b: '有经验与约束', icon: faListCheck, c: C.purple },
    { x: 8.0, y: 2.1, t: '结果可验证', b: '能判断是否完成', icon: faCircleCheck, c: C.green },
    { x: 10.4, y: 2.1, t: '持续跟踪', b: '需要长期运行', icon: faClock, c: C.orange }
  ];
  fit.forEach(f => card(agentFit, f.x, f.y, 2.12, 2.15, { fill: C.panel, line: f.c, icon: f.icon, iconColor: f.c, title: f.t, titleSize: 14.5, body: f.b, bodySize: 10 }));
  card(agentFit, 0.8, 4.75, 5.55, 1.15, { fill: C.panel2, line: C.green, icon: faCircleCheck, iconColor: C.green, title: '更适合', body: '客户运营、销售跟进、采购分析、财务监测、IT 运维', bodySize: 10.1, titleSize: 12.5, shadow: false });
  card(agentFit, 6.7, 4.75, 5.55, 1.15, { fill: '2A1B1A', line: C.orange, icon: faTriangleExclamation, iconColor: C.orange, title: '谨慎交付', body: '战略决策、高风险审批、复杂伦理判断、不可逆动作', bodySize: 10.1, titleSize: 12.5, shadow: false });
  addFooter(agentFit);

  caseSlide({ title: '主案例：重点客户订单延期风险 Agent', subtitle: '同一个场景，贯穿“监测—分析—决策—行动—跟踪”', background: 'B2B 企业拥有大量重点客户与跨部门订单。', traditional: '销售每周人工查看 CRM、订单、邮件和投诉\n风险往往在客户不满后才被发现。', aiFlow: '每天自动扫描客户与订单状态\n→ 识别延期、投诉、互动下降等信号\n→ 检查合同、SOP 与历史案例\n→ 生成风险等级与原因\n→ 提出行动建议并请求负责人确认\n→ 更新系统、持续跟踪', value: '风险更早发现\n减少客户流失\n提升销售与交付协同', human: 'AI：监测、整理、提醒\n人：商业判断、客户关系、最终行动', accent: C.purple });
  workflowSlide('重点客户订单延期风险 Agent：执行链路', '一个目标，跨越多个系统与人类控制点', [
    { title: '读取 CRM', icon: faBriefcase, color: C.blue },
    { title: '检查订单', icon: faCartShopping, color: C.cyan },
    { title: '发现风险', icon: faTriangleExclamation, color: C.orange },
    { title: '检索依据', icon: faBookOpen, color: C.purple },
    { title: '生成方案', icon: faLightbulb, color: C.cyan },
    { title: '人工确认', icon: faUserCheck, color: C.green },
    { title: '执行跟进', icon: faArrowTrendUp, color: C.blue }
  ], C.purple);

  const hitl = pptx.addSlide();
  addHeader(hitl, '客户风险 Agent 的人机分工', 'AI 负责感知与准备，人负责判断与关系', 'Module 2 · Part 3 · AI Agent 数字员工', C.orange);
  card(hitl, 0.75, 1.95, 4.65, 3.92, { fill: C.panel2, line: C.purple, icon: faRobot, iconColor: C.purple, title: 'AI Agent 负责', titleSize: 18, body: '跨系统扫描数据\n发现异常与风险信号\n聚合证据与历史信息\n生成风险等级与方案\n持续提醒与跟踪状态', bodySize: 12 });
  card(hitl, 7.93, 1.95, 4.65, 3.92, { fill: C.panel2, line: C.green, icon: faUserTie, iconColor: C.green, title: '业务负责人负责', titleSize: 18, body: '判断客户真实影响\n选择商业应对策略\n处理关系与情绪\n批准关键动作\n承担最终责任', bodySize: 12 });
  slideShapeRound(hitl, 5.65, 2.72, 1.75, 2.0, C.orange, 75, 'HITL\n\n关键节点\n人工把关');
  callout(hitl, 'Human-in-the-loop 不是妨碍自动化，而是让自动化能够进入高价值业务。', 1.05, 6.17, 11.2, C.orange, faShieldHalved);
  addFooter(hitl);

  caseSlide({ title: '案例：采购 Agent', subtitle: '从被动下单，到主动预测与建议', background: '库存变化快，采购需要同时关注需求、供应商、价格与交付。', traditional: '人工查库存 → 问业务 → 比价 → 写申请\n信息滞后，容易过量或短缺。', aiFlow: '监测库存与销售趋势\n→ 预测未来需求\n→ 比较供应商价格、交期与风险\n→ 生成采购建议\n→ 人工审批\n→ 创建订单并跟踪', value: '降低缺货与积压\n提高采购响应速度\n让采购聚焦谈判与供应关系', human: 'AI：预测、比较、准备\n人：预算、谈判、审批与责任', accent: C.green });
  caseSlide({ title: '案例：财务经营分析 Agent', subtitle: '从月末报表，到持续发现经营异常', background: '财务与经营团队定期拉取多系统数据。', traditional: '取数 → 清洗 → 做表 → 查异常 → 写报告\n大量时间耗费在重复准备。', aiFlow: '定时读取经营数据\n→ 核对口径与完整性\n→ 识别收入、成本、现金与毛利异常\n→ 分析驱动因素\n→ 生成管理简报并提醒负责人', value: '更快发现经营问题\n缩短报告周期\n形成持续监测机制', human: 'AI：取数、检查、分析\n人：解释业务原因、决策与合规', accent: C.green });

  const rpa = pptx.addSlide();
  addHeader(rpa, 'Agent 不是 RPA 升级版', '固定自动化与目标驱动执行，适用于不同任务', 'Module 2 · Part 3 · AI Agent 数字员工', C.purple);
  card(rpa, 0.75, 1.92, 5.55, 4.05, { fill: C.panel3, line: C.dim, icon: faRepeat, iconColor: C.muted, title: '传统自动化 / RPA', titleSize: 17, body: '逻辑：固定规则\n流程：预先配置\n输入：结构化、稳定\n变化：遇到例外容易中断\n适合：重复、确定的操作', bodySize: 12 });
  card(rpa, 7.05, 1.92, 5.55, 4.05, { fill: C.panel2, line: C.purple, icon: faRobot, iconColor: C.purple, title: 'AI Agent', titleSize: 17, body: '逻辑：理解与推理\n流程：动态规划\n输入：复杂、多来源\n变化：观察反馈并调整\n适合：目标明确但路径可变的任务', bodySize: 12 });
  callout(rpa, '最好的企业方案通常不是二选一，而是“规则自动化 + Agent 判断 + 人工控制”的组合。', 0.95, 6.25, 11.4, C.purple, faPuzzlePiece);
  addFooter(rpa);

  const prereq = pptx.addSlide();
  addHeader(prereq, '企业落地 Agent，需要五个基础', '买一个 Agent 平台，不等于拥有一个可工作的数字员工', 'Module 2 · Part 3 · AI Agent 数字员工', C.purple);
  const ps = [
    { t: '业务目标', b: '要改善什么结果', icon: faBullseye, c: C.blue },
    { t: '可用数据', b: '事实从哪里来', icon: faDatabase, c: C.cyan },
    { t: '系统连接', b: '如何执行动作', icon: faLink, c: C.green },
    { t: '流程规则', b: '如何稳定做事', icon: faDiagramProject, c: C.purple },
    { t: '控制机制', b: '如何安全负责', icon: faShieldHalved, c: C.orange }
  ];
  ps.forEach((p, i) => card(prereq, 0.62 + i * 2.55, 2.1, 2.3, 3.7, { fill: C.panel, line: p.c, icon: p.icon, iconColor: p.c, title: p.t, titleSize: 15.5, body: p.b, bodySize: 10.8 }));
  callout(prereq, 'Agent 项目本质上是业务流程重新设计项目，而不是单纯模型项目。', 1.4, 6.2, 10.5, C.purple, faSitemap);
  addFooter(prereq);

  const native = pptx.addSlide();
  addHeader(native, 'AI Native：不是给企业加 AI，而是重新设计企业', '从“软件提供功能”到“系统帮助完成目标”', 'Module 2 · Part 4 · AI Native 企业', C.green);
  const stages = [
    { x: 0.72, t: '传统企业', b: '人执行工作\n流程协调人\n软件记录结果', icon: faPeopleGroup, c: C.dim },
    { x: 4.55, t: '数字化企业', b: '系统承载流程\n数据支持管理\n效率持续提升', icon: faServer, c: C.blue },
    { x: 8.38, t: 'AI Native 企业', b: '目标驱动协作\nAI 参与理解与执行\n人负责判断与创新', icon: faBuilding, c: C.green }
  ];
  stages.forEach((s, i) => { card(native, s.x, 2.05, 3.35, 3.75, { fill: C.panel, line: s.c, icon: s.icon, iconColor: s.c, title: s.t, titleSize: 17, body: s.b, bodySize: 12 }); if (i < 2) arrow(native, s.x + 3.48, 3.6, 0.42, s.c); });
  callout(native, 'AI Native 的判断标准：从设计之初，就假设 AI 会参与完成工作。', 1.35, 6.2, 10.65, C.green, faBuilding);
  addFooter(native);

  const product = pptx.addSlide();
  addHeader(product, 'AI Native 产品：从提供功能，到提供结果', '用户不再学习每个按钮，而是表达目标', 'Module 2 · Part 4 · AI Native 企业', C.green);
  card(product, 0.75, 1.95, 5.35, 3.95, { fill: C.panel3, line: C.dim, icon: faWrench, iconColor: C.muted, title: '传统软件', titleSize: 17, body: '提供功能与界面\n用户理解流程\n用户逐步操作\n用户负责组合成结果', bodySize: 12.4 });
  arrow(product, 6.36, 3.55, 0.55, C.green);
  card(product, 7.2, 1.95, 5.4, 3.95, { fill: C.panel2, line: C.green, icon: faBullseye, iconColor: C.green, title: 'AI Native 产品', titleSize: 17, body: '用户表达目标\nAI 理解上下文\n主动规划并执行\n用户审核关键变化与结果', bodySize: 12.4 });
  callout(product, '产品范式变化：功能 → 操作 → 结果，升级为目标 → AI 协作 → 结果。', 1.2, 6.18, 10.9, C.green, faArrowRightArrowLeft);
  addFooter(product);

  caseSlide({ title: '案例：AI Native CRM', subtitle: '从记录系统，变成销售决策与行动系统', background: '传统 CRM 存放客户、联系人与跟进记录。', traditional: '销售录入信息\n管理者看报表\n系统很少主动帮助成交。', aiFlow: '持续分析客户互动、机会阶段与历史数据\n→ 主动发现高价值机会与风险\n→ 解释原因\n→ 推荐下一步行动\n→ 自动准备客户画像、议程与邮件', value: '销售更聚焦高价值客户\n管理更早发现风险\n系统从“记录”走向“参与”', human: 'AI：洞察、建议、准备\n销售：关系、判断、承诺与成交', accent: C.green });

  const sidekick = pptx.addSlide();
  addHeader(sidekick, '官方实践：Shopify Sidekick 展现了 AI Native 产品形态', '在业务上下文中理解、生成、执行，并在变更前让人确认', 'Module 2 · Part 4 · AI Native 企业', C.green);
  const sk = [
    { t: '在商店上下文中工作', icon: faStore, c: C.blue },
    { t: '生成内容与分析数据', icon: faPenNib, c: C.cyan },
    { t: '管理订单与编辑商品', icon: faCartShopping, c: C.purple },
    { t: '长任务后台运行', icon: faClock, c: C.orange },
    { t: '记忆与保存 Skills', icon: faMemory, c: C.green },
    { t: '变更前提交审核', icon: faUserCheck, c: C.yellow }
  ];
  sk.forEach((s, i) => card(sidekick, 0.58 + (i % 3) * 4.18, 1.85 + Math.floor(i / 3) * 2.15, 3.78, 1.72, { fill: C.panel, line: s.c, icon: s.icon, iconColor: s.c, title: s.t, titleSize: 13.2, shadow: false }));
  callout(sidekick, '泛化启示：AI Native 不只是聊天入口，而是“业务上下文 + 工具 + 长任务 + 记忆 + 人工确认”。', 0.85, 6.25, 11.6, C.green, faLightbulb);
  addFooter(sidekick, '', 'Source: Shopify Help Center — Sidekick（官方产品说明）');

  const service = pptx.addSlide();
  addHeader(service, 'AI Native 服务：从回答问题，到解决问题', '客服是企业 Agent 最成熟、最容易量化价值的场景之一', 'Module 2 · Part 4 · AI Native 企业', C.cyan);
  const svc = [
    { x: 0.5, t: '客户提出需求', icon: faUser, c: C.blue },
    { x: 2.85, t: 'AI 理解意图', icon: faBrain, c: C.cyan },
    { x: 5.2, t: '查询账户与订单', icon: faDatabase, c: C.purple },
    { x: 7.55, t: '执行退款 / 修改', icon: faWrench, c: C.orange },
    { x: 9.9, t: '解决或升级人工', icon: faHeadset, c: C.green }
  ];
  svc.forEach((s, i) => { card(service, s.x, 2.35, 2.0, 2.15, { fill: C.panel, line: s.c, icon: s.icon, iconColor: s.c, title: s.t, titleSize: 12.4, shadow: false }); if (i < svc.length - 1) arrow(service, s.x + 2.05, 3.28, 0.25, s.c); });
  card(service, 0.75, 5.05, 5.7, 1.0, { fill: C.panel2, line: C.cyan, icon: faArrowTrendUp, iconColor: C.cyan, title: '官方案例共性', body: '从静态问答转向“理解—调用系统—完成解决—必要时转人”。', bodySize: 9.8, titleSize: 11.5, shadow: false });
  card(service, 6.8, 5.05, 5.75, 1.0, { fill: C.panel2, line: C.green, icon: faShieldHalved, iconColor: C.green, title: '企业落地关键', body: '以“解决率、重复咨询、处理时长、客户满意度”衡量，而不是只看回复质量。', bodySize: 9.8, titleSize: 11.5, shadow: false });
  addFooter(service, '', 'Sources: OpenAI — Zendesk / Klarna customer stories（官方）');

  const realCases = pptx.addSlide();
  addHeader(realCases, '来自官方企业案例的四个信号', '企业 AI 正在从个人提效，走向 Agent 与业务流程重构', 'Module 2 · 实践观察', C.cyan);
  const cases = [
    { x: 0.65, t: 'Zapier', metric: '89% 员工采用\n800+ 内部 Agent', lesson: '让非技术员工成为构建者', c: C.orange, icon: faPuzzlePiece },
    { x: 3.8, t: 'Capita', metric: '每月节省 9,000 小时', lesson: '个人 Copilot 走向 Agent 服务', c: C.blue, icon: faClock },
    { x: 6.95, t: 'Klarna', metric: '客服处理从 11 分钟降至 <2 分钟', lesson: '以解决率而不是回复量衡量', c: C.purple, icon: faHeadset },
    { x: 10.1, t: 'SCSK', metric: '从 10+ 场景筛选 3 个 Agent', lesson: '先识别场景，再快速试点', c: C.green, icon: faFilterCircleDollar }
  ];
  cases.forEach(c => card(realCases, c.x, 1.95, 2.6, 4.15, { fill: C.panel, line: c.c, icon: c.icon, iconColor: c.c, title: c.t, titleSize: 17, body: `${c.metric}\n\n${c.lesson}`, bodySize: 11.2 }));
  addFooter(realCases, '', 'Sources: Anthropic Zapier; Microsoft Capita / SCSK; OpenAI Klarna（官方客户案例）');

  const roadmap = pptx.addSlide();
  addHeader(roadmap, '企业 AI 应用进化路线', '不是一步进入 AI Native，而是从个人、团队、流程逐步升级', 'Module 2 · 总结', C.cyan);
  const r = [
    { x: 0.75, t: '阶段 1', h: '个人提效', s: 'AI Assistant', c: C.blue, icon: faUser },
    { x: 3.75, t: '阶段 2', h: '团队复用', s: 'AI Workflow', c: C.cyan, icon: faPeopleGroup },
    { x: 6.75, t: '阶段 3', h: '业务执行', s: 'AI Agent', c: C.purple, icon: faRobot },
    { x: 9.75, t: '阶段 4', h: '企业重构', s: 'AI Native', c: C.green, icon: faBuilding }
  ];
  r.forEach((a, i) => { card(roadmap, a.x, 2.1 + (3 - i) * 0.35, 2.7, 3.2 + i * 0.35, { fill: C.panel, line: a.c, icon: a.icon, iconColor: a.c, title: `${a.t} · ${a.h}`, titleSize: 15.5, body: a.s, bodySize: 12 }); if (i < 3) arrow(roadmap, a.x + 2.78, 3.7, 0.2, a.c); });
  callout(roadmap, '第一代 AI 帮助人完成任务；下一代 AI 帮助组织完成目标。', 1.35, 6.25, 10.65, C.cyan, faRocket);
  addFooter(roadmap);
}

function frameworkSlide(title, subtitle, items, accent = C.purple, module = 'Module 3 · 落地 AI：从机会发现到业务成果') {
  const slide = pptx.addSlide();
  addHeader(slide, title, subtitle, module, accent);
  items.forEach((it, i) => {
    const cols = items.length <= 4 ? items.length : 3;
    const row = Math.floor(i / cols);
    const col = i % cols;
    const w = cols === 4 ? 2.85 : 3.75;
    const gap = cols === 4 ? 0.25 : 0.35;
    const x = 0.65 + col * (w + gap);
    const y = 1.9 + row * 2.15;
    card(slide, x, y, w, 1.75, { fill: C.panel, line: it.color || accent, icon: it.icon, iconColor: it.color || accent, title: it.title, titleSize: 14.2, body: it.body, bodySize: 10.2, shadow: false });
  });
  addFooter(slide);
}

function module3Slides() {
  sectionSlide(3, '落地 AI：从机会发现到业务成果', '不是做更多 Demo，而是找到第一个值得投入、能够证明价值的 AI 项目。', ['发现机会', '选择场景', '设计方案', '验证价值'], C.purple);
  const fail = pptx.addSlide();
  addHeader(fail, '为什么很多 AI 项目失败？', '通常不是模型不够强，而是起点与落地方式错误', 'Module 3 · Part 1 · 发现机会', C.orange);
  const fs = [
    { x: 0.72, t: '从技术开始', b: '先买平台\n再找场景', icon: faWrench, c: C.red },
    { x: 3.78, t: '从 Demo 开始', b: '证明能做\n没有业务指标', icon: faFlask, c: C.orange },
    { x: 6.84, t: '没有业务 Owner', b: 'IT 推动\n业务不负责结果', icon: faUserTie, c: C.purple },
    { x: 9.9, t: '流程没有改变', b: 'AI 只是外挂\n用户仍按旧方式工作', icon: faDiagramProject, c: C.blue }
  ];
  fs.forEach(f => card(fail, f.x, 2.05, 2.72, 3.8, { fill: C.panel, line: f.c, icon: f.icon, iconColor: f.c, title: f.t, titleSize: 16.5, body: f.b, bodySize: 11.7 }));
  callout(fail, 'AI 项目首先是业务项目，其次才是技术项目。', 2.15, 6.25, 9.05, C.orange, faTriangleExclamation);
  addFooter(fail);

  const jobTask = pptx.addSlide();
  addHeader(jobTask, '从工作任务，而不是岗位寻找 AI 机会', '岗位太宽，任务才是可以重新设计的基本单元', 'Module 3 · Part 1 · 发现机会', C.purple);
  card(jobTask, 0.72, 2.0, 2.4, 3.9, { fill: C.panel2, line: C.purple, icon: faUserTie, iconColor: C.purple, title: '销售经理', titleSize: 18, body: '一个岗位\n包含许多不同任务', bodySize: 12.5 });
  arrow(jobTask, 3.35, 3.63, 0.55, C.purple);
  const tasks = [
    { x: 4.05, y: 1.85, t: '找客户', icon: faBullseye, c: C.blue },
    { x: 6.15, y: 1.85, t: '研究客户', icon: faMagnifyingGlassChart, c: C.cyan },
    { x: 8.25, y: 1.85, t: '准备方案', icon: faPenNib, c: C.purple },
    { x: 4.05, y: 4.0, t: '跟进沟通', icon: faEnvelope, c: C.orange },
    { x: 6.15, y: 4.0, t: '预测成交', icon: faChartLine, c: C.green },
    { x: 8.25, y: 4.0, t: '管理风险', icon: faShieldHalved, c: C.yellow }
  ];
  tasks.forEach(t => card(jobTask, t.x, t.y, 1.85, 1.55, { fill: C.panel, line: t.c, icon: t.icon, iconColor: t.c, title: t.t, titleSize: 12.5, shadow: false }));
  card(jobTask, 10.45, 2.0, 2.15, 3.9, { fill: C.panel3, line: C.green, icon: faLightbulb, iconColor: C.green, title: 'AI 机会', titleSize: 16, body: '哪些任务高频？\n哪些信息密集？\n哪些结果可验证？', bodySize: 11.1 });
  addFooter(jobTask);

  const canvas = pptx.addSlide();
  addHeader(canvas, 'AI Opportunity Canvas：把模糊想法变成业务机会', '六个问题，完成从痛点到价值的第一次结构化', 'Module 3 · Part 1 · 发现机会', C.purple);
  const cvs = [
    { x: 0.7, y: 1.82, t: '1 业务目标', b: '希望改善什么业务结果？', icon: faBullseye, c: C.blue },
    { x: 4.55, y: 1.82, t: '2 当前流程', b: '现在由谁、如何完成？', icon: faDiagramProject, c: C.cyan },
    { x: 8.4, y: 1.82, t: '3 最大痛点', b: '哪里最耗时、易错或体验差？', icon: faTriangleExclamation, c: C.orange },
    { x: 0.7, y: 4.05, t: '4 AI 机会', b: 'AI 可以理解、分析、生成或执行什么？', icon: faBrain, c: C.purple },
    { x: 4.55, y: 4.05, t: '5 人机分工', b: 'AI 做什么？人保留什么？', icon: faHandshake, c: C.green },
    { x: 8.4, y: 4.05, t: '6 成功指标', b: '如何证明效率、质量或业务价值？', icon: faGaugeHigh, c: C.yellow }
  ];
  cvs.forEach(c => card(canvas, c.x, c.y, 3.35, 1.75, { fill: C.panel, line: c.c, icon: c.icon, iconColor: c.c, title: c.t, titleSize: 14.5, body: c.b, bodySize: 10.4, shadow: false }));
  addFooter(canvas);

  const oppExample = pptx.addSlide();
  addHeader(oppExample, '机会卡示例：客户投诉处理 AI 助手', '从一个真实痛点，形成可讨论、可评估的机会定义', 'Module 3 · Part 1 · 发现机会', C.purple);
  const ex = [
    { x: 0.7, y: 1.85, t: '业务目标', b: '缩短响应时间\n提升客户满意度', c: C.blue, icon: faBullseye },
    { x: 3.85, y: 1.85, t: '当前痛点', b: '查询慢\n新人经验不足\n回复质量不一致', c: C.orange, icon: faTriangleExclamation },
    { x: 7.0, y: 1.85, t: 'AI 机会', b: '理解投诉\n检索历史案例\n推荐方案与回复', c: C.purple, icon: faBrain },
    { x: 10.15, y: 1.85, t: '成功指标', b: '平均响应时间\n一次解决率\n满意度', c: C.green, icon: faGaugeHigh },
    { x: 0.7, y: 4.2, t: 'AI 负责', b: '分类、检索、整理、起草', c: C.cyan, icon: faRobot },
    { x: 4.0, y: 4.2, t: '人负责', b: '赔偿判断、关系与例外', c: C.green, icon: faUserTie },
    { x: 7.3, y: 4.2, t: '所需数据', b: '工单、客户记录、制度与历史案例', c: C.blue, icon: faDatabase },
    { x: 10.0, y: 4.2, t: '下一步', b: '抽取 100 个历史工单做离线验证', c: C.yellow, icon: faFlask }
  ];
  ex.forEach(e => card(oppExample, e.x, e.y, e.x >= 10 ? 2.45 : 2.85, 1.65, { fill: C.panel, line: e.c, icon: e.icon, iconColor: e.c, title: e.t, titleSize: 12.8, body: e.b, bodySize: 9.5, shadow: false }));
  addFooter(oppExample);

  frameworkSlide('什么工作最值得 AI 化？', '用五个维度判断任务潜力，而不是凭感觉追热点', [
    { title: '频率', body: '发生次数越多，累计价值越高', icon: faRepeat, color: C.blue },
    { title: '信息密集度', body: '需要处理大量文档、数据与沟通', icon: faDatabase, color: C.cyan },
    { title: '规则与经验', body: '已有 SOP、标准或可学习案例', icon: faListCheck, color: C.purple },
    { title: '可验证性', body: '能够判断结果是否正确或完成', icon: faCircleCheck, color: C.green },
    { title: '持续性', body: '需要长期监测、跟踪或响应', icon: faClock, color: C.orange },
    { title: '风险可控', body: '可以设置权限、确认、回滚与升级', icon: faShieldHalved, color: C.yellow }
  ], C.purple);

  const map = pptx.addSlide();
  addHeader(map, '企业 AI 机会地图：从个人到企业', '帮助各部门同时看到短期效率与长期重构机会', 'Module 3 · Part 1 · 发现机会', C.purple);
  const maps = [
    { x: 0.72, t: '个人效率', q: '我如何更高效？', examples: '研究、写作、分析、会议', c: C.blue, icon: faUser },
    { x: 3.78, t: '团队协作', q: '团队如何复制？', examples: '知识助手、模板、Workflow', c: C.cyan, icon: faPeopleGroup },
    { x: 6.84, t: '业务流程', q: '任务如何自动完成？', examples: '销售、采购、客服、财务 Agent', c: C.purple, icon: faDiagramProject },
    { x: 9.9, t: '企业创新', q: '业务如何重新设计？', examples: 'AI Native 产品、服务与组织', c: C.green, icon: faBuilding }
  ];
  maps.forEach(m => card(map, m.x, 2.05, 2.72, 3.85, { fill: C.panel, line: m.c, icon: m.icon, iconColor: m.c, title: m.t, titleSize: 16, body: `${m.q}\n\n${m.examples}`, bodySize: 10.7 }));
  addFooter(map);

  const funnel = pptx.addSlide();
  addHeader(funnel, '不是所有 AI 机会都值得投入', '从大量想法中，筛选 1–2 个能够证明价值的试点', 'Module 3 · Part 2 · 选择场景', C.orange);
  const fs2 = [
    { y: 1.92, w: 11.4, t: '100 个想法：来自访谈、痛点与案例启发', c: C.blue },
    { y: 2.88, w: 9.5, t: '20 个机会：完成 Opportunity Canvas', c: C.cyan },
    { y: 3.84, w: 7.6, t: '5 个重点：价值 × 可行性排序', c: C.purple },
    { y: 4.80, w: 5.7, t: '1–2 个 PoC：小范围验证业务价值', c: C.green }
  ];
  fs2.forEach(f => { const x = (13.333 - f.w) / 2; slide.addShape(pptx.ShapeType.chevron, { x, y: f.y, w: f.w, h: 0.72, fill: { color: f.c, transparency: 57 }, line: { color: f.c, width: 1 } }); slide.addText(f.t, { x: x + 0.35, y: f.y + 0.18, w: f.w - 0.7, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink' }); });
  callout(funnel, '企业不缺 AI 想法，缺的是选择第一个正确项目的纪律。', 1.65, 6.15, 10.05, C.orange, faFilterCircleDollar);
  addFooter(funnel);

  const matrix = pptx.addSlide();
  addHeader(matrix, 'AI 项目优先级矩阵', '业务价值 × 实施复杂度，决定先做什么', 'Module 3 · Part 2 · 选择场景', C.orange);
  slide.addShape(pptx.ShapeType.rect, { x: 2.1, y: 1.85, w: 8.9, h: 4.5, fill: { color: C.panel3, transparency: 10 }, line: { color: C.border, width: 1 } });
  slide.addShape(pptx.ShapeType.line, { x: 6.55, y: 1.85, w: 0, h: 4.5, line: { color: C.border, width: 1.2 } });
  slide.addShape(pptx.ShapeType.line, { x: 2.1, y: 4.1, w: 8.9, h: 0, line: { color: C.border, width: 1.2 } });
  slide.addText('业务价值 ↑', { x: 0.82, y: 3.55, w: 1.0, h: 0.3, fontFace: FONT, fontSize: 10, bold: true, color: C.muted, rotate: 270, margin: 0 });
  slide.addText('实施复杂度 →', { x: 5.65, y: 6.55, w: 2.1, h: 0.28, fontFace: FONT, fontSize: 10, bold: true, color: C.muted, margin: 0, align: 'center' });
  const qs = [
    { x: 2.35, y: 2.08, t: 'Quick Win', b: '高价值 · 低复杂\n优先启动', c: C.green, icon: faRocket },
    { x: 6.8, y: 2.08, t: 'Strategic', b: '高价值 · 高复杂\n战略规划', c: C.purple, icon: faFlagCheckered },
    { x: 2.35, y: 4.35, t: 'Nice to Have', b: '低价值 · 低复杂\n顺手优化', c: C.blue, icon: faPuzzlePiece },
    { x: 6.8, y: 4.35, t: 'Avoid / Later', b: '低价值 · 高复杂\n暂缓投入', c: C.red, icon: faCircleXmark }
  ];
  qs.forEach(q => card(matrix, q.x, q.y, 3.95, 1.55, { fill: C.panel, line: q.c, icon: q.icon, iconColor: q.c, title: q.t, titleSize: 14.5, body: q.b, bodySize: 10.2, shadow: false }));
  addFooter(matrix);

  frameworkSlide('第一个 AI 项目应该怎么选？', '不是最酷，而是最容易证明价值并获得组织信心', [
    { title: '真实业务问题', body: '业务负责人愿意为结果负责', icon: faBullseye, color: C.blue },
    { title: '数据已经存在', body: '不把第一个项目变成大型数据工程', icon: faDatabase, color: C.cyan },
    { title: '用户清晰', body: '明确谁使用、何时使用、为什么使用', icon: faUserCheck, color: C.green },
    { title: '结果可衡量', body: '效率、质量、收入、风险或体验', icon: faGaugeHigh, color: C.orange },
    { title: '风险可控制', body: '可以小范围、可回退、有人审', icon: faShieldHalved, color: C.purple },
    { title: '可复制', body: '成功后能够扩展到团队或相似流程', icon: faRepeat, color: C.yellow }
  ], C.orange);

  const system = pptx.addSlide();
  addHeader(system, 'AI 应用不是模型项目，而是业务系统项目', '模型只是一层，价值来自完整组合', 'Module 3 · Part 3 · 设计方案', C.purple);
  const sys = [
    { t: '模型', icon: faBrain, c: C.blue }, { t: '数据', icon: faDatabase, c: C.cyan },
    { t: '知识', icon: faBookOpen, c: C.green }, { t: '工具', icon: faWrench, c: C.orange },
    { t: '流程', icon: faDiagramProject, c: C.purple }, { t: '人机协同', icon: faHandshake, c: C.yellow }
  ];
  sys.forEach((s, i) => {
    const angle = (Math.PI * 2 * i) / sys.length - Math.PI / 2;
    const x = 6.35 + Math.cos(angle) * 3.7 - 0.75;
    const y = 3.75 + Math.sin(angle) * 2.0 - 0.55;
    card(system, x, y, 1.65, 1.05, { fill: C.panel, line: s.c, icon: s.icon, iconColor: s.c, title: s.t, titleSize: 11.5, shadow: false });
  });
  slideShapeRound(system, 4.9, 2.65, 3.55, 2.15, C.purple, 75, 'AI 应用\n\n业务目标驱动的\n完整工作系统');
  callout(system, '模型决定能力上限；数据、流程、治理与用户设计决定项目能否产生价值。', 1.05, 6.25, 11.2, C.purple, faLayerGroup);
  addFooter(system);

  const solution = pptx.addSlide();
  addHeader(solution, 'AI Solution Canvas：把机会变成可实施方案', '明确用户、输入、能力、工具、控制与成功标准', 'Module 3 · Part 3 · 设计方案', C.purple);
  const sc = [
    { x: 0.7, y: 1.82, t: '用户', b: '谁在什么场景使用？', c: C.blue, icon: faUser },
    { x: 3.75, y: 1.82, t: '目标', b: '希望完成什么结果？', c: C.cyan, icon: faBullseye },
    { x: 6.8, y: 1.82, t: '输入', b: '需要哪些数据、知识与上下文？', c: C.purple, icon: faDatabase },
    { x: 9.85, y: 1.82, t: 'AI 能力', b: '理解、分析、生成还是执行？', c: C.green, icon: faBrain },
    { x: 0.7, y: 4.15, t: '工具连接', b: '需要访问哪些系统？', c: C.orange, icon: faLink },
    { x: 3.75, y: 4.15, t: '人机控制', b: '何时确认、升级与回滚？', c: C.yellow, icon: faShieldHalved },
    { x: 6.8, y: 4.15, t: '输出', b: '回答、文件、动作还是状态变化？', c: C.cyan, icon: faFileCircleCheck },
    { x: 9.85, y: 4.15, t: '衡量', b: '如何判断成功？', c: C.green, icon: faGaugeHigh }
  ];
  sc.forEach(s => card(solution, s.x, s.y, 2.75, 1.75, { fill: C.panel, line: s.c, icon: s.icon, iconColor: s.c, title: s.t, titleSize: 13.8, body: s.b, bodySize: 9.8, shadow: false }));
  addFooter(solution);

  workflowSlideModule3('Agent 设计六步法', '从一个业务目标，设计出可执行、可控制的 Agent', [
    { title: '定义目标', icon: faBullseye, color: C.blue },
    { title: '拆解任务', icon: faListCheck, color: C.cyan },
    { title: '配置知识', icon: faBookOpen, color: C.green },
    { title: '连接工具', icon: faLink, color: C.orange },
    { title: '设计循环', icon: faArrowsRotate, color: C.purple },
    { title: '设置控制', icon: faShieldHalved, color: C.yellow }
  ], C.purple);

  frameworkSlide('数据与知识准备', 'AI 输出质量不会超过它能够访问的事实质量', [
    { title: '数据在哪里？', body: '数据库、Excel、SaaS、文档与消息', icon: faDatabase, color: C.blue },
    { title: '是否准确？', body: '字段口径、重复、缺失与历史错误', icon: faCircleCheck, color: C.green },
    { title: '是否可访问？', body: 'API、权限、网络与系统边界', icon: faLink, color: C.cyan },
    { title: '是否及时？', body: '静态资料还是实时状态', icon: faClock, color: C.orange },
    { title: '谁负责更新？', body: '明确知识 Owner 与版本机制', icon: faUserTie, color: C.purple },
    { title: '能否追溯？', body: '来源、证据、版本与审计记录', icon: faFingerprint, color: C.yellow }
  ], C.blue);

  frameworkSlide('系统连接与工具能力', 'Agent 要有“手脚”，但必须只获得完成任务所需的最小权限', [
    { title: 'API / MCP', body: '标准化连接工具与系统', icon: faLink, color: C.blue },
    { title: '数据库', body: '读取结构化业务事实', icon: faDatabase, color: C.cyan },
    { title: 'SaaS 应用', body: 'CRM、ERP、工单、协作平台', icon: faCloud, color: C.purple },
    { title: 'Computer Use', body: '在无 API 场景操作界面', icon: faLaptop, color: C.orange },
    { title: '文件系统', body: '读取、创建、修改与版本控制', icon: faFileLines, color: C.green },
    { title: '最小权限', body: '只开放必要范围与动作', icon: faLock, color: C.yellow }
  ], C.cyan);

  frameworkSlide('安全与治理设计', '能力与控制必须成对出现', [
    { title: '权限', body: '谁能看什么、做什么', icon: faLock, color: C.blue },
    { title: '人工确认', body: '关键节点必须由人批准', icon: faUserCheck, color: C.green },
    { title: '日志与审计', body: '记录过程、证据和责任', icon: faFileContract, color: C.cyan },
    { title: '失败与回退', body: '停止、重试、撤销和恢复', icon: faClockRotateLeft, color: C.orange },
    { title: '质量评估', body: '规则、模型与人工多层检查', icon: faClipboardCheck, color: C.purple },
    { title: '异常升级', body: '不确定时及时交给人', icon: faTriangleExclamation, color: C.yellow }
  ], C.orange);

  const demo = pptx.addSlide();
  addHeader(demo, '不要做 Demo，要做业务验证', '“AI 能做”与“业务值得做”是两个不同问题', 'Module 3 · Part 4 · 验证价值', C.green);
  card(demo, 0.75, 2.0, 5.5, 3.85, { fill: C.panel3, line: C.dim, icon: faFlask, iconColor: C.muted, title: 'Demo', titleSize: 18, body: '证明模型能够完成某个任务\n关注：展示效果、技术可能性\n常见结果：看起来很惊艳，但无法持续使用', bodySize: 12.2 });
  arrow(demo, 6.48, 3.55, 0.52, C.green);
  card(demo, 7.25, 2.0, 5.35, 3.85, { fill: C.panel2, line: C.green, icon: faGaugeHigh, iconColor: C.green, title: 'PoC / Pilot', titleSize: 18, body: '证明业务问题值得解决\n关注：用户、流程、指标与治理\n目标：决定继续、调整还是停止', bodySize: 12.2 });
  callout(demo, 'Demo 证明“能不能”；PoC 证明“值不值得”；Pilot 证明“能否在真实环境持续运行”。', 0.85, 6.2, 11.6, C.green, faCheckDouble);
  addFooter(demo);

  const value = pptx.addSlide();
  addHeader(value, 'AI 价值评估四层模型', '不要只计算省了多少时间，更要看质量、体验与业务结果', 'Module 3 · Part 4 · 验证价值', C.green);
  const vl = [
    { y: 5.35, w: 7.6, t: 'Level 1 · 体验：用户是否愿意使用？', c: C.blue },
    { y: 4.55, w: 8.6, t: 'Level 2 · 效率：节省多少时间与成本？', c: C.cyan },
    { y: 3.75, w: 9.6, t: 'Level 3 · 质量：错误率、完整性、一致性是否改善？', c: C.purple },
    { y: 2.95, w: 10.6, t: 'Level 4 · 业务：收入、客户价值、风险是否发生变化？', c: C.green }
  ];
  vl.forEach((v, i) => { const x = (13.333 - v.w) / 2; slideShapeRound(value, x, v.y, v.w, 0.56, v.c, 63, v.t); });
  callout(value, '越接近业务结果，证明价值越难，但也越有资格获得规模化投入。', 1.25, 6.25, 10.8, C.green, faArrowTrendUp);
  addFooter(value);

  const roi = pptx.addSlide();
  addHeader(roi, 'AI 项目 ROI：不只有节省成本', '用完整价值公式避免把 AI 限定为裁员与自动化', 'Module 3 · Part 4 · 验证价值', C.green);
  const roiItems = [
    { x: 0.7, t: '节省成本', b: '时间、人力、外包与运营成本', icon: faClock, c: C.blue },
    { x: 3.75, t: '提升收入', b: '转化率、客单价、产能与新产品', icon: faMoneyBillTrendUp, c: C.green },
    { x: 6.8, t: '降低风险', b: '错误、流失、合规与运营中断', icon: faShieldHalved, c: C.orange },
    { x: 9.85, t: '改善体验', b: '客户、员工与合作伙伴体验', icon: faMedal, c: C.purple }
  ];
  roiItems.forEach(r => card(roi, r.x, 2.05, 2.75, 3.85, { fill: C.panel, line: r.c, icon: r.icon, iconColor: r.c, title: r.t, titleSize: 16.5, body: r.b, bodySize: 11.5 }));
  callout(roi, 'AI 项目价值 = 成本节省 + 收入提升 + 风险降低 + 体验改善 − 建设与运营投入。', 0.85, 6.2, 11.6, C.green, faFilterCircleDollar);
  addFooter(roi);

  frameworkSlide('一个好的 AI PoC 应该是什么样？', '4–8 周，小范围、强指标、快速学习', [
    { title: '问题明确', body: '只验证一个核心业务假设', icon: faBullseye, color: C.blue },
    { title: '范围受控', body: '限定用户、数据、流程与权限', icon: faCube, color: C.cyan },
    { title: '基线存在', body: '知道当前耗时、质量与成本', icon: faChartLine, color: C.purple },
    { title: '指标可测', body: '事前定义成功与停止条件', icon: faGaugeHigh, color: C.green },
    { title: '用户参与', body: '业务用户持续反馈与共同设计', icon: faPeopleArrows, color: C.orange },
    { title: '快速学习', body: '验证、复盘、调整或停止', icon: faRepeat, color: C.yellow }
  ], C.green);

  frameworkSlide('AI 项目验收清单', '上线不是完成，持续产生价值才是完成', [
    { title: '业务', body: '是否解决真实问题？指标是否改善？', icon: faBriefcase, color: C.blue },
    { title: '用户', body: '是否愿意使用？是否融入工作？', icon: faUserCheck, color: C.cyan },
    { title: '技术', body: '是否稳定、可扩展、可维护？', icon: faServer, color: C.purple },
    { title: '质量', body: '是否准确、一致、可解释？', icon: faClipboardCheck, color: C.green },
    { title: '治理', body: '权限、隐私、审计与责任是否清楚？', icon: faShieldHalved, color: C.orange },
    { title: '运营', body: '谁监控、更新、评估与持续改进？', icon: faArrowsRotate, color: C.yellow }
  ], C.green);

  const scale = pptx.addSlide();
  addHeader(scale, '从一个 AI 案例，到企业 AI 能力', '不要建设 100 个孤立工具，要沉淀可复用的底层资产', 'Module 3 · Part 5 · 规模推广', C.cyan);
  const layers = [
    { y: 5.38, w: 9.0, t: '治理与安全：权限、审计、评估、责任', c: C.orange },
    { y: 4.65, w: 9.7, t: '数据与知识：统一事实、知识 Owner、版本管理', c: C.blue },
    { y: 3.92, w: 10.4, t: '工具与连接：API、MCP、身份与系统访问', c: C.cyan },
    { y: 3.19, w: 11.1, t: 'Agent 与 Workflow：可复用能力、组件和方法', c: C.purple },
    { y: 2.46, w: 11.8, t: '业务应用：销售、客服、采购、财务、产品与服务', c: C.green }
  ];
  layers.forEach(l => { const x = (13.333 - l.w) / 2; slideShapeRound(scale, x, l.y, l.w, 0.52, l.c, 68, l.t); });
  callout(scale, '规模化的关键不是复制应用界面，而是复用知识、连接、方法、评估与治理。', 1.05, 6.25, 11.2, C.cyan, faLayerGroup);
  addFooter(scale);

  const roadmap = pptx.addSlide();
  addHeader(roadmap, '企业 AI 12 个月路线图', '从机会发现，到价值验证，再到规模复制', 'Module 3 · Part 5 · 规模推广', C.cyan);
  const phases = [
    { x: 0.75, t: '0–3 个月', h: '发现机会', b: '认知统一\n业务访谈\n流程梳理\n场景排序', c: C.blue, icon: faMagnifyingGlassChart },
    { x: 3.8, t: '3–6 个月', h: '验证价值', b: '方案设计\n数据准备\nPoC / Pilot\n指标评估', c: C.purple, icon: faFlask },
    { x: 6.85, t: '6–12 个月', h: '复制能力', b: '标准组件\n知识与连接\n治理机制\n跨部门复制', c: C.cyan, icon: faRepeat },
    { x: 9.9, t: '12 个月+', h: '业务创新', b: 'AI Native 产品\n服务模式\n混合团队\n持续进化', c: C.green, icon: faRocket }
  ];
  phases.forEach((p, i) => { card(roadmap, p.x, 2.0, 2.7, 4.0, { fill: C.panel, line: p.c, icon: p.icon, iconColor: p.c, title: `${p.t} · ${p.h}`, titleSize: 14.8, body: p.b, bodySize: 10.7 }); if (i < 3) arrow(roadmap, p.x + 2.78, 3.65, 0.2, p.c); });
  addFooter(roadmap);

  const workshop = pptx.addSlide();
  addHeader(workshop, 'AI Discovery Workshop：培训之后如何开始？', '把认知与案例，转化成企业自己的机会地图与 PoC 候选', 'Module 3 · 行动入口', C.purple);
  const ws = [
    { x: 0.45, t: '业务访谈', icon: faComments, c: C.blue },
    { x: 2.55, t: '流程梳理', icon: faDiagramProject, c: C.cyan },
    { x: 4.65, t: '机会识别', icon: faLightbulb, c: C.purple },
    { x: 6.75, t: '优先排序', icon: faRankingStar, c: C.orange },
    { x: 8.85, t: '方案设计', icon: faSitemap, c: C.green },
    { x: 10.95, t: 'PoC 规划', icon: faFlask, c: C.yellow }
  ];
  ws.forEach((w, i) => { card(workshop, w.x, 2.2, 1.75, 2.15, { fill: C.panel, line: w.c, icon: w.icon, iconColor: w.c, title: w.t, titleSize: 12.3, shadow: false }); if (i < ws.length - 1) arrow(workshop, w.x + 1.78, 3.13, 0.25, w.c); });
  card(workshop, 0.85, 4.85, 3.6, 1.15, { fill: C.panel2, line: C.blue, icon: faMap, iconColor: C.blue, title: '产出 1', body: '企业 AI 机会地图', bodySize: 10, titleSize: 11.8, shadow: false });
  card(workshop, 4.85, 4.85, 3.6, 1.15, { fill: C.panel2, line: C.purple, icon: faRankingStar, iconColor: C.purple, title: '产出 2', body: 'Top 场景优先级', bodySize: 10, titleSize: 11.8, shadow: false });
  card(workshop, 8.85, 4.85, 3.6, 1.15, { fill: C.panel2, line: C.green, icon: faRocket, iconColor: C.green, title: '产出 3', body: 'PoC 方案与路线建议', bodySize: 10, titleSize: 11.8, shadow: false });
  addFooter(workshop);
}

function workflowSlideModule3(title, subtitle, steps, accent = C.purple) {
  const slide = pptx.addSlide();
  addHeader(slide, title, subtitle, 'Module 3 · Part 3 · 设计方案', accent);
  steps.forEach((st, i) => {
    const x = 0.55 + i * 2.08;
    slideShapeRound(slide, x, 2.35, 1.72, 1.78, st.color, 77, `${i + 1}\n${st.title}`);
    addIcon(slide, st.icon, x + 0.65, 2.67, 0.4, st.color);
    if (i < steps.length - 1) arrow(slide, x + 1.78, 3.12, 0.25, accent);
  });
  card(slide, 1.05, 4.85, 11.2, 1.15, { fill: C.panel2, line: accent, icon: faCircleInfo, iconColor: accent, title: '设计原则', body: '先明确业务目标与人机边界，再选择模型、工具和技术实现；不要反过来。', bodySize: 10.4, titleSize: 11.8, shadow: false });
  addFooter(slide);
}

function module4Slides() {
  sectionSlide(4, 'AI 时代的人与组织', '从“员工使用 AI”，走向“企业管理人 + AI 的混合工作系统”。', ['任务', '人才', '管理', '组织'], C.green);
  const unit = pptx.addSlide();
  addHeader(unit, 'AI 改变的不只是工作，而是工作的基本单元', '从“岗位 = 一个人”到“任务 = 人 + AI 的组合”', 'Module 4 · AI 时代的人与组织', C.green);
  card(unit, 0.75, 2.0, 5.35, 3.9, { fill: C.panel3, line: C.dim, icon: faUserTie, iconColor: C.muted, title: '过去：以岗位组织工作', titleSize: 17, body: '一个岗位包含大量不同任务\n同一个人负责搜集、分析、生成、执行与沟通\n岗位说明书相对稳定', bodySize: 12 });
  arrow(unit, 6.37, 3.55, 0.52, C.green);
  card(unit, 7.2, 2.0, 5.4, 3.9, { fill: C.panel2, line: C.green, icon: faSitemap, iconColor: C.green, title: '未来：以任务重新组合能力', titleSize: 17, body: '任务被拆解为可自动化、可增强与必须由人负责的部分\n人 + AI Agent 按目标动态协作', bodySize: 12 });
  callout(unit, '组织设计开始从“配置多少岗位”，转向“如何组合人、Agent、知识与系统”。', 1.0, 6.2, 11.3, C.green, faSitemap);
  addFooter(unit);

  const employee = pptx.addSlide();
  addHeader(employee, '员工角色变化：从执行者到 AI 协作者', '人的价值不会消失，但会从低价值处理迁移到高价值判断', 'Module 4 · AI 时代的人与组织', C.green);
  const shifts = [
    { old: '信息查找', now: '问题定义', icon: faMagnifyingGlassChart, c: C.blue },
    { old: '重复执行', now: '判断与优化', icon: faRepeat, c: C.cyan },
    { old: '内容生产', now: '创造与指导', icon: faPenNib, c: C.purple },
    { old: '流程操作', now: '任务与 Agent 管理', icon: faDiagramProject, c: C.orange },
    { old: '个人经验', now: '方法沉淀与团队复用', icon: faMedal, c: C.green }
  ];
  shifts.forEach((s, i) => {
    const y = 1.82 + i * 0.98;
    card(employee, 0.8, y, 4.55, 0.75, { fill: C.panel3, line: C.dim, icon: s.icon, iconColor: C.muted, title: s.old, titleSize: 12.8, shadow: false });
    arrow(employee, 5.7, y + 0.22, 0.58, s.c);
    card(employee, 6.65, y, 5.8, 0.75, { fill: C.panel2, line: s.c, icon: faArrowTrendUp, iconColor: s.c, title: s.now, titleSize: 12.8, shadow: false });
  });
  addFooter(employee);

  const talent = pptx.addSlide();
  addHeader(talent, 'AI 时代人才能力模型', '专业能力仍是根基，但需要增加 AI 协作与系统设计能力', 'Module 4 · AI 时代的人与组织', C.green);
  const layers = [
    { y: 5.4, w: 9.6, t: '专业能力：行业知识、岗位技能与业务判断', c: C.blue },
    { y: 4.52, w: 10.35, t: 'AI 使用能力：表达目标、提供上下文、验证结果', c: C.cyan },
    { y: 3.64, w: 11.1, t: 'AI 协作能力：任务拆解、人机分工、持续迭代', c: C.purple },
    { y: 2.76, w: 11.85, t: '系统能力：流程重构、Agent 设计、组织创新', c: C.green }
  ];
  layers.forEach(l => { const x = (13.333 - l.w) / 2; slideShapeRound(talent, x, l.y, l.w, 0.62, l.c, 65, l.t); });
  callout(talent, '最稀缺的不是“会使用某个 AI 工具”，而是能把业务问题转化为可靠人机系统的人。', 0.95, 6.3, 11.4, C.green, faMedal);
  addFooter(talent);

  const manager = pptx.addSlide();
  addHeader(manager, '管理者的新角色：管理人 + AI 团队', '目标设计、人机分工、结果管理与风险控制成为新管理能力', 'Module 4 · AI 时代的人与组织', C.green);
  slideShapeRound(manager, 5.05, 2.3, 3.2, 2.3, C.green, 78, '管理者\n\n定义目标\n配置能力\n管理结果');
  const nodes = [
    { x: 0.75, y: 1.85, t: '员工团队', icon: faPeopleGroup, c: C.blue },
    { x: 0.75, y: 4.45, t: '业务 Owner', icon: faUserTie, c: C.cyan },
    { x: 9.9, y: 1.85, t: 'AI 助手', icon: faPeopleArrows, c: C.purple },
    { x: 9.9, y: 4.45, t: 'AI Agent', icon: faRobot, c: C.orange }
  ];
  nodes.forEach(n => card(manager, n.x, n.y, 2.7, 1.35, { fill: C.panel, line: n.c, icon: n.icon, iconColor: n.c, title: n.t, titleSize: 14.5, shadow: false }));
  callout(manager, '未来管理者不仅要“带人”，还要知道如何给 AI 设目标、配权限、看证据和做纠偏。', 0.95, 6.25, 11.4, C.green, faPersonChalkboard);
  addFooter(manager);

  const office = pptx.addSlide();
  addHeader(office, '企业需要新的 AI 转型推动机制', 'AI 不是单一 IT 项目，需要业务、流程、数据、技术与治理共同负责', 'Module 4 · AI 时代的人与组织', C.green);
  const roles = [
    { x: 0.65, t: '高管 Sponsor', b: '方向、资源与组织阻力', icon: faFlagCheckered, c: C.orange },
    { x: 3.15, t: '业务 Owner', b: '问题、用户与业务指标', icon: faUserTie, c: C.blue },
    { x: 5.65, t: '流程 / 产品', b: '任务重构与方案设计', icon: faSitemap, c: C.cyan },
    { x: 8.15, t: '数据 / IT / AI', b: '数据、连接与工程交付', icon: faServer, c: C.purple },
    { x: 10.65, t: '治理与运营', b: '安全、评估、推广与优化', icon: faShieldHalved, c: C.green }
  ];
  roles.forEach(r => card(office, r.x, 2.1, 2.05, 3.75, { fill: C.panel, line: r.c, icon: r.icon, iconColor: r.c, title: r.t, titleSize: 13.2, body: r.b, bodySize: 9.7 }));
  callout(office, 'AI 转型办公室的价值，不是集中做所有应用，而是建立发现、验证、沉淀与复制的机制。', 0.85, 6.22, 11.6, C.green, faNetworkWired);
  addFooter(office);

  const maturity = pptx.addSlide();
  addHeader(maturity, '企业 AI 成熟度模型', '从认识 AI，到重新设计产品、流程与组织', 'Module 4 · AI 时代的人与组织', C.green);
  const lv = [
    { y: 5.42, w: 7.4, t: 'L0 · Awareness：知道 AI 重要，零散探索', c: C.dim },
    { y: 4.72, w: 8.25, t: 'L1 · AI User：员工开始提升个人效率', c: C.blue },
    { y: 4.02, w: 9.1, t: 'L2 · AI Team：知识、模板与 Workflow 团队复用', c: C.cyan },
    { y: 3.32, w: 9.95, t: 'L3 · AI Process：Agent 进入业务流程并持续运行', c: C.purple },
    { y: 2.62, w: 10.8, t: 'L4 · AI Native：企业重新设计产品、服务与组织', c: C.green }
  ];
  lv.forEach(l => { const x = (13.333 - l.w) / 2; slideShapeRound(maturity, x, l.y, l.w, 0.5, l.c, 63, l.t); });
  callout(maturity, '成熟度不是采购清单，而是组织能否稳定把 AI 转化为业务结果。', 1.25, 6.25, 10.8, C.green, faRankingStar);
  addFooter(maturity);

  const comp = pptx.addSlide();
  addHeader(comp, 'AI 时代的组织竞争力', '真正拉开差距的，不是一次领先，而是持续学习与重构速度', 'Module 4 · AI 时代的人与组织', C.green);
  const cs = [
    { x: 0.72, t: 'AI 能力', b: '模型、Agent、评估与工程能力', icon: faBrain, c: C.blue },
    { x: 3.78, t: '数据能力', b: '统一事实、知识资产与实时数据', icon: faDatabase, c: C.cyan },
    { x: 6.84, t: '流程能力', b: '任务拆解、人机协同与持续优化', icon: faDiagramProject, c: C.purple },
    { x: 9.9, t: '组织能力', b: '决策、人才、治理与变革机制', icon: faBuilding, c: C.green }
  ];
  cs.forEach(c => card(comp, c.x, 2.05, 2.72, 3.9, { fill: C.panel, line: c.c, icon: c.icon, iconColor: c.c, title: c.t, titleSize: 16, body: c.b, bodySize: 10.7 }));
  addFooter(comp);

  const action = pptx.addSlide();
  addHeader(action, '从今天开始：你的第一个 AI 行动是什么？', '把课程认知转化为一个可执行的下一步', 'Module 4 · 课程收束', C.orange);
  const qs = [
    { x: 0.72, n: '01', q: '哪个工作最值得 AI 重新设计？', icon: faBullseye, c: C.blue },
    { x: 4.55, n: '02', q: '哪个流程最浪费人的时间？', icon: faClock, c: C.purple },
    { x: 8.38, n: '03', q: '哪个业务结果最值得提升？', icon: faArrowTrendUp, c: C.green }
  ];
  qs.forEach(q => card(action, q.x, 2.05, 3.35, 3.75, { fill: C.panel, line: q.c, icon: q.icon, iconColor: q.c, title: `${q.n}`, titleSize: 18, body: q.q, bodySize: 14 }));
  callout(action, '先选一个真实任务，完成一张 AI Opportunity Card。', 2.15, 6.15, 9.05, C.orange, faPenNib);
  addFooter(action);

  const path = pptx.addSlide();
  addHeader(path, '从培训到企业 AI 落地', '认知只是起点，真正价值来自持续调研、设计、构建与评估', 'Module 4 · 课程收束', C.cyan);
  const ps = [
    { x: 0.42, t: '认知培训', icon: faPersonChalkboard, c: C.blue },
    { x: 2.48, t: 'AI Discovery', icon: faMagnifyingGlassChart, c: C.cyan },
    { x: 4.54, t: 'Solution Design', icon: faSitemap, c: C.purple },
    { x: 6.6, t: 'PoC Delivery', icon: faFlask, c: C.orange },
    { x: 8.66, t: 'Pilot & Eval', icon: faClipboardCheck, c: C.green },
    { x: 10.72, t: 'Scale Up', icon: faRocket, c: C.yellow }
  ];
  ps.forEach((p, i) => { card(path, p.x, 2.25, 1.75, 2.15, { fill: C.panel, line: p.c, icon: p.icon, iconColor: p.c, title: p.t, titleSize: 11.4, shadow: false }); if (i < ps.length - 1) arrow(path, p.x + 1.78, 3.16, 0.25, p.c); });
  card(path, 0.85, 4.95, 11.6, 1.0, { fill: C.panel2, line: C.cyan, icon: faHandshake, iconColor: C.cyan, title: '最终目标', body: '不是建设更多孤立工具，而是帮助企业形成持续发现机会、交付价值与组织进化的 AI 能力。', bodySize: 10.5, titleSize: 12.2, shadow: false });
  addFooter(path);
}

function endingSlides() {
  const slide = pptx.addSlide();
  slideNo += 1;
  slide.background = { color: C.bg };
  addBg(slide, C.cyan, true);
  slide.addShape(pptx.ShapeType.ellipse, { x: 8.75, y: 0.85, w: 3.5, h: 3.5, fill: { color: C.cyan, transparency: 80 }, line: { color: C.cyan, width: 1.2 } });
  addIcon(slide, faRobot, 9.63, 1.68, 1.7, C.cyan);
  slide.addText('AI 不是终点，\n而是企业智能化的起点。', { x: 0.78, y: 1.45, w: 7.6, h: 1.65, fontFace: FONT, fontSize: 34, bold: true, color: C.white, margin: 0, breakLine: false, fit: 'shrink' });
  slide.addText('理解能力 · 发现机会 · 设计方案 · 落地迭代 · 形成组织能力', { x: 0.82, y: 3.55, w: 7.8, h: 0.45, fontFace: FONT, fontSize: 15.5, color: C.cyan, margin: 0, fit: 'shrink' });
  slide.addText('第一阶段，人学习如何使用 AI。\n第二阶段，企业学习如何与 AI 协作。\n第三阶段，企业重新设计未来。', { x: 0.84, y: 4.55, w: 7.2, h: 1.1, fontFace: FONT, fontSize: 16, color: C.text, margin: 0.02, breakLine: false, fit: 'shrink' });
  slide.addText('ENTERPRISE AI TRANSFORMATION & APPLICATION WORKSHOP', { x: 0.82, y: 6.65, w: 6.8, h: 0.28, fontFace: 'Aptos', fontSize: 9.5, color: C.dim, charSpacing: 1.2, margin: 0 });

  const sources = pptx.addSlide();
  addHeader(sources, '案例与参考来源', '所有外部案例均优先采用官方产品说明与官方客户故事', '附录', C.dim);
  const srcs = [
    'OpenAI — Introducing Frontier（2026）/ Introducing Presence（2026）',
    'OpenAI Customer Stories — Klarna / Zendesk',
    'Anthropic Customer Stories — Zapier / GitLab / IG Group / Trellix',
    'Microsoft Customer Stories — Capita / SCSK / HealthEquity / HEINEKEN',
    'Shopify Help Center — Sidekick 官方产品说明',
    '课程中的销售、采购、客户风险、经营分析与医疗案例为基于企业常见场景的泛化教学案例。'
  ];
  srcs.forEach((s, i) => {
    card(sources, 0.8, 1.75 + i * 0.82, 11.75, 0.62, { fill: C.panel3, line: C.dim, icon: i < 5 ? faBook : faCircleInfo, iconColor: i < 5 ? C.muted : C.orange, title: s, titleSize: 10.5, shadow: false });
  });
  addFooter(sources, '建议在正式对外发布前，根据客户行业补充对应行业案例与本地法规来源。');
}

coverSlide();
threeEraSlide();
beforeAfterWorkSlide();
progressionSlide();
whyDifferentSlide();
rightQuestionSlide();
courseMapSlide();
module1Slides();
module2Slides();
module3Slides();
module4Slides();
endingSlides();

pptx.writeFile({ fileName: '企业AI转型与应用实践_完整培训课件_V1.pptx' });
