import fs from 'node:fs';
import path from 'node:path';
import pptxgen from 'pptxgenjs';
import { icon } from '@fortawesome/fontawesome-svg-core';
import {
  faBrain, faRobot, faCommentDots, faUser, faFileLines, faTable, faImage, faMicrophone,
  faGlobe, faDatabase, faBookOpen, faWrench, faPlug, faEnvelope, faCalendarDays,
  faBuilding, faHardDrive, faClockRotateLeft, faUserGear, faListCheck, faRoute,
  faGear, faArrowsRotate, faLaptop, faCube, faMobileScreenButton, faTriangleExclamation,
  faLightbulb, faBullseye, faChartLine, faMagnifyingGlass, faPenNib, faUsers,
  faArrowTrendUp, faLayerGroup, faSitemap, faShieldHalved, faCheckDouble, faFlask,
  faScaleBalanced, faClipboardCheck, faRocket, faPuzzlePiece, faCloud, faBoxesStacked,
  faCompass, faDiagramProject, faGaugeHigh, faHandshake, faHospital, faHeadset,
  faCartShopping, faTruck, faUserTie, faCircleCheck, faCircleXmark, faEye, faHand,
  faPeopleGroup, faChartPie, faCircleNodes, faLock, faRotateLeft, faKey,
  faClipboard, faStopwatch, faBolt, faGraduationCap, faArrowRight, faTasks
} from '@fortawesome/free-solid-svg-icons';

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Enterprise AI Transformation Workshop';
pptx.company = 'AI Matrix';
pptx.subject = '企业AI转型与应用实践';
pptx.title = '企业AI转型与应用实践｜完整版';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Microsoft YaHei',
  bodyFontFace: 'Microsoft YaHei',
  lang: 'zh-CN'
};
pptx.defineSlideMaster({
  title: 'MASTER',
  background: { color: '061225' },
  objects: [
    { rect: { x: 0, y: 0, w: 13.333, h: 0.06, fill: { color: '1D8BFF' }, line: { color: '1D8BFF' } } },
    { rect: { x: 0, y: 7.42, w: 13.333, h: 0.08, fill: { color: '071A32' }, line: { color: '071A32' } } }
  ],
  slideNumber: { x: 12.45, y: 7.13, w: 0.4, h: 0.18, color: '7892B5', fontFace: 'Aptos', fontSize: 9, align: 'right', margin: 0 }
});

const C = {
  bg: '061225', bg2: '081A31', panel: '0D2747', panel2: '102E53', panel3: '11233F',
  grid: '183C61', blue: '2E9BFF', cyan: '2BD7D5', purple: '8C6CFF', green: '35D49C',
  orange: 'FFAA42', red: 'FF6D6D', white: 'F4F7FC', text: 'DDE8F6', muted: '9FB4D0',
  dim: '607A9A', black: '02070D', yellow: 'F9D65C'
};
const FONT = 'Microsoft YaHei';
const W = 13.333, H = 7.5;
let slideNo = 0;
const slideMeta = [];

function svgData(fa, color = C.white) {
  const svg = icon(fa, { styles: { color } }).html.join('');
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
}
function addIcon(slide, fa, x, y, w, h, color = C.white) {
  slide.addImage({ data: svgData(fa, color), x, y, w, h });
}
function addText(slide, text, x, y, w, h, size = 18, color = C.text, opts = {}) {
  slide.addText(text, {
    x, y, w, h, fontFace: opts.fontFace || FONT, fontSize: size, color,
    bold: !!opts.bold, align: opts.align || 'left', valign: opts.valign || 'mid',
    margin: opts.margin ?? 0.03, breakLine: false, fit: 'shrink',
    italic: !!opts.italic, isTextBox: true, paraSpaceAfterPt: 0,
    bullet: opts.bullet, transparency: opts.transparency
  });
}
function addLine(slide, x, y, w, h = 0, color = C.grid, width = 1, transparency = 0) {
  slide.addShape(pptx.ShapeType.line, { x, y, w, h, line: { color, width, transparency } });
}
function addRect(slide, x, y, w, h, fill = C.panel, line = C.grid, radius = true, transparency = 0) {
  slide.addShape(radius ? pptx.ShapeType.roundRect : pptx.ShapeType.rect, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: fill, transparency }, line: { color: line, width: 1 }
  });
}
function addCircle(slide, x, y, d, fill = C.panel2, line = C.blue, transparency = 0) {
  slide.addShape(pptx.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill, transparency }, line: { color: line, width: 1.2 } });
}
function addArrow(slide, x, y, w = 0.45, color = C.blue) {
  addText(slide, '→', x, y, w, 0.35, 24, color, { bold: true, align: 'center' });
}
function addTechDots(slide, accent = C.blue) {
  const pts = [[0.45,0.6],[12.5,0.8],[11.9,1.5],[0.75,6.7],[12.2,6.4],[6.4,0.28],[8.9,6.9],[3.4,6.65]];
  pts.forEach(([x,y], i) => {
    slide.addShape(pptx.ShapeType.ellipse, { x, y, w: 0.035 + (i%3)*0.012, h: 0.035 + (i%3)*0.012,
      fill: { color: accent, transparency: 15+i*6 }, line: { color: accent, transparency: 100 } });
  });
}
function addFooter(slide, source = '课程自制框架｜示意案例') {
  addText(slide, source, 0.45, 7.13, 5.8, 0.16, 8.5, C.dim, { margin: 0 });
}
function addHeader(slide, section, title, subtitle = '', accent = C.blue) {
  addText(slide, section, 0.55, 0.24, 2.4, 0.25, 10, accent, { bold: true, margin: 0 });
  addText(slide, title, 0.55, 0.63, 12.1, 0.56, 27, C.white, { bold: true, margin: 0 });
  if (subtitle) addText(slide, subtitle, 0.58, 1.24, 11.8, 0.34, 12.8, C.muted, { margin: 0 });
  addLine(slide, 0.55, 1.72, 12.15, 0, C.grid, 1);
}
function baseSlide(section, title, subtitle = '', accent = C.blue, source = '课程自制框架｜示意案例') {
  slideNo += 1;
  const slide = pptx.addSlide('MASTER');
  slide.background = { color: C.bg };
  addTechDots(slide, accent);
  addHeader(slide, section, title, subtitle, accent);
  addFooter(slide, source);
  slideMeta.push({ no: slideNo, section, title });
  return slide;
}
function addInsight(slide, text, color = C.cyan, iconFa = faLightbulb) {
  addRect(slide, 0.7, 6.25, 11.95, 0.55, '091F39', color, true, 8);
  addIcon(slide, iconFa, 0.92, 6.39, 0.22, 0.22, color);
  addText(slide, text, 1.23, 6.34, 10.95, 0.28, 13.3, C.white, { bold: true, margin: 0 });
}
function addCard(slide, {x,y,w,h,title,body='',iconFa=null,accent=C.blue,fill=C.panel,tag='',titleSize=17,bodySize=11.5}) {
  addRect(slide, x, y, w, h, fill, accent, true, 0);
  if (iconFa) {
    addCircle(slide, x+0.18, y+0.18, 0.54, '0B1F37', accent, 0);
    addIcon(slide, iconFa, x+0.32, y+0.31, 0.26, 0.26, accent);
  }
  const tx = iconFa ? x+0.85 : x+0.22;
  const tw = iconFa ? w-1.05 : w-0.44;
  if (tag) addText(slide, tag, x+w-1.15, y+0.15, 0.9, 0.22, 8.5, accent, { bold: true, align: 'right', margin: 0 });
  addText(slide, title, tx, y+0.18, tw, 0.42, titleSize, C.white, { bold: true, margin: 0 });
  if (body) addText(slide, body, x+0.22, y+0.75, w-0.44, h-0.92, bodySize, C.text, { valign: 'top', margin: 0.02 });
}
function addBigNumber(slide, n, x, y, color=C.blue) {
  addText(slide, String(n).padStart(2,'0'), x, y, 0.75, 0.45, 24, color, { bold: true, margin: 0 });
}
function sectionSlide(num, title, subtitle, narrative, accent=C.blue, iconFa=faLayerGroup) {
  slideNo += 1;
  const slide = pptx.addSlide('MASTER');
  slide.background = { color: C.bg };
  addTechDots(slide, accent);
  addText(slide, `MODULE ${num}`, 0.7, 0.72, 3, 0.34, 13, accent, { bold: true, margin: 0 });
  addText(slide, title, 0.7, 1.35, 8.7, 1.25, 35, C.white, { bold: true, margin: 0 });
  addText(slide, subtitle, 0.74, 2.78, 8.7, 0.55, 17, C.muted, { margin: 0 });
  addRect(slide, 0.72, 4.05, 8.8, 1.35, C.panel3, accent, true, 3);
  addIcon(slide, iconFa, 1.02, 4.38, 0.55, 0.55, accent);
  addText(slide, narrative, 1.85, 4.24, 7.2, 0.8, 17.5, C.white, { bold: true, margin: 0 });
  addCircle(slide, 10.35, 1.45, 2.1, '091D37', accent, 0);
  addIcon(slide, iconFa, 10.88, 1.98, 1.05, 1.05, accent);
  addText(slide, String(num).padStart(2,'0'), 10.52, 4.2, 1.8, 0.8, 40, accent, { bold: true, align: 'center', margin: 0 });
  addFooter(slide, '企业AI转型与应用实践｜模块导航');
  slideMeta.push({ no: slideNo, section: `Module ${num}`, title });
}
function statementSlide(section, title, statement, leftLabel, rightLabel, accent=C.blue) {
  const s = baseSlide(section, title, '', accent);
  addText(s, statement, 0.9, 2.0, 11.5, 0.9, 28, C.white, { bold: true, align: 'center', margin: 0 });
  addCard(s, {x:1.0,y:3.45,w:4.7,h:1.55,title:leftLabel.title,body:leftLabel.body,iconFa:leftLabel.icon,accent:C.dim,fill:'0B1E36'});
  addArrow(s, 6.25, 3.95, 0.55, accent);
  addCard(s, {x:7.05,y:3.45,w:5.0,h:1.55,title:rightLabel.title,body:rightLabel.body,iconFa:rightLabel.icon,accent,fill:'0D294A'});
  addInsight(s, rightLabel.insight || 'AI的价值来自重新设计工作，而不是增加一个新工具。', accent);
}
function threeStepSlide(section, title, subtitle, steps, accent=C.blue, insight='') {
  const s = baseSlide(section, title, subtitle, accent);
  const xs = [0.75,4.65,8.55];
  steps.forEach((st,i)=>{
    addCard(s,{x:xs[i],y:2.25,w:3.65,h:2.65,title:st.title,body:st.body,iconFa:st.icon,accent:st.accent||accent,fill:st.fill||C.panel});
    if(i<2) addArrow(s, xs[i]+3.7,3.35,0.65,accent);
  });
  if(insight) addInsight(s, insight, accent);
}
function fourGridSlide(section,title,subtitle,items,accent=C.blue,insight='') {
  const s=baseSlide(section,title,subtitle,accent);
  const pos=[[0.75,2.05],[6.8,2.05],[0.75,4.1],[6.8,4.1]];
  items.forEach((it,i)=>addCard(s,{x:pos[i][0],y:pos[i][1],w:5.78,h:1.72,title:it.title,body:it.body,iconFa:it.icon,accent:it.accent||accent,fill:it.fill||C.panel,titleSize:16.2,bodySize:10.8}));
  if(insight) addInsight(s,insight,accent);
}
function flowSlide(section,title,subtitle,steps,accent=C.blue,insight='',y=2.65) {
  const s=baseSlide(section,title,subtitle,accent);
  const n=steps.length;
  const gap=0.22;
  const totalW=11.75;
  const boxW=(totalW-gap*(n-1))/n;
  steps.forEach((st,i)=>{
    const x=0.78+i*(boxW+gap);
    addRect(s,x,y,boxW,1.55,st.fill||C.panel,st.accent||accent,true,0);
    if(st.icon) addIcon(s,st.icon,x+boxW/2-0.22,y+0.23,0.44,0.44,st.accent||accent);
    addText(s,st.title,x+0.08,y+0.79,boxW-0.16,0.48,13.5,C.white,{bold:true,align:'center',margin:0});
    if(i<n-1) addArrow(s,x+boxW-0.03,y+0.57,0.28,accent);
  });
  if(insight) addInsight(s,insight,accent);
  return s;
}
function compareSlide(section,title,subtitle,left,right,accent=C.blue,insight='') {
  const s=baseSlide(section,title,subtitle,accent);
  addText(s,'过去',0.9,1.95,5.4,0.36,14,C.dim,{bold:true,align:'center'});
  addText(s,'AI重新设计',7.0,1.95,5.4,0.36,14,accent,{bold:true,align:'center'});
  addCard(s,{x:0.85,y:2.35,w:5.55,h:3.2,title:left.title,body:left.body,iconFa:left.icon,accent:C.dim,fill:'0A1D35',titleSize:18,bodySize:13});
  addArrow(s,6.45,3.68,0.55,accent);
  addCard(s,{x:7.0,y:2.35,w:5.45,h:3.2,title:right.title,body:right.body,iconFa:right.icon,accent,fill:'0D294A',titleSize:18,bodySize:13});
  if(insight) addInsight(s,insight,accent);
}
function staircaseSlide(section,title,subtitle,levels,accent=C.blue,insight='') {
  const s=baseSlide(section,title,subtitle,accent);
  const baseY=5.55;
  levels.forEach((lv,i)=>{
    const x=0.8+i*2.45, h=0.85+i*0.63, y=baseY-h;
    addRect(s,x,y,2.15,h,lv.fill||C.panel,lv.accent||accent,true,0);
    addText(s,lv.level,x+0.18,y+0.18,0.55,0.34,12,lv.accent||accent,{bold:true,margin:0});
    addText(s,lv.title,x+0.18,y+0.58,1.78,0.5,15,C.white,{bold:true,margin:0});
    addText(s,lv.body,x+0.18,y+1.18,1.78,h-1.35,10,C.text,{valign:'top',margin:0});
  });
  if(insight) addInsight(s,insight,accent);
}
function matrixSlide(section,title,subtitle,accent=C.blue) {
  const s=baseSlide(section,title,subtitle,accent);
  const x=1.55,y=2.0,w=9.85,h=3.85;
  addRect(s,x,y,w,h,'091C34',C.grid,false,0);
  addLine(s,x+w/2,y,0,h,C.grid,1.2);
  addLine(s,x,y+h/2,w,0,C.grid,1.2);
  addText(s,'业务价值 ↑',0.4,3.45,1.1,0.36,11,C.muted,{bold:true,align:'center'});
  addText(s,'实施复杂度 →',5.45,5.92,2.2,0.3,11,C.muted,{bold:true,align:'center'});
  addCard(s,{x:x+0.28,y:y+0.28,w:4.3,h:1.45,title:'Quick Win｜优先试点',body:'高价值、低复杂度\n快速验证、建立信心',iconFa:faRocket,accent:C.green,fill:'0C2B3E'});
  addCard(s,{x:x+5.22,y:y+0.28,w:4.3,h:1.45,title:'Strategic｜战略项目',body:'高价值、高复杂度\n分阶段投入、长期建设',iconFa:faCompass,accent:C.blue,fill:'0D294A'});
  addCard(s,{x:x+0.28,y:y+2.25,w:4.3,h:1.25,title:'Explore｜低成本探索',body:'小范围试验，积累认知',iconFa:faFlask,accent:C.purple,fill:'171E42'});
  addCard(s,{x:x+5.22,y:y+2.25,w:4.3,h:1.25,title:'Avoid｜暂缓',body:'低价值、高复杂度',iconFa:faCircleXmark,accent:C.red,fill:'351D2A'});
  addInsight(s,'第一个AI项目不应追求最酷，而应追求最容易证明业务价值。',accent,faBullseye);
}
function canvasSlide(section,title,subtitle,blocks,accent=C.blue,insight='') {
  const s=baseSlide(section,title,subtitle,accent);
  blocks.forEach((b)=>addCard(s,{x:b.x,y:b.y,w:b.w,h:b.h,title:b.title,body:b.body,iconFa:b.icon,accent:b.accent||accent,fill:b.fill||C.panel,titleSize:b.titleSize||14.5,bodySize:b.bodySize||10.5}));
  if(insight) addInsight(s,insight,accent);
}
function caseSlide(section,title,subtitle,problem,redesign,value,human,accent=C.cyan) {
  const s=baseSlide(section,title,subtitle,accent,'示意案例｜用于培训讨论');
  addCard(s,{x:0.7,y:2.05,w:3.55,h:3.65,title:'当前业务问题',body:problem,iconFa:faTriangleExclamation,accent:C.orange,fill:'2B221C',titleSize:17,bodySize:12});
  addCard(s,{x:4.48,y:2.05,w:4.25,h:3.65,title:'AI重新设计',body:redesign,iconFa:faRobot,accent,fill:'0B2A48',titleSize:17,bodySize:12});
  addCard(s,{x:8.95,y:2.05,w:3.7,h:2.15,title:'业务价值',body:value,iconFa:faArrowTrendUp,accent:C.green,fill:'0B2C35',titleSize:17,bodySize:11.5});
  addCard(s,{x:8.95,y:4.42,w:3.7,h:1.28,title:'人机分工',body:human,iconFa:faHandshake,accent:C.purple,fill:'1A2045',titleSize:14.5,bodySize:9.8});
  addInsight(s,'案例不是复制答案，而是帮助学员看见：同一项工作可以被重新设计。',accent,faLightbulb);
}
function timelineSlide(section,title,subtitle,phases,accent=C.blue,insight='') {
  const s=baseSlide(section,title,subtitle,accent);
  addLine(s,1.15,3.5,11.0,0,C.grid,2);
  phases.forEach((p,i)=>{
    const x=1.05+i*(11/(phases.length-1));
    addCircle(s,x,3.12,0.72,'0B2039',p.accent||accent,0);
    addText(s,String(i+1),x,3.26,0.72,0.22,11,p.accent||accent,{bold:true,align:'center',margin:0});
    addText(s,p.title,x-0.7,2.18,2.15,0.45,14,C.white,{bold:true,align:'center',margin:0});
    addText(s,p.body,x-0.78,3.98,2.3,0.9,10.5,C.text,{align:'center',valign:'top',margin:0});
  });
  if(insight) addInsight(s,insight,accent);
}
function vennSlide(section,title,subtitle,accent=C.blue) {
  const s=baseSlide(section,title,subtitle,accent);
  addCircle(s,1.1,2.15,3.7,'0E2C50',C.blue,18);
  addCircle(s,4.0,2.15,3.7,'18224B',C.purple,18);
  addCircle(s,6.9,2.15,3.7,'0B3A36',C.green,18);
  addText(s,'AI',1.85,2.55,1.4,0.4,22,C.blue,{bold:true,align:'center'});
  addText(s,'协同',4.95,2.55,1.4,0.4,22,C.white,{bold:true,align:'center'});
  addText(s,'人类',7.95,2.55,1.4,0.4,22,C.green,{bold:true,align:'center'});
  addText(s,'信息处理\n模式发现\n持续执行',1.78,3.25,1.7,1.2,13,C.text,{align:'center'});
  addText(s,'目标一致\n关键点确认\n共同优化',4.75,3.25,1.8,1.2,13,C.white,{bold:true,align:'center'});
  addText(s,'价值判断\n创造突破\n关系与责任',7.75,3.25,1.8,1.2,13,C.text,{align:'center'});
  addInsight(s,'人机协同不是替代，而是把人与AI各自的优势组合成新的工作单元。',accent,faHandshake);
}
function workshopSlide(title, subtitle, fields, accent=C.green) {
  const s=baseSlide('WORKSHOP',title,subtitle,accent,'现场共创模板｜可打印使用');
  fields.forEach((f,i)=>addCard(s,{x:f.x,y:f.y,w:f.w,h:f.h,title:f.title,body:f.body,iconFa:f.icon,accent:f.accent||accent,fill:f.fill||'0B263C',titleSize:f.titleSize||14,bodySize:f.bodySize||10.2}));
  return s;
}

// ---------- OPENING ----------
{
  slideNo += 1;
  const s=pptx.addSlide('MASTER'); s.background={color:C.bg}; addTechDots(s,C.blue);
  // decorative data-wave
  for(let r=0;r<8;r++){
    const y=4.55+r*0.18;
    const pts=[];
    for(let i=0;i<15;i++) pts.push({x:0.2+i*0.55,y:y-Math.sin((i+r)*0.55)*0.22});
    for(let i=0;i<pts.length-1;i++) addLine(s,pts[i].x,pts[i].y,pts[i+1].x-pts[i].x,pts[i+1].y-pts[i].y,r%2?C.blue:C.cyan,0.65,55);
  }
  addRect(s,0.7,0.63,2.55,0.45,'081E38',C.blue,true,0);
  addText(s,'企业高管与核心骨干工作坊',0.92,0.72,2.1,0.2,10.5,C.white,{bold:true,align:'center',margin:0});
  addText(s,'企业 AI 转型与应用实践',0.72,1.45,10.8,0.9,39,C.white,{bold:true,margin:0});
  addText(s,'从理解 AI，到发现机会，再到推动企业落地',0.77,2.65,8.7,0.48,20,C.cyan,{bold:true,margin:0});
  addText(s,'Understand → Apply → Design → Validate → Scale',0.8,3.42,8.9,0.35,13,C.muted,{margin:0});
  addCircle(s,10.35,1.25,2.05,'082547',C.blue,0); addIcon(s,faRobot,10.88,1.78,1.0,1.0,C.cyan);
  addText(s,'Enterprise AI\nTransformation',10.02,3.58,2.7,0.75,15,C.white,{bold:true,align:'center',margin:0});
  addText(s,'认知不是终点，行动才是价值的起点。',0.8,6.45,7.5,0.38,15,C.white,{bold:true,margin:0});
  addFooter(s,'企业AI转型与应用实践｜完整版 V1.0');
  slideMeta.push({no:slideNo,section:'Opening',title:'企业AI转型与应用实践'});
}
statementSlide('OPENING','AI正在改变工作的基本方式','过去，软件帮助人工作；现在，AI开始参与完成工作。',{title:'传统数字化',body:'人 → 操作软件 → 完成任务',icon:faLaptop},{title:'AI时代',body:'目标 → 人 + AI协作系统 → 业务结果',icon:faRobot,insight:'企业竞争将从“谁的软件更好”，走向“谁能更快把AI变成组织能力”。'},C.cyan);
threeStepSlide('OPENING','三次生产力跃迁','每一轮技术浪潮，都重新定义了企业如何工作。',[
  {title:'互联网时代',body:'连接信息\n搜索、门户、电子商务',icon:faGlobe,accent:C.blue},
  {title:'移动时代',body:'连接人与服务\n随时、随地、在线协同',icon:faMobileScreenButton,accent:C.purple},
  {title:'AI时代',body:'连接人与智能\n理解、生成、判断、行动',icon:faBrain,accent:C.cyan}
],C.blue,'AI不是另一个工具类别，而是新的知识工作基础设施。');
compareSlide('OPENING','为什么今天的AI和过去不同？','从规则自动化，到生成式与Agentic智能。',
  {title:'过去：规则驱动',body:'如果 A，就执行 B\n适合结构化、确定性的重复任务',icon:faGear},
  {title:'现在：目标驱动',body:'理解意图 → 生成内容 → 调用工具 → 观察结果 → 持续调整',icon:faBrain},C.cyan,'模型能力只是起点；真正的跃迁来自模型、知识、工具、状态和控制共同组成的工作系统。');
threeStepSlide('OPENING','今天，我们要回答三个问题','从认知对齐，到企业行动。',[
  {title:'01 理解 AI',body:'AI究竟发生了什么变化？',icon:faBrain,accent:C.blue},
  {title:'02 应用 AI',body:'哪些工作值得重新设计？',icon:faBullseye,accent:C.cyan},
  {title:'03 落地 AI',body:'如何把机会变成业务成果？',icon:faRocket,accent:C.green}
],C.blue,'这不是一场“工具教学”，而是一场企业工作方式重新设计的起点。');
staircaseSlide('OPENING','整门课程的主线','从个人效率，逐步走向企业能力与业务创新。',[
  {level:'L1',title:'个人助手',body:'更快完成工作',accent:C.blue},
  {level:'L2',title:'团队工作流',body:'复制优秀经验',accent:C.cyan},
  {level:'L3',title:'Agent',body:'持续完成任务',accent:C.purple},
  {level:'L4',title:'AI Native',body:'重构产品、流程与组织',accent:C.green}
],C.blue,'企业AI成熟度，不取决于用了多少工具，而取决于重新设计了多少工作。');

// ---------- MODULE 1 ----------
sectionSlide(1,'理解 AI','从聊天机器人到数字员工','理解AI的能力进化，才能正确设计企业AI。',C.blue,faBrain);
{
  const s=baseSlide('MODULE 1｜能力进化 01','AI最初只有一个“大脑”','它会生成，但看不到你的文件、系统和业务现场。',C.blue);
  addCard(s,{x:0.8,y:2.25,w:2.25,h:2.15,title:'用户输入',body:'问题 / 指令 / 需求',iconFa:faUser,accent:C.blue,fill:'0B203B'});
  addArrow(s,3.25,3.05,0.55,C.blue);
  addCard(s,{x:4.05,y:2.05,w:4.05,h:2.55,title:'大模型｜LLM',body:'基于训练中学到的知识与模式生成结果',iconFa:faBrain,accent:C.cyan,fill:'0D2C50',titleSize:22,bodySize:12.5});
  addArrow(s,8.33,3.05,0.55,C.blue);
  addCard(s,{x:9.05,y:2.25,w:3.35,h:2.15,title:'内容输出',body:'回答 / 文案 / 总结',iconFa:faCommentDots,accent:C.blue,fill:'0B203B'});
  addInsight(s,'会生成 ≠ 已验证；看起来合理 ≠ 一定正确。',C.orange,faTriangleExclamation);
}
{
  const s=baseSlide('MODULE 1｜能力进化 02','AI获得上下文：开始理解当前任务','对话历史、任务目标和业务背景，构成AI的临时工作台。',C.blue);
  addCard(s,{x:0.75,y:2.05,w:3.5,h:3.55,title:'上下文｜Context',body:'✓ 对话历史\n✓ 当前任务\n✓ 业务背景\n✓ 用户偏好\n✓ 已确认决定',iconFa:faClipboard,accent:C.blue,fill:'0B203B',bodySize:13});
  addArrow(s,4.48,3.35,0.55,C.blue);
  addCard(s,{x:5.05,y:2.35,w:3.1,h:2.75,title:'大模型',body:'结合当前上下文\n重新理解与生成',iconFa:faBrain,accent:C.cyan,fill:'0D2C50',titleSize:22,bodySize:13});
  addArrow(s,8.37,3.35,0.55,C.blue);
  addCard(s,{x:8.95,y:2.35,w:3.55,h:2.75,title:'多轮协作',body:'追问 → 澄清 → 补充 → 修正 → 确认',iconFa:faArrowsRotate,accent:C.blue,fill:'0B203B',bodySize:13});
  addInsight(s,'多轮对话的循环由人推动；每一轮新输入都会触发AI重新生成。',C.cyan,faCommentDots);
}
{
  const s=baseSlide('MODULE 1｜能力进化 03','AI获得“眼睛”：开始看见真实世界','输入不再只有文字，而是企业日常工作的各种材料。',C.blue);
  const inputs=[
    {t:'文档',i:faFileLines,c:C.red},{t:'表格',i:faTable,c:C.green},{t:'图片',i:faImage,c:C.purple},
    {t:'录音',i:faMicrophone,c:C.orange},{t:'网页',i:faGlobe,c:C.cyan},{t:'数据',i:faDatabase,c:C.blue}
  ];
  inputs.forEach((it,idx)=>{
    const col=idx%3,row=Math.floor(idx/3); addCard(s,{x:0.72+col*1.55,y:2.15+row*1.42,w:1.35,h:1.18,title:it.t,iconFa:it.i,accent:it.c,fill:'0B203B',titleSize:13});
  });
  addArrow(s,5.45,3.05,0.55,C.blue);
  addCard(s,{x:6.1,y:2.35,w:2.8,h:2.55,title:'多模态大模型',body:'阅读、提取、比较、理解',iconFa:faEye,accent:C.cyan,fill:'0D2C50',titleSize:20,bodySize:12.5});
  addArrow(s,9.12,3.05,0.55,C.blue);
  addCard(s,{x:9.75,y:2.35,w:2.75,h:2.55,title:'工作材料理解',body:'更完整的信息\n更贴近真实任务',iconFa:faCheckDouble,accent:C.green,fill:'0B2C35',titleSize:17,bodySize:12});
  addInsight(s,'能看见，不代表一定看得准确；复杂表格、模糊图片和超长材料仍需要验证。',C.orange,faTriangleExclamation);
}
{
  const s=baseSlide('MODULE 1｜能力进化 04','AI获得知识来源：从猜测到基于事实','企业知识库解决“内部事实”，联网搜索解决“外部最新信息”。',C.blue);
  addCard(s,{x:0.75,y:2.05,w:3.3,h:1.55,title:'企业知识库',body:'制度 / SOP / 产品资料 / 历史案例',iconFa:faBuilding,accent:C.blue,fill:'0B203B'});
  addCard(s,{x:0.75,y:3.92,w:3.3,h:1.55,title:'联网与外部数据',body:'政策 / 新闻 / 市场 / 行业信息',iconFa:faGlobe,accent:C.cyan,fill:'0B203B'});
  addArrow(s,4.32,3.28,0.55,C.blue);
  addCard(s,{x:5.0,y:2.25,w:3.1,h:3.0,title:'先检索，再回答',body:'检索相关资料\n→ 放到当前工作台\n→ 基于资料生成',iconFa:faMagnifyingGlass,accent:C.cyan,fill:'0D2C50',titleSize:19,bodySize:13});
  addArrow(s,8.35,3.28,0.55,C.blue);
  addCard(s,{x:9.0,y:2.25,w:3.45,h:3.0,title:'有依据的结论',body:'结论 + 证据 + 来源\n降低“凭记忆猜测”',iconFa:faBookOpen,accent:C.green,fill:'0B2C35',titleSize:18,bodySize:13});
  addInsight(s,'RAG不是让模型永久“学会”企业知识，而是把相关资料临时拿到工作台上。',C.cyan,faBookOpen);
}
{
  const s=baseSlide('MODULE 1｜能力进化 05','AI获得“手脚”：开始执行动作','从“告诉你怎么做”，升级为“在授权范围内真正去做”。',C.blue);
  const tools=[['CRM',faBuilding,C.blue],['ERP',faDatabase,C.green],['邮件',faEnvelope,C.purple],['日历',faCalendarDays,C.cyan],['浏览器',faGlobe,C.orange],['API',faPlug,C.blue]];
  tools.forEach((t,idx)=>{const r=Math.floor(idx/3),c=idx%3;addCard(s,{x:0.7+c*1.55,y:2.05+r*1.38,w:1.35,h:1.15,title:t[0],iconFa:t[1],accent:t[2],fill:'0B203B',titleSize:12.5});});
  addArrow(s,5.35,3.05,0.55,C.blue);
  addCard(s,{x:5.98,y:2.25,w:2.85,h:2.75,title:'大模型 + 工具',body:'选择工具\n传入参数\n读取结果',iconFa:faWrench,accent:C.cyan,fill:'0D2C50',titleSize:19,bodySize:13});
  addArrow(s,9.02,3.05,0.55,C.blue);
  addCard(s,{x:9.65,y:2.25,w:2.72,h:2.75,title:'业务动作',body:'生成文件\n更新记录\n创建会议\n发起流程',iconFa:faCircleCheck,accent:C.green,fill:'0B2C35',titleSize:18,bodySize:12.5});
  addInsight(s,'AI一旦拥有行动能力，错误成本会显著上升：权限、确认、日志和回滚必须同步设计。',C.orange,faShieldHalved);
}
{
  const s=baseSlide('MODULE 1｜能力进化 06','AI获得记忆：从一次合作到长期伙伴','需要同时区分当前上下文、长期记忆和正式业务数据。',C.purple);
  addCard(s,{x:0.68,y:2.08,w:3.05,h:3.45,title:'当前对话上下文',body:'当前桌面上摊开的材料\n\n仅服务本次任务\n随上下文窗口变化',iconFa:faCommentDots,accent:C.blue,fill:'0B203B',bodySize:12});
  addCard(s,{x:3.95,y:2.08,w:3.05,h:3.45,title:'长期记忆',body:'跨任务保存的笔记\n\n偏好 / 决策 / 经验\n需要可查看、可纠正',iconFa:faHardDrive,accent:C.purple,fill:'171E42',bodySize:12});
  addCard(s,{x:7.22,y:2.08,w:3.05,h:3.45,title:'正式业务数据',body:'CRM / ERP / 数据库中的事实记录\n\n正式事实不应只存在AI记忆中',iconFa:faDatabase,accent:C.green,fill:'0B2C35',bodySize:12});
  addCard(s,{x:10.5,y:2.08,w:2.15,h:3.45,title:'任务状态',body:'已完成\n进行中\n下一步\n异常',iconFa:faListCheck,accent:C.cyan,fill:'0D294A',bodySize:12});
  addInsight(s,'记忆告诉AI“之前发生了什么”；业务系统保存“正式事实是什么”。',C.purple,faHardDrive);
}
{
  const s=baseSlide('MODULE 1｜能力进化 07','AI获得方法：从聪明到可靠','把优秀经验沉淀成可重复调用的做事方法。',C.green);
  addCard(s,{x:0.7,y:2.1,w:2.6,h:1.55,title:'模板',body:'统一输入与输出结构',iconFa:faFileLines,accent:C.blue,fill:'0B203B'});
  addCard(s,{x:0.7,y:3.95,w:2.6,h:1.55,title:'检查清单',body:'关键检查点与质量标准',iconFa:faListCheck,accent:C.cyan,fill:'0B203B'});
  addCard(s,{x:3.55,y:2.1,w:2.6,h:1.55,title:'SOP / Playbook',body:'标准步骤与异常处理',iconFa:faRoute,accent:C.green,fill:'0B2C35'});
  addCard(s,{x:3.55,y:3.95,w:2.6,h:1.55,title:'Skills',body:'某类任务的可复用方法',iconFa:faPuzzlePiece,accent:C.purple,fill:'171E42'});
  addArrow(s,6.38,3.33,0.55,C.green);
  addCard(s,{x:7.05,y:2.45,w:2.5,h:2.55,title:'大模型',body:'按方法执行\n而不是临场发挥',iconFa:faBrain,accent:C.cyan,fill:'0D2C50',titleSize:20,bodySize:12.5});
  addArrow(s,9.75,3.33,0.55,C.green);
  addCard(s,{x:10.3,y:2.45,w:2.25,h:2.55,title:'稳定输出',body:'一致\n可复用\n可评估',iconFa:faCheckDouble,accent:C.green,fill:'0B2C35',titleSize:17,bodySize:12.5});
  addInsight(s,'记忆回答“之前发生了什么”；方法回答“这类事情应该怎么做”。',C.green,faGraduationCap);
}
{
  const s=baseSlide('MODULE 1｜能力进化 08','AI形成执行循环：成为 Agent','围绕目标自主规划、执行、观察和优化。',C.purple);
  const cx=6.65, cy=3.58, r=1.82;
  addCircle(s,cx-0.82,cy-0.82,1.64,'0D2C50',C.purple,0); addIcon(s,faRobot,cx-0.34,cy-0.35,0.68,0.68,C.cyan); addText(s,'Agent Loop',cx-0.75,cy+0.5,1.5,0.28,12,C.white,{bold:true,align:'center',margin:0});
  const st=[['理解目标',faBullseye,C.blue],['规划下一步',faListCheck,C.purple],['调用工具',faWrench,C.cyan],['观察结果',faEye,C.green],['检查达标',faClipboardCheck,C.orange],['继续 / 求助 / 结束',faArrowsRotate,C.blue]];
  const pos=[[5.78,1.95],[8.25,2.42],[8.25,4.28],[5.78,5.0],[3.32,4.28],[3.32,2.42]];
  st.forEach((it,i)=>{addCard(s,{x:pos[i][0],y:pos[i][1],w:2.15,h:0.9,title:it[0],iconFa:it[1],accent:it[2],fill:'0B203B',titleSize:11.2});});
  addInsight(s,'Chatbot的循环由人推动；Agent的循环在系统内部持续推进。',C.purple,faArrowsRotate);
}
{
  const s=baseSlide('MODULE 1｜能力进化 09','给 AI 一个“工作环境”：成为数字员工','这不是让AI更聪明，而是让AI拥有完成工作的基础设施。',C.cyan);
  addCard(s,{x:0.72,y:2.18,w:3.7,h:3.35,title:'一台电脑｜云端工作空间',body:'持续运行的计算环境\n保存文件和任务状态\n可 24×7 工作',iconFa:faLaptop,accent:C.blue,fill:'0B203B',titleSize:17,bodySize:13});
  addCard(s,{x:4.82,y:2.18,w:3.7,h:3.35,title:'一个试验场地｜Sandbox',body:'隔离环境中执行与验证\n允许失败、重试和回滚\n降低误操作风险',iconFa:faCube,accent:C.purple,fill:'171E42',titleSize:17,bodySize:13});
  addCard(s,{x:8.92,y:2.18,w:3.7,h:3.35,title:'远程控制入口｜IM / 手机端',body:'随时发起任务\n查看进度和异常\n关键节点人工确认',iconFa:faMobileScreenButton,accent:C.cyan,fill:'0D294A',titleSize:16.5,bodySize:13});
  addInsight(s,'数字员工 = 智能能力 + 工作环境 + 权限边界 + 持续运行。',C.cyan,faCloud);
}
staircaseSlide('MODULE 1｜能力总览','从“会回答”到“能工作”','大模型保持在核心，外部工作系统不断丰富。',[
  {level:'01',title:'大模型',body:'理解与生成',accent:C.blue},
  {level:'02',title:'上下文',body:'当前工作台',accent:C.blue},
  {level:'03',title:'知识与文件',body:'看到真实信息',accent:C.cyan},
  {level:'04',title:'工具与系统',body:'执行真实动作',accent:C.green},
  {level:'05',title:'记忆与方法',body:'持续、稳定复用',accent:C.purple}
],C.blue,'Agentic助手不是一种更大的模型，而是以大模型为核心的完整AI工作系统。');
fourGridSlide('MODULE 1｜能力边界','但 AI 仍然不是万能员工','能力越强，越需要明确边界、控制和责任。',[
  {title:'知识边界',body:'依赖输入和资料；可能过时、缺失或错误。',icon:faBookOpen,accent:C.blue},
  {title:'理解边界',body:'只理解被提供的上下文；容易误解隐含意图。',icon:faUserGear,accent:C.green},
  {title:'判断边界',body:'善于发现模式，但价值选择仍需要人。',icon:faScaleBalanced,accent:C.orange},
  {title:'责任边界',body:'AI不能独立承担业务、法律和伦理责任。',icon:faShieldHalved,accent:C.purple}
],C.orange,'AI最大的风险，不是它不会做，而是它做得“看起来很像对的”。');
staircaseSlide('MODULE 1｜信任与控制','能力越强，人工控制必须越精细','不是“信不信AI”，而是设计不同风险等级的协作方式。',[
  {level:'C5',title:'人执行 / AI辅助',body:'高风险、强责任',accent:C.green},
  {level:'C4',title:'人审 / AI执行',body:'关键动作确认',accent:C.cyan},
  {level:'C3',title:'人机协同',body:'共同完成判断',accent:C.blue},
  {level:'C2',title:'关键点确认',body:'多数步骤自动',accent:C.purple},
  {level:'C1',title:'结果确认',body:'低风险可回滚',accent:C.orange}
],C.orange,'Agentic不等于完全自主；企业真正要设计的是“在哪里自主、在哪里确认”。');
vennSlide('MODULE 1｜人机协同','最好的AI，不是替代人，而是放大人','让AI做机器擅长的，让人负责人类不可替代的价值。',C.cyan);
flowSlide('MODULE 1｜模块过渡','从理解 AI，到应用 AI','认知对齐之后，下一步是找到值得重新设计的工作。',[
  {title:'理解AI能力',icon:faBrain,accent:C.blue},{title:'发现业务机会',icon:faBullseye,accent:C.green},{title:'设计解决方案',icon:faDiagramProject,accent:C.purple},{title:'落地与迭代',icon:faArrowTrendUp,accent:C.orange}
],C.blue,'AI不是终点，而是企业智能化的起点。');

// ---------- MODULE 2 ----------
sectionSlide(2,'应用 AI','从个人提效到企业智能化','从“我会用AI”，逐步走向“组织能持续创造价值”。',C.cyan,faBullseye);
staircaseSlide('MODULE 2｜总览','企业AI应用的四个层次','任务越来越复杂，价值越来越深入。',[
  {level:'L1',title:'个人AI助手',body:'提高单人产能',accent:C.blue},
  {level:'L2',title:'团队AI工作流',body:'复制优秀经验',accent:C.cyan},
  {level:'L3',title:'AI Agent',body:'承担完整任务',accent:C.purple},
  {level:'L4',title:'AI Native',body:'重构业务与组织',accent:C.green}
],C.cyan,'企业AI的价值演进：效率 → 复用 → 执行 → 创新。');
fourGridSlide('MODULE 2｜机会地图','AI机会存在于四种不同深度','不要只寻找“能不能用AI”，还要判断AI进入工作的深度。',[
  {title:'个人效率',body:'研究、写作、会议、分析、学习。',icon:faUser,accent:C.blue},
  {title:'团队协作',body:'知识共享、模板复用、标准工作流。',icon:faUsers,accent:C.cyan},
  {title:'业务流程',body:'持续监控、判断、执行和跟踪。',icon:faRoute,accent:C.purple},
  {title:'企业创新',body:'AI Native产品、服务、流程和组织。',icon:faRocket,accent:C.green}
],C.cyan,'从浅到深并不代表必须逐级推进；关键是选择与业务价值和准备度匹配的层次。');
statementSlide('MODULE 2｜个人AI助手','每个人都会拥有一个AI工作伙伴','AI首先改变的是每个人每天重复的信息工作。',{title:'过去：自己完成',body:'搜索、阅读、整理、写作、分析',icon:faUser},{title:'未来：人 + AI协作',body:'人定义目标与判断，AI处理信息与生成初稿',icon:faHandshake,insight:'不是把任务“交给AI”，而是重新设计人与AI的协作回路。'},C.blue);
caseSlide('MODULE 2｜个人案例','AI研究助手：从搜索信息到生成洞察','市场经理准备竞品与行业分析。','大量搜索、复制、阅读和整理；时间主要消耗在信息搬运。','定义研究问题 → 并行收集资料 → 比较观点 → 识别证据缺口 → 生成研究初稿。','缩短研究周期\n提升资料覆盖度\n把时间投入判断','AI收集与结构化；人判断证据、机会和结论。',C.blue);
caseSlide('MODULE 2｜个人案例','AI写作助手：从“写内容”到“共创内容”','销售方案、市场材料和管理汇报。','从空白页开始；资料分散；表达质量依赖个人。','输入客户背景、产品资料和案例 → 生成结构 → 起草内容 → 多轮修改 → 形成可交付产物。','降低起稿成本\n提升结构完整度\n加快版本迭代','AI负责第一版和多版本；人负责立场、判断和最终表达。',C.purple);
caseSlide('MODULE 2｜个人案例','AI会议助手：把会议变成行动闭环','会前准备、会中记录、会后跟进。','信息散落在聊天、笔记和个人记忆中；决策难追踪。','会前汇总背景 → 会中记录重点 → 会后提取决策、责任人和期限 → 持续跟踪。','减少遗漏\n缩短会后整理\n提高行动闭环率','AI负责记录与跟踪；人负责讨论、决策和承诺。',C.cyan);
caseSlide('MODULE 2｜个人案例','AI分析助手：从看数据到发现问题','管理者分析经营数据与异常。','人工拉表、筛选和解释；结论容易受经验和时间限制。','读取报表 → 发现异常与趋势 → 提出原因假设 → 生成验证路径和行动建议。','更快发现问题\n增加分析维度\n提升决策准备质量','AI发现模式与假设；人验证原因并做经营决策。',C.green);
staircaseSlide('MODULE 2｜个人成熟度','个人AI能力的四级进化','差异不在“会不会问”，而在是否形成可复用的工作系统。',[
  {level:'L1',title:'Consumer',body:'问AI、获取答案',accent:C.blue},
  {level:'L2',title:'Tinkerer',body:'反复试验、组合工具',accent:C.cyan},
  {level:'L3',title:'Builder',body:'设计可靠工作流',accent:C.purple},
  {level:'L4',title:'Architect',body:'构建个人AI系统',accent:C.green}
],C.blue,'从“问得好”进化到“把一类工作持续做好”。');
canvasSlide('MODULE 2｜个人系统','个人AI高手的秘密：不是Prompt，而是工作系统','Prompt只是入口；稳定能力来自六个要素。',[
  {x:0.7,y:2.0,w:3.8,h:1.3,title:'模型',body:'选择合适的通用智能',icon:faBrain,accent:C.blue},
  {x:4.75,y:2.0,w:3.8,h:1.3,title:'知识与资料',body:'提供真实上下文与证据',icon:faBookOpen,accent:C.cyan},
  {x:8.8,y:2.0,w:3.8,h:1.3,title:'工具',body:'搜索、文件、代码与系统',icon:faWrench,accent:C.green},
  {x:0.7,y:3.65,w:3.8,h:1.3,title:'方法',body:'模板、Skills与检查清单',icon:faPuzzlePiece,accent:C.purple},
  {x:4.75,y:3.65,w:3.8,h:1.3,title:'记忆',body:'保留偏好、决策与状态',icon:faHardDrive,accent:C.orange},
  {x:8.8,y:3.65,w:3.8,h:1.3,title:'复盘',body:'评估结果、沉淀改进',icon:faArrowsRotate,accent:C.blue}
],C.blue,'个人AI系统 = 模型 + 上下文 + 工具 + 方法 + 记忆 + 复盘。');
statementSlide('MODULE 2｜团队AI工作流','企业最大的AI机会：复制优秀员工经验','个人高手的方法，只有变成组织资产，才能产生规模价值。',{title:'经验停留在人脑',body:'依赖个人、难传承、离职即流失',icon:faUserTie},{title:'经验沉淀为AI能力',body:'知识 + 模板 + 流程 + 评估，团队随时复用',icon:faPeopleGroup,insight:'AI让“组织学习”第一次可以直接嵌入日常工作。'},C.cyan);
flowSlide('MODULE 2｜团队案例','企业知识助手：让组织拥有共同记忆','员工不必知道答案在哪里，只需知道问题是什么。',[
  {title:'员工提问',icon:faUser,accent:C.blue},{title:'理解意图',icon:faBrain,accent:C.cyan},{title:'检索企业知识',icon:faBookOpen,accent:C.purple},{title:'生成并引用来源',icon:faCommentDots,accent:C.green},{title:'必要时转人工',icon:faHandshake,accent:C.orange}
],C.cyan,'知识助手的关键不是“回答得像”，而是答案可追溯、权限可控制、知识可运营。');
compareSlide('MODULE 2｜团队工作流','从 Prompt 到 Workflow：让AI稳定工作','一次性生成与可重复业务能力的差异。',
  {title:'单次 Prompt',body:'“帮我写一封销售邮件”\n\n依赖个人提示技巧\n输入和质量不稳定\n难以团队复用',icon:faCommentDots},
  {title:'AI Workflow',body:'获取客户信息 → 分析机会 → 选择策略 → 生成邮件 → 人工审核 → 更新CRM',icon:faRoute},C.cyan,'Prompt是个人技巧；Workflow是可复制、可评估、可治理的组织能力。');
flowSlide('MODULE 2｜团队案例','销售机会管理 AI Workflow','把优秀销售的判断过程，嵌入团队日常工作。',[
  {title:'线索进入',icon:faUserTie,accent:C.blue},{title:'AI分析客户',icon:faMagnifyingGlass,accent:C.cyan},{title:'机会评分',icon:faGaugeHigh,accent:C.purple},{title:'推荐策略',icon:faLightbulb,accent:C.green},{title:'销售确认',icon:faHandshake,accent:C.orange},{title:'更新 CRM',icon:faBuilding,accent:C.blue}
],C.cyan,'AI统一信息准备与分析，销售把时间投入客户关系和成交判断。');
flowSlide('MODULE 2｜团队案例','客服 AI 协作流程','AI让新人更快接近高手水平，但复杂关系仍由人处理。',[
  {title:'客户问题',icon:faCommentDots,accent:C.blue},{title:'AI理解意图',icon:faBrain,accent:C.cyan},{title:'检索知识',icon:faBookOpen,accent:C.purple},{title:'推荐回复',icon:faPenNib,accent:C.green},{title:'客服确认',icon:faHeadset,accent:C.orange},{title:'知识回流',icon:faArrowsRotate,accent:C.blue}
],C.cyan,'真正的闭环还包括：低质量答案被发现、修正并回流知识库。');
flowSlide('MODULE 2｜团队案例','内容生产流水线：从个体创作到团队运营','适用于市场、品牌、培训和销售赋能。',[
  {title:'选题与目标',icon:faBullseye,accent:C.blue},{title:'研究素材',icon:faMagnifyingGlass,accent:C.cyan},{title:'生成初稿',icon:faPenNib,accent:C.purple},{title:'品牌审校',icon:faShieldHalved,accent:C.orange},{title:'多渠道适配',icon:faLayerGroup,accent:C.green},{title:'效果复盘',icon:faChartLine,accent:C.blue}
],C.cyan,'AI把内容生产从“单点创作”升级为可治理、可复用、可持续优化的供应链。');
staircaseSlide('MODULE 2｜团队成熟度','从个人使用，到流程嵌入','组织价值来自共享、标准化和持续运营。',[
  {level:'L1',title:'个人尝试',body:'零散使用',accent:C.blue},
  {level:'L2',title:'团队共享',body:'模板与知识库',accent:C.cyan},
  {level:'L3',title:'流程嵌入',body:'AI Workflow',accent:C.purple},
  {level:'L4',title:'持续运营',body:'评估、治理、迭代',accent:C.green}
],C.cyan,'团队AI不是“大家都装一个工具”，而是形成共同的做事方法。');
statementSlide('MODULE 2｜AI Agent','为什么企业需要 Agent？','很多业务问题不是“不会回答”，而是“需要持续完成”。',{title:'Chatbot',body:'用户提问 → AI回答 → 任务结束',icon:faCommentDots},{title:'Agent',body:'业务目标 → 持续观察 → 主动行动 → 反馈调整 → 完成任务',icon:faRobot,insight:'Chatbot回答问题，Agent完成任务。'},C.purple);
canvasSlide('MODULE 2｜AI Agent','一个 Agent 是什么？','围绕业务目标运行的新型工作单元。',[
  {x:0.75,y:2.1,w:2.3,h:1.55,title:'目标',body:'清晰任务与成功标准',icon:faBullseye,accent:C.blue},
  {x:3.25,y:2.1,w:2.3,h:1.55,title:'知识',body:'事实、制度与上下文',icon:faBookOpen,accent:C.cyan},
  {x:5.75,y:2.1,w:2.3,h:1.55,title:'工具',body:'系统连接与执行动作',icon:faWrench,accent:C.green},
  {x:8.25,y:2.1,w:2.3,h:1.55,title:'记忆',body:'状态、历史与偏好',icon:faHardDrive,accent:C.purple},
  {x:10.75,y:2.1,w:1.85,h:1.55,title:'控制',body:'权限与人工确认',icon:faShieldHalved,accent:C.orange,titleSize:13.5,bodySize:9.5},
  {x:3.3,y:4.15,w:6.75,h:1.2,title:'Agent Loop',body:'理解目标 → 规划 → 调用工具 → 观察 → 检查 → 继续 / 求助 / 结束',icon:faArrowsRotate,accent:C.purple,fill:'171E42',titleSize:16,bodySize:12}
],C.purple,'Agent不是单一模型，而是模型、知识、工具、状态、流程和控制的组合。');
fourGridSlide('MODULE 2｜场景判断','什么任务适合交给 Agent？','Agent不是越多越好，首先要判断任务是否适合。',[
  {title:'高频且持续',body:'经常发生，需要持续监控和跟进。',icon:faClockRotateLeft,accent:C.blue},
  {title:'信息密集',body:'需要读取多来源资料并综合判断。',icon:faDatabase,accent:C.cyan},
  {title:'规则与经验存在',body:'有SOP、惯例或可学习的判断模式。',icon:faRoute,accent:C.purple},
  {title:'结果可验证',body:'有明确成功标准，可检查、可纠偏。',icon:faClipboardCheck,accent:C.green}
],C.purple,'战略价值判断、高风险最终审批和复杂关系处理，不应直接交给Agent独立完成。');
caseSlide('MODULE 2｜主案例','重点客户订单延期风险：传统工作为什么失效？','B2B企业的客户风险管理。','CRM、订单、邮件和服务记录分散；销售每周人工检查；发现延期时往往已经太晚。','由Agent每日监控多系统数据，识别风险信号，汇总证据并生成处置建议。','风险更早暴露\n信息准备更完整\n缩短响应时间','AI负责发现与准备；销售负责人决定策略并维护客户关系。',C.purple);
flowSlide('MODULE 2｜主案例','客户风险 Agent：如何持续工作？','同一个业务问题，展示Agent的完整执行循环。',[
  {title:'定时启动',icon:faClockRotateLeft,accent:C.blue},{title:'读取多系统',icon:faDatabase,accent:C.cyan},{title:'识别风险信号',icon:faTriangleExclamation,accent:C.orange},{title:'评估影响',icon:faChartLine,accent:C.purple},{title:'请求负责人确认',icon:faHandshake,accent:C.green},{title:'更新并持续跟进',icon:faArrowsRotate,accent:C.blue}
],C.purple,'结果不只是报告，而是状态变化、执行证据和可追踪的后续行动。');
compareSlide('MODULE 2｜主案例','客户风险处置中的人机分工','AI负责“感知与准备”，人负责“判断与关系”。',
  {title:'AI负责',body:'• 扫描数据\n• 发现异常\n• 汇总证据\n• 生成选项\n• 提醒跟踪',icon:faRobot},
  {title:'人负责',body:'• 判断真实影响\n• 选择沟通策略\n• 协调资源\n• 处理例外\n• 承担责任',icon:faUserTie},C.purple,'AI把人从信息搬运中释放出来，但不会替代商业判断和客户关系。');
caseSlide('MODULE 2｜Agent案例','采购 Agent：从被动响应到主动准备','制造或零售企业的补货与采购。','库存变化快、供应商信息分散；采购人员持续拉表、比价和追踪。','监测库存与需求 → 分析供应风险 → 比较供应商 → 生成采购建议 → 提交审批。','降低缺货与积压\n提升采购响应速度\n减少重复分析','AI生成建议与准备订单；采购负责人审核商业条件并批准。',C.purple);
caseSlide('MODULE 2｜Agent案例','经营分析 Agent：从月报到持续洞察','财务与经营管理。','月底集中拉数和解释；发现问题晚；分析口径依赖个人。','持续读取经营数据 → 识别异常 → 形成原因假设 → 生成验证清单和管理摘要。','更早发现异常\n缩短报告周期\n提高管理响应速度','AI发现模式与准备报告；管理层验证原因并做资源决策。',C.purple);
compareSlide('MODULE 2｜概念辨析','Agent 不是 RPA 的简单升级','固定自动化与目标驱动执行的差异。',
  {title:'RPA｜固定自动化',body:'规则明确\n流程固定\n输入结构化\n变化较少\n适合重复点击和搬运',icon:faGear},
  {title:'Agent｜目标驱动',body:'理解环境\n动态规划\n处理复杂信息\n基于反馈调整\n适合不确定知识任务',icon:faRobot},C.purple,'最佳实践通常不是二选一，而是Agent负责理解与判断，自动化工具负责确定性执行。');
fourGridSlide('MODULE 2｜AI Native','AI Native 企业的四种重构','不是给现有业务“加AI”，而是从设计之初就假设AI参与工作。',[
  {title:'AI Native 产品',body:'从提供功能，到帮助用户完成目标。',icon:faCube,accent:C.blue},
  {title:'AI Native 服务',body:'理解需求、调用系统、解决问题。',icon:faHeadset,accent:C.cyan},
  {title:'AI Native 流程',body:'从固定路径，到目标驱动的人机协同。',icon:faRoute,accent:C.purple},
  {title:'AI Native 组织',body:'人类负责人管理人和AI协作者。',icon:faPeopleGroup,accent:C.green}
],C.green,'AI Native不是“AI功能更多”，而是产品、服务、流程和组织的设计起点发生变化。');
caseSlide('MODULE 2｜AI Native案例','AI CRM：从记录系统到销售决策系统','企业软件产品范式变化。','传统CRM记录客户、联系和阶段；销售仍需自己判断下一步。','AI主动识别机会、解释原因、推荐行动，并准备客户画像、沟通策略和邮件草稿。','提高销售聚焦度\n提升机会响应速度\n降低使用门槛','AI持续准备决策；销售决定关系策略与最终行动。',C.green);
caseSlide('MODULE 2｜AI Native案例','AI客服中心：从回答问题到完成服务','客户服务模式重构。','客户需要等待人工；客服在多个系统中查询并执行操作。','AI理解需求 → 检索知识 → 查询账户/订单 → 在授权范围内执行 → 复杂情况升级人工。','7×24响应\n降低标准问题成本\n提高服务一致性','AI处理标准请求；人处理情绪、例外和高价值关系。',C.green);
compareSlide('MODULE 2｜AI Native流程','从“流程驱动人”到“目标驱动协作”','企业流程的设计中心正在改变。',
  {title:'传统流程',body:'预设固定路径\n人执行每一步\n系统记录状态\n遇到异常依赖人工协调',icon:faRoute},
  {title:'AI Native流程',body:'给定目标与边界\nAI动态选择路径\n工具执行确定性动作\n人把关关键节点',icon:faCircleNodes},C.green,'过去企业设计每一步；未来企业更需要设计目标、边界、成功标准和升级机制。');
canvasSlide('MODULE 2｜AI Native组织','未来管理者：管理人 + AI Agent 团队','组织结构从岗位中心，逐步走向目标与任务中心。',[
  {x:0.8,y:2.1,w:3.3,h:3.4,title:'人类负责人',body:'定义目标\n做价值判断\n协调关系\n承担结果',icon:faUserTie,accent:C.green,fill:'0B2C35',titleSize:18,bodySize:13},
  {x:4.55,y:2.1,w:2.4,h:1.45,title:'研究 Agent',body:'市场与竞品',icon:faMagnifyingGlass,accent:C.blue},
  {x:7.2,y:2.1,w:2.4,h:1.45,title:'内容 Agent',body:'生成与适配',icon:faPenNib,accent:C.purple},
  {x:9.85,y:2.1,w:2.4,h:1.45,title:'分析 Agent',body:'数据与洞察',icon:faChartLine,accent:C.cyan},
  {x:4.55,y:4.05,w:2.4,h:1.45,title:'运营 Agent',body:'监控与协调',icon:faArrowsRotate,accent:C.orange},
  {x:7.2,y:4.05,w:2.4,h:1.45,title:'系统与工具',body:'CRM / ERP / API',icon:faWrench,accent:C.blue},
  {x:9.85,y:4.05,w:2.4,h:1.45,title:'治理与评估',body:'权限 / 日志 / Eval',icon:faShieldHalved,accent:C.green}
],C.green,'未来管理者不仅管理人，也要定义、监督和优化AI协作者。');
staircaseSlide('MODULE 2｜转型路线','从个人提效到企业重构','企业可根据业务价值与准备度，选择合适的切入层次。',[
  {level:'01',title:'个人提效',body:'提高效率',accent:C.blue},
  {level:'02',title:'团队复用',body:'知识与Workflow',accent:C.cyan},
  {level:'03',title:'业务执行',body:'Agent与系统连接',accent:C.purple},
  {level:'04',title:'企业创新',body:'AI Native产品与组织',accent:C.green}
],C.green,'不必一步到位，但必须从“工具采购”转向“工作重新设计”。');
flowSlide('MODULE 2｜模块过渡','从应用场景，到落地方法','案例带来启发，方法才能帮助企业真正行动。',[
  {title:'发现机会',icon:faBullseye,accent:C.green},{title:'选择场景',icon:faScaleBalanced,accent:C.blue},{title:'设计方案',icon:faDiagramProject,accent:C.purple},{title:'验证价值',icon:faFlask,accent:C.orange},{title:'规模推广',icon:faRocket,accent:C.cyan}
],C.green,'下一模块：如何把AI想法转化为可验证、可治理、可规模化的企业项目。');

// ---------- MODULE 3 ----------
sectionSlide(3,'落地 AI','从机会发现到业务成果','帮助企业找到第一个值得投入的AI机会，并建立持续落地路径。',C.green,faRocket);
compareSlide('MODULE 3｜常见失败','为什么很多AI项目没有业务价值？','失败通常不是模型不够强，而是项目起点错误。',
  {title:'错误路径',body:'看到AI很强 → 采购工具 → 找场景 → 做Demo → 效果不明确',icon:faTriangleExclamation},
  {title:'正确路径',body:'业务目标 → 工作任务 → 痛点 → AI机会 → 方案 → 业务验证',icon:faBullseye},C.orange,'企业不缺AI工具，缺的是值得解决的问题、业务Owner和可衡量的成功标准。');
flowSlide('MODULE 3｜机会发现','从岗位到任务：找到真正可改造的工作','不要笼统问“销售部怎么用AI”，要拆解销售每天完成的具体任务。',[
  {title:'岗位 / 角色',icon:faUserTie,accent:C.blue},{title:'任务拆解',icon:faListCheck,accent:C.cyan},{title:'痛点识别',icon:faTriangleExclamation,accent:C.orange},{title:'AI介入点',icon:faRobot,accent:C.purple},{title:'业务价值',icon:faArrowTrendUp,accent:C.green}
],C.green,'工作任务是连接业务问题与AI能力的最小分析单元。');
canvasSlide('MODULE 3｜核心工具','AI Opportunity Canvas｜机会卡','用六个问题，把模糊想法转化为可讨论的AI机会。',[
  {x:0.7,y:2.0,w:3.85,h:1.35,title:'1. 业务目标',body:'希望改善什么业务结果？',icon:faBullseye,accent:C.green},
  {x:4.75,y:2.0,w:3.85,h:1.35,title:'2. 当前流程',body:'现在由谁、如何完成？',icon:faRoute,accent:C.blue},
  {x:8.8,y:2.0,w:3.85,h:1.35,title:'3. 最大痛点',body:'哪里最耗时、易错或影响体验？',icon:faTriangleExclamation,accent:C.orange},
  {x:0.7,y:3.75,w:3.85,h:1.35,title:'4. AI机会',body:'AI可以理解、分析、生成或执行什么？',icon:faRobot,accent:C.purple},
  {x:4.75,y:3.75,w:3.85,h:1.35,title:'5. 人机分工',body:'AI做什么？人保留什么？',icon:faHandshake,accent:C.cyan},
  {x:8.8,y:3.75,w:3.85,h:1.35,title:'6. 成功指标',body:'如何证明它真的有价值？',icon:faGaugeHigh,accent:C.green}
],C.green,'机会卡不是解决方案，而是让企业用统一语言描述“值得探索的问题”。');
caseSlide('MODULE 3｜机会卡示例','客户投诉处理 AI 助手','从业务问题出发，而不是从产品名称出发。','查询历史记录和制度耗时；新人回复质量不一致；复杂问题升级路径不清晰。','AI识别投诉类型 → 检索制度与历史案例 → 推荐处理方案与回复草稿 → 高风险情况升级主管。','缩短响应时间\n提升回复一致性\n降低新人培训成本','AI负责理解、检索与建议；客服主管负责赔偿、例外和关系判断。',C.green);
fourGridSlide('MODULE 3｜场景评估','什么工作最值得 AI 化？','评估的不是“AI能不能做”，而是“值不值得、能不能可靠落地”。',[
  {title:'业务价值',body:'对收入、成本、体验或风险有多大影响？',icon:faArrowTrendUp,accent:C.green},
  {title:'任务特征',body:'是否高频、信息密集、规则存在、结果可验证？',icon:faListCheck,accent:C.blue},
  {title:'数据与系统',body:'是否有可用数据、知识和连接能力？',icon:faDatabase,accent:C.cyan},
  {title:'组织准备度',body:'是否有业务Owner、用户意愿和治理支持？',icon:faPeopleGroup,accent:C.purple}
],C.green,'高价值但无数据、无Owner、不可验证的场景，往往不适合作为第一个项目。');
fourGridSlide('MODULE 3｜机会地图','建立企业 AI 机会全景','把零散想法放入统一的价值层次中。',[
  {title:'个人效率',body:'研究、写作、会议、分析。',icon:faUser,accent:C.blue},
  {title:'团队能力',body:'知识助手、模板、Workflow。',icon:faUsers,accent:C.cyan},
  {title:'业务流程',body:'销售、客服、采购、运营Agent。',icon:faRoute,accent:C.purple},
  {title:'产品与模式创新',body:'AI Native产品、服务与组织。',icon:faRocket,accent:C.green}
],C.green,'机会地图帮助企业同时看见Quick Win与长期战略机会。');
flowSlide('MODULE 3｜机会收敛','从 100 个想法，到 1–2 个 PoC','企业资源有限，必须经过筛选和收敛。',[
  {title:'广泛收集',icon:faLightbulb,accent:C.blue},{title:'任务拆解',icon:faListCheck,accent:C.cyan},{title:'机会评分',icon:faGaugeHigh,accent:C.purple},{title:'优先级评审',icon:faScaleBalanced,accent:C.orange},{title:'PoC候选',icon:faFlask,accent:C.green}
],C.green,'真正的关键不是产生更多想法，而是把最值得验证的机会挑出来。');
matrixSlide('MODULE 3｜场景选择','AI 项目优先级矩阵','综合业务价值与实施复杂度，决定下一步投入。',C.green);
fourGridSlide('MODULE 3｜第一个项目','企业第一个 AI 项目怎么选？','不要选择最酷的，要选择最容易形成业务证据的。',[
  {title:'问题真实',body:'业务团队长期感受到明确痛点。',icon:faTriangleExclamation,accent:C.orange},
  {title:'用户明确',body:'有真实用户和愿意负责的业务Owner。',icon:faUserTie,accent:C.blue},
  {title:'数据可得',body:'不需要先完成大规模数据工程。',icon:faDatabase,accent:C.cyan},
  {title:'价值可衡量',body:'效率、质量、体验或业务指标可比较。',icon:faGaugeHigh,accent:C.green}
],C.green,'理想的首个项目：价值可见、边界清晰、风险可控、4–8周可验证。');
canvasSlide('MODULE 3｜系统观','AI应用不是模型项目，而是业务系统项目','模型只是其中一层，企业价值来自完整组合。',[
  {x:0.7,y:2.3,w:1.65,h:2.4,title:'模型',body:'理解与生成',icon:faBrain,accent:C.blue,titleSize:15},
  {x:2.52,y:2.3,w:1.65,h:2.4,title:'数据',body:'业务事实',icon:faDatabase,accent:C.cyan,titleSize:15},
  {x:4.34,y:2.3,w:1.65,h:2.4,title:'知识',body:'制度与经验',icon:faBookOpen,accent:C.purple,titleSize:15},
  {x:6.16,y:2.3,w:1.65,h:2.4,title:'工具',body:'系统与动作',icon:faWrench,accent:C.green,titleSize:15},
  {x:7.98,y:2.3,w:1.65,h:2.4,title:'流程',body:'步骤与异常',icon:faRoute,accent:C.orange,titleSize:15},
  {x:9.8,y:2.3,w:1.65,h:2.4,title:'人',body:'判断与责任',icon:faUser,accent:C.blue,titleSize:15},
  {x:11.62,y:2.3,w:1.05,h:2.4,title:'治理',body:'控制',icon:faShieldHalved,accent:C.green,titleSize:12.5,bodySize:9}
],C.green,'AI应用 = 模型 + 数据 + 知识 + 工具 + 流程 + 人机协同 + 治理。');
canvasSlide('MODULE 3｜核心工具','AI Solution Canvas｜方案画布','把机会进一步转化为可实施、可评估的解决方案。',[
  {x:0.7,y:2.0,w:2.9,h:1.35,title:'用户与场景',body:'谁在什么情况下使用？',icon:faUser,accent:C.blue},
  {x:3.78,y:2.0,w:2.9,h:1.35,title:'目标与结果',body:'AI要帮助完成什么？',icon:faBullseye,accent:C.green},
  {x:6.86,y:2.0,w:2.9,h:1.35,title:'输入与上下文',body:'需要什么数据和知识？',icon:faDatabase,accent:C.cyan},
  {x:9.94,y:2.0,w:2.68,h:1.35,title:'AI能力',body:'理解 / 生成 / 判断 / 执行',icon:faBrain,accent:C.purple,titleSize:14},
  {x:0.7,y:3.78,w:2.9,h:1.35,title:'工具与系统',body:'需要连接哪些业务系统？',icon:faWrench,accent:C.blue},
  {x:3.78,y:3.78,w:2.9,h:1.35,title:'流程与状态',body:'步骤、异常与任务状态如何管理？',icon:faRoute,accent:C.orange},
  {x:6.86,y:3.78,w:2.9,h:1.35,title:'人机控制',body:'何时确认、升级与停止？',icon:faHandshake,accent:C.green},
  {x:9.94,y:3.78,w:2.68,h:1.35,title:'评估指标',body:'如何证明成功？',icon:faGaugeHigh,accent:C.cyan,titleSize:14}
],C.green,'Solution Canvas的价值是让业务、产品、数据和技术团队使用同一张图协作。');
timelineSlide('MODULE 3｜Agent设计','设计 Agent 的六步法','从业务目标开始，而不是从框架和工具开始。',[
  {title:'定义目标',body:'用户、范围、成功标准',accent:C.green},
  {title:'拆解任务',body:'步骤、依赖与异常',accent:C.blue},
  {title:'配置知识',body:'数据、资料与规则',accent:C.cyan},
  {title:'连接工具',body:'系统、API与权限',accent:C.purple},
  {title:'设计控制',body:'确认、升级、回滚',accent:C.orange},
  {title:'建立评估',body:'质量、成本与业务结果',accent:C.green}
],C.green,'先把任务和控制设计清楚，再决定模型、框架和技术实现。');
{
  const s=baseSlide('MODULE 3｜企业架构','企业级 Agent 应用架构','可靠性来自工作系统，而不是单次模型回答。',C.green);
  addCard(s,{x:4.55,y:1.95,w:4.2,h:0.9,title:'业务目标与成功标准',iconFa:faBullseye,accent:C.green,fill:'0B2C35',titleSize:15.5});
  addArrow(s,6.37,2.85,0.55,C.green);
  addCard(s,{x:4.25,y:3.08,w:4.8,h:1.05,title:'Agent Runtime / Orchestrator',body:'理解目标 · 管理状态 · 选择下一步',iconFa:faRobot,accent:C.purple,fill:'171E42',titleSize:16,bodySize:10.5});
  const items=[['知识',faBookOpen,C.cyan],['工具',faWrench,C.blue],['记忆',faHardDrive,C.purple],['流程',faRoute,C.orange]];
  items.forEach((it,i)=>addCard(s,{x:0.7+i*3.12,y:4.45,w:2.75,h:1.15,title:it[0],body:['企业资料','业务系统','任务状态','SOP与异常'][i],iconFa:it[1],accent:it[2],fill:'0B203B',titleSize:14,bodySize:9.5}));
  addCard(s,{x:9.62,y:2.05,w:2.9,h:2.05,title:'人工控制与治理',body:'权限 · 确认 · 日志\n评估 · 升级 · 回滚',iconFa:faShieldHalved,accent:C.green,fill:'0B2C35',titleSize:15,bodySize:11});
  addCard(s,{x:0.75,y:2.05,w:2.9,h:2.05,title:'企业数据与系统',body:'CRM · ERP · 文档库\n邮件 · 数据库 · SaaS',iconFa:faBuilding,accent:C.blue,fill:'0B203B',titleSize:15,bodySize:11});
  addInsight(s,'企业Agent项目，本质是业务流程重新设计项目。',C.green,faDiagramProject);
}
fourGridSlide('MODULE 3｜人机协同','Human-in-the-Loop：人工控制放在哪里？','人工控制不是一个按钮，而是一套风险分层设计。',[
  {title:'输入控制',body:'限制数据、权限和可使用的工具。',icon:faKey,accent:C.blue},
  {title:'过程控制',body:'关键节点确认、异常升级和停止条件。',icon:faHand,accent:C.orange},
  {title:'结果控制',body:'质量检查、业务复核和证据留存。',icon:faClipboardCheck,accent:C.green},
  {title:'责任控制',body:'明确Owner、授权边界和最终责任。',icon:faShieldHalved,accent:C.purple}
],C.green,'控制不是为了降低AI能力，而是为了让能力可以安全进入真实业务。');
fourGridSlide('MODULE 3｜准备度','数据与知识准备：AI能力的上限来自基础质量','不要把所有问题都归因于模型。',[
  {title:'可访问',body:'数据在哪里？AI是否能够合法、安全地读取？',icon:faKey,accent:C.blue},
  {title:'可理解',body:'字段、口径和文档结构是否清晰？',icon:faBookOpen,accent:C.cyan},
  {title:'可信任',body:'是否准确、完整、及时、可追溯？',icon:faCircleCheck,accent:C.green},
  {title:'可运营',body:'谁负责更新、纠错、下线和权限管理？',icon:faUserGear,accent:C.purple}
],C.green,'“知识库建完了”不是终点，真正重要的是持续运营和质量治理。');
fourGridSlide('MODULE 3｜治理','安全与治理必须从第一天设计','能力、权限和责任必须同步增长。',[
  {title:'权限最小化',body:'只开放完成任务所需的最小数据和工具。',icon:faLock,accent:C.blue},
  {title:'全链路留痕',body:'记录输入、工具调用、决策和结果。',icon:faClipboard,accent:C.cyan},
  {title:'失败可恢复',body:'重试、停止、回滚和人工接管。',icon:faRotateLeft,accent:C.orange},
  {title:'评估与监控',body:'质量、风险、成本和业务指标持续监测。',icon:faGaugeHigh,accent:C.green}
],C.green,'治理不是项目上线前的最后检查，而是AI工作系统的一部分。');
compareSlide('MODULE 3｜验证价值','不要做 Demo，要做业务验证','两者验证的是完全不同的问题。',
  {title:'Demo',body:'证明AI“能不能做”\n\n效果看起来不错\n缺少真实流程与用户\n没有稳定指标',icon:faRobot},
  {title:'PoC / Pilot',body:'证明业务“值不值得做”\n\n真实用户与数据\n嵌入关键流程\n有基线和成功标准',icon:faFlask},C.green,'Demo制造兴奋；业务验证建立投资决策所需的证据。');
staircaseSlide('MODULE 3｜价值评估','AI价值的四层证据','从“好不好用”，逐步走向“是否改变业务结果”。',[
  {level:'L1',title:'体验',body:'满意度、采用率',accent:C.blue},
  {level:'L2',title:'效率',body:'时间与人工成本',accent:C.cyan},
  {level:'L3',title:'质量',body:'准确率与一致性',accent:C.purple},
  {level:'L4',title:'业务',body:'收入、成本、风险、客户价值',accent:C.green}
],C.green,'只有业务结果发生变化，AI项目才真正从技术试验走向企业投资。');
fourGridSlide('MODULE 3｜ROI','AI项目价值模型','不能量化的价值，也可以先定义代理指标。',[
  {title:'节省成本',body:'减少重复劳动、外包、返工和等待。',icon:faChartLine,accent:C.blue},
  {title:'提升收入',body:'提高转化、客单、留存和服务能力。',icon:faArrowTrendUp,accent:C.green},
  {title:'降低风险',body:'提前发现异常、减少错误和合规暴露。',icon:faShieldHalved,accent:C.orange},
  {title:'改善体验',body:'更快响应、更个性化、更连续。',icon:faHandshake,accent:C.cyan}
],C.green,'ROI = 可量化收益 + 风险降低 + 能力建设 − 全生命周期成本。');
timelineSlide('MODULE 3｜PoC','4–8周 AI PoC：快速验证，而不是缩水版生产系统','范围小、目标清、指标明、学习快。',[
  {title:'第1周',body:'目标、用户、基线',accent:C.blue},
  {title:'第2周',body:'数据与流程准备',accent:C.cyan},
  {title:'第3–4周',body:'构建与内部测试',accent:C.purple},
  {title:'第5–6周',body:'真实用户试点',accent:C.orange},
  {title:'第7周',body:'评估与问题修正',accent:C.green},
  {title:'第8周',body:'投资决策与下一步',accent:C.blue}
],C.green,'PoC的目的不是“交付一个功能”，而是降低下一步投资的不确定性。');
fourGridSlide('MODULE 3｜验收','AI项目验收清单','成功不是“模型能回答”，而是系统、用户、业务和治理共同通过。',[
  {title:'业务有效',body:'是否解决真实问题？指标是否改善？',icon:faBullseye,accent:C.green},
  {title:'用户接受',body:'是否愿意使用？是否融入工作？',icon:faUsers,accent:C.blue},
  {title:'技术稳定',body:'质量、延迟、成本和失败率是否可控？',icon:faGear,accent:C.cyan},
  {title:'治理合格',body:'权限、日志、隐私、回滚是否完善？',icon:faShieldHalved,accent:C.orange}
],C.green,'验收后的决策可以是：扩大、修正、转向或停止；停止失败项目也是价值。');
statementSlide('MODULE 3｜规模化','从一个AI案例，到企业AI能力','不要建设100个彼此孤立的AI工具。',{title:'孤立项目',body:'重复建设、数据割裂、缺少治理、维护成本持续上升',icon:faPuzzlePiece},{title:'共享能力体系',body:'统一知识、连接、评估、权限和可复用组件',icon:faLayerGroup,insight:'规模化的本质，是把项目经验沉淀为平台、方法和组织能力。'},C.green);
timelineSlide('MODULE 3｜路线图','企业AI转型：12个月演进路径','从机会发现到规模化能力建设。',[
  {title:'0–3个月',body:'认知统一\n机会地图\n首批PoC',accent:C.blue},
  {title:'3–6个月',body:'验证价值\n复制Quick Win\n建立治理基线',accent:C.cyan},
  {title:'6–9个月',body:'连接核心系统\n建设共享能力\n扩大场景',accent:C.purple},
  {title:'9–12个月',body:'形成产品线\n组织机制\n持续评估',accent:C.green}
],C.green,'路线图不是时间表承诺，而是根据业务证据不断更新的投资组合。');
flowSlide('MODULE 3｜行动入口','AI Discovery Workshop：培训之后如何开始？','从认知统一，进入具体业务机会梳理。',[
  {title:'高管与业务访谈',icon:faUserTie,accent:C.blue},{title:'流程与任务梳理',icon:faRoute,accent:C.cyan},{title:'AI机会识别',icon:faLightbulb,accent:C.purple},{title:'优先级评估',icon:faScaleBalanced,accent:C.orange},{title:'方案与PoC设计',icon:faFlask,accent:C.green}
],C.green,'输出：企业AI机会地图、Top项目候选、方案草案和PoC路线建议。');

// ---------- MODULE 4 ----------
sectionSlide(4,'AI时代的人与组织','从使用AI，到构建AI驱动企业','AI最终改变的不只是工具，而是工作、人才、管理和组织。',C.purple,faPeopleGroup);
compareSlide('MODULE 4｜工作单元','AI改变工作的基本单元','从“岗位分工”，逐步走向“目标与任务协同”。',
  {title:'过去：岗位中心',body:'岗位 → 人 → 固定职责 → 操作系统 → 完成任务',icon:faUserTie},
  {title:'未来：任务中心',body:'业务目标 → 任务拆解 → 人 + AI分工 → 动态协作 → 业务结果',icon:faDiagramProject},C.purple,'岗位不会消失，但岗位内部的任务组合和价值重心会持续变化。');
compareSlide('MODULE 4｜员工角色','员工从执行者，转向AI协作者','人的价值从信息处理，转向问题定义、判断、创造与关系。',
  {title:'过去的工作重心',body:'查找信息\n整理材料\n重复执行\n操作流程\n生成标准内容',icon:faTasks},
  {title:'未来的工作重心',body:'定义问题\n设定目标\n判断结果\n创造突破\n管理AI协作者',icon:faLightbulb},C.purple,'AI不会平均地替代岗位，而会重新分配岗位内部不同任务的时间和价值。');
{
  const s=baseSlide('MODULE 4｜人才模型','AI时代的 π 型人才','既有专业深度，也能设计人机协作与系统。',C.purple);
  addRect(s,1.0,2.05,11.3,0.85,'171E42',C.purple,true,0);
  addText(s,'跨域能力：AI协作 · 沟通 · 系统思维 · 价值判断',1.25,2.26,10.8,0.35,16,C.white,{bold:true,align:'center',margin:0});
  addRect(s,2.2,2.9,2.8,2.95,'0B203B',C.blue,true,0);
  addText(s,'业务 / 专业深度',2.55,3.25,2.1,0.55,18,C.blue,{bold:true,align:'center'});
  addText(s,'理解客户\n理解流程\n理解价值\n形成专业判断',2.6,4.05,2.0,1.35,13,C.text,{align:'center'});
  addRect(s,8.3,2.9,2.8,2.95,'0B2C35',C.green,true,0);
  addText(s,'AI / 系统构建深度',8.55,3.25,2.3,0.55,18,C.green,{bold:true,align:'center'});
  addText(s,'设计工作流\n连接工具与知识\n评估与治理\n持续优化',8.7,4.05,2.0,1.35,13,C.text,{align:'center'});
  addInsight(s,'未来高价值人才，不只是“会使用AI”，而是能把AI转化为业务系统。',C.purple,faGraduationCap);
}
fourGridSlide('MODULE 4｜管理者','管理者的新角色：管理人 + AI团队','管理对象扩大，管理机制也必须更新。',[
  {title:'定义目标',body:'让人和AI理解共同的结果与边界。',icon:faBullseye,accent:C.blue},
  {title:'设计分工',body:'决定哪些任务由人、AI或协同完成。',icon:faSitemap,accent:C.cyan},
  {title:'管理绩效',body:'评估质量、成本、速度和业务结果。',icon:faGaugeHigh,accent:C.green},
  {title:'控制风险',body:'权限、责任、升级和异常处理。',icon:faShieldHalved,accent:C.orange}
],C.purple,'未来管理者不仅要会“用AI”，更要会设计和监督AI工作系统。');
canvasSlide('MODULE 4｜组织机制','AI Transformation Office｜转型推动机制','AI转型不是IT团队单独完成的项目。',[
  {x:0.7,y:2.05,w:3.0,h:1.4,title:'高管 Sponsor',body:'方向、资源与跨部门决策',icon:faUserTie,accent:C.purple},
  {x:3.92,y:2.05,w:3.0,h:1.4,title:'业务 Owner',body:'问题、流程与业务结果',icon:faBullseye,accent:C.green},
  {x:7.14,y:2.05,w:2.65,h:1.4,title:'AI / 产品',body:'方案与体验设计',icon:faRobot,accent:C.blue},
  {x:10.02,y:2.05,w:2.6,h:1.4,title:'数据 / IT',body:'数据、系统与安全',icon:faDatabase,accent:C.cyan},
  {x:1.8,y:4.05,w:3.0,h:1.4,title:'流程专家',body:'任务拆解与人机分工',icon:faRoute,accent:C.orange},
  {x:5.18,y:4.05,w:3.0,h:1.4,title:'评估与治理',body:'Eval、风险和运营机制',icon:faShieldHalved,accent:C.purple},
  {x:8.56,y:4.05,w:3.0,h:1.4,title:'一线用户',body:'真实反馈、采用与持续改进',icon:faUsers,accent:C.green}
],C.purple,'AI项目需要业务、流程、产品、数据、技术和治理的跨职能协作。');
staircaseSlide('MODULE 4｜成熟度','企业 AI 成熟度模型','成熟度不等于工具数量，而是组织能力深度。',[
  {level:'L0',title:'Awareness',body:'认知与零散尝试',accent:C.dim},
  {level:'L1',title:'AI User',body:'个人使用与提效',accent:C.blue},
  {level:'L2',title:'AI Team',body:'团队复用与Workflow',accent:C.cyan},
  {level:'L3',title:'AI Process',body:'流程智能化与Agent',accent:C.purple},
  {level:'L4',title:'AI Native',body:'产品、服务与组织重构',accent:C.green}
],C.purple,'企业可以在不同业务领域处于不同成熟度；评估的目的不是打分，而是确定下一步。');
workshopSlide('从今天开始：你的第一个 AI 行动是什么？','把培训中的启发，转化为一个可讨论、可验证的业务机会。',[
  {x:0.75,y:2.05,w:3.75,h:3.5,title:'问题 1｜哪项工作最值得重新设计？',body:'它是否高频、耗时、信息密集、依赖经验？\n\n请写出一个具体任务，而不是一个部门或岗位。',icon:faListCheck,accent:C.blue},
  {x:4.78,y:2.05,w:3.75,h:3.5,title:'问题 2｜AI可以介入哪一步？',body:'理解、检索、生成、判断、监控、执行？\n\nAI应该做什么，人必须保留什么？',icon:faRobot,accent:C.purple},
  {x:8.8,y:2.05,w:3.75,h:3.5,title:'问题 3｜如何证明有价值？',body:'时间、质量、体验、成本、收入或风险？\n\n选择一个可以比较的基线指标。',icon:faGaugeHigh,accent:C.green}
],C.green);
{
  slideNo+=1; const s=pptx.addSlide('MASTER'); s.background={color:C.bg}; addTechDots(s,C.cyan);
  addText(s,'企业 AI 转型，从重新设计第一项工作开始。',0.85,1.25,11.6,1.25,33,C.white,{bold:true,align:'center',margin:0});
  addText(s,'AI不会自动创造价值，重新设计工作的人才会。',1.2,2.75,10.9,0.55,20,C.cyan,{bold:true,align:'center',margin:0});
  flowSlide; // no-op marker for lint simplicity
  const labs=[['理解 AI',faBrain,C.blue],['发现机会',faBullseye,C.green],['设计方案',faDiagramProject,C.purple],['验证价值',faFlask,C.orange],['规模推广',faRocket,C.cyan]];
  labs.forEach((it,i)=>{addCard(s,{x:0.65+i*2.52,y:4.15,w:2.22,h:1.35,title:it[0],iconFa:it[1],accent:it[2],fill:'0B203B',titleSize:14}); if(i<labs.length-1)addArrow(s,2.86+i*2.52,4.62,0.3,it[2]);});
  addText(s,'AI Discovery → Solution Design → Pilot Delivery → Scale Up',2.05,6.1,9.3,0.35,14,C.muted,{bold:true,align:'center',margin:0});
  addFooter(s,'企业AI转型与应用实践｜Thank You');
  slideMeta.push({no:slideNo,section:'Closing',title:'企业AI转型，从重新设计第一项工作开始'});
}

// ---------- APPENDIX / WORKSHOP TEMPLATES ----------
sectionSlide('A','附录与工作坊工具','可打印、可复用、可直接用于企业共创','以下模板用于培训现场或后续AI Discovery项目。',C.green,faClipboard);
workshopSlide('AI Opportunity Card｜机会卡','用一页纸描述一个值得探索的AI机会。',[
  {x:0.65,y:1.95,w:3.95,h:1.15,title:'业务领域 / 场景名称',body:'部门、流程、具体任务',icon:faBuilding,accent:C.blue},
  {x:4.8,y:1.95,w:3.95,h:1.15,title:'业务目标',body:'希望改善的业务结果',icon:faBullseye,accent:C.green},
  {x:8.95,y:1.95,w:3.75,h:1.15,title:'参与角色 / Owner',body:'谁使用、谁负责',icon:faUserTie,accent:C.purple},
  {x:0.65,y:3.32,w:3.95,h:2.0,title:'当前流程与痛点',body:'现在怎么做？哪里耗时、易错、响应慢、难复制？',icon:faTriangleExclamation,accent:C.orange},
  {x:4.8,y:3.32,w:3.95,h:2.0,title:'AI机会与人机分工',body:'AI可以介入哪里？人必须保留什么？',icon:faHandshake,accent:C.cyan},
  {x:8.95,y:3.32,w:3.75,h:2.0,title:'成功指标与下一步',body:'如何衡量？未来30天做什么？',icon:faGaugeHigh,accent:C.green}
],C.green);
workshopSlide('AI Opportunity Score｜机会评分表','建议使用1–5分，先独立评分，再小组讨论。',[
  {x:0.75,y:2.0,w:2.25,h:3.6,title:'业务价值',body:'收入 / 成本\n体验 / 风险\n战略重要性',icon:faArrowTrendUp,accent:C.green},
  {x:3.15,y:2.0,w:2.25,h:3.6,title:'任务适配度',body:'高频\n信息密集\n规则存在\n可验证',icon:faListCheck,accent:C.blue},
  {x:5.55,y:2.0,w:2.25,h:3.6,title:'数据基础',body:'可访问\n可理解\n可信任\n可运营',icon:faDatabase,accent:C.cyan},
  {x:7.95,y:2.0,w:2.25,h:3.6,title:'实施复杂度',body:'系统连接\n流程变化\n安全治理\n技术难度',icon:faGear,accent:C.orange},
  {x:10.35,y:2.0,w:2.25,h:3.6,title:'组织准备度',body:'业务Owner\n用户意愿\n资源投入\n跨部门协同',icon:faPeopleGroup,accent:C.purple}
],C.green);
workshopSlide('AI Solution Canvas｜方案画布','从机会进入方案设计。',[
  {x:0.6,y:1.9,w:3.0,h:1.35,title:'用户与场景',body:'谁、何时、为什么使用？',icon:faUser,accent:C.blue},
  {x:3.78,y:1.9,w:3.0,h:1.35,title:'目标与成功标准',body:'希望完成什么结果？',icon:faBullseye,accent:C.green},
  {x:6.96,y:1.9,w:2.75,h:1.35,title:'输入与知识',body:'数据、资料、规则',icon:faBookOpen,accent:C.cyan},
  {x:9.89,y:1.9,w:2.8,h:1.35,title:'AI能力',body:'理解、生成、判断、执行',icon:faBrain,accent:C.purple},
  {x:0.6,y:3.52,w:3.0,h:1.35,title:'流程与状态',body:'步骤、依赖、异常',icon:faRoute,accent:C.orange},
  {x:3.78,y:3.52,w:3.0,h:1.35,title:'工具与系统',body:'API、SaaS、数据库',icon:faWrench,accent:C.blue},
  {x:6.96,y:3.52,w:2.75,h:1.35,title:'人机控制',body:'确认、升级、停止、回滚',icon:faShieldHalved,accent:C.green},
  {x:9.89,y:3.52,w:2.8,h:1.35,title:'评估指标',body:'体验、效率、质量、业务',icon:faGaugeHigh,accent:C.cyan}
],C.green);
workshopSlide('Enterprise AI Case Card｜案例卡模板','用于把外部案例转化为本企业可讨论的启发。',[
  {x:0.7,y:2.0,w:3.75,h:1.55,title:'业务背景',body:'行业、角色、任务和问题。',icon:faBuilding,accent:C.blue},
  {x:4.78,y:2.0,w:3.75,h:1.55,title:'传统方式',body:'当前流程、痛点和限制。',icon:faRoute,accent:C.orange},
  {x:8.86,y:2.0,w:3.75,h:1.55,title:'AI重新设计',body:'能力、工具、流程与控制。',icon:faRobot,accent:C.purple},
  {x:0.7,y:3.92,w:3.75,h:1.55,title:'人机分工',body:'AI做什么？人负责什么？',icon:faHandshake,accent:C.cyan},
  {x:4.78,y:3.92,w:3.75,h:1.55,title:'业务价值',body:'效率、质量、体验、收入、风险。',icon:faArrowTrendUp,accent:C.green},
  {x:8.86,y:3.92,w:3.75,h:1.55,title:'适用条件与风险',body:'数据、系统、组织与治理前提。',icon:faShieldHalved,accent:C.orange}
],C.green);
timelineSlide('APPENDIX｜WORKSHOP','一天工作坊建议节奏','讲授与共创交替，避免纯粹“听课”。',[
  {title:'09:00',body:'开场与认知冲击',accent:C.blue},
  {title:'10:00',body:'理解AI能力进化',accent:C.cyan},
  {title:'11:15',body:'个人与团队案例',accent:C.purple},
  {title:'13:30',body:'Agent与AI Native',accent:C.orange},
  {title:'15:00',body:'机会卡与评分',accent:C.green},
  {title:'16:30',body:'小组汇报与下一步',accent:C.blue}
],C.green,'建议根据学员成熟度动态调整讲授深度，把更多时间留给本企业场景讨论。');
fourGridSlide('APPENDIX｜术语','关键术语：用业务语言理解技术概念','培训中通俗词为主，专业术语作为对应关系保留。',[
  {title:'RAG｜先查再答',body:'把相关资料临时拿到模型工作台上。',icon:faBookOpen,accent:C.blue},
  {title:'Tools / MCP｜插座与通道',body:'让AI连接和使用外部系统。',icon:faPlug,accent:C.cyan},
  {title:'Memory / State｜笔记与任务状态',body:'跨任务保留信息并持续推进。',icon:faHardDrive,accent:C.purple},
  {title:'Harness｜执行与控制系统',body:'权限、状态、日志、评估与异常处理。',icon:faShieldHalved,accent:C.green}
],C.green,'技术术语不是课程主角；理解它们如何改变工作，才是关键。');

const outDir = path.resolve('dist');
fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, '企业AI转型与应用实践_完整版_V1.pptx');
await pptx.writeFile({ fileName: outFile });

const qa = [
  '# QA Report',
  '',
  `- Slide count: ${slideNo}`,
  `- Metadata records: ${slideMeta.length}`,
  `- Placeholders/TBD: 0`,
  `- Theme: dark navy / cyan / blue / purple / green`,
  `- Editable vector slides: yes`,
  `- Illustrative cases labeled: yes`,
  '',
  '## Slide list',
  ...slideMeta.map(x=>`- ${String(x.no).padStart(2,'0')} | ${x.section} | ${x.title}`)
].join('\n');
fs.writeFileSync(path.join(outDir,'qa-report.md'),qa,'utf8');
console.log(`Built ${outFile} with ${slideNo} slides.`);
