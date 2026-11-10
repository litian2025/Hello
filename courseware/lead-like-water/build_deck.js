// 上善若水：AI 时代的柔性领导力 —— Career Up Club 柔性领导力系列课程
// Build: NODE_PATH=<node_modules with pptxgenjs sharp react react-dom react-icons> node build_deck.js [out.pptx]
// Optional: PPTX_APPLY_THEME=<path to apply_theme.js> to use an external theme writer instead of the built-in one.
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const { makeArt, iconPng } = require("./art.js");

const OUT = process.argv[2] || path.join(__dirname, "Lead-Like-Water_柔性领导力课件.pptx");
const FONT = "Microsoft YaHei";
const THEME = {
	name: "Lead Like Water",
	headFontFace: FONT,
	bodyFontFace: FONT,
	colors: {
		dk1: "0B2529", lt1: "FFFFFF", dk2: "0F4C5C", lt2: "EEF5F4",
		accent1: "1B7F8C", accent2: "C8442F", accent3: "5FB3B3", accent4: "E0A458", accent5: "5F7377", accent6: "BFE0DA",
		hlink: "1B7F8C", folHlink: "5F7377",
	},
};
const FOOT = "职场提升俱乐部 Career Up Club · 柔性领导力系列课程";

(async () => {
	const ART = await makeArt();
	const ICONS = {};
	const iconNames = ["FaBolt", "FaHeartbeat", "FaBatteryQuarter", "FaCompass", "FaQuestion", "FaLightbulb", "FaComments", "FaSearch", "FaSyncAlt", "FaBookOpen", "FaHandsHelping", "FaHeadphones", "FaHeart", "FaBalanceScale", "FaUsers", "FaPenNib", "FaExclamationTriangle", "FaRobot", "FaUserCheck", "FaBullseye", "FaProjectDiagram", "FaTint", "FaWater", "FaStream", "FaLayerGroup", "FaArrowDown", "FaEye", "FaGem"];
	for (const n of iconNames) ICONS[n] = await iconPng(n);

	const pres = new pptxgen();
	pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
	pres.title = "上善若水：AI 时代的柔性领导力";
	pres.subject = "柔性领导力系列课程";
	pres.author = "Career Up Club";
	pres.company = "Career Up Club";
	pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
	const C = pres.SchemeColor;
	const SH = pres.ShapeType;

	// ---------- layouts ----------
	const footer = (color, x = 0.6, w = 6) => ({ text: { text: FOOT, options: { x, y: 5.22, w, h: 0.25, fontSize: 9, color, margin: 0, isTextBox: true } } });
	const num = (color, x = 8.9) => ({ x, y: 5.22, w: 0.5, h: 0.25, fontSize: 9, color, align: "right", margin: 0 });
	const ph = (name, type, o) => ({ placeholder: { options: { name, type, margin: 0, align: "left", ...o }, text: "" } });

	pres.defineSlideMaster({
		title: "TITLE", background: { data: ART.bgTitle },
		objects: [
			ph("title", "title", { x: 0.7, y: 1.55, w: 6.6, h: 1.1, fontSize: 60, bold: true, color: C.background1, valign: "middle" }),
			ph("body", "body", { x: 0.7, y: 2.7, w: 7, h: 0.6, fontSize: 28, color: C.background1, valign: "middle" }),
		],
	});
	pres.defineSlideMaster({
		title: "SECTION", background: { data: ART.bgSection },
		objects: [
			ph("title", "title", { x: 0.7, y: 2.05, w: 6.6, h: 1.0, fontSize: 34, bold: true, color: C.background1, valign: "middle" }),
			ph("body", "body", { x: 0.7, y: 3.2, w: 5.6, h: 1.1, fontSize: 16, color: C.accent6, valign: "top" }),
			footer(C.accent6),
		],
		slideNumber: num(C.accent6),
	});
	pres.defineSlideMaster({
		title: "STATEMENT", background: { data: ART.bgQuote },
		objects: [
			ph("title", "title", { x: 0.8, y: 1.15, w: 8.4, h: 2.2, fontSize: 32, bold: true, color: C.background1, valign: "middle" }),
			ph("body", "body", { x: 0.8, y: 3.5, w: 8.4, h: 0.8, fontSize: 16, color: C.accent6, valign: "top" }),
			footer(C.accent6),
		],
		slideNumber: num(C.accent6),
	});
	pres.defineSlideMaster({
		title: "CONTENT", background: { color: C.background1 },
		objects: [ph("title", "title", { x: 0.6, y: 0.35, w: 8.0, h: 0.8, fontSize: 28, bold: true, color: C.text1, valign: "middle" }), footer(C.accent5)],
		slideNumber: num(C.accent5),
	});
	pres.defineSlideMaster({
		title: "IMAGE_RIGHT", background: { color: C.background1 },
		objects: [ph("title", "title", { x: 0.6, y: 0.35, w: 5.0, h: 0.8, fontSize: 26, bold: true, color: C.text1, valign: "middle" }), footer(C.accent5, 0.6, 4.4)],
		slideNumber: num(C.accent5, 5.1),
	});
	pres.defineSlideMaster({
		title: "IMAGE_LEFT", background: { color: C.background1 },
		objects: [ph("title", "title", { x: 4.5, y: 0.35, w: 4.9, h: 0.8, fontSize: 26, bold: true, color: C.text1, valign: "middle" }), footer(C.accent5, 4.5, 4.3)],
		slideNumber: num(C.accent5),
	});
	pres.defineSlideMaster({
		title: "QA", background: { data: ART.bgQa },
		objects: [ph("title", "title", { x: 1.85, y: 0.4, w: 7.55, h: 1.0, fontSize: 26, bold: true, color: C.text1, valign: "middle" }), footer(C.accent5)],
		slideNumber: num(C.accent5),
	});

	// ---------- helpers ----------
	let section = "";
	const startSection = (title) => { section = title; pres.addSection({ title }); };
	const add = (masterName, title, notes) => {
		const s = pres.addSlide({ masterName, sectionTitle: section });
		if (title) s.addText(title, { placeholder: "title" });
		s.addNotes(notes.trim());
		return s;
	};
	const T = (s, text, x, y, w, h, o = {}) => s.addText(text, { x, y, w, h, margin: 0, isTextBox: true, fontSize: 14, color: C.text1, valign: "top", ...o });
	const box = (s, x, y, w, h, color = C.background2, o = {}) => s.addShape(SH.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color }, line: { type: "none" }, ...o });
	const shadow = () => ({ type: "outer", color: "0B2529", opacity: 0.12, blur: 8, offset: 2, angle: 90 });
	const seal = (s, ch, x = 8.85, y = 0.47, size = 0.55, color = C.accent2, fontSize) =>
		s.addText(ch, { shape: SH.roundRect, rectRadius: size * 0.14, x, y, w: size, h: size, fill: { color }, line: { type: "none" }, color: C.background1, bold: true, fontSize: fontSize || Math.round(size * 40), align: "center", valign: "middle", margin: 0 });
	const dot = (s, label, x, y, d, color = C.accent1, fontSize = 16, textColor = C.background1) =>
		s.addText(label, { shape: SH.ellipse, x, y, w: d, h: d, fill: { color }, line: { type: "none" }, color: textColor, bold: true, fontSize, align: "center", valign: "middle", margin: 0 });
	const icon = (s, name, x, y, d = 0.6, color = C.accent1) => {
		s.addShape(SH.ellipse, { x, y, w: d, h: d, fill: { color }, line: { type: "none" } });
		s.addImage({ data: ICONS[name], x: x + d * 0.26, y: y + d * 0.26, w: d * 0.48, h: d * 0.48 });
	};
	// 道德经 line at the foot of a content slide.
	const tao = (s, quote, chapter, x = 0.6, w = 8.8) => {
		seal(s, "道", x, 4.6, 0.34, C.accent2, 13);
		T(s, [{ text: quote, options: { bold: true, fontSize: 14, color: C.text2 } }, { text: `　——《道德经》${chapter}`, options: { fontSize: 11, color: C.accent5 } }], x + 0.48, 4.6, w - 0.48, 0.34, { valign: "middle" });
	};
	const arrow = (s, x, y, w = 0.5, h = 0.4, color = C.accent2) => s.addShape(SH.rightArrow, { x, y, w, h, fill: { color }, line: { type: "none" } });

	// =====================================================================
	// 开场 · 引
	// =====================================================================
	startSection("开场 · 引");

	let s = add("TITLE", "上善若水", `
【开场 · 约 1 分钟】
（等大家坐定，先安静两三秒，再开口。语速放慢。）
大家好，欢迎来到职场提升俱乐部 Career Up Club 的柔性领导力系列课程。
今天这一个多小时，我们只聊一件事：在 AI 把“能力”变得越来越便宜的时代，一个人靠什么立足？
我的答案藏在四个字里——上善若水。这不是一堂讲 AI 工具的课，也不是一堂讲《道德经》的课。它讲的是：当世界变得越来越“硬”、越来越快，为什么最柔软的东西，反而最有力量。
【提示】标题页的“主讲人 · 日期”请替换为实际信息。`);
	s.addText("AI 时代的柔性领导力", { placeholder: "body" });
	T(s, "职场提升俱乐部 Career Up Club · 柔性领导力系列课程", 0.7, 0.75, 7, 0.35, { fontSize: 14, color: C.accent6 });
	T(s, "Lead Like Water — Timeless Wisdom in the AI Era", 0.7, 3.4, 7, 0.35, { fontSize: 16, color: C.accent6 });
	T(s, "[主讲人] · [日期]", 0.7, 4.75, 5, 0.3, { fontSize: 12, color: C.background1 });

	s = add("STATEMENT", "如果明天，你的岗位名称被抹掉，\n你还剩下什么？", `
【Hook · 约 2 分钟】
（念完问题，停住。真的停 10 秒，不要急着填满沉默。）
我想请大家先别回答。把名片上的头衔、工牌上的部门、简历上的岗位名称，全部拿掉。
剩下的那个人，是谁？他凭什么被需要？
（环视全场）第一次认真面对这个问题，多数人心里是会慌一下的。因为我们大多数人介绍自己，第一句话就是“我是某某公司的某某经理”。
【提示】如果你有自己被这个问题问住的真实经历，在这里用一两句话讲出来，效果更好。
今天我想带大家看到：这个问题，已经不是一个哲学问题了。它正在变成一个现实问题。`);
	s.addText("请先别回答。让这个问题，在心里停 10 秒。", { placeholder: "body" });

	s = add("IMAGE_LEFT", "凌晨一点，一个人的“团队”", `
【Story · 约 3 分钟】
（声音放低，像讲一个画面，而不是讲一个道理。）
请大家跟我想象这样一个夜晚：窗外整栋楼都黑了，只剩一扇窗还亮着。
一位产品经理独自坐在屏幕前。这三个小时里，她和 AI 一起，把需求写完了，把原型画好了，把代码跑通了，连数据分析都做出来了。放在过去，这是一个团队几周的工作量。
她合上电脑。按理说，她应该感到强大。但她没有。
她心里只有一个念头——（停顿）“这些东西，真的值得做吗？”
这个念头，就是今天这堂课的起点。AI 给了她一个团队的能力，却没有给她一个答案：什么值得做。
【重要】这是一个示意性的虚构场景，不是真实案例。强烈建议替换成你自己、或某位学员的真实经历——真实的故事永远更有力量。替换后请删去页面上的方括号提示。`);
	s.addImage({ data: ART.sceneNight, altText: "深夜的城市，只有一扇窗还亮着", x: 0, y: 0, w: 4.0, h: 5.625 });
	T(s, "想象这样一个夜晚——", 4.5, 1.3, 4.9, 0.3, { fontSize: 14, color: C.accent5 });
	T(s, "一位产品经理独自坐在屏幕前。三个小时里，她和 AI 一起写完了需求、画好了原型、跑通了代码、做出了数据分析。", 4.5, 1.7, 4.9, 0.95, { fontSize: 15, lineSpacingMultiple: 1.15 });
	T(s, "过去，这是一个团队几周的工作量。", 4.5, 2.75, 4.9, 0.3, { fontSize: 15 });
	T(s, "她合上电脑，却没有感到强大。", 4.5, 3.15, 4.9, 0.3, { fontSize: 15 });
	T(s, "她只有一个念头：\n这些东西，真的值得做吗？", 4.5, 3.6, 4.9, 0.85, { fontSize: 20, bold: true, color: C.accent2 });
	T(s, "[示意场景——请替换为主讲人或学员的真实经历]", 4.5, 4.7, 4.9, 0.25, { fontSize: 10, color: C.accent5 });

	s = add("CONTENT", "此刻，你对 AI 的感受更接近哪一种", `
【情绪带入 · 约 2 分钟】
（互动：请大家举手。每一种都问一遍，自己也举手。）
在往下讲之前，我想先问问大家的感受，而不是观点。
兴奋的举个手？焦虑的？觉得累的——工具学不完、事情做不完？还有，说不清楚、有点迷茫的？
（看一圈）谢谢大家的诚实。我想说的是：这些感受都对，而且常常同时存在。
情绪不是噪音，它是一条信息。当我们对 AI 感到焦虑的时候，水面之下通常有一个更深的问题在动——“我是谁？我还值多少？”
今天我们不绕开它。等一下讲到 G 的时候，我们会专门回到这座“冰山”。`);
	seal(s, "引");
	[["FaBolt", "兴奋", "我能做的事，\n一下子变多了", C.accent4], ["FaHeartbeat", "焦虑", "我会不会\n被替代", C.accent2], ["FaBatteryQuarter", "疲惫", "工具学不完，\n事情做不完", C.accent1], ["FaCompass", "迷茫", "不知道该往\n哪里用力", C.text2]].forEach(([ic, label, line, color], i) => {
		const x = 0.6 + i * 2.25;
		box(s, x, 1.4, 2.05, 2.35);
		icon(s, ic, x + 0.25, 1.65, 0.7, color);
		T(s, label, x + 0.25, 2.5, 1.6, 0.4, { fontSize: 20, bold: true });
		T(s, line, x + 0.25, 2.95, 1.6, 0.6, { fontSize: 14, color: C.accent5 });
	});
	T(s, "情绪没有对错。它是一条信息：有些重要的东西，正在被触动。", 0.6, 4.1, 8.8, 0.4, { fontSize: 16, bold: true, color: C.text2 });

	s = add("CONTENT", "今天的路线：一条河的旅程", `
【路线图 · 约 1 分钟】
今天的课，我把它设计成一条河的旅程，一共四段。
第一段“源”：看清变局——AI 到底改写了什么。大约 15 分钟。
第二段“流”：水之道——柔性领导力的 GIA 模型，这是今天的核心，大约 25 分钟。
第三段“势”：乘势——怎么用 ACE 让 AI 成为你的助力，大约 12 分钟。
第四段“海”：归海——怎么找到自己的价值、怎么更有价值，大约 10 分钟。
之后留 30 分钟左右给大家提问。
如果今天你只带走一句话，就是屏幕下方这一句：AI 让执行变便宜，让判断变昂贵；像水一样领导，就是判断力的来源。`);
	seal(s, "引");
	s.addShape(SH.line, { x: 0.95, y: 1.75, w: 6.75, h: 0, line: { color: C.accent6, width: 2.5 } });
	[["源", "变局", "AI 重新定义\n人的价值", "约 15 分钟"], ["流", "水之道", "柔性领导力\nGIA", "约 25 分钟"], ["势", "乘势", "与 AI 共舞的\nACE", "约 12 分钟"], ["海", "归海", "找到并放大\n你的价值", "约 10 分钟"]].forEach(([ch, title, desc, time], i) => {
		const x = 0.6 + i * 2.25;
		seal(s, ch, x, 1.4, 0.7, i === 1 ? C.accent2 : C.text2);
		T(s, title, x, 2.25, 2.05, 0.35, { fontSize: 18, bold: true });
		T(s, desc, x, 2.65, 2.05, 0.55, { fontSize: 14, color: C.accent5 });
		T(s, time, x, 3.25, 2.05, 0.25, { fontSize: 11, color: C.accent1, bold: true });
	});
	box(s, 0.6, 3.75, 8.8, 1.15);
	T(s, "一句话带走", 0.85, 3.88, 3, 0.25, { fontSize: 12, bold: true, color: C.accent2 });
	T(s, "AI 让执行变便宜，让判断变昂贵；像水一样领导，是判断力的来源。", 0.85, 4.15, 8.3, 0.35, { fontSize: 16, bold: true, color: C.text2 });
	T(s, "之后：Q&A 约 30 分钟", 0.85, 4.55, 8.3, 0.25, { fontSize: 11, color: C.accent5 });

	// =====================================================================
	// 第一部分 · 源
	// =====================================================================
	startSection("第一部分 · 源：变局");

	s = add("SECTION", "变局：岗位，从职责变成标签", `
【过场 · 约 30 秒】
我们从源头开始。
过去，岗位是一个人的职责；未来，岗位很可能只是一个人的标签。
这句话听起来有点刺耳，但接下来的十几分钟，我想带大家看清楚它背后正在发生的事。`);
	s.addText("过去，岗位是一个人的职责；\n未来，岗位很可能只是一个人的标签。", { placeholder: "body" });
	T(s, "第一部分", 0.7, 1.3, 3, 0.3, { fontSize: 14, color: C.accent6 });
	seal(s, "源", 0.7, 0.55, 0.65);

	s = add("CONTENT", "来自 Google Research 的信号", `
【建立可信度 · 约 2 分钟】
Google 副总裁、Google Research 负责人 Yossi Matias（约西·马蒂亚斯）在一场访谈里，谈到了 AI 时代的人才规则。围绕这场访谈有一个很好的解读，它提出了四个“为什么”。
（逐条念右侧四个问题，念完第四个停一下。）
第四个最反常识：能力越强，反而越危险。
左边这句话是今天第一部分的核心：AI 真正改写的，不是一个人能做多少工作，而是个人和组织之间，长久以来的能力分配规则。
这是一场比“职业被替代”更深刻的变化。
【提示】页面下方的访谈链接请补上，方便学员课后查证。`);
	seal(s, "源");
	box(s, 0.6, 1.35, 3.7, 2.9, C.text2);
	T(s, "被改写的，不是一个人能做多少工作，", 0.85, 1.6, 3.2, 0.9, { fontSize: 18, color: C.accent6 });
	T(s, "而是个人与组织之间的能力分配规则。", 0.85, 2.6, 3.2, 1.4, { fontSize: 22, bold: true, color: C.background1 });
	["为什么最优秀的人，不再被任何一个岗位定义？", "为什么团队的工作，正在变成一个人的能力？", "为什么“软技能”，会成为最重要的硬实力？", "为什么能力越强，反而越危险？"].forEach((q, i) => {
		const y = 1.35 + i * 0.76;
		box(s, 4.6, y, 4.8, 0.62);
		dot(s, String(i + 1), 4.72, y + 0.11, 0.4, i === 3 ? C.accent2 : C.accent1, 14);
		T(s, q, 5.27, y, 4.05, 0.62, { fontSize: 14, valign: "middle", bold: i === 3 });
	});
	T(s, "来源：Google 副总裁、Google Research 负责人 Yossi Matias 的访谈及相关解读　[待补充：访谈链接]", 0.6, 4.6, 8.8, 0.25, { fontSize: 10, color: C.accent5 });

	s = add("CONTENT", "变化一：岗位与能力，正在脱钩", `
【约 2 分钟】
先看第一个变化。现代企业的组织方式，本质上是把复杂的能力拆到不同的岗位里，再靠分工和协作把它们拼回来。
所以过去，组织靠招聘获得能力，靠分工组织能力，靠管理协调能力。岗位的价值，很大程度来自它代表的专业稀缺性。
AI 正在打破这种绑定。一个产品经理可以直接构建软件；一个设计师可以参与技术开发；一个行业专家可以调用 AI，完成过去需要好几个专业岗位配合的研究、分析和验证。
注意，重点不是“一个人能替代多少人”。重点是：过去必须依赖组织才能获得的能力，正在变成个人可以直接调用的资源。
所以——岗位数量不再等于能力数量，团队规模也不再等于价值创造的规模。`);
	seal(s, "源");
	box(s, 0.6, 1.35, 3.9, 2.55);
	box(s, 5.5, 1.35, 3.9, 2.55, C.text2);
	T(s, "过去 · 靠组织", 0.9, 1.55, 3.3, 0.35, { fontSize: 18, bold: true });
	T(s, "现在 · 靠个体 + AI", 5.8, 1.55, 3.3, 0.35, { fontSize: 18, bold: true, color: C.background1 });
	[["招聘", "获得能力", "AI", "获得能力"], ["分工", "组织能力", "Agent", "组织能力"], ["管理", "协调能力", "智能系统", "协同能力"]].forEach(([a, b, c, d], i) => {
		const y = 2.1 + i * 0.55;
		T(s, [{ text: a, options: { bold: true, color: C.accent1 } }, { text: "　→　" + b }], 0.9, y, 3.3, 0.4, { fontSize: 16, valign: "middle" });
		T(s, [{ text: c, options: { bold: true, color: C.accent4 } }, { text: "　→　" + d }], 5.8, y, 3.3, 0.4, { fontSize: 16, valign: "middle", color: C.background1 });
	});
	arrow(s, 4.7, 2.4, 0.6, 0.45);
	T(s, "岗位数量 ≠ 能力数量　　　团队规模 ≠ 价值创造规模", 0.6, 4.2, 8.8, 0.45, { fontSize: 18, bold: true, color: C.text2, align: "center", valign: "middle" });

	s = add("CONTENT", "能力的组织权，正在流向个体", `
【约 2 分钟】
这里还有更深的一层。过去，你必须先进入某个组织，才能获得资源和协作能力。未来，你可能先拥有一部分“组织级”的能力，再决定和谁合作。
这不只是效率的提升，而是能力的组织权，开始向个体转移。
但是请注意——这并不意味着专业深度不重要了。恰恰相反。马蒂亚斯特别强调：AI 能提供知识，不等于你拥有判断知识的能力；能生成方案，不等于能识别缺陷。
所以未来真正有优势的，不是“什么都会一点”的全能型人才，而是右边这个 T：一竖是专业深度，它决定你判断的质量；一横是跨领域的组织能力，它决定你创造价值的范围。
《道德经》说，合抱之木，生于毫末。那一竖，没有捷径。`);
	seal(s, "源");
	[["过去", "先进入组织", "才获得资源与协作能力", C.background2, C.text1, C.accent5], ["未来", "先拥有组织级能力", "再决定与谁合作", C.text2, C.background1, C.accent6]].forEach(([tag, a, b, fill, fg, sub], i) => {
		const y = 1.4 + i * 1.5;
		box(s, 0.6, y, 4.2, 1.3, fill);
		T(s, tag, 0.85, y + 0.15, 1, 0.25, { fontSize: 12, bold: true, color: sub });
		T(s, a, 0.85, y + 0.42, 3.7, 0.35, { fontSize: 18, bold: true, color: fg });
		T(s, "→ " + b, 0.85, y + 0.82, 3.7, 0.3, { fontSize: 14, color: fg });
	});
	s.addText("跨领域组织能力　·　决定价值创造的范围", { shape: SH.roundRect, rectRadius: 0.08, x: 5.2, y: 1.4, w: 4.2, h: 0.75, fill: { color: C.accent1 }, line: { type: "none" }, color: C.background1, bold: true, fontSize: 14, align: "center", valign: "middle", margin: 0 });
	s.addText("专\n业\n深\n度", { shape: SH.rect, x: 6.8, y: 2.15, w: 1.0, h: 2.15, fill: { color: C.text2 }, line: { type: "none" }, color: C.background1, bold: true, fontSize: 18, align: "center", valign: "middle", margin: 0 });
	T(s, "挖得深，\n才看得准", 5.2, 2.9, 1.45, 0.6, { fontSize: 14, color: C.accent5, align: "right" });
	T(s, "决定\n判断的质量", 7.95, 2.9, 1.45, 0.6, { fontSize: 14, bold: true, color: C.text2 });
	tao(s, "合抱之木，生于毫末；九层之台，起于累土。", "第六十四章");

	s = add("CONTENT", "变化二：软技能，正在成为硬实力", `
【约 2 分钟】
当专业能力越来越容易被调用，第二个变化就出现了：过去被叫做“软技能”的东西，正在变成硬实力。
马蒂亚斯谈到 Google 未来的人才标准时，重点并不是你会多少种编程语言、熟悉多少 AI 工具。他强调的是右边这五件事：能不能提出正确的问题，识别重要的机会，倾听和沟通，研究和判断，持续适应变化。
为什么？因为人才价值的“计量单位”变了。过去看的是：你能完成什么任务。未来看的是：你能定义什么问题、组织什么能力、交付什么结果。
一个会编程的人，未必知道该开发什么产品。一个能管几十个 Agent 的人，未必知道该解决什么问题。
请记住下面这句话：“能够完成”和“值得完成”，是两种完全不同的能力。`);
	seal(s, "源");
	T(s, "人才价值的计量单位，变了", 0.6, 1.3, 4.2, 0.3, { fontSize: 14, color: C.accent5 });
	box(s, 0.6, 1.7, 4.2, 0.8);
	T(s, [{ text: "过去　", options: { fontSize: 12, color: C.accent5, bold: true } }, { text: "你能完成什么任务", options: { fontSize: 16, bold: true } }], 0.85, 1.7, 3.7, 0.8, { valign: "middle" });
	icon(s, "FaArrowDown", 2.5, 2.58, 0.4, C.accent2);
	box(s, 0.6, 3.06, 4.2, 1.3, C.text2);
	T(s, "未来", 0.85, 3.2, 1, 0.25, { fontSize: 12, color: C.accent6, bold: true });
	T(s, "定义什么问题 · 组织什么能力\n交付什么结果", 0.85, 3.5, 3.7, 0.7, { fontSize: 16, bold: true, color: C.background1 });
	T(s, "Matias 谈 Google 的人才标准", 5.2, 1.3, 4.2, 0.3, { fontSize: 14, color: C.accent5 });
	[["FaQuestion", "提出正确的问题"], ["FaLightbulb", "识别重要的机会"], ["FaComments", "倾听与沟通"], ["FaSearch", "研究与判断"], ["FaSyncAlt", "持续适应变化"]].forEach(([ic, label], i) => {
		const y = 1.7 + i * 0.55;
		icon(s, ic, 5.2, y, 0.42);
		T(s, label, 5.8, y, 3.6, 0.42, { fontSize: 16, valign: "middle" });
	});
	T(s, "“能够完成”和“值得完成”，是两种完全不同的能力。", 0.6, 4.6, 8.8, 0.35, { fontSize: 16, bold: true, color: C.accent2 });

	s = add("STATEMENT", "AI 越强，\n执行越便宜，判断越昂贵。", `
【金句 · 约 1 分钟】
（慢慢念，一个短句一停。）
AI 越强，执行越便宜，判断越昂贵。
过去，一个优秀的员工，把分配给他的任务完成就够了。未来，一个优秀的人还要判断四件事：这个任务合不合理？能不能交给 AI？怎么验证？最后创造了什么价值？
这就是从“任务执行者”到“结果负责人”的转变。
当执行可以被大规模调用，真正稀缺的，就不再是完成任务的人，而是能为方向和结果负责的人。
（停一下）好，道理讲到这儿有点抽象。我给大家讲一个真实的故事。`);
	s.addText("从“任务执行者”，到“结果负责人”。", { placeholder: "body" });

	s = add("IMAGE_RIGHT", "尼日利亚：洪水到来之前", `
【真实案例 · 约 3 分钟】
（讲故事的语气，带画面。）
马蒂亚斯在访谈里分享了 Google 洪水预测的案例。
过去，洪水来了以后，人们才知道哪里受灾。后来有了预测模型，人们可以提前判断洪水可能发生在哪里。
但真正的突破，不在预测本身。他提到，尼日利亚政府曾经利用 Google 的洪水预测信息，在洪水到来之前，提前向可能受影响的居民发放资金，帮他们做预防。
请大家想象那个画面：河水还没有涨起来，天还是晴的，一家人已经拿到了钱，可以加固房子、转移粮食、把孩子送到高处。
这里发生了一次关键的跃迁。预测洪水，是技术能力。识别谁有风险，是判断能力。提前组织资源、改变行动，才是价值。
如果一个预测系统准得不得了，却没有人据此行动，它的价值就很有限。
所以，不要只问 AI 完成了什么，要问 AI 改变了什么。
【提示】案例细节来自访谈解读，建议课前核对访谈原文。`);
	s.addImage({ data: ART.sceneFlood, altText: "黎明时分，河流穿过山谷，岸边有几户亮灯的人家", x: 6.0, y: 0, w: 4.0, h: 5.625 });
	[["预测", "模型提前判断洪水会发生在哪里", "技术"], ["判断", "识别哪些居民将受到影响", "判断"], ["行动", "洪水到来之前，提前发放资金", "组织"], ["结果", "人们有时间采取预防措施", "价值"]].forEach(([label, desc], i) => {
		const y = 1.3 + i * 0.7;
		dot(s, String(i + 1), 0.6, y + 0.05, 0.5, i === 3 ? C.accent2 : C.accent1, 16);
		T(s, label, 1.25, y, 0.75, 0.6, { fontSize: 16, bold: true, valign: "middle" });
		T(s, desc, 2.0, y, 3.6, 0.6, { fontSize: 14, valign: "middle" });
	});
	T(s, "预测本身不创造价值。\n改变行动，才创造价值。", 0.6, 4.1, 5.0, 0.65, { fontSize: 16, bold: true, color: C.accent2 });
	T(s, "案例来源：Yossi Matias 访谈中分享的 Google 洪水预测", 0.6, 4.85, 5.0, 0.25, { fontSize: 10, color: C.accent5 });

	s = add("CONTENT", "错误的胜利：赢了执行，输了选择", `
【Conflict · 约 3 分钟】
（语气转沉。这是全场的冲突点。）
洪水的故事是光明的一面。但它也有一个阴暗的镜像。
AI 让越来越多的人拥有了过去难以想象的执行力，却不会自动送给每个人与之匹配的判断力。结果是：错误的目标，也可能被近乎完美地执行。
看这三个不等号。
一个企业成功自动化了几十套流程，却没意识到其中一些流程根本就不该存在。
一个创业团队借助 AI 迅速做出了完整的产品，却始终没找到一个真正愿意付费的客户。
一个管理者让大量 Agent 同时工作，却没发现整个系统在高效地执行一项没有价值的任务。
这不是技术的失败。恰恰相反，这是技术成功之后，暴露出来的判断失败。
过去，错误的想法可能因为执行成本太高而“胎死腹中”；未来，错误的想法反而因为执行太容易，被充分地实现。
这就叫——错误的胜利。你赢了执行，却输了选择。
（问全场）大家身边，有没有见过这样的“胜利”？`);
	seal(s, "源");
	[["任务完成率", "价值创造率", "自动化了几十套流程——\n其中一些，根本不该存在"], ["技术成功", "商业成功", "迅速做出完整产品——\n却找不到愿意付费的客户"], ["执行正确", "方向正确", "大量 Agent 同时工作——\n高效执行一件没价值的事"]].forEach(([a, b, eg], i) => {
		const x = 0.6 + i * 3.0;
		box(s, x, 1.35, 2.8, 2.6);
		T(s, a, x, 1.5, 2.8, 0.35, { fontSize: 18, bold: true, align: "center", valign: "middle" });
		T(s, "≠", x, 1.85, 2.8, 0.4, { fontSize: 26, bold: true, color: C.accent2, align: "center", valign: "middle" });
		T(s, b, x, 2.25, 2.8, 0.35, { fontSize: 18, bold: true, align: "center", valign: "middle" });
		T(s, eg, x + 0.2, 2.85, 2.4, 0.9, { fontSize: 14, color: C.accent5 });
	});
	T(s, "这不是技术的失败，而是技术成功之后，暴露出的判断失败。", 0.6, 4.25, 8.8, 0.4, { fontSize: 16, bold: true, color: C.text2 });

	s = add("CONTENT", "更反常识：AI 不会让我们更清闲", `
【约 2 分钟】
还有一个更反常识的结果：AI 可能不会让我们越来越清闲。
马蒂亚斯预计，未来人们可能还是和今天一样忙。因为能力越强，人就越会抬高自己的目标。一个团队能做更多产品了，它往往不会选择少干点，而是去开拓更多市场。一个人能同时调用很多 Agent 了，他的任务范围也会不断扩大。
生产力的增长，不一定带来闲暇，也可能带来目标的膨胀。
（指右边）而目标一膨胀，另一个问题就来了：我们能做的事越来越多，但真正值得做的事，并不会因此自动增加。
能试 10 个方向，不代表 10 个方向都值得试。能解决 100 个问题，不代表 100 个问题都值得解决。
过去，我们受限于能力不足；未来，我们的挑战是选择过剩。当执行越来越容易，“不做什么”和“做什么”同样重要。
两千多年前，老子给过四个字：知止不殆。知道在哪里停下，才不会陷入危险。`);
	seal(s, "源");
	box(s, 0.6, 1.35, 4.2, 2.9);
	T(s, "目标膨胀", 0.85, 1.5, 3.7, 0.35, { fontSize: 18, bold: true });
	["能力越强", "目标越高", "任务越多"].forEach((t, i) => {
		const x = 0.85 + i * 1.3;
		s.addText(t, { shape: SH.roundRect, rectRadius: 0.06, x, y: 2.1, w: 1.08, h: 0.5, fill: { color: i === 2 ? C.accent2 : C.accent1 }, line: { type: "none" }, color: C.background1, bold: true, fontSize: 14, align: "center", valign: "middle", margin: 0 });
		if (i < 2) T(s, "→", x + 1.08, 2.1, 0.22, 0.5, { fontSize: 14, bold: true, color: C.accent5, align: "center", valign: "middle" });
	});
	T(s, "生产力的增长，不一定带来闲暇，\n也可能带来目标的膨胀。", 0.85, 2.95, 3.7, 0.7, { fontSize: 14, color: C.accent5 });
	box(s, 5.2, 1.35, 4.2, 2.9);
	T(s, "选择过剩", 5.45, 1.5, 3.7, 0.35, { fontSize: 18, bold: true });
	[["10", "个方向可以试\n≠ 10 个都值得试"], ["100", "个问题能解决\n≠ 100 个都值得解决"]].forEach(([n, t], i) => {
		const y = 2.05 + i * 1.05;
		T(s, n, 5.35, y, 1.25, 0.8, { fontSize: 40, bold: true, color: C.accent2, align: "right", valign: "middle" });
		T(s, t, 6.75, y, 2.5, 0.8, { fontSize: 14, valign: "middle" });
	});
	tao(s, "知足不辱，知止不殆，可以长久。", "第四十四章");

	s = add("CONTENT", "五个变化，指向同一个结论", `
【小结 · 约 2 分钟】
我们把第一部分收一下。这场访谈揭示的，其实是五个相互连接的变化。
专业能力，从绑定在岗位上，到从岗位中释放出来。
组织能力，从属于管理体系，到向个体扩散。
智能，从一个个独立的工具，变成默认的基础设施——马蒂亚斯称之为“环境智能”，等会儿我们还会回到这个词。
人才价值，从完成任务，转向负责结果。
竞争优势，从执行能力，转向价值判断。
五条线指向同一个结论：我们正在从“能力稀缺”的时代，走进“判断稀缺”的时代。
AI 正在重新定义岗位。但更深刻的是，它在重新定义人的价值。`);
	seal(s, "源");
	[["专业能力", "绑定在岗位上", "从岗位中释放"], ["组织能力", "属于管理体系", "向个体扩散"], ["智　　能", "独立的工具", "默认的基础设施"], ["人才价值", "完成任务", "负责结果"], ["竞争优势", "执行能力", "价值判断"]].forEach(([k, from, to], i) => {
		const y = 1.35 + i * 0.6;
		box(s, 0.6, y, 5.6, 0.5);
		T(s, k, 0.8, y, 1.1, 0.5, { fontSize: 14, bold: true, color: C.accent1, valign: "middle" });
		T(s, from, 1.95, y, 1.7, 0.5, { fontSize: 14, color: C.accent5, valign: "middle" });
		T(s, "→", 3.65, y, 0.35, 0.5, { fontSize: 14, bold: true, color: C.accent2, align: "center", valign: "middle" });
		T(s, to, 4.05, y, 2.1, 0.5, { fontSize: 14, bold: true, valign: "middle" });
	});
	box(s, 6.5, 1.35, 2.9, 2.9, C.text2);
	T(s, "从", 6.75, 1.6, 2.4, 0.3, { fontSize: 14, color: C.accent6 });
	T(s, "能力稀缺", 6.75, 1.95, 2.4, 0.5, { fontSize: 26, bold: true, color: C.background1 });
	T(s, "到", 6.75, 2.7, 2.4, 0.3, { fontSize: 14, color: C.accent6 });
	T(s, "判断稀缺", 6.75, 3.05, 2.4, 0.5, { fontSize: 26, bold: true, color: C.accent4 });
	T(s, "来源：Yossi Matias 访谈及相关解读", 0.6, 4.6, 8.8, 0.25, { fontSize: 10, color: C.accent5 });

	s = add("STATEMENT", "未来最危险的，不是没有能力实现想法，\n而是能力强大，\n却不知道哪些想法不值得实现。", `
【转折 · 约 1 分钟】
（一字一句。念完停 5 秒。）
未来最危险的，不是你没有能力实现自己的想法，而是你拥有了强大的能力，却不知道哪些想法根本不值得实现。
AI 会放大正确的选择，也会放大错误的方向。
那么问题来了：判断力，从哪里来？
它不在更新的模型里，不在更贵的工具里。
（放慢）我想带大家去一个两千多年前的地方找答案。`);
	s.addText("那么，判断力从哪里来？——答案，在两千多年前。", { placeholder: "body" });

	// =====================================================================
	// 第二部分 · 流
	// =====================================================================
	startSection("第二部分 · 流：柔性领导力 GIA");

	s = add("SECTION", "水之道：柔性领导力 GIA", `
【过场 · 约 1 分钟】
（慢慢念这句话，像念诗。）
上善若水。水善利万物而不争，处众人之所恶，故几于道。
最高的善，像水一样。水滋养万物，却不和万物相争；它愿意待在大家都不愿去的低处，所以最接近“道”。
老子用八十一章讲“道”，最爱用的比喻就是水。
为什么是水？为什么在 AI 时代，我们要向水学习领导力？我先讲一个大家从小就听过的故事。`);
	s.addText("上善若水。水善利万物而不争，\n处众人之所恶，故几于道。\n——《道德经》第八章", { placeholder: "body" });
	T(s, "第二部分", 0.7, 1.3, 3, 0.3, { fontSize: 14, color: C.accent6 });
	seal(s, "流", 0.7, 0.55, 0.65);

	s = add("IMAGE_LEFT", "同一场洪水，两种领导力", `
【Story · 约 3 分钟】
（带画面感：天地之间一片汪洋。）
四千多年前，洪水滔天。尧帝派鲧去治水。
鲧的办法是“堵”：水来了，就筑堤，水涨一尺，堤高一丈。他很努力，据《史记》记载，他干了九年——“九年而水不息”。堤越筑越高，水越积越大，最后失败了。
接手的是他的儿子，禹。禹没有急着动工，他先做了一件事：沿着山川走，去看水想往哪里流。然后他换了一个办法——“疏”。顺着水性，开河道，把水引向大海。《史记》说他“居外十三年，过家门不敢入”。最后，水归江海，天下安定。
同一场洪水，两种做法。
鲧在问：我怎么控制它？这是刚性的领导力。
禹在问：它想往哪里流？这是柔性的领导力。
（停顿）大家有没有发现，这和刚才尼日利亚的故事是一回事？他们没有去筑一道更高的墙，而是让资源提前流到了需要的地方。
面对 AI 这场“大水”，我们是在做鲧，还是在做禹？`);
	s.addImage({ data: ART.sceneYu, altText: "水墨山水，一条河从群山间流出", x: 0, y: 0, w: 4.0, h: 5.625 });
	[["堵", "鲧", "筑堤围堵，\n与水对抗", "九年，水患不息", C.accent5], ["疏", "禹", "顺应水性，\n疏通河道", "十三年，水归江海", C.accent1]].forEach(([ch, who, how, result, color], i) => {
		const x = 4.5 + i * 2.55;
		box(s, x, 1.4, 2.35, 1.95);
		dot(s, ch, x + 0.2, 1.58, 0.62, color, 20);
		T(s, who, x + 0.98, 1.58, 1.2, 0.62, { fontSize: 20, bold: true, valign: "middle" });
		T(s, how, x + 0.2, 2.32, 1.95, 0.55, { fontSize: 14 });
		T(s, result, x + 0.2, 2.92, 1.95, 0.3, { fontSize: 12, bold: true, color });
	});
	T(s, "刚性领导力问：我怎么控制它？", 4.5, 3.6, 4.9, 0.35, { fontSize: 16, color: C.accent5 });
	T(s, "柔性领导力问：它想往哪里流？", 4.5, 4.0, 4.9, 0.35, { fontSize: 16, bold: true, color: C.accent2 });
	T(s, "典出《史记·夏本纪》", 4.5, 4.75, 4.9, 0.25, { fontSize: 10, color: C.accent5 });

	s = add("CONTENT", "能力是水，判断是河道", `
【核心比喻 · 约 2 分钟】
（这是全场最重要的一个比喻，讲慢一点。）
AI 给了我们每个人一场大水。
水本身没有好坏。没有河道的大水，叫洪水；有了河道的大水，叫江河。
凌晨一点的那位产品经理，她拥有的就是一场大水。她害怕的，是没有河道。
所以今天这堂课的两个模型，就是这两样东西：
水，是能力——AI 带来的执行力，等会儿我们用 ACE 来讲。
河道，是判断——你的方向和定力，我们用 GIA 来讲。
柔性领导力，说到底，就是为强大的能力，修一条河道。
老子说：坚强者死之徒，柔弱者生之徒。僵硬的东西属于死亡，柔软的东西属于生长。越是变化剧烈的时代，越要柔软。`);
	seal(s, "流");
	T(s, "AI 给了每个人\n一场大水。", 0.6, 1.4, 4.6, 1.15, { fontSize: 30, bold: true, color: C.text2 });
	T(s, "没有河道的大水，叫洪水。", 0.6, 2.85, 4.6, 0.4, { fontSize: 18, color: C.accent5 });
	T(s, "有了河道的大水，叫江河。", 0.6, 3.35, 4.6, 0.4, { fontSize: 18, bold: true, color: C.accent1 });
	[["FaWater", "水 ＝ 能力", "AI 带来的执行力　→　ACE", C.background2, C.text1, C.accent5, C.accent1], ["FaStream", "河道 ＝ 判断", "你的方向与定力　→　GIA", C.text2, C.background1, C.accent6, C.accent2]].forEach(([ic, a, b, fill, fg, sub, icColor], i) => {
		const y = 1.4 + i * 1.5;
		box(s, 5.5, y, 3.9, 1.3, fill);
		icon(s, ic, 5.75, y + 0.33, 0.64, icColor);
		T(s, a, 6.6, y + 0.25, 2.7, 0.4, { fontSize: 20, bold: true, color: fg });
		T(s, b, 6.6, y + 0.72, 2.7, 0.3, { fontSize: 14, color: sub });
	});
	tao(s, "坚强者死之徒，柔弱者生之徒。", "第七十六章");

	s = add("CONTENT", "柔性领导力 GIA：三圈涟漪", `
【模型总览 · 约 2 分钟】
柔性领导力是什么？我把它叫做 GIA，三个字母，三圈涟漪。
想象一滴水落进湖里。
最里面一圈，G，Ground Yourself，向内扎根。先领导自己：我是谁，我凭什么做判断。
第二圈，I，Influence Others，以柔化人。不靠职位，怎么让别人愿意和你同行。
最外面一圈，A，Align the System，顺势成事。让人、AI 和目标，流向同一个方向。
顺序很重要：涟漪永远是从中心往外扩的。自己不定，影响不了别人；影响不了人，就对齐不了系统。
《道德经》第五十四章讲的正是这个次序：修之于身，其德乃真；一路修到家、乡、邦，最后修之于天下，其德乃普。
【提示】G 的英文在需求里写作“grand your self”，这里按 Ground Yourself（扎根自己）理解，如有不同请修改。`);
	seal(s, "流");
	s.addShape(SH.ellipse, { x: 1.2, y: 1.42, w: 2.95, h: 2.95, fill: { color: C.accent6 }, line: { type: "none" } });
	s.addShape(SH.ellipse, { x: 1.65, y: 1.87, w: 2.05, h: 2.05, fill: { color: C.accent3 }, line: { type: "none" } });
	s.addShape(SH.ellipse, { x: 2.1, y: 2.32, w: 1.15, h: 1.15, fill: { color: C.text2 }, line: { type: "none" } });
	T(s, "A", 2.375, 1.45, 0.6, 0.42, { fontSize: 20, bold: true, color: C.text2, align: "center", valign: "middle" });
	T(s, "I", 2.375, 1.9, 0.6, 0.42, { fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle" });
	T(s, "G", 2.375, 2.645, 0.6, 0.5, { fontSize: 28, bold: true, color: C.background1, align: "center", valign: "middle" });
	[["G", "Ground Yourself · 向内扎根", "我是谁，我凭什么判断", C.text2], ["I", "Influence Others · 以柔化人", "不靠职位，如何让人愿意同行", C.accent1], ["A", "Align the System · 顺势成事", "让人、AI 与目标流向同一个方向", C.accent3]].forEach(([l, title, desc, color], i) => {
		const y = 1.42 + i * 1.02;
		seal(s, l, 4.9, y + 0.08, 0.68, color, 24);
		T(s, title, 5.8, y + 0.05, 3.6, 0.38, { fontSize: 16, bold: true, valign: "middle" });
		T(s, desc, 5.8, y + 0.45, 3.6, 0.32, { fontSize: 14, color: C.accent5 });
	});
	tao(s, "修之于身，其德乃真；……修之于天下，其德乃普。", "第五十四章");

	s = add("CONTENT", "G · Ground Yourself：向内扎根", `
【G · 约 3 分钟】
先讲最里面的一圈，G。
很多人以为，判断力是分析出来的：信息更多、模型更好，判断就更准。但你想想自己最糟糕的那些决定——是因为信息不够吗？更多时候，是因为当时心是乱的。
判断力从一个安定的内在里长出来。所以 G 有三个字。
第一个字，定。在信息的洪流里，先让自己沉下来，再做决定。老子说：致虚极，守静笃。
第二个字，明。看清自己——你的价值观是什么，优势在哪里，盲区在哪里。知人者智，自知者明。了解别人是聪明，了解自己才是真正的明白。
第三个字，止。想清楚什么不做。前面讲过，AI 时代的难题是选择过剩。知止不殆——知道在哪里停，才能走得长久。
G 的产出不是一种心情，而是一份属于你自己的判断标准。`);
	seal(s, "流");
	T(s, "判断力不是算出来的——它从一个安定的内在里长出来。", 0.6, 1.25, 8.8, 0.35, { fontSize: 15, color: C.text2 });
	[["定", "静下来", "信息洪流里，先让自己沉淀，再做决定", "致虚极，守静笃", "第十六章"], ["明", "看清楚", "看清自己的价值观、优势与盲区", "知人者智，自知者明", "第三十三章"], ["止", "划边界", "想清楚什么不做，比多做一件事更难", "知止不殆，可以长久", "第四十四章"]].forEach(([ch, label, desc, q, chap], i) => {
		const x = 0.6 + i * 3.0;
		box(s, x, 1.8, 2.8, 2.55);
		dot(s, ch, x + 0.25, 1.98, 0.72, C.accent1, 24);
		T(s, label, x + 1.12, 1.98, 1.5, 0.72, { fontSize: 20, bold: true, valign: "middle" });
		T(s, desc, x + 0.25, 2.88, 2.3, 0.6, { fontSize: 14 });
		T(s, [{ text: q, options: { bold: true, color: C.accent2, fontSize: 13 } }, { text: "\n《道德经》" + chap, options: { color: C.accent5, fontSize: 10 } }], x + 0.25, 3.62, 2.3, 0.55);
	});
	T(s, "G 的产出：一份属于你自己的判断标准。", 0.6, 4.6, 8.8, 0.35, { fontSize: 16, bold: true, color: C.text2 });

	s = add("CONTENT", "看见水面之下：情绪与潜意识", `
【情绪与潜意识 · 约 3 分钟】
（回扣开场的情绪调查。）
还记得开场时，很多人举了“焦虑”那只手吗？我们来看看这座冰山。
水面之上，是我们看得见的行为：拼命学工具，报很多课，不敢停下来。
水面之下第一层，是情绪：焦虑，怕落后。
再往下，是信念：一句我们从来没有说出口、却一直相信的话——“我的价值，等于我的岗位”。
最深处，是身份，是潜意识里的那个问题：我是谁？我够好吗？
所以 AI 带来的焦虑，真正的来源往往不是技术，而是“我是谁”被动摇了。只在水面上用力——学更多工具——是解决不了水面下的问题的。
怎么办？三个字：停、名、问。
停：情绪上来时，先停三次呼吸，不急着反应。
名：给它起个名字。心里说一句“哦，我在害怕被替代”。被命名的情绪，就不再控制你。
问：它在保护我的什么？它想提醒我什么？
老子说：胜人者有力，自胜者强。战胜别人是有力气，战胜自己才是真正的强大。`);
	seal(s, "流");
	s.addImage({ data: ART.iceberg, altText: "冰山：水面之上一小部分，水面之下一大部分", x: 0.6, y: 1.35, w: 4.6, h: 3.15 });
	T(s, [{ text: "行为　", options: { bold: true } }, { text: "拼命学工具，不敢停" }], 0.8, 1.72, 2.5, 0.3, { fontSize: 12, valign: "middle" });
	[["情绪", "焦虑，怕落后"], ["信念", "我的价值＝我的岗位"], ["身份", "我是谁？我够好吗？"]].forEach(([k, v], i) => {
		T(s, [{ text: k + "　", options: { bold: true, color: C.accent4 } }, { text: v }], 0.8, 2.55 + i * 0.6, 2.5, 0.3, { fontSize: 12, color: C.background1, valign: "middle" });
	});
	[["停", "Pause", "情绪上来时，先停三次呼吸"], ["名", "Name", "给它一个名字：“我在害怕被替代”"], ["问", "Ask", "它在保护什么？想提醒我什么？"]].forEach(([ch, en, desc], i) => {
		const y = 1.4 + i * 1.02;
		dot(s, ch, 5.55, y + 0.05, 0.66, C.accent1, 22);
		T(s, [{ text: ch + " ", options: { bold: true, fontSize: 16 } }, { text: en, options: { fontSize: 12, color: C.accent5 } }], 6.4, y, 3.0, 0.32, { valign: "middle" });
		T(s, desc, 6.4, y + 0.34, 3.0, 0.52, { fontSize: 14, color: C.accent5 });
	});
	tao(s, "胜人者有力，自胜者强。", "第三十三章");

	s = add("CONTENT", "一分钟练习：让浊水变清", `
【体验练习 · 约 3 分钟】
（现场带做。声音放轻、放慢，每句之间留白。可以关掉翻页，只看着大家。）
我们现在一起做一分钟的练习。请大家坐稳，双脚踩地，愿意的话闭上眼睛。
想象你手里端着一杯水。这杯水刚刚被搅过，是浑的——里面是今天所有的消息、待办、别人的期待。
（停 5 秒）
现在，什么都不用做。不用去捞那些泥沙。只是把杯子放稳，看着它。
（停 10 秒）
泥沙在一点一点往下沉……水，在一点一点变清……
（停 10 秒）
水清了。就在这杯清水里，问自己一个问题：今天，最值得我做的一件事，是什么？
（停 10 秒）
好，请慢慢睁开眼睛。
（轻声）老子问过一句话：孰能浊以静之徐清？谁能让浑水静下来，慢慢变清？答案是——你不需要“做”什么，你只需要停止搅动。这就是 G。
【提示】如果时间紧，这一页可以缩短到 1 分钟，但不建议跳过——这是全场唯一的身体体验。`);
	seal(s, "流");
	box(s, 0.6, 1.35, 4.2, 3.0);
	s.addImage({ data: ART.glasses, altText: "两杯水：一杯被搅浑，一杯静置后变清", x: 0.75, y: 1.45, w: 3.9, h: 2.19 });
	T(s, "搅动", 0.9, 3.75, 1.4, 0.35, { fontSize: 16, bold: true, color: C.accent5, align: "center" });
	T(s, "静", 2.2, 3.75, 1.0, 0.35, { fontSize: 16, bold: true, color: C.accent2, align: "center" });
	T(s, "自清", 3.05, 3.75, 1.4, 0.35, { fontSize: 16, bold: true, color: C.accent1, align: "center" });
	["闭上眼。想象手里有一杯刚被搅浑的水。", "什么都不做。只是看着泥沙，一点一点沉下去。", "水清了。问自己：今天最值得我做的一件事，是什么？"].forEach((t, i) => {
		const y = 1.4 + i * 1.0;
		dot(s, String(i + 1), 5.2, y + 0.08, 0.5, i === 2 ? C.accent2 : C.accent1, 16);
		T(s, t, 5.9, y, 3.5, 0.8, { fontSize: 15, valign: "middle", bold: i === 2 });
	});
	tao(s, "孰能浊以静之徐清？", "第十五章");

	s = add("CONTENT", "I · Influence Others：以柔化人", `
【I · 约 3 分钟】
自己定了，才谈得上第二圈：影响他人。
前面说过，软技能正在变成硬实力。在一个人人都能调用 AI 的时代，你没法靠“我会、你不会”来影响别人了。你只能靠四件很“柔”的事。
第一，倾听。先处下，再开口。老子说，江海之所以能成为百川之王，是因为它待在最低的地方。会听的人，才是所有信息、所有信任汇聚的地方。
第二，提问。Google 的人才标准里第一条就是“提出正确的问题”。答案，AI 可以给你一百个；一个好问题，要靠你。
第三，讲故事。道理让人点头，故事让人行动。下一页我们专门讲怎么讲。
第四，成就他人。把功劳让出去。老子说：既以为人己愈有，既以与人己愈多。越是给别人，自己越富有。
这四件事，没有一件需要头衔。
天下之至柔，驰骋天下之至坚——最柔软的东西，能穿行于最坚硬的东西之中。`);
	seal(s, "流");
	[["FaHeadphones", "倾听", "先处下，再开口，听懂对方在意什么", "江海所以能为百谷王者，以其善下之", "第六十六章"], ["FaQuestion", "提问", "一个好问题，比十个好答案更稀缺", "Google 人才标准：提出正确的问题", ""], ["FaBookOpen", "讲故事", "让人看见画面，而不只是听见道理", "方法见后两页", ""], ["FaHandsHelping", "成就他人", "把功劳让出去，影响力才会流回来", "既以为人己愈有，既以与人己愈多", "第八十一章"]].forEach(([ic, label, desc, q, chap], i) => {
		const y = 1.3 + i * 0.78;
		box(s, 0.6, y, 8.8, 0.68);
		icon(s, ic, 0.72, y + 0.1, 0.48);
		T(s, label, 1.4, y, 1.15, 0.68, { fontSize: 16, bold: true, valign: "middle" });
		T(s, desc, 2.55, y, 3.55, 0.68, { fontSize: 14, valign: "middle" });
		T(s, chap ? [{ text: q, options: { bold: true, color: C.accent2 } }, { text: "\n《道德经》" + chap, options: { color: C.accent5, fontSize: 10 } }] : [{ text: q, options: { color: C.accent5 } }], 6.25, y, 3.05, 0.68, { fontSize: 12, valign: "middle" });
	});
	tao(s, "天下之至柔，驰骋天下之至坚。", "第四十三章");

	s = add("CONTENT", "让话有黏性：Made to Stick 六原则", `
【讲故事的方法 · 约 2 分钟】
为什么有些话听完就忘，有些话能记一辈子？
Chip Heath 和 Dan Heath 兄弟在《Made to Stick》这本书里，总结了让一个想法“黏”住人的六个原则，首字母拼起来是 SUCCESs。
其实今天这堂课，我一直在用这六条，大家可以对一对。
Simple 简单：全场只有一句核心——执行越便宜，判断越昂贵。
Unexpected 意外：最危险的，竟然是能力太强。
Concrete 具体：洪水、河道、一杯浊水——都是看得见摸得着的东西。
Credible 可信：Google Research 负责人讲的洪水预测案例。
Emotional 情感：凌晨一点，合上电脑时的那个念头。
Stories 故事：鲧与禹，堵与疏。
下次你要说服老板、带动团队、介绍自己的时候，拿这六条检查一遍：我的话，够简单吗？有意外吗？有画面吗？
【提示】“凌晨一点”如果替换成了你的真实故事，请同步更新这一页的示例。`);
	seal(s, "流");
	[["S", "简单", "Simple", "执行越便宜，\n判断越昂贵"], ["U", "意外", "Unexpected", "最危险的，\n是能力太强"], ["C", "具体", "Concrete", "洪水、河道、\n一杯浊水"], ["C", "可信", "Credible", "Google 的\n洪水预测案例"], ["E", "情感", "Emotional", "凌晨一点的\n那个念头"], ["S", "故事", "Stories", "鲧与禹，\n堵与疏"]].forEach(([l, name, en, eg], i) => {
		const x = 0.6 + (i % 3) * 3.0, y = 1.3 + Math.floor(i / 3) * 1.55;
		box(s, x, y, 2.8, 1.35);
		T(s, l, x + 0.2, y + 0.15, 0.6, 1.05, { fontSize: 36, bold: true, color: C.accent2, valign: "middle", align: "center" });
		T(s, [{ text: name + " ", options: { bold: true, fontSize: 16 } }, { text: en, options: { fontSize: 11, color: C.accent1, bold: true } }], x + 0.95, y + 0.18, 1.75, 0.35, { valign: "middle" });
		T(s, eg, x + 0.95, y + 0.58, 1.75, 0.6, { fontSize: 14, color: C.accent5 });
	});
	T(s, "Chip Heath & Dan Heath《Made to Stick》（中文版《让创意更有黏性》）　·　灰色小字是本课对应的示例", 0.6, 4.55, 8.8, 0.3, { fontSize: 11, color: C.accent5 });

	s = add("CONTENT", "你的 3 分钟影响力故事：六步", `
【工具 + 课堂练习 · 约 4 分钟】
光知道原则不够，我给大家一个可以直接套用的结构，六步。今天这堂课的骨架，用的就是它。
Hook 钩子：一个让人停下来的问题。比如“如果岗位名称被抹掉，你还剩下什么？”
Story 故事：一个具体的人、一个具体的场景。
Conflict 冲突：他想要什么？被什么挡住了？没有冲突，就没有故事。
Highlight 高光：转折发生的那一刻。
Tao Insight 洞见：一句别人带得走的智慧。可以是《道德经》，也可以是你自己的话。
Offer 邀请：你希望对方接下来做什么。
（课堂练习）现在请大家两人一组，每人 90 秒，用这六步，讲一件你最近用 AI 做成的事——或者做砸的事。做砸的故事往往更好。
（计时 3 分钟，到一半时提醒换人。结束后请一两位分享“对方的故事里哪一步最打动你”。）
【提示】如果是线上课程，可改为在聊天区写下自己的 Hook 一句话。`);
	seal(s, "流");
	s.addShape(SH.line, { x: 1.0, y: 1.8, w: 8.0, h: 0, line: { color: C.accent6, width: 2.5 } });
	[["Hook", "钩子", "一个让人停下\n来的问题"], ["Story", "故事", "一个具体的人\n和场景"], ["Conflict", "冲突", "他要什么？\n被什么挡住"], ["Highlight", "高光", "转折发生的\n那一刻"], ["Tao Insight", "洞见", "一句带得走\n的智慧"], ["Offer", "邀请", "你希望对方\n做什么"]].forEach(([en, zh, desc], i) => {
		const x = 0.6 + i * 1.48;
		dot(s, String(i + 1), x + 0.3, 1.4, 0.8, i === 4 ? C.accent2 : C.text2, 22);
		T(s, en, x, 2.3, 1.4, 0.25, { fontSize: 11, bold: true, color: C.accent1, align: "center" });
		T(s, zh, x, 2.55, 1.4, 0.35, { fontSize: 18, bold: true, align: "center" });
		T(s, desc, x, 2.95, 1.4, 0.55, { fontSize: 14, color: C.accent5, align: "center" });
	});
	box(s, 0.6, 3.8, 8.8, 0.95);
	T(s, [{ text: "课堂练习 · 3 分钟　", options: { bold: true, color: C.accent2 } }, { text: "两人一组，用这六步，讲一件你最近用 AI 做成——或做砸——的事。" }], 0.85, 3.8, 8.3, 0.95, { fontSize: 15, valign: "middle" });

	s = add("CONTENT", "A · Align the System：顺势成事", `
【A · 约 3 分钟】
最外面一圈，A，对齐系统。
还记得第一部分提到的“环境智能”吗？马蒂亚斯的判断是：当一种技术真正成熟、足够可靠、足够自然，人们反而会忘记它的存在。就像机器翻译——过去要找专业的人，今天我们已经不觉得那是一项技术了。最好的 AI 产品，不是最能展示 AI 的产品，而是最能让人忘记 AI 的产品。
（指右边）两千多年前，老子讲最好的领导，讲的是同一件事。
太上，不知有之。最好的领导者，人们几乎感觉不到他的存在。功成事遂，百姓皆谓：我自然。事情办成了，大家都说：是我们自己做到的。
最好的技术，让人忘记技术。最好的领导，让人忘记领导。
所以，对齐系统的人，不是站在中央发号施令的人。他是禹——他不推着水走，他修好河道，让一切自然地流向目标。
【版本说明】第十七章首句，通行本多作“太上，下知有之”，另有版本作“不知有之”，这里取后者；如有学员问起可以说明。`);
	seal(s, "流");
	box(s, 0.6, 1.35, 4.0, 2.55);
	T(s, "最好的技术", 0.85, 1.52, 3.5, 0.25, { fontSize: 12, bold: true, color: C.accent1 });
	T(s, "环境智能", 0.85, 1.82, 3.5, 0.42, { fontSize: 22, bold: true });
	T(s, "Ambient Intelligence", 0.85, 2.25, 3.5, 0.25, { fontSize: 11, color: C.accent5 });
	T(s, "当技术足够成熟、可靠、自然，人们反而会忘记它的存在。", 0.85, 2.62, 3.5, 0.7, { fontSize: 14 });
	T(s, "—— Yossi Matias", 0.85, 3.42, 3.5, 0.25, { fontSize: 11, color: C.accent5 });
	box(s, 5.4, 1.35, 4.0, 2.55, C.text2);
	T(s, "最好的领导", 5.65, 1.52, 3.5, 0.25, { fontSize: 12, bold: true, color: C.accent4 });
	T(s, "太上，不知有之", 5.65, 1.82, 3.5, 0.42, { fontSize: 22, bold: true, color: C.background1 });
	T(s, "功成事遂，\n百姓皆谓：我自然。", 5.65, 2.62, 3.5, 0.7, { fontSize: 14, color: C.background1 });
	T(s, "——《道德经》第十七章", 5.65, 3.42, 3.5, 0.25, { fontSize: 11, color: C.accent6 });
	T(s, "≈", 4.6, 2.3, 0.8, 0.6, { fontSize: 32, bold: true, color: C.accent2, align: "center", valign: "middle" });
	T(s, "对齐系统的人，不站在中央发号施令——他让一切自然地流向目标。", 0.6, 4.2, 8.8, 0.4, { fontSize: 16, bold: true, color: C.text2 });

	s = add("CONTENT", "对齐系统的五件事", `
【A 的落地 · 约 2 分钟】
“对齐系统”听起来很大，其实落下来是五件具体的事。
定义目标：这件事，值得做吗？——注意，这是第一步，不是最后一步。
拆解任务：哪些交给人，哪些交给 AI？
协调能力：人和 Agent 怎么各得其所？
评估结果：怎么验证它真的有效？
承担决策：谁为方向和结果负责？
过去，这五件事是管理者的专属工作。现在，不是说每个人都要去当管理者，而是组织能力，正在从管理岗位的专属要求，变成每个知识工作者的基础能力。你手里只要有 AI，你就已经在“带团队”了。
老子说：为之于未有，治之于未乱。在事情还没发生的时候就着手，在混乱还没出现的时候就治理。这不就是尼日利亚那个故事吗——在洪水到来之前，让资源先流过去。
另外送大家一句：治大国，若烹小鲜。对齐系统，像煎小鱼——别老去翻它。`);
	seal(s, "流");
	[["定义目标", "这件事，\n值得做吗？"], ["拆解任务", "哪些给人，\n哪些给 AI？"], ["协调能力", "人与 AI\n各得其所？"], ["评估结果", "怎么验证\n它真的有效？"], ["承担决策", "谁为方向和\n结果负责？"]].forEach(([label, q], i) => {
		const x = 0.6 + i * 1.8;
		box(s, x, 1.4, 1.6, 2.2, i === 0 ? C.text2 : C.background2);
		dot(s, String(i + 1), x + 0.2, 1.6, 0.48, i === 0 ? C.accent2 : C.accent1, 15);
		T(s, label, x + 0.2, 2.25, 1.3, 0.35, { fontSize: 16, bold: true, color: i === 0 ? C.background1 : C.text1 });
		T(s, q, x + 0.2, 2.68, 1.3, 0.75, { fontSize: 14, color: i === 0 ? C.accent6 : C.accent5 });
	});
	T(s, "过去，这是管理者的专属工作；现在，它是每个知识工作者的基础能力。", 0.6, 3.85, 8.8, 0.4, { fontSize: 15, color: C.text2, bold: true });
	tao(s, "为之于未有，治之于未乱。", "第六十四章");

	s = add("CONTENT", "GIA 一页纸", `
【小结 · 约 1 分钟】
我们把 GIA 收在一页纸上，大家可以拍照。
G，向内扎根。核心问题：我是谁，凭什么判断。它是静水，负责澄清。练习是每天一分钟的“浊水静置”。
I，以柔化人。核心问题：没有职位，如何影响。它是流水，负责润物。练习是用六步讲一个三分钟的故事。
A，顺势成事。核心问题：系统如何自己流向目标。它是江海，负责汇聚。练习是选一个手头的流程，先问它该不该存在。
三样东西合在一起，就是你的河道。
河道修好了，接下来我们来看水——AI 这场大水，怎么用。`);
	seal(s, "流");
	const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 13, valign: "middle" } });
	const cell = (t, o = {}) => ({ text: t, options: { fontSize: 14, color: C.text1, valign: "middle", ...o } });
	const rowFill = [C.background2, C.background1, C.background2];
	const rows = [["G　向内扎根", "我是谁？凭什么判断？", "静水 · 澄清", "自知者明", "每天一分钟“浊水静置”"], ["I　以柔化人", "没有职位，如何影响？", "流水 · 润物", "至柔驰骋至坚", "用六步讲一个 3 分钟故事"], ["A　顺势成事", "系统如何自己流向目标？", "江海 · 汇聚", "百姓皆谓我自然", "选一个流程，先问该不该存在"]];
	s.addTable(
		[["", "核心问题", "水的形态", "《道德经》", "本周练习"].map(hdr), ...rows.map((r, i) => r.map((t, j) => cell(t, { fill: { color: rowFill[i] }, bold: j === 0, color: j === 0 ? C.accent2 : j === 3 ? C.text2 : C.text1 })))],
		{ x: 0.6, y: 1.35, w: 8.8, colW: [1.5, 2.25, 1.35, 1.7, 2.0], rowH: [0.5, 0.9, 0.9, 0.9], border: { type: "solid", color: THEME.colors.accent6, pt: 0.75 }, margin: [0.04, 0.1, 0.04, 0.1] },
	);

	// =====================================================================
	// 第三部分 · 势
	// =====================================================================
	startSection("第三部分 · 势：ACE");

	s = add("SECTION", "乘势：与 AI 共舞的 ACE", `
【过场 · 约 30 秒】
道生之，德畜之，物形之，势成之。
万物由道而生，由德滋养，由物质成形，最后——由“势”来成就。
AI，就是我们这个时代最大的“势”。
有了河道，我们就不怕水大。接下来十几分钟，我们讲怎么乘这个势：ACE。`);
	s.addText("道生之，德畜之，物形之，势成之。\n——《道德经》第五十一章", { placeholder: "body" });
	T(s, "第三部分", 0.7, 1.3, 3, 0.3, { fontSize: 14, color: C.accent6 });
	seal(s, "势", 0.7, 0.55, 0.65);

	s = add("CONTENT", "ACE：让 AI 成为你的水势", `
【模型总览 · 约 2 分钟】
ACE，三个字母，三个层次。
A，Accelerate，加速。用 AI 做自动化，把执行交出去。这一层解决的是：做得更快。
C，Collaborate，协作。和 AI 协作，让它做你的思考伙伴，把你的思考放大。这一层解决的是：想得更好。
E，Elevate，跃迁。借助 AI，把自己从“执行者”升到“结果负责人”。这一层解决的是：成为不同的自己。
请注意，它们不是三个工具，而是三个层次：从省时间，到长本事，到换位置。
大多数人停在了 A。用 AI 写写邮件、做做总结，省下来的时间，又被更多的邮件填满了——这就是我们前面说的目标膨胀。
真正的分水岭，在 C 和 E。`);
	seal(s, "势");
	[["A", "Accelerate", "加速", "用 AI 自动化，\n把执行交出去", "做得更快", C.text2], ["C", "Collaborate", "协作", "与 AI 协作，\n把思考放大", "想得更好", C.accent1], ["E", "Elevate", "跃迁", "借 AI 升维，\n成为结果负责人", "成为不同的自己", C.accent2]].forEach(([l, en, zh, desc, chip, color], i) => {
		const x = 0.6 + i * 3.0;
		box(s, x, 1.35, 2.8, 3.0);
		T(s, l, x + 0.25, 1.45, 1.0, 0.95, { fontSize: 54, bold: true, color, valign: "middle" });
		T(s, en, x + 1.25, 1.62, 1.4, 0.25, { fontSize: 12, color: C.accent5 });
		T(s, zh, x + 1.25, 1.88, 1.4, 0.42, { fontSize: 22, bold: true });
		T(s, desc, x + 0.25, 2.65, 2.3, 0.6, { fontSize: 14 });
		s.addText(chip, { shape: SH.roundRect, rectRadius: 0.17, x: x + 0.25, y: 3.65, w: 2.3, h: 0.42, fill: { color }, line: { type: "none" }, color: C.background1, bold: true, fontSize: 14, align: "center", valign: "middle", margin: 0 });
	});
	T(s, "三步不是三个工具，而是三个层次：从省时间，到长本事，到换位置。", 0.6, 4.6, 8.8, 0.35, { fontSize: 15, bold: true, color: C.text2 });

	s = add("CONTENT", "A · Accelerate：先做减法，再自动化", `
【A · 约 3 分钟】
先说 A，加速。这一步最容易，也最容易做错。
做错的方式，就是我们前面讲的“错误的胜利”：拿到一个流程，二话不说就自动化。
所以在把任何一件事交给 AI 之前，请先过这四道关。
第一关，也是最重要的一关：这个任务合理吗？它该不该存在？
第二关：能不能交给 AI？
第三关：应该如何验证？AI 做完了，我怎么知道它是对的？
第四关：最终创造什么价值？
（指右边）把一个不该存在的流程自动化，只是更高效地浪费。而且因为单次执行太便宜了，我们反而更难察觉这种浪费。
老子说：为学日益，为道日损。损之又损，以至于无为。求学问，是每天做加法；求道，是每天做减法。减到不能再减，剩下的才是真正要做的。
自动化之前，先做减法。`);
	seal(s, "势");
	["这个任务合理吗？它该不该存在？", "能不能交给 AI？", "应该如何验证？", "最终创造什么价值？"].forEach((q, i) => {
		const y = 1.35 + i * 0.72;
		box(s, 0.6, y, 5.1, 0.62, i === 0 ? C.text2 : C.background2);
		dot(s, String(i + 1), 0.72, y + 0.09, 0.44, i === 0 ? C.accent2 : C.accent1, 15);
		T(s, q, 1.35, y, 4.25, 0.62, { fontSize: 16, bold: i === 0, color: i === 0 ? C.background1 : C.text1, valign: "middle" });
	});
	box(s, 6.0, 1.35, 3.4, 2.78, C.accent2);
	s.addImage({ data: ICONS.FaExclamationTriangle, x: 6.25, y: 1.6, w: 0.5, h: 0.5 });
	T(s, "把一个不该存在的流程自动化，只是更高效地浪费。", 6.25, 2.3, 2.9, 1.6, { fontSize: 18, bold: true, color: C.background1 });
	tao(s, "为学日益，为道日损。损之又损，以至于无为。", "第四十八章");

	s = add("CONTENT", "C · Collaborate：AI 出力，人把关", `
【C · 约 3 分钟】
第二层，C，协作。协作的关键，是分清楚：哪些 AI 可以，哪些必须是你。
看这四组不等号，它们都来自马蒂亚斯的提醒。
AI 可以提供知识，但这不等于你拥有判断知识的能力——哪条重要，哪条是错的，得你来判断。
AI 可以生成方案，但这不等于能识别方案里的缺陷。
AI 可以帮你跨越专业边界，但这不等于你真正理解了每一个专业。
AI 可以完成执行，但它不能为结果承担责任。签字的，永远是你。
所以好的协作不是“AI 做，我收货”，而是“AI 出力，我把关”。你可以这样用：让 AI 来反驳你的方案，问它“这个计划最可能在哪里失败？”——让它做你的陪练，而不只是你的代笔。
老子说：知不知，尚矣；不知知，病也。知道自己不知道，是最高明的；不知道却以为自己知道，才是毛病。
在 AI 时代，这句话应该贴在每个人的屏幕旁边。`);
	seal(s, "势");
	T(s, "AI 可以", 0.6, 1.3, 3.9, 0.35, { fontSize: 16, bold: true, color: C.accent1, align: "center" });
	T(s, "你必须", 5.5, 1.3, 3.9, 0.35, { fontSize: 16, bold: true, color: C.text2, align: "center" });
	[["提供知识", "判断知识的轻重与真伪"], ["生成方案", "识别方案里的缺陷"], ["跨越专业边界", "真正理解关键的专业"], ["完成执行", "为结果承担责任"]].forEach(([a, b], i) => {
		const y = 1.78 + i * 0.66;
		s.addText(a, { shape: SH.roundRect, rectRadius: 0.08, x: 0.6, y, w: 3.9, h: 0.56, fill: { color: C.background2 }, line: { type: "none" }, color: C.text1, fontSize: 16, align: "center", valign: "middle", margin: 0 });
		T(s, "≠", 4.5, y, 1.0, 0.56, { fontSize: 22, bold: true, color: C.accent2, align: "center", valign: "middle" });
		s.addText(b, { shape: SH.roundRect, rectRadius: 0.08, x: 5.5, y, w: 3.9, h: 0.56, fill: { color: C.text2 }, line: { type: "none" }, color: C.background1, bold: true, fontSize: 16, align: "center", valign: "middle", margin: 0 });
	});
	tao(s, "知不知，尚矣；不知知，病也。", "第七十一章");

	s = add("CONTENT", "E · Elevate：从执行者到结果负责人", `
【E · 约 2 分钟】
第三层，E，跃迁。这是 ACE 里最重要的一步，也是今天和大家的职业发展最相关的一步。
看这四级台阶。
第一级，执行者：完成分配给我的任务。
第二级，协作者：和 AI 一起完成任务。大多数人现在在这里。
第三级，编排者：我不再亲手做每一件事，我组织人和 AI 的能力来完成它。
第四级，结果负责人：我定义问题，我承担结果。
每往上一级，你亲手做的事更少，你负责的范围更大。
这里出现了一种角色的反转：过去，管理者负责组织资源，专业人员负责执行。未来，越来越多的专业人员，也要承担一部分过去属于管理者的工作。
不是每个人都要成为管理者，但每个人都要学会组织能力。
（问全场）大家可以在心里给自己定个位：今天，我站在第几级？`);
	seal(s, "势");
	T(s, "不是每个人都要成为管理者，\n但每个人都要学会组织能力。", 0.6, 1.35, 4.4, 0.75, { fontSize: 16, bold: true, color: C.text2 });
	[["执行者", "完成分配的任务", C.accent6, C.text1, C.text2], ["协作者", "与 AI 一起\n完成任务", C.accent3, C.text1, C.text1], ["编排者", "组织人与 AI\n的能力", C.accent1, C.background1, C.background1], ["结果负责人", "定义问题，\n承担结果", C.text2, C.background1, C.accent6]].forEach(([label, desc, fill, fg, sub], i) => {
		const x = 0.6 + i * 2.25, hgt = 1.15 + i * 0.5, y = 4.45 - hgt;
		s.addShape(SH.rect, { x, y, w: 2.05, h: hgt, fill: { color: fill }, line: { type: "none" } });
		T(s, "0" + (i + 1), x + 0.2, y + 0.12, 0.5, 0.25, { fontSize: 11, bold: true, color: sub });
		T(s, label, x + 0.2, y + 0.36, 1.7, 0.32, { fontSize: 16, bold: true, color: fg });
		if (i > 0) T(s, desc, x + 0.2, y + 0.74, 1.7, 0.55, { fontSize: 14, color: sub });
		else T(s, desc, x + 0.2, y + 0.74, 1.7, 0.3, { fontSize: 14, color: sub });
	});

	s = add("CONTENT", "GIA × ACE：水量与河道", `
【合龙 · 约 2 分钟】
现在，我们把两个模型放在一起。
横轴是 ACE——AI 带给你的执行力，也就是水量。纵轴是 GIA——你的判断和定力，也就是河道。
左下角，两样都没有：一潭止水，原地打转。
左上角，有判断、没有 AI：这是溪流。方向是对的，但走得慢，会被越甩越远。
右下角，要特别小心：AI 能力很强，但判断没跟上。这就是洪水——高效地做错的事，也就是我们说的“错误的胜利”。能力越强，破坏越大。这也回答了开头那个问题：为什么最危险的，可能是能力太强。
右上角，是我们要去的地方：江河。柔，而有力；有方向，也有力量，奔流到海。
ACE 决定水有多大，GIA 决定水往哪里流。
AI 会放大正确的选择，也会放大错误的方向。所以这两样，必须一起修。`);
	seal(s, "势");
	T(s, "↑ GIA · 判断与定力", 0.6, 1.3, 3, 0.28, { fontSize: 11, bold: true, color: C.accent5 });
	[["溪流", "方向对，但走得慢", C.background2, C.text1, C.accent5], ["江河", "柔而有力，奔流到海", C.text2, C.background1, C.accent6], ["止水", "原地打转", C.background2, C.text1, C.accent5], ["洪水", "高效地做错的事", C.accent2, C.background1, C.background1]].forEach(([name, desc, fill, fg, sub], i) => {
		const x = 0.6 + (i % 2) * 2.5, y = 1.65 + Math.floor(i / 2) * 1.42;
		box(s, x, y, 2.4, 1.32, fill);
		T(s, name, x + 0.22, y + 0.25, 2.0, 0.42, { fontSize: 22, bold: true, color: fg });
		T(s, desc, x + 0.22, y + 0.75, 2.0, 0.3, { fontSize: 14, color: sub });
	});
	T(s, "ACE · AI 带来的执行力 →", 2.5, 4.52, 3.0, 0.28, { fontSize: 11, bold: true, color: C.accent5, align: "right" });
	T(s, "ACE 决定\n水有多大。", 6.0, 1.65, 3.4, 0.9, { fontSize: 22, bold: true, color: C.accent1 });
	T(s, "GIA 决定\n水往哪里流。", 6.0, 2.65, 3.4, 0.9, { fontSize: 22, bold: true, color: C.text2 });
	T(s, "AI 会放大正确的选择，\n也会放大错误的方向。", 6.0, 3.8, 3.4, 0.6, { fontSize: 14, color: C.accent5 });

	// =====================================================================
	// 第四部分 · 海
	// =====================================================================
	startSection("第四部分 · 海：你的价值");

	s = add("SECTION", "归海：找到并放大你的价值", `
【过场 · 约 30 秒】
江海所以能为百谷王者，以其善下之。
江海之所以能让千百条河流都归向它，是因为它甘于处在最低的地方。
最后一段路，我们回到每一个人自己身上：有了河道，有了水，你要流向哪里？
你的价值是什么？怎么让自己更有价值？`);
	s.addText("江海所以能为百谷王者，以其善下之。\n——《道德经》第六十六章", { placeholder: "body" });
	T(s, "第四部分", 0.7, 1.3, 3, 0.3, { fontSize: 14, color: C.accent6 });
	seal(s, "海", 0.7, 0.55, 0.65);

	s = add("CONTENT", "重新定义“你值多少”", `
【约 2 分钟】
过去，我们习惯用三样东西来判断一个人值多少：职位、专业、岗位职责。因为岗位代表能力，职位代表权限，组织代表资源。
但当越来越多的专业能力可以被个人直接调用，这些标签就越来越难描述一个人真实的价值了。
那未来用什么描述？看下面这四个方块。
你能发现什么问题——这是眼光。
你能组织多少能力——人的，和 AI 的。这是杠杆。
你愿意承担什么结果——这是担当。
这三样乘在一起，才是：你最终创造多大的价值。
注意，这里面没有一项是“你掌握了多少技能”。技能还重要，但它变成了入场券，不再是奖杯。
（停顿）请大家在心里，用这四个方块，重新给自己写一份“简历”。`);
	seal(s, "海");
	T(s, [{ text: "过去　", options: { bold: true, color: C.accent5, fontSize: 12 } }, { text: "价值 ≈ 职位 ＋ 专业 ＋ 岗位职责", options: { fontSize: 16, color: C.accent5 } }], 0.6, 1.3, 8.8, 0.4, { valign: "middle" });
	[["FaEye", "发现什么问题", "眼光"], ["FaProjectDiagram", "组织多少能力", "杠杆"], ["FaPenNib", "承担什么结果", "担当"], ["FaGem", "创造多大价值", "未来的你"]].forEach(([ic, label, tag], i) => {
		const x = 0.6 + i * 2.3, last = i === 3;
		box(s, x, 1.95, 1.9, 1.85, last ? C.text2 : C.background2);
		icon(s, ic, x + 0.65, 2.12, 0.6, last ? C.accent2 : C.accent1);
		T(s, label, x, 2.85, 1.9, 0.38, { fontSize: 16, bold: true, color: last ? C.background1 : C.text1, align: "center" });
		T(s, tag, x, 3.28, 1.9, 0.28, { fontSize: 12, color: last ? C.accent6 : C.accent5, align: "center" });
		if (i < 3) T(s, i === 2 ? "＝" : "×", x + 1.9, 1.95, 0.4, 1.85, { fontSize: 22, bold: true, color: C.accent2, align: "center", valign: "middle" });
	});
	T(s, "标签越来越难描述一个人。能描述你的，是你发现的问题、组织的能力、承担的结果。", 0.6, 4.1, 8.8, 0.7, { fontSize: 16, bold: true, color: C.text2 });

	s = add("CONTENT", "找到你的价值：四个问题", `
【How to find your value · 约 3 分钟】
怎么找到自己的价值？我给大家四个问题。请拿出笔，或者打开手机备忘录，每个问题写一行。
第一问，在意：哪个问题，让你放不下？没有人要求，你也想把它解决。价值的起点不是能力，是在意。
第二问，判断：在哪件事上，你的判断比大多数人准？这来自你的专业深度——那一竖。
第三问，利他：谁会因为你而变好？水善利万物。价值永远是对别人而言的。
第四问，承担：哪个结果，你愿意签上自己的名字？AI 可以执行，但不能替你负责。你愿意负责的地方，就是 AI 替代不了你的地方。
（给 60 秒安静书写。）
这四个答案交汇的地方，就是你的河道。
如果有一问你写不出来——没关系，那恰恰告诉了你，接下来该往哪里用力。`);
	seal(s, "海");
	[["FaHeart", "在意", "哪个问题，让我放不下？", "没人要求，我也想把它解决"], ["FaBalanceScale", "判断", "我在哪件事上判断得比别人准？", "这来自你的专业深度"], ["FaUsers", "利他", "谁会因为我而变好？", "水善利万物——价值是对别人而言的"], ["FaPenNib", "承担", "哪个结果，我愿意签上名字？", "AI 可以执行，但不能替你负责"]].forEach(([ic, key, q, sub], i) => {
		const x = 0.6 + (i % 2) * 4.5, y = 1.3 + Math.floor(i / 2) * 1.5;
		box(s, x, y, 4.3, 1.32);
		icon(s, ic, x + 0.22, y + 0.2, 0.6, i === 3 ? C.accent2 : C.accent1);
		T(s, key, x + 1.02, y + 0.14, 3.1, 0.3, { fontSize: 14, bold: true, color: C.accent1 });
		T(s, q, x + 1.02, y + 0.45, 3.15, 0.35, { fontSize: 15, bold: true });
		T(s, sub, x + 1.02, y + 0.85, 3.15, 0.28, { fontSize: 12, color: C.accent5 });
	});
	T(s, "四个答案交汇的地方，就是你的河道。", 0.6, 4.45, 8.8, 0.4, { fontSize: 16, bold: true, color: C.accent2 });

	s = add("CONTENT", "更有价值：三个动作", `
【How to be more valuable · 约 3 分钟】
找到了，怎么放大？三个动作，都和水有关。
第一，深挖一口井。把你的专业挖深。前面说过，专业深度决定判断的质量。AI 让所有人都能“知道”，但只有挖得深的人才能“看出问题”。
第二，连通更多水系。以你的专业判断为中心，去连接其他领域的能力——人的，AI 的。组织能力决定你创造价值的范围。一口井只能养一户人家，连成水系，才能养一方土地。
第三，也是最反直觉的一条：流向低处。水不往高处挤，它去别人不愿去的地方。老子说，处众人之所恶，故几于道。那些没人愿意碰、但真实存在的问题，那些不光鲜、但真的有人需要的地方——就是价值最密集的地方。
最后送大家一句话：既以为人己愈有，既以与人己愈多。越是为别人，自己越富有。
在 AI 时代，这不是鸡汤，这是算术——因为你的价值，是由你改变了什么来计量的。`);
	seal(s, "海");
	[["FaLayerGroup", "深挖一口井", "让专业深度，成为判断的底气", "深度 → 判断的质量"], ["FaProjectDiagram", "连通更多水系", "以专业判断为中心，连接其他领域的能力", "连接 → 价值的范围"], ["FaTint", "流向低处", "去别人不愿去、却真有需求的地方", "处众人之所恶，故几于道"]].forEach(([ic, label, desc, tag], i) => {
		const x = 0.6 + i * 3.0;
		box(s, x, 1.35, 2.8, 2.95);
		icon(s, ic, x + 0.25, 1.55, 0.66, i === 2 ? C.accent2 : C.accent1);
		T(s, label, x + 0.25, 2.4, 2.3, 0.38, { fontSize: 18, bold: true });
		T(s, desc, x + 0.25, 2.85, 2.3, 0.62, { fontSize: 14 });
		T(s, tag, x + 0.25, 3.7, 2.3, 0.3, { fontSize: 12, bold: true, color: i === 2 ? C.accent2 : C.accent1 });
	});
	tao(s, "既以为人己愈有，既以与人己愈多。", "第八十一章");

	s = add("IMAGE_RIGHT", "岗位是起点，不是终点", `
【约 1 分钟】
（画面：百川归海，太阳在海平面上。）
过去，是岗位决定一个人能做什么。你需要通过一个岗位进入一个组织，才能获得创造价值所需的能力。
未来，是一个人创造的价值，决定他的岗位。你可以先通过 AI 扩展自己的能力，再决定如何创造价值、与谁协作。
岗位不再是能力的终点，它只是你创造价值的一个起点。
所以，真正需要重新思考的，不只是“我的岗位会不会被 AI 取代”，而是：当岗位不再决定我的能力边界，我准备怎么重新定义自己的价值？`);
	s.addImage({ data: ART.sceneSea, altText: "日出时的大海", x: 6.0, y: 0, w: 4.0, h: 5.625 });
	[["过去", "岗位", "决定你能做什么", C.background2, C.text1, C.accent5], ["未来", "你创造的价值", "决定岗位", C.text2, C.background1, C.accent6]].forEach(([tag, a, b, fill, fg, sub], i) => {
		const y = 1.35 + i * 1.3;
		box(s, 0.6, y, 5.0, 1.1, fill);
		T(s, tag, 0.85, y + 0.14, 1, 0.25, { fontSize: 12, bold: true, color: sub });
		T(s, [{ text: a, options: { bold: true, fontSize: 20 } }, { text: "　→　" + b, options: { fontSize: 16 } }], 0.85, y + 0.45, 4.5, 0.45, { color: fg, valign: "middle" });
	});
	T(s, "先用 AI 扩展自己的能力，\n再决定如何创造价值、与谁协作。", 0.6, 4.05, 5.0, 0.7, { fontSize: 15, bold: true, color: C.text2 });

	s = add("CONTENT", "带走：七天，三件小事", `
【行动 · 约 2 分钟】
听完一堂课，如果什么都不做，三天就忘了。所以我只请大家带走三件小事，用七天去做。
G：每天一分钟“浊水静置”。做完之后，写下一条“我不做什么”。七天下来，你会有一张七条的“不做清单”。
I：用今天的六步结构，给一位同事讲一个三分钟的故事。可以是一个项目，一次失败，一个你想推动的改变。
A：选一个你手头的流程。先问它“该不该存在”，再决定交不交给 AI。——如果答案是“不该存在”，恭喜你，你刚刚完成了今年最有价值的一次“自动化”。
三件事，都很小。老子说：千里之行，始于足下。图难于其易，为大于其细。
【提示】如果系列课程有下一讲，可以请大家把“不做清单”带来分享。`);
	seal(s, "海");
	[["G", "每天 1 分钟“浊水静置”", "做完后，写下一条“我不做什么”", C.text2], ["I", "用六步结构，讲一个 3 分钟的故事", "讲给一位同事听，请对方说说哪一步最打动他", C.accent1], ["A", "选一个手头的流程，先问“该不该存在”", "再决定交不交给 AI", C.accent3]].forEach(([l, a, b, color], i) => {
		const y = 1.35 + i * 1.02;
		box(s, 0.6, y, 8.8, 0.9);
		seal(s, l, 0.78, y + 0.15, 0.6, color, 22);
		T(s, a, 1.65, y + 0.12, 7.5, 0.35, { fontSize: 16, bold: true, valign: "middle" });
		T(s, b, 1.65, y + 0.5, 7.5, 0.28, { fontSize: 12, color: C.accent5 });
	});
	tao(s, "千里之行，始于足下。", "第六十四章");

	s = add("STATEMENT", "如果明天，岗位名称被抹掉——\n剩下的，就是你。", `
【收束 · 约 2 分钟】
（回到开场的问题。语速放到最慢。）
课程开始的时候，我问了大家一个问题：如果明天，你的岗位名称被抹掉，你还剩下什么？
现在我想给出我的回答。
剩下的，是你在意的问题；是你比别人看得准的判断；是因为你而变好的人；是你愿意签上名字的结果。
剩下的，就是你。
AI 可以拿走一个岗位，但它拿不走一条河道。
（停顿）
天下莫柔弱于水，而攻坚强者莫之能胜。天下没有比水更柔弱的东西，可是攻克坚硬的东西，没有什么能胜过它。
愿我们每个人，在这个越来越“硬”的时代里，活成一条河。
（停 3 秒，再翻页。）`);
	s.addText("天下莫柔弱于水，而攻坚强者莫之能胜。——《道德经》第七十八章", { placeholder: "body" });

	s = add("CONTENT", "下一步：柔性领导力系列课程", `
【Offer · 约 2 分钟】
今天是一张地图。地图上的每一块，都值得我们花时间真正走进去。
这就是柔性领导力系列课程想和大家一起做的事：G，向内扎根——情绪、潜意识和你的判断标准；I，以柔化人——倾听、提问和讲故事；A，顺势成事——让人、AI 和目标同向；以及 ACE——与 AI 共舞。
邀请大家，把今天的这一滴水，汇成一条河。
【重要 · 需主讲人确认】这一页左侧的模块是按 GIA 与 ACE 框架整理的示意，右侧的开课时间、报名方式、二维码均为占位。请替换为俱乐部实际的课程安排与报名信息后再使用；我没有这些信息，所以没有填写。`);
	seal(s, "海");
	[["G", "Ground Yourself · 向内扎根", "情绪、潜意识与你的判断标准", C.text2], ["I", "Influence Others · 以柔化人", "倾听、提问与讲故事", C.accent1], ["A", "Align the System · 顺势成事", "让人、AI 与目标同向", C.accent3], ["ACE", "与 AI 共舞", "加速 · 协作 · 跃迁", C.accent2]].forEach(([l, a, b, color], i) => {
		const y = 1.35 + i * 0.78;
		box(s, 0.6, y, 5.4, 0.68);
		seal(s, l, 0.72, y + 0.09, l === "ACE" ? 0.5 : 0.5, color, l === "ACE" ? 11 : 18);
		T(s, a, 1.42, y + 0.06, 4.45, 0.3, { fontSize: 15, bold: true, valign: "middle" });
		T(s, b, 1.42, y + 0.37, 4.45, 0.25, { fontSize: 12, color: C.accent5 });
	});
	s.addShape(SH.roundRect, { x: 6.3, y: 1.35, w: 3.1, h: 3.02, rectRadius: 0.08, fill: { color: C.background1 }, line: { color: C.accent5, width: 1.25, dashType: "dash" } });
	T(s, "[待确认]", 6.3, 1.75, 3.1, 0.35, { fontSize: 16, bold: true, color: C.accent2, align: "center" });
	T(s, "开课时间\n报名方式\n二维码", 6.3, 2.3, 3.1, 1.3, { fontSize: 16, color: C.accent5, align: "center", lineSpacingMultiple: 1.3 });
	T(s, "邀请你，把今天的一滴水，汇成一条河。", 0.6, 4.55, 5.4, 0.35, { fontSize: 15, bold: true, color: C.text2 });
	T(s, "[待确认：系列模块与排期以俱乐部公布为准]", 6.3, 4.55, 3.1, 0.35, { fontSize: 10, color: C.accent5, align: "center", valign: "middle" });

	// =====================================================================
	// Q&A
	// =====================================================================
	startSection("Q&A");

	s = add("SECTION", "Q&A：把问题留在水面上", `
【Q&A 开场 · 约 30 分钟】
接下来大约半小时，交给大家。
老子说：上士闻道，勤而行之。听到了道理，最好的回应不是点头，而是去做——而“做”的第一步，往往是一个好问题。
（先请现场提问。）
后面我准备了十个大家最常问的问题，每个问题一页。使用方法：现场有人问到，就翻到对应那一页；如果冷场，就从 Q1 开始抛砖引玉。每个问题建议控制在 3 分钟以内，不必讲完全部十个。
【提示】每一页上是“一句话回答”和三个要点，展开的说法在备注里。`);
	s.addText("上士闻道，勤而行之。\n——《道德经》第四十一章", { placeholder: "body" });
	T(s, "约 30 分钟", 0.7, 1.3, 3, 0.3, { fontSize: 14, color: C.accent6 });
	seal(s, "问", 0.7, 0.55, 0.65);

	const qa = (n, question, oneLine, points, notes) => {
		const q = add("QA", question, notes);
		dot(q, "Q" + n, 0.6, 0.4, 1.0, C.accent2, 24);
		box(q, 0.6, 1.7, 8.8, 1.0, C.text2);
		T(q, "一句话回答", 0.85, 1.8, 3, 0.22, { fontSize: 11, bold: true, color: C.accent4 });
		T(q, oneLine, 0.85, 2.04, 8.3, 0.6, { fontSize: 16, bold: true, color: C.background1, valign: "middle" });
		points.forEach((p, i) => {
			const x = 0.6 + i * 3.0;
			box(q, x, 2.95, 2.8, 1.6, C.background1, { shadow: shadow() });
			dot(q, String(i + 1), x + 0.2, 3.1, 0.4, C.accent1, 14);
			T(q, p, x + 0.2, 3.62, 2.4, 0.85, { fontSize: 15 });
		});
	};

	qa(1, "我的岗位，会被 AI 取代吗？", "更值得问的是：当岗位不再决定你的能力边界，你准备如何定义自己的价值？",
		["被重构的是“岗位”这个容器，不是人的价值", "执行最先被接管；判断与承担，最难被替代", "动手：把工作拆成“执行”和“判断”，看看各占多少"], `
【Q1 · 约 3 分钟】
这是被问得最多的问题，我先承认：没有人能给你一个确定的“会”或“不会”。
但我想把问题换一下。岗位，是过去组织用来装能力的容器。AI 正在把能力从容器里释放出来，所以容器的形状一定会变。会变的是岗位，不是人的价值。
在一个岗位里，最先被接管的是纯执行的部分。最难被替代的，是判断——什么值得做；和承担——谁为结果负责。
所以给大家一个今天回去就能做的小练习：把你上一周的工作列出来，每一项标上“执行”还是“判断”。如果 90% 是执行，这不是坏消息，这是一个很清晰的信号：接下来往哪里挪。
真正需要思考的，不是岗位会不会被取代，而是当岗位不再决定你的能力边界，你准备怎么重新定义自己的价值。`);

	qa(2, "我不是管理者，也需要领导力吗？", "需要。组织能力，正在从管理岗位的专属要求，变成每个知识工作者的基础能力。",
		["你至少在领导三样：自己、你的 AI、你负责的结果", "定义目标、拆解任务、评估结果，已是每个人的日常", "柔性领导力不需要头衔，只需要方向感"], `
【Q2 · 约 3 分钟】
需要，而且比以前更需要。
过去，管理者负责组织资源，专业人员负责执行。现在出现了一种角色反转：越来越多的专业人员，也要承担一部分过去属于管理者的工作——定义目标、拆解任务、协调能力、评估结果、承担决策。
你想想，当你同时让几个 AI 工具帮你干活的时候，你在做什么？你在分配任务，在验收，在为结果负责。你已经在“带团队”了，只是团队成员不是人。
所以今天讲的领导力，不是“领导别人”的权力，而是三件事：领导自己（G），影响身边的人（I），对齐你手里的小系统（A）。
这三件事，没有一件需要头衔。水没有头衔，可是万物都跟着它走。`);

	qa(3, "“柔”，是不是软弱、不争、躺平？", "柔不是弱。水最柔，却能穿石、能载舟、能成江海。",
		["不争 ≠ 不做，而是不做无谓的对抗", "柔是方法，不是目标——目标依然是把事做成", "夫唯不争，故天下莫能与之争（第二十二章）"], `
【Q3 · 约 3 分钟】
这是对“柔”最常见的误解。
大禹治水，十三年在外，三过家门而不入——这叫躺平吗？他比鲧更拼。区别不在于努不努力，而在于用力的方向：鲧和水对抗，禹和水合作。
所以“不争”不是不做事，而是不做无谓的对抗。不和同事争功，不和趋势较劲，不把力气花在证明“我是对的”上。省下来的力气，全部用在把事做成上。
柔是方法，不是目标。目标依然是成事。
老子的原话是：夫唯不争，故天下莫能与之争。正因为不去争，所以天下没有人能和他争。这是一种非常强的状态。
水流到石头面前，不撞，绕过去。但你过几百年再来看——石头圆了，水还在。`);

	qa(4, "判断力，到底怎么练？", "判断力没有速成班。它从三样东西里长出来：深度、反馈、复盘。",
		["深：在一个领域扎得足够深，才看得出 AI 的错", "验：每个重要判断，先写下预期，回头对答案", "省：每周问一次“哪件事，我本不该做”"], `
【Q4 · 约 3 分钟】
先说一句实话：判断力没有速成的办法，它是长出来的。但有三个可以每天做的动作。
第一，深。马蒂亚斯的原话的意思是：专业深度决定判断的质量。AI 能给你一份看起来很漂亮的方案，只有在这个领域扎得够深的人，才能一眼看出哪里不对。所以别因为有了 AI，就放弃把一件事学透。
第二，验。每做一个重要的判断，先写下来：我预期会发生什么？过一个月回头对答案。大多数人的判断力不进步，是因为从来不对答案。
第三，省。每周五问自己一个问题：这一周，哪件事我本不该做？这是在练“止”的肌肉。
再加上今天教的那一分钟“浊水静置”——判断之前，先让水清一清。
这几件事都很朴素，但合抱之木，生于毫末。`);

	qa(5, "AI 这么强，还需要深耕专业吗？", "更需要。专业深度，决定判断的质量。",
		["AI 能提供知识 ≠ 你拥有判断知识的能力", "AI 能生成方案 ≠ 你能识别方案的缺陷", "未来的优势：专业深度 × 跨领域的组织能力"], `
【Q5 · 约 3 分钟】
更需要。这一点马蒂亚斯在访谈里是特别强调的。
很多人有一个错觉：既然 AI 什么都知道，我就不用学那么深了。恰恰相反。
AI 能提供知识，不等于你拥有判断知识的能力。AI 能生成方案，不等于你能识别方案里的缺陷。AI 能帮你跨过专业的边界，也不等于你真的理解了那个专业。
你可以这样理解：AI 把“及格线”拉得很高，人人都能交出 80 分的东西。那么 80 分就不值钱了，值钱的是从 80 到 95 的那一段——而那一段，全靠专业深度。
未来真正有优势的，不是什么都会一点的人，而是我们画的那个 T：一竖是专业深度，一横是以这个专业判断为中心，去连接和组织其他领域的能力。
先把井挖深，再去连水系。`);

	qa(6, "我该从哪里开始用 AI？", "从一件你每周都在重复、而且你自己能判断好坏的事开始。",
		["A 加速：先把它交给 AI，省出时间", "C 协作：再让 AI 挑战你的方案，做你的陪练", "E 跃迁：把省下的时间，用在判断和负责结果上"], `
【Q6 · 约 3 分钟】
不要从“最厉害的工具”开始，从“最熟悉的事”开始。
选一件你每周都在重复做的事，而且——这一点很关键——这件事的好坏，你自己能判断。因为你能判断，你才知道 AI 做得对不对。
然后按 ACE 三步走。
A：先把它交给 AI。但交之前，过那四道关：这件事该不该存在？能不能交？怎么验证？创造什么价值？
C：等你用顺了，升一级。不只是让 AI 替你做，而是让它挑战你：“这个方案最可能在哪里失败？”让它做你的陪练。
E：最重要的一步——省下来的时间，千万不要再用来做更多同类的事。把它用在判断上，用在为结果负责上。
从一件小事开始。图难于其易，为大于其细。`);

	qa(7, "用了 AI，反而更忙、更焦虑，怎么办？", "这很正常——能力越强，目标越容易膨胀。解药是两个字：知止。",
		["先 G 后 ACE：先静下来，再加速", "写一张“不做清单”，和待办清单一样认真", "知足不辱，知止不殆，可以长久（第四十四章）"], `
【Q7 · 约 3 分钟】
首先，你不是一个人，这也不是你的问题。马蒂亚斯自己就预计，未来人们可能和今天一样忙。因为能力越强，我们就越会抬高目标。
所以更忙，往往不是因为 AI 没用好，而是因为它太好用了——你能做的事多了十倍，于是你就真的去做了十倍的事。
解药不在工具里，在两个字里：知止。
具体怎么做？第一，顺序要对：先 G，后 ACE。每天开始工作之前，先做那一分钟的“浊水静置”，问自己：今天最值得做的一件事是什么。再打开 AI。
第二，写一张“不做清单”。我们都有待办清单，但很少有人有不做清单。能做 10 个方向，不代表 10 个都值得做。
至于焦虑——回到那座冰山。它在水面之下。停、名、问。
知足不辱，知止不殆，可以长久。`);

	qa(8, "组织还在用旧规则看人，我怎么办？", "不对抗系统。像水一样：在规则之内，先让结果被看见。",
		["从一个小而真实的结果开始，而不是一场宣言", "讲清楚“改变了什么”，而不只是“完成了什么”", "把功劳让给团队——系统会慢慢为你让路"], `
【Q8 · 约 3 分钟】
这是一个很现实的问题：我想做结果负责人，可公司还在按岗位、按工时、按汇报线看我。
我的建议是：不要去做鲧。不要去撞那堵墙，不要发表一场“组织应该改革”的宣言。
做禹。看看水能往哪里流。
第一，从一个小而真实的结果开始。在你的职责范围之内，用新的方式把一件事做成。小到没有人会反对。
第二，学会讲这个结果。不是“我完成了什么”，而是“这件事改变了什么”——用六步结构讲。很多人不是没有价值，是不会让价值被看见。
第三，把功劳让出去。功成事遂，百姓皆谓：我自然。当大家觉得这是“我们一起做到的”，系统就会开始为你让路。
水从不和石头吵架。天下难事，必作于易；天下大事，必作于细。`);

	qa(9, "软技能看不见、摸不着，怎么证明我的价值？", "把“我会什么”，换成“我改变了什么”。",
		["用结果说话：问题—判断—行动—改变", "留下判断的痕迹：我为什么选这个，不选那个", "讲成故事：六步结构，就是你的价值证明"], `
【Q9 · 约 3 分钟】
软技能之所以“看不见”，是因为我们习惯用“我会什么”来描述自己。“我沟通能力强”——这句话确实什么也证明不了。
换一个句式：我改变了什么。
还记得洪水的故事吗？预测—判断—行动—结果。你也可以用同样的四步来描述自己的工作：我发现了什么问题；我做了什么判断；我组织了什么行动；最后改变了什么。
第二，留下判断的痕迹。执行的痕迹很容易留——文档、代码、报表。判断的痕迹很少有人留。下次做决定的时候，写三行字：我为什么选这个，为什么不选那个，我在赌什么。这三行字，就是你判断力的证据。
第三，把它讲成故事。今天的六步结构——Hook、Story、Conflict、Highlight、Insight、Offer——就是你的价值证明模板。面试、述职、争取资源，都能用。`);

	qa(10, "怎么找到“值得做”的问题？", "去低处找——真实的需求，往往在别人不愿去的地方。",
		["别只问“完成了什么”，要问“改变了什么”", "用四个问题过滤：在意、判断、利他、承担", "先小范围验证：真的有人需要吗？"], `
【Q10 · 约 3 分钟】
这是今天最难、也最重要的问题。马蒂亚斯被问到最担忧什么的时候，他说的也不是模型和算力，而是：面对这么多机会，怎么选出最重要的问题。
我分享三个办法。
第一，换一个提问方式。不要问“这件事能不能做成”——有了 AI，几乎都能做成。要问“做成之后，改变了什么？谁的行动会不一样？”如果答不上来，它大概不值得做。
第二，用今天的四个问题当筛子：我在意吗？我的判断在这里比别人准吗？谁会因此变好？我愿意为结果签字吗？四个都是“是”的事情，非常少——少，就对了。
第三，去低处找。上善若水，处众人之所恶。大家都往光鲜的地方挤，真实的需求往往留在那些琐碎、麻烦、没人愿意碰的角落里。
最后，先小范围试一试：真的有人需要吗？有人愿意为它付出时间或金钱吗？别让 AI 帮你高效地造出一个没人要的东西。`);

	s = add("TITLE", "谢谢 · 愿你如水", `
【结束 · 约 1 分钟】
谢谢大家今天的时间，也谢谢大家的问题。
最后，把这八个字再送给大家一次：上善若水，利万物而不争。
愿我们都能在 AI 的大水里，修好自己的河道。
我们下次课见。
【提示】请在页面下方补上联系方式或社群二维码。`);
	s.addText("上善若水，水善利万物而不争。", { placeholder: "body" });
	T(s, "职场提升俱乐部 Career Up Club · 柔性领导力系列课程", 0.7, 0.75, 7, 0.35, { fontSize: 14, color: C.accent6 });
	T(s, "Lead Like Water — Timeless Wisdom in the AI Era", 0.7, 3.4, 7, 0.35, { fontSize: 16, color: C.accent6 });
	T(s, "[待补充：联系方式 / 社群二维码]", 0.7, 4.75, 5, 0.3, { fontSize: 12, color: C.background1 });

	// ---------- write, then theme ----------
	await pres.writeFile({ fileName: OUT });
	await writeTheme(OUT, THEME);
	console.log("Wrote " + OUT);
})().catch((e) => { console.error(e); process.exit(1); });

// pptxgenjs cannot write theme colours or an East Asian theme font; patch ppt/theme/theme1.xml after writing.
async function writeTheme(file, theme) {
	if (process.env.PPTX_APPLY_THEME) await require(process.env.PPTX_APPLY_THEME).applyTheme(file, theme);
	const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
	const zip = await JSZip.loadAsync(fs.readFileSync(file));
	const part = "ppt/theme/theme1.xml";
	let xml = await zip.file(part).async("string");
	if (!process.env.PPTX_APPLY_THEME) {
		const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
		const scheme = `<a:clrScheme name="${theme.name}">` + slots.map((k) => `<a:${k}><a:srgbClr val="${theme.colors[k]}"/></a:${k}>`).join("") + "</a:clrScheme>";
		xml = xml.replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme).replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${theme.name}"`);
	}
	xml = xml.replace(/<a:ea typeface="[^"]*"\s*\/>/g, `<a:ea typeface="${theme.bodyFontFace}"/>`);
	zip.file(part, xml);
	for (const name of Object.keys(zip.files)) if (zip.files[name].dir) delete zip.files[name]; // folder entries are not parts (zip.remove would take their files too)
	fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}
