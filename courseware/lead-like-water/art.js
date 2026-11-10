// Generated artwork for the deck: water / ink-wash backgrounds and story scenes.
// Everything is drawn as SVG and rasterised with sharp, so the deck needs no external images.
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const fa = require("react-icons/fa");

const INK = "#0B2529", DEEP = "#0F4C5C", WATER = "#1B7F8C", AQUA = "#5FB3B3", PALE = "#BFE0DA", MIST = "#EEF5F4", SEAL = "#C8442F", SAND = "#E0A458";

// A filled band whose top edge is a soft double sine wave.
function wave(w, h, y0, amp, len, phase, fill, op) {
	let d = `M0 ${h} L0 ${y0.toFixed(1)}`;
	for (let x = 0; x <= w; x += 16) {
		const y = y0 + amp * Math.sin((x / len) * 2 * Math.PI + phase) + amp * 0.35 * Math.sin((x / (len * 0.37)) * 2 * Math.PI + phase * 1.7);
		d += ` L${x} ${y.toFixed(1)}`;
	}
	return `<path d="${d} L${w} ${h} Z" fill="${fill}" fill-opacity="${op}"/>`;
}

// A mountain ridge: a few overlapping sines give irregular peaks.
function ridge(w, h, y0, amp, seed, fill, op, extra = "") {
	let d = `M0 ${h} L0 ${y0}`;
	for (let x = 0; x <= w; x += 12) {
		const t = x / w;
		const y = y0 - amp * (0.55 * Math.abs(Math.sin(t * 3.1 * Math.PI + seed)) + 0.3 * Math.abs(Math.sin(t * 7.3 * Math.PI + seed * 2.1)) + 0.15 * Math.sin(t * 17 * Math.PI + seed * 3.3));
		d += ` L${x} ${y.toFixed(1)}`;
	}
	return `<path d="${d} L${w} ${h} Z" fill="${fill}" fill-opacity="${op}" ${extra}/>`;
}

const svg = (w, h, body) => `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${body}</svg>`;
const vgrad = (id, stops) => `<linearGradient id="${id}" x1="0" y1="0" x2="0" y2="1">${stops.map(([o, c, a = 1]) => `<stop offset="${o}" stop-color="${c}" stop-opacity="${a}"/>`).join("")}</linearGradient>`;

function bgTitle() {
	const w = 1920, h = 1080;
	return svg(w, h, `<defs>${vgrad("g", [[0, "#07191D"], [0.6, "#0D3F4C"], [1, "#12606E"]])}
		<radialGradient id="glow"><stop offset="0" stop-color="#F4EBD3" stop-opacity="0.35"/><stop offset="1" stop-color="#F4EBD3" stop-opacity="0"/></radialGradient></defs>
		<rect width="${w}" height="${h}" fill="url(#g)"/>
		<circle cx="1540" cy="270" r="330" fill="url(#glow)"/><circle cx="1540" cy="270" r="112" fill="#F4EBD3" fill-opacity="0.93"/>
		${ridge(w, h, 760, 250, 0.7, "#0C3640", 0.85)}${ridge(w, h, 810, 170, 2.2, "#092A32", 0.95)}
		${wave(w, h, 835, 14, 520, 0.3, WATER, 0.55)}${wave(w, h, 890, 16, 430, 1.4, "#2A93A0", 0.45)}
		${wave(w, h, 950, 14, 610, 2.6, AQUA, 0.32)}${wave(w, h, 1010, 12, 380, 0.9, PALE, 0.22)}
		${[0, 1, 2, 3, 4].map((i) => `<rect x="${1540 - 70 + i * 9}" y="${850 + i * 34}" width="${140 - i * 18}" height="5" rx="2.5" fill="#F4EBD3" fill-opacity="${0.42 - i * 0.07}"/>`).join("")}`);
}

function bgSection() {
	const w = 1920, h = 1080, cx = 1560, cy = 790;
	let rip = "";
	for (let i = 1; i <= 10; i++) rip += `<ellipse cx="${cx}" cy="${cy}" rx="${i * 92}" ry="${i * 92 * 0.3}" fill="none" stroke="${PALE}" stroke-width="${i < 4 ? 4 : 3}" stroke-opacity="${(0.62 - i * 0.055).toFixed(3)}"/>`;
	return svg(w, h, `<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#11596A"/><stop offset="1" stop-color="#081E22"/></linearGradient></defs>
		<rect width="${w}" height="${h}" fill="url(#g)"/>${rip}
		<path d="M${cx} ${cy - 330} C${cx + 46} ${cy - 262} ${cx + 58} ${cy - 228} ${cx} ${cy - 196} C${cx - 58} ${cy - 228} ${cx - 46} ${cy - 262} ${cx} ${cy - 330} Z" fill="${PALE}" fill-opacity="0.9"/>`);
}

function bgQuote() {
	const w = 1920, h = 1080;
	return svg(w, h, `<defs>${vgrad("g", [[0, "#07191D"], [1, "#0F4855"]])}</defs><rect width="${w}" height="${h}" fill="url(#g)"/>
		${wave(w, h, 925, 16, 560, 0.4, WATER, 0.4)}${wave(w, h, 975, 14, 430, 1.7, AQUA, 0.26)}${wave(w, h, 1025, 12, 640, 2.9, PALE, 0.18)}`);
}

function bgQa() {
	const w = 1920, h = 1080, cx = 1700, cy = 1010;
	let rip = "";
	for (let i = 1; i <= 8; i++) rip += `<ellipse cx="${cx}" cy="${cy}" rx="${i * 105}" ry="${i * 105 * 0.3}" fill="none" stroke="${WATER}" stroke-width="3" stroke-opacity="${(0.2 - i * 0.02).toFixed(3)}"/>`;
	return svg(w, h, `<rect width="${w}" height="${h}" fill="${MIST}"/>${rip}`);
}

// Story scenes, portrait 1200 x 1688 (a 4.0" x 5.625" half-bleed panel).
function sceneNight() {
	const w = 1200, h = 1688;
	const back = [[-40, 260, 760], [170, 190, 1060], [330, 240, 880], [560, 170, 1180], [730, 250, 940], [960, 280, 1100]];
	const front = [[0, 230, 620], [210, 270, 900], [460, 190, 560], [630, 300, 1040], [910, 170, 720], [1060, 160, 500]];
	let b = back.map(([x, bw, bh]) => `<rect x="${x}" y="${h - bh}" width="${bw}" height="${bh}" fill="#0E3038"/>`).join("");
	let lit = "";
	front.forEach(([x, bw, bh], bi) => {
		b += `<rect x="${x}" y="${h - bh}" width="${bw}" height="${bh}" fill="#061619"/>`;
		for (let r = 0; r * 86 + 110 < bh; r++) for (let c = 0; c * 64 + 60 < bw; c++) {
			const wx = x + 30 + c * 64, wy = h - bh + 50 + r * 86;
			if (bi === 3 && r === 3 && c === 2) lit = `<circle cx="${wx + 15}" cy="${wy + 22}" r="150" fill="url(#lamp)"/><rect x="${wx}" y="${wy}" width="30" height="44" fill="${SAND}"/>`;
			else b += `<rect x="${wx}" y="${wy}" width="30" height="44" fill="#12333B"/>`;
		}
	});
	const stars = [[120, 180], [300, 90], [460, 240], [640, 130], [1050, 150], [1120, 420], [220, 430], [760, 330]].map(([x, y]) => `<circle cx="${x}" cy="${y}" r="3" fill="#F4EBD3" fill-opacity="0.7"/>`).join("");
	return svg(w, h, `<defs>${vgrad("g", [[0, "#06171A"], [1, "#1A4E5B"]])}<radialGradient id="lamp"><stop offset="0" stop-color="${SAND}" stop-opacity="0.55"/><stop offset="1" stop-color="${SAND}" stop-opacity="0"/></radialGradient></defs>
		<rect width="${w}" height="${h}" fill="url(#g)"/>${stars}<circle cx="930" cy="250" r="64" fill="#F4EBD3" fill-opacity="0.9"/>${b}${lit}`);
}

function sceneFlood() {
	const w = 1200, h = 1688;
	const house = (x, y, s) => `<g transform="translate(${x} ${y}) scale(${s})"><path d="M-10 0 L60 -52 L130 0 Z" fill="#061619"/><rect x="0" y="0" width="120" height="78" fill="#0A2228"/><rect x="72" y="22" width="26" height="30" fill="${SAND}"/></g>`;
	return svg(w, h, `<defs>${vgrad("g", [[0, "#0E3A45"], [0.42, "#C9925A"], [0.56, "#F3D49C"]])}${vgrad("r", [[0, "#F3D49C"], [0.25, "#8FCBC6"], [1, "#2A93A0"]])}</defs>
		<rect width="${w}" height="${h}" fill="url(#g)"/><circle cx="800" cy="850" r="96" fill="#FBE9BE"/>
		${ridge(w, h, 900, 210, 1.3, "#2C6F78", 0.9)}${ridge(w, h, 960, 130, 2.9, "#1B5560", 1)}
		<rect x="0" y="950" width="${w}" height="${h - 950}" fill="#12404A"/>
		<path d="M690 950 C830 1080 380 1170 470 1320 C550 1450 260 1540 60 1688 L900 1688 C960 1530 800 1420 730 1310 C650 1180 980 1070 780 950 Z" fill="url(#r)"/>
		${wave(w, h, 1600, 12, 380, 0.6, "#0B2E36", 0.55)}
		${house(880, 1190, 0.8)}${house(1010, 1290, 1)}${house(900, 1420, 1.2)}${house(150, 1250, 0.9)}${house(60, 1390, 1.15)}`);
}

function sceneYu() {
	const w = 1200, h = 1688, blur = (n) => `filter="url(#b${n})"`;
	return svg(w, h, `<defs>${vgrad("g", [[0, "#F4F8F7"], [1, "#D5E6E3"]])}<filter id="b1"><feGaussianBlur stdDeviation="9"/></filter><filter id="b2"><feGaussianBlur stdDeviation="4"/></filter></defs>
		<rect width="${w}" height="${h}" fill="url(#g)"/><circle cx="900" cy="330" r="62" fill="${SEAL}" fill-opacity="0.9"/>
		${ridge(w, h, 900, 420, 0.4, "#A9C2C1", 0.7, blur(1))}${ridge(w, h, 1040, 380, 1.9, "#6F979A", 0.8, blur(2))}${ridge(w, h, 1180, 300, 3.4, "#2F5E69", 0.92, blur(2))}
		<rect x="0" y="1170" width="${w}" height="${h - 1170}" fill="#1D4954"/>
		<path d="M520 1170 C700 1260 250 1330 380 1440 C470 1520 240 1600 120 1688 L760 1688 C820 1590 660 1500 600 1430 C520 1330 860 1260 640 1170 Z" fill="#F4F8F7"/>`);
}

function sceneSea() {
	const w = 1200, h = 1688;
	let sea = "";
	const tones = [["#12606E", 1], [WATER, 1], ["#2A93A0", 1], ["#46A6AC", 1], [AQUA, 1], ["#8FCBC6", 1], [PALE, 1]];
	tones.forEach(([c, a], i) => (sea += wave(w, h, 900 + i * 115, 10 + i * 3, 420 + i * 90, i * 1.3, c, a)));
	let glint = "";
	for (let i = 0; i < 7; i++) glint += `<rect x="${600 - (150 - i * 16) / 2}" y="${920 + i * 60}" width="${150 - i * 16}" height="7" rx="3.5" fill="#FBE9BE" fill-opacity="${0.6 - i * 0.07}"/>`;
	return svg(w, h, `<defs>${vgrad("g", [[0, "#0E4756"], [0.38, "#D49A5A"], [0.53, "#F6DCA6"]])}</defs>
		<rect width="${w}" height="${h}" fill="url(#g)"/><circle cx="600" cy="890" r="150" fill="#FBE9BE"/>${sea}${glint}`);
}

// Two glasses: stirred (murky) and settled (clear, sediment at the bottom). Transparent background.
function glasses() {
	const w = 1600, h = 900;
	const glass = (x, id) => `<path d="M${x} 190 L${x + 300} 190 L${x + 268} 760 L${x + 32} 760 Z" fill="url(#${id})" stroke="${DEEP}" stroke-width="10" stroke-linejoin="round"/><path d="M${x - 12} 120 L${x} 190 M${x + 312} 120 L${x + 300} 190" stroke="${DEEP}" stroke-width="10" stroke-linecap="round"/>`;
	let dots = "";
	for (let i = 0; i < 46; i++) dots += `<circle cx="${250 + ((i * 97) % 240)}" cy="${240 + ((i * 61) % 480)}" r="${4 + (i % 4) * 2}" fill="#7A5C36" fill-opacity="0.55"/>`;
	return svg(w, h, `<defs>${vgrad("murky", [[0, "#C6AE84"], [1, "#A88A5C"]])}${vgrad("clear", [[0, "#DDF0EC"], [0.84, "#BFE0DA"], [0.86, "#8A6B3F"], [1, "#7A5C36"]])}</defs>
		${glass(220, "murky")}${dots}${glass(1080, "clear")}
		<path d="M640 475 L900 475" stroke="${SEAL}" stroke-width="14" stroke-linecap="round"/><path d="M880 430 L950 475 L880 520 Z" fill="${SEAL}"/>`);
}

// Iceberg, 1380 x 945 (a 4.6" x 3.15" panel). The left half is kept clear for labels.
function iceberg() {
	const w = 1380, h = 945, wl = 290;
	return svg(w, h, `<defs>${vgrad("sea", [[0, AQUA], [1, "#0B3944"]])}<clipPath id="r"><rect width="${w}" height="${h}" rx="36"/></clipPath></defs>
		<g clip-path="url(#r)"><rect width="${w}" height="${wl}" fill="#E4F0EE"/><rect y="${wl}" width="${w}" height="${h - wl}" fill="url(#sea)"/>
		<path d="M860 ${wl} L935 150 L1000 205 L1060 70 L1130 180 L1175 150 L1240 ${wl} Z" fill="#FFFFFF" stroke="${PALE}" stroke-width="5"/>
		<path d="M1060 70 L1130 180 L1090 ${wl} L1040 ${wl} Z" fill="#DCEBE8"/>
		<path d="M800 ${wl} L1300 ${wl} L1335 480 L1260 650 L1180 760 L1090 880 L1000 800 L900 720 L820 560 L770 420 Z" fill="#DFF1EE" fill-opacity="0.5" stroke="#FFFFFF" stroke-opacity="0.5" stroke-width="4"/>
		<path d="M0 ${wl} L${w} ${wl}" stroke="#FFFFFF" stroke-width="5" stroke-opacity="0.9"/></g>`);
}

const ART = { bgTitle, bgSection, bgQuote, bgQa, sceneNight, sceneFlood, sceneYu, sceneSea, glasses, iceberg };

async function makeArt() {
	const out = {};
	for (const [name, fn] of Object.entries(ART)) {
		const buf = await sharp(Buffer.from(fn())).png({ compressionLevel: 9 }).toBuffer();
		out[name] = "image/png;base64," + buf.toString("base64");
	}
	return out;
}

async function iconPng(name, color = "#FFFFFF") {
	const Icon = fa[name];
	if (!Icon) throw new Error("No such icon: " + name);
	const markup = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color, size: 256 }));
	const buf = await sharp(Buffer.from(markup)).resize(256, 256).png().toBuffer();
	return "image/png;base64," + buf.toString("base64");
}

module.exports = { makeArt, iconPng };
