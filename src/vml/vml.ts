import { DocumentParser } from '../document-parser';
import { convertLength, LengthUsage } from '../document/common';
import { OpenXmlElementBase, DomType } from '../document/dom';
import xml from '../parser/xml-parser';
import { clamp } from '../utils';

const vmlNamespace = "urn:schemas-microsoft-com:vml";

const modifierTags = new Set(["stroke", "fill", "path", "formulas", "f", "handles", "h", "shadow", "textbox",
	"imagedata", "shapetype", "lock", "wrap", "anchorlock"]);

export type VmlSvgTag = "path" | "rect" | "ellipse" | "line" | "image" | "g";

export interface VmlSvgPart {
	tagName: VmlSvgTag;
	attrs: Record<string, string>;
}

export interface VmlBox {
	position?: string;
	hAlign?: string;
	vAlign?: string;
	hRelative?: string;
	vRelative?: string;
	x: number;
	y: number;
	width?: number;
	height?: number;
	zIndex?: number;
	rotation?: number;
	hidden: boolean;
}

export interface VmlTextBox {
	padding: [string, string, string, string];
	justify: string;
}

export interface VmlShapeType {
	path?: string;
	coordsize?: string;
	coordorigin?: string;
	filled?: string;
	stroked?: string;
	fillcolor?: string;
	strokecolor?: string;
}

export interface VmlSubpath {
	d: string;
	filled: boolean;
	stroked: boolean;
}

export class VmlElement extends OpenXmlElementBase {
	type: DomType = DomType.VmlElement;
	notDrawn?: string;
	box: VmlBox;
	parts: VmlSvgPart[] = [];
	viewBox?: string;
	flip?: string;
	textBox?: VmlTextBox;
	wrapType?: string;
	imageHref?: {
		id: string,
		title: string
	}
}

export function parseVmlElement(elem: Element, parser: DocumentParser, watermarkAllowed = false): VmlElement {
	if (elem.namespaceURI != vmlNamespace)
		return null;

	if (elem.localName == "shapetype") {
		const id = xml.attr(elem, "id");
		id && parser.vmlShapeTypes.set(id, readShapeType(elem));
		return null;
	}

	switch (elem.localName) {
		case "rect":
		case "oval":
		case "line":
		case "shape":
			break;

		default:
			return modifierTags.has(elem.localName) ? null : notDrawn("vml-element");
	}

	let shapeType: VmlShapeType = {};
	const typeRef = elem.localName == "shape" ? xml.attr(elem, "type") : null;

	if (typeRef) {
		shapeType = typeRef.startsWith("#") ? parser.vmlShapeTypes.get(typeRef.substring(1)) : undefined;

		if (!shapeType && !xml.elements(elem).some(el => el.localName == "imagedata"))
			return notDrawn("vml-type");

		shapeType ??= {};
	}

	const result = new VmlElement();

	let stroke: Element = null, fill: Element = null, textBox: Element = null, imageData: Element = null;
	let hasTextPath = false;
	const innerMarks: VmlSvgPart[] = [];

	for (const el of xml.elements(elem)) {
		switch (el.localName) {
			case "stroke": stroke = el; break;
			case "fill": fill = el; break;
			case "textbox": textBox = el; break;
			case "imagedata": imageData = el; break;
			case "textpath": hasTextPath = true; break;

			case "wrap":
				result.wrapType = xml.attr(el, "type");
				break;

			default:
				if (el.namespaceURI == vmlNamespace && !modifierTags.has(el.localName))
					innerMarks.push(markPart("vml-element"));
				break;
		}
	}

	if (hasTextPath)
		return notDrawn(watermarkAllowed && elem.localName == "shape" ? "watermark" : "vml-element");

	const style = parseVmlStyle(xml.attr(elem, "style"));
	result.box = parseBox(style);

	const imageId = imageData && xml.attr(imageData, "id");
	const paint = parsePaint(elem, shapeType, stroke, fill);

	if (imageId) {
		result.imageHref = { id: imageId, title: xml.attr(imageData, "title") };
		result.parts.push({ tagName: "image", attrs: { width: "100%", height: "100%", preserveAspectRatio: "none" } });
	} else {
		switch (elem.localName) {
			case "rect":
				result.parts.push({ tagName: "rect", attrs: { width: "100%", height: "100%", ...paintAttrs(paint, true, true) } });
				break;

			case "oval":
				result.parts.push({ tagName: "ellipse", attrs: { cx: "50%", cy: "50%", rx: "50%", ry: "50%", ...paintAttrs(paint, true, true) } });
				break;

			case "line":
				result.parts.push(parseLine(elem, result.box, paint));
				break;

			case "shape": {
				const pathText = xml.attr(elem, "path") || shapeType.path;
				const subpaths = pathText ? parseVmlPath(pathText) : null;

				if (subpaths) {
					result.viewBox = parseViewBox(xml.attr(elem, "coordsize") ?? shapeType.coordsize, xml.attr(elem, "coordorigin") ?? shapeType.coordorigin);
					for (const subpath of subpaths) {
						result.parts.push({
							tagName: "path",
							attrs: { d: subpath.d, "vector-effect": "non-scaling-stroke", ...paintAttrs(paint, subpath.filled, subpath.stroked) }
						});
					}
				} else if (!textBox) {
					return notDrawn(pathText ? "vml-path" : "vml-shape");
				} else if (pathText) {
					innerMarks.push(markPart("vml-path"));
				}
				break;
			}
		}
	}

	result.parts.push(...innerMarks);

	const flip = style.get("flip");
	if (flip)
		result.flip = flip.length > 1 ? "scale(-1, -1)" : flip == "x" ? "scale(-1, 1)" : "scale(1, -1)";

	if (textBox) {
		result.textBox = parseTextBox(textBox, style.get("v-text-anchor"));
		const content = xml.element(textBox, "txbxContent");
		content && result.children.push(...parser.parseBodyElements(content));
	}

	return result;
}

function notDrawn(reason: string): VmlElement {
	const result = new VmlElement();
	result.notDrawn = reason;
	return result;
}

function markPart(reason: string): VmlSvgPart {
	return { tagName: "g", attrs: { "data-docx-not-drawn": reason } };
}

function readShapeType(elem: Element): VmlShapeType {
	const result: VmlShapeType = {};

	for (const key of ["path", "coordsize", "coordorigin", "filled", "stroked", "fillcolor", "strokecolor"] as const) {
		const value = xml.attr(elem, key);
		if (value)
			result[key] = value;
	}

	return result;
}

const lengthPattern = /^[+-]?(\d+\.?\d*|\.\d+)(pt|px|in|cm|mm|pc|pi)?$/;
const styleRules = new Map<string, RegExp>([
	["position", /^(absolute|relative|static)$/],
	["left", lengthPattern],
	["top", lengthPattern],
	["margin-left", lengthPattern],
	["margin-top", lengthPattern],
	["width", lengthPattern],
	["height", lengthPattern],
	["z-index", /^-?\d{1,10}$/],
	["rotation", /^-?\d+(\.\d+)?$/],
	["flip", /^(x|y|x y|y x)$/],
	["visibility", /^(hidden|visible|inherit)$/],
	["mso-position-horizontal", /^(absolute|left|center|right|inside|outside|top|bottom)$/],
	["mso-position-vertical", /^(absolute|left|center|right|inside|outside|top|bottom)$/],
	["mso-position-horizontal-relative", /^(page|margin|text|char|(left|right|inner|outer)-margin-area)$/],
	["mso-position-vertical-relative", /^(page|margin|text|line|(top|bottom|inner|outer)-margin-area)$/],
	["v-text-anchor", /^(top|middle|bottom)(-center|-baseline|-center-baseline)?$/],
]);

export function parseVmlStyle(text: string): Map<string, string> {
	const result = new Map<string, string>();

	for (const item of (text ?? "").split(";")) {
		const colon = item.indexOf(":");
		if (colon > 0)
			result.set(item.substring(0, colon).trim().toLowerCase(), item.substring(colon + 1).trim().toLowerCase());
	}

	for (const [key, value] of result) {
		if (!styleRules.get(key)?.test(value))
			result.delete(key);
	}

	return result;
}

function lengthPoints(value: string, unitlessAsPixels = false): number {
	const match = value == null ? null : lengthPattern.exec(value.trim().toLowerCase());

	if (!match)
		return null;

	const number = value.trim().replace(/^\+/, "").replace(/[a-z]+$/i, "");

	if (!match[2]) {
		if (unitlessAsPixels)
			return clamp(parseFloat(number) * 0.75, -10000, 10000);
		return parseFloat(number) == 0 ? 0 : null;
	}

	const css = convertLength(number + match[2], LengthUsage.Point);
	return clamp(parseFloat(css) * (css.endsWith("px") ? 0.75 : 1), -10000, 10000);
}

function parseBox(style: Map<string, string>): VmlBox {
	const position = style.get("position");
	const hAlign = style.get("mso-position-horizontal");
	const zIndex = style.get("z-index");
	const rotation = style.get("rotation");

	return {
		position,
		hAlign,
		vAlign: style.get("mso-position-vertical"),
		hRelative: style.get("mso-position-horizontal-relative"),
		vRelative: style.get("mso-position-vertical-relative"),
		x: lengthPoints(style.get("margin-left")) ?? lengthPoints(style.get("left")) ?? 0,
		y: lengthPoints(style.get("margin-top")) ?? lengthPoints(style.get("top")) ?? 0,
		width: nonNegative(lengthPoints(style.get("width"))),
		height: nonNegative(lengthPoints(style.get("height"))),
		zIndex: zIndex ? parseInt(zIndex) : undefined,
		rotation: rotation ? parseFloat(rotation) : undefined,
		hidden: style.get("visibility") == "hidden",
	};
}

function nonNegative(value: number): number {
	return value == null ? undefined : Math.max(0, value);
}

function parsePoint(value: string): [number, number] {
	const [x, y] = (value ?? "").split(",");
	return [lengthPoints(x || "0", true) ?? 0, lengthPoints(y || "0", true) ?? 0];
}

function parseLine(elem: Element, box: VmlBox, paint: VmlPaint): VmlSvgPart {
	const [x1, y1] = parsePoint(xml.attr(elem, "from"));
	const [x2, y2] = parsePoint(xml.attr(elem, "to"));

	box.x += Math.min(x1, x2);
	box.y += Math.min(y1, y2);
	box.width = Math.abs(x2 - x1);
	box.height = Math.abs(y2 - y1);

	return {
		tagName: "line",
		attrs: {
			x1: x1 <= x2 ? "0" : "100%",
			y1: y1 <= y2 ? "0" : "100%",
			x2: x1 <= x2 ? "100%" : "0",
			y2: y1 <= y2 ? "100%" : "0",
			...paintAttrs(paint, false, true)
		}
	};
}

function parseViewBox(coordsize: string, coordorigin: string): string {
	const integer = (part: string) => /^-?\d+$/.test(part?.trim() ?? "") ? parseInt(part) : null;
	const size = (coordsize ?? "1000,1000").split(",").map(integer);
	const origin = (coordorigin ?? "").split(",").map(integer);

	return [origin[0] ?? 0, origin[1] ?? 0, size[0] > 0 ? size[0] : 1, size[1] > 0 ? size[1] : 1].join(" ");
}

const textAnchor: Record<string, string> = { top: "flex-start", middle: "center", bottom: "flex-end" };
const defaultInset = ["7.2pt", "3.6pt", "7.2pt", "3.6pt"];

function parseTextBox(elem: Element, anchor: string): VmlTextBox {
	const parts = (xml.attr(elem, "inset") ?? "").split(",");
	const [left, top, right, bottom] = defaultInset.map((fallback, i) => {
		const value = lengthPoints(parts[i]);
		return value == null ? fallback : `${formatNumber(Math.max(0, value))}pt`;
	});

	return {
		padding: [top, right, bottom, left],
		justify: textAnchor[anchor?.split("-")[0]] ?? "flex-start",
	};
}

const trueValues = new Set(["t", "true", "1", "yes", "on"]);
const falseValues = new Set(["f", "false", "0", "no", "off"]);

export function vmlBool(value: string, defaultValue: boolean): boolean {
	const key = value?.trim().toLowerCase();
	return trueValues.has(key) ? true : falseValues.has(key) ? false : defaultValue;
}

const namedColors = new Set(("aliceblue antiquewhite aqua aquamarine azure beige bisque black blanchedalmond blue blueviolet " +
	"brown burlywood cadetblue chartreuse chocolate coral cornflowerblue cornsilk crimson cyan darkblue darkcyan " +
	"darkgoldenrod darkgray darkgreen darkgrey darkkhaki darkmagenta darkolivegreen darkorange darkorchid darkred " +
	"darksalmon darkseagreen darkslateblue darkslategray darkslategrey darkturquoise darkviolet deeppink deepskyblue " +
	"dimgray dimgrey dodgerblue firebrick floralwhite forestgreen fuchsia gainsboro ghostwhite gold goldenrod gray " +
	"green greenyellow grey honeydew hotpink indianred indigo ivory khaki lavender lavenderblush lawngreen " +
	"lemonchiffon lightblue lightcoral lightcyan lightgoldenrodyellow lightgray lightgreen lightgrey lightpink " +
	"lightsalmon lightseagreen lightskyblue lightslategray lightslategrey lightsteelblue lightyellow lime limegreen " +
	"linen magenta maroon mediumaquamarine mediumblue mediumorchid mediumpurple mediumseagreen mediumslateblue " +
	"mediumspringgreen mediumturquoise mediumvioletred midnightblue mintcream mistyrose moccasin navajowhite navy " +
	"oldlace olive olivedrab orange orangered orchid palegoldenrod palegreen paleturquoise palevioletred papayawhip " +
	"peachpuff peru pink plum powderblue purple rebeccapurple red rosybrown royalblue saddlebrown salmon sandybrown " +
	"seagreen seashell sienna silver skyblue slateblue slategray slategrey snow springgreen steelblue tan teal " +
	"thistle tomato turquoise violet wheat white whitesmoke yellow yellowgreen").split(" "));

// Windows system colours Word writes for default fills and lines
const systemColors = new Map([["window", "white"], ["windowtext", "black"]]);

export function vmlColor(value: string): string {
	const color = value?.replace(/\s*\[\d+\]\s*$/, "").trim().toLowerCase();

	if (!color)
		return null;

	if (/^#([0-9a-f]{3}|[0-9a-f]{6})$/.test(color) || namedColors.has(color))
		return color;

	if (systemColors.has(color))
		return systemColors.get(color);

	const rgb = /^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$/.exec(color);

	if (rgb && rgb.slice(1).every(c => parseInt(c) <= 255))
		return `rgb(${rgb.slice(1).map(c => parseInt(c)).join(", ")})`;

	return null;
}

const dashPatterns: Record<string, number[]> = {
	dash: [4, 3],
	shortdash: [3, 1],
	dot: [1, 3],
	shortdot: [1, 1],
	dashdot: [4, 3, 1, 3],
	longdash: [8, 3],
	longdashdot: [8, 3, 1, 3],
	shortdashdot: [3, 1, 1, 1],
	shortdashdotdot: [3, 1, 1, 1, 1, 1],
	longdashdotdot: [8, 3, 1, 3, 1, 3],
};

interface VmlPaint {
	fill: string;
	stroke: string;
	strokeWidth: number;
	dash?: number[];
}

function parsePaint(elem: Element, shapeType: VmlShapeType, stroke: Element, fill: Element): VmlPaint {
	const filled = vmlBool(xml.attr(elem, "filled") ?? shapeType.filled, true) && vmlBool(fill && xml.attr(fill, "on"), true);
	const stroked = vmlBool(xml.attr(elem, "stroked") ?? shapeType.stroked, true) && vmlBool(stroke && xml.attr(stroke, "on"), true);
	const strokeColor = (stroke && xml.attr(stroke, "color")) ?? xml.attr(elem, "strokecolor") ?? shapeType.strokecolor ?? "black";
	const weight = (stroke && xml.attr(stroke, "weight")) ?? xml.attr(elem, "strokeweight");
	const dashStyle = stroke && xml.attr(stroke, "dashstyle")?.trim().toLowerCase();

	return {
		fill: filled ? vmlColor(xml.attr(elem, "fillcolor") ?? shapeType.fillcolor ?? "white") : null,
		stroke: stroked ? vmlColor(strokeColor) : null,
		strokeWidth: strokeWeight(weight),
		dash: Object.prototype.hasOwnProperty.call(dashPatterns, dashStyle) ? dashPatterns[dashStyle] : undefined,
	};
}

function strokeWeight(value: string): number {
	const text = value?.trim().toLowerCase();

	if (!text || !lengthPattern.test(text))
		return 0.75;

	const css = /[a-z]$/.test(text) ? convertLength(text.replace(/^\+/, ""), LengthUsage.Point) : convertLength(text, LengthUsage.Emu);
	const points = parseFloat(css) * (css.endsWith("px") ? 0.75 : 1);

	return clamp(isNaN(points) ? 0.75 : points, 0.25, 12);
}

function paintAttrs(paint: VmlPaint, filled: boolean, stroked: boolean): Record<string, string> {
	const result: Record<string, string> = {
		fill: (filled && paint.fill) || "none",
		stroke: (stroked && paint.stroke) || "none",
	};

	if (result.stroke != "none") {
		result["stroke-width"] = `${formatNumber(paint.strokeWidth)}pt`;

		if (paint.dash)
			result["stroke-dasharray"] = paint.dash.map(k => `${formatNumber(k * paint.strokeWidth)}pt`).join(" ");
	}

	return result;
}

export function formatNumber(value: number): string {
	return `${parseFloat(value.toFixed(2))}`;
}

const pathCommands = ["nf", "ns", "m", "l", "c", "x", "e", "r", "t", "v"];
const pathGroupSize: Record<string, number> = { m: 2, l: 2, r: 2, t: 2, c: 6, v: 6 };
const svgCommands: Record<string, [string, string]> = {
	m: ["M", "L"], t: ["m", "l"], l: ["L", "L"], r: ["l", "l"], c: ["C", "C"], v: ["c", "c"]
};
const maxPathTokens = 10000;
const maxPathValue = 2 ** 31;

export function parseVmlPath(text: string): VmlSubpath[] {
	const segments: { command: string, values: number[] }[] = [];
	const token = /(-?\d+)|(,)|(\s+)/y;
	let tokens = 0;
	let afterComma = false;
	let started = false;

	const endSegment = () => {
		const last = segments[segments.length - 1];
		if (last && afterComma)
			last.values.push(0);
		afterComma = false;
		started = false;
	};

	for (let i = 0; i < text.length;) {
		if (++tokens > maxPathTokens)
			return null;

		const command = pathCommands.find(c => text.startsWith(c, i));

		if (command) {
			endSegment();
			segments.push({ command, values: [] });
			i += command.length;
			continue;
		}

		token.lastIndex = i;
		const match = token.exec(text);

		if (!match || !segments.length && !match[3])
			return null;

		i = token.lastIndex;
		const values = segments[segments.length - 1]?.values;

		if (match[1]) {
			const value = parseInt(match[1]);
			if (Math.abs(value) > maxPathValue)
				return null;
			values.push(value);
			afterComma = false;
			started = true;
		} else if (match[2]) {
			if (!started || afterComma)
				values.push(0);
			afterComma = true;
			started = true;
		} else {
			tokens--;
		}
	}

	endSegment();

	const subpaths: { d: string[], filled: boolean, stroked: boolean, drawn: number }[] = [];
	let current: typeof subpaths[number] = null;
	const start = () => {
		current = { d: [], filled: true, stroked: true, drawn: 0 };
		subpaths.push(current);
		return current;
	};

	for (const { command, values } of segments) {
		const size = pathGroupSize[command];

		if (!size) {
			if (values.length)
				return null;

			if (command == "nf")
				(current ?? start()).filled = false;
			else if (command == "ns")
				(current ?? start()).stroked = false;
			else if (command == "x" && current?.d.length)
				current.d.push("Z");

			continue;
		}

		if (!values.length || values.length % size)
			return null;

		if (command == "m") {
			if (!current || current.drawn)
				start();
			else
				current.d = [];
		} else if (!current) {
			start();
		}

		const move = command == "m" || command == "t";

		if (!current.d.length && !move)
			current.d.push("M 0 0");

		const [first, rest] = svgCommands[command];

		for (let k = 0; k < values.length; k += size) {
			current.d.push(`${k ? rest : first} ${values.slice(k, k + size).join(" ")}`);
			if (k || !move)
				current.drawn++;
		}
	}

	const drawn = subpaths.filter(s => s.drawn);
	return drawn.length ? drawn.map(s => ({ d: s.d.join(" "), filled: s.filled, stroked: s.stroked })) : null;
}
