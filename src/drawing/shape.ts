import { DocumentParser } from '../document-parser';
import { DomType, OpenXmlElement, OpenXmlElementBase } from '../document/dom';
import xml from '../parser/xml-parser';
import { DmlTheme } from '../theme/theme';
import { clamp } from '../utils';
import { formatNumber, namedColors, VmlTextBox } from '../vml/vml';

const shapeNamespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const drawnPresets = new Set(["rect", "roundRect", "ellipse", "line"]);
const maxEmu = 1e12;

type Rgb = number[];

export interface DmlColor {
	kind: string;
	values: string[];
	mods: [string, string][];
}

export interface DmlStyleRef {
	idx: number;
	color: DmlColor;
}

export interface ResolvedColor {
	rgb?: Rgb;
	name?: string;
	alpha: number;
}

export class DmlShape extends OpenXmlElementBase {
	type: DomType = DomType.Shape;
	geometry: string;
	radius: number;
	width: number;
	height: number;
	rotation: number;
	flipH: boolean;
	flipV: boolean;
	fill: DmlColor | null | undefined;
	line: { width: number, fill: DmlColor | null | undefined, dash: string };
	fillRef: DmlStyleRef;
	lnRef: DmlStyleRef;
	fontRef: DmlStyleRef;
	textBox?: VmlTextBox;
	autoFit: boolean;
}

export interface DmlAxis {
	relative?: string;
	align?: string;
	offset?: number | null;
}

export interface DmlAnchor {
	inline: boolean;
	nested: boolean;
	wrap: string;
	behindDoc: boolean;
	width: number;
	height: number;
	h: DmlAxis;
	v: DmlAxis;
}

export interface DmlDrawing extends OpenXmlElement {
	anchor: DmlAnchor;
}

export function integer(text: string): number {
	const value = /^-?\d+$/.test(text?.trim() ?? "") ? parseInt(text) : NaN;
	return Math.abs(value) <= maxEmu ? value : null;
}

function emuPoints(text: string): number {
	const value = integer(text);
	return value == null ? null : value / 12700;
}

const descendant = (root: Element, localName: string) => Array.from(root.getElementsByTagNameNS("*", localName));

export function isDrawableShape(wsp: Element): boolean {
	const spPr = xml.element(wsp, "spPr");
	const prst = spPr && xml.elementAttr(spPr, "prstGeom", "prst");
	const bodyPr = xml.element(wsp, "bodyPr");
	const warp = bodyPr && xml.element(bodyPr, "prstTxWarp");

	return drawnPresets.has(prst)
		&& !xml.element(spPr, "custGeom")
		&& ["gradFill", "pattFill", "blipFill"].every(name => !descendant(spPr, name).length)
		&& !xml.element(wsp, "linkedTxbx")
		&& (!bodyPr || (xml.attr(bodyPr, "vert") ?? "horz") == "horz" && !xml.boolAttr(bodyPr, "fromWordArt", false))
		&& (!warp || xml.attr(warp, "prst") == "textNoShape");
}

export function canDrawChoice(choice: Element): boolean {
	const requires = xml.attr(choice, "Requires");
	const children = xml.elements(choice);

	return !!requires
		&& requires.split(" ").every(prefix => prefix && choice.lookupNamespaceURI(prefix) == shapeNamespace)
		// the parser takes one drawing from a Choice, so a Choice with more keeps its Fallback
		&& children.length == 1
		&& children.every(child => child.localName == "drawing" && drawsShapes(child));
}

function drawsShapes(drawing: Element): boolean {
	const wrapper = xml.elements(drawing).find(n => n.localName == "inline" || n.localName == "anchor");
	const graphic = wrapper && xml.element(wrapper, "graphic");
	const graphicData = graphic && xml.element(graphic, "graphicData");
	const shapes = graphicData ? xml.elements(graphicData) : [];

	return shapes.length > 0 && shapes.every(s => s.namespaceURI == shapeNamespace && s.localName == "wsp" && isDrawableShape(s));
}

export function parseShape(wsp: Element, parser: DocumentParser): OpenXmlElement {
	const spPr = xml.element(wsp, "spPr");
	const geometry = xml.element(spPr, "prstGeom");
	const xfrm = xml.element(spPr, "xfrm");
	const ext = xfrm && xml.element(xfrm, "ext");
	const ln = xml.element(spPr, "ln");
	const style = xml.element(wsp, "style");
	const bodyPr = xml.element(wsp, "bodyPr");
	const txbx = xml.element(wsp, "txbx");
	const content = txbx && xml.element(txbx, "txbxContent");

	const result = new DmlShape();
	result.geometry = xml.attr(geometry, "prst");
	result.radius = adjustment(geometry);
	result.width = ext ? emuPoints(xml.attr(ext, "cx")) : null;
	result.height = ext ? emuPoints(xml.attr(ext, "cy")) : null;
	result.rotation = ((xfrm && integer(xml.attr(xfrm, "rot"))) ?? 0) % 21600000 / 60000;
	result.flipH = !!xfrm && xml.boolAttr(xfrm, "flipH", false);
	result.flipV = !!xfrm && xml.boolAttr(xfrm, "flipV", false);
	result.fill = parsePaint(spPr);
	result.line = { width: ln ? emuPoints(xml.attr(ln, "w")) : null, fill: parsePaint(ln), dash: ln && xml.elementAttr(ln, "prstDash", "val") };
	result.fillRef = parseStyleRef(style, "fillRef");
	result.lnRef = parseStyleRef(style, "lnRef");
	result.fontRef = parseStyleRef(style, "fontRef");
	result.autoFit = !!bodyPr && !!xml.element(bodyPr, "spAutoFit");

	if (content) {
		result.children = parser.parseTextBoxContent(content);
		result.textBox = parseBodyProps(bodyPr);
	}

	return result;
}

function adjustment(geometry: Element): number {
	const list = xml.element(geometry, "avLst");
	const guide = list && xml.elements(list, "gd").find(gd => xml.attr(gd, "name") == "adj");
	const match = /^val (-?\d+)$/.exec((guide && xml.attr(guide, "fmla")) ?? "");
	return clamp(match ? parseInt(match[1]) : 16667, 0, 50000) / 100000;
}

const textAnchor = new Map([["t", "flex-start"], ["ctr", "center"], ["b", "flex-end"]]);

function parseBodyProps(bodyPr: Element): VmlTextBox {
	const inset = (name: string, fallback: number) =>
		`${formatNumber(Math.max(0, (bodyPr && integer(xml.attr(bodyPr, name))) ?? fallback) / 12700)}pt`;

	return {
		padding: [inset("tIns", 45720), inset("rIns", 91440), inset("bIns", 45720), inset("lIns", 91440)],
		justify: textAnchor.get(bodyPr && xml.attr(bodyPr, "anchor")) ?? "flex-start",
	};
}

const colorAttrs: Record<string, string[]> = {
	srgbClr: ["val"], schemeClr: ["val"], sysClr: ["lastClr"], prstClr: ["val"], scrgbClr: ["r", "g", "b"], hslClr: ["hue", "sat", "lum"],
};

function parseColor(parent: Element): DmlColor {
	const el = xml.elements(parent).find(e => Object.prototype.hasOwnProperty.call(colorAttrs, e.localName));

	return el ? {
		kind: el.localName,
		values: colorAttrs[el.localName].map(name => xml.attr(el, name)),
		mods: xml.elements(el).map(mod => [mod.localName, xml.attr(mod, "val")]),
	} : null;
}

function parsePaint(parent: Element): DmlColor | null | undefined {
	if (!parent || xml.element(parent, "noFill"))
		return parent ? null : undefined;

	const solid = xml.element(parent, "solidFill");
	return solid ? parseColor(solid) ?? { kind: "", values: [], mods: [] } : undefined;
}

function parseStyleRef(style: Element, name: string): DmlStyleRef {
	const ref = style && xml.element(style, name);
	return ref ? { idx: integer(xml.attr(ref, "idx")), color: parseColor(ref) } : null;
}

export function parseShapeDrawing(node: Element, shape: DmlShape, nested: boolean): DmlDrawing {
	const docPr = xml.element(node, "docPr");

	if (docPr && xml.boolAttr(docPr, "hidden", false))
		return null;

	const extent = xml.element(node, "extent");
	const simplePos = xml.boolAttr(node, "simplePos", false) && xml.element(node, "simplePos");
	const axis = (name: string): DmlAxis => {
		const el = xml.element(node, name);
		const align = el && xml.element(el, "align");
		const offset = el && xml.element(el, "posOffset");
		return el ? { relative: xml.attr(el, "relativeFrom"), align: align?.textContent.trim(), offset: offset ? emuPoints(offset.textContent) : undefined } : {};
	};

	return {
		type: DomType.Drawing,
		children: [shape],
		anchor: {
			inline: node.localName == "inline",
			nested,
			wrap: xml.elements(node).find(n => n.localName.startsWith("wrap"))?.localName,
			behindDoc: xml.boolAttr(node, "behindDoc", false),
			width: extent ? emuPoints(xml.attr(extent, "cx")) : null,
			height: extent ? emuPoints(xml.attr(extent, "cy")) : null,
			h: simplePos ? { relative: "page", offset: emuPoints(xml.attr(simplePos, "x")) } : axis("positionH"),
			v: simplePos ? { relative: "page", offset: emuPoints(xml.attr(simplePos, "y")) } : axis("positionV"),
		}
	};
}

const hex6 = /^[0-9A-Fa-f]{6}$/;
const schemeAliases = new Map([["bg1", "lt1"], ["tx1", "dk1"], ["bg2", "lt2"], ["tx2", "dk2"]]);
const modifiers = new Set(["lumMod", "lumOff", "satMod", "shade", "tint", "alpha"]);

export function resolveColor(spec: DmlColor, theme: DmlTheme, placeholder: ResolvedColor = null): ResolvedColor {
	const color = spec.kind == "schemeClr" && spec.values[0] == "phClr"
		? placeholder && { ...placeholder }
		: baseColor(spec, theme);

	if (!color)
		return null;

	for (const [name, value] of spec.mods) {
		if (!modifiers.has(name))
			continue;

		if (!/^-?\d+$/.test(value ?? "") || name != "alpha" && !color.rgb)
			return null;

		const amount = parseInt(value) / 100000;

		if (name == "alpha")
			color.alpha = clamp(amount, 0, 1);
		else
			color.rgb = modify(color.rgb, name, amount);
	}

	return color;
}

function baseColor(spec: DmlColor, theme: DmlTheme): ResolvedColor {
	const [first] = spec.values;
	const numbers = spec.values.map(integer);

	switch (spec.kind) {
		case "srgbClr":
		case "sysClr":
			return rgbColor(hexRgb(first));

		case "schemeClr":
			return rgbColor(hexRgb(theme?.colorScheme?.colors[schemeAliases.get(first) ?? first]));

		case "prstClr": {
			const name = first?.toLowerCase().replace(/^dk/, "dark").replace(/^lt/, "light").replace(/^med/, "medium");
			return namedColors.has(name) ? { name, alpha: 1 } : null;
		}

		case "scrgbClr":
			return numbers.every(n => n != null) ? rgbColor(numbers.map(n => toGamma(clamp(n / 100000, 0, 1)))) : null;

		case "hslClr":
			return numbers.every(n => n != null)
				? rgbColor(fromHsl([(numbers[0] / 21600000 % 1 + 1) % 1, clamp(numbers[1] / 100000, 0, 1), clamp(numbers[2] / 100000, 0, 1)]))
				: null;
	}

	return null;
}

function rgbColor(rgb: Rgb): ResolvedColor {
	return rgb ? { rgb, alpha: 1 } : null;
}

function hexRgb(value: unknown): Rgb {
	return typeof value == "string" && hex6.test(value) ? [0, 2, 4].map(i => parseInt(value.substring(i, i + 2), 16) / 255) : null;
}

export function formatColor(color: ResolvedColor): string {
	return color.name ?? `#${color.rgb.map(c => Math.round(clamp(c, 0, 1) * 255).toString(16).padStart(2, "0")).join("")}`;
}

function modify(rgb: Rgb, name: string, amount: number): Rgb {
	if (name == "shade")
		return rgb.map(c => toGamma(clamp(toLinear(c) * amount, 0, 1)));

	if (name == "tint")
		return rgb.map(c => toGamma(clamp(1 - (1 - toLinear(c)) * amount, 0, 1)));

	const [h, s, l] = toHsl(rgb);

	switch (name) {
		case "lumMod": return fromHsl([h, s, clamp(l * amount, 0, 1)]);
		case "lumOff": return fromHsl([h, s, clamp(l + amount, 0, 1)]);
		default: return fromHsl([h, clamp(s * amount, 0, 1), l]);
	}
}

const toLinear = (c: number) => c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
const toGamma = (c: number) => c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055;

function toHsl([r, g, b]: Rgb): Rgb {
	const max = Math.max(r, g, b), min = Math.min(r, g, b);
	const l = (max + min) / 2, d = max - min;

	if (!d)
		return [0, 0, l];

	const h = max == r ? (g - b) / d + (g < b ? 6 : 0) : max == g ? (b - r) / d + 2 : (r - g) / d + 4;
	return [h / 6, d / (1 - Math.abs(2 * l - 1)), l];
}

function fromHsl([h, s, l]: Rgb): Rgb {
	const a = s * Math.min(l, 1 - l);
	const k = (n: number) => (n + h * 12) % 12;
	return [0, 8, 4].map(n => l - a * Math.max(-1, Math.min(k(n) - 3, 9 - k(n), 1)));
}
