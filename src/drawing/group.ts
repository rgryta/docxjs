import { strictInteger } from '../utils';
import { DomType, OpenXmlElementBase } from '../document/dom';

const maxPercent = 1e6;
const pxPerPt = 4 / 3;

export interface GroupFrame {
	ox: number;
	oy: number;
	cw: number;
	ch: number;
	width: number;
	height: number;
	flipH: boolean;
	flipV: boolean;
}

export interface GroupPlace {
	left: string;
	top: string;
	width: string;
	height: string;
	rotation: number;
}

export interface GroupChild {
	place: GroupPlace;
	width: number;
	height: number;
}

export class DrawingGroup extends OpenXmlElementBase {
	type: DomType = DomType.Group;
	width: number;
	height: number;
	place?: GroupPlace;
	rotation?: number;
}

export function groupFrame([ox, oy]: number[], [cw, ch]: number[], width: number, height: number, flipH: boolean, flipV: boolean): GroupFrame {
	// a group of lines on one axis has no height or width
	return [ox, oy, cw, ch, width, height].every(Number.isFinite) && Math.min(cw, ch, width, height) >= 0 && cw + ch > 0 && width + height > 0
		? { ox, oy, cw, ch, width, height, flipH, flipV }
		: null;
}

export function placeChild(frame: GroupFrame, [x, y, w, h]: number[], line: boolean, rotation: number): GroupChild {
	if (![x, y, w, h].every(Number.isFinite) || w < 0 || h < 0 || !(line ? w + h > 0 : w > 0 && h > 0))
		return null;

	const left = frame.flipH ? frame.cw - (x - frame.ox) - w : x - frame.ox;
	const top = frame.flipV ? frame.ch - (y - frame.oy) - h : y - frame.oy;
	const [cw, ch] = [frame.cw || 1, frame.ch || 1];
	const percents = [left / cw, top / ch, w / cw, h / ch].map(v => v * 100);

	if (!percents.every(v => Math.abs(v) <= maxPercent))
		return null;

	const [l, t, pw, ph] = percents.map(v => `${v.toFixed(4)}%`);

	return {
		place: { left: l, top: t, width: pw, height: ph, rotation: frame.flipH != frame.flipV ? -rotation : rotation },
		width: w / cw * frame.width,
		height: h / ch * frame.height,
	};
}

export function integerPair(text: string, fallback: number[]): number[] {
	const parts = text.split(",");
	return fallback.map((value, i) => {
		const part = parts[i]?.trim();
		return !part ? value : strictInteger(part);
	});
}

/** Where an arrowhead's base sits on the line, in px: the line stops there so its end does not poke past the tip. */
export function arrowBase(tip: number[], from: number[], strokeWidth: number, lengthScale: number): number[] {
	const [tx, ty] = tip.map(v => v * pxPerPt);
	const [fx, fy] = from.map(v => v * pxPerPt);
	const distance = Math.hypot(tx - fx, ty - fy) || 1;
	const length = lengthScale * strokeWidth * pxPerPt;
	return [tx - (tx - fx) / distance * length, ty - (ty - fy) / distance * length];
}

export function arrowHead(tip: number[], from: number[], strokeWidth: number, widthScale: number, lengthScale: number): string {
	const [tx, ty] = tip.map(v => v * pxPerPt);
	const [fx, fy] = from.map(v => v * pxPerPt);
	const distance = Math.hypot(tx - fx, ty - fy) || 1;
	const [ux, uy] = [(tx - fx) / distance, (ty - fy) / distance];
	const half = widthScale * strokeWidth * pxPerPt / 2;
	const [bx, by] = arrowBase(tip, from, strokeWidth, lengthScale);
	const f = (v: number) => `${parseFloat(v.toFixed(2))}`;

	return `M ${f(tx)} ${f(ty)} L ${f(bx - uy * half)} ${f(by + ux * half)} L ${f(bx + uy * half)} ${f(by - ux * half)} Z`;
}
