import { Length } from "./document/common";
import { ParagraphTab } from "./document/paragraph";

interface TabStop {
	pos: number;
	leader: string;
	style: string;
}

interface Registration {
	stops: ParagraphTab[];
	defaultTabSize: Length;
}

interface TabGroup {
	p: HTMLElement;
	spans: HTMLElement[];
	stops: TabStop[];
	clears: number[];
	defaultTabSize: number;
}

const registered = new WeakMap<HTMLElement, Registration>();
const posTolerance = 0.01;
const leaderBorders: Record<string, string> = {
	dot: "1px dotted",
	middleDot: "1px dotted",
	hyphen: "1px dashed",
	underscore: "1px solid",
	heavy: "2px solid",
};

export function registerTabStop(span: HTMLElement, stops: ParagraphTab[], defaultTabSize: Length) {
	registered.set(span, { stops, defaultTabSize });
}

export function mergeTabs(base: ParagraphTab[] | undefined, own: ParagraphTab[] | undefined): ParagraphTab[] | undefined {
	if (!base?.length) return own;
	if (!own?.length) return base;
	const samePos = (a: ParagraphTab, b: ParagraphTab) => Math.abs(lengthToPoint(a.position) - lengthToPoint(b.position)) < posTolerance;
	return [...base.filter(b => !own.some(o => samePos(o, b))), ...own];
}

export function updateTabStop(elem: HTMLElement, tabs: ParagraphTab[], defaultTabSize: Length) {
	registerTabStop(elem, tabs, defaultTabSize);
	measureTabSpans([elem]);
}

export function remeasureTabStops(root: ParentNode, className: string = "docx"): void {
	measureTabSpans(Array.from(root.querySelectorAll<HTMLElement>(`.${className}-tab-stop`)));
}

export function measureTabSpans(spans: HTMLElement[]): void {
	const groups = groupByParagraph(spans.filter(s => registered.has(s) && s.isConnected && s.closest("p")?.getBoundingClientRect().width > 0));

	if (groups.length == 0)
		return;

	groups.forEach(g => g.spans.forEach(reset));

	const probe = document.createElement("div");
	probe.style.cssText = "position:absolute;visibility:hidden;width:100pt";
	(groups[0].p.closest("section") ?? groups[0].p.parentElement).appendChild(probe);
	const scale = probe.getBoundingClientRect().width / 100;
	probe.remove();

	if (!(scale > 0))
		return;

	const rounds = Math.max(...groups.map(g => g.spans.length));

	for (let k = 0; k < rounds; k++) {
		const round = groups.filter(g => k < g.spans.length);

		for (const g of round) {
			const span = g.spans[k];
			span.textContent = "";
			span.style.display = "inline-block";
			span.style.width = "0";
		}

		const measured = round.map(g => measureTab(g, k, scale));

		round.forEach((g, i) => {
			const span = g.spans[k];
			span.style.width = `${measured[i].width.toFixed(2)}pt`;
			const border = leaderBorders[measured[i].leader];
			if (border) span.style.borderBottom = border;
		});
	}
}

function reset(span: HTMLElement) {
	span.removeAttribute("style");
	span.textContent = "\u2003";
}

function groupByParagraph(spans: HTMLElement[]): TabGroup[] {
	const sorted = [...spans].sort((a, b) => a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1);
	const groups = new Map<HTMLElement, TabGroup>();

	for (const span of sorted) {
		const p = span.closest("p");
		let group = groups.get(p);

		if (!group) {
			const { stops, defaultTabSize } = registered.get(span);
			const all = (stops ?? []).map(t => ({ pos: lengthToPoint(t.position), leader: t.leader, style: t.style }));
			group = {
				p,
				spans: [],
				stops: all.filter(t => t.style != "clear").sort((a, b) => a.pos - b.pos),
				clears: all.filter(t => t.style == "clear").map(t => t.pos),
				defaultTabSize: lengthToPoint(defaultTabSize) || 36,
			};
			groups.set(p, group);
		}

		group.spans.push(span);
	}

	return [...groups.values()];
}

function nextStop(g: TabGroup, x: number): TabStop {
	const declared = g.stops.find(t => t.pos > x + posTolerance);

	if (declared)
		return declared;

	const size = g.defaultTabSize;
	let pos = (Math.floor((x + posTolerance) / size) + 1) * size;

	while (g.clears.some(c => Math.abs(c - pos) < posTolerance))
		pos += size;

	return { pos, leader: "none", style: "left" };
}

function measureTab(g: TabGroup, k: number, scale: number): { width: number; leader: string } {
	const span = g.spans[k];
	const pbb = g.p.getBoundingClientRect();
	const pcs = getComputedStyle(g.p);
	const cssPx = scale * 0.75; // computed lengths are in the element's unzoomed CSS px
	const origin = pbb.left - parseFloat(pcs.marginLeft) * cssPx;
	const sbb = span.getBoundingClientRect();
	const x = (sbb.left - origin) / scale;
	const tab = nextStop(g, x);
	const segment = (stopAt?: string) => segmentWidth(span, g.spans[k + 1], g.p, sbb, stopAt) / scale;

	let width: number;

	switch (tab.style) {
		case "right":
		case "end":
			width = tab.pos - x - segment();
			break;
		case "center":
			width = tab.pos - x - segment() / 2;
			break;
		case "decimal":
			width = tab.pos - x - segment(".");
			break;
		default:
			width = tab.pos - x;
	}

	const room = (pbb.right - parseFloat(pcs.paddingRight) * cssPx - sbb.left) / scale;

	return { width: Math.max(0, Math.min(width, room)), leader: tab.leader };
}

function segmentWidth(span: HTMLElement, next: HTMLElement | undefined, p: HTMLElement, sbb: DOMRect, stopAt?: string): number {
	const range = document.createRange();
	range.setStartAfter(span);

	if (next)
		range.setEndBefore(next);
	else
		range.setEnd(p, p.childNodes.length);

	if (stopAt) {
		const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
		walker.currentNode = span;

		for (let n = walker.nextNode(); n && range.isPointInRange(n, 0); n = walker.nextNode()) {
			const i = n.textContent.indexOf(stopAt);

			if (i >= 0) {
				range.setEnd(n, i);
				break;
			}
		}
	}

	// the empty inline-block span sits on the baseline: boxes starting above it are on the first line
	const rights = Array.from(range.getClientRects()).filter(r => r.width > 0 && r.top < sbb.bottom).map(r => r.right);

	return rights.length ? Math.max(0, Math.max(...rights) - sbb.left) : 0;
}

function lengthToPoint(length: Length): number {
	return parseFloat(length);
}
