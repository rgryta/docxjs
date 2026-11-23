// Splits rendered pages whose content is taller than the page, the way Word
// lays out text: whole blocks move to a new page, paragraphs split at a line
// boundary (at least two lines per page), tables split between rows with their
// header rows (<thead>) repeated, and a block marked keep-with-next
// (break-after: avoid) moves along with the block after it.
//
// It measures the layout, so the pages must be in the document and the fonts
// loaded. renderAsync runs it with the `paginate` option; call it yourself
// after renderDocument when you insert the nodes elsewhere.

import { cloneWithTabStops, remeasureTabStops } from "./javascript";

const MAX_PAGES = 2000;
const MIN_SPLIT_HEIGHT = 24;

/**
 * Splits every rendered page (section.<className>) under root until each fits its page height,
 * then numbers the pages, fills their PAGE and NUMPAGES fields and shows the header and footer
 * each page needs. With budgetMs it stops before the next page once the budget is spent, keeping
 * the pages split so far; NUMPAGES then keeps its saved value. Returns false when it stopped early.
 */
export function paginate(root: HTMLElement, className: string = "docx", options?: { budgetMs?: number }): boolean {
	const deadline = options?.budgetMs === undefined ? Infinity : performance.now() + options.budgetMs;
	const pages = Array.from(root.querySelectorAll<HTMLElement>(`section.${className}`));
	const savedCounts = new Map(Array.from(root.querySelectorAll<HTMLElement>('[data-docx-field="NUMPAGES"]')).map((f): [HTMLElement, string] => [f, f.textContent]));
	let done = splitPages(pages, className, deadline);

	if (fillPages(pages, done)) {
		const count = pages.length;
		done = splitPages(pages, className, deadline) && done;

		if (!done)
			restoreCounts(pages, savedCounts);

		if (pages.length != count)
			fillPages(pages, done);
	}

	// pages shown or cloned by the split were hidden or absent when the tabs were first measured
	remeasureTabStops(root, className);
	return done;
}

function splitPages(pages: HTMLElement[], className: string, deadline: number): boolean {
	let number = 0;

	for (let i = 0; i < pages.length; i++) {
		if (pages.length >= MAX_PAGES || performance.now() >= deadline)
			return false;

		// Pages before i are final, so i is this page's final position.
		[number] = pageNumber(pages[i], number);
		showHeaderFooter(pages[i], number);
		const next = splitPage(pages[i], className);

		if (next)
			pages.splice(i + 1, 0, next);
	}

	return true;
}

/**
 * The page's number, and whether Word puts a blank page before it: an odd- or even-page section
 * that would start on the wrong side gets one, which takes a number but is not rendered.
 */
function pageNumber(page: HTMLElement, previous: number): [number, boolean] {
	const start = page.getAttribute("data-page-start");
	const number = start === null ? previous + 1 : Number(start);
	const parity = page.getAttribute("data-page-parity");
	const blank = !!parity && (number % 2 == 0) != (parity == "even");
	return [blank ? number + 1 : number, blank];
}

/** The page number in the section's format, or null for a format it does not write (the field keeps its saved text). */
function formatPageNumber(number: number, format: string): string {
	const letter = (base: number) => String.fromCharCode(base + (number - 1) % 26).repeat(Math.ceil(number / 26));

	switch (format) {
		case null: return String(number);
		case "upperRoman": return roman(number);
		case "lowerRoman": return roman(number).toLowerCase();
		case "upperLetter": return number > 0 ? letter(65) : null;
		case "lowerLetter": return number > 0 ? letter(97) : null;
		default: return null;
	}
}

function roman(number: number): string {
	if (!(number > 0 && number < 4000))
		return String(number);

	const digits: [number, string][] = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
	let result = "";

	for (const [value, text] of digits)
		for (; number >= value; number -= value)
			result += text;

	return result;
}

/**
 * Sets each page's PAGE fields to its number (unless its section numbers pages in another format)
 * and, with withCount, each NUMPAGES field to the page count. Returns true when a text changed its
 * length, because only then can the page layout change.
 */
function fillPages(pages: HTMLElement[], withCount: boolean): boolean {
	let number = 0;
	let changed = false;
	const fill = (field: HTMLElement, text: string) => {
		const before = field.textContent.length;
		return setFieldText(field, text) && before != text.length;
	};

	const numbers = pages.map(page => {
		const [value, blank] = pageNumber(page, number);
		number = value;
		return { value, blank };
	});
	const count = String(pages.length + numbers.filter(n => n.blank).length);

	pages.forEach((page, i) => {
		const text = formatPageNumber(numbers[i].value, page.getAttribute("data-page-format"));

		if (text != null)
			fields(page, "PAGE").forEach(f => changed = fill(f, text) || changed);

		if (withCount)
			fields(page, "NUMPAGES").forEach(f => changed = fill(f, count) || changed);

		showHeaderFooter(page, numbers[i].value);
	});

	return changed;
}

/** Puts back the saved NUMPAGES text; a field on a page split off later takes the text of the same field on the page before. */
function restoreCounts(pages: HTMLElement[], saved: Map<HTMLElement, string>) {
	let previous: string[] = [];

	for (const page of pages) {
		previous = fields(page, "NUMPAGES").map((f, i) => {
			const text = saved.get(f) ?? previous[i];
			text != null && setFieldText(f, text);
			return text;
		});
	}
}

function fields(page: HTMLElement, field: "PAGE" | "NUMPAGES"): HTMLElement[] {
	return Array.from(page.querySelectorAll<HTMLElement>(`[data-docx-field="${field}"]`));
}

/** Puts the text in the field's first text node and empties the others, so the run formatting stays. Returns true when the text changed. */
function setFieldText(field: HTMLElement, text: string): boolean {
	const before = field.textContent;
	const walker = document.createTreeWalker(field, NodeFilter.SHOW_TEXT);
	let first = true;

	for (let n = walker.nextNode() as Text; n; n = walker.nextNode() as Text) {
		n.data = first ? text : "";
		first = false;
	}

	return field.textContent != before;
}

/**
 * Moves the blocks that do not fit on a page to a new page right after it. The flow is the
 * blocks of every article in order, then the notes lists. Returns the new page, or null when
 * the content fits or cannot be split.
 */
export function splitPage(page: HTMLElement, className: string = "docx"): HTMLElement | null {
	const children = Array.from(page.children) as HTMLElement[];
	const articles = children.filter(c => c.tagName === "ARTICLE");

	if (!articles.length)
		return null;

	const style = getComputedStyle(page);
	const pageHeight = parseFloat(style.minHeight);

	if (!pageHeight)
		return null;

	const footer = children.find(c => c.tagName === "FOOTER" && getComputedStyle(c).display !== "none");
	const below = footer ? outerHeight(footer) : 0;
	const limitY = page.getBoundingClientRect().top + pageHeight - (parseFloat(style.paddingBottom) || 0) - below + 0.5;

	const pageSpace = limitY - articles[0].getBoundingClientRect().top;

	if (pageSpace < MIN_SPLIT_HEIGHT)
		return null;

	const flow = [
		...articles.flatMap(article => (Array.from(article.children) as HTMLElement[]).map(block => ({ block, article }))),
		...children.filter(c => c.tagName === "OL").map(block => ({ block, article: null as HTMLElement })),
	];
	// Empty paragraphs (such as the one holding a page break) never start a page.
	// the page also grows by the last block's bottom margin: margins do not collapse in the page's flex column
	const last = flow[flow.length - 1].block;
	const end = (block: HTMLElement) => block.getBoundingClientRect().bottom + (block === last ? parseFloat(getComputedStyle(block).marginBottom) || 0 : 0);
	let over = flow.findIndex(f => !isBlank(f.block) && end(f.block) > limitY);
	let splitLimit = limitY;
	let keepNotes = false;

	// Notes that do not fit push the last body lines down, as in Word, instead of taking a page of their own.
	if (over > 0 && !flow[over].article && flow[over - 1].article) {
		const bodyLimit = limitY - (end(last) - flow[over - 1].block.getBoundingClientRect().bottom);
		const bodyOver = flow.findIndex(f => f.article && !isBlank(f.block) && f.block.getBoundingClientRect().bottom > bodyLimit);

		if (bodyOver > 0 && bodyLimit - articles[0].getBoundingClientRect().top >= MIN_SPLIT_HEIGHT) {
			over = bodyOver;
			splitLimit = bodyLimit;
			keepNotes = true;
		}
	}

	if (over < 0) {
		hideBlankOverflow(flow, limitY);
		return null;
	}

	const overArticle = flow[over].article;
	let cut = over;
	let carry: HTMLElement = null;

	if (flow[over].block.tagName === "TABLE")
		carry = splitTable(flow[over].block as HTMLTableElement, splitLimit);
	else if (flow[over].block.tagName === "P")
		carry = splitParagraph(flow[over].block, splitLimit, className);

	const stays = !carry && unbreakableHeight(flow[over].block) > pageSpace;

	if (carry || stays) {
		cut = over + 1;  // a block that cannot split and is taller than a page would overflow the next page too
	} else {
		while (cut > 0 && flow[cut - 1].article === overArticle && overArticle && getComputedStyle(flow[cut - 1].block).breakAfter === "avoid")
			cut--;

		if (cut === 0)
			cut = over || 1;
	}

	// the notes of a page whose tall block stays where it starts stay with it
	if (!carry && flow.slice(cut).every(f => f.article ? isBlank(f.block) : stays)) {
		hideBlankOverflow(flow, limitY);
		return null;
	}

	const moved = new Map<HTMLElement, HTMLElement[]>();
	const notes: HTMLElement[] = [];

	if (carry)
		moved.set(overArticle, [carry]);

	for (const f of flow.slice(cut)) {
		if (f.article)
			moved.set(f.article, [...(moved.get(f.article) ?? []), f.block]);
		else if (!keepNotes)
			notes.push(f.block);
	}

	const next = page.cloneNode(false) as HTMLElement;
	next.removeAttribute("data-docx-first-page");
	next.removeAttribute("data-page-start");
	next.removeAttribute("data-page-parity");

	for (const c of children) {
		if (c.tagName === "HEADER" && c.dataset.docxHf !== "first")
			next.appendChild(cloneWithTabStops(c));
	}

	for (const [article, blocks] of moved) {
		if (article !== overArticle || !carry && blocks[0] === article.firstElementChild) {
			next.appendChild(article);
		} else {
			const rest = article.cloneNode(false) as HTMLElement;
			rest.append(...blocks);
			next.appendChild(rest);
		}
	}

	next.append(...notes);

	for (const c of children) {
		if (c.tagName === "FOOTER" && c.dataset.docxHf !== "first")
			next.appendChild(cloneWithTabStops(c));
	}

	page.after(next);
	return next;
}

function showHeaderFooter(page: HTMLElement, number: number) {
	for (const tag of ["HEADER", "FOOTER"]) {
		const alternates = (Array.from(page.children) as HTMLElement[]).filter(c => c.tagName === tag && c.dataset.docxHf);
		const find = (type: string) => alternates.find(c => c.dataset.docxHf === type);
		const pick = (page.hasAttribute("data-docx-first-page") ? find("first") : null)
			?? (number % 2 == 0 ? find("even") : null)
			?? find("default");

		for (const c of alternates) {
			if ((c === pick) === (c.style.display === "none"))
				c.style.display = c === pick ? "" : "none";
		}
	}
}

function hideBlankOverflow(flow: { block: HTMLElement, article: HTMLElement }[], limitY: number) {
	for (const f of flow) {
		if (f.article && isBlank(f.block) && f.block.getBoundingClientRect().bottom > limitY)
			f.block.style.display = "none";
	}
}

/**
 * Height of the leading part of a block that cannot be split off: a table's header and first body row,
 * a whole paragraph kept together, or the whole of any other block. Any other paragraph counts as 0: one
 * with text splits between its lines, and one of pictures only moves to a new page instead of staying.
 */
function unbreakableHeight(block: HTMLElement): number {
	if (block instanceof HTMLTableElement) {
		const first = Array.from(block.rows).find(r => r.parentElement !== block.tHead);
		return first ? first.getBoundingClientRect().bottom - block.getBoundingClientRect().top : 0;
	}

	if (block.tagName === "P" && getComputedStyle(block).breakInside !== "avoid")
		return 0;

	return outerHeight(block);
}

function outerHeight(el: HTMLElement): number {
	const style = getComputedStyle(el);
	return el.getBoundingClientRect().height + (parseFloat(style.marginTop) || 0) + (parseFloat(style.marginBottom) || 0);
}

/**
 * Splits a table at the first body row that ends below limitY. The rows from
 * there move to a new table with the same attributes, columns and header rows;
 * cells merged across the cut continue in the first moved row.
 * Returns that table, or null when no body row fits above limitY.
 */
export function splitTable(table: HTMLTableElement, limitY: number): HTMLTableElement | null {
	const head = table.tHead;
	const rows = Array.from(table.rows).filter(r => r.parentElement !== head);
	const at = rows.findIndex(r => r.getBoundingClientRect().bottom > limitY);

	if (at <= 0)
		return null;

	const rest = table.cloneNode(false) as HTMLTableElement;

	for (const c of Array.from(table.children)) {
		if (c.tagName === "COLGROUP") {
			rest.appendChild(c.cloneNode(true));
		} else if (c === head) {
			const copy = cloneWithTabStops(head);
			copy.setAttribute("data-docx-repeat", "");
			clampRowSpans(Array.from(copy.rows), copy.rows.length);
			rest.appendChild(copy);
		}
	}

	const first = rows[at];

	for (const { row, cell, column, span } of clampRowSpans(rows.slice(0, at), at)) {
		const continued = cellAtColumn(first, column);

		if (continued?.style.display === "none") {
			continued.style.cssText = cell.style.cssText;
			continued.rowSpan = span - (at - row);
		}
	}

	for (const r of rows.slice(at))
		rest.appendChild(r);

	return rest;
}

/**
 * Splits a paragraph at a line boundary so the lines ending above limitY stay,
 * with at least two lines on each side (Word's widow and orphan control).
 * Returns the continuation (inserted after p), or null when it cannot split.
 */
export function splitParagraph(p: HTMLElement, limitY: number, className: string = "docx"): HTMLElement | null {
	if (getComputedStyle(p).breakInside === "avoid")
		return null;

	const lines = paragraphLines(p);
	let keep = lines.filter(l => l.bottom <= limitY).length;

	if (lines.length - keep < 2)
		keep = lines.length - 2;

	if (keep < 2)
		return null;

	const at = textPositionAt(p, lines[keep].top - 1);

	if (!at)
		return null;

	const range = document.createRange();
	range.setStart(at.node, at.offset);
	range.setEnd(p, p.childNodes.length);

	const rest = p.cloneNode(false) as HTMLElement;
	rest.removeAttribute("id");
	rest.appendChild(range.extractContents());

	// The first part ends the page: no space after it. The continuation is not
	// a new paragraph: no space before it, no first-line indent, no list marker.
	p.style.marginBottom = "0";
	rest.style.marginTop = "0";
	rest.style.textIndent = "0";
	rest.classList.add(`${className}-continued`);

	p.after(rest);
	return rest;
}

/**
 * Clamps the visible cells of rows whose rowSpan reaches past rows.length.
 * Returns each clamped cell, with its row index, grid column and span before the clamp.
 */
function clampRowSpans(rows: HTMLTableRowElement[], end: number): { row: number, cell: HTMLTableCellElement, column: number, span: number }[] {
	const clamped: { row: number, cell: HTMLTableCellElement, column: number, span: number }[] = [];

	rows.forEach((row, i) => {
		let column = 0;

		for (const cell of Array.from(row.cells)) {
			if (cell.style.display !== "none" && i + cell.rowSpan > end) {
				clamped.push({ row: i, cell, column, span: cell.rowSpan });
				cell.rowSpan = end - i;
			}

			column += cell.colSpan;
		}
	});

	return clamped;
}

function cellAtColumn(row: HTMLTableRowElement, column: number): HTMLTableCellElement | null {
	let at = 0;

	for (const cell of Array.from(row.cells)) {
		if (at === column)
			return cell;

		at += cell.colSpan;
	}

	return null;
}

/** Line boxes of a paragraph's text ({top, bottom}, in order); pictures are not lines. */
export function paragraphLines(p: HTMLElement): { top: number, bottom: number }[] {
	const range = document.createRange();
	const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
	const rects: DOMRect[] = [];

	for (let n = walker.nextNode(); n; n = walker.nextNode()) {
		range.selectNodeContents(n);
		rects.push(...Array.from(range.getClientRects()));
	}

	const lines: { top: number, bottom: number }[] = [];

	for (const r of rects) {
		if (!r.width && !r.height)
			continue;

		const line = lines.find(l => Math.abs(l.top - r.top) < 2 || (r.top < l.bottom - 2 && r.bottom > l.top + 2));

		if (line) {
			line.top = Math.min(line.top, r.top);
			line.bottom = Math.max(line.bottom, r.bottom);
		} else {
			lines.push({ top: r.top, bottom: r.bottom });
		}
	}

	return lines.sort((a, b) => a.top - b.top);
}

/** Text position of the first character on or below y, by binary search over the text nodes. */
function textPositionAt(p: HTMLElement, y: number): { node: Text, offset: number } | null {
	const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
	const range = document.createRange();

	const topOf = (node: Text, offset: number) => {
		range.setStart(node, offset);
		range.setEnd(node, offset + 1);
		const rects = range.getClientRects();
		return rects.length ? rects[0].top : -Infinity;
	};

	for (let n = walker.nextNode() as Text; n; n = walker.nextNode() as Text) {
		if (!n.length || topOf(n, n.length - 1) < y)
			continue;

		let lo = 0;
		let hi = n.length - 1;

		while (lo < hi) {
			const mid = (lo + hi) >> 1;

			if (topOf(n, mid) >= y)
				hi = mid;
			else
				lo = mid + 1;
		}

		return { node: n, offset: lo };
	}

	return null;
}

function isBlank(el: HTMLElement): boolean {
	return !el.textContent.trim() && !el.querySelector("img, svg, table, canvas, image, div");
}
