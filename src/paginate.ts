// Splits rendered pages whose content is taller than the page, the way Word
// lays out text: whole blocks move to a new page, paragraphs split at a line
// boundary (at least two lines per page), tables split between rows with their
// header rows (<thead>) repeated, and a block marked keep-with-next
// (break-after: avoid) moves along with the block after it.
//
// It measures the layout, so the pages must be in the document and the fonts
// loaded. renderAsync runs it with the `paginate` option; call it yourself
// after renderDocument when you insert the nodes elsewhere.

const MAX_PAGES = 2000;
const MIN_SPLIT_HEIGHT = 24;

/**
 * Splits every rendered page (section.<className>) under root until each fits its page height,
 * then shows the header and footer each page needs. With budgetMs it stops before the next page
 * once the budget is spent, keeping the pages split so far. Returns false when it stopped early.
 */
export function paginate(root: HTMLElement, className: string = "docx", options?: { budgetMs?: number }): boolean {
	const deadline = options?.budgetMs === undefined ? Infinity : performance.now() + options.budgetMs;
	const pages = Array.from(root.querySelectorAll<HTMLElement>(`section.${className}`));
	let done = true;

	for (let i = 0; i < pages.length; i++) {
		if (pages.length >= MAX_PAGES || performance.now() >= deadline) {
			done = false;
			break;
		}

		// Pages before i are final, so i is this page's final position.
		showHeaderFooter(pages[i], i);
		const next = splitPage(pages[i], className);

		if (next)
			pages.splice(i + 1, 0, next);
	}

	pages.forEach(showHeaderFooter);
	return done;
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

	if (limitY - articles[0].getBoundingClientRect().top < MIN_SPLIT_HEIGHT)
		return null;

	// A footnote list stays on the page with its references: the page keeps as many notes as it keeps
	// references, and the body gives up room for them. The other notes move to a list on the next page.
	const footnotes = children.find(c => c.tagName === "OL" && c.dataset.docxNotes === "footnotes");
	const items = footnotes ? Array.from(footnotes.children) as HTMLElement[] : [];
	const refs = footnotes ? articles.flatMap(a => Array.from(a.querySelectorAll<HTMLElement>("sup[data-docx-footnote]"))) : [];
	let keptNotes = items.length;
	let notesLimit = limitY;

	if (items.length && refs.length) {
		const chrome = outerHeight(footnotes) - (items[items.length - 1].getBoundingClientRect().bottom - items[0].getBoundingClientRect().top);
		keptNotes = 0;

		for (let k = Math.min(items.length, refs.length); k > 0; k--) {
			const limit = limitY - chrome - (items[k - 1].getBoundingClientRect().bottom - items[0].getBoundingClientRect().top);

			if (limit - articles[0].getBoundingClientRect().top >= MIN_SPLIT_HEIGHT && refs.filter(r => r.getBoundingClientRect().bottom <= limit).length >= k) {
				keptNotes = k;
				notesLimit = limit;
				break;
			}
		}
	} else if (items.length && limitY - outerHeight(footnotes) - articles[0].getBoundingClientRect().top >= MIN_SPLIT_HEIGHT) {
		notesLimit = limitY - outerHeight(footnotes);  // notes carried over from the page before keep their room
	}

	const flow = [
		...articles.flatMap(article => (Array.from(article.children) as HTMLElement[]).map(block => ({ block, article }))),
		...children.filter(c => c.tagName === "OL" && c !== footnotes).map(block => ({ block, article: null as HTMLElement })),
	];
	// Empty paragraphs (such as the one holding a page break) never start a page.
	// the page also grows by the last block's bottom margin: margins do not collapse in the page's flex column
	const last = flow[flow.length - 1].block;
	const end = (block: HTMLElement) => block.getBoundingClientRect().bottom + (block === last ? parseFloat(getComputedStyle(block).marginBottom) || 0 : 0);
	const bodyEnd = notesLimit;
	let over = flow.findIndex(f => !isBlank(f.block) && end(f.block) > bodyEnd);
	let splitLimit = bodyEnd;
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

	if (over < 0 && keptNotes == items.length) {
		for (const f of flow) {
			if (f.article && isBlank(f.block) && f.block.getBoundingClientRect().bottom > limitY)
				f.block.style.display = "none";
		}

		return null;
	}

	const overArticle = over < 0 ? null : flow[over].article;
	let cut = over < 0 ? flow.length : over;
	let carry: HTMLElement = null;

	if (over >= 0 && flow[over].block.tagName === "TABLE")
		carry = splitTable(flow[over].block as HTMLTableElement, splitLimit);
	else if (over >= 0 && flow[over].block.tagName === "P")
		carry = splitParagraph(flow[over].block, splitLimit, className);

	if (carry) {
		cut = over + 1;
	} else if (over >= 0) {
		while (cut > 0 && flow[cut - 1].article === overArticle && overArticle && getComputedStyle(flow[cut - 1].block).breakAfter === "avoid")
			cut--;

		if (cut === 0)
			cut = over || 1;

		if (cut >= flow.length && keptNotes == items.length)
			return null;  // a single block taller than a page stays where it is
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

	for (const c of children) {
		if (c.tagName === "HEADER" && c.dataset.docxHf !== "first")
			next.appendChild(c.cloneNode(true));
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

	if (keptNotes < items.length) {
		const rest = keptNotes ? footnotes.cloneNode(false) as HTMLElement : footnotes;

		if (keptNotes) {
			rest.setAttribute("start", String(Number(footnotes.getAttribute("start") ?? 1) + keptNotes));
			rest.append(...items.slice(keptNotes));
		}

		next.appendChild(rest);
	}

	next.append(...notes);

	for (const c of children) {
		if (c.tagName === "FOOTER" && c.dataset.docxHf !== "first")
			next.appendChild(c.cloneNode(true));
	}

	page.after(next);
	return next;
}

function showHeaderFooter(page: HTMLElement, index: number) {
	for (const tag of ["HEADER", "FOOTER"]) {
		const alternates = (Array.from(page.children) as HTMLElement[]).filter(c => c.tagName === tag && c.dataset.docxHf);
		const find = (type: string) => alternates.find(c => c.dataset.docxHf === type);
		const pick = (page.hasAttribute("data-docx-first-page") ? find("first") : null)
			?? (index % 2 == 1 ? find("even") : null)
			?? find("default");

		for (const c of alternates) {
			if ((c === pick) === (c.style.display === "none"))
				c.style.display = c === pick ? "" : "none";
		}
	}
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
			const copy = head.cloneNode(true) as HTMLTableSectionElement;
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
