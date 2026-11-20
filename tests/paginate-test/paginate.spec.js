describe("Paginate", function () {
  let div;

  beforeAll(async () => {
    const docBlob = await fetch(`/base/tests/paginate-test/document.docx`).then(r => r.blob());
    div = document.createElement("div");
    document.body.appendChild(div);
    await docx.renderAsync(docBlob, div, null, { className: "paginate-fixture" });
    const longest = [...div.querySelectorAll("article p")].sort((a, b) => b.textContent.length - a.textContent.length)[0];
    longest.append(...[...longest.childNodes].map(n => n.cloneNode(true)));
    docx.paginate(div, "paginate-fixture");
  });

  afterAll(() => div.remove());

  const pages = () => Array.from(div.querySelectorAll("section.paginate-fixture"));
  const blocks = (page) => Array.from(page.querySelector("article").children).filter(b => b.style.display !== "none");

  it("splits the document into more pages than its explicit breaks", () => {
    expect(pages().length).toBeGreaterThan(1);
  });

  it("keeps every page within its page height", () => {
    for (const page of pages()) {
      const height = parseFloat(getComputedStyle(page).minHeight);
      expect(page.getBoundingClientRect().height).toBeLessThanOrEqual(height + 1);
    }
  });

  it("never ends a page with a keep-with-next heading", () => {
    for (const page of pages().slice(0, -1)) {
      const last = blocks(page).pop();
      expect(getComputedStyle(last).breakAfter).not.toBe("avoid");
    }
  });

  it("repeats the header row of a table split across pages", () => {
    const tables = Array.from(div.querySelectorAll("table"));
    expect(tables.length).toBeGreaterThan(1);
    for (const table of tables) {
      expect(table.tHead.rows[0].textContent).toContain("Header A");
    }
  });

  it("splits a paragraph with at least two lines on each page", () => {
    const continued = Array.from(div.querySelectorAll("p.paginate-fixture-continued"));
    expect(continued.length).toBeGreaterThan(0);
    for (const rest of continued) {
      const first = rest.closest("section").previousElementSibling.querySelector("article").lastElementChild;
      expect(lineCount(first)).toBeGreaterThanOrEqual(2);
      expect(lineCount(rest)).toBeGreaterThanOrEqual(2);
      expect(getComputedStyle(rest, "::before").content).toBe("none");
    }
  });
});

function lineCount(p) {
  const range = document.createRange();
  range.selectNodeContents(p);
  return new Set(Array.from(range.getClientRects()).filter(r => r.width).map(r => Math.round(r.top))).size;
}
