(() => {
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';
const LANDSCAPE = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';
const FOOTER = '<w:footerReference w:type="default" r:id="rIdF"/>';

const r = (inner, rPr = "") => `<w:r><w:rPr>${rPr}</w:rPr>${inner}</w:r>`;
const run = (text, rPr = "") => r(`<w:t xml:space="preserve">${text}</w:t>`, rPr);
const para = (text) => `<w:p>${run(text)}</w:p>`;
const fld = (type) => `<w:fldChar w:fldCharType="${type}"/>`;
const instr = (text) => `<w:instrText xml:space="preserve"> ${text} </w:instrText>`;
const field = (instruction, saved, rPr = "") => r(fld("begin"), rPr) + r(instr(instruction), rPr) + r(fld("separate"), rPr) + run(saved, rPr) + r(fld("end"), rPr);
const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
const sectionEnd = (sectPr) => `<w:p><w:pPr><w:sectPr>${sectPr}</w:sectPr></w:pPr></w:p>`;
const part = (id, paragraphs, type = "footer") => ({
  id, type, target: `${id}.xml`,
  content: `<w:${type === "header" ? "hdr" : "ftr"} ${NS}>${paragraphs.join("")}</w:${type === "header" ? "hdr" : "ftr"}>`,
});
const pageFooter = (saved = "9") => part("rIdF", [`<w:p>${field("PAGE", saved)}</w:p>`]);

function build({ body, sectPr = PAGE, parts = [] }) {
  const zip = new JSZip();
  zip.file("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file("word/document.xml", `<w:document ${NS}><w:body>${body}<w:sectPr>${sectPr}</w:sectPr></w:body></w:document>`);
  zip.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + parts.map(p => `<Relationship Id="${p.id}" Type="${REL}${p.type}" Target="${p.target}"/>`).join("") + "</Relationships>");

  for (const p of parts)
    zip.file(`word/${p.target}`, p.content);

  return zip.generateAsync({ type: "blob" });
}

async function render(doc, options = {}) {
  const div = document.createElement("div");
  document.body.appendChild(div);
  await docx.renderAsync(await build(doc), div, null, { ignoreLastRenderedPageBreak: true, ...options });
  return div;
}

const pages = (div) => Array.from(div.querySelectorAll("section.docx"));
const shown = (page, tag) => Array.from(page.children).filter(c => c.tagName === tag && getComputedStyle(c).display !== "none");
const shownText = (div, tag) => pages(div).map(p => shown(p, tag).map(c => c.textContent).join("|"));
const bodyText = (page) => Array.from(page.querySelectorAll("article")).map(a => a.textContent).join("");

describe("Page fields and section breaks", function () {
  let div;

  afterEach(() => div?.remove());

  it("keeps a continuous section on the page", async () => {
    div = await render({
      body: para("A") + sectionEnd(PAGE) + para("B"),
      sectPr: `<w:type w:val="continuous"/>${PAGE}`,
    });

    expect(pages(div).map(bodyText)).toEqual(["AB"]);
  });

  it("starts a page at a next-page section, at the section's first piece", async () => {
    div = await render({
      body: para("A") + sectionEnd(PAGE) + para("B") + pageBreak + para("C") + sectionEnd(PAGE) + para("D"),
      sectPr: `<w:type w:val="continuous"/>${PAGE}`,
    });

    expect(pages(div).map(bodyText)).toEqual(["A", "B", "CD"]);
  });

  it("starts the page of a new page size at the section's first piece", async () => {
    div = await render({
      body: para("A") + sectionEnd(PAGE) + para("B"),
      sectPr: `<w:type w:val="continuous"/>${LANDSCAPE}`,
    }, { ignoreLastRenderedPageBreak: false });

    const [first, second] = pages(div);
    expect(pages(div).map(bodyText)).toEqual(["A", "B"]);
    expect(parseFloat(second.style.width)).toBeGreaterThan(parseFloat(first.style.width));
  });

  it("adds no page for a leading lastRenderedPageBreak after a next-page section", async () => {
    div = await render({
      body: para("A") + sectionEnd(FOOTER + PAGE) + `<w:p><w:r><w:lastRenderedPageBreak/><w:t>B</w:t></w:r></w:p>`,
      sectPr: FOOTER + PAGE,
      parts: [pageFooter()],
    });

    expect(docx.paginate(div)).toBe(true);
    expect(shownText(div, "FOOTER")).toEqual(["1", "2"]);
  });

  it("adds no page for a page-break-before paragraph after a section end", async () => {
    div = await render({
      body: para("A") + sectionEnd(FOOTER + PAGE) + `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${run("B")}</w:p>`,
      sectPr: FOOTER + PAGE,
      parts: [pageFooter()],
    });

    expect(docx.paginate(div)).toBe(true);
    expect(shownText(div, "FOOTER")).toEqual(["1", "2"]);
  });

  it("fills PAGE and NUMPAGES in the number run's formatting", async () => {
    const rPr = '<w:b/><w:sz w:val="18"/>';
    div = await render({
      body: para("A") + pageBreak + para("B") + pageBreak + para("C"),
      sectPr: FOOTER + PAGE,
      parts: [part("rIdF", [`<w:p>${run("Page ")}${field("PAGE", "1", rPr)}${run(" of ")}${field("NUMPAGES", "1", rPr)}</w:p>`])],
    });

    expect(docx.paginate(div)).toBe(true);
    expect(shownText(div, "FOOTER")).toEqual(["Page 1 of 3", "Page 2 of 3", "Page 3 of 3"]);

    const number = shown(pages(div)[1], "FOOTER")[0].querySelector('[data-docx-field="PAGE"] span');
    expect(number.textContent).toBe("2");
    expect(getComputedStyle(number).fontWeight).toBe("700");
    expect(parseFloat(getComputedStyle(number).fontSize)).toBeCloseTo(12, 1);
  });

  it("draws a number for a field with its characters in one run and no result run", async () => {
    div = await render({
      body: para("A") + pageBreak + para("B"),
      sectPr: FOOTER + PAGE,
      parts: [part("rIdF", [`<w:p>${r(fld("begin") + instr("PAGE") + fld("separate") + fld("end"))}</w:p>`])],
    });

    expect(docx.paginate(div)).toBe(true);
    expect(shownText(div, "FOOTER")).toEqual(["1", "2"]);
  });

  it("keeps the saved text of roman, nested and roman-section fields", async () => {
    const ifField = r(fld("begin")) + r(instr("IF")) + field("PAGE", "7") + r(instr('= 1 "one" "other"')) + r(fld("separate")) + run("other") + r(fld("end"));
    div = await render({
      body: para("A") + sectionEnd(FOOTER + PAGE) + para("B"),
      sectPr: `${FOOTER}<w:pgNumType w:fmt="lowerRoman"/>${PAGE}`,
      parts: [part("rIdF", [`<w:p>${field("PAGE", "9")}</w:p>`, `<w:p>${field("PAGE \\* roman", "ix")}</w:p>`, `<w:p>${ifField}</w:p>`])],
    });

    expect(docx.paginate(div)).toBe(true);
    const footers = pages(div).map(p => Array.from(shown(p, "FOOTER")[0].querySelectorAll("p")).map(e => e.textContent));
    expect(footers).toEqual([["1", "ix", "7other"], ["9", "ix", "7other"]]);
    expect(div.querySelectorAll("footer p:nth-child(n+2) [data-docx-field]").length).toBe(0);
  });

  it("numbers pages from pgNumType starts", async () => {
    const section = (pgNumType) => FOOTER + pgNumType + PAGE;
    div = await render({
      body: para("A") + pageBreak + para("B") + pageBreak + para("C") + sectionEnd(section("<w:pgNumType/>"))
        + para("D") + sectionEnd(section('<w:pgNumType w:start="5"/>'))
        + para("E") + sectionEnd(section('<w:pgNumType w:start="0"/>'))
        + para("F"),
      sectPr: section('<w:pgNumType w:start="x"/>'),
      parts: [pageFooter()],
    });

    expect(docx.paginate(div)).toBe(true);
    expect(shownText(div, "FOOTER")).toEqual(["1", "2", "3", "5", "0", "1"]);
  });

  it("fills a paragraph-level simple PAGE field in a header", async () => {
    div = await render({
      body: para("A") + pageBreak + para("B"),
      sectPr: `<w:headerReference w:type="default" r:id="rIdH"/>${PAGE}`,
      parts: [part("rIdH", [`<w:p><w:fldSimple w:instr=" PAGE ">${run("9")}</w:fldSimple></w:p>`], "header")],
    });

    expect(docx.paginate(div)).toBe(true);
    expect(shownText(div, "HEADER")).toEqual(["1", "2"]);
  });

  it("keeps the saved NUMPAGES text when it stops at the time budget", async () => {
    div = await render({
      body: para("A") + pageBreak + para("B") + pageBreak + para("C"),
      sectPr: FOOTER + PAGE,
      parts: [part("rIdF", [`<w:p>${field("PAGE", "7")}${run(" of ")}${field("NUMPAGES", "12")}</w:p>`])],
    });
    let now = 0;
    spyOn(performance, "now").and.callFake(() => ++now);

    // The first split pass finishes inside the budget; the pass after the fill stops.
    expect(docx.paginate(div, "docx", { budgetMs: 4 })).toBe(false);
    expect(shownText(div, "FOOTER")).toEqual(["1 of 12", "2 of 12", "3 of 12"]);
  });

  it("picks even and odd footers by page number", async () => {
    div = await render({
      body: para("A") + pageBreak + para("B"),
      sectPr: `${FOOTER}<w:footerReference w:type="even" r:id="rIdE"/><w:pgNumType w:start="2"/>${PAGE}`,
      parts: [
        part("rIdF", [para("odd")]), part("rIdE", [para("even")]),
        { id: "rIdS", type: "settings", target: "settings.xml", content: `<w:settings ${NS}><w:evenAndOddHeaders/></w:settings>` },
      ],
    });

    expect(docx.paginate(div)).toBe(true);
    expect(shownText(div, "FOOTER")).toEqual(["even", "odd"]);
  });

  it("shows the first-page footer on each title-page section's first page", async () => {
    const section = `<w:footerReference w:type="first" r:id="rIdF1"/>${FOOTER}<w:titlePg/>${PAGE}`;
    div = await render({
      body: para("A") + pageBreak + para("B") + sectionEnd(section) + para("C"),
      sectPr: section,
      parts: [part("rIdF1", [para("first")]), pageFooter()],
    });

    expect(docx.paginate(div)).toBe(true);
    expect(shownText(div, "FOOTER")).toEqual(["first", "2", "first"]);
  });
});
})();
