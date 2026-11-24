(() => {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';
  const FOOTER = '<w:footerReference w:type="default" r:id="rIdF"/>';

  const r = (inner, rPr = "") => `<w:r><w:rPr>${rPr}</w:rPr>${inner}</w:r>`;
  const run = (text, rPr = "") => r(`<w:t xml:space="preserve">${text}</w:t>`, rPr);
  const para = (text) => `<w:p>${run(text)}</w:p>`;
  const fld = (type) => `<w:fldChar w:fldCharType="${type}"/>`;
  const field = (instruction, saved, rPr = "") => r(fld("begin"), rPr) + r(`<w:instrText xml:space="preserve"> ${instruction} </w:instrText>`, rPr)
    + r(fld("separate"), rPr) + run(saved, rPr) + r(fld("end"), rPr);
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const sectionEnd = (sectPr) => `<w:p><w:pPr><w:sectPr>${sectPr}</w:sectPr></w:pPr></w:p>`;
  const footer = (id, content) => ({ id, type: "footer", target: `${id}.xml`, content: `<w:ftr ${NS}>${content}</w:ftr>` });
  const pageFooter = footer("rIdF", `<w:p>${field("PAGE", "9")}</w:p>`);

  async function render({ body, sectPr = PAGE, parts = [] }, options = {}) {
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/document.xml", `<w:document ${NS}><w:body>${body}<w:sectPr>${sectPr}</w:sectPr></w:body></w:document>`);
    zip.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + parts.map(p => `<Relationship Id="${p.id}" Type="${REL}${p.type}" Target="${p.target}"/>`).join("") + "</Relationships>");
    parts.forEach(p => zip.file(`word/${p.target}`, p.content));

    const div = document.createElement("div");
    document.body.appendChild(div);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), div, null, { ignoreLastRenderedPageBreak: true, ...options });
    return div;
  }

  const pages = (div) => Array.from(div.querySelectorAll("section.docx"));
  const shown = (page, tag) => Array.from(page.children).filter(c => c.tagName === tag && getComputedStyle(c).display !== "none");
  const footers = (div) => pages(div).map(p => shown(p, "FOOTER").map(c => c.textContent).join("|"));
  const bodyText = (page) => Array.from(page.querySelectorAll("article"), a => a.textContent).join("");

  describe("Page fields and section breaks", function () {
    let div;

    afterEach(() => div?.remove());

    it("fills PAGE and NUMPAGES in the number run's formatting", async () => {
      const rPr = '<w:b/>';
      div = await render({
        body: para("A") + pageBreak + para("B") + pageBreak + para("C"),
        sectPr: FOOTER + PAGE,
        parts: [footer("rIdF", `<w:p>${run("Page ")}${field("PAGE", "1", rPr)}${run(" of ")}${field("NUMPAGES", "1", rPr)}</w:p>`)],
      });

      expect(docx.paginate(div)).toBe(true);
      expect(footers(div)).toEqual(["Page 1 of 3", "Page 2 of 3", "Page 3 of 3"]);
      expect(getComputedStyle(shown(pages(div)[1], "FOOTER")[0].querySelector('[data-docx-field="PAGE"] span')).fontWeight).toBe("700");
    });

    it("starts a page at a next-page section, at the section's first piece", async () => {
      div = await render({
        body: para("A") + sectionEnd(PAGE) + para("B") + pageBreak + para("C") + sectionEnd(PAGE) + para("D"),
        sectPr: `<w:type w:val="continuous"/>${PAGE}`,
      });

      expect(pages(div).map(bodyText)).toEqual(["A", "B", "CD"]);
    });

    it("numbers pages from pgNumType starts", async () => {
      div = await render({
        body: para("A") + pageBreak + para("B") + sectionEnd(FOOTER + PAGE) + para("C"),
        sectPr: `${FOOTER}<w:pgNumType w:start="5"/>${PAGE}`,
        parts: [pageFooter],
      });

      expect(docx.paginate(div)).toBe(true);
      expect(footers(div)).toEqual(["1", "2", "5"]);
    });

    it("picks even and odd footers by page number", async () => {
      div = await render({
        body: para("A") + pageBreak + para("B"),
        sectPr: `${FOOTER}<w:footerReference w:type="even" r:id="rIdE"/><w:pgNumType w:start="2"/>${PAGE}`,
        parts: [
          footer("rIdF", para("odd")), footer("rIdE", para("even")),
          { id: "rIdS", type: "settings", target: "settings.xml", content: `<w:settings ${NS}><w:evenAndOddHeaders/></w:settings>` },
        ],
      }, { paginate: true });

      expect(footers(div)).toEqual(["even", "odd"]);
    });

    it("starts numbering at the first section's pgNumType start when that section is continuous", async () => {
      div = await render({
        body: para("A") + pageBreak + para("B"),
        sectPr: `${FOOTER}<w:type w:val="continuous"/><w:pgNumType w:start="5"/>${PAGE}`,
        parts: [pageFooter],
      });

      expect(docx.paginate(div)).toBe(true);
      expect(footers(div)).toEqual(["5", "6"]);
    });

    it("numbers a roman odd-page section after the blank page Word adds before it", async () => {
      div = await render({
        body: para("A") + sectionEnd(FOOTER + PAGE) + para("B"),
        sectPr: `${FOOTER}<w:type w:val="oddPage"/><w:pgNumType w:fmt="lowerRoman"/>${PAGE}`,
        parts: [pageFooter],
      });

      expect(docx.paginate(div)).toBe(true);
      expect(footers(div)).toEqual(["1", "iii"]);
    });

    it("keeps the page after a page break that ends the body", async () => {
      div = await render({
        body: para("A") + pageBreak,
        sectPr: FOOTER + PAGE,
        parts: [footer("rIdF", `<w:p>${field("PAGE", "9")}${run(" of ")}${field("NUMPAGES", "9")}</w:p>`)],
      }, { paginate: true });

      expect(footers(div)).toEqual(["1 of 2", "2 of 2"]);
    });

    it("starts the page after a mid-body page break with the next paragraph", async () => {
      div = await render({ body: para("A") + pageBreak + para("B") });

      expect(pages(div).map(p => p.querySelector("article").firstElementChild.textContent)).toEqual(["A", "B"]);
    });

    it("keeps a section's leading table on that section's page", async () => {
      const table = '<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
      div = await render({
        body: para("A") + sectionEnd(PAGE) + table + `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${run("B")}</w:p>`,
        sectPr: '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>',
      });

      expect(div.querySelector("article table").closest("section").style.width).toBe("792pt");
    });
  });
})();
