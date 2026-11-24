(() => {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const PAGE = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="500" w:footer="500"/>';

  const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const paras = (count, label) => Array.from({ length: count }, (_, i) => para(`${label} ${i}`)).join("");
  const cell = (content) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${content}</w:tc>`;

  async function render(body, headers = {}, options = {}) {
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}officeDocument" Target="word/document.xml"/></Relationships>`);
    const refs = Object.keys(headers).map(type => `<w:headerReference w:type="${type}" r:id="rId${type}"/>`).join("") + (headers.first ? "<w:titlePg/>" : "");
    zip.file("word/document.xml", `<w:document ${NS}><w:body>${body}<w:sectPr>${refs}${PAGE}</w:sectPr></w:body></w:document>`);
    zip.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + Object.keys(headers).map(type => `<Relationship Id="rId${type}" Type="${REL}header" Target="${type}.xml"/>`).join("") + "</Relationships>");
    Object.entries(headers).forEach(([type, content]) => zip.file(`word/${type}.xml`, `<w:hdr ${NS}>${content}</w:hdr>`));

    const div = document.createElement("div");
    document.body.appendChild(div);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), div, null, { ignoreLastRenderedPageBreak: true, ...options });
    return div;
  }

  describe("Paginate a block taller than a page", function () {
    let div;

    afterEach(() => div?.remove());

    it("keeps a one-row table taller than a page on the page it starts, without a blank page after it", async () => {
      const row = [paras(10, "left"), paras(25, "middle"), paras(60, "right")].map(cell).join("");
      div = await render(para("Title")
        + `<w:tbl><w:tblPr><w:tblW w:w="13500" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr>${row}</w:tr></w:tbl>`
        + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + para("Second page"));
      const tall = div.querySelector("article > table");
      expect(tall.getBoundingClientRect().height).toBeGreaterThan(parseFloat(getComputedStyle(tall.closest("section.docx")).minHeight));

      expect(docx.paginate(div)).toBe(true);

      const pages = Array.from(div.querySelectorAll("section.docx"));
      expect(pages.length).toBe(2);
      expect(tall.closest("section.docx")).toBe(pages[0]);
      expect(pages[1].textContent).toContain("Second page");
    });

    it("moves a row that fits only under the next page's shorter header", async () => {
      // exact heights: 15 x 14pt of first-page letterhead leaves page 1 about 320pt, the 14pt default header leaves page 2 about 500pt
      const lines = (count) => Array.from({ length: count }, (_, i) => `<w:p><w:pPr><w:spacing w:line="280" w:lineRule="exact"/></w:pPr><w:r><w:t>${i}</w:t></w:r></w:p>`).join("");
      const row = [para("a"), para("b"), para("c")].map(cell).join("");
      div = await render(para("Title")
        + `<w:tbl><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="8000" w:hRule="exact"/></w:trPr>${row}</w:tr></w:tbl>`,
        { first: lines(15), default: lines(1) }, { paginate: true });

      const pages = Array.from(div.querySelectorAll("section.docx"));
      expect(pages.length).toBe(2);
      expect(div.querySelector("article > table").closest("section.docx")).toBe(pages[1]);
    });
  });
})();
