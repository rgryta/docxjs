(() => {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const LANDSCAPE = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="500" w:footer="500"/>'
    + '<w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>';

  const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const paras = (count, label) => Array.from({ length: count }, (_, i) => para(`${label} ${i}`)).join("");
  const cell = (content) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
  const row = (...cells) => `<w:tr>${cells.map(cell).join("")}</w:tr>`;
  const table = (...rows) => `<w:tbl><w:tblPr><w:tblW w:w="13500" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>${rows.join("")}</w:tbl>`;
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const part = (tag, text) => `<w:${tag} ${NS}>${para(text)}</w:${tag}>`;

  function build(body) {
    const zip = new JSZip();
    zip.file("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file("word/document.xml", `<w:document ${NS}><w:body>${body}<w:sectPr>${LANDSCAPE}</w:sectPr></w:body></w:document>`);
    zip.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rIdH" Type="${REL}header" Target="header1.xml"/><Relationship Id="rIdF" Type="${REL}footer" Target="footer1.xml"/></Relationships>`);
    zip.file("word/header1.xml", part("hdr", "Header"));
    zip.file("word/footer1.xml", part("ftr", "Footer"));
    return zip.generateAsync({ type: "blob" });
  }

  describe("Paginate a block taller than a page", function () {
    let div;

    afterEach(() => div?.remove());

    it("keeps a one-row table taller than a page on the page it starts, without a blank page after it", async () => {
      const body = para("Title") + para("Subtitle") + table(row(para("a"), para("b"), para("c")))
        + table(row(paras(10, "left"), paras(25, "middle"), paras(60, "right")))
        + pageBreak + para("Second page") + table(row(para("x"), para("y"), para("z")));

      div = document.createElement("div");
      document.body.appendChild(div);
      await docx.renderAsync(await build(body), div, null, { ignoreLastRenderedPageBreak: true, experimental: true });

      const tall = div.querySelectorAll("article > table")[1];
      const page = tall.closest("section.docx");
      expect(tall.getBoundingClientRect().height).toBeGreaterThan(parseFloat(getComputedStyle(page).minHeight));

      expect(docx.paginate(div, "docx", { budgetMs: 1000 })).toBe(true);

      const pages = Array.from(div.querySelectorAll("section.docx"));
      expect(tall.closest("section.docx")).toBe(pages[0]);
      expect(pages.length).toBe(2);
      expect(pages[1].textContent).toContain("Second page");
    });
  });
})();
