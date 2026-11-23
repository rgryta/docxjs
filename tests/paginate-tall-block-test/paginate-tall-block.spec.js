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

  const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
  const EMU = 12700;
  const picture = (w, h) => `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${w * EMU}" cy="${h * EMU}"/><wp:docPr id="1" name="p"/>`
    + `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`
    + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPic"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

  function build(body, footnotes = null, firstHeaderLines = 0) {
    const zip = new JSZip();
    zip.file("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="word/document.xml"/></Relationships>');
    const titlePage = firstHeaderLines ? '<w:headerReference w:type="first" r:id="rIdHF"/><w:titlePg/>' : "";
    zip.file("word/document.xml", `<w:document ${NS}><w:body>${body}<w:sectPr>${LANDSCAPE}${titlePage}</w:sectPr></w:body></w:document>`);
    zip.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rIdH" Type="${REL}header" Target="header1.xml"/><Relationship Id="rIdF" Type="${REL}footer" Target="footer1.xml"/>`
      + (footnotes ? `<Relationship Id="rIdN" Type="${REL}footnotes" Target="footnotes.xml"/>` : "")
      + (firstHeaderLines ? `<Relationship Id="rIdHF" Type="${REL}header" Target="header2.xml"/>` : "")
      + `<Relationship Id="rIdPic" Type="${REL}image" Target="media/p.png"/></Relationships>`);
    zip.file("word/media/p.png", PNG, { base64: true });
    if (firstHeaderLines)
      zip.file("word/header2.xml", `<w:hdr ${NS}>${Array.from({ length: firstHeaderLines }, (_, i) =>
        `<w:p><w:pPr><w:spacing w:line="280" w:lineRule="exact"/></w:pPr><w:r><w:t>Letterhead ${i}</w:t></w:r></w:p>`).join("")}</w:hdr>`);
    if (footnotes)
      zip.file("word/footnotes.xml", `<w:footnotes ${NS}><w:footnote w:id="1">${para(footnotes)}</w:footnote></w:footnotes>`);
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

    it("moves a block that only the next page can hold, under a tall first-page header", async () => {
      // exact heights keep the case independent of fonts: 15 x 14pt of letterhead leaves page 1 about 300pt, page 2 about 500pt
      const tallRow = `<w:tr><w:trPr><w:trHeight w:val="8000" w:hRule="exact"/></w:trPr>${[para("a"), para("b"), para("c")].map(cell).join("")}</w:tr>`;
      const body = para("Title") + table(tallRow);

      div = document.createElement("div");
      document.body.appendChild(div);
      await docx.renderAsync(await build(body, null, 15), div, null, { ignoreLastRenderedPageBreak: true, experimental: true, paginate: true });

      const pages = Array.from(div.querySelectorAll("section.docx"));
      const tall = div.querySelector("article > table");
      expect(pages.length).toBe(2);
      expect(tall.closest("section.docx")).toBe(pages[1]);
      expect(pages[1].getBoundingClientRect().height).toBeLessThanOrEqual(parseFloat(getComputedStyle(pages[1]).minHeight) + 1);
    });

    it("moves a picture-only paragraph taller than a page to a new page", async () => {
      div = document.createElement("div");
      document.body.appendChild(div);
      await docx.renderAsync(await build(para("Title") + picture(300, 700)), div, null, { ignoreLastRenderedPageBreak: true, experimental: true, paginate: true });

      const pages = Array.from(div.querySelectorAll("section.docx"));
      expect(pages.length).toBe(2);
      expect(div.querySelector("article img").closest("section.docx")).toBe(pages[1]);
    });

    it("keeps the notes on the page of a tall block that stays where it starts", async () => {
      const body = '<w:p><w:r><w:t>Title</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
        + table(row(paras(10, "left"), paras(25, "middle"), paras(60, "right")));

      div = document.createElement("div");
      document.body.appendChild(div);
      await docx.renderAsync(await build(body, "Note"), div, null, { ignoreLastRenderedPageBreak: true, experimental: true });

      expect(docx.paginate(div, "docx", { budgetMs: 1000 })).toBe(true);

      const pages = Array.from(div.querySelectorAll("section.docx"));
      expect(pages.length).toBe(1);
      expect(pages[0].querySelector(":scope > ol")).not.toBeNull();
    });

    it("keeps a paragraph with keep-lines taller than a page where it starts, even after earlier text", async () => {
      const lines = '<w:r><w:t>tall</w:t></w:r>' + '<w:r><w:br/><w:t>line</w:t></w:r>'.repeat(70);
      const body = para("Title") + para("Subtitle") + `<w:p><w:pPr><w:keepLines/></w:pPr>${lines}</w:p>`
        + pageBreak + para("Second page");

      div = document.createElement("div");
      document.body.appendChild(div);
      await docx.renderAsync(await build(body), div, null, { ignoreLastRenderedPageBreak: true, experimental: true });

      const tall = Array.from(div.querySelectorAll("article > p")).find(p => p.textContent.startsWith("tall"));
      const page = tall.closest("section.docx");
      expect(tall.getBoundingClientRect().height).toBeGreaterThan(parseFloat(getComputedStyle(page).minHeight));

      expect(docx.paginate(div, "docx", { budgetMs: 1000 })).toBe(true);

      const pages = Array.from(div.querySelectorAll("section.docx"));
      expect(tall.closest("section.docx")).toBe(pages[0]);
      expect(pages.length).toBe(2);
    });
  });
})();
