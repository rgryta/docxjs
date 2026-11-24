(() => {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';
  const SENTENCE = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.";
  const LOREM = `${SENTENCE} `.repeat(4);

  const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const paras = (count) => Array.from({ length: count }, (_, i) => para(`${i} ${LOREM}`)).join("");
  const part = (id, type, root, content) => ({ id, type, target: `${type}${id}.xml`, content: `<w:${root} ${NS}>${content}</w:${root}>` });

  async function render({ body, sectPr = PAGE, parts = [] }, options = {}) {
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/document.xml", `<w:document ${NS}><w:body>${body}<w:sectPr>${sectPr}</w:sectPr></w:body></w:document>`);
    zip.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + parts.map(p => `<Relationship Id="${p.id}" Type="${REL}${p.type}" Target="${p.target}"/>`).join("") + "</Relationships>");
    parts.forEach(p => zip.file(`word/${p.target}`, p.content));

    const div = document.createElement("div");
    document.body.appendChild(div);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), div, null, options);
    return div;
  }

  const pages = (div) => Array.from(div.querySelectorAll("section.docx"));
  const fits = (page) => page.getBoundingClientRect().height <= parseFloat(getComputedStyle(page).minHeight) + 1;
  const headers = (div) => pages(div).map(p => Array.from(p.children).filter(c => c.tagName === "HEADER" && getComputedStyle(c).display !== "none").map(h => h.textContent).join("|"));
  const bodyText = (div) => Array.from(div.querySelectorAll("section.docx > article"), a => a.textContent).join("");

  describe("Paginate flow", function () {
    let div;

    afterEach(() => div?.remove());

    it("fills pages across continuous sections without duplicating them", async () => {
      const sect = `<w:type w:val="continuous"/><w:headerReference w:type="default" r:id="rIdH"/>${PAGE}`;
      const sectionEnd = `<w:p><w:pPr><w:sectPr>${sect}</w:sectPr></w:pPr></w:p>`;
      div = await render({
        body: paras(9) + sectionEnd + paras(9) + sectionEnd + paras(9),
        sectPr: sect,
        parts: [part("rIdH", "header", "hdr", para("Header"))],
      }, { ignoreLastRenderedPageBreak: false });
      const before = bodyText(div);

      expect(docx.paginate(div)).toBe(true);

      expect(pages(div).length).toBeGreaterThanOrEqual(3);
      expect(bodyText(div)).toBe(before);
      pages(div).forEach(page => expect(fits(page)).toBe(true));
      expect(headers(div).every(h => h === "Header")).toBe(true);
    });

    it("shows the first-page header only on the first page", async () => {
      div = await render({
        body: paras(24),
        sectPr: `<w:headerReference w:type="first" r:id="rIdF"/><w:headerReference w:type="default" r:id="rIdD"/><w:titlePg/>${PAGE}`,
        parts: [part("rIdF", "header", "hdr", para("First")), part("rIdD", "header", "hdr", para("Default"))],
      }, { paginate: true });

      const shown = headers(div);
      expect(shown.length).toBeGreaterThanOrEqual(3);
      expect(shown).toEqual(["First", ...shown.slice(1).map(() => "Default")]);
    });

    it("shows even and odd headers by the final page position", async () => {
      div = await render({
        body: paras(16) + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + paras(1),
        sectPr: `<w:headerReference w:type="default" r:id="rIdD"/><w:headerReference w:type="even" r:id="rIdE"/>${PAGE}`,
        parts: [
          part("rIdD", "header", "hdr", para("Odd")), part("rIdE", "header", "hdr", para("Even")),
          part("rIdS", "settings", "settings", "<w:evenAndOddHeaders/>"),
        ],
      }, { paginate: true });

      const shown = headers(div);
      expect(shown.length).toBeGreaterThanOrEqual(3);
      shown.forEach((text, i) => expect(text).toBe(i % 2 ? "Even" : "Odd"));
    });

    it("keeps a footnote on the page of its reference", async () => {
      div = await render({
        body: `<w:p><w:r><w:t>Ref</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` + paras(24),
        parts: [part("rIdN", "footnotes", "footnotes", `<w:footnote w:id="1">${para("Note")}</w:footnote>`)],
      });
      const before = bodyText(div);

      docx.paginate(div);

      expect(pages(div).length).toBeGreaterThanOrEqual(3);
      expect(div.querySelectorAll("section.docx > ol").length).toBe(1);
      expect(pages(div)[0].querySelector(":scope > ol")).not.toBeNull();
      expect(bodyText(div)).toBe(before);
      pages(div).forEach(page => expect(fits(page)).toBe(true));
    });

    it("pushes the last body lines down instead of giving the endnotes a page of their own", async () => {
      for (let count = 36; count <= 56; count++) {
        div?.remove();
        div = await render({
          body: Array.from({ length: count }, (_, i) => para(`${i} ${SENTENCE}`)).join("")
            + `<w:p><w:r><w:t>Ref</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>`,
          parts: [part("rIdE", "endnotes", "endnotes", `<w:endnote w:id="1">${para("End")}${para("Second end line")}</w:endnote>`)],
        });

        docx.paginate(div);

        const notesPage = pages(div).find(p => p.querySelector(":scope > ol"));
        expect(notesPage.querySelector(":scope > article > p")).withContext(`${count} paragraphs`).not.toBeNull();
        pages(div).forEach(page => expect(fits(page)).withContext(`${count} paragraphs`).toBe(true));
      }
    });

    it("continues a vertically merged cell on the next page", async () => {
      const rows = Array.from({ length: 80 }, (_, i) => '<w:tr>'
        + `<w:tc><w:tcPr><w:vMerge${i ? "" : ' w:val="restart"'}/></w:tcPr>${para(i ? "" : "Merged")}</w:tc>`
        + `<w:tc>${para(`Row ${i}`)}</w:tc></w:tr>`).join("");
      div = await render({ body: `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6000"/></w:tblGrid>${rows}</w:tbl>` });

      docx.paginate(div);

      const tables = Array.from(div.querySelectorAll("table"));
      expect(tables.length).toBeGreaterThan(1);

      for (const table of tables) {
        const [merged, text] = Array.from(table.rows[0].cells);
        expect(getComputedStyle(merged).display).not.toBe("none");
        expect(merged.rowSpan).toBe(table.rows.length);
        expect(text.textContent).toMatch(/^Row \d+$/);
      }
    });

    it("renders a page with an empty body", async () => {
      const bodies = ['<w:p><w:r><w:br w:type="page"/></w:r></w:p>', '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>', ""];

      for (const body of bodies) {
        div?.remove();
        div = await render({ body }, { paginate: true });
        expect(pages(div).length).toBeGreaterThan(0);
      }
    });
  });
})();
