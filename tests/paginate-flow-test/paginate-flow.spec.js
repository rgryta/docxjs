(() => {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
  const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';
  const SENTENCE = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ";

  const lorem = (sentences) => SENTENCE.repeat(sentences).trim();
  const run = (text, rPr = "") => `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const para = (text) => `<w:p>${run(text)}</w:p>`;
  const paras = (count, sentences = 4) => Array.from({ length: count }, (_, i) => para(`${i} ${lorem(sentences)}`)).join("");
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const sectionEnd = (sectPr) => `<w:p><w:pPr><w:sectPr>${sectPr}</w:sectPr></w:pPr></w:p>`;
  const headerPart = (id, text, type = "header") => ({
    id, type, target: `${id}.xml`,
    content: `<w:${type === "header" ? "hdr" : "ftr"} ${NS}>${typeof text === "string" ? para(text) : text.join("")}</w:${type === "header" ? "hdr" : "ftr"}>`,
  });

  function build({ body, sectPr = PAGE, parts = [] }) {
    const zip = new JSZip();
    zip.file("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file("word/document.xml", `<w:document ${NS}><w:body>${body}<w:sectPr>${sectPr}</w:sectPr></w:body></w:document>`);
    zip.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + parts.map(p => `<Relationship Id="${p.id}" Type="${REL}${p.type}" Target="${p.target}"/>`).join("") + "</Relationships>");

    for (const p of parts)
      zip.file(`word/${p.target}`, p.content, { base64: !!p.base64 });

    return zip.generateAsync({ type: "blob" });
  }

  async function render(doc, options = {}) {
    const div = document.createElement("div");
    document.body.appendChild(div);
    await docx.renderAsync(await build(doc), div, null, options);
    return div;
  }

  const pages = (div) => Array.from(div.querySelectorAll("section.docx"));
  const pageHeight = (page) => parseFloat(getComputedStyle(page).minHeight);
  const fits = (page) => page.getBoundingClientRect().height <= pageHeight(page) + 1;
  const shown = (page, tag) => Array.from(page.children).filter(c => c.tagName === tag && getComputedStyle(c).display !== "none");

  function flowText(div) {
    return Array.from(div.querySelectorAll("section.docx > article, section.docx > ol")).map(el => {
      const copy = el.cloneNode(true);
      copy.querySelectorAll("[data-docx-repeat]").forEach(r => r.remove());
      return copy.textContent;
    }).join("");
  }

  function textLines(el) {
    const range = document.createRange();
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    const tops = new Map();

    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      range.selectNodeContents(n);

      for (const r of Array.from(range.getClientRects())) {
        if (r.width)
          tops.set(Math.round(r.top), Math.max(tops.get(Math.round(r.top)) ?? 0, r.bottom));
      }
    }

    return Array.from(tops.entries()).sort((a, b) => a[0] - b[0]).map(([top, bottom]) => ({ top, bottom }));
  }

  // Moves the page bottom so the content limit sits at y (viewport coordinates).
  function setLimit(page, y) {
    const style = getComputedStyle(page);
    page.style.minHeight = `${y - page.getBoundingClientRect().top + parseFloat(style.paddingBottom)}px`;
  }

  describe("Paginate flow", function () {
    let div;

    afterEach(() => div?.remove());

    it("keeps the paragraph line height of a run (no rounding to whole pixels)", async () => {
      div = await render({ body: `<w:p><w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr>${run("Text", '<w:sz w:val="20"/>')}</w:p>` });
      const span = div.querySelector("article p span");
      const style = getComputedStyle(span);
      expect(parseFloat(style.fontSize)).toBeCloseTo(40 / 3, 1);
      expect(parseFloat(style.lineHeight)).toBeCloseTo(1.15 * parseFloat(style.fontSize), 1);
    });

    it("fills pages across continuous sections without duplicating them", async () => {
      const refs = '<w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>';
      const sect = `<w:type w:val="continuous"/>${refs}${PAGE}`;
      div = await render({
        body: paras(9) + sectionEnd(sect) + paras(9) + sectionEnd(sect) + paras(9),
        sectPr: sect,
        parts: [headerPart("rIdH", "Header"), headerPart("rIdF", "Footer", "footer")],
      }, { ignoreLastRenderedPageBreak: false });
      expect(pages(div).length).toBe(1);
      expect(div.querySelectorAll("section.docx > article").length).toBe(3);
      const before = flowText(div);

      expect(docx.paginate(div)).toBe(true);

      expect(pages(div).length).toBeGreaterThanOrEqual(3);
      expect(flowText(div)).toBe(before);

      for (const page of pages(div)) {
        expect(fits(page)).toBe(true);
        expect(shown(page, "HEADER").length).toBe(1);
        expect(shown(page, "FOOTER").length).toBe(1);
      }
    });

    it("fills a page down to the footer when header and footer sit in the margins", async () => {
      div = await render({
        body: paras(1, 120),
        sectPr: `<w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>${PAGE}`,
        parts: [headerPart("rIdH", "Header"), headerPart("rIdF", "Footer", "footer")],
      });

      docx.paginate(div);

      for (const page of pages(div).slice(0, -1)) {
        const lines = Array.from(page.querySelectorAll("article > p")).flatMap(textLines);
        const last = lines[lines.length - 1];
        const footerTop = shown(page, "FOOTER")[0].getBoundingClientRect().top;
        expect(fits(page)).toBe(true);
        expect(footerTop - last.bottom).toBeLessThan(last.bottom - last.top + 1);
      }
    });

    it("shows the first-page header only on the first page", async () => {
      div = await render({
        body: paras(24),
        sectPr: `<w:headerReference w:type="first" r:id="rIdF"/><w:headerReference w:type="default" r:id="rIdD"/><w:titlePg/>${PAGE}`,
        parts: [headerPart("rIdF", "First"), headerPart("rIdD", "Default")],
      }, { paginate: true });

      const headers = pages(div).map(p => shown(p, "HEADER").map(h => h.textContent).join("|"));
      expect(headers.length).toBeGreaterThanOrEqual(3);
      expect(headers).toEqual(["First", ...headers.slice(1).map(() => "Default")]);
      expect(div.querySelectorAll('[data-docx-hf="first"]').length).toBe(1);
    });

    it("shows even and odd headers by the final page position", async () => {
      div = await render({
        body: paras(16) + pageBreak + paras(1),
        sectPr: `<w:headerReference w:type="default" r:id="rIdD"/><w:headerReference w:type="even" r:id="rIdE"/>${PAGE}`,
        parts: [
          headerPart("rIdD", "Odd"), headerPart("rIdE", "Even"),
          { id: "rIdS", type: "settings", target: "settings.xml", content: `<w:settings ${NS}><w:evenAndOddHeaders/></w:settings>` },
        ],
      }, { paginate: true });

      const headers = pages(div).map(p => shown(p, "HEADER").map(h => h.textContent).join("|"));
      expect(headers.length).toBeGreaterThanOrEqual(3);
      headers.forEach((text, i) => expect(text).toBe(i % 2 ? "Even" : "Odd"));
    });

    it("keeps one notes list, on the page after the last body block", async () => {
      div = await render({
        body: `<w:p>${run(lorem(2))}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>` + paras(24),
        parts: [{ id: "rIdN", type: "footnotes", target: "footnotes.xml", content: `<w:footnotes ${NS}><w:footnote w:id="1">${para("Note")}</w:footnote></w:footnotes>` }],
      });
      const before = flowText(div);

      docx.paginate(div);

      expect(pages(div).length).toBeGreaterThanOrEqual(3);
      expect(div.querySelectorAll("section.docx > ol").length).toBe(1);
      expect(flowText(div)).toBe(before);
      pages(div).forEach(page => expect(fits(page)).toBe(true));
    });

    it("pushes body lines down instead of giving the notes a page of their own", async () => {
      for (let count = 36; count <= 56; count++) {
        div?.remove();
        div = await render({
          body: `<w:p>${run("Ref")}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>` + paras(count, 1),
          parts: [{ id: "rIdN", type: "footnotes", target: "footnotes.xml", content: `<w:footnotes ${NS}><w:footnote w:id="1">${para("Note")}${para("Second note line")}</w:footnote></w:footnotes>` }],
        });

        docx.paginate(div);

        const notesPage = pages(div).find(p => p.querySelector(":scope > ol"));
        expect(notesPage.querySelector(":scope > article > p")).withContext(`${count} paragraphs`).not.toBeNull();
        pages(div).forEach(page => expect(fits(page)).withContext(`${count} paragraphs`).toBe(true));
      }
    });

    it("splits the pages before a notes list taller than a page", async () => {
      div = await render({
        body: `<w:p>${run(lorem(2))}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>` + paras(24),
        parts: [{ id: "rIdN", type: "footnotes", target: "footnotes.xml", content: `<w:footnotes ${NS}><w:footnote w:id="1">${paras(30)}</w:footnote></w:footnotes>` }],
      });
      const before = flowText(div);

      docx.paginate(div);

      const all = pages(div);
      const notesPage = all.findIndex(p => p.querySelector(":scope > ol"));
      expect(div.querySelectorAll("section.docx > ol").length).toBe(1);
      expect(notesPage).toBeGreaterThanOrEqual(2);
      all.slice(0, notesPage).forEach(page => expect(fits(page)).toBe(true));
      expect(flowText(div)).toBe(before);
    });

    it("continues a vertically merged cell on the next page in its own column", async () => {
      const rows = Array.from({ length: 80 }, (_, i) => '<w:tr>'
        + `<w:tc><w:tcPr><w:vMerge${i ? "" : ' w:val="restart"'}/></w:tcPr>${para(i ? "" : "Merged")}</w:tc>`
        + `<w:tc>${para(`Row ${i}`)}</w:tc></w:tr>`).join("");
      div = await render({ body: `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6000"/></w:tblGrid>${rows}</w:tbl>` });
      const firstTable = div.querySelector("table");
      const textLeft = firstTable.rows[0].cells[1].getBoundingClientRect().left - firstTable.getBoundingClientRect().left;

      docx.paginate(div);

      const tables = Array.from(div.querySelectorAll("table"));
      expect(tables.length).toBeGreaterThan(1);

      for (const table of tables.slice(1)) {
        const [merged, text] = Array.from(table.rows[0].cells);
        expect(getComputedStyle(merged).display).not.toBe("none");
        expect(merged.rowSpan).toBe(table.rows.length);
        expect(text.textContent).toMatch(/^Row \d+$/);
        expect(text.getBoundingClientRect().left - table.getBoundingClientRect().left).toBeCloseTo(textLeft, 0);
      }

      expect(tables[0].rows[0].cells[0].rowSpan).toBe(tables[0].rows.length);
    });

    it("splits a paragraph by its text lines, not by a picture below the text", async () => {
      const drawing = '<w:r><w:br/></w:r><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="342900"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        + '<pic:pic><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill><pic:spPr><a:xfrm><a:ext cx="914400" cy="342900"/></a:xfrm></pic:spPr></pic:pic>'
        + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
      div = await render({
        body: para("Before") + `<w:p><w:r><w:t>One</w:t><w:br/><w:t>Two</w:t><w:br/><w:t>Three</w:t></w:r>${drawing}</w:p>`,
        parts: [{ id: "rIdImg", type: "image", target: "media/image1.png", content: PNG, base64: true }],
      });
      const [, target] = Array.from(div.querySelectorAll("article > p"));
      const lines = textLines(target);
      expect(lines.length).toBe(3);
      setLimit(pages(div)[0], lines[1].bottom + 2);

      docx.paginate(div);

      for (const rest of Array.from(div.querySelectorAll("p.docx-continued"))) {
        expect(textLines(rest).length).toBeGreaterThanOrEqual(2);
        expect(textLines(rest.closest("section").previousElementSibling.querySelector("article:last-of-type > p:last-child")).length).toBeGreaterThanOrEqual(2);
      }
      expect(pages(div)[1].querySelector("article > p").textContent).toBe(target.textContent);
    });

    it("moves a later section whole when its first block overflows after a keep-with-next block", async () => {
      const sect = `<w:type w:val="continuous"/>${PAGE}`;
      div = await render({
        body: paras(2) + `<w:p><w:pPr><w:keepNext/><w:sectPr>${sect}</w:sectPr></w:pPr>${run("Heading")}</w:p>` + paras(3),
        sectPr: sect,
      }, { ignoreLastRenderedPageBreak: false });
      const [first, second] = Array.from(div.querySelectorAll("section.docx > article"));
      const firstText = first.textContent;
      const secondText = second.textContent;
      const heading = first.lastElementChild;
      expect(heading.textContent).toBe("Heading");
      expect(getComputedStyle(heading).breakAfter).toBe("avoid");
      setLimit(pages(div)[0], heading.getBoundingClientRect().bottom + 2);

      docx.paginate(div);

      const [page1, ...rest] = pages(div);
      expect(Array.from(page1.querySelectorAll("article")).map(a => a.textContent)).toEqual([firstText]);
      expect(rest.map(p => Array.from(p.querySelectorAll("article")).map(a => a.textContent).join("")).join("")).toBe(secondText);
    });

    it("stops at the time budget and keeps the pages split so far", async () => {
      div = await render({ body: paras(48) });
      const before = flowText(div);
      let now = 0;
      spyOn(performance, "now").and.callFake(() => (now += 400));

      expect(docx.paginate(div, "docx", { budgetMs: 1000 })).toBe(false);

      const all = pages(div);
      expect(all.length).toBeGreaterThan(1);
      all.slice(0, -1).forEach(page => expect(fits(page)).toBe(true));
      expect(fits(all[all.length - 1])).toBe(false);
      expect(flowText(div)).toBe(before);
    });

    it("changes nothing with a zero budget", async () => {
      div = await render({ body: paras(48) });
      const before = div.innerHTML;

      expect(docx.paginate(div, "docx", { budgetMs: 0 })).toBe(false);
      expect(div.innerHTML).toBe(before);
    });

    it("does not split a page under a header taller than the page", async () => {
      div = await render({
        body: paras(6),
        sectPr: `<w:headerReference w:type="default" r:id="rIdH"/>${PAGE}`,
        parts: [headerPart("rIdH", [paras(60)])],
      });

      expect(docx.paginate(div)).toBe(true);
      expect(pages(div).length).toBe(1);
    });
  });
})();
