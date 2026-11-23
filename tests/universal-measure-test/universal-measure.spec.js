(() => {
  // Chrome writes an atLeast line back in px (36pt becomes 48px), so compare the length.
  function extraLinePoints(element) {
      const match = /calc\(100% \+ ([\d.]+)(pt|px)\)/.exec(element.style.lineHeight);
      if (!match) return NaN;
      const value = parseFloat(match[1]);
      return match[2] === "px" ? value * 0.75 : value;
  }

  describe("universal measure lengths", function () {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    const paragraphs = [
      ["sz-3mm", '', '<w:sz w:val="3mm"/>'],
      ["sz-4.3mm", '', '<w:sz w:val="4.3mm"/>'],
      ["sz-22", '', '<w:sz w:val="22"/>'],
      ["ind-1cm", '<w:ind w:left="1cm"/>', ''],
      ["ind-1in", '<w:ind w:left="1in"/>', ''],
      ["ind-25.4mm", '<w:ind w:left="25.4mm"/>', ''],
      ["ind-1pc", '<w:ind w:left="1pc"/>', ''],
      ["ind-1pi", '<w:ind w:left="1pi"/>', ''],
      ["ind-.5in", '<w:ind w:left=".5in"/>', ''],
      ["ind-1440", '<w:ind w:left="1440"/>', ''],
      ["ind-12pt", '<w:ind w:left="12pt"/>', ''],
      ["line-12pt-exact", '<w:spacing w:line="12pt" w:lineRule="exact"/>', ''],
      ["line-0.5in-atLeast", '<w:spacing w:line="0.5in" w:lineRule="atLeast"/>', ''],
      ["line-300-norule", '<w:spacing w:line="300"/>', ''],
      ["line-1in-auto", '<w:spacing w:line="1in" w:lineRule="auto"/>', ''],
      ["line-240-exact", '<w:spacing w:line="240" w:lineRule="exact"/>', ''],
      ["line-360-atLeast", '<w:spacing w:line="360" w:lineRule="atLeast"/>', ''],
    ];

    function paragraphXml([text, pPr, rPr]) {
      return `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;
    }

    async function buildDocx() {
      const zip = new JSZip();
      zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  </Types>`);
      zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>
  </Relationships>`);
      zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHeader" Type="${REL}/header" Target="header1.xml"/>
  </Relationships>`);
      zip.file("word/header1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <w:hdr ${W}><w:p><w:r><w:t>header</w:t></w:r></w:p></w:hdr>`);
      zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <w:document ${W}><w:body>
  ${paragraphs.map(paragraphXml).join("\n")}
  <w:sectPr>
  <w:headerReference w:type="default" r:id="rIdHeader"/>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1in" w:right="1in" w:bottom="1in" w:left="1in" w:header="0.5in" w:footer="0.5in" w:gutter="0"/>
  </w:sectPr>
  </w:body></w:document>`);
      return zip.generateAsync({ type: "blob" });
    }

    let container;

    beforeAll(async () => {
      container = document.createElement("div");
      document.body.appendChild(container);
      await docx.renderAsync(await buildDocx(), container);
    });

    afterAll(() => container.remove());

    function paragraph(text) {
      return Array.from(container.querySelectorAll("article p")).find(p => p.textContent === text);
    }

    function points(value) {
      return parseFloat(value);
    }

    function pxToPoints(value) {
      return parseFloat(value) * 0.75;
    }

    it("reads mm font sizes as points", () => {
      expect(points(paragraph("sz-3mm").querySelector("span").style.fontSize)).toBeCloseTo(8.5, 2);
      expect(points(paragraph("sz-4.3mm").querySelector("span").style.fontSize)).toBeCloseTo(12.19, 2);
    });

    it("reads inch page and header margins as points", () => {
      const page = container.querySelector("section.docx");
      expect(points(page.style.paddingLeft)).toBeCloseTo(72, 2);
      expect(points(page.style.paddingTop)).toBeCloseTo(72, 2);
      expect(pxToPoints(getComputedStyle(page.querySelector("header")).marginTop)).toBeCloseTo(-36, 2);
    });

    it("reads every universal measure unit in indents", () => {
      const expected = { "ind-1cm": 28.35, "ind-1in": 72, "ind-25.4mm": 72, "ind-1pc": 12, "ind-1pi": 12, "ind-.5in": 36 };
      for (const [text, value] of Object.entries(expected))
        expect(points(paragraph(text).style.getPropertyValue("margin-inline-start"))).withContext(text).toBeCloseTo(value, 2);
    });

    it("keeps a pt line as written for an exact rule", () => {
      const p = paragraph("line-12pt-exact");
      expect(points(p.style.lineHeight)).toBeCloseTo(12, 2);
      expect(points(p.style.minHeight)).toBeCloseTo(12, 2);
    });

    it("converts a unit line for an atLeast rule", () => {
      expect(extraLinePoints(paragraph("line-0.5in-atLeast"))).toBeCloseTo(36, 2);
    });

    it("treats a line without lineRule as auto", () => {
      expect(paragraph("line-300-norule").style.lineHeight).toBe("1.25");
    });

    it("does not collapse the lines when an auto rule carries a unit", () => {
      expect(paragraph("line-1in-auto").style.lineHeight).toBe("");
    });

    it("keeps unitless and pt lengths unchanged", () => {
      expect(points(paragraph("ind-1440").style.getPropertyValue("margin-inline-start"))).toBeCloseTo(72, 2);
      expect(points(paragraph("sz-22").querySelector("span").style.fontSize)).toBeCloseTo(11, 2);
      expect(paragraph("ind-12pt").style.getPropertyValue("margin-inline-start")).toBe("12pt");
    });

    it("keeps unitless exact and atLeast lines unchanged", () => {
      const exact = paragraph("line-240-exact");
      expect(exact.style.lineHeight).toBe("12pt");
      expect(exact.style.minHeight).toBe("12pt");
      expect(extraLinePoints(paragraph("line-360-atLeast"))).toBeCloseTo(18, 2);
    });
  });
})();
