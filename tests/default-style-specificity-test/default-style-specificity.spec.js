(function () {
  describe("default style selector specificity", function () {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:color w:val="00AA00"/><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="216" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`;

    const body = `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>heading</w:t></w:r></w:p>
<w:p><w:r><w:t>plain</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

    async function build() {
      const zip = new JSZip();
      zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
      zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
      zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdstyles" Type="${REL}/styles" Target="styles.xml"/></Relationships>`);
      zip.file("word/styles.xml", styles);
      zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
${body}
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`);
      return zip.generateAsync({ type: "blob" });
    }

    let container;

    beforeAll(async () => {
      container = document.createElement("div");
      document.body.appendChild(container);
      await docx.renderAsync(await build(), container, null, { className: "dss" });
    });

    afterAll(() => container.remove());

    // Chrome resolves pt to px in computed values.
    const points = element => parseFloat(getComputedStyle(element).fontSize) * 0.75;
    const span = text => Array.from(container.querySelectorAll("article span")).find(s => s.textContent == text);

    it("lets a paragraph style size win even when Normal is listed after it", () => {
      expect(points(span("heading"))).toBeCloseTo(20, 2);
    });

    it("gives runs of an unstyled paragraph Normal's values over the document defaults", () => {
      expect(points(span("plain"))).toBeCloseTo(9, 2);
      expect(getComputedStyle(span("plain")).color).toBe("rgb(0, 170, 0)");
    });

    it("keeps Normal's size on the paragraph itself, so its line box is sized from it", () => {
      expect(points(span("plain").closest("p"))).toBeCloseTo(9, 2);
    });

    it("puts the default table style's cell margins on the cells", () => {
      expect(parseFloat(getComputedStyle(container.querySelector("article td")).paddingLeft) * 0.75).toBeCloseTo(10.8, 2);
    });
  });
})();
