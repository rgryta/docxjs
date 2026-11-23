(function () {
  describe("default style selector specificity", function () {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:color w:val="000000"/><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Link"><w:name w:val="Link"/><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>
</w:styles>`;

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
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>plain</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Link"/></w:rPr><w:t>linked</w:t></w:r></w:p>
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

    const spans = () => Array.from(container.querySelectorAll("article p span"));

    it("lets a paragraph style size win even when Normal is listed after it", () => {
      const plain = spans().find(s => s.textContent == "plain");
      expect(parseFloat(getComputedStyle(plain).fontSize) * 0.75).toBeCloseTo(20, 2);
    });

  });
})();
