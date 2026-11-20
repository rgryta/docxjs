(() => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
  const CLASS = "repeated-tabs";
  let container;

  afterEach(() => container?.remove());

  it("remeasures right tabs in repeated headers, footers and table headings", async () => {
    const zip = new JSZip();
    const tabParagraph = '<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="6000"/></w:tabs></w:pPr><w:r><w:t>Left</w:t><w:tab/></w:r><w:r><w:t>Right</w:t></w:r></w:p>';
    zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>');
    zip.file("_rels/.rels", `<Relationships xmlns="${PKG}"><Relationship Id="doc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="${PKG}"><Relationship Id="header" Type="${REL}/header" Target="header.xml"/><Relationship Id="footer" Type="${REL}/footer" Target="footer.xml"/></Relationships>`);
    zip.file("word/header.xml", `<w:hdr ${W}>${tabParagraph}</w:hdr>`);
    zip.file("word/footer.xml", `<w:ftr ${W}>${tabParagraph}</w:ftr>`);
    const rows = Array.from({ length: 30 }, () => '<w:tr><w:tc><w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact"/></w:pPr><w:r><w:t>Row</w:t></w:r></w:p></w:tc></w:tr>').join("");
    zip.file("word/document.xml", `<w:document ${W} xmlns:r="${REL}"><w:body><w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc>${tabParagraph}</w:tc></w:tr>${rows}</w:tbl><w:sectPr><w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/><w:pgSz w:w="12240" w:h="7200"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="360" w:footer="360"/></w:sectPr></w:body></w:document>`);
    container = document.createElement("div");
    document.body.appendChild(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container, null, { experimental: true, className: CLASS });
    docx.remeasureTabStops(container, CLASS);
    docx.paginate(container, CLASS);
    const pages = [...container.querySelectorAll(`section.${CLASS}`)];
    expect(pages.length).toBeGreaterThan(1);
    const repeated = pages[1];
    for (const selector of ["header p", "footer p", "thead p"]) {
      const p = repeated.querySelector(selector);
      expect(p).withContext(selector).not.toBeNull();
      const right = [...p.querySelectorAll("span")].find(s => s.textContent === "Right");
      right.style.fontSize = "22pt";
    }
    container.querySelector(`.${CLASS}-wrapper`).style.zoom = "0.75";
    docx.remeasureTabStops(container, CLASS);
    for (const selector of ["header p", "footer p", "thead p"]) {
      const p = repeated.querySelector(selector);
      const right = [...p.querySelectorAll("span")].find(s => s.textContent === "Right");
      const expected = p.getBoundingClientRect().left + 300;
      expect(Math.abs(right.getBoundingClientRect().right - expected)).withContext(selector).toBeLessThan(1);
    }
  });
})();
