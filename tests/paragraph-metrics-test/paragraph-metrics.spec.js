(() => {
  const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
  let container;
  afterEach(() => container?.remove());

  it("uses document and paragraph font metrics for line struts instead of browser defaults", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>');
    zip.file("_rels/.rels", `<Relationships xmlns="${PKG}"><Relationship Id="doc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="${PKG}"><Relationship Id="styles" Type="${REL}/styles" Target="styles.xml"/></Relationships>`);
    zip.file("word/styles.xml", `<w:styles xmlns:w="${WORD}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="16"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Small"><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="12"/><w:b/></w:rPr></w:style></w:styles>`);
    const body = Array.from({ length: 500 }, (_, i) => `<w:p>${i === 499 ? '<w:pPr><w:pStyle w:val="Small"/></w:pPr>' : ""}<w:r><w:t>Line ${i + 1}</w:t></w:r></w:p>`).join("");
    zip.file("word/document.xml", `<w:document xmlns:w="${WORD}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
    container = document.createElement("div");
    document.body.appendChild(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container, null, { className: "paragraph-metrics" });
    const first = container.querySelector("p");
    const last = container.querySelectorAll("p")[499];
    expect(parseFloat(getComputedStyle(first).fontSize)).toBeCloseTo(8 * 96 / 72, 3);
    expect(getComputedStyle(first).fontFamily).toBe(getComputedStyle(first.querySelector("span")).fontFamily);
    expect(parseFloat(getComputedStyle(last).fontSize)).toBeCloseTo(8, 3);
    expect(getComputedStyle(last).fontWeight).toBe("700");
    expect(docx.paginate(container, "paragraph-metrics")).toBeTrue();
    expect(container.querySelectorAll("section").length).toBeGreaterThan(1);
    expect(container.querySelectorAll("p").length).toBe(500);
  });
})();
