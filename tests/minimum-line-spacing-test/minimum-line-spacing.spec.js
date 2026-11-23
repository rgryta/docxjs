(() => {
  const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
  const CLASS = "minimum-line-spacing";
  let container;

  async function render(body) {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>');
    zip.file("_rels/.rels", `<Relationships xmlns="${PKG}"><Relationship Id="doc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="${PKG}"><Relationship Id="styles" Type="${REL}/styles" Target="styles.xml"/></Relationships>`);
    zip.file("word/styles.xml", `<w:styles xmlns:w="${WORD}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="240" w:lineRule="atLeast"/></w:pPr></w:style></w:styles>`);
    zip.file("word/document.xml", `<w:document xmlns:w="${WORD}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
    container = document.createElement("div");
    document.body.appendChild(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container, null, { className: CLASS });
    return [...container.querySelectorAll("article > p")];
  }

  const paragraph = (props = "", size = "24", text = "Line") => `<w:p><w:pPr>${props}</w:pPr><w:r><w:rPr><w:sz w:val="${size}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const spacing = (line, rule) => `<w:spacing w:line="${line}" w:lineRule="${rule}"/>`;
  afterEach(() => container?.remove());

  it("uses a minimum strut while larger runs keep their natural line boxes", async () => {
    const ps = await render([
      paragraph(),
      paragraph("", "48"),
      paragraph(spacing(480, "atLeast")),
      paragraph(spacing(480, "atLeast"), "24", ""),
      paragraph(spacing("24pt", "atLeast")),
    ].join(""));
    expect(parseFloat(getComputedStyle(ps[0]).lineHeight)).toBeCloseTo(16, 2);
    expect(ps[0].getBoundingClientRect().height).toBeLessThan(24);
    expect(ps[1].getBoundingClientRect().height).toBeGreaterThan(ps[0].getBoundingClientRect().height * 1.5);
    expect(getComputedStyle(ps[1].querySelector("span")).lineHeight).toBe("normal");
    for (const p of ps.slice(2)) expect(p.getBoundingClientRect().height).toBeCloseTo(32, 0);
  });

  it("lets direct exact and auto spacing override inherited atLeast run metrics", async () => {
    const ps = await render([
      paragraph(),
      paragraph(spacing(240, "exact"), "48"),
      paragraph(spacing(480, "auto")),
    ].join(""));
    expect(getComputedStyle(ps[0].querySelector("span")).lineHeight).toBe("normal");
    expect(parseFloat(getComputedStyle(ps[1].querySelector("span")).lineHeight)).toBeCloseTo(16, 2);
    expect(parseFloat(getComputedStyle(ps[2].querySelector("span")).lineHeight)).toBeCloseTo(32, 2);
    expect(ps[2].getBoundingClientRect().height).toBeCloseTo(32, 0);
  });

  it("paginates repeated minimum-spaced lines without adding a font size per line", async () => {
    await render(Array.from({ length: 200 }, () => paragraph()).join(""));
    expect(docx.paginate(container, CLASS)).toBeTrue();
    expect(container.querySelectorAll(`section.${CLASS}`).length).toBeLessThanOrEqual(5);
    expect(container.querySelectorAll("article p").length).toBe(200);
  });
})();
