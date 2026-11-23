describe("default font size", function () {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const normalStyle = rPr => `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>${rPr}</w:style>`;
  const stylesXml = (defaultRPr, styles) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr>${defaultRPr}</w:rPr></w:rPrDefault></w:docDefaults>${styles}</w:styles>`;

  const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

  const plainParagraph = `<w:p><w:r><w:t>plain</w:t></w:r></w:p>`;

  async function buildDocx({ styles, numbering, body }) {
    const parts = [["styles.xml", "styles", styles], ["numbering.xml", "numbering", numbering]].filter(([, , xml]) => xml);
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>
</Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${parts.map(([file, type]) => `<Relationship Id="rId${type}" Type="${REL}/${type}" Target="${file}"/>`).join("\n")}
</Relationships>`);
    for (const [file, , xml] of parts)
      zip.file(`word/${file}`, xml);
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
${body}
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`);
    return zip.generateAsync({ type: "blob" });
  }

  const files = {
    unsized: {
      styles: stylesXml('<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:szCs w:val="24"/>',
        normalStyle("") + '<w:style w:type="character" w:styleId="Big"><w:name w:val="Big"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style>'),
      numbering: numberingXml,
      body: [
        plainParagraph,
        `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>numbered</w:t></w:r></w:p>`,
        `<w:p/>`,
        `<w:p><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>sized</w:t><w:tab/><w:sym w:font="Symbol" w:char="F0B7"/></w:r></w:p>`,
        `<w:p><w:r><w:rPr><w:rStyle w:val="Big"/></w:rPr><w:t>styled</w:t></w:r></w:p>`,
      ].join("\n"),
    },
    noStyles: { body: plainParagraph },
    defaultsSized: { styles: stylesXml('<w:sz w:val="22"/>', normalStyle("")), body: plainParagraph },
    normalSized: { styles: stylesXml("", normalStyle('<w:rPr><w:sz w:val="18"/></w:rPr>')), body: plainParagraph },
  };

  const containers = {};

  beforeAll(async () => {
    for (const [name, file] of Object.entries(files)) {
      const container = document.createElement("div");
      document.body.appendChild(container);
      await docx.renderAsync(await buildDocx(file), container, null, { className: `dfs-${name.toLowerCase()}` });
      containers[name] = container;
    }
  });

  afterAll(() => Object.values(containers).forEach(c => c.remove()));

  function paragraph(name, text) {
    return Array.from(containers[name].querySelectorAll("article p")).find(p => p.textContent.startsWith(text));
  }

  // Chrome resolves pt to px in computed values.
  function points(element, pseudo) {
    return parseFloat(getComputedStyle(element, pseudo).fontSize) * 0.75;
  }

  function styleText(name) {
    return Array.from(containers[name].querySelectorAll("style")).map(s => s.textContent).join("\n");
  }

  it("draws text at 10pt when no default sets a size", () => {
    const p = paragraph("unsized", "plain");
    expect(points(p.querySelector("span"))).toBeCloseTo(10, 2);
    expect(points(p)).toBeCloseTo(10, 2);
  });

  it("draws text at 10pt in a file without styles", () => {
    expect(points(paragraph("noStyles", "plain").querySelector("span"))).toBeCloseTo(10, 2);
  });

  it("draws list markers at 10pt", () => {
    expect(points(paragraph("unsized", "numbered"), "::before")).toBeCloseTo(10, 2);
  });

  it("sizes an empty paragraph from 10pt", () => {
    const empty = Array.from(containers.unsized.querySelectorAll("article p")).find(p => p.textContent === "");
    expect(empty.getBoundingClientRect().height * 0.75).toBeCloseTo(10, 0);
  });

  it("keeps a docDefaults size and adds no rule", () => {
    expect(points(paragraph("defaultsSized", "plain").querySelector("span"))).toBeCloseTo(11, 2);
    expect(styleText("defaultsSized")).not.toMatch(/font-size:\s*10pt/);
  });

  it("keeps a Normal size and adds no rule", () => {
    expect(points(paragraph("normalSized", "plain").querySelector("span"))).toBeCloseTo(9, 2);
    expect(styleText("normalSized")).not.toMatch(/font-size:\s*10pt/);
  });

  it("keeps a run size on the run and its tab and symbol spans", () => {
    const run = paragraph("unsized", "sized").querySelector("span");
    const nested = run.querySelectorAll("span");
    expect(points(run)).toBeCloseTo(14, 2);
    expect(nested.length).toBe(2);
    nested.forEach(span => expect(points(span)).toBeCloseTo(14, 2));
  });

  it("keeps a character style size", () => {
    expect(points(paragraph("unsized", "styled").querySelector("span"))).toBeCloseTo(16, 2);
  });
});
