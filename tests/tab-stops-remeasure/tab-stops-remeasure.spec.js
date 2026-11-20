describe("tab stop measure", function () {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PX_PER_PT = 96 / 72;
  const TOLERANCE_PX = 1;

  function tabs(...stops) {
    return `<w:tabs>${stops.map(([val, pos, leader]) => `<w:tab w:val="${val}" w:pos="${pos}"${leader ? ` w:leader="${leader}"` : ""}/>`).join("")}</w:tabs>`;
  }

  function paragraph(pPr, ...parts) {
    const runs = parts.map(part => part === "\t" ? "<w:r><w:tab/></w:r>" : `<w:r><w:t xml:space="preserve">${part}</w:t></w:r>`);
    return `<w:p><w:pPr>${pPr}</w:pPr>${runs.join("")}</w:p>`;
  }

  async function buildDocx(body, { settings = true, styles = "" } = {}) {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>
</Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${settings ? `<Relationship Id="rIdSettings" Type="${REL}/settings" Target="settings.xml"/>` : ""}
${styles ? `<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>` : ""}
</Relationships>`);
    if (settings)
      zip.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${W}><w:defaultTabStop w:val="720"/></w:settings>`);
    if (styles)
      zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>${styles}</w:styles>`);
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
${body.join("\n")}
<w:sectPr>
<w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
</w:body></w:document>`);
    return zip.generateAsync({ type: "blob" });
  }

  const containers = [];

  async function render(body, options, attach = true) {
    const container = document.createElement("div");
    containers.push(container);
    if (attach) document.body.appendChild(container);
    await docx.renderAsync(await buildDocx(body, options), container, null, { experimental: true });
    return container;
  }

  afterEach(() => containers.splice(0).forEach(c => c.remove()));

  function origin(container) {
    return container.querySelector("section.docx > article").getBoundingClientRect().left;
  }

  function run(container, text) {
    return Array.from(container.querySelectorAll("article span")).find(s => s.textContent === text);
  }

  function tabSpans(container) {
    return Array.from(container.querySelectorAll("span.docx-tab-stop"));
  }

  function expectAt(actualPx, container, points, label, zoom = 1) {
    expect(Math.abs(actualPx - origin(container) - points * PX_PER_PT * zoom)).withContext(label).toBeLessThan(TOLERANCE_PX);
  }

  function centre(rect) {
    return (rect.left + rect.right) / 2;
  }

  function dotLeft(container, text) {
    const node = run(container, text).firstChild;
    const range = document.createRange();
    range.setStart(node, text.indexOf("."));
    range.setEnd(node, text.indexOf(".") + 1);
    return range.getBoundingClientRect().left;
  }

  const separateStops = [
    paragraph(tabs(["left", 2880]), "a", "\t", "Left"),
    paragraph(tabs(["center", 4680]), "b", "\t", "Centre"),
    paragraph(tabs(["right", 9360, "dot"]), "c", "\t", "Right"),
  ];

  it("places left, centre and right stops with a dot leader", async () => {
    const container = await render(separateStops);
    docx.remeasureTabStops(container);

    expectAt(run(container, "Left").getBoundingClientRect().left, container, 144, "left");
    expectAt(centre(run(container, "Centre").getBoundingClientRect()), container, 234, "centre");
    expectAt(run(container, "Right").getBoundingClientRect().right, container, 468, "right");
    expect(tabSpans(container)[2].style.borderBottomStyle).toBe("dotted");
  });

  it("measures each tab of one paragraph after the tabs before it", async () => {
    const container = await render([
      paragraph(tabs(["left", 2880], ["center", 4680], ["right", 9360]), "a", "\t", "One", "\t", "Two", "\t", "Three"),
    ]);
    docx.remeasureTabStops(container);

    expectAt(run(container, "One").getBoundingClientRect().left, container, 144, "left");
    expectAt(centre(run(container, "Two").getBoundingClientRect()), container, 234, "centre");
    expectAt(run(container, "Three").getBoundingClientRect().right, container, 468, "right");
  });

  it("counts stops from the text column in an indented paragraph", async () => {
    const container = await render([
      paragraph(`<w:ind w:left="1440"/>${tabs(["right", 9360])}`, "a", "\t", "Indented"),
    ]);
    docx.remeasureTabStops(container);

    expectAt(run(container, "Indented").getBoundingClientRect().right, container, 468, "right");
  });

  it("gives the same widths on every call and follows a changed layout", async () => {
    const container = await render([
      paragraph(tabs(["left", 2880], ["right", 9360]), "a", "\t", "One", "\t", "Last"),
    ]);
    docx.remeasureTabStops(container);
    const first = tabSpans(container).map(s => s.getBoundingClientRect().width);
    docx.remeasureTabStops(container);
    const second = tabSpans(container).map(s => s.getBoundingClientRect().width);

    second.forEach((w, i) => expect(w).withContext(`tab ${i}`).toBeCloseTo(first[i], 2));

    run(container, "Last").style.fontSize = "24pt";
    docx.remeasureTabStops(container);

    expectAt(run(container, "Last").getBoundingClientRect().right, container, 468, "right after font change");
  });

  it("skips a detached render and measures it once attached", async () => {
    const container = await render(separateStops, undefined, false);
    await new Promise(r => setTimeout(r, 600));

    tabSpans(container).forEach(s => {
      expect(s.textContent).toBe("\u2003");
      expect(s.getAttribute("style")).toBeNull();
    });

    document.body.appendChild(container);
    docx.remeasureTabStops(container);

    expectAt(run(container, "Right").getBoundingClientRect().right, container, 468, "right");
  });

  it("uses 36pt default stops without settings or declared stops", async () => {
    const container = await render([paragraph("", "ab", "\t", "Default")], { settings: false });
    docx.remeasureTabStops(container);

    expectAt(run(container, "Default").getBoundingClientRect().left, container, 36, "default");
  });

  it("merges style stops with paragraph stops and a clear", async () => {
    const styles = `
<w:style w:type="paragraph" w:styleId="TabBase"><w:name w:val="Tab Base"/><w:pPr>${tabs(["center", 4680])}</w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TabStyle"><w:name w:val="Tab Style"/><w:basedOn w:val="TabBase"/><w:pPr>${tabs(["right", 9360])}</w:pPr></w:style>`;
    const container = await render([
      paragraph(`<w:pStyle w:val="TabStyle"/>${tabs(["left", 1440], ["clear", 4680])}`, "a", "\t", "First", "\t", "Second"),
    ], { styles });
    docx.remeasureTabStops(container);

    expectAt(run(container, "First").getBoundingClientRect().left, container, 72, "left");
    expectAt(run(container, "Second").getBoundingClientRect().right, container, 468, "right");
  });

  it("clamps a right stop at zero when the text is wider than the room", async () => {
    const container = await render([
      paragraph(tabs(["right", 1440]), "a", "\t", "Text much wider than one inch"),
    ]);
    docx.remeasureTabStops(container);

    expect(tabSpans(container)[0].getBoundingClientRect().width).toBe(0);
    expect(run(container, "Text much wider than one inch").getBoundingClientRect().left)
      .not.toBeLessThan(run(container, "a").getBoundingClientRect().right - 0.5);
  });

  it("aligns the first full stop on a decimal stop", async () => {
    const decimal = tabs(["decimal", 5760]);
    const container = await render([
      paragraph(decimal, "a", "\t", "12.5"),
      paragraph(decimal, "b", "\t", "1234.75"),
      paragraph(decimal, "c", "\t", "100"),
    ]);
    docx.remeasureTabStops(container);

    expectAt(dotLeft(container, "12.5"), container, 288, "12.5");
    expectAt(dotLeft(container, "1234.75"), container, 288, "1234.75");
    expectAt(run(container, "100").getBoundingClientRect().right, container, 288, "100");
  });

  it("measures inside a zoomed wrapper", async () => {
    const container = await render(separateStops);
    container.querySelector(".docx-wrapper").style.zoom = "0.5";
    docx.remeasureTabStops(container);

    expectAt(run(container, "Left").getBoundingClientRect().left, container, 144, "left", 0.5);
    expectAt(centre(run(container, "Centre").getBoundingClientRect()), container, 234, "centre", 0.5);
    expectAt(run(container, "Right").getBoundingClientRect().right, container, 468, "right", 0.5);
  });
});
