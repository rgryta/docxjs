describe("tab stops", function () {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PX_PER_PT = 96 / 72;

  async function render(pPr, ...texts) {
    const runs = texts.map(t => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`).join("<w:r><w:tab/></w:r>");
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/document.xml", `<w:document ${W}><w:body><w:p><w:pPr>${pPr}</w:pPr>${runs}</w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);

    container = document.createElement("div");
    document.body.appendChild(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container, null, { experimental: true });
    docx.remeasureTabStops(container);
  }

  let container;
  afterEach(() => container.remove());

  // right edge of a run, in points from the left edge of the page's text column
  function rightOf(text) {
    const span = Array.from(container.querySelectorAll("article span")).find(s => s.textContent === text);
    return (span.getBoundingClientRect().right - container.querySelector("article").getBoundingClientRect().left) / PX_PER_PT;
  }

  it("counts stops from the text column in an indented paragraph", async () => {
    await render('<w:ind w:left="1440"/><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs>', "a", "Indented");

    expect(rightOf("Indented")).toBeCloseTo(468, 0);
  });

  it("gives the same widths on every call and follows a changed layout", async () => {
    await render('<w:tabs><w:tab w:val="left" w:pos="2880"/><w:tab w:val="right" w:pos="9360"/></w:tabs>', "a", "One", "Last");
    const widths = () => Array.from(container.querySelectorAll("span.docx-tab-stop")).map(s => s.getBoundingClientRect().width);
    const first = widths();
    docx.remeasureTabStops(container);
    expect(widths()).toEqual(first);

    Array.from(container.querySelectorAll("article span")).find(s => s.textContent === "Last").style.fontSize = "24pt";
    docx.remeasureTabStops(container);

    expect(rightOf("Last")).toBeCloseTo(468, 0);
  });
});
