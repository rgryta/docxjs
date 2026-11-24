describe("DrawingML shapes", function () {
  const PT = 12700;
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const wsp = (w, h, extra = "") => `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * PT}" cy="${h * PT}"/></a:xfrm><a:prstGeom prst="rect"/>${extra}</wps:wsp></a:graphicData></a:graphic>`;
  const anchor = (h, v, wrap = "<wp:wrapNone/>", w = 20, ht = 10) =>
    `<w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" relativeHeight="1" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>${h}${v}<wp:extent cx="${w * PT}" cy="${ht * PT}"/>${wrap}<wp:docPr id="1" name="s"/>${wsp(w, ht, "</wps:spPr><wps:bodyPr/>")}</wp:anchor></w:drawing></w:r>`;
  const pos = (axis, relative, inner) => `<wp:position${axis} relativeFrom="${relative}">${inner}</wp:position${axis}>`;
  const para = (id, ...runs) => `<w:p><w:r><w:t xml:space="preserve">${id} </w:t></w:r>${runs.join("")}</w:p>`;
  const top = pos("V", "paragraph", "<wp:posOffset>0</wp:posOffset>");
  const textBox = (text, bodyPr, shapeType = "") => `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="${100 * PT}" cy="${20 * PT}"/><wp:docPr id="1" name="s"/>` +
    `${wsp(100, 20, `</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>${bodyPr}`)}</wp:inline></w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict>${shapeType}<v:shape type="#_x0000_t202" style="width:100pt;height:20pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

  const body = [
    para("page", anchor(pos("H", "page", `<wp:posOffset>${72 * PT}</wp:posOffset>`), pos("V", "margin", "<wp:align>top</wp:align>"))),
    para("right", anchor(pos("H", "margin", "<wp:align>right</wp:align>"), top)),
    para("centre", anchor(pos("H", "margin", "<wp:align>center</wp:align>"), top)),
    para("float", anchor(pos("H", "column", `<wp:posOffset>${150 * PT}</wp:posOffset>`), top, '<wp:wrapSquare wrapText="bothSides"/>', 40, 30),
      `<w:r><w:t>${"words that wrap beside the shape ".repeat(12)}</w:t></w:r>`),
    para("colour", `<w:r><w:drawing><wp:inline><wp:extent cx="${100 * PT}" cy="${20 * PT}"/><wp:docPr id="1" name="s"/>${wsp(100, 20,
      '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="85000"/></a:schemeClr></a:solidFill></wps:spPr>' +
      '<wps:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>' +
      '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>' +
      '<wps:txbx><w:txbxContent><w:p><w:r><w:t>White text</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/>')}</wp:inline></w:drawing></w:r>`),
    para("first box", textBox("First box", "<wps:bodyPr/>", '<v:shapetype id="_x0000_t202" coordsize="21600,21600" path="m,l,21600r21600,l21600,xe"/>')),
    para("second box", textBox("Second box", '<wps:bodyPr vert="vert270"/>')),
    '<w:tbl><w:tblGrid><w:gridCol w:w="2280"/><w:gridCol w:w="2280"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2280" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>left cell</w:t></w:r></w:p></w:tc>' +
      `<w:tc><w:tcPr><w:tcW w:w="2280" w:type="dxa"/></w:tcPr><w:p>${anchor(pos("H", "column", "<wp:posOffset>0</wp:posOffset>"), top)}</w:p></w:tc></w:tr></w:tbl>`,
  ].join("");

  let container;

  beforeAll(async () => {
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/theme" Target="theme1.xml"/></Relationships>`);
    zip.file("word/theme1.xml", `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="c">` +
      '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme></a:themeElements></a:theme>');
    zip.file("word/document.xml", `<w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="6000" w:h="6000"/>` +
      '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:body></w:document>');
    container = document.createElement("div");
    document.body.prepend(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container);
  });

  afterAll(() => container.remove());

  const MARGIN = 48;
  const page = () => container.querySelector("section.docx").getBoundingClientRect();
  const paragraph = id => Array.from(container.querySelectorAll("article p")).find(p => p.textContent.startsWith(`${id} `));
  const box = id => paragraph(id).querySelector("svg").parentElement.getBoundingClientRect();
  const expectNear = (actual, expected) => expect(Math.abs(actual - expected)).withContext(`${actual} vs ${expected}`).toBeLessThanOrEqual(1);

  it("places a page-offset, margin-top shape against the section", () => {
    expectNear(box("page").left - page().left, 96);
    expectNear(box("page").top - page().top, MARGIN);
  });

  it("aligns a margin-relative shape right and centre in the text column", () => {
    expectNear(box("right").right, page().right - MARGIN);
    expectNear((box("centre").left + box("centre").right) / 2, page().left + page().width / 2);
  });

  it("floats a square-wrapped shape so the text wraps beside it", () => {
    const wrapper = paragraph("float").querySelector("div");
    expect(getComputedStyle(wrapper).cssFloat).toBe("right");
    const shape = wrapper.getBoundingClientRect();
    expectNear(shape.left - page().left, MARGIN + 150 * 4 / 3);
    const range = document.createRange();
    range.selectNodeContents(paragraph("float").lastChild);
    const beside = Array.from(range.getClientRects()).filter(r => r.top < shape.bottom - 1);
    expect(beside.length).toBeGreaterThan(0);
    beside.forEach(r => expect(r.right).toBeLessThanOrEqual(shape.left + 1));
  });

  it("keeps the shape type of a Fallback whose Choice is drawn for a later Fallback", () => {
    expect(paragraph("second box").querySelector("[data-docx-not-drawn]")).toBeNull();
    expect(Array.from(paragraph("second box").querySelectorAll("p")).some(p => p.textContent == "Second box")).toBeTrue();
  });

  it("places a no-wrap column-relative shape in a table cell from the cell", () => {
    const cell = container.querySelectorAll("article td")[1];
    expectNear(cell.querySelector("svg").parentElement.getBoundingClientRect().left, cell.querySelector("p").getBoundingClientRect().left);
  });

  it("resolves theme colours with modifiers and the font reference colour", () => {
    expect(paragraph("colour").querySelector("svg rect").getAttribute("fill")).toBe("#3660ab");
    expect(getComputedStyle(Array.from(paragraph("colour").querySelectorAll("p")).find(p => p.textContent == "White text")).color).toBe("rgb(255, 255, 255)");
  });
});
