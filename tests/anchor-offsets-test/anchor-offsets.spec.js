describe("Anchored pictures", function () {
  const EMU = 12700;
  const SIZE = 21.75 * EMU;
  const picture = `<w:r><w:drawing><wp:anchor simplePos="0" allowOverlap="1" behindDoc="0" locked="0" layoutInCell="1" relativeHeight="1"><wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="page"><wp:posOffset>${45.36 * EMU}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${23.04 * EMU}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${SIZE}" cy="${SIZE}"/><wp:wrapNone/><wp:docPr id="1" name=""/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SIZE}" cy="${SIZE}"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  const inCell = picture.replace('<wp:positionH relativeFrom="page">', '<wp:positionH relativeFrom="column">').replace(`${45.36 * EMU}</wp:posOffset></wp:positionH>`, `${10 * EMU}</wp:posOffset></wp:positionH>`);
  let container;

  beforeAll(async () => {
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>' +
      `<w:p>${picture}<w:r><w:br/><w:t>Body line under the heading</w:t></w:r></w:p>` +
      '<w:tbl><w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>left cell</w:t></w:r></w:p></w:tc>' +
      `<w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p>${inCell}</w:p></w:tc></w:tr></w:tbl>` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>');
    container = document.createElement("div");
    document.body.prepend(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container);
  });

  afterAll(() => container.remove());

  const rect = element => element.getBoundingClientRect();
  const expectNear = (actual, expected) => expect(Math.abs(actual - expected)).withContext(`${actual} vs ${expected}`).toBeLessThanOrEqual(1);

  it("places a page-relative no-wrap picture in the left margin, beside the text", () => {
    const page = rect(container.querySelector("section.docx"));
    const picture = rect(container.querySelector("article img"));
    const text = rect(container.querySelector("article span"));
    expectNear(picture.left - page.left, 45.36 * 4 / 3);
    expectNear(text.left - page.left, 96);
    expect(picture.right).toBeLessThanOrEqual(text.left);
  });

  it("places a column-relative no-wrap picture in a table cell from the cell", () => {
    const cell = container.querySelectorAll("article td")[1];
    expectNear(rect(cell.querySelector("img")).left - rect(cell.querySelector("p")).left, 10 * 4 / 3);
  });
});
