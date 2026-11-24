describe("VML shape placement", function () {
  const shape = (id, style) => `<w:p><w:r><w:t xml:space="preserve">${id} </w:t></w:r><w:r><w:pict><v:rect style="width:20pt;height:10pt;${style}"/></w:pict></w:r></w:p>`;
  let container;

  beforeAll(async () => {
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/document.xml", `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>
${shape("margin", "position:absolute;margin-left:24pt;mso-position-horizontal-relative:margin;mso-position-vertical-relative:text")}
${shape("centre", "position:absolute;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page")}
<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>left cell</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:pict><v:rect style="width:100pt;height:50pt;position:absolute;margin-left:0;mso-position-horizontal-relative:text"/></w:pict></w:r></w:p></w:tc></w:tr></w:tbl>
<w:sectPr><w:pgSz w:w="12240" w:h="4000"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`);
    container = document.createElement("div");
    document.body.appendChild(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container);
  });

  afterAll(() => container.remove());

  const page = () => container.querySelector("section.docx").getBoundingClientRect();
  const paragraph = id => Array.from(container.querySelectorAll("article p")).find(p => p.textContent.startsWith(`${id} `));
  const box = id => paragraph(id).querySelector("svg").parentElement.getBoundingClientRect();
  const expectNear = (actual, expected) => expect(Math.abs(actual - expected)).withContext(`${actual} vs ${expected}`).toBeLessThanOrEqual(1);

  it("places a margin-relative shape from the margin, not from the text before it", () => {
    expectNear(box("margin").left - page().left, 96 + 32);
  });

  it("centres a shape on the page on both axes", () => {
    expectNear(box("centre").left + box("centre").width / 2 - page().left, page().width / 2);
    expectNear(box("centre").top + box("centre").height / 2 - page().top, 100 / 0.75);
  });

  it("places a text-relative shape in a table cell from the cell", () => {
    const cell = container.querySelectorAll("article td")[1];
    expectNear(cell.querySelector("svg").parentElement.getBoundingClientRect().left, cell.querySelector("p").getBoundingClientRect().left);
  });
});
