describe("shape groups", function () {
  const PT = 12700;
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';

  const xfrm = (w, h, attrs = "") => `<a:xfrm${attrs}><a:off x="0" y="0"/><a:ext cx="${w * PT}" cy="${h * PT}"/><a:chOff x="0" y="0"/><a:chExt cx="${w * PT}" cy="${h * PT}"/></a:xfrm>`;
  const box = (hex, w, h, extra = "") => `<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * PT}" cy="${h * PT}"/></a:xfrm><a:prstGeom prst="rect"/>` +
    `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></wps:spPr>${extra}<wps:bodyPr/></wps:wsp>`;
  const group = (w, h, attrs, ...children) => `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">` +
    `<wpg:wgp><wpg:grpSpPr>${xfrm(w, h, attrs)}</wpg:grpSpPr>${children.join("")}</wpg:wgp></a:graphicData></a:graphic>`;
  const inline = (w, h, content) => `<w:drawing><wp:inline><wp:extent cx="${w * PT}" cy="${h * PT}"/><wp:docPr id="1" name="g"/>${content}</wp:inline></w:drawing>`;
  const connector = x => `<wps:wsp><wps:spPr><a:xfrm><a:off x="${x * PT}" y="0"/><a:ext cx="${40 * PT}" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"/>` +
    `<a:ln w="${PT}"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>`;
  const para = (id, run) => `<w:p><w:r><w:t xml:space="preserve">${id} </w:t></w:r><w:r>${run}</w:r></w:p>`;

  const body = [
    para("flip", inline(200, 100, group(200, 100, ' flipH="1"', box("00FF00", 50, 100, "<wps:txbx><w:txbxContent><w:p><w:r><w:t>Not mirrored</w:t></w:r></w:p></w:txbxContent></wps:txbx>")))),
    para("fallback", `<mc:AlternateContent><mc:Choice Requires="wpg">${inline(100, 50, group(100, 50, "", box("FF0000", 50, 20),
      '<wps:wsp><wps:spPr><a:custGeom><a:pathLst/></a:custGeom></wps:spPr><wps:bodyPr/></wps:wsp>'))}</mc:Choice>` +
      '<mc:Fallback><w:pict><v:group style="width:100pt;height:50pt" coordsize="100,50"><v:rect style="width:50;height:20" fillcolor="#0000ff"/></v:group></w:pict></mc:Fallback></mc:AlternateContent>'),
    para("flat", `<mc:AlternateContent><mc:Choice Requires="wpg">${inline(100, 0, group(100, 0, "", connector(0), connector(60)))}</mc:Choice>` +
      '<mc:Fallback><w:pict><v:oval style="width:20pt;height:10pt"/></w:pict></mc:Fallback></mc:AlternateContent>'),
  ].join("");

  let container;

  beforeAll(async () => {
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/document.xml", `<w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>');
    container = document.createElement("div");
    document.body.prepend(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container);
  });

  afterAll(() => container.remove());

  const paragraph = id => Array.from(container.querySelectorAll("article p")).find(p => p.textContent.startsWith(`${id} `));
  const filled = (id, hex) => Array.from(paragraph(id).querySelectorAll("svg rect")).filter(r => r.getAttribute("fill") == hex);
  const rect = element => element.getBoundingClientRect();
  const expectNear = (actual, expected) => expect(Math.abs(actual - expected)).withContext(`${actual} vs ${expected}`).toBeLessThanOrEqual(1);

  it("mirrors a flipped group's child positions but not its text", () => {
    const box = filled("flip", "#00ff00")[0].closest("svg").parentElement;
    const group = paragraph("flip").querySelector("div").firstElementChild;
    const text = Array.from(box.querySelectorAll("p")).find(p => p.textContent == "Not mirrored");
    expectNear(rect(box).left - rect(group).left, 200);
    expect(getComputedStyle(box.querySelector("svg")).transform).toBe("matrix(-1, 0, 0, 1, 0, 0)");
    for (let el = text; el != group; el = el.parentElement)
      expect(getComputedStyle(el).transform).toBe("none");
  });

  it("draws the VML fallback when a group child cannot be drawn", () => {
    expect(filled("fallback", "#0000ff").length).toBe(1);
    expect(filled("fallback", "#ff0000").length).toBe(0);
  });

  it("draws a group with no height", () => {
    expect(paragraph("flat").querySelectorAll("svg line").length).toBe(2);
    expect(paragraph("flat").querySelector("[data-docx-not-drawn]")).toBeNull();
  });
});
