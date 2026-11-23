describe("Shape groups", function () {
  const SHAPE_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
  const GROUP_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const NS = [
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    `xmlns:r="${REL}"`,
    'xmlns:v="urn:schemas-microsoft-com:vml"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
    `xmlns:wps="${SHAPE_NS}"`,
    `xmlns:wpg="${GROUP_NS}"`,
  ].join(" ");
  const PT = 12700;
  const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

  const para = (id, ...runs) => `<w:p><w:r><w:t xml:space="preserve">${id} </w:t></w:r>${runs.join("")}</w:p>`;
  const run = inner => `<w:r>${inner}</w:r>`;

  const xfrm = ({ x = 0, y = 0, w, h, ch = null, attrs = "" }) =>
    `<a:xfrm${attrs}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/>` +
    (ch ? `<a:chOff x="${ch[0]}" y="${ch[1]}"/><a:chExt cx="${ch[2]}" cy="${ch[3]}"/>` : "") + "</a:xfrm>";
  const identity = (w, h, attrs = "") => xfrm({ w: w * PT, h: h * PT, ch: [0, 0, w * PT, h * PT], attrs });
  const solid = hex => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
  const shape = ({ geom = "rect", frame, spPr = "", text = null }) =>
    `<wps:wsp><wps:cNvSpPr/><wps:spPr>${frame}<a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${spPr}</wps:spPr>` +
    `${text == null ? "" : `<wps:txbx><w:txbxContent>${text}</w:txbxContent></wps:txbx>`}<wps:bodyPr/></wps:wsp>`;
  const box = (hex, x, y, w, h) => shape({ frame: xfrm({ x: x * PT, y: y * PT, w: w * PT, h: h * PT }), spPr: solid(hex) });
  const picture = frame => '<pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="p"/><pic:cNvPicPr/></pic:nvPicPr>' +
    `<pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr>${frame}<a:prstGeom prst="rect"/></pic:spPr></pic:pic>`;

  const wgp = (frame, ...children) => `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>${frame}</wpg:grpSpPr>${children.join("")}</wpg:wgp>`;
  const grpSp = (frame, ...children) => `<wpg:grpSp><wpg:cNvGrpSpPr/><wpg:grpSpPr>${frame}</wpg:grpSpPr>${children.join("")}</wpg:grpSp>`;
  const groupGraphic = group => `<a:graphic><a:graphicData uri="${GROUP_NS}">${group}</a:graphicData></a:graphic>`;
  const inline = (w, h, content) =>
    `<w:drawing><wp:inline><wp:extent cx="${w * PT}" cy="${h * PT}"/><wp:docPr id="1" name="g"/>${content}</wp:inline></w:drawing>`;
  const anchor = ({ h, v, w, ht }, content) =>
    `<w:drawing><wp:anchor behindDoc="0" simplePos="0" relativeHeight="1" locked="0" layoutInCell="1" allowOverlap="1" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:simplePos x="0" y="0"/>${h}${v}<wp:extent cx="${w * PT}" cy="${ht * PT}"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="g"/>${content}</wp:anchor></w:drawing>`;
  const posH = (relative, inner) => `<wp:positionH relativeFrom="${relative}">${inner}</wp:positionH>`;
  const posV = (relative, inner) => `<wp:positionV relativeFrom="${relative}">${inner}</wp:positionV>`;
  const offset = emu => `<wp:posOffset>${emu}</wp:posOffset>`;
  const pict = inner => `<w:pict>${inner}</w:pict>`;

  const deepDml = level => level > 20 ? "" : (level == 1 ? wgp : grpSp)(identity(100, 50), box("00AA00", 0, 0, 10, 5), deepDml(level + 1));
  const deepVml = level => level > 20 ? "" :
    `<v:group style="${level == 1 ? "width:100pt;height:50pt" : "width:100;height:50"}" coordsize="100,50">` +
    `<v:rect style="width:10;height:5" fillcolor="#00aa00"/>${deepVml(level + 1)}</v:group>`;
  const throughTextBoxes = level => groupGraphic(level == 17
    ? wgp(identity(100, 50), box("00AA00", 0, 0, 10, 5))
    : wgp(identity(100, 50), shape({
      frame: xfrm({ w: 100 * PT, h: 50 * PT }),
      text: `<w:p><w:r>${inline(100, 50, throughTextBoxes(level + 2))}</w:r></w:p>`,
    })));

  const body = [
    para("case-nested", run(inline(200, 100, groupGraphic(wgp(identity(200, 100),
      grpSp(xfrm({ x: 20 * PT, y: 10 * PT, w: 100 * PT, h: 50 * PT, ch: [1000, 2000, 200 * PT, 100 * PT] }),
        shape({ frame: xfrm({ x: 1000 + 100 * PT, y: 2000 + 50 * PT, w: 20 * PT, h: 10 * PT }), spPr: solid("00FF00") }))))))),
    para("case-vml-left", run(pict('<v:group style="width:200pt;height:100pt" coordsize="200,100">' +
      '<v:rect style="top:10;width:50;height:20" fillcolor="#00ff00"/></v:group>'))),
    para("case-arrow", run(inline(100, 20, groupGraphic(wgp(identity(100, 20), shape({
      geom: "straightConnector1",
      frame: xfrm({ x: 10 * PT, y: 10 * PT, w: 80 * PT, h: 0 }),
      spPr: `<a:ln w="${2 * PT}">${solid("FF0000")}<a:tailEnd type="triangle"/></a:ln>`,
    })))))),
    para("case-vml-arrow", run(pict('<v:line from="0,0" to="100pt,0" strokeweight="2pt"><v:stroke endarrow="block"/></v:line>'))),
    para("case-hidden", run(inline(200, 100, groupGraphic(wgp(identity(200, 100),
      box("123456", 0, 0, 50, 50),
      box("654321", 100, 0, 50, 50).replace("<wps:wsp>", '<wps:wsp><wps:cNvPr id="7" name="h" hidden="1"/>'),
      picture(xfrm({ x: 0, y: 60 * PT, w: 40 * PT, h: 20 * PT })).replace('<pic:cNvPr id="2" name="p"/>', '<pic:cNvPr id="2" name="p" hidden="1"/>')))))),
    para("case-flip-picture", run(inline(200, 100, groupGraphic(wgp(identity(200, 100, ' flipH="1"'),
      picture(xfrm({ x: 0, y: 0, w: 40 * PT, h: 20 * PT }))))))),
    para("case-flip-cropped", run(inline(200, 100, groupGraphic(wgp(identity(200, 100, ' flipH="1"'),
      picture(xfrm({ x: 0, y: 0, w: 40 * PT, h: 20 * PT })).replace('<a:blip r:embed="rIdImage"/>', '<a:blip r:embed="rIdImage"/><a:srcRect l="10000" r="10000"/>')))))),
    para("case-picture", run(inline(200, 100, groupGraphic(wgp(identity(200, 100),
      picture(xfrm({ x: 100 * PT, y: 50 * PT, w: 40 * PT, h: 20 * PT }))))))),
    para("case-flip", run(inline(200, 100, groupGraphic(wgp(identity(200, 100, ' flipH="1"'), shape({
      frame: xfrm({ w: 50 * PT, h: 100 * PT }),
      spPr: solid("00FF00"),
      text: "<w:p><w:r><w:t>Not mirrored</w:t></w:r></w:p>",
    })))))),
    para("case-rotated", run(inline(200, 100, groupGraphic(wgp(identity(200, 100, ' rot="5400000"'), box("00AA00", 0, 0, 10, 5)))))),
    para("case-centre-before"),
    `<w:p>${run(anchor({ h: posH("page", "<wp:align>center</wp:align>"), v: posV("paragraph", offset(292608)), w: 200, ht: 40 },
      groupGraphic(wgp(identity(200, 40), box("AA0000", 0, 0, 200, 40)))))}</w:p>`,
    `<w:p>${run(anchor({ h: posH("column", offset(36 * PT)), v: posV("paragraph", offset(45720)), w: 100, ht: 40 },
      groupGraphic(wgp(identity(100, 40), box("0000AA", 0, 0, 100, 40)))))}</w:p>`,
    para("case-fallback", run(`<mc:AlternateContent><mc:Choice Requires="wpg">${inline(100, 50, groupGraphic(wgp(identity(100, 50),
      box("FF0000", 0, 0, 50, 20), '<wps:wsp><wps:spPr><a:custGeom><a:pathLst/></a:custGeom></wps:spPr><wps:bodyPr/></wps:wsp>')))}</mc:Choice>` +
      `<mc:Fallback>${pict('<v:group style="width:100pt;height:50pt" coordsize="100,50"><v:rect style="width:50;height:20" fillcolor="#0000ff"/></v:group>')}</mc:Fallback></mc:AlternateContent>`)),
    para("case-bad-chext", run(inline(100, 50, groupGraphic(wgp(xfrm({ w: 100 * PT, h: 50 * PT, ch: [0, 0, 100 * PT, 0] }), box("00FF00", 0, 0, 10, 5)))))),
    para("case-bad-xfrm", run(inline(100, 50, groupGraphic(wgp(identity(100, 50), shape({ frame: "", spPr: solid("00FF00") })))))),
    para("case-bad-coordsize", run(pict('<v:group style="width:100pt;height:50pt" coordsize="a,b"><v:rect style="width:10;height:5" fillcolor="#00ff00"/></v:group>'))),
    para("case-depth-dml", run(inline(100, 50, groupGraphic(deepDml(1))))),
    para("case-depth-vml", run(pict(deepVml(1)))),
    para("case-depth-text", run(inline(100, 50, throughTextBoxes(1)))),
    para("case-bent", run(inline(100, 50, groupGraphic(wgp(identity(100, 50),
      box("00FF00", 0, 0, 10, 5), shape({ geom: "bentConnector3", frame: xfrm({ w: 50 * PT, h: 20 * PT }), spPr: `<a:ln>${solid("000000")}</a:ln>` })))))),
  ].join("\n");

  async function buildDocx() {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>
</Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdImage" Type="${REL}/image" Target="media/image1.png"/>
</Relationships>`);
    zip.file("word/media/image1.png", PNG, { base64: true });
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>
${body}
<w:sectPr>
<w:pgSz w:w="6000" w:h="6000"/>
<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>
</w:sectPr>
</w:body></w:document>`);
    return zip.generateAsync({ type: "blob" });
  }

  let container;

  beforeAll(async () => {
    container = document.createElement("div");
    document.body.prepend(container);
    await docx.renderAsync(await buildDocx(), container, null, { className: "groups" });
  });

  afterAll(() => container.remove());

  const MARGIN = 48;
  const section = () => container.querySelector("section.groups");
  const paragraph = id => Array.from(container.querySelectorAll("article p")).find(p => p.textContent.startsWith(`${id} `));
  const filled = (root, hex) => Array.from(root.querySelectorAll("svg rect")).filter(r => r.getAttribute("fill") == hex);
  const boxOf = (root, hex) => filled(root, hex)[0].closest("svg").parentElement;
  const dmlGroup = id => paragraph(id).querySelector("div").firstElementChild;
  const vmlGroup = id => paragraph(id).querySelector("div");
  const rect = element => element.getBoundingClientRect();
  const expectNear = (actual, expected, tolerance = 1) => expect(Math.abs(actual - expected)).withContext(`${actual} vs ${expected}`).toBeLessThanOrEqual(tolerance);
  const reasons = root => Array.from(root.querySelectorAll("[data-docx-not-drawn]"), el => el.getAttribute("data-docx-not-drawn"));

  it("maps a child of a scaled nested group through both frames", () => {
    const child = rect(boxOf(paragraph("case-nested"), "#00ff00"));
    const group = rect(dmlGroup("case-nested"));
    expectNear((child.left + child.right) / 2 - group.left, 100);
    expectNear((child.top + child.bottom) / 2 - group.top, 50);
  });

  it("puts a VML group child without left at the group's left edge", () => {
    const child = rect(boxOf(paragraph("case-vml-left"), "#00ff00"));
    const group = rect(vmlGroup("case-vml-left"));
    expectNear(child.left, group.left);
    expectNear(child.top - group.top, 40 / 3);
  });

  it("draws a zero-height connector with a filled arrowhead at its tail", () => {
    const line = paragraph("case-arrow").querySelector("svg line");
    const box = line.closest("svg").parentElement;
    const head = box.querySelector("svg path");
    const bounds = head.getBBox();
    expect(rect(box).width).toBeGreaterThan(0);
    expect(head.getAttribute("fill")).toBe("#ff0000");
    expectNear(bounds.width, 8);
    expectNear(rect(head).right, rect(box).right);
    expectNear(rect(line).right, rect(head).left);
    expect(reasons(paragraph("case-arrow"))).toEqual([]);
  });

  it("draws an arrowhead on a standalone VML line", () => {
    const heads = paragraph("case-vml-arrow").querySelectorAll("svg path");
    expect(heads.length).toBe(1);
    expect(heads[0].getAttribute("fill")).toBe("black");
  });

  it("skips group children the author hid", () => {
    const group = paragraph("case-hidden");
    expect(filled(group, "#123456").length).toBe(1);
    expect(filled(group, "#654321").length).toBe(0);
    expect(group.querySelector("img")).toBeNull();
  });

  it("mirrors a picture in a flipped group", () => {
    expect(new DOMMatrix(getComputedStyle(paragraph("case-flip-picture").querySelector("img")).transform).a).toBe(-1);
    expect(new DOMMatrix(getComputedStyle(paragraph("case-flip-cropped").querySelector("img")).transform).a).toBeLessThan(0);
  });

  it("fills a picture's child box with its image", () => {
    const image = rect(paragraph("case-picture").querySelector("img"));
    const group = rect(dmlGroup("case-picture"));
    expectNear(image.left - group.left, 400 / 3);
    expectNear(image.top - group.top, 200 / 3);
    expectNear(image.width, 160 / 3);
    expectNear(image.height, 80 / 3);
  });

  it("rotates a top-level group by its own rotation", () => {
    const matrix = new DOMMatrix(getComputedStyle(dmlGroup("case-rotated")).transform);
    expect(matrix.b).toBeCloseTo(1, 5);
    expect(matrix.a).toBeCloseTo(0, 5);
  });

  it("mirrors a flipped group's child positions but not its text", () => {
    const box = boxOf(paragraph("case-flip"), "#00ff00");
    const group = dmlGroup("case-flip");
    const text = Array.from(box.querySelectorAll("p")).find(p => p.textContent == "Not mirrored");
    expectNear(rect(box).left - rect(group).left, 200);
    expect(getComputedStyle(box.querySelector("svg")).transform).toBe("matrix(-1, 0, 0, 1, 0, 0)");
    for (let el = text; el != group; el = el.parentElement)
      expect(getComputedStyle(el).transform).toBe("none");
  });

  it("centres a page-aligned square-wrapped group and offsets it below its paragraph", () => {
    const page = rect(section());
    const group = rect(boxOf(container, "#aa0000").parentElement);
    expectNear((group.left + group.right) / 2, page.left + page.width / 2);
    expectNear(group.top - rect(paragraph("case-centre-before")).bottom, 0.32 * 96);
  });

  it("offsets a column-relative group from the text column", () => {
    const group = rect(boxOf(container, "#0000aa").parentElement);
    expectNear(group.left - rect(section()).left, MARGIN + 48);
  });

  it("draws the VML fallback when a group child cannot be drawn", () => {
    expect(filled(paragraph("case-fallback"), "#0000ff").length).toBe(1);
    expect(filled(paragraph("case-fallback"), "#ff0000").length).toBe(0);
    expect(reasons(paragraph("case-fallback"))).toEqual([]);
  });

  it("marks groups and children with unusable sizes without writing NaN", () => {
    expect(reasons(paragraph("case-bad-chext"))).toEqual(["group-child"]);
    expect(reasons(paragraph("case-bad-xfrm"))).toEqual(["group-child"]);
    expect(reasons(paragraph("case-bad-coordsize"))).toEqual(["group-child"]);
    for (const id of ["case-bad-chext", "case-bad-xfrm", "case-bad-coordsize"])
      expect(paragraph(id).innerHTML).not.toContain("NaN");
  });

  it("draws groups nested 16 deep and marks the 17th on both walks", () => {
    for (const id of ["case-depth-dml", "case-depth-vml"]) {
      expect(filled(paragraph(id), "#00aa00").length).withContext(id).toBe(16);
      expect(reasons(paragraph(id))).withContext(id).toEqual(["depth"]);
    }
  });

  it("counts text boxes and groups on one depth counter", () => {
    expect(reasons(paragraph("case-depth-text"))).toEqual(["depth"]);
  });

  it("marks a connector kind it does not draw and keeps its siblings", () => {
    expect(reasons(paragraph("case-bent"))).toEqual(["group-child"]);
    expect(filled(paragraph("case-bent"), "#00ff00").length).toBe(1);
  });
});
