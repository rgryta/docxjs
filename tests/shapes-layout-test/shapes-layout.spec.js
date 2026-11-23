describe("DrawingML shapes", function () {
  const SHAPE_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
  const NS = [
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    'xmlns:v="urn:schemas-microsoft-com:vml"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
    `xmlns:wps="${SHAPE_NS}"`,
  ].join(" ");
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PT = 12700;

  const para = (id, ...runs) => `<w:p><w:r><w:t xml:space="preserve">${id} </w:t></w:r>${runs.join("")}</w:p>`;
  const run = inner => `<w:r>${inner}</w:r>`;
  const textParas = (...texts) => texts.map(t => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`).join("");

  const wsp = ({ geom = "rect", w = 20, h = 10, adj = "", xfrm = "", spPr = "", style = "", text = null, bodyPr = "<wps:bodyPr/>" } = {}) =>
    `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm${xfrm}><a:off x="0" y="0"/><a:ext cx="${w * PT}" cy="${h * PT}"/></a:xfrm>` +
    `<a:prstGeom prst="${geom}"><a:avLst>${adj}</a:avLst></a:prstGeom>${spPr}</wps:spPr>${style}` +
    `${text == null ? "" : `<wps:txbx><w:txbxContent>${text}</w:txbxContent></wps:txbx>`}${bodyPr}</wps:wsp>`;
  const custGeom = '<wps:wsp><wps:spPr><a:custGeom><a:pathLst/></a:custGeom></wps:spPr><wps:bodyPr/></wps:wsp>';

  const graphic = (uri, data) => `<a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic>`;
  const shapeGraphic = shape => graphic(SHAPE_NS, shape);
  const picture = graphic("http://schemas.openxmlformats.org/drawingml/2006/picture",
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill>' +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${20 * PT}" cy="${10 * PT}"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic>`);

  const inline = (w, h, content) =>
    `<w:drawing><wp:inline><wp:extent cx="${w * PT}" cy="${h * PT}"/><wp:docPr id="1" name="s"/>${content}</wp:inline></w:drawing>`;
  const anchor = ({ h, v, w = 20, ht = 10, wrap = "<wp:wrapNone/>", behind = 0 }, content) =>
    `<w:drawing><wp:anchor behindDoc="${behind}" simplePos="0" relativeHeight="1" locked="0" layoutInCell="1" allowOverlap="1" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:simplePos x="0" y="0"/>${h}${v}<wp:extent cx="${w * PT}" cy="${ht * PT}"/>${wrap}<wp:docPr id="1" name="s"/>${content}</wp:anchor></w:drawing>`;
  const posH = (relative, inner) => `<wp:positionH relativeFrom="${relative}">${inner}</wp:positionH>`;
  const posV = (relative, inner) => `<wp:positionV relativeFrom="${relative}">${inner}</wp:positionV>`;
  const offset = points => `<wp:posOffset>${points * PT}</wp:posOffset>`;
  const align = value => `<wp:align>${value}</wp:align>`;

  const solid = color => `<a:solidFill>${color}</a:solidFill>`;
  const alternate = (choice, fallback) => `<mc:AlternateContent>${choice}${fallback == null ? "" : `<mc:Fallback>${fallback}</mc:Fallback>`}</mc:AlternateContent>`;
  const ovalFallback = '<w:pict><v:oval style="width:20pt;height:10pt"/></w:pict>';
  const rectChoice = (requires, extra = "") => `<mc:Choice Requires="${requires}"${extra}>${inline(20, 10, shapeGraphic(wsp()))}</mc:Choice>`;
  const nestedBoxes = depth => depth == 0 ? textParas("innermost")
    : `<w:p><w:r>${inline(200, 20, shapeGraphic(wsp({ w: 200, h: 20, text: nestedBoxes(depth - 1) })))}</w:r></w:p>`;

  const body = [
    para("case-behind behind behind behind behind behind",
      run(anchor({ h: posH("margin", offset(0)), v: posV("paragraph", offset(0)), w: 150, ht: 8, behind: 1 },
        shapeGraphic(wsp({ w: 150, h: 8, spPr: solid('<a:srgbClr val="FFFF00"/>') }))))),
    para("case-text", run(inline(150, 30, shapeGraphic(wsp({ w: 150, h: 30, text: textParas("Layer text") }))))),
    para("case-radius-dml", run(alternate(
      `<mc:Choice Requires="wps">${inline(120, 30, shapeGraphic(wsp({ geom: "roundRect", w: 120, h: 30, adj: '<a:gd name="adj" fmla="val 49143"/>' })))}</mc:Choice>`,
      '<w:pict><v:rect style="width:120pt;height:30pt"/></w:pict>'))),
    para("case-radius-vml", run('<w:pict><v:roundrect arcsize="32206f" style="width:120pt;height:30pt"/></w:pict>')),
    para("case-colour", run(inline(100, 20, shapeGraphic(wsp({
      w: 100, h: 20,
      spPr: solid('<a:schemeClr val="accent1"><a:lumMod val="85000"/></a:schemeClr>'),
      style: '<wps:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>' +
        '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>',
      text: textParas("White text"),
    }))))),
    para("case-bad-colour", run(inline(40, 20, shapeGraphic(wsp({
      w: 40, h: 20, spPr: solid('<a:srgbClr val="url(x)"/>') + `<a:ln w="${PT}">${solid('<a:srgbClr val="FF0000"/>')}</a:ln>`,
    }))))),
    para("case-page", run(anchor({ h: posH("page", offset(72)), v: posV("margin", align("top")) }, shapeGraphic(wsp())))),
    para("case-right", run(anchor({ h: posH("margin", align("right")), v: posV("paragraph", offset(0)) }, shapeGraphic(wsp())))),
    para("case-center", run(anchor({ h: posH("margin", align("center")), v: posV("paragraph", offset(0)) }, shapeGraphic(wsp())))),
    para("case-v-para", run(anchor({ h: posH("column", offset(0)), v: posV("paragraph", offset(12)) }, shapeGraphic(wsp())))),
    para("case-flat-line", run(inline(100, 0, shapeGraphic(wsp({ geom: "line", w: 100, h: 0 }))))),
    para("case-flip", run(inline(100, 20, shapeGraphic(wsp({ w: 100, h: 20, xfrm: ' flipH="1"', text: textParas("Flipped") }))))),
    para("case-autofit", run(inline(100, 10, shapeGraphic(wsp({
      w: 100, h: 10, text: textParas("one", "two", "three"), bodyPr: "<wps:bodyPr><a:spAutoFit/></wps:bodyPr>",
    }))))),
    para("case-pic-none", run(anchor({ h: posH("page", offset(72)), v: posV("paragraph", offset(0)) }, picture))),
    para("case-pic-square", run(anchor({ h: posH("column", offset(72)), v: posV("paragraph", offset(0)), wrap: '<wp:wrapSquare wrapText="bothSides"/>' }, picture))),
    para("case-float-right", run(anchor({ h: posH("column", offset(150)), v: posV("paragraph", offset(0)), w: 40, ht: 30, wrap: '<wp:wrapSquare wrapText="bothSides"/>' }, shapeGraphic(wsp({ w: 40, h: 30 })))),
      `<w:r><w:t xml:space="preserve">${"words that wrap beside the shape ".repeat(12)}</w:t></w:r>`),
    para("case-float-left", run(anchor({ h: posH("margin", align("left")), v: posV("paragraph", offset(0)), w: 40, ht: 30, wrap: '<wp:wrapSquare wrapText="bothSides"/>' }, shapeGraphic(wsp({ w: 40, h: 30 })))),
      `<w:r><w:t xml:space="preserve">${"words that wrap beside the shape ".repeat(12)}</w:t></w:r>`),
    para("case-rtl", run(anchor({ h: posH("margin", offset(50)), v: posV("paragraph", offset(0)), wrap: "<wp:wrapTopAndBottom/>" }, shapeGraphic(wsp())))).replace("<w:p>", "<w:p><w:pPr><w:bidi/></w:pPr>"),
    para("case-square-page", run(anchor({ h: posH("page", offset(72)), v: posV("page", offset(200)), wrap: '<wp:wrapSquare wrapText="bothSides"/>' }, shapeGraphic(wsp())))),
    para("case-pic-behind", run(anchor({ h: posH("page", offset(72)), v: posV("paragraph", offset(0)), behind: 1 }, picture))),
    para("case-choice-empty", run(alternate('<mc:Choice Requires="wps"/>', ovalFallback))),
    para("case-choice-blank", run(alternate(rectChoice(""), ovalFallback))),
    para("case-choice-two", run(alternate(`<mc:Choice Requires="wps">${inline(20, 10, shapeGraphic(wsp()))}${inline(20, 10, shapeGraphic(wsp()))}</mc:Choice>`, ovalFallback))),
    para("case-choice-local", run(alternate(rectChoice("wq", ` xmlns:wq="${SHAPE_NS}"`), ovalFallback))),
    para("case-choice-wordart", run(alternate(`<mc:Choice Requires="wps">${inline(20, 10, shapeGraphic(wsp({
      bodyPr: '<wps:bodyPr><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp></wps:bodyPr>',
    })))}</mc:Choice>`, ovalFallback))),
    para("case-marks",
      run(inline(20, 10, shapeGraphic(custGeom))),
      run(alternate(`<mc:Choice Requires="wps">${inline(20, 10, shapeGraphic(custGeom))}</mc:Choice>`)),
      run(alternate('<mc:Choice Requires="wps"/>', `${ovalFallback}${ovalFallback}`))),
    para("case-depth-16", run(inline(200, 20, shapeGraphic(wsp({ w: 200, h: 20, text: nestedBoxes(15) }))))),
    para("case-depth-17", run(inline(200, 20, shapeGraphic(wsp({ w: 200, h: 20, text: nestedBoxes(16) }))))),
    para("case-unsafe", run(anchor({ h: posH("page", "<wp:posOffset>5px</wp:posOffset>"), v: posV("paragraph", offset(0)) }, shapeGraphic(wsp())))),
    ...Array.from({ length: 30 }, (_, i) => para(`filler-${i}`)),
  ].join("\n");

  const footer = `<w:ftr ${NS}><w:p>${run(anchor({ h: posH("page", offset(0)), v: posV("page", offset(270)), w: 50, ht: 12 }, shapeGraphic(wsp({ w: 50, h: 12 }))))}</w:p></w:ftr>`;

  const theme = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t"><a:themeElements>
<a:clrScheme name="c"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme>
<a:fontScheme name="f"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
<a:fmtScheme name="m"><a:lnStyleLst><a:ln w="6350"/><a:ln w="12700"/><a:ln w="19050"/></a:lnStyleLst></a:fmtScheme>
</a:themeElements></a:theme>`;

  async function buildDocx(content = body, columns = "") {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>
</Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdFooter" Type="${REL}/footer" Target="footer1.xml"/>
<Relationship Id="rIdTheme" Type="${REL}/theme" Target="theme/theme1.xml"/>
</Relationships>`);
    zip.file("word/theme/theme1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${theme}`);
    zip.file("word/footer1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${footer}`);
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>
${content}
<w:sectPr>${columns}
<w:footerReference w:type="default" r:id="rIdFooter"/>
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
    await docx.renderAsync(await buildDocx(), container);
  });

  afterAll(() => container.remove());

  const MARGIN = 48;
  const section = () => container.querySelector("section.docx");
  const paragraph = id => Array.from(container.querySelectorAll("article p")).find(p => p.textContent.startsWith(`${id} `));
  const shapeBox = id => paragraph(id).querySelector("svg").parentElement;
  const wrapperOf = id => paragraph(id).querySelector("div");
  const textOf = (id, text) => Array.from(paragraph(id).querySelectorAll("p")).find(p => p.textContent == text);
  const rect = element => element.getBoundingClientRect();
  const px = value => parseFloat(value) * (value.endsWith("pt") ? 4 / 3 : 1);
  const expectNear = (actual, expected, tolerance = 1) => expect(Math.abs(actual - expected)).withContext(`${actual} vs ${expected}`).toBeLessThanOrEqual(tolerance);
  const reasons = root => Array.from(root.querySelectorAll("[data-docx-not-drawn]"), el => el.getAttribute("data-docx-not-drawn"));

  it("rounds a DrawingML rectangle from its adjustment and a VML one from its arc size alike", () => {
    expectNear(paragraph("case-radius-dml").querySelector("svg rect").rx.baseVal.value, 19.66, 0.1);
    expectNear(paragraph("case-radius-vml").querySelector("svg rect").rx.baseVal.value, 19.66, 0.1);
  });

  it("resolves theme colours with modifiers and the font reference colour", () => {
    expect(paragraph("case-colour").querySelector("svg rect").getAttribute("fill")).toBe("#3660ab");
    expect(getComputedStyle(textOf("case-colour", "White text")).color).toBe("rgb(255, 255, 255)");
  });

  it("draws an unresolvable fill as none, keeps the outline and marks the shape as partly drawn", () => {
    const shape = paragraph("case-bad-colour").querySelector("svg rect");
    expect(shape.getAttribute("fill")).toBe("none");
    expect(shape.getAttribute("stroke")).toBe("#ff0000");
    expect(reasons(paragraph("case-bad-colour"))).toEqual(["partial"]);
  });

  it("places a page-offset, margin-top shape against the section", () => {
    const box = rect(shapeBox("case-page"));
    expectNear(box.left - rect(section()).left, 96);
    expectNear(box.top - rect(section()).top, MARGIN);
  });

  it("aligns a margin-relative shape right and centre in the text column", () => {
    const page = rect(section());
    expectNear(rect(shapeBox("case-right")).right, page.right - MARGIN);
    const centre = rect(shapeBox("case-center"));
    expectNear((centre.left + centre.right) / 2, page.left + page.width / 2);
  });

  it("draws a line with no height flat", () => {
    const line = paragraph("case-flat-line").querySelector("svg line");
    expect(line.getAttribute("y2")).toBe(line.getAttribute("y1"));
  });

  // the client rects of the paragraph's text that share a band with the box
  const textBeside = (id, box) => {
    const range = document.createRange();
    range.selectNodeContents(paragraph(id));
    return Array.from(range.getClientRects()).filter(r => r.width > 0 && r.top < box.bottom - 1 && r.bottom > box.top + 1 && !(r.left >= box.left - 0.5 && r.right <= box.right + 0.5));
  };

  it("floats a square-wrapped shape to the right of the column so the text wraps beside it", () => {
    const wrapper = wrapperOf("case-float-right");
    const box = rect(wrapper);
    expect(getComputedStyle(wrapper).cssFloat).toBe("right");
    expectNear(box.left - rect(wrapper.closest("article")).left, 150 * 4 / 3);
    const beside = textBeside("case-float-right", box);
    expect(beside.length).toBeGreaterThan(0);
    beside.forEach(r => expect(r.right).toBeLessThanOrEqual(box.left + 1));
  });

  it("floats a left-aligned square-wrapped shape to the left of the column", () => {
    const wrapper = wrapperOf("case-float-left");
    const box = rect(wrapper);
    expect(getComputedStyle(wrapper).cssFloat).toBe("left");
    expectNear(box.left, rect(wrapper.closest("article")).left);
    const beside = textBeside("case-float-left", box);
    expect(beside.length).toBeGreaterThan(0);
    beside.forEach(r => expect(r.left).toBeGreaterThanOrEqual(box.right - 1));
  });

  it("offsets a top-and-bottom shape from the left in a right-to-left paragraph", () => {
    const wrapper = wrapperOf("case-rtl");
    expectNear(rect(wrapper).left - rect(wrapper.closest("article")).left, 50 * 4 / 3);
  });

  it("places a square-wrapped shape anchored to the page on the page, out of the text flow", () => {
    const box = shapeBox("case-square-page");
    expect(getComputedStyle(box.parentElement).position).toBe("absolute");
    expectNear(rect(box).top - rect(section()).top, 200 / 0.75);
    expectNear(rect(box).left - rect(section()).left, 96);
  });

  it("offsets a paragraph-relative shape from the top of its line", () => {
    const marker = paragraph("case-v-para").querySelector("span");
    expectNear(rect(shapeBox("case-v-para")).top - rect(marker).top, 16);
  });

  it("places a page-relative footer shape from the page bottom when the section is taller", () => {
    const page = section();
    const box = page.querySelector("footer svg").parentElement;
    expect(rect(page).height).toBeGreaterThan(400);
    expectNear(rect(page).bottom - rect(box).bottom, 24);
  });

  it("puts a behind-text shape below the paragraph text", () => {
    const wrapper = wrapperOf("case-behind");
    const centre = rect(shapeBox("case-behind"));
    const hit = document.elementFromPoint((centre.left + centre.right) / 2, (centre.top + centre.bottom) / 2);
    expect(getComputedStyle(wrapper).zIndex).toBe("-1");
    expect(paragraph("case-behind").contains(hit)).toBeTrue();
    expect(wrapper.contains(hit)).toBeFalse();
  });

  it("draws text in a selectable HTML layer that takes pointer events while the shape does not", () => {
    const text = textOf("case-text", "Layer text");
    const layer = text.parentElement;
    const box = shapeBox("case-text");
    expect(text.closest("svg")).toBeNull();
    expect(layer.parentElement).toBe(box);
    expect(getComputedStyle(layer).pointerEvents).toBe("auto");
    expect(getComputedStyle(box).pointerEvents).toBe("none");
    getSelection().selectAllChildren(text);
    expect(getSelection().toString()).toBe("Layer text");
    getSelection().removeAllRanges();
  });

  it("mirrors only the geometry of a flipped shape", () => {
    const box = shapeBox("case-flip");
    expect(getComputedStyle(box.querySelector("svg")).transform).toBe("matrix(-1, 0, 0, 1, 0, 0)");
    for (let el = textOf("case-flip", "Flipped"); el != box; el = el.parentElement)
      expect(getComputedStyle(el).transform).toBe("none");
  });

  it("grows a shape that fits its text", () => {
    expect(rect(shapeBox("case-autofit")).height).toBeGreaterThan(13.4);
  });

  it("keeps the picture wrapper as before", () => {
    const none = wrapperOf("case-pic-none");
    expect(none.style.display).toBe("inline-block");
    expect(none.style.position).toBe("relative");
    expect(none.style.cssFloat).toBe("");
    expect(px(getComputedStyle(none).width)).toBe(0);
    expectNear(px(getComputedStyle(none).left), 96);

    const square = wrapperOf("case-pic-square");
    expect(square.style.cssFloat).toBe("left");
    expect(square.style.left).toBe("");
    expectNear(px(getComputedStyle(square).width), 80 / 3);

    const behind = wrapperOf("case-pic-behind");
    expect(behind.style.zIndex).toBe("");
    expect(px(getComputedStyle(behind).width)).toBe(0);
  });

  it("takes the shape Choice only when it requires the shape namespace and every drawing in it draws", () => {
    const branch = id => paragraph(id).querySelector("svg ellipse") ? "fallback" : paragraph(id).querySelector("svg rect") ? "choice" : null;
    expect(branch("case-choice-empty")).toBe("fallback");
    expect(branch("case-choice-blank")).toBe("fallback");
    expect(branch("case-choice-local")).toBe("choice");
    expect(branch("case-choice-two")).toBe("fallback");
    expect(branch("case-choice-wordart")).toBe("fallback");
    expect(reasons(paragraph("case-choice-empty"))).toEqual([]);
  });

  it("marks graphics and alternate content it does not draw", () => {
    expect(reasons(paragraph("case-marks"))).toEqual(["graphic", "alternate", "alternate"]);
    expect(paragraph("case-marks").querySelectorAll("svg ellipse").length).toBe(1);
  });

  it("draws text boxes nested 16 deep and marks the 17th", () => {
    expect(reasons(paragraph("case-depth-16"))).toEqual([]);
    expect(Array.from(paragraph("case-depth-16").querySelectorAll("p")).some(p => p.textContent == "innermost")).toBeTrue();
    expect(reasons(paragraph("case-depth-17"))).toEqual(["depth"]);
  });

  describe("in a two-column section", () => {
    const square = (id, h) => para(id, run(anchor({ h, v: posV("paragraph", offset(0)), w: 40, ht: 30, wrap: '<wp:wrapSquare wrapText="bothSides"/>' }, shapeGraphic(wsp({ w: 40, h: 30 })))));
    let columns;

    beforeAll(async () => {
      columns = document.createElement("div");
      document.body.prepend(columns);
      const cell = `<w:tbl><w:tr><w:tc>${square("case-cell", posH("column", align("center")))}</w:tc></w:tr></w:tbl>`;
      await docx.renderAsync(await buildDocx(square("case-columns", posH("column", align("center"))) + cell, '<w:cols w:num="2" w:space="240"/>'), columns);
    });

    afterAll(() => columns.remove());

    const wrapperIn = id => Array.from(columns.querySelectorAll("article p")).find(p => p.textContent.startsWith(`${id} `)).querySelector("div");

    it("centres a column-relative shape in its own column", () => {
      const wrapper = wrapperIn("case-columns");
      const box = rect(wrapper);
      expect(getComputedStyle(wrapper).cssFloat).toBe("left");
      expectNear((box.left + box.right) / 2 - rect(wrapper.closest("article")).left, 54 * 4 / 3);
    });

    it("does not float a shape in a table cell", () => {
      expect(getComputedStyle(wrapperIn("case-cell")).cssFloat).toBe("none");
    });
  });

  it("ignores an offset that is not a whole number of EMU", () => {
    const wrapper = wrapperOf("case-unsafe");
    expect(wrapper.style.left).toBe("");
    expect(wrapper.style.marginLeft).toBe("");
  });
});
