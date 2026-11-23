describe("VML shapes", function () {
  const NS = [
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    'xmlns:v="urn:schemas-microsoft-com:vml"',
    'xmlns:o="urn:schemas-microsoft-com:office:office"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  ].join(" ");
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PAGE = "position:absolute;mso-position-horizontal-relative:page;";

  const pict = inner => `<w:r><w:pict>${inner}</w:pict></w:r>`;
  const para = (id, ...runs) => `<w:p><w:r><w:t xml:space="preserve">${id} </w:t></w:r>${runs.join("")}</w:p>`;
  const textBox = text => `<v:textbox><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></v:textbox>`;
  const textPath = '<v:shape style="width:100pt;height:20pt"><v:textpath string="DRAFT"/></v:shape>';

  const body = [
    para("case-first"),
    para("case-h-page", pict(`<v:rect style="${PAGE}margin-left:72pt;width:20pt;height:10pt;mso-position-vertical-relative:text"/>`)),
    para("case-h-margin", pict(`<v:rect style="position:absolute;margin-left:24pt;width:20pt;height:10pt;mso-position-horizontal-relative:margin;mso-position-vertical-relative:text"/>`)),
    para("case-v-text", pict(`<v:rect style="${PAGE}margin-top:12pt;width:20pt;height:5pt;mso-position-vertical-relative:text"/>`)),
    para("case-bar", pict(`<v:shape coordsize="300,7" path="m300,l,,,7r300,l300,xe" fillcolor="black" stroked="f" style="${PAGE}margin-top:2pt;width:225pt;height:.6pt;mso-position-vertical-relative:text"/>`)),
    para("case-line", pict(`<v:line from="0,0" to="100pt,0" strokeweight="2pt" style="${PAGE}mso-position-vertical-relative:text"/>`)),
    para("case-dash", pict(`<v:rect strokecolor="red [3204]" style="width:20pt;height:10pt"><v:stroke dashstyle="dash"/></v:rect>`)),
    para("case-plain", '<w:r><w:t>Boxed text</w:t></w:r>'),
    para("case-text-box", pict(
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" path="m,l,21600r21600,l21600,xe"/>' +
      `<v:shape type="#_x0000_t202" filled="f" stroked="f" style="width:200pt;height:50pt">${textBox("Boxed text")}</v:shape>`)),
    para("case-stacking", pict(`<v:shape type="#_x0000_t202" style="width:100pt;height:30pt;z-index:-5;visibility:hidden;flip:x">${textBox("Flipped text")}</v:shape>`)),
    para("case-unsafe", pict('<v:rect style="background:url(http://x/);position:fixed;width:10pt;height:10pt"/>')),
    para("case-marks",
      pict('<v:shape style="width:10pt;height:10pt" path="m0,0qx10,10e"/>'),
      pict('<v:polyline points="0,0 10,10"/>'),
      pict('<v:shape style="width:10pt;height:10pt"/>'),
      pict('<v:shape type="#missing" style="width:10pt;height:10pt"/>')),
    para("case-picture-no-type", pict('<v:shape type="#missing" style="width:10pt;height:10pt"><v:imagedata r:id="rIdMissing"/></v:shape>')),
    para("case-watermark-alternate",
      `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing/></mc:Choice><mc:Fallback>${`<w:pict>${textPath}</w:pict>`}</mc:Fallback></mc:AlternateContent></w:r>`),
    para("case-watermark-bare", pict(textPath)),
    ...Array.from({ length: 20 }, (_, i) => para(`filler-${i}`)),
  ].join("\n");

  const footer = `<w:ftr ${NS}><w:p>${pict(`<v:rect style="${PAGE}margin-top:176pt;width:50pt;height:12pt;mso-position-vertical-relative:page"/>`)}</w:p></w:ftr>`;

  const header = `<w:hdr ${NS}><w:p>${pict(`<v:rect style="${PAGE}width:50pt;height:12pt;z-index:-1;mso-position-vertical-relative:page"/>`)}</w:p></w:hdr>`;

  async function buildDocx() {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>
</Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdHeader" Type="${REL}/header" Target="header1.xml"/>
<Relationship Id="rIdFooter" Type="${REL}/footer" Target="footer1.xml"/>
</Relationships>`);
    zip.file("word/header1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
${header}`);
    zip.file("word/footer1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${footer}`);
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>
${body}
<w:sectPr>
<w:headerReference w:type="default" r:id="rIdHeader"/>
<w:footerReference w:type="default" r:id="rIdFooter"/>
<w:pgSz w:w="12240" w:h="4000"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
</w:body></w:document>`);
    return zip.generateAsync({ type: "blob" });
  }

  let container;

  beforeAll(async () => {
    container = document.createElement("div");
    document.body.appendChild(container);
    await docx.renderAsync(await buildDocx(), container);
  });

  afterAll(() => container.remove());

  const section = () => container.querySelector("section.docx");
  const paragraph = id => Array.from(container.querySelectorAll("article p")).find(p => p.textContent.startsWith(`${id} `));
  const shapeBox = id => paragraph(id).querySelector("svg").parentElement;
  const rect = element => element.getBoundingClientRect();
  const expectNear = (actual, expected) => expect(Math.abs(actual - expected)).withContext(`${actual} vs ${expected}`).toBeLessThanOrEqual(1);

  it("draws a path with empty values as zero", () => {
    expect(paragraph("case-bar").querySelector("svg path").getAttribute("d")).toBe("M 300 0 L 0 0 L 0 7 l 300 0 L 300 0 Z");
  });

  it("places a page-relative shape from the page edge", () => {
    expectNear(rect(shapeBox("case-h-page")).left - rect(section()).left, 96);
  });

  it("places a margin-relative shape from the margin, not from the text before it", () => {
    expectNear(rect(shapeBox("case-h-margin")).left - rect(section()).left, 96 + 32);
  });

  it("places a text-relative shape below the top of its line", () => {
    const marker = paragraph("case-v-text").querySelector("span");
    expectNear(rect(shapeBox("case-v-text")).top - rect(marker).top, 16);
  });

  it("places a page-relative footer shape from the page bottom when the section is taller", () => {
    const page = section();
    const box = page.querySelector("footer svg").parentElement;
    expect(rect(page).height).toBeGreaterThan(200 / 0.75);
    expectNear(rect(page).bottom - rect(box).bottom, 16);
  });

  it("draws a thin filled bar without a stroke", () => {
    const box = shapeBox("case-bar");
    const path = box.querySelector("path");
    expect(rect(box).height).toBeGreaterThanOrEqual(0.8);
    expect(path.getAttribute("fill")).toBe("black");
    expect(path.getAttribute("stroke")).toBe("none");
  });

  it("sizes a line box from its ends", () => {
    const box = shapeBox("case-line");
    expectNear(rect(box).width, 133.33);
    expect(rect(box).height).toBeGreaterThanOrEqual(1);
    expect(rect(box.querySelector("line")).width).toBeGreaterThan(0);
  });

  it("strips the theme index from a stroke colour and dashes the stroke", () => {
    const shape = shapeBox("case-dash").querySelector("rect");
    expect(shape.getAttribute("stroke")).toBe("red");
    expect(shape.getAttribute("stroke-dasharray")).toBeTruthy();
  });

  it("draws text box text outside the svg at its own size", () => {
    const boxed = Array.from(paragraph("case-text-box").querySelectorAll("p")).find(p => p.textContent == "Boxed text");
    const plain = Array.from(paragraph("case-plain").querySelectorAll("span")).find(s => s.textContent == "Boxed text");
    expect(boxed.closest("svg")).toBeNull();
    expect(rect(boxed.querySelector("span")).height).toBeCloseTo(rect(plain).height, 0);
  });

  it("maps negative z-index, hidden visibility and flips only the geometry", () => {
    const box = shapeBox("case-stacking");
    const text = Array.from(box.querySelectorAll("p")).find(p => p.textContent == "Flipped text");
    expect(getComputedStyle(box).zIndex).toBe("-1");
    expect(getComputedStyle(box).visibility).toBe("hidden");
    expect(getComputedStyle(box.querySelector("svg")).transform).toBe("matrix(-1, 0, 0, 1, 0, 0)");
    expect(text.closest("svg")).toBeNull();
    for (let el = text; el != box; el = el.parentElement)
      expect(getComputedStyle(el).transform).toBe("none");
  });

  it("drops style values outside the allow-list", () => {
    for (const el of paragraph("case-unsafe").querySelectorAll("*")) {
      expect(getComputedStyle(el).backgroundImage).toBe("none");
      expect(getComputedStyle(el).position).not.toBe("fixed");
    }
  });

  it("marks the shapes it does not draw", () => {
    const reasons = Array.from(paragraph("case-marks").querySelectorAll("[data-docx-not-drawn]"), el => el.getAttribute("data-docx-not-drawn"));
    expect(reasons).toEqual(["vml-path", "vml-element", "vml-shape", "vml-type"]);
    expect(paragraph("case-marks").querySelector("svg")).toBeNull();
  });

  it("marks a watermark as such only in a bare run", () => {
    const reason = id => paragraph(id).querySelector("[data-docx-not-drawn]")?.getAttribute("data-docx-not-drawn");
    expect(reason("case-watermark-alternate")).toBe("vml-element");
    expect(reason("case-watermark-bare")).toBe("watermark");
  });

  it("keeps a shape behind the text visible in the header", () => {
    expect(getComputedStyle(section().querySelector("header")).zIndex).toBe("1");
  });

  it("does not mark a picture shape whose shape type is missing", () => {
    const reasons = Array.from(paragraph("case-picture-no-type").querySelectorAll("[data-docx-not-drawn]")).map(e => e.getAttribute("data-docx-not-drawn"));
    expect(reasons).not.toContain("vml-type");
  });
});
