(() => {
  describe("Anchored pictures", function () {
    const NS = [
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
    ].join(" ");
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
    const EMU = 12700;
    const SIZE = 21.75 * EMU;

    const avatar = `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" allowOverlap="1" behindDoc="0" locked="0" layoutInCell="1" relativeHeight="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${45.36 * EMU}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${23.04 * EMU}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${SIZE}" cy="${SIZE}"/><wp:effectExtent b="0" l="0" r="0" t="0"/><wp:wrapNone/><wp:docPr id="1" name="" descr=""/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPic"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SIZE}" cy="${SIZE}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    const block = id => `<w:p><w:pPr><w:spacing w:line="300"/></w:pPr>${avatar}<w:r><w:rPr><w:b/></w:rPr><w:br/><w:t xml:space="preserve">${id} </w:t></w:r>` +
      `<w:r><w:t>Body line under the heading</w:t></w:r></w:p>`;

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
<Relationship Id="rIdPic" Type="${REL}/image" Target="media/avatar.png"/>
</Relationships>`);
      zip.file("word/media/avatar.png", PNG, { base64: true });
      zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>
${block("first")}${block("second")}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body></w:document>`);
      return zip.generateAsync({ type: "blob" });
    }

    let container;

    beforeAll(async () => {
      container = document.createElement("div");
      document.body.prepend(container);
      await docx.renderAsync(await buildDocx(), container, null, { className: "anchor-offsets", ignoreLastRenderedPageBreak: true, experimental: true });
    });

    afterAll(() => container.remove());

    const MARGIN = 96;
    const rect = element => element.getBoundingClientRect();
    const paragraph = id => Array.from(container.querySelectorAll("article p")).find(p => p.textContent.startsWith(`${id} `));
    const expectNear = (actual, expected, tolerance = 1) => expect(Math.abs(actual - expected)).withContext(`${actual} vs ${expected}`).toBeLessThanOrEqual(tolerance);

    it("places a page-relative no-wrap picture in the left margin, beside the text", () => {
      const page = rect(container.querySelector("section.anchor-offsets"));

      for (const id of ["first", "second"]) {
        const picture = rect(paragraph(id).querySelector("img"));
        const text = rect(paragraph(id).querySelector("span"));
        expectNear(picture.left - page.left, 45.36 * 4 / 3);
        expectNear(picture.width, 29);
        expect(picture.right - page.left).toBeLessThan(MARGIN);
        expectNear(text.left - page.left, MARGIN);
        expect(picture.right).toBeLessThanOrEqual(text.left);
      }
    });
  });
})();
