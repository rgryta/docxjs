(() => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  let container;

  async function render(setting) {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`);
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/document.xml", `<w:document ${W}><w:body><w:p><w:r><w:t>discretionary\u00adhyphen</w:t></w:r></w:p><w:p><w:pPr><w:suppressAutoHyphens/></w:pPr><w:r><w:t>Suppressed</w:t></w:r></w:p><w:p><w:pPr><w:suppressAutoHyphens w:val="false"/></w:pPr><w:r><w:t>Inherited</w:t></w:r></w:p></w:body></w:document>`);
    if (setting !== null) {
      zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="${REL}/settings" Target="settings.xml"/></Relationships>`);
      zip.file("word/settings.xml", `<w:settings ${W}>${setting}</w:settings>`);
    }
    container = document.createElement("div");
    document.body.appendChild(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container, null, { className: "hyphenation-test" });
    return container.querySelector("section");
  }

  afterEach(() => container?.remove());

  describe("document automatic hyphenation", () => {
    it("honours the setting and preserves discretionary hyphens and paragraph suppression", async () => {
      for (const [name, setting, expected] of [
      ["missing settings", null, "manual"],
      ["missing setting", "", "manual"],
      ["false", '<w:autoHyphenation w:val="false"/>', "manual"],
      ["zero", '<w:autoHyphenation w:val="0"/>', "manual"],
      ["off", '<w:autoHyphenation w:val="off"/>', "manual"],
      ["true", '<w:autoHyphenation w:val="true"/>', "auto"],
      ["one", '<w:autoHyphenation w:val="1"/>', "auto"],
      ["on", '<w:autoHyphenation w:val="on"/>', "auto"],
      ["empty element", "<w:autoHyphenation/>", "auto"],
      ]) {
        const section = await render(setting);
        expect(getComputedStyle(section).hyphens).withContext(name).toBe(expected);
        expect(getComputedStyle(section.querySelector("p")).hyphens).toBe(expected);
        expect(section.querySelector("p").textContent).toBe("discretionary\u00adhyphen");
        expect(getComputedStyle(section.querySelectorAll("p")[1]).hyphens).toBe("none");
        expect(getComputedStyle(section.querySelectorAll("p")[2]).hyphens).toBe(expected);
        container.remove();
      }
    });
  });
})();
