describe("tab stops", function () {
    it("measures a declared stop from the paragraph text area", async () => {
        const docBlob = await fetch(`/base/tests/tab-stop-test/document.docx`).then(r => r.blob());

        const div = document.createElement("div");
        document.body.appendChild(div);

        await docx.renderAsync(docBlob, div, null, { experimental: true });

        docx.remeasureTabStops(div);

        // ind left=1440 with a left stop at 1440: the run after the tab starts on the paragraph's own left edge
        const p = div.querySelector("section.docx article > p");
        const spans = Array.from(p.querySelectorAll("span")).filter(s => s.textContent.trim());
        const after = spans[spans.length - 1];

        expect(after.textContent).toBe("Heading text after the tab");

        const delta = after.getBoundingClientRect().left - p.getBoundingClientRect().left;
        expect(Math.abs(delta)).toBeLessThan(6);

        div.remove();
    })
})
