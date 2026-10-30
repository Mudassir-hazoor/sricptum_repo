import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import {
  formatFileSize,
  getFileExtension,
  isOperationId,
  parsePageRanges,
  resolvePresentationSlidePaths,
  safeFileStem,
  splitTextIntoSlides,
  validateInputFiles,
} from "./conversion-utils";
import { createWordDocument } from "./word-document";

function makeFile(name: string, size = 1): File {
  return { name, size } as File;
}

describe("document conversion helpers", () => {
  it("parses page numbers and ascending ranges", () => {
    expect(parsePageRanges("1-3, 5, 7-8", 10)).toEqual([1, 2, 3, 5, 7, 8]);
  });

  it.each(["", "a", "1,", "0", "4-2", "1-8"])(
    "rejects invalid page ranges: %s",
    (range) => {
      expect(() => parsePageRanges(range, 7)).toThrow();
    },
  );

  it("rejects duplicated pages in overlapping ranges", () => {
    expect(() => parsePageRanges("1-3, 3-4", 5)).toThrow(
      /included more than once/,
    );
  });

  it("splits text into readable slide-sized chunks", () => {
    const slides = splitTextIntoSlides("one two three four five", 9);
    expect(slides).toEqual(["one two", "three", "four five"]);
    expect(slides.every((slide) => slide.length <= 9)).toBe(true);
  });

  it("handles single words longer than a slide", () => {
    expect(splitTextIntoSlides("abcdefghij", 4)).toEqual([
      "abcd",
      "efgh",
      "ij",
    ]);
  });

  it("creates safe output names from uploaded file names", () => {
    expect(safeFileStem("Quarterly report?.pdf")).toBe("Quarterly report-");
    expect(safeFileStem("...")).toBe("document");
  });

  it("checks input extensions and merge selection count", () => {
    expect(
      validateInputFiles("word-to-pdf", [makeFile("brief.docx")]),
    ).toBeUndefined();
    expect(validateInputFiles("word-to-pdf", [makeFile("brief.rtf")])).toMatch(
      /isn’t supported/,
    );
    expect(validateInputFiles("merge-pdf", [makeFile("one.pdf")])).toMatch(
      /at least two/,
    );
  });

  it("limits the combined size of selected files", () => {
    expect(
      validateInputFiles("merge-pdf", [
        makeFile("one.pdf", 27 * 1024 * 1024),
        makeFile("two.pdf", 27 * 1024 * 1024),
        makeFile("three.pdf", 27 * 1024 * 1024),
      ]),
    ).toMatch(/processing limit/);
  });

  it("recognises only available conversion tools", () => {
    expect(isOperationId("images-to-pdf")).toBe(true);
    expect(isOperationId("does-not-exist")).toBe(false);
  });

  it("uses the presentation relationship list to preserve slide order", () => {
    const relationships = new Map([
      ["rId1", "slides/slide2.xml"],
      ["rId2", "slides/slide1.xml"],
    ]);
    const availableSlides = new Set([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);

    expect(
      resolvePresentationSlidePaths(
        ["rId1", "rId2"],
        relationships,
        availableSlides,
      ),
    ).toEqual(["ppt/slides/slide2.xml", "ppt/slides/slide1.xml"]);
    expect(() =>
      resolvePresentationSlidePaths(["rId3"], relationships, availableSlides),
    ).toThrow(/invalid slide reference/);
  });

  it("formats bytes for the file list", () => {
    expect(formatFileSize(420)).toBe("420 B");
    expect(formatFileSize(2048)).toBe("2 KB");
    expect(formatFileSize(1572864)).toBe("1.5 MB");
  });

  it("recognises extensions without depending on filename case", () => {
    expect(getFileExtension("PRESENTATION.PPTX")).toBe(".pptx");
    expect(getFileExtension("missing-extension")).toBe("");
  });

  it("creates a styled Word document with page breaks for source pages", async () => {
    const blob = await createWordDocument("Quarterly report", [
      { heading: "Page 1", text: "First page text." },
      { heading: "Page 2", text: "Second page text." },
    ]);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await archive.file("word/document.xml")?.async("text");
    const stylesXml = await archive.file("word/styles.xml")?.async("text");

    expect(documentXml).toContain("SCRIPTORIUM");
    expect(documentXml).toContain("First page text.");
    expect(documentXml).toContain("Second page text.");
    expect(documentXml).toContain("<w:pageBreakBefore");
    expect(stylesXml).toContain("Arial");
    expect(stylesXml).toContain("203B2D");
  });
});
