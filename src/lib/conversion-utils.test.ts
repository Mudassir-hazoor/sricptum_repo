import { describe, expect, it, vi } from "vitest";
import JSZip from "jszip";
import {
  Document,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
} from "docx";
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
import { extractPdfPageText, extractWordContent } from "./document-extraction";
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
    const imageData = Uint8Array.from(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4AWJiYGD4D8IgBpBmYAAAAAD//7vS9wEAAAAGSURBVAMAGDACA6ybwrYAAAAASUVORK5CYII=",
        "base64",
      ),
    );
    const blob = await createWordDocument("Quarterly report", [
      {
        heading: "Page 1",
        text: "First page text.",
        image: {
          data: imageData,
          contentType: "image/png",
          width: 1,
          height: 1,
        },
      },
      { heading: "Page 2", text: "Second page text." },
    ]);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await archive.file("word/document.xml")?.async("text");
    const stylesXml = await archive.file("word/styles.xml")?.async("text");

    expect(documentXml).toContain("SCRIPTORIUM");
    expect(documentXml).toContain("First page text.");
    expect(documentXml).toContain("Second page text.");
    expect(documentXml).toContain("<w:pageBreakBefore");
    expect(
      Object.keys(archive.files).some((fileName) =>
        fileName.startsWith("word/media/"),
      ),
    ).toBe(true);
    expect(stylesXml).toContain("Arial");
    expect(stylesXml).toContain("203B2D");
  });

  it("keeps paragraph and line boundaries when extracting PDF text", () => {
    expect(
      extractPdfPageText([
        {
          str: "Quarterly",
          transform: [1, 0, 0, 1, 40, 700],
          width: 48,
          height: 12,
        },
        {
          str: "results",
          transform: [1, 0, 0, 1, 92, 700],
          width: 36,
          height: 12,
          hasEOL: true,
        },
        {
          str: "Revenue increased",
          transform: [1, 0, 0, 1, 40, 680],
          width: 100,
          height: 12,
          hasEOL: true,
        },
      ]),
    ).toBe("Quarterly results\nRevenue increased");
  });

  it("extracts Word paragraphs and table content and rejects text-free files", async () => {
    const imageData = Uint8Array.from(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4AWJiYGD4D8IgBpBmYAAAAAD//7vS9wEAAAAGSURBVAMAGDACA6ybwrYAAAAASUVORK5CYII=",
        "base64",
      ),
    );
    vi.stubGlobal("createImageBitmap", async () => ({
      width: 1,
      height: 1,
      close: () => undefined,
    }));
    try {
      const source = new Document({
        sections: [
          {
            children: [
              new Paragraph("Executive summary"),
              new Paragraph("The project met its targets."),
              new Table({
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({ children: [new Paragraph("Metric")] }),
                      new TableCell({ children: [new Paragraph("Result")] }),
                    ],
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [new Paragraph("Completion")],
                      }),
                      new TableCell({ children: [new Paragraph("100%")] }),
                    ],
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new ImageRun({
                    data: imageData,
                    type: "png",
                    transformation: { width: 1, height: 1 },
                  }),
                ],
              }),
            ],
          },
        ],
      });
      const sourceBlob = await Packer.toBlob(source);
      const extracted = await extractWordContent(sourceBlob, "report.docx");
      const imageOnlyBlob = await Packer.toBlob(
        new Document({
          sections: [
            {
              children: [
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: imageData,
                      type: "png",
                      transformation: { width: 1, height: 1 },
                    }),
                  ],
                }),
              ],
            },
          ],
        }),
      );
      const imageOnly = await extractWordContent(imageOnlyBlob, "image.docx");
      const emptyBlob = await Packer.toBlob(
        new Document({ sections: [{ children: [] }] }),
      );

      expect(extracted.text).toContain("Executive summary");
      expect(extracted.text).toContain("The project met its targets.");
      expect(extracted.text).toContain("Metric");
      expect(extracted.text).toContain("100%");
      expect(extracted.images).toHaveLength(1);
      expect(extracted.images[0].contentType).toBe("image/png");
      expect(imageOnly.text).toContain("No selectable text");
      expect(imageOnly.images).toHaveLength(1);
      await expect(extractWordContent(emptyBlob, "empty.docx")).rejects.toThrow(
        /no extractable text/,
      );
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
