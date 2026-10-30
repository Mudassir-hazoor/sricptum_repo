import { describe, expect, it, vi } from "vitest";
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import {
  Document,
  HeadingLevel,
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
import {
  extractPdfPageText,
  extractWordContent,
  parseWordHtml,
  splitDocumentBlocksIntoSlides,
  structurePlainText,
} from "./document-extraction";
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

  it("preserves headings, lists and tables as structured Word blocks", () => {
    vi.stubGlobal("DOMParser", DOMParser);
    try {
      expect(
        parseWordHtml(
          "<h1>Executive summary</h1><p>Project status.</p><ul><li>First action</li><li>Second action<ol><li>Follow up</li></ol></li></ul><table><tr><th>Metric</th><th>Result</th></tr><tr><td>Completion</td><td>100%</td></tr></table>",
        ),
      ).toEqual([
        { type: "heading", text: "Executive summary", level: 1 },
        { type: "paragraph", text: "Project status." },
        { type: "list-item", text: "First action", level: 0, ordered: false },
        { type: "list-item", text: "Second action", level: 0, ordered: false },
        { type: "list-item", text: "Follow up", level: 1, ordered: true },
        {
          type: "table",
          rows: [
            ["Metric", "Result"],
            ["Completion", "100%"],
          ],
          hasHeader: true,
        },
      ]);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("splits structured content at headings without losing table rows", () => {
    expect(
      splitDocumentBlocksIntoSlides([
        { type: "heading", text: "Summary", level: 1 },
        { type: "paragraph", text: "Short overview." },
        { type: "heading", text: "Results", level: 1 },
        {
          type: "table",
          rows: [
            ["Metric", "Value"],
            ["Accuracy", "98%"],
          ],
          hasHeader: true,
        },
      ]),
    ).toEqual([
      [
        { type: "heading", text: "Summary", level: 1 },
        { type: "paragraph", text: "Short overview." },
      ],
      [
        { type: "heading", text: "Results", level: 1 },
        {
          type: "table",
          rows: [
            ["Metric", "Value"],
            ["Accuracy", "98%"],
          ],
          hasHeader: true,
        },
      ],
    ]);
  });

  it("keeps oversized paragraphs and tables complete across slides", () => {
    const text = `${"Alpha beta gamma delta. ".repeat(4)}Omega`;
    const rows = [
      ["Metric", "Value"],
      ...Array.from({ length: 10 }, (_, index) => [
        `Metric ${index + 1}`,
        `Value ${index + 1}`,
      ]),
    ];
    const slides = splitDocumentBlocksIntoSlides(
      [
        { type: "heading", text: "Summary", level: 1 },
        { type: "paragraph", text },
        { type: "table", rows, hasHeader: true },
      ],
      24,
    );
    const paragraphs = slides.flatMap((slide) =>
      slide
        .filter((block) => block.type === "paragraph")
        .map((block) => block.text),
    );
    const tables = slides.flat().filter((block) => block.type === "table");
    const outputRows = [
      tables[0]?.rows[0],
      ...tables.flatMap((block) => block.rows.slice(1)),
    ];

    expect(slides[0][0]).toEqual({
      type: "heading",
      text: "Summary",
      level: 1,
    });
    expect(paragraphs.join(" ")).toBe(text);
    expect(
      tables.every((table) => table.rows[0].join("|") === "Metric|Value"),
    ).toBe(true);
    expect(outputRows).toEqual(rows);
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
        blocks: [
          { type: "heading", text: "Key results", level: 1 },
          {
            type: "list-item",
            text: "First action",
            level: 0,
            ordered: false,
          },
          {
            type: "list-item",
            text: "Second action",
            level: 0,
            ordered: true,
          },
          {
            type: "table",
            rows: [
              ["Metric", "Value"],
              ["Completion", "100%"],
            ],
            hasHeader: true,
          },
          { type: "paragraph", text: "First page text." },
        ],
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
    expect(documentXml).toContain("Key results");
    expect(documentXml).toContain("First action");
    expect(documentXml).toContain("Second action");
    expect(documentXml).toContain("<w:numPr>");
    expect(documentXml).toContain("<w:tbl>");
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

  it("reconstructs headings and list items from extracted PDF text", () => {
    expect(
      structurePlainText(
        "QUARTERLY RESULTS\nRevenue increased by 12%.\n• Expand the team\n2. Review the plan",
      ),
    ).toEqual([
      { type: "heading", text: "QUARTERLY RESULTS", level: 1 },
      { type: "paragraph", text: "Revenue increased by 12%." },
      { type: "list-item", text: "Expand the team", level: 0, ordered: false },
      { type: "list-item", text: "Review the plan", level: 0, ordered: true },
    ]);
  });

  it("extracts Word paragraphs and table content and rejects text-free files", async () => {
    const imageData = Uint8Array.from(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4AWJiYGD4D8IgBpBmYAAAAAD//7vS9wEAAAAGSURBVAMAGDACA6ybwrYAAAAASUVORK5CYII=",
        "base64",
      ),
    );
    vi.stubGlobal("DOMParser", DOMParser);
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
              new Paragraph({
                text: "Executive summary",
                heading: HeadingLevel.HEADING_1,
              }),
              new Paragraph("The project met its targets."),
              new Paragraph({
                text: "Review the findings",
                bullet: { level: 0 },
              }),
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
      expect(extracted.blocks).toContainEqual({
        type: "heading",
        text: "Executive summary",
        level: 1,
      });
      expect(extracted.blocks).toContainEqual({
        type: "list-item",
        text: "Review the findings",
        level: 0,
        ordered: false,
      });
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
