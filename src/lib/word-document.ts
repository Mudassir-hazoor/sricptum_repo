import type { DocumentBlock, ExtractedImage } from "./document-extraction";

export interface WordDocumentSection {
  heading?: string;
  text: string;
  blocks?: DocumentBlock[];
  image?: ExtractedImage;
}

export async function createWordDocument(
  title: string,
  sections: WordDocumentSection[],
): Promise<Blob> {
  const {
    Document,
    HeadingLevel,
    ImageRun,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
    WidthType,
  } = await import("docx");
  const children: Array<
    InstanceType<typeof Paragraph> | InstanceType<typeof Table>
  > = [
    new Paragraph({
      children: [
        new TextRun({
          text: "SCRIPTORIUM  /  DOCUMENT CONVERSION",
          font: "Arial",
          size: 16,
          bold: true,
          color: "66806C",
        }),
      ],
      spacing: { after: 180 },
    }),
    new Paragraph({
      text: title,
      heading: HeadingLevel.TITLE,
      keepNext: true,
    }),
  ];

  for (const [index, section] of sections.entries()) {
    if (section.heading) {
      children.push(
        new Paragraph({
          text: section.heading,
          heading: HeadingLevel.HEADING_1,
          pageBreakBefore: index > 0,
          keepNext: true,
        }),
      );
    }

    if (section.image) {
      const { data, contentType } = section.image;
      const type =
        contentType === "image/jpeg"
          ? "jpg"
          : contentType.slice("image/".length);
      if (
        type !== "png" &&
        type !== "jpg" &&
        type !== "gif" &&
        type !== "bmp"
      ) {
        throw new Error(`Unsupported embedded image format: ${contentType}.`);
      }
      const scale = Math.min(
        620 / section.image.width,
        460 / section.image.height,
        1,
      );
      children.push(
        new Paragraph({
          children: [
            new ImageRun({
              data,
              type,
              transformation: {
                width: section.image.width * scale,
                height: section.image.height * scale,
              },
              altText: {
                name: section.heading ?? "Original document image",
                description:
                  "Embedded image retained from the source document.",
              },
            }),
          ],
          spacing: { after: 180 },
        }),
      );
    }

    const blocks =
      section.blocks ??
      section.text
        .split(/\n+/)
        .filter((paragraph) => paragraph.trim())
        .map((text) => ({ type: "paragraph" as const, text }));
    for (const block of blocks) {
      if (block.type === "table") {
        children.push(
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: block.rows.map(
              (row, rowIndex) =>
                new TableRow({
                  children: row.map(
                    (cell) =>
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({
                                text: cell,
                                font: "Arial",
                                size: 20,
                                color: "343A35",
                                bold: block.hasHeader && rowIndex === 0,
                              }),
                            ],
                          }),
                        ],
                      }),
                  ),
                }),
            ),
          }),
        );
        continue;
      }

      if (block.type === "heading") {
        const heading = [
          HeadingLevel.HEADING_1,
          HeadingLevel.HEADING_2,
          HeadingLevel.HEADING_3,
          HeadingLevel.HEADING_4,
          HeadingLevel.HEADING_5,
          HeadingLevel.HEADING_6,
        ][Math.min(Math.max(block.level, 1), 6) - 1];
        children.push(
          new Paragraph({
            text: block.text,
            heading,
            keepNext: true,
          }),
        );
        continue;
      }

      const isListItem = block.type === "list-item";
      const level = isListItem ? Math.min(block.level, 5) : 0;
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: block.text,
              font: "Arial",
              size: 22,
              color: "343A35",
            }),
          ],
          bullet: isListItem && !block.ordered ? { level } : undefined,
          numbering:
            isListItem && block.ordered
              ? { reference: "scriptorium-ordered-list", level }
              : undefined,
          indent: isListItem
            ? { left: (level + 1) * 360, hanging: 180 }
            : undefined,
          spacing: { after: 180, line: 360 },
        }),
      );
    }
  }

  return Packer.toBlob(
    new Document({
      numbering: {
        config: [
          {
            reference: "scriptorium-ordered-list",
            levels: Array.from({ length: 6 }, (_, level) => ({
              level,
              format: "decimal",
              text: `%${level + 1}.`,
              alignment: "start",
            })),
          },
        ],
      },
      styles: {
        default: {
          document: {
            run: { font: "Arial", size: 22, color: "343A35" },
            paragraph: { spacing: { after: 180, line: 360 } },
          },
          title: {
            run: { font: "Arial", size: 42, bold: true, color: "203B2D" },
            paragraph: { spacing: { after: 280 } },
          },
          heading1: {
            run: { font: "Arial", size: 30, bold: true, color: "356348" },
            paragraph: {
              spacing: { before: 240, after: 140 },
              keepNext: true,
            },
          },
        },
      },
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 1080,
                right: 1080,
                bottom: 1080,
                left: 1080,
              },
            },
          },
          children,
        },
      ],
    }),
  );
}
