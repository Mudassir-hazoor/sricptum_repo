export async function createWordDocument(
  title: string,
  sections: Array<{ heading?: string; text: string }>,
): Promise<Blob> {
  const { Document, HeadingLevel, Packer, Paragraph, TextRun } =
    await import("docx");
  const children = [
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

    for (const paragraph of section.text
      .split(/\n{1,}/)
      .filter((line) => line.trim())) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: paragraph,
              font: "Arial",
              size: 22,
              color: "343A35",
            }),
          ],
          spacing: { after: 180, line: 360 },
        }),
      );
    }
  }

  return Packer.toBlob(
    new Document({
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
