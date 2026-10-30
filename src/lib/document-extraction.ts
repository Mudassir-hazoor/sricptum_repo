export interface PdfTextItem {
  str: string;
  transform?: number[];
  width?: number;
  height?: number;
  hasEOL?: boolean;
}

export interface ExtractedImage {
  data: Uint8Array;
  contentType: string;
  width: number;
  height: number;
}

export interface ExtractedWordContent {
  text: string;
  blocks: DocumentBlock[];
  images: ExtractedImage[];
}

export type DocumentBlock =
  | { type: "paragraph"; text: string }
  | { type: "heading"; text: string; level: number }
  | {
      type: "list-item";
      text: string;
      level: number;
      ordered: boolean;
    }
  | { type: "table"; rows: string[][]; hasHeader: boolean };

export type TextDocumentBlock = Exclude<DocumentBlock, { type: "table" }>;

function elementChildren(node: Element): Element[] {
  return Array.from(node.childNodes).filter(
    (child): child is Element => child.nodeType === 1,
  );
}

function readElementText(element: Element): string {
  return (element.textContent ?? "").replace(/\s+/g, " ").trim();
}

function readList(list: Element, level: number, blocks: DocumentBlock[]): void {
  const ordered = list.tagName.toLowerCase() === "ol";
  for (const item of elementChildren(list).filter(
    (child) => child.tagName.toLowerCase() === "li",
  )) {
    const textContainer = item.cloneNode(true) as Element;
    for (const nestedList of Array.from(
      textContainer.getElementsByTagName("ul"),
    )) {
      nestedList.parentNode?.removeChild(nestedList);
    }
    for (const nestedList of Array.from(
      textContainer.getElementsByTagName("ol"),
    )) {
      nestedList.parentNode?.removeChild(nestedList);
    }
    const text = readElementText(textContainer);
    if (text) blocks.push({ type: "list-item", text, level, ordered });
    for (const nested of elementChildren(item).filter((child) =>
      ["ul", "ol"].includes(child.tagName.toLowerCase()),
    )) {
      readList(nested, level + 1, blocks);
    }
  }
}

export function parseWordHtml(html: string): DocumentBlock[] {
  const document = new DOMParser().parseFromString(
    `<scriptorium>${html}</scriptorium>`,
    "application/xml",
  );
  if (document.getElementsByTagName("parsererror").length > 0) {
    throw new Error("The Word document contains unreadable formatted content.");
  }

  const blocks: DocumentBlock[] = [];
  const visit = (element: Element): void => {
    const tag = element.tagName.toLowerCase();
    if (/^h[1-6]$/.test(tag)) {
      const text = readElementText(element);
      if (text) {
        blocks.push({
          type: "heading",
          text,
          level: Number(tag.slice(1)),
        });
      }
      return;
    }
    if (tag === "p") {
      const text = readElementText(element);
      if (text) blocks.push({ type: "paragraph", text });
      return;
    }
    if (tag === "ul" || tag === "ol") {
      readList(element, 0, blocks);
      return;
    }
    if (tag === "table") {
      const rows = Array.from(element.getElementsByTagName("tr"))
        .map((row) =>
          elementChildren(row)
            .filter((cell) => ["td", "th"].includes(cell.tagName.toLowerCase()))
            .map(readElementText),
        )
        .filter((row) => row.length > 0);
      if (rows.length > 0) {
        const firstRow = element.getElementsByTagName("tr")[0];
        const hasHeader = elementChildren(firstRow ?? element).some(
          (cell) => cell.tagName.toLowerCase() === "th",
        );
        blocks.push({ type: "table", rows, hasHeader });
      }
      return;
    }
    for (const child of elementChildren(element)) visit(child);
  };

  const root = document.documentElement;
  if (root) {
    for (const child of elementChildren(root)) visit(child);
  }
  return blocks;
}

export function documentBlockText(block: DocumentBlock): string {
  if (block.type === "table") {
    return block.rows.map((row) => row.join(" | ")).join("\n");
  }
  if (block.type === "list-item") {
    return `${"  ".repeat(block.level)}${block.ordered ? "1." : "•"} ${block.text}`;
  }
  return block.text;
}

export function splitDocumentBlocksIntoSlides(
  blocks: DocumentBlock[],
  maxCharacters = 520,
): DocumentBlock[][] {
  const slides: DocumentBlock[][] = [];
  let slide: DocumentBlock[] = [];
  let length = 0;

  const finishSlide = (): void => {
    if (slide.length > 0) slides.push(slide);
    slide = [];
    length = 0;
  };

  for (const block of blocks) {
    const blockLength = documentBlockText(block).length;
    if (block.type === "heading" && slide.length > 0) finishSlide();
    if (block.type === "paragraph" && blockLength > maxCharacters) {
      let remaining = "";
      for (const word of block.text.split(/\s+/).filter(Boolean)) {
        if (word.length > maxCharacters) {
          if (remaining) {
            slide.push({ type: "paragraph", text: remaining });
            length += remaining.length;
            finishSlide();
            remaining = "";
          }
          let wordOffset = 0;
          if (slide.length > 0) {
            const available = maxCharacters - length;
            if (available > 0) {
              slide.push({
                type: "paragraph",
                text: word.slice(0, available),
              });
              wordOffset = available;
              finishSlide();
            } else {
              finishSlide();
            }
          }
          for (
            let offset = wordOffset;
            offset < word.length;
            offset += maxCharacters
          ) {
            slide.push({
              type: "paragraph",
              text: word.slice(offset, offset + maxCharacters),
            });
            finishSlide();
          }
          continue;
        }
        const next = remaining ? `${remaining} ${word}` : word;
        if (next.length > maxCharacters - length && remaining) {
          slide.push({ type: "paragraph", text: remaining });
          length += remaining.length;
          finishSlide();
          remaining = word;
        } else if (next.length > maxCharacters - length && slide.length > 0) {
          finishSlide();
          remaining = word;
        } else {
          remaining = next;
        }
      }
      if (remaining) {
        slide.push({ type: "paragraph", text: remaining });
        length += remaining.length;
      }
      finishSlide();
      continue;
    }

    if (block.type === "table" && blockLength > maxCharacters) {
      if (slide.length > 0) finishSlide();
      const [header, ...dataRows] = block.rows;
      let rows = block.hasHeader && header ? [header] : [];
      let rowsLength = rows.reduce(
        (total, row) => total + row.join(" | ").length,
        0,
      );
      for (const row of block.hasHeader ? dataRows : block.rows) {
        const rowLength = row.join(" | ").length;
        if (
          rows.length > (block.hasHeader ? 1 : 0) &&
          rowsLength + rowLength > maxCharacters
        ) {
          slide.push({
            type: "table",
            rows,
            hasHeader: block.hasHeader,
          });
          finishSlide();
          rows = block.hasHeader && header ? [header] : [];
          rowsLength = rows.reduce(
            (total, headerRow) => total + headerRow.join(" | ").length,
            0,
          );
        }
        rows.push(row);
        rowsLength += rowLength;
      }
      if (rows.length > 0) {
        slide.push({
          type: "table",
          rows,
          hasHeader: block.hasHeader,
        });
        finishSlide();
      }
      continue;
    }

    if (slide.length > 0 && length + blockLength > maxCharacters) {
      finishSlide();
    }
    slide.push(block);
    length += blockLength;
  }
  finishSlide();
  return slides;
}

function copyImageData(data: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(data.byteLength);
  copy.set(data);
  return copy.buffer;
}

export async function normalizeEmbeddedImage(
  data: Uint8Array,
  contentType: string,
  fileName: string,
): Promise<ExtractedImage> {
  let bitmap: ImageBitmap;
  try {
    bitmap = await createImageBitmap(
      new Blob([copyImageData(data)], { type: contentType }),
    );
  } catch (error) {
    throw new Error(
      `${fileName} contains an embedded image that this browser cannot preserve. Convert it to PNG or JPEG and try again.`,
      { cause: error },
    );
  }
  if (bitmap.width * bitmap.height > 80_000_000) {
    bitmap.close();
    throw new Error(
      `${fileName} contains an image larger than the 80-megapixel preservation limit.`,
    );
  }

  const { width, height } = bitmap;
  if (
    ["image/png", "image/jpeg", "image/gif", "image/bmp"].includes(contentType)
  ) {
    const image = {
      data,
      contentType,
      width,
      height,
    };
    bitmap.close();
    return image;
  }

  const canvas = document.createElement("canvas");
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  const context = canvas.getContext("2d", { alpha: false });
  if (!context) {
    bitmap.close();
    throw new Error("Your browser couldn’t create an image canvas.");
  }
  context.drawImage(bitmap, 0, 0);
  bitmap.close();
  const png = await new Promise<Blob>((resolve, reject) => {
    canvas.toBlob(
      (output) =>
        output
          ? resolve(output)
          : reject(
              new Error("Your browser couldn’t convert an embedded image."),
            ),
      "image/png",
    );
  });
  canvas.width = 0;
  canvas.height = 0;
  return {
    data: new Uint8Array(await png.arrayBuffer()),
    contentType: "image/png",
    width,
    height,
  };
}

export function extractPdfPageText(items: PdfTextItem[]): string {
  const normalized = items
    .map((item) => {
      const transform = item.transform ?? [1, 0, 0, 1, 0, 0];
      const x = transform[4] ?? 0;
      const y = transform[5] ?? 0;
      const text = item.str.replace(/\u00a0/g, " ").trim();
      return {
        ...item,
        text,
        x,
        y,
        height: item.height ?? 12,
      };
    })
    .filter(({ text }) => text.length > 0);

  if (normalized.length === 0) return "";

  normalized.sort((first, second) => {
    const yDifference = (second.y ?? 0) - (first.y ?? 0);
    if (Math.abs(yDifference) > 0.001) return yDifference;
    return (first.x ?? 0) - (second.x ?? 0);
  });

  const lines: string[] = [];
  let currentLine:
    | {
        y: number;
        text: string;
        lastX: number;
      }
    | undefined;

  for (const item of normalized) {
    const lineTolerance = Math.max(2, item.height * 0.6);
    const nextItemX = item.x ?? 0;
    const nextItemEnd = nextItemX + (item.width ?? 0);

    if (!currentLine || Math.abs(item.y - currentLine.y) > lineTolerance) {
      if (currentLine) lines.push(currentLine.text.trimEnd());
      currentLine = {
        y: item.y,
        text: item.text,
        lastX: nextItemEnd,
      };
      continue;
    }

    const previousEnd = currentLine.lastX;
    const gap = nextItemX - previousEnd;
    const spaceThreshold = Math.max(2, item.height * 0.18);
    if (
      gap > spaceThreshold &&
      !/\s$/.test(currentLine.text) &&
      !/^\s/.test(item.text)
    ) {
      currentLine.text += " ";
    }
    currentLine.text += item.text;
    currentLine.lastX = nextItemEnd;
  }

  if (currentLine) lines.push(currentLine.text.trimEnd());

  return lines
    .join("\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export function structurePlainText(text: string): TextDocumentBlock[] {
  const blocks: TextDocumentBlock[] = [];
  for (const rawLine of text.split(/\n+/)) {
    const line = rawLine.trim();
    if (!line) continue;
    const listItem = /^([•●▪◦*-]|\d+[.)])\s+(.+)$/.exec(line);
    if (listItem) {
      blocks.push({
        type: "list-item",
        text: listItem[2],
        level: Math.min(
          Math.floor((rawLine.length - rawLine.trimStart().length) / 2),
          5,
        ),
        ordered: /^\d/.test(listItem[1]),
      });
      continue;
    }

    const letters = line.match(/[A-Za-z]/g) ?? [];
    if (
      line.length <= 90 &&
      letters.length >= 3 &&
      letters.filter((character) => character === character.toUpperCase())
        .length /
        letters.length >=
        0.85 &&
      !/[.!?:;]$/.test(line)
    ) {
      blocks.push({ type: "heading", text: line, level: 1 });
    } else {
      blocks.push({ type: "paragraph", text: line });
    }
  }
  return blocks;
}

export async function extractWordContent(
  file: Blob,
  fileName: string,
): Promise<ExtractedWordContent> {
  const { default: mammoth } = await import("mammoth");
  const arrayBuffer = await file.arrayBuffer();
  const input =
    typeof Buffer === "undefined"
      ? { arrayBuffer }
      : { buffer: Buffer.from(arrayBuffer) };
  const [textResult, htmlResult] = await Promise.all([
    mammoth.extractRawText(input),
    mammoth.convertToHtml(input, {
      convertImage: mammoth.images.imgElement(async (image) => ({
        src: `data:${image.contentType};base64,${await image.read("base64")}`,
      })),
    }),
  ]);
  const blocks = parseWordHtml(htmlResult.value);
  const { default: JSZip } = await import("jszip");
  const archive = await JSZip.loadAsync(new Uint8Array(arrayBuffer));
  const additionalFiles = Object.keys(archive.files).filter((path) =>
    /^word\/(?:header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i.test(
      path,
    ),
  );
  for (const path of additionalFiles) {
    const xmlText = await archive.file(path)?.async("string");
    if (!xmlText) continue;
    if (!/<(?:[A-Za-z0-9_]+:)?t(?:\s|>)/.test(xmlText)) continue;
    const xml = new DOMParser().parseFromString(xmlText, "application/xml");
    if (xml.querySelector("parsererror")) {
      throw new Error(`${fileName} contains unreadable supplemental text.`);
    }
    const paragraphs = Array.from(xml.getElementsByTagNameNS("*", "p"))
      .map((paragraph) =>
        Array.from(paragraph.getElementsByTagNameNS("*", "t"))
          .map((run) => run.textContent ?? "")
          .join(""),
      )
      .filter((paragraph) => paragraph.trim());
    if (paragraphs.length > 0) {
      blocks.push(
        ...paragraphs.map((text) => ({ type: "paragraph" as const, text })),
      );
    }
  }
  const images: ExtractedImage[] = [];
  const imagePattern =
    /<img\b[^>]*\bsrc=["'](data:(image\/[^;]+);base64,([A-Za-z0-9+/=]+))["']/gi;
  for (const match of htmlResult.value.matchAll(imagePattern)) {
    const binary = atob(match[3]);
    const data = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    images.push(
      await normalizeEmbeddedImage(data, match[2].toLowerCase(), fileName),
    );
    if (
      images.reduce((total, image) => total + image.data.byteLength, 0) >
      80 * 1024 * 1024
    ) {
      throw new Error(
        `${fileName} contains more than 80 MB of embedded images. Reduce the image count or size and try again.`,
      );
    }
  }
  const blockText = blocks.map(documentBlockText).join("\n");
  const text = (blockText || textResult.value)
    .replaceAll(String.fromCharCode(0), "")
    .trim();
  if (!text && images.length === 0) {
    throw new Error(
      `${fileName} has no extractable text or embedded images. It may contain only unsupported drawings or scanned text.`,
    );
  }
  return {
    text: text || "(No selectable text. Original images are included below.)",
    blocks:
      blocks.length > 0
        ? blocks
        : text
            .split(/\n+/)
            .filter((paragraph: string) => paragraph.trim())
            .map((paragraph: string) => ({
              type: "paragraph" as const,
              text: paragraph,
            })),
    images,
  };
}
