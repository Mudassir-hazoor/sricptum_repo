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
  images: ExtractedImage[];
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
  let text = "";
  let previous: PdfTextItem | undefined;

  for (const item of items) {
    if (!item.str) continue;

    if (previous && item.transform && previous.transform) {
      const currentX = item.transform[4] ?? 0;
      const currentY = item.transform[5] ?? 0;
      const previousX = previous.transform[4] ?? 0;
      const previousY = previous.transform[5] ?? 0;
      const lineTolerance = Math.max(
        2,
        Math.max(item.height ?? 0, previous.height ?? 0) * 0.25,
      );
      if (Math.abs(currentY - previousY) > lineTolerance) {
        text = text.trimEnd() + "\n";
      } else {
        const previousEnd = previousX + (previous.width ?? 0);
        const gap = currentX - previousEnd;
        const spaceThreshold =
          Math.max(item.height ?? 0, previous.height ?? 0) * 0.12;
        if (
          gap > spaceThreshold &&
          !/\s$/.test(text) &&
          !/^\s/.test(item.str)
        ) {
          text += " ";
        }
      }
    } else if (text && !/\s$/.test(text) && !/^\s/.test(item.str)) {
      text += " ";
    }

    text += item.str;
    if (item.hasEOL) text = text.trimEnd() + "\n";
    previous = item;
  }

  return text.replace(/[ \t]+\n/g, "\n").trim();
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
  const { default: JSZip } = await import("jszip");
  const archive = await JSZip.loadAsync(new Uint8Array(arrayBuffer));
  const additionalText: string[] = [];
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
    if (paragraphs.length > 0) additionalText.push(...paragraphs);
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
  const text = [textResult.value, ...additionalText]
    .join("\n")
    .replaceAll(String.fromCharCode(0), "")
    .trim();
  if (!text && images.length === 0) {
    throw new Error(
      `${fileName} has no extractable text or embedded images. It may contain only unsupported drawings or scanned text.`,
    );
  }
  return {
    text: text || "(No selectable text. Original images are included below.)",
    images,
  };
}
