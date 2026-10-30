export const MAX_FILE_SIZE = 30 * 1024 * 1024;
export const MAX_TOTAL_FILE_SIZE = 80 * 1024 * 1024;
export const MAX_PDF_PAGES = 100;
export const MAX_POWERPOINT_SLIDES = 80;

export const TOOL_GROUPS = [
  {
    name: "PDF tools",
    ids: [
      "pdf-to-word",
      "pdf-to-powerpoint",
      "compress-pdf",
      "merge-pdf",
      "split-pdf",
      "pdf-to-images",
    ],
  },
  {
    name: "Word tools",
    ids: ["word-to-pdf", "word-to-powerpoint", "powerpoint-to-word"],
  },
  {
    name: "PowerPoint tools",
    ids: ["powerpoint-to-pdf"],
  },
  {
    name: "Image tools",
    ids: ["images-to-pdf"],
  },
] as const;

export const TOOLS = {
  "pdf-to-word": {
    title: "PDF to Word",
    description:
      "Extract selectable text from a PDF into an editable Word document.",
    inputFormats: [".pdf"],
    outputFormat: "DOCX",
    maxFiles: 1,
    detail:
      "Text is extracted page by page. Scanned pages need OCR, which isn’t included.",
  },
  "word-to-pdf": {
    title: "Word to PDF",
    description: "Create a clean, shareable PDF from a Word document.",
    inputFormats: [".docx"],
    outputFormat: "PDF",
    maxFiles: 1,
    detail:
      "Text is carried across; complex page layouts, fonts and embedded objects may differ.",
  },
  "powerpoint-to-word": {
    title: "PowerPoint to Word",
    description:
      "Collect presentation slide text in an editable Word document.",
    inputFormats: [".pptx"],
    outputFormat: "DOCX",
    maxFiles: 1,
    detail:
      "Slide text and slide order are kept. Slide design, images, charts and animations are not.",
  },
  "word-to-powerpoint": {
    title: "Word to PowerPoint",
    description: "Turn Word paragraphs into a simple, editable slide deck.",
    inputFormats: [".docx"],
    outputFormat: "PPTX",
    maxFiles: 1,
    detail:
      "The text is organised into clean slides; the original Word layout is not reproduced.",
  },
  "powerpoint-to-pdf": {
    title: "PowerPoint to PDF",
    description:
      "Save the text from a PowerPoint presentation as a readable PDF.",
    inputFormats: [".pptx"],
    outputFormat: "PDF",
    maxFiles: 1,
    detail:
      "Creates a text-first handout with one page per slide; slide visuals and animations are not reproduced.",
  },
  "pdf-to-powerpoint": {
    title: "PDF to PowerPoint",
    description: "Create an editable slide deck from the text in a PDF.",
    inputFormats: [".pdf"],
    outputFormat: "PPTX",
    maxFiles: 1,
    detail:
      "Creates one text-first slide per PDF page. Original page design and images are not recreated.",
  },
  "compress-pdf": {
    title: "Compress PDF",
    description: "Reduce PDF size by turning its pages into optimised images.",
    inputFormats: [".pdf"],
    outputFormat: "PDF",
    maxFiles: 1,
    detail:
      "Compression rasterises each page. Text will no longer be selectable and interactive forms are removed.",
  },
  "merge-pdf": {
    title: "Merge PDFs",
    description: "Combine PDF files into one document in the order you select.",
    inputFormats: [".pdf"],
    outputFormat: "PDF",
    maxFiles: 10,
    detail:
      "Select two or more PDFs. Their pages are appended in the file order shown.",
  },
  "split-pdf": {
    title: "Split PDF",
    description: "Save selected pages from a PDF as a separate document.",
    inputFormats: [".pdf"],
    outputFormat: "PDF",
    maxFiles: 1,
    detail: "Enter page numbers or ranges, for example 1-3, 5.",
  },
  "images-to-pdf": {
    title: "Images to PDF",
    description: "Arrange images as full-page PDF sheets.",
    inputFormats: [".png", ".jpg", ".jpeg", ".webp"],
    outputFormat: "PDF",
    maxFiles: 20,
    detail:
      "Select up to 20 PNG, JPEG or WebP images. Each image gets its own page.",
  },
  "pdf-to-images": {
    title: "PDF to images",
    description: "Export every page of a PDF as a crisp PNG image.",
    inputFormats: [".pdf"],
    outputFormat: "ZIP",
    maxFiles: 1,
    detail:
      "Creates one PNG per page in a ZIP archive. Maximum 50 pages per export.",
  },
} as const;

export type OperationId = keyof typeof TOOLS;

export function getFileExtension(name: string): string {
  const lastDot = name.lastIndexOf(".");
  return lastDot < 0 ? "" : name.slice(lastDot).toLowerCase();
}

export function isOperationId(value: string): value is OperationId {
  return Object.hasOwn(TOOLS, value);
}

export function resolvePresentationSlidePaths(
  relationshipIds: string[],
  relationships: ReadonlyMap<string, string>,
  availableFiles: ReadonlySet<string>,
): string[] {
  if (
    relationshipIds.length === 0 ||
    relationshipIds.length > MAX_POWERPOINT_SLIDES
  ) {
    throw new Error(
      `This presentation must contain between 1 and ${MAX_POWERPOINT_SLIDES} standard PowerPoint slides.`,
    );
  }

  const paths = relationshipIds.map((id) => {
    const target = relationships.get(id);
    if (
      !target ||
      target.startsWith("/") ||
      target.includes("\\") ||
      /^[a-z]+:/i.test(target)
    ) {
      throw new Error("This presentation contains an invalid slide reference.");
    }

    const path = new URL(
      target,
      "https://scriptorium.invalid/ppt/presentation.xml",
    ).pathname.slice(1);
    if (
      !/^ppt\/slides\/slide\d+\.xml$/.test(path) ||
      !availableFiles.has(path)
    ) {
      throw new Error(
        "This presentation references a slide that cannot be read.",
      );
    }
    return path;
  });

  if (new Set(paths).size !== paths.length) {
    throw new Error(
      "This presentation references the same slide more than once.",
    );
  }
  return paths;
}

export function safeFileStem(name: string): string {
  return (
    name
      .replace(/\.[^.]+$/, "")
      .normalize("NFKC")
      .replace(/[<>:"/\\|?*]/g, "-")
      .split("")
      .map((character) => (character.charCodeAt(0) < 32 ? "-" : character))
      .join("")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/[. ]+$/g, "")
      .slice(0, 100) || "document"
  );
}

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function parsePageRanges(input: string, pageCount: number): number[] {
  const value = input.trim();
  if (!value) throw new Error("Enter at least one page number or range.");
  if (!Number.isInteger(pageCount) || pageCount < 1) {
    throw new Error("The document has no pages to select.");
  }

  const pages = new Set<number>();
  for (const rawPart of value.split(",")) {
    const part = rawPart.trim();
    const match = /^(\d+)(?:\s*-\s*(\d+))?$/.exec(part);
    if (!match)
      throw new Error(`“${part}” isn’t a valid page number or range.`);

    const first = Number(match[1]);
    const last = Number(match[2] ?? match[1]);
    if (first < 1 || last < first || last > pageCount) {
      throw new Error(
        `Choose pages between 1 and ${pageCount}; ranges must ascend.`,
      );
    }

    for (let page = first; page <= last; page += 1) {
      if (pages.has(page))
        throw new Error(`Page ${page} is included more than once.`);
      pages.add(page);
    }
  }

  return [...pages];
}

export function splitTextIntoSlides(
  text: string,
  maxCharacters = 620,
): string[] {
  const words = text.replace(/\r\n?/g, "\n").split(/\s+/).filter(Boolean);
  if (words.length === 0) return [];

  const slides: string[] = [];
  let slide = "";
  for (const word of words) {
    if (word.length > maxCharacters) {
      if (slide) slides.push(slide);
      for (let offset = 0; offset < word.length; offset += maxCharacters) {
        slides.push(word.slice(offset, offset + maxCharacters));
      }
      slide = "";
      continue;
    }

    const next = slide ? `${slide} ${word}` : word;
    if (next.length > maxCharacters && slide) {
      slides.push(slide);
      slide = word;
    } else {
      slide = next;
    }
  }

  if (slide) slides.push(slide);
  return slides;
}

export function validateInputFiles(
  operation: OperationId,
  files: File[],
): string | undefined {
  const tool = TOOLS[operation];
  if (files.length === 0) return "Choose a file to get started.";
  if (files.length > tool.maxFiles) {
    return `Choose no more than ${tool.maxFiles} file${tool.maxFiles === 1 ? "" : "s"} for this tool.`;
  }
  if (operation === "merge-pdf" && files.length < 2) {
    return "Choose at least two PDF files to merge.";
  }

  for (const file of files) {
    const extension = getFileExtension(file.name);
    if (!tool.inputFormats.some((format) => format === extension)) {
      return `“${file.name}” isn’t supported by ${tool.title}. Choose ${tool.inputFormats.join(", ")} files.`;
    }
    if (file.size === 0) return `“${file.name}” is empty. Choose another file.`;
    if (file.size > MAX_FILE_SIZE) {
      return `“${file.name}” is larger than the 30 MB per-file limit.`;
    }
  }
  if (
    files.reduce((total, file) => total + file.size, 0) > MAX_TOTAL_FILE_SIZE
  ) {
    return `Selected files together exceed the ${formatFileSize(MAX_TOTAL_FILE_SIZE)} processing limit.`;
  }
}
