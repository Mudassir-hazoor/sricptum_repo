import {
  formatFileSize,
  isOperationId,
  MAX_FILE_SIZE,
  MAX_PDF_PAGES,
  MAX_POWERPOINT_SLIDES,
  MAX_TOTAL_FILE_SIZE,
  parsePageRanges,
  resolvePresentationSlidePaths,
  safeFileStem,
  splitTextIntoSlides,
  TOOLS,
  type OperationId,
  validateInputFiles,
} from "../lib/conversion-utils";
import {
  extractPdfPageText,
  extractWordContent,
  normalizeEmbeddedImage,
  structurePlainText,
  splitDocumentBlocksIntoSlides,
  type DocumentBlock,
  type ExtractedImage,
} from "../lib/document-extraction";
import {
  createWordDocument,
  type WordDocumentSection,
} from "../lib/word-document";

type PdfDocumentProxy = import("pdfjs-dist").PDFDocumentProxy;

let pdfjsModule: Promise<typeof import("pdfjs-dist")> | undefined;

function getPdfjs(): Promise<typeof import("pdfjs-dist")> {
  pdfjsModule ??= Promise.all([
    import("pdfjs-dist"),
    import("pdfjs-dist/build/pdf.worker.min.mjs?url"),
  ]).then(([pdfjs, worker]) => {
    pdfjs.GlobalWorkerOptions.workerSrc = worker.default;
    return pdfjs;
  });
  return pdfjsModule;
}

function imageDataBuffer(data: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(data.byteLength);
  copy.set(data);
  return copy.buffer;
}

const FILE_LIMIT = MAX_FILE_SIZE;
const PDF_EXPORT_LIMIT = 50;

const toolButtons = Array.from(
  document.querySelectorAll<HTMLButtonElement>("[data-tool]"),
);
if (toolButtons.length > 0) {
  const fileInput = requireElement<HTMLInputElement>("#file-input");
  const dropZone = requireElement<HTMLElement>("#drop-zone");
  const fileList = requireElement<HTMLElement>("#file-list");
  const runButton = requireElement<HTMLButtonElement>("#run-conversion");
  const status = requireElement<HTMLElement>("#workspace-status");
  const progress = requireElement<HTMLProgressElement>("#progress-bar");
  const progressLabel = requireElement<HTMLElement>("#progress-label");
  const results = requireElement<HTMLElement>("#conversion-results");
  const rangeOptions = requireElement<HTMLElement>("#range-options");
  const compressionOptions = requireElement<HTMLElement>(
    "#compression-options",
  );
  const compressionWarning = requireElement<HTMLInputElement>(
    "#compression-confirmation",
  );
  const pageRange = requireElement<HTMLInputElement>("#page-range");
  const workspaceTitle = requireElement<HTMLElement>("#workspace-title");
  const workspaceDescription = requireElement<HTMLElement>(
    "#workspace-description",
  );
  const outputFormat = requireElement<HTMLElement>("#output-format");
  const inputLabel = requireElement<HTMLElement>("#input-label");
  const dropTitle = requireElement<HTMLElement>("#drop-title");
  const dropHelp = requireElement<HTMLElement>("#drop-help");
  const conversionDetail = requireElement<HTMLElement>("#conversion-detail");
  const uploadAction = requireElement<HTMLButtonElement>("#select-files");

  let activeOperation: OperationId = "pdf-to-word";
  let selectedFiles: File[] = [];
  let resultUrl: string | undefined;
  let isConverting = false;

  const activeConfig = () => TOOLS[activeOperation];

  function setStatus(message: string, isError = false): void {
    status.textContent = message;
    status.classList.toggle("error", isError);
  }

  function setProgress(value: number, label: string): void {
    progress.value = value;
    progressLabel.textContent = label;
  }

  function clearResult(): void {
    if (resultUrl) URL.revokeObjectURL(resultUrl);
    resultUrl = undefined;
    results.replaceChildren();
    results.hidden = true;
  }

  function clearSelection(): void {
    selectedFiles = [];
    fileInput.value = "";
    fileList.replaceChildren();
    clearResult();
    setStatus("");
    setProgress(0, "Files stay on this device");
    renderFileList();
  }

  function requireElement<T extends Element>(selector: string): T {
    const element = document.querySelector<T>(selector);
    if (!element) throw new Error(`The workspace is missing ${selector}.`);
    return element;
  }

  function renderFileList(): void {
    fileList.replaceChildren();
    if (selectedFiles.length === 0) return;

    for (const [index, file] of selectedFiles.entries()) {
      const row = document.createElement("li");
      const name = document.createElement("span");
      const size = document.createElement("span");
      const actions = document.createElement("span");
      name.className = "file-name";
      name.textContent = file.name;
      size.className = "file-size";
      size.textContent = formatFileSize(file.size);
      actions.className = "file-actions";
      if (activeOperation === "merge-pdf") {
        for (const [action, label, disabled] of [
          ["up", "Move file up", index === 0],
          ["down", "Move file down", index === selectedFiles.length - 1],
        ] as const) {
          const button = document.createElement("button");
          button.type = "button";
          button.dataset.fileAction = action;
          button.dataset.fileIndex = String(index);
          button.disabled = disabled;
          button.setAttribute("aria-label", `${label}: ${file.name}`);
          button.textContent = action === "up" ? "↑" : "↓";
          actions.append(button);
        }
      }
      const removeButton = document.createElement("button");
      removeButton.type = "button";
      removeButton.dataset.fileAction = "remove";
      removeButton.dataset.fileIndex = String(index);
      removeButton.setAttribute("aria-label", `Remove ${file.name}`);
      removeButton.textContent = "×";
      actions.append(removeButton);
      row.append(name, size, actions);
      if (activeOperation === "merge-pdf") {
        const rowHandle = document.createElement("span");
        rowHandle.className = "file-order-hint";
        rowHandle.textContent = String(index + 1).padStart(2, "0");
        row.prepend(rowHandle);
      }
      fileList.append(row);
    }

    const numberOfFiles = selectedFiles.length;
    if (activeOperation === "merge-pdf" && numberOfFiles > 0) {
      setStatus(
        numberOfFiles > 1
          ? `${numberOfFiles} PDFs selected. Use the arrow buttons to arrange page order.`
          : "Add at least one more PDF to merge.",
      );
    }
  }

  function renderOperation(operation: OperationId): void {
    activeOperation = operation;
    const config = activeConfig();
    workspaceTitle.textContent = config.title;
    workspaceDescription.textContent = config.description;
    conversionDetail.textContent = config.detail;
    inputLabel.textContent = config.inputFormats.join(", ");
    dropTitle.textContent = `Add ${config.inputFormats.join(", ")} files`;
    dropHelp.textContent =
      operation === "merge-pdf"
        ? `Choose 2–10 PDFs; arrange them in the order you want below. ${formatFileSize(MAX_TOTAL_FILE_SIZE)} total.`
        : `Up to ${config.maxFiles} file${config.maxFiles === 1 ? "" : "s"} · ${formatFileSize(FILE_LIMIT)} per file, ${formatFileSize(MAX_TOTAL_FILE_SIZE)} total`;
    fileInput.accept = config.inputFormats.join(",");
    fileInput.multiple = config.maxFiles > 1;
    rangeOptions.hidden = operation !== "split-pdf";
    compressionOptions.hidden = operation !== "compress-pdf";
    compressionWarning.checked = false;
    outputFormat.textContent = config.outputFormat;
    uploadAction.textContent =
      config.maxFiles === 1 ? "Choose file" : "Choose files";
    toolButtons.forEach((button) => {
      const selected = button.dataset.tool === operation;
      button.setAttribute("aria-pressed", String(selected));
      button.classList.toggle("selected", selected);
    });
    clearSelection();
    runButton.textContent = `Convert to ${config.outputFormat}`;
  }

  function chooseFiles(incomingFiles: FileList | File[]): void {
    if (isConverting) return;
    clearResult();
    setStatus("");
    const files = Array.from(incomingFiles);
    const error = validateInputFiles(activeOperation, files);
    fileInput.value = "";
    if (error) {
      selectedFiles = [];
      fileInput.value = "";
      renderFileList();
      setStatus(error, true);
      return;
    }
    selectedFiles = files;
    renderFileList();
    if (activeOperation !== "merge-pdf") {
      setStatus(
        `${files.length} file${files.length === 1 ? "" : "s"} ready. Select Convert when you’re ready.`,
      );
    }
  }

  function downloadBlob(blob: Blob, fileName: string): HTMLAnchorElement {
    if (resultUrl) URL.revokeObjectURL(resultUrl);
    resultUrl = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.className = "download-button";
    link.href = resultUrl;
    link.download = fileName;
    link.textContent = `Download ${TOOLS[activeOperation].outputFormat}`;
    return link;
  }

  function finish(blob: Blob, fileName: string, detail?: string): void {
    const container = document.createElement("div");
    container.className = "result-card";
    const label = document.createElement("span");
    label.className = "result-check";
    label.setAttribute("aria-hidden", "true");
    label.textContent = "✓";
    const description = document.createElement("div");
    description.className = "result-copy";
    const heading = document.createElement("strong");
    heading.textContent = "Your file is ready";
    const filename = document.createElement("span");
    filename.textContent = `${fileName} · ${formatFileSize(blob.size)}`;
    description.append(heading, filename);
    if (detail) {
      const note = document.createElement("p");
      note.textContent = detail;
      description.append(note);
    }
    container.append(label, description, downloadBlob(blob, fileName));
    results.replaceChildren(container);
    results.hidden = false;
    results.scrollIntoView({ behavior: "smooth", block: "nearest" });
    setStatus("Conversion complete. Your original file has not been changed.");
    setProgress(100, "Conversion complete");
  }

  async function loadPdf(
    file: File,
    maxPages = MAX_PDF_PAGES,
  ): Promise<{ pdf: PdfDocumentProxy; close: () => Promise<void> }> {
    if (file.size > FILE_LIMIT)
      throw new Error(`${file.name} is larger than the 30 MB limit.`);
    const bytes = new Uint8Array(await file.arrayBuffer());
    const pdfjs = await getPdfjs();
    const loadingTask = pdfjs.getDocument({ data: bytes });
    const pdf = await loadingTask.promise;
    if (pdf.numPages > maxPages) {
      await loadingTask.destroy();
      throw new Error(
        `This tool supports up to ${maxPages} pages at a time; this PDF has ${pdf.numPages}.`,
      );
    }
    return { pdf, close: () => loadingTask.destroy() };
  }

  function createPageCanvas(viewport: import("pdfjs-dist").PageViewport): {
    canvas: HTMLCanvasElement;
    context: CanvasRenderingContext2D;
  } {
    const width = Math.ceil(viewport.width);
    const height = Math.ceil(viewport.height);
    if (width > 10_000 || height > 10_000 || width * height > 30_000_000) {
      throw new Error(
        "A page exceeds the safe rendering size for this browser.",
      );
    }

    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d", { alpha: false });
    if (!context)
      throw new Error("Your browser couldn’t create a page canvas.");
    context.fillStyle = "#fff";
    context.fillRect(0, 0, width, height);
    return { canvas, context };
  }

  async function extractPdfPages(
    file: File,
    maxPages: number,
    report: (progressValue: number, label: string) => void,
  ): Promise<{
    pages: string[];
    blocks: DocumentBlock[][];
    images: ExtractedImage[];
  }> {
    const { pdf, close } = await loadPdf(file, maxPages);
    try {
      const pages: string[] = [];
      const images: ExtractedImage[] = [];
      for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
        report(
          (pageNumber / pdf.numPages) * 65,
          `Reading page ${pageNumber} of ${pdf.numPages}`,
        );
        const page = await pdf.getPage(pageNumber);
        const content = await page.getTextContent();
        const text = extractPdfPageText(
          content.items.filter(
            (item): item is typeof item & { str: string } => "str" in item,
          ),
        );
        pages.push(text);
        report(
          65 + (pageNumber / pdf.numPages) * 25,
          `Preserving the original page ${pageNumber} of ${pdf.numPages}`,
        );
        const viewport = page.getViewport({ scale: 1.3 });
        const { canvas, context } = createPageCanvas(viewport);
        report(
          65 + (pageNumber / pdf.numPages) * 25,
          `Rendering page ${pageNumber} of ${pdf.numPages}`,
        );
        await page.render({ canvas, canvasContext: context, viewport }).promise;
        report(
          65 + (pageNumber / pdf.numPages) * 25,
          `Encoding page image ${pageNumber} of ${pdf.numPages}`,
        );
        const imageBlob = await canvasToBlob(canvas, "image/jpeg", 0.84);
        const imageData = new Uint8Array(await imageBlob.arrayBuffer());
        images.push({
          data: imageData,
          contentType: "image/jpeg",
          width: canvas.width,
          height: canvas.height,
        });
        canvas.width = 0;
        canvas.height = 0;
        const imageSize = images.reduce(
          (total, image) => total + image.data.byteLength,
          0,
        );
        if (imageSize > 80 * 1024 * 1024) {
          throw new Error(
            "The rendered PDF pages exceed the 80 MB visual-preservation limit. Split the document into smaller files and try again.",
          );
        }
      }
      return {
        pages,
        blocks: pages.map((text) =>
          structurePlainText(
            text || "(No selectable text found on this page.)",
          ),
        ),
        images,
      };
    } finally {
      await close();
    }
  }

  async function extractSlides(file: File): Promise<{
    slides: string[];
    blocks: DocumentBlock[][];
    images: ExtractedImage[];
  }> {
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(await file.arrayBuffer(), {
      checkCRC32: true,
    });
    const presentationText = await zip
      .file("ppt/presentation.xml")
      ?.async("string");
    const relationshipsText = await zip
      .file("ppt/_rels/presentation.xml.rels")
      ?.async("string");
    if (!presentationText || !relationshipsText) {
      throw new Error(
        `${file.name} is missing the standard PowerPoint slide order.`,
      );
    }
    const presentationXml = new DOMParser().parseFromString(
      presentationText,
      "application/xml",
    );
    const relationshipsXml = new DOMParser().parseFromString(
      relationshipsText,
      "application/xml",
    );
    if (
      presentationXml.querySelector("parsererror") ||
      relationshipsXml.querySelector("parsererror")
    ) {
      throw new Error(`${file.name} has an unreadable slide list.`);
    }
    const relationships = new Map<string, string>();
    for (const relationship of Array.from(
      relationshipsXml.getElementsByTagNameNS("*", "Relationship"),
    )) {
      const id = relationship.getAttribute("Id");
      const target = relationship.getAttribute("Target");
      if (id && target) relationships.set(id, target);
    }
    const rawRelationshipIds = Array.from(
      presentationXml.getElementsByTagNameNS("*", "sldId"),
    ).map((slide) =>
      slide.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id",
      ),
    );
    const relationshipIds = rawRelationshipIds.filter(
      (id): id is string => id !== null,
    );
    if (relationshipIds.length !== rawRelationshipIds.length) {
      throw new Error(`${file.name} contains an invalid slide reference.`);
    }
    const entries = resolvePresentationSlidePaths(
      relationshipIds,
      relationships,
      new Set(Object.keys(zip.files)),
    );

    const slides: string[] = [];
    const slideBlocks: DocumentBlock[][] = [];
    const images: ExtractedImage[] = [];
    const includedImages = new Set<string>();
    for (const [index, path] of entries.entries()) {
      const xmlText = await zip.file(path)?.async("string");
      if (!xmlText)
        throw new Error(`Couldn’t read slide ${index + 1} from ${file.name}.`);
      const xml = new DOMParser().parseFromString(xmlText, "application/xml");
      if (xml.querySelector("parsererror"))
        throw new Error(`${file.name} has an unreadable slide.`);
      const blocks: DocumentBlock[] = [];
      const handledTables = new Set<Element>();
      const paragraphs = Array.from(xml.getElementsByTagNameNS("*", "p"));
      for (const paragraph of paragraphs) {
        let parent = paragraph.parentNode;
        let table: Element | undefined;
        while (parent && parent.nodeType === 1) {
          const element = parent as Element;
          if (element.localName === "tbl") {
            table = element;
            break;
          }
          parent = parent.parentNode;
        }
        if (table) {
          if (handledTables.has(table)) continue;
          handledTables.add(table);
          const rows = Array.from(table.getElementsByTagNameNS("*", "tr"))
            .map((row) =>
              Array.from(row.getElementsByTagNameNS("*", "tc")).map((cell) =>
                Array.from(cell.getElementsByTagNameNS("*", "p"))
                  .map((cellParagraph) =>
                    Array.from(cellParagraph.getElementsByTagNameNS("*", "t"))
                      .map((run) => run.textContent ?? "")
                      .join(""),
                  )
                  .join("\n"),
              ),
            )
            .filter((row) => row.length > 0);
          if (rows.length > 0) {
            blocks.push({
              type: "table",
              rows,
              hasHeader:
                table
                  .getElementsByTagNameNS("*", "tblPr")[0]
                  ?.getAttribute("firstRow") === "1",
            });
          }
          continue;
        }

        const text = Array.from(paragraph.getElementsByTagNameNS("*", "t"))
          .map((run) => run.textContent ?? "")
          .join("")
          .trim();
        if (!text) continue;
        const paragraphProperties = paragraph.getElementsByTagNameNS(
          "*",
          "pPr",
        )[0];
        const isBullet =
          (paragraphProperties?.getElementsByTagNameNS("*", "buChar").length ??
            0) > 0;
        const isNumbered =
          (paragraphProperties?.getElementsByTagNameNS("*", "buAutoNum")
            .length ?? 0) > 0;
        if (isBullet || isNumbered) {
          blocks.push({
            type: "list-item",
            text,
            level: Math.min(
              Math.max(
                Number(paragraphProperties?.getAttribute("lvl") ?? 0),
                0,
              ),
              5,
            ),
            ordered: isNumbered,
          });
          continue;
        }

        let placeholderType: string | null = null;
        let shape: Node | null = paragraph.parentNode;
        while (shape && shape.nodeType === 1) {
          const element = shape as Element;
          const placeholder = element.getElementsByTagNameNS("*", "ph")[0];
          if (placeholder) {
            placeholderType = placeholder.getAttribute("type");
            break;
          }
          shape = shape.parentNode;
        }
        if (placeholderType === "title" || placeholderType === "ctrTitle") {
          blocks.push({ type: "heading", text, level: 1 });
        } else if (placeholderType === "subTitle") {
          blocks.push({ type: "heading", text, level: 2 });
        } else {
          blocks.push({ type: "paragraph", text });
        }
      }
      const slideText = blocks
        .map((block) =>
          block.type === "table"
            ? block.rows.map((row) => row.join(" | ")).join("\n")
            : block.type === "list-item"
              ? `${"  ".repeat(block.level)}${block.ordered ? "1." : "•"} ${block.text}`
              : block.text,
        )
        .join("\n");
      const slideTextParts = slideText.trim() ? [slideText.trim()] : [];
      const slideName = path.split("/").pop();
      const slideRelationships = slideName
        ? await zip.file(`ppt/slides/_rels/${slideName}.rels`)?.async("string")
        : undefined;
      if (!slideRelationships) {
        slides.push(slideTextParts.join("\n\n"));
        slideBlocks.push(blocks);
        continue;
      }
      const relationshipsXml = new DOMParser().parseFromString(
        slideRelationships,
        "application/xml",
      );
      if (relationshipsXml.querySelector("parsererror")) {
        throw new Error(
          `${file.name} has unreadable slide content references.`,
        );
      }
      for (const relationship of Array.from(
        relationshipsXml.getElementsByTagNameNS("*", "Relationship"),
      )) {
        const target = relationship.getAttribute("Target");
        const type = relationship.getAttribute("Type");
        if (!target || !type) continue;
        const isImage = type.endsWith("/image");
        const isChart = type.endsWith("/chart");
        const isNotes = type.endsWith("/notesSlide");
        if (!isImage && !isChart && !isNotes) continue;
        if (
          target.startsWith("/") ||
          target.includes("\\") ||
          /^[a-z]+:/i.test(target)
        ) {
          throw new Error(
            `${file.name} contains an invalid content reference.`,
          );
        }
        const relatedPath = new URL(
          target,
          `https://scriptorium.invalid/${path}`,
        ).pathname.slice(1);
        if (isChart) {
          if (!relatedPath.startsWith("ppt/charts/")) {
            throw new Error(
              `${file.name} contains an invalid chart reference.`,
            );
          }
          const chartText = await zip.file(relatedPath)?.async("string");
          if (!chartText) {
            throw new Error(`${file.name} is missing embedded chart data.`);
          }
          const chartXml = new DOMParser().parseFromString(
            chartText,
            "application/xml",
          );
          if (chartXml.querySelector("parsererror")) {
            throw new Error(`${file.name} contains unreadable chart data.`);
          }
          const values = Array.from(chartXml.getElementsByTagNameNS("*", "v"))
            .map((value) => value.textContent?.trim() ?? "")
            .filter(Boolean);
          if (values.length > 0) {
            const chartData = `Chart data: ${values.join(" | ")}`;
            slideTextParts.push(chartData);
            blocks.push({ type: "paragraph", text: chartData });
          }
          continue;
        }
        if (isNotes) {
          if (!relatedPath.startsWith("ppt/notesSlides/")) {
            throw new Error(
              `${file.name} contains an invalid speaker-note reference.`,
            );
          }
          const notesText = await zip.file(relatedPath)?.async("string");
          if (!notesText) {
            throw new Error(`${file.name} is missing its speaker notes.`);
          }
          const notesXml = new DOMParser().parseFromString(
            notesText,
            "application/xml",
          );
          if (notesXml.querySelector("parsererror")) {
            throw new Error(`${file.name} contains unreadable speaker notes.`);
          }
          const notes = Array.from(notesXml.getElementsByTagNameNS("*", "p"))
            .filter((paragraph) => {
              const placeholder = paragraph.getElementsByTagNameNS(
                "*",
                "ph",
              )[0];
              return !["sldNum", "dt", "hdr", "ftr"].includes(
                placeholder?.getAttribute("type") ?? "",
              );
            })
            .map((paragraph) =>
              Array.from(paragraph.getElementsByTagNameNS("*", "t"))
                .map((run) => run.textContent ?? "")
                .join(""),
            )
            .filter((paragraph) => paragraph.trim());
          if (notes.length > 0) {
            slideTextParts.push(`Speaker notes:\n${notes.join("\n")}`);
            blocks.push({
              type: "heading",
              text: "Speaker notes",
              level: 2,
            });
            blocks.push(
              ...notes.map((text) => ({ type: "paragraph" as const, text })),
            );
          }
          continue;
        }
        if (!isImage) continue;
        const imagePath = relatedPath;
        if (!imagePath.startsWith("ppt/media/")) {
          throw new Error(`${file.name} contains an invalid image reference.`);
        }
        if (includedImages.has(imagePath)) continue;
        includedImages.add(imagePath);
        const extension = imagePath.split(".").pop()?.toLowerCase();
        const contentType =
          extension === "jpg" || extension === "jpeg"
            ? "image/jpeg"
            : extension === "png"
              ? "image/png"
              : extension === "gif"
                ? "image/gif"
                : extension === "bmp"
                  ? "image/bmp"
                  : extension === "svg"
                    ? "image/svg+xml"
                    : extension === "webp"
                      ? "image/webp"
                      : undefined;
        if (!contentType) {
          throw new Error(
            `${file.name} contains an embedded image format that cannot be preserved. Convert slide images to PNG or JPEG and try again.`,
          );
        }
        const imageFile = zip.file(imagePath);
        if (!imageFile)
          throw new Error(`${file.name} is missing an embedded slide image.`);
        const data = await imageFile.async("uint8array");
        images.push(await normalizeEmbeddedImage(data, contentType, file.name));
        const imageSize = images.reduce(
          (total, image) => total + image.data.byteLength,
          0,
        );
        if (imageSize > 80 * 1024 * 1024) {
          throw new Error(
            "The embedded PowerPoint images exceed the 80 MB preservation limit.",
          );
        }
      }
      slides.push(slideTextParts.join("\n\n"));
      slideBlocks.push(blocks);
    }
    if (!slides.some((slide) => slide.length > 0) && images.length === 0) {
      throw new Error(`${file.name} contains no extractable slide text.`);
    }
    return {
      slides: slides.map(
        (slide) =>
          slide || "(No selectable text on this slide. See retained images.)",
      ),
      blocks: slideBlocks.map((blocks) =>
        blocks.length > 0
          ? blocks
          : [
              {
                type: "paragraph",
                text: "(No selectable text on this slide. See retained images.)",
              },
            ],
      ),
      images,
    };
  }

  async function createTextPdf(
    title: string,
    sections: Array<{
      heading?: string;
      text: string;
      blocks?: DocumentBlock[];
    }>,
    images: ExtractedImage[] = [],
  ): Promise<Blob> {
    const { PDFDocument, PageSizes, StandardFonts, rgb } =
      await import("pdf-lib");
    const pdf = await PDFDocument.create();
    const titleFont = await pdf.embedFont(StandardFonts.HelveticaBold);
    const headingFont = await pdf.embedFont(StandardFonts.HelveticaBold);
    const bodyFont = await pdf.embedFont(StandardFonts.Helvetica);
    const [pageWidth, pageHeight] = PageSizes.A4;
    const margin = 52;
    const usableWidth = pageWidth - margin * 2;
    let page = pdf.addPage(PageSizes.A4);
    let cursor = pageHeight - margin;
    const pages = [page];

    const validateFontSupport = (text: string, font: typeof bodyFont): void => {
      try {
        font.encodeText(text);
      } catch {
        throw new Error(
          "This PDF export can’t render a character or writing system in this document with its built-in fonts. Try Word to PowerPoint, or export the Word file directly.",
        );
      }
    };

    const ensureSpace = (height: number): void => {
      if (cursor - height < margin) {
        page = pdf.addPage(PageSizes.A4);
        pages.push(page);
        cursor = pageHeight - margin;
      }
    };

    const drawWrappedText = (
      text: string,
      font: typeof bodyFont,
      size: number,
      color: ReturnType<typeof rgb>,
      lineHeight: number,
      afterParagraph = 0,
    ): void => {
      for (const paragraph of text.split(/\n/).filter((line) => line.trim())) {
        validateFontSupport(paragraph, font);
        const words = paragraph.trim().split(/\s+/);
        let line = "";
        for (const word of words) {
          if (font.widthOfTextAtSize(word, size) > usableWidth) {
            if (line) {
              ensureSpace(lineHeight);
              page.drawText(line, {
                x: margin,
                y: cursor - lineHeight,
                font,
                size,
                color,
              });
              cursor -= lineHeight;
            }
            let segment = "";
            for (const character of Array.from(word)) {
              const nextSegment = segment + character;
              if (
                segment &&
                font.widthOfTextAtSize(nextSegment, size) > usableWidth
              ) {
                ensureSpace(lineHeight);
                page.drawText(segment, {
                  x: margin,
                  y: cursor - lineHeight,
                  font,
                  size,
                  color,
                });
                cursor -= lineHeight;
                segment = character;
              } else {
                segment = nextSegment;
              }
            }
            line = segment;
            continue;
          }

          const nextLine = line ? `${line} ${word}` : word;
          if (line && font.widthOfTextAtSize(nextLine, size) > usableWidth) {
            ensureSpace(lineHeight);
            page.drawText(line, {
              x: margin,
              y: cursor - lineHeight,
              font,
              size,
              color,
            });
            cursor -= lineHeight;
            line = word;
          } else {
            line = nextLine;
          }
        }

        if (line) {
          ensureSpace(lineHeight);
          page.drawText(line, {
            x: margin,
            y: cursor - lineHeight,
            font,
            size,
            color,
          });
          cursor -= lineHeight;
        }
        cursor -= afterParagraph;
      }
    };

    const drawTable = (rows: string[][], hasHeader: boolean): void => {
      const columnCount = Math.max(...rows.map((row) => row.length));
      const cellWidth = usableWidth / columnCount;
      const padding = 5;
      const lineHeight = 12;
      const wrapCell = (value: string): string[] => {
        const lines: string[] = [];
        let line = "";
        for (const word of value.split(/\s+/).filter(Boolean)) {
          let segment = word;
          while (
            bodyFont.widthOfTextAtSize(segment, 8.5) >
            cellWidth - 2 * padding
          ) {
            let prefix = "";
            for (const character of Array.from(segment)) {
              const next = prefix + character;
              if (
                prefix &&
                bodyFont.widthOfTextAtSize(next, 8.5) > cellWidth - 2 * padding
              ) {
                break;
              }
              prefix = next;
            }
            if (!prefix) break;
            lines.push(prefix);
            segment = segment.slice(prefix.length);
          }
          const nextLine = line ? `${line} ${segment}` : segment;
          if (
            line &&
            bodyFont.widthOfTextAtSize(nextLine, 8.5) > cellWidth - 2 * padding
          ) {
            lines.push(line);
            line = segment;
          } else {
            line = nextLine;
          }
        }
        if (line) lines.push(line);
        return lines.length > 0 ? lines : [""];
      };

      for (const [rowIndex, row] of rows.entries()) {
        const cellLines = Array.from({ length: columnCount }, (_, column) =>
          wrapCell(row[column] ?? ""),
        );
        const rowHeight =
          Math.max(...cellLines.map((lines) => lines.length)) * lineHeight +
          padding * 2;
        if (rowHeight > pageHeight - margin * 2) {
          throw new Error(
            "A table row is too large to fit safely in the generated PDF.",
          );
        }
        ensureSpace(rowHeight);
        const top = cursor;
        for (let column = 0; column < columnCount; column += 1) {
          const x = margin + column * cellWidth;
          page.drawRectangle({
            x,
            y: top - rowHeight,
            width: cellWidth,
            height: rowHeight,
            borderColor: rgb(0.82, 0.86, 0.82),
            borderWidth: 0.5,
            color:
              hasHeader && rowIndex === 0
                ? rgb(0.92, 0.96, 0.92)
                : rgb(1, 1, 1),
          });
          cellLines[column].forEach((line, lineIndex) => {
            page.drawText(line, {
              x: x + padding,
              y: top - padding - 8.5 - lineIndex * lineHeight,
              font: hasHeader && rowIndex === 0 ? headingFont : bodyFont,
              size: 8.5,
              color: rgb(0.2, 0.22, 0.2),
            });
          });
        }
        cursor -= rowHeight;
      }
      cursor -= 8;
    };

    try {
      drawWrappedText(title, titleFont, 24, rgb(0.13, 0.23, 0.18), 32, 16);
      for (const section of sections) {
        if (section.heading) {
          drawWrappedText(
            section.heading,
            headingFont,
            15,
            rgb(0.23, 0.41, 0.29),
            20,
            5,
          );
        }
        if (section.blocks) {
          const orderedIndices = new Map<number, number>();
          for (const block of section.blocks) {
            if (block.type === "table") {
              orderedIndices.clear();
              drawTable(block.rows, block.hasHeader);
            } else if (block.type === "heading") {
              orderedIndices.clear();
              drawWrappedText(
                block.text,
                headingFont,
                block.level < 3 ? 14 : 12,
                rgb(0.23, 0.41, 0.29),
                block.level < 3 ? 19 : 16,
                5,
              );
            } else if (block.type === "list-item") {
              const level = Math.min(block.level, 4);
              const orderedIndex = block.ordered
                ? (orderedIndices.get(level) ?? 0) + 1
                : 0;
              if (block.ordered) orderedIndices.set(level, orderedIndex);
              else orderedIndices.delete(level);
              const prefix = block.ordered ? `${orderedIndex}.` : "•";
              drawWrappedText(
                `${"  ".repeat(level)}${prefix} ${block.text}`,
                bodyFont,
                10.5,
                rgb(0.2, 0.22, 0.2),
                15,
                3,
              );
            } else {
              orderedIndices.clear();
              drawWrappedText(
                block.text,
                bodyFont,
                10.5,
                rgb(0.2, 0.22, 0.2),
                15,
                5,
              );
            }
          }
        } else {
          drawWrappedText(
            section.text,
            bodyFont,
            10.5,
            rgb(0.2, 0.22, 0.2),
            15,
            5,
          );
        }
      }
    } catch (error) {
      if (error instanceof Error && error.message.startsWith("This PDF export"))
        throw error;
      if (error instanceof Error && error.message.includes("WinAnsi")) {
        throw new Error(
          "This PDF export can’t render a character or writing system in this document with its built-in fonts. Try Word to PowerPoint, or export the Word file directly.",
          { cause: error },
        );
      }
      throw error;
    }

    for (const [index, image] of images.entries()) {
      let embedded;
      if (image.contentType === "image/png") {
        embedded = await pdf.embedPng(image.data);
      } else if (image.contentType === "image/jpeg") {
        embedded = await pdf.embedJpg(image.data);
      } else {
        const bitmap = await createImageBitmap(
          new Blob([imageDataBuffer(image.data)], {
            type: image.contentType,
          }),
        );
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
        const png = await canvasToBlob(canvas);
        embedded = await pdf.embedPng(await png.arrayBuffer());
        canvas.width = 0;
        canvas.height = 0;
      }

      const imagePage = pdf.addPage(PageSizes.A4);
      pages.push(imagePage);
      const maxWidth = pageWidth - margin * 2;
      const maxHeight = pageHeight - margin * 2 - 24;
      const scale = Math.min(
        maxWidth / embedded.width,
        maxHeight / embedded.height,
      );
      imagePage.drawText(`Original embedded image ${index + 1}`, {
        x: margin,
        y: pageHeight - margin,
        size: 12,
        font: headingFont,
        color: rgb(0.23, 0.41, 0.29),
      });
      imagePage.drawImage(embedded, {
        x: (pageWidth - embedded.width * scale) / 2,
        y: (pageHeight - 24 - embedded.height * scale) / 2,
        width: embedded.width * scale,
        height: embedded.height * scale,
      });
    }

    pages.forEach((pageItem, index) => {
      pageItem.drawText("SCRIPTORIUM  /  DOCUMENT CONVERSION", {
        x: margin,
        y: pageHeight - 27,
        size: 7,
        font: headingFont,
        color: rgb(0.35, 0.47, 0.38),
      });
      pageItem.drawLine({
        start: { x: margin, y: pageHeight - 37 },
        end: { x: pageWidth - margin, y: pageHeight - 37 },
        thickness: 0.8,
        color: rgb(0.82, 0.86, 0.82),
      });
      pageItem.drawLine({
        start: { x: margin, y: 33 },
        end: { x: pageWidth - margin, y: 33 },
        thickness: 0.6,
        color: rgb(0.86, 0.89, 0.86),
      });
      pageItem.drawText(`SCRIPTORIUM  ·  ${index + 1} / ${pages.length}`, {
        x: margin,
        y: 20,
        size: 7,
        font: bodyFont,
        color: rgb(0.42, 0.49, 0.44),
      });
    });

    return new Blob([Uint8Array.from(await pdf.save()).buffer], {
      type: "application/pdf",
    });
  }

  async function savePresentation(
    title: string,
    slides: string[],
    images: ExtractedImage[] = [],
    useImageSlidesWithNotes = false,
    contentSlides?: DocumentBlock[][],
    pageTextGroups?: string[][],
  ): Promise<Blob> {
    if (
      useImageSlidesWithNotes &&
      (!pageTextGroups ||
        slides.length !== images.length ||
        pageTextGroups.length !== images.length ||
        pageTextGroups.some((group) => group.length === 0))
    ) {
      throw new Error(
        "The original page images and extracted text could not be matched.",
      );
    }
    const textSlideCount = contentSlides?.length ?? slides.length;
    const outputSlideCount = useImageSlidesWithNotes
      ? (pageTextGroups?.reduce((total, group) => total + group.length, 0) ??
        images.length)
      : textSlideCount + images.length;
    if (outputSlideCount > MAX_POWERPOINT_SLIDES) {
      throw new Error(
        `This conversion would create more than ${MAX_POWERPOINT_SLIDES} slides when preserving the extracted text and images. Convert fewer pages or images at a time.`,
      );
    }
    const { default: PptxGenJS } = await import("pptxgenjs");
    const presentation = new PptxGenJS();
    type PresentationSlide = ReturnType<
      InstanceType<typeof PptxGenJS>["addSlide"]
    >;
    presentation.layout = "LAYOUT_WIDE";
    presentation.author = "Scriptorium";
    presentation.subject = "Text extracted and reorganised in Scriptorium";
    presentation.title = title;

    const addStructuredText = (
      slide: PresentationSlide,
      text: string,
      frame: {
        x: number;
        y: number;
        w: number;
        h: number;
        fontSize: number;
        charsPerLine: number;
      },
    ): void => {
      let cursor = frame.y;
      const orderedIndices = new Map<number, number>();
      for (const block of structurePlainText(text)) {
        const isHeading = block.type === "heading";
        const isListItem = block.type === "list-item";
        const level = isListItem ? Math.min(block.level, 4) : 0;
        const content =
          isListItem && block.ordered
            ? `${(orderedIndices.get(level) ?? 0) + 1}. ${block.text}`
            : block.text;
        if (isListItem && block.ordered) {
          orderedIndices.set(level, (orderedIndices.get(level) ?? 0) + 1);
        } else if (isListItem) {
          orderedIndices.delete(level);
        } else {
          orderedIndices.clear();
        }

        const lines = Math.max(
          1,
          Math.ceil(content.length / frame.charsPerLine),
        );
        const fontSize = isHeading ? frame.fontSize + 2 : frame.fontSize;
        const height = isHeading
          ? 0.32 * lines
          : Math.min(1.3, 0.16 + (lines * fontSize * 1.35) / 72);
        const indent = level * 0.18;
        if (cursor + height > frame.y + frame.h) {
          throw new Error(
            "The extracted PDF text does not fit in the available presentation layout.",
          );
        }
        slide.addText(content, {
          x: frame.x + indent,
          y: cursor,
          w: frame.w - indent,
          h: height,
          fontFace: isHeading ? "Aptos Display" : "Aptos",
          fontSize,
          bold: isHeading,
          color: isHeading ? "2F3B2E" : "33372F",
          bullet:
            isListItem && !block.ordered
              ? { indent: 18 + level * 18 }
              : undefined,
          margin: 0,
          breakLine: false,
          fit: "shrink",
        });
        cursor += height + (isHeading ? 0.05 : 0.04);
      }
    };

    const textSlides = useImageSlidesWithNotes || contentSlides ? [] : slides;
    textSlides.forEach((text, index) => {
      const slide = presentation.addSlide();
      slide.background = { color: "F7F6F2" };
      slide.addText(title, {
        x: 0.7,
        y: 0.35,
        w: 11.9,
        h: 0.3,
        fontFace: "Aptos",
        fontSize: 9,
        color: "75816D",
        breakLine: false,
        margin: 0,
      });
      slide.addShape(presentation.ShapeType.line, {
        x: 0.7,
        y: 0.85,
        w: 11.9,
        h: 0,
        line: { color: "DEDFD8", width: 1 },
      });
      slide.addText(`Page ${index + 1}`, {
        x: 0.7,
        y: 0.98,
        w: 11.9,
        h: 0.45,
        fontFace: "Aptos Display",
        fontSize: 22,
        bold: true,
        color: "2F3B2E",
        margin: 0,
        breakLine: false,
      });
      slide.addText(text, {
        x: 0.7,
        y: 1.62,
        w: 11.9,
        h: 5.4,
        fontFace: "Aptos",
        fontSize: 16,
        color: "33372F",
        valign: "top",
        breakLine: false,
        paraSpaceAfter: 12,
        margin: 0,
        fit: "shrink",
      });
      slide.addText(`SCRIPTORIUM  ·  TEXT-ONLY CONVERSION`, {
        x: 0.7,
        y: 7.12,
        w: 11.9,
        h: 0.2,
        fontFace: "Aptos",
        fontSize: 8,
        charSpacing: 0.7,
        color: "8B8D83",
        margin: 0,
        breakLine: false,
      });
    });
    contentSlides?.forEach((blocks, index) => {
      const slide = presentation.addSlide();
      slide.background = { color: "F7F6F2" };
      slide.addText(title, {
        x: 0.7,
        y: 0.35,
        w: 11.9,
        h: 0.3,
        fontFace: "Aptos",
        fontSize: 9,
        color: "75816D",
        margin: 0,
      });
      slide.addShape(presentation.ShapeType.line, {
        x: 0.7,
        y: 0.85,
        w: 11.9,
        h: 0,
        line: { color: "DEDFD8", width: 1 },
      });

      let cursor = 1.05;
      const orderedIndices = new Map<number, number>();
      for (const block of blocks) {
        if (block.type === "table") {
          orderedIndices.clear();
          const tableRows = block.rows.map((row, rowIndex) =>
            row.map((text) => ({
              text,
              options: {
                bold: block.hasHeader && rowIndex === 0,
                color: "33372F",
                fill: {
                  color:
                    block.hasHeader && rowIndex === 0 ? "EAF2EB" : "FFFFFF",
                },
              },
            })),
          );
          slide.addTable(tableRows, {
            x: 0.8,
            y: cursor,
            w: 11.7,
            fontFace: "Aptos",
            fontSize: 12,
            color: "33372F",
            border: { type: "solid", color: "D8E1D7", pt: 0.75 },
            margin: 0.08,
          });
          cursor += Math.max(0.55, block.rows.length * 0.4) + 0.15;
          continue;
        }

        const isHeading = block.type === "heading";
        const isListItem = block.type === "list-item";
        const level = isListItem ? Math.min(block.level, 4) : 0;
        const text =
          isListItem && block.ordered
            ? `${(orderedIndices.get(level) ?? 0) + 1}. ${block.text}`
            : block.text;
        if (isListItem && block.ordered) {
          orderedIndices.set(level, (orderedIndices.get(level) ?? 0) + 1);
        } else if (isListItem) {
          orderedIndices.delete(level);
        } else {
          orderedIndices.clear();
        }
        const lines = Math.max(1, Math.ceil(block.text.length / 105));
        const height = isHeading ? 0.42 : Math.min(1.3, 0.2 + lines * 0.25);
        const indent = level * 0.24;
        slide.addText(text, {
          x: 0.85 + indent,
          y: cursor,
          w: 11.6 - indent,
          h: height,
          fontFace: isHeading ? "Aptos Display" : "Aptos",
          fontSize:
            block.type === "heading" ? (block.level === 1 ? 21 : 16) : 14,
          bold: isHeading,
          color: isHeading ? "2F3B2E" : "33372F",
          bullet:
            isListItem && !block.ordered
              ? { indent: 18 + level * 18 }
              : undefined,
          margin: 0,
          breakLine: false,
          fit: "shrink",
        });
        cursor += height + (isHeading ? 0.08 : 0.1);
      }

      slide.addText(`SCRIPTORIUM  ·  ${index + 1}`, {
        x: 0.7,
        y: 7.12,
        w: 11.9,
        h: 0.2,
        fontFace: "Aptos",
        fontSize: 8,
        charSpacing: 0.7,
        color: "8B8D83",
        margin: 0,
      });
    });
    for (const [index, image] of images.entries()) {
      const slide = presentation.addSlide();
      slide.background = { color: "F7F6F2" };
      slide.addText(
        `${useImageSlidesWithNotes ? "Original PDF page" : "Original embedded image"} ${index + 1}`,
        {
          x: 0.7,
          y: 0.35,
          w: 11.9,
          h: 0.4,
          fontFace: "Aptos Display",
          fontSize: 18,
          bold: true,
          color: "2F3B2E",
          margin: 0,
        },
      );
      let outputImage = image;
      if (
        image.contentType !== "image/png" &&
        image.contentType !== "image/jpeg"
      ) {
        const bitmap = await createImageBitmap(
          new Blob([imageDataBuffer(image.data)], {
            type: image.contentType,
          }),
        );
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
        const png = await canvasToBlob(canvas);
        outputImage = {
          data: new Uint8Array(await png.arrayBuffer()),
          contentType: "image/png",
          width: canvas.width,
          height: canvas.height,
        };
        canvas.width = 0;
        canvas.height = 0;
      }
      const encoded = outputImage.data.reduce((chunks, byte, byteIndex) => {
        if (byteIndex % 0x8000 === 0) chunks.push([]);
        chunks[chunks.length - 1].push(byte);
        return chunks;
      }, [] as number[][]);
      const base64 = btoa(
        encoded.map((chunk) => String.fromCharCode(...chunk)).join(""),
      );
      const data = `data:${outputImage.contentType};base64,${base64}`;
      const imageWidth = useImageSlidesWithNotes ? 4.7 : 11.9;
      const imageHeight = useImageSlidesWithNotes ? 5.7 : 6.1;
      const scale = Math.min(
        imageWidth / outputImage.width,
        imageHeight / outputImage.height,
      );
      const width = outputImage.width * scale;
      const height = outputImage.height * scale;
      slide.addImage({
        data,
        x: useImageSlidesWithNotes
          ? 0.7 + (imageWidth - width) / 2
          : (13.333 - width) / 2,
        y: 1.05 + (imageHeight - height) / 2,
        w: width,
        h: height,
      });
      if (useImageSlidesWithNotes && slides[index]) {
        const textSegments = pageTextGroups?.[index] ?? [slides[index]];
        slide.addText("Editable extracted text", {
          x: 5.7,
          y: 1.05,
          w: 6.1,
          h: 0.4,
          fontFace: "Aptos Display",
          fontSize: 17,
          bold: true,
          color: "2F3B2E",
          margin: 0,
        });
        addStructuredText(slide, textSegments[0], {
          x: 5.7,
          y: 1.55,
          w: 6.1,
          h: 5.2,
          fontSize: 12,
          charsPerLine: 62,
        });
        slide.addNotes(slides[index]);
        for (const [continuationIndex, text] of textSegments
          .slice(1)
          .entries()) {
          const continuation = presentation.addSlide();
          continuation.background = { color: "F7F6F2" };
          continuation.addText(title, {
            x: 0.7,
            y: 0.35,
            w: 11.9,
            h: 0.3,
            fontFace: "Aptos",
            fontSize: 9,
            color: "75816D",
            margin: 0,
          });
          continuation.addShape(presentation.ShapeType.line, {
            x: 0.7,
            y: 0.85,
            w: 11.9,
            h: 0,
            line: { color: "DEDFD8", width: 1 },
          });
          continuation.addText(
            `Page ${index + 1} · Continued ${continuationIndex + 1}`,
            {
              x: 0.7,
              y: 0.98,
              w: 11.9,
              h: 0.45,
              fontFace: "Aptos Display",
              fontSize: 22,
              bold: true,
              color: "2F3B2E",
              margin: 0,
            },
          );
          addStructuredText(continuation, text, {
            x: 0.7,
            y: 1.62,
            w: 11.9,
            h: 5.4,
            fontSize: 16,
            charsPerLine: 110,
          });
          continuation.addText("SCRIPTORIUM  ·  CONTINUED TEXT", {
            x: 0.7,
            y: 7.12,
            w: 11.9,
            h: 0.2,
            fontFace: "Aptos",
            fontSize: 8,
            charSpacing: 0.7,
            color: "8B8D83",
            margin: 0,
          });
        }
      }
    }
    const output = await presentation.write({
      compression: true,
      outputType: "blob",
    });
    if (!(output instanceof Blob))
      throw new Error("PowerPoint export did not create a downloadable file.");
    return new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
  }

  function canvasToBlob(
    canvas: HTMLCanvasElement,
    type = "image/png",
    quality?: number,
  ): Promise<Blob> {
    return new Promise((resolve, reject) => {
      canvas.toBlob(
        (blob) => {
          if (blob) resolve(blob);
          else
            reject(new Error("Your browser couldn’t export this page image."));
        },
        type,
        quality,
      );
    });
  }

  async function imageFromFile(file: File): Promise<ImageBitmap> {
    try {
      return await createImageBitmap(file);
    } catch {
      throw new Error(
        `Couldn’t open ${file.name}. Check that it is a valid PNG, JPEG or WebP image.`,
      );
    }
  }

  async function convert(
    report: (progressValue: number, label: string) => void,
  ): Promise<{ blob: Blob; name: string; detail?: string }> {
    const first = selectedFiles[0];
    if (!first) throw new Error("Choose a file to get started.");
    const stem = safeFileStem(first.name);

    switch (activeOperation) {
      case "pdf-to-word": {
        const { pages, blocks, images } = await extractPdfPages(
          first,
          MAX_PDF_PAGES,
          report,
        );
        const sections = pages.flatMap((text, index) => [
          {
            heading: `Original page ${index + 1}`,
            text: "",
            image: images[index],
          },
          {
            heading: `Editable text — Page ${index + 1}`,
            text: text || "(No selectable text found on this page.)",
            blocks: blocks[index],
          },
        ]);
        report(90, "Preparing your Word document");
        return {
          blob: await createWordDocument(stem, sections),
          name: `${stem}-preserved.docx`,
          detail:
            "Includes a visual copy of each original PDF page followed by its editable extracted text. Scanned pages are preserved visually but are not OCR-converted.",
        };
      }
      case "word-to-pdf": {
        report(35, "Reading Word document");
        const content = await extractWordContent(first, first.name);
        report(78, "Creating your PDF");
        return {
          blob: await createTextPdf(
            stem,
            [{ text: content.text, blocks: content.blocks }],
            content.images,
          ),
          name: `${stem}.pdf`,
          detail: `Editable text is preserved, with ${content.images.length} embedded image${content.images.length === 1 ? "" : "s"} included on separate pages. Original Word layout and formatting are not reproduced.`,
        };
      }
      case "powerpoint-to-word": {
        report(20, "Reading PowerPoint slides");
        const { slides, blocks, images } = await extractSlides(first);
        report(83, "Preparing your Word document");
        const sections: WordDocumentSection[] = slides.map((text, index) => ({
          heading: `Slide ${index + 1}`,
          text,
          blocks: blocks[index],
        }));
        sections.push(
          ...images.map((image, index) => ({
            heading: `Original presentation image ${index + 1}`,
            text: "",
            image,
          })),
        );
        return {
          blob: await createWordDocument(stem, sections),
          name: `${stem}-text.docx`,
          detail: `Slide text, speaker notes, cached chart values and order are preserved; ${images.length} embedded ${images.length === 1 ? "image is" : "images are"} included. Slide layout, chart visuals and animations may differ.`,
        };
      }
      case "word-to-powerpoint": {
        report(30, "Reading Word document");
        const content = await extractWordContent(first, first.name);
        const contentSlides = splitDocumentBlocksIntoSlides(content.blocks);
        const slides = contentSlides.length;
        if (slides + content.images.length > MAX_POWERPOINT_SLIDES) {
          throw new Error(
            `This conversion would create more than ${MAX_POWERPOINT_SLIDES} slides after preserving the extracted text and images.`,
          );
        }
        report(78, `Creating ${slides} editable slides`);
        return {
          blob: await savePresentation(
            stem,
            [],
            content.images,
            false,
            contentSlides,
          ),
          name: `${stem}.pptx`,
          detail: `${slides} structured editable slides created with headings, lists and tables where present. ${content.images.length} embedded image${content.images.length === 1 ? " is" : "s are"} preserved on additional slides.`,
        };
      }
      case "powerpoint-to-pdf": {
        report(20, "Reading PowerPoint slides");
        const { slides, blocks, images } = await extractSlides(first);
        const sections = slides.map((text, index) => ({
          heading: `Slide ${index + 1}`,
          text,
          blocks: blocks[index],
        }));
        report(82, "Creating a text-first PDF");
        return {
          blob: await createTextPdf(stem, sections, images),
          name: `${stem}-text.pdf`,
          detail: `Editable slide text, speaker notes and cached chart values are preserved, followed by ${images.length} embedded image${images.length === 1 ? "" : "s"} on separate pages. Original slide layout and chart visuals may differ.`,
        };
      }
      case "pdf-to-powerpoint": {
        const { pages, images } = await extractPdfPages(
          first,
          MAX_POWERPOINT_SLIDES,
          report,
        );
        const slides = pages.map(
          (text, index) =>
            text || `(No selectable text found on page ${index + 1}.)`,
        );
        const pageTextGroups = slides.map((text) =>
          splitTextIntoSlides(text, 850),
        );
        const outputSlideCount = pageTextGroups.reduce(
          (total, group) => total + group.length,
          0,
        );
        if (outputSlideCount > MAX_POWERPOINT_SLIDES) {
          throw new Error(
            `The editable text requires ${outputSlideCount} slides, above the ${MAX_POWERPOINT_SLIDES}-slide limit. Convert fewer pages at a time.`,
          );
        }
        report(82, `Creating ${outputSlideCount} editable slides`);
        return {
          blob: await savePresentation(
            stem,
            slides,
            images,
            true,
            undefined,
            pageTextGroups,
          ),
          name: `${stem}.pptx`,
          detail: `Each PDF page image is paired with editable extracted text; lengthy pages continue on separate text slides. Full extracted text is also included in speaker notes. Scanned-page text is not OCR-converted.`,
        };
      }
      case "compress-pdf": {
        const { PDFDocument } = await import("pdf-lib");
        const { pdf: source, close } = await loadPdf(first, PDF_EXPORT_LIMIT);
        try {
          const quality = document.querySelector<HTMLInputElement>(
            'input[name="compression-quality"]:checked',
          )?.value;
          const presets: Record<string, { scale: number; jpeg: number }> = {
            standard: { scale: 1.35, jpeg: 0.78 },
            balanced: { scale: 1, jpeg: 0.64 },
            compact: { scale: 0.75, jpeg: 0.48 },
          };
          const preset = presets[quality ?? "balanced"];
          if (!preset)
            throw new Error("Choose a compression setting before continuing.");
          const output = await PDFDocument.create();
          for (
            let pageNumber = 1;
            pageNumber <= source.numPages;
            pageNumber += 1
          ) {
            report(
              ((pageNumber - 1) / source.numPages) * 80,
              `Optimising page ${pageNumber} of ${source.numPages}`,
            );
            const page = await source.getPage(pageNumber);
            const viewport = page.getViewport({ scale: preset.scale });
            const { canvas, context } = createPageCanvas(viewport);
            await page.render({ canvas, canvasContext: context, viewport })
              .promise;
            const image = await canvasToBlob(canvas, "image/jpeg", preset.jpeg);
            canvas.width = 0;
            canvas.height = 0;
            const embedded = await output.embedJpg(await image.arrayBuffer());
            const outputPage = output.addPage([
              viewport.width,
              viewport.height,
            ]);
            outputPage.drawImage(embedded, {
              x: 0,
              y: 0,
              width: viewport.width,
              height: viewport.height,
            });
          }
          report(90, "Saving the compressed PDF");
          const bytes = Uint8Array.from(
            await output.save({ useObjectStreams: true }),
          ).buffer;
          const blob = new Blob([bytes], { type: "application/pdf" });
          const comparison =
            blob.size < first.size
              ? `File size reduced by ${Math.round((1 - blob.size / first.size) * 100)}%.`
              : "This PDF did not become smaller at this setting. Try the compact option.";
          return {
            blob,
            name: `${stem}-compressed.pdf`,
            detail: `Pages are now image-only. ${comparison}`,
          };
        } finally {
          await close();
        }
      }
      case "merge-pdf": {
        const { PDFDocument } = await import("pdf-lib");
        const output = await PDFDocument.create();
        for (const [fileIndex, file] of selectedFiles.entries()) {
          report(
            (fileIndex / selectedFiles.length) * 85,
            `Adding ${file.name}`,
          );
          const bytes = await file.arrayBuffer();
          const source = await PDFDocument.load(bytes);
          if (output.getPageCount() + source.getPageCount() > MAX_PDF_PAGES) {
            throw new Error(
              `Merged documents are limited to ${MAX_PDF_PAGES} pages in total.`,
            );
          }
          const pages = await output.copyPages(source, source.getPageIndices());
          pages.forEach((page) => output.addPage(page));
        }
        report(90, "Saving your merged PDF");
        const bytes = Uint8Array.from(await output.save()).buffer;
        const blob = new Blob([bytes], { type: "application/pdf" });
        return {
          blob,
          name: `${stem}-merged.pdf`,
          detail: `${selectedFiles.length} PDFs were combined in the displayed order.`,
        };
      }
      case "split-pdf": {
        const { PDFDocument } = await import("pdf-lib");
        const sourceBytes = await first.arrayBuffer();
        const source = await PDFDocument.load(sourceBytes);
        if (source.getPageCount() > MAX_PDF_PAGES) {
          throw new Error(
            `Splitting is limited to ${MAX_PDF_PAGES} pages at a time.`,
          );
        }
        const pages = parsePageRanges(pageRange.value, source.getPageCount());
        const output = await PDFDocument.create();
        const selected = await output.copyPages(
          source,
          pages.map((pageNumber) => pageNumber - 1),
        );
        selected.forEach((page) => output.addPage(page));
        report(90, "Saving the selected pages");
        return {
          blob: new Blob([Uint8Array.from(await output.save()).buffer], {
            type: "application/pdf",
          }),
          name: `${stem}-selected-pages.pdf`,
          detail: `${pages.length} page${pages.length === 1 ? "" : "s"} copied in the order requested.`,
        };
      }
      case "images-to-pdf": {
        const { PDFDocument } = await import("pdf-lib");
        const output = await PDFDocument.create();
        for (const [index, file] of selectedFiles.entries()) {
          report(
            (index / selectedFiles.length) * 82,
            `Preparing image ${index + 1} of ${selectedFiles.length}`,
          );
          const bitmap = await imageFromFile(file);
          try {
            if (bitmap.width * bitmap.height > 80_000_000) {
              throw new Error(
                `${file.name} exceeds the 80-megapixel image limit.`,
              );
            }
            const canvas = document.createElement("canvas");
            canvas.width = bitmap.width;
            canvas.height = bitmap.height;
            const context = canvas.getContext("2d", { alpha: false });
            if (!context)
              throw new Error("Your browser couldn’t create an image canvas.");
            context.fillStyle = "#ffffff";
            context.fillRect(0, 0, canvas.width, canvas.height);
            context.drawImage(bitmap, 0, 0);
            const image = await canvasToBlob(canvas, "image/jpeg", 0.9);
            canvas.width = 0;
            canvas.height = 0;
            const embedded = await output.embedJpg(await image.arrayBuffer());
            const landscape = bitmap.width > bitmap.height;
            const pageWidth = landscape ? 842 : 595;
            const pageHeight = landscape ? 595 : 842;
            const margin = 32;
            const fit = Math.min(
              (pageWidth - margin * 2) / bitmap.width,
              (pageHeight - margin * 2) / bitmap.height,
            );
            const width = bitmap.width * fit;
            const height = bitmap.height * fit;
            output.addPage([pageWidth, pageHeight]).drawImage(embedded, {
              x: (pageWidth - width) / 2,
              y: (pageHeight - height) / 2,
              width,
              height,
            });
          } finally {
            bitmap.close();
          }
        }
        report(90, "Saving your image PDF");
        return {
          blob: new Blob([Uint8Array.from(await output.save()).buffer], {
            type: "application/pdf",
          }),
          name: `${stem}-images.pdf`,
          detail: `${selectedFiles.length} image${selectedFiles.length === 1 ? "" : "s"} placed on separate pages.`,
        };
      }
      case "pdf-to-images": {
        const { default: JSZip } = await import("jszip");
        const { pdf: source, close } = await loadPdf(first, PDF_EXPORT_LIMIT);
        const archive = new JSZip();
        const pageCount = source.numPages;
        try {
          for (
            let pageNumber = 1;
            pageNumber <= source.numPages;
            pageNumber += 1
          ) {
            report(
              ((pageNumber - 1) / source.numPages) * 80,
              `Rendering page ${pageNumber} of ${source.numPages}`,
            );
            const page = await source.getPage(pageNumber);
            const viewport = page.getViewport({ scale: 1.5 });
            const canvas = document.createElement("canvas");
            canvas.width = Math.ceil(viewport.width);
            canvas.height = Math.ceil(viewport.height);
            const context = canvas.getContext("2d", { alpha: false });
            if (!context)
              throw new Error("Your browser couldn’t create a page canvas.");
            await page.render({ canvas, canvasContext: context, viewport })
              .promise;
            const image = await canvasToBlob(canvas);
            canvas.width = 0;
            canvas.height = 0;
            archive.file(
              `${stem}-page-${String(pageNumber).padStart(3, "0")}.png`,
              image,
            );
          }
        } finally {
          await close();
        }
        report(90, "Packing the page images");
        return {
          blob: await archive.generateAsync({ type: "blob" }),
          name: `${stem}-pages.zip`,
          detail: `${pageCount} PNG page images, bundled in a ZIP file.`,
        };
      }
    }
  }

  async function runConversion(): Promise<void> {
    if (isConverting) return;
    const validationError = validateInputFiles(activeOperation, selectedFiles);
    if (validationError) {
      setStatus(validationError, true);
      return;
    }
    if (activeOperation === "compress-pdf" && !compressionWarning.checked) {
      setStatus(
        "Confirm that you understand compression turns each page into an image.",
        true,
      );
      compressionWarning.focus();
      return;
    }

    clearResult();
    isConverting = true;
    runButton.disabled = true;
    toolButtons.forEach((button) => (button.disabled = true));
    fileList
      .querySelectorAll<HTMLButtonElement>("button")
      .forEach((button) => (button.disabled = true));
    uploadAction.disabled = true;
    fileInput.disabled = true;
    setProgress(1, "Starting conversion");
    setStatus("Your file is being processed in this browser.");
    try {
      const converted = await convert(setProgress);
      setProgress(95, "Preparing your download");
      finish(converted.blob, converted.name, converted.detail);
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : "The conversion could not be completed.";
      setStatus(message, true);
      setProgress(0, "Conversion not completed");
    } finally {
      isConverting = false;
      runButton.disabled = false;
      toolButtons.forEach((button) => (button.disabled = false));
      uploadAction.disabled = false;
      fileInput.disabled = false;
    }
  }

  toolButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const operation = button.dataset.tool;
      if (operation && isOperationId(operation)) renderOperation(operation);
    });
  });
  dropZone.addEventListener("click", (event) => {
    if (
      event.target === fileInput ||
      (event.target as HTMLElement).closest("button")
    )
      return;
    fileInput.click();
  });
  dropZone.addEventListener("keydown", (event) => {
    if (
      event.target !== dropZone ||
      (event.key !== "Enter" && event.key !== " ")
    )
      return;
    event.preventDefault();
    fileInput.click();
  });
  uploadAction.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => {
    if (fileInput.files) chooseFiles(fileInput.files);
  });
  runButton.addEventListener("click", () => void runConversion());

  dropZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
  dropZone.addEventListener("dragleave", (event) => {
    if (!dropZone.contains(event.relatedTarget as Node | null)) {
      dropZone.classList.remove("dragging");
    }
  });
  dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
    if (event.dataTransfer?.files) chooseFiles(event.dataTransfer.files);
  });

  fileList.addEventListener("click", (event) => {
    if (isConverting) return;
    const target = (event.target as HTMLElement).closest<HTMLButtonElement>(
      "[data-file-action]",
    );
    if (!target) return;
    const index = Number(target.dataset.fileIndex);
    if (!Number.isInteger(index) || index < 0 || index >= selectedFiles.length)
      return;
    if (target.dataset.fileAction === "remove") {
      selectedFiles = selectedFiles.filter(
        (_, fileIndex) => fileIndex !== index,
      );
    } else {
      const destination =
        target.dataset.fileAction === "up" ? index - 1 : index + 1;
      if (destination < 0 || destination >= selectedFiles.length) return;
      const next = [...selectedFiles];
      [next[index], next[destination]] = [next[destination], next[index]];
      selectedFiles = next;
    }
    renderFileList();
    setStatus(
      selectedFiles.length === 0
        ? "Add a file to get started."
        : `${selectedFiles.length} PDF${selectedFiles.length === 1 ? "" : "s"} selected. Use the arrow buttons to arrange page order.`,
    );
  });

  renderOperation(activeOperation);
}
