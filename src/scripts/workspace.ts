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

  function requireText(text: string, fileName: string): string {
    const trimmed = text.replaceAll(String.fromCharCode(0), "").trim();
    if (!trimmed) {
      throw new Error(
        `${fileName} has no extractable text. Scanned pages or image-only documents need OCR, which isn’t available in this browser tool.`,
      );
    }
    return trimmed;
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
  ): Promise<string[]> {
    const { pdf, close } = await loadPdf(file, maxPages);
    try {
      const pages: string[] = [];
      for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
        report(
          (pageNumber / pdf.numPages) * 70,
          `Reading page ${pageNumber} of ${pdf.numPages}`,
        );
        const page = await pdf.getPage(pageNumber);
        const content = await page.getTextContent();
        pages.push(
          content.items
            .filter((item) => "str" in item)
            .map((item) => ("str" in item ? item.str : ""))
            .join(" ")
            .replace(/[ \t]+/g, " ")
            .trim(),
        );
      }
      if (!pages.some((page) => page.length > 0)) {
        throw new Error(
          "This PDF contains no extractable text. It may be scanned or image-only; OCR is not available in this browser tool.",
        );
      }
      return pages;
    } finally {
      await close();
    }
  }

  async function extractWordText(file: File): Promise<string> {
    const { default: mammoth } = await import("mammoth");
    const result = await mammoth.extractRawText({
      arrayBuffer: await file.arrayBuffer(),
    });
    return requireText(result.value, file.name);
  }

  async function extractSlides(file: File): Promise<string[]> {
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
    for (const [index, path] of entries.entries()) {
      const xmlText = await zip.file(path)?.async("string");
      if (!xmlText)
        throw new Error(`Couldn’t read slide ${index + 1} from ${file.name}.`);
      const xml = new DOMParser().parseFromString(xmlText, "application/xml");
      if (xml.querySelector("parsererror"))
        throw new Error(`${file.name} has an unreadable slide.`);
      const text = Array.from(xml.getElementsByTagNameNS("*", "p"))
        .map((paragraph) =>
          Array.from(paragraph.getElementsByTagNameNS("*", "t"))
            .map((run) => run.textContent ?? "")
            .join(""),
        )
        .filter((paragraph) => paragraph.trim())
        .join("\n");
      slides.push(text.trim());
    }
    if (!slides.some((slide) => slide.length > 0)) {
      throw new Error(`${file.name} contains no extractable slide text.`);
    }
    return slides;
  }

  async function createWordDocument(
    title: string,
    sections: Array<{ heading?: string; text: string }>,
  ): Promise<Blob> {
    const { Document, HeadingLevel, Packer, Paragraph, TextRun } =
      await import("docx");
    const children = [
      new Paragraph({ text: title, heading: HeadingLevel.TITLE }),
    ];
    for (const section of sections) {
      if (section.heading) {
        children.push(
          new Paragraph({
            text: section.heading,
            heading: HeadingLevel.HEADING_1,
          }),
        );
      }
      for (const paragraph of section.text
        .split(/\n{1,}/)
        .filter((line) => line.trim())) {
        children.push(
          new Paragraph({
            children: [new TextRun({ text: paragraph, size: 22 })],
            spacing: { after: 140 },
          }),
        );
      }
    }
    return Packer.toBlob(new Document({ sections: [{ children }] }));
  }

  async function createTextPdf(
    title: string,
    sections: Array<{ heading?: string; text: string }>,
  ): Promise<Blob> {
    const { PDFDocument, PageSizes, StandardFonts, rgb } =
      await import("pdf-lib");
    const pdf = await PDFDocument.create();
    const titleFont = await pdf.embedFont(StandardFonts.TimesRomanBold);
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

    try {
      drawWrappedText(title, titleFont, 22, rgb(0.17, 0.26, 0.2), 29, 12);
      for (const section of sections) {
        if (section.heading) {
          drawWrappedText(
            section.heading,
            headingFont,
            14,
            rgb(0.23, 0.41, 0.29),
            20,
            5,
          );
        }
        drawWrappedText(
          section.text,
          bodyFont,
          10.5,
          rgb(0.2, 0.22, 0.2),
          15,
          5,
        );
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

    pages.forEach((pageItem, index) => {
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
  ): Promise<Blob> {
    const { default: PptxGenJS } = await import("pptxgenjs");
    const presentation = new PptxGenJS();
    presentation.layout = "LAYOUT_WIDE";
    presentation.author = "Scriptorium";
    presentation.subject = "Text extracted and reorganised in Scriptorium";
    presentation.title = title;

    slides.forEach((text, index) => {
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
        const pages = await extractPdfPages(first, MAX_PDF_PAGES, report);
        const sections = pages.map((text, index) => ({
          heading: `Page ${index + 1}`,
          text: text || "(No selectable text found on this page.)",
        }));
        report(90, "Preparing your Word document");
        return {
          blob: await createWordDocument(stem, sections),
          name: `${stem}-text.docx`,
          detail:
            "Extracted text only. Original page design, images, tables and formatting are not preserved. Scanned pages require OCR.",
        };
      }
      case "word-to-pdf": {
        report(35, "Reading Word document");
        const text = await extractWordText(first);
        report(78, "Creating your PDF");
        return {
          blob: await createTextPdf(stem, [{ text }]),
          name: `${stem}.pdf`,
          detail:
            "Text-only PDF. Original page layout, embedded images and formatting are not preserved.",
        };
      }
      case "powerpoint-to-word": {
        report(20, "Reading PowerPoint slides");
        const slides = await extractSlides(first);
        report(83, "Preparing your Word document");
        return {
          blob: await createWordDocument(
            stem,
            slides.map((text, index) => ({
              heading: `Slide ${index + 1}`,
              text: text || "(No text on this slide.)",
            })),
          ),
          name: `${stem}-text.docx`,
          detail:
            "Slide text and order are kept. Visual design, images, charts, notes and animations are not.",
        };
      }
      case "word-to-powerpoint": {
        report(30, "Reading Word document");
        const text = await extractWordText(first);
        const slides = splitTextIntoSlides(text);
        if (slides.length > MAX_POWERPOINT_SLIDES) {
          throw new Error(
            `This conversion would create more than ${MAX_POWERPOINT_SLIDES} slides.`,
          );
        }
        report(78, `Creating ${slides.length} editable slides`);
        return {
          blob: await savePresentation(stem, slides),
          name: `${stem}.pptx`,
          detail: `${slides.length} text-first slides created. Original Word layout and embedded content are not included.`,
        };
      }
      case "powerpoint-to-pdf": {
        report(20, "Reading PowerPoint slides");
        const slides = await extractSlides(first);
        const sections = slides.map((text, index) => ({
          heading: `Slide ${index + 1}`,
          text: text || "(No text on this slide.)",
        }));
        report(82, "Creating a text-first PDF");
        return {
          blob: await createTextPdf(stem, sections),
          name: `${stem}-text.pdf`,
          detail:
            "Text-first PDF with one page per slide. Slide visuals, fonts, charts and animations are not reproduced.",
        };
      }
      case "pdf-to-powerpoint": {
        const pages = await extractPdfPages(
          first,
          MAX_POWERPOINT_SLIDES,
          report,
        );
        const slides = pages.map(
          (text, index) =>
            text || `(No selectable text found on page ${index + 1}.)`,
        );
        report(82, `Creating ${slides.length} editable slides`);
        return {
          blob: await savePresentation(stem, slides),
          name: `${stem}.pptx`,
          detail: `${slides.length} text-first slides created. PDF layout and images are not recreated.`,
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
