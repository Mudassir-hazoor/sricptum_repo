# Scriptorium

Scriptorium is a browser-based document workspace for everyday PDF, Word,
PowerPoint and image tasks. Files are processed on the visitor’s device; the
website does not upload documents or require an account.

**Live website:** <https://mudassir-hazoor.github.io/sricptum_repo/>

## Available tools

- **PDF:** extract text into Word or PowerPoint, merge documents, extract
  selected pages, export pages as PNG, or create an image-only compressed PDF.
- **Word:** create a text-first PDF or an editable PowerPoint from a `.docx`.
- **PowerPoint:** create a text-first PDF or a Word document from a `.pptx`.
- **Images:** place PNG, JPEG or WebP images on pages in a PDF.

### Conversion limits

Office-to-PDF, PDF-to-Office and PowerPoint conversions carry text into a new
document with a consistent Scriptorium type style, clear headings and readable
spacing. They do not reproduce the source document’s exact design, layout,
fonts, images, charts, forms, or animations. PDF text extraction does not
include OCR: scanned or image-only pages cannot be converted to editable text.
PDF compression rasterises pages, which removes selectable text, forms, links,
and other interactive or accessibility features. Review the source and output
before relying on a converted document.

The site accepts `.pdf`, `.docx`, `.pptx`, `.png`, `.jpg`, `.jpeg` and `.webp`.
Legacy `.doc` and `.ppt` files are not supported. Each file is limited to 30 MB
and selected files together to 80 MB. Processing limits vary by tool and are
shown in the workspace.

## Develop locally

Requires Node.js 24 or later and npm.

```powershell
npm install
npm run dev
```

Open the address printed by Astro, usually <http://localhost:4321>.

Run the checks before publishing changes:

```powershell
npm run build
npm test
npm run lint
npm audit
```

The static site is generated in `dist`. The `main` branch is deployed to
GitHub Pages by `.github/workflows/deploy-pages.yml`.

## Privacy

This is a static GitHub Pages website. Conversion runs in the browser; this
project has no document-upload endpoint or server-side conversion service.
Never close the page before downloading your result, and keep your original
file until you have checked the conversion.
