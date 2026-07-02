# PDF Freely

PDF Freely is a free browser tool for PDF assembly and page-level PDF work.

Live site: https://pdffreely.com/

Use PDF Freely to merge, split, reorder, rotate, number, stamp, watermark, and convert images or documents to PDF. Use Fill Freely when you need to fill PDF forms or add visual signatures.

## Features

- Add multiple PDFs, JPGs, PNGs, WebP images, Office files, OpenDocument files, text, CSV, HTML, and EPUB.
- Merge PDFs and images into one PDF.
- Convert documents to PDF locally with LibreOffice WASM when needed.
- Reorder, select, duplicate, delete, and rotate pages.
- Download all pages, selected pages, or each page as a split ZIP.
- Convert images to PDF pages with page size, fit, margin, and quality settings.
- Add page numbers, text stamps, watermarks, title metadata, and author metadata.
- Runs locally in the browser with no upload step.

## Run Locally

This is a static site. From this directory:

```sh
node tools/serve-local.mjs
```

Then open `http://127.0.0.1:4176`.

The local server sends COOP/COEP headers because the LibreOffice document
converter needs `SharedArrayBuffer`.

## Limits

- Encrypted PDFs cannot be edited without the password.
- Existing PDF image streams are preserved rather than recompressed.
- Large PDFs and document conversions are limited by browser memory.
- LibreOffice's large runtime files and the global font bundle are loaded from
  `https://data.pdffreely.com/libreoffice/2.6.0/` only when document conversion
  is needed.
- Redaction is intentionally not included because fake redaction can leave underlying content extractable.

## Vendor Libraries

- `pdf-lib` handles PDF creation, page copying, image embedding, text drawing, metadata, and export.
- PDF.js renders local thumbnails for page previews.
- JSZip handles ZIP-based font bundles and Office package cleanup.
- `@matbee/libreoffice-converter` provides LibreOffice WASM document-to-PDF conversion.

## License

CC0-1.0. See `LICENSE`.
