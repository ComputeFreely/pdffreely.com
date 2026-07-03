# PDF Freely

PDF Freely is a free browser tool for PDF assembly and page-level PDF work.

Live site: https://pdffreely.com/

Use PDF Freely to merge, split, reorder, rotate, number, stamp, watermark, and convert images to PDF. Use Fill Freely when you need to fill PDF forms or add visual signatures.

## Features

- Add multiple PDFs, JPGs, PNGs, and WebP images.
- Merge PDFs and images into one PDF.
- Reorder, select, duplicate, delete, and rotate pages.
- Download all pages, selected pages, or each page as a split ZIP.
- Convert images to PDF pages with page size, fit, margin, and quality settings.
- Convert Word, PowerPoint, Excel, and OpenDocument files to PDF locally with
  LibreOffice WebAssembly (zetajs-based). The converter is the main tool at `/`;
  the page organizer lives at `/organize/` and also accepts office documents.
  The ~60 MB engine downloads on demand only when a file needs it (images and
  PDFs never trigger it) and is cached by the browser.
- Add page numbers, text stamps, watermarks, title metadata, and author metadata.
- Runs locally in the browser with no upload step.

## Run Locally

This is a static site plus a small Cloudflare Worker. From this directory:

```sh
node tools/serve-local.mjs
```

Then open `http://127.0.0.1:4176`.

The local server sends COOP/COEP headers because the LibreOffice document
converter needs `SharedArrayBuffer`.

For the converter to work locally, place the LibreOffice WASM engine files
(`soffice.js`, `soffice.wasm`, `soffice.data`, `soffice.data.js.metadata`)
in `wasm/zeta-24-2/` (gitignored). Download them from the live site, e.g.:

```sh
mkdir -p wasm/zeta-24-2 && cd wasm/zeta-24-2
for f in soffice.js soffice.wasm soffice.data soffice.data.js.metadata; do
  curl -LO "https://pdffreely.com/wasm/zeta-24-2/$f"
done
```

In production the same files are served from R2 by `worker.js` under
`/wasm/zeta-24-2/` (brotli-compressed when the client accepts it). They are
built from LibreOffice core branch `distro/allotropia/zeta-24-2` with
Qt 5.15.2 (wasm) and Emscripten 3.1.65 — see
[static/README.wasm.md](https://github.com/LibreOffice/core/blob/master/static/README.wasm.md).

## Limits

- Encrypted PDFs cannot be edited without the password.
- Existing PDF image streams are preserved rather than recompressed.
- Large PDFs and image conversions are limited by browser memory.
- Redaction is intentionally not included because fake redaction can leave underlying content extractable.

## Vendor Libraries

- `pdf-lib` handles PDF creation, page copying, image embedding, text drawing, metadata, and export.
- PDF.js renders local thumbnails for page previews.
- `zetajs` (MIT) scripts the LibreOffice WebAssembly engine for document-to-PDF conversion.
- LibreOffice WASM engine (MPL-2.0, with Qt 5.15 LGPL-3.0) served from `/wasm/`.

## License

CC0-1.0. See `LICENSE`.
