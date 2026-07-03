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
- Add page numbers, text stamps, watermarks, title metadata, and author metadata.
- Runs locally in the browser with no upload step.

## Run Locally

This is a static site. From this directory:

```sh
node tools/serve-local.mjs
```

Then open `http://127.0.0.1:4176`.

## Limits

- Encrypted PDFs cannot be edited without the password.
- Existing PDF image streams are preserved rather than recompressed.
- Large PDFs and image conversions are limited by browser memory.
- Redaction is intentionally not included because fake redaction can leave underlying content extractable.

## Vendor Libraries

- `pdf-lib` handles PDF creation, page copying, image embedding, text drawing, metadata, and export.
- PDF.js renders local thumbnails for page previews.

## License

CC0-1.0. See `LICENSE`.
