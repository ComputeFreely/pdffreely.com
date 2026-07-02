# Vendor Notice

Vendored browser libraries:

- `pdf-lib` 1.17.1, MIT license, https://pdf-lib.js.org/
- PDF.js 3.11.174, Apache-2.0 license, https://mozilla.github.io/pdf.js/
- JSZip 3.10.1, MIT or GPLv3 license, https://stuk.github.io/jszip/
- `@matbee/libreoffice-converter` 2.6.0, MPL-2.0 license, https://github.com/matbeedotcom/libreoffice-document-converter
- Freely font bundle, built from Google Fonts and Noto Emoji repositories.
  Font licenses are included inside the R2-hosted
  `https://data.pdffreely.com/libreoffice/2.6.0/fonts/freely-fonts.zip` and
  summarized in `assets/vendor/fonts/freely-fonts-manifest.json`.

The small wrapper libraries are served locally from this repository so PDF
Freely has no third-party CDN dependency at runtime.

Large LibreOffice runtime files are served from Cloudflare R2 because they
exceed the Cloudflare Workers Assets per-file limit. Current runtime URLs:

- `https://data.pdffreely.com/libreoffice/2.6.0/wasm/soffice.wasm`
- `https://data.pdffreely.com/libreoffice/2.6.0/wasm/soffice.data`
- `https://data.pdffreely.com/libreoffice/2.6.0/fonts/freely-fonts.zip`

The small browser loader files kept in the repository are:

- `assets/vendor/libreoffice/dist/browser.js`
- `assets/vendor/libreoffice/dist/browser.worker.global.js`
- `assets/vendor/libreoffice/wasm/soffice.js`
- `assets/vendor/libreoffice/wasm/soffice.worker.js`
