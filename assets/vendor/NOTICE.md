# Vendor Notice

Vendored browser libraries:

- `pdf-lib` 1.17.1, MIT license, https://pdf-lib.js.org/
- PDF.js 3.11.174, Apache-2.0 license, https://mozilla.github.io/pdf.js/
- `zetajs` 1.2.0, MIT license, https://github.com/allotropia/zetajs
  (`zetajs/zeta.js`, `zetajs/zetaHelper.js`)

The LibreOffice WebAssembly engine used by /convert/ is served same-origin
from `/wasm/zeta-24-2/` (hosted in R2, not in this repository). It is built
from LibreOffice core branch `distro/allotropia/zeta-24-2` (MPL-2.0 and other
open-source licenses; see https://www.libreoffice.org/about-us/licenses)
statically linked with Qt 5.15.2 for wasm (LGPL-3.0,
https://github.com/allotropia/qtbase/tree/5.15.2%2Bwasm) using Emscripten
3.1.65 (https://github.com/allotropia/emscripten/tree/fixed-3.1.65).

These browser libraries are served locally from this repository so PDF Freely
has no third-party CDN dependency at runtime.
