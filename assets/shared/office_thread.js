// SPDX-License-Identifier: MIT
// Runs inside the LibreOffice WASM office thread (web worker).
// Receives {cmd:'convert', id, from, to} messages, loads the document
// hidden, exports it as PDF with the filter matching its document type.
import { ZetaHelperThread } from '/assets/vendor/zetajs/zetaHelper.js';

const zHT = new ZetaHelperThread();
const zetajs = zHT.zetajs;
const css = zHT.css;

const loadProps = [
  new css.beans.PropertyValue({ Name: 'Hidden', Value: true }),
];

// Service name → PDF export filter. Order matters: an Impress model also
// implements drawing services, so presentations are checked first.
const FILTERS = [
  ['com.sun.star.presentation.PresentationDocument', 'impress_pdf_Export'],
  ['com.sun.star.sheet.SpreadsheetDocument', 'calc_pdf_Export'],
  ['com.sun.star.text.WebDocument', 'writer_web_pdf_Export'],
  ['com.sun.star.text.GlobalDocument', 'writer_globaldocument_pdf_Export'],
  ['com.sun.star.text.TextDocument', 'writer_pdf_Export'],
  ['com.sun.star.drawing.DrawingDocument', 'draw_pdf_Export'],
  ['com.sun.star.formula.FormulaProperties', 'math_pdf_Export'],
];

function pickFilter(xModel) {
  try {
    const xInfo = xModel.queryInterface(zetajs.type.interface(css.lang.XServiceInfo));
    for (const [service, filter] of FILTERS) {
      if (xInfo.supportsService(service)) return filter;
    }
  } catch (e) {
    console.warn('service detection failed, defaulting to writer filter', e);
  }
  return 'writer_pdf_Export';
}

zHT.thrPort.onmessage = (e) => {
  const m = e.data;
  switch (m.cmd) {
    case 'convert': {
      let xModel;
      try {
        xModel = zHT.desktop.loadComponentFromURL('file://' + m.from, '_blank', 0, loadProps);
        if (!xModel) throw Error('Document could not be loaded (unsupported or corrupt file?)');
        const filter = pickFilter(xModel);
        const storeProps = [
          new css.beans.PropertyValue({ Name: 'Overwrite', Value: true }),
          new css.beans.PropertyValue({ Name: 'FilterName', Value: filter }),
        ];
        xModel.storeToURL('file://' + m.to, storeProps);
        zetajs.mainPort.postMessage({ cmd: 'converted', id: m.id, from: m.from, to: m.to, filter });
      } catch (err) {
        let message;
        try {
          const exc = zetajs.catchUnoException(err);
          message = zetajs.getAnyType(exc) + ': ' + (exc.Message || '(no message)');
        } catch {
          message = String(err && err.message || err);
        }
        zetajs.mainPort.postMessage({ cmd: 'convert-error', id: m.id, from: m.from, message });
      } finally {
        // Close the document to free memory before the next file.
        try {
          if (xModel && xModel.queryInterface(zetajs.type.interface(css.util.XCloseable))) {
            xModel.close(false);
          }
        } catch {}
      }
      break;
    }
    default:
      // Don't throw: an uncaught error here dies inside the worker with
      // nothing reported back to the page.
      console.warn('office_thread: unknown message command:', m && m.cmd, m);
  }
};

zHT.thrPort.postMessage({ cmd: 'ready' });
