// SPDX-License-Identifier: CC0-1.0
// Shared, lazily-loaded LibreOffice WebAssembly document converter.
//
// The ~60 MB engine is only fetched when convertToPdf() is first called with
// a file that actually needs it — pages should gate on needsLibreOffice()
// so PDFs and images never trigger the download.
//
// Requirements on the embedding page:
//   - cross-origin isolation (COOP/COEP headers; SharedArrayBuffer)
//   - a <canvas id="qtcanvas"> element (hidden is fine)

const WASM_BASE = '/wasm/zeta-24-2/';
// Version param matters: /assets/* is served with immutable caching.
const THREAD_JS = '/assets/shared/office_thread.js?v=2026-07-02-2';

const OFFICE_EXTENSIONS = new Set([
  'doc', 'docx', 'docm', 'odt', 'fodt', 'ott', 'rtf', 'txt',
  'ppt', 'pptx', 'pptm', 'odp', 'fodp', 'otp',
  'xls', 'xlsx', 'xlsm', 'ods', 'fods', 'ots', 'csv',
  'html', 'htm', 'wps', 'wpd', 'pages', 'key', 'numbers', 'abw',
]);

export function fileExtension(name) {
  const n = String(name || '');
  const dot = n.lastIndexOf('.');
  return dot > 0 ? n.slice(dot + 1).toLowerCase() : '';
}

/** Does this file need the LibreOffice engine to become a PDF? */
export function needsLibreOffice(name) {
  return OFFICE_EXTENSIONS.has(fileExtension(name));
}

let enginePromise = null;
let engine = null;
let jobCounter = 0;
const inFlight = new Map();
const statusListeners = new Set();

/** Subscribe to human-readable engine status updates. Returns unsubscribe. */
export function onStatus(listener) {
  statusListeners.add(listener);
  return () => statusListeners.delete(listener);
}

function emitStatus(text, level) {
  for (const listener of statusListeners) {
    try { listener(text, level); } catch {}
  }
}

/** True once the engine is downloaded and running. */
export function engineReady() {
  return engine !== null;
}

/**
 * Start (or join) the engine download+boot. Safe to call repeatedly.
 * Resolves when the office thread is ready to convert.
 * On boot failure the attempt is discarded so a later call can retry
 * (e.g. after a transient network error), and the failure is surfaced
 * on the status badge instead of leaving "Downloading…" up forever.
 */
export function ensureEngine() {
  if (!enginePromise) {
    enginePromise = startEngine().catch((err) => {
      enginePromise = null;
      emitStatus('Document engine failed: ' + (err && err.message || err), 'danger');
      throw err;
    });
  }
  return enginePromise;
}

async function startEngine() {
  if (!window.crossOriginIsolated) {
    throw new Error('This page is not cross-origin isolated; the document converter cannot run.');
  }
  emitStatus('Downloading document engine (one-time, cached)…', 'warn');
  const { ZetaHelperMain } = await import('/assets/vendor/zetajs/zetaHelper.js');
  const zHM = new ZetaHelperMain(THREAD_JS, {
    threadJsType: 'module',
    wasmPkg: 'url:' + WASM_BASE,
  });
  await new Promise((resolve, reject) => {
    const bootTimeout = setTimeout(
      () => reject(new Error('Document engine took too long to start.')), 300000);
    zHM.start(() => {
      zHM.thrPort.onmessage = (e) => {
        const m = e.data;
        switch (m.cmd) {
          case 'ready': {
            clearTimeout(bootTimeout);
            engine = zHM;
            emitStatus('Document engine ready', '');
            resolve();
            break;
          }
          case 'converted': {
            const job = inFlight.get(m.id);
            if (!job) break;
            inFlight.delete(m.id);
            try { zHM.FS.unlink(m.from); } catch {}
            let bytes;
            try {
              bytes = zHM.FS.readFile(m.to);
              zHM.FS.unlink(m.to);
            } catch (err) {
              job.reject(new Error('Conversion produced no output.'));
              break;
            }
            job.resolve(bytes);
            break;
          }
          case 'convert-error': {
            const job = inFlight.get(m.id);
            if (!job) break;
            inFlight.delete(m.id);
            try { zHM.FS.unlink(m.from); } catch {}
            job.reject(new Error(m.message || 'Conversion failed.'));
            break;
          }
          default:
            console.warn('lo-convert: unknown message from office thread:', m);
        }
      };
    });
  });
}

/**
 * Convert a File/Blob (named with an office extension) to PDF.
 * Lazily boots the engine on first use. Returns a Uint8Array of PDF bytes.
 */
export async function convertToPdf(file) {
  await ensureEngine();
  const id = ++jobCounter;
  const ext = fileExtension(file.name);
  const from = `/tmp/lo_input_${id}` + (ext ? '.' + ext : '');
  const to = `/tmp/lo_output_${id}.pdf`;
  const buf = new Uint8Array(await file.arrayBuffer());
  engine.FS.writeFile(from, buf);
  return new Promise((resolve, reject) => {
    inFlight.set(id, { resolve, reject });
    engine.thrPort.postMessage({ cmd: 'convert', id, from, to });
  });
}
