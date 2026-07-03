// SPDX-License-Identifier: CC0-1.0
// Document → PDF converter (main tool). Routes each file by type:
//   - office documents → LibreOffice WASM (lazily downloaded on first need)
//   - jpg/png/webp     → pdf-lib, instantly, no engine download
//   - pdf              → passthrough (no conversion needed)
import { needsLibreOffice, onStatus, convertToPdf } from '/assets/shared/lo-convert.js?v=2026-07-02-2';

const badge = document.getElementById('engineBadge');
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('file-input');
const table = document.getElementById('files-table');
const tbody = document.getElementById('files-body');
const toolbar = document.getElementById('toolbar');
const summaryEl = document.getElementById('summary');
const downloadAllBtn = document.getElementById('download-all');
const previewWrap = document.getElementById('preview-wrap');
const previewTitle = document.getElementById('preview-title');
const previewFrame = document.getElementById('preview');

const jobs = new Map();
let nextId = 1;
const queue = [];
let pumping = false;

// Test hooks.
window.__results = [];
window.__pageReady = true;

function setBadge(cls, text) {
  badge.className = 'engine-badge' + (cls ? ' ' + cls : '');
  badge.textContent = text;
}

// Engine status updates (only fire once an office file triggers the download).
onStatus((text, level) => setBadge(level || '', text));

function fmtSize(n) {
  if (n < 1024) return n + ' B';
  if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
  return (n / 1048576).toFixed(1) + ' MB';
}

function isImage(name) {
  return /\.(png|jpe?g|webp)$/i.test(name);
}

function isPdf(file) {
  return file.type === 'application/pdf' || /\.pdf$/i.test(file.name);
}

function addFiles(fileList) {
  for (const file of fileList) {
    const id = nextId++;
    const row = tbody.insertRow();
    row.insertCell().textContent = file.name;
    const sizeCell = row.insertCell();
    sizeCell.className = 'size';
    sizeCell.textContent = fmtSize(file.size);
    const statusCell = row.insertCell();
    statusCell.className = 'convert-status queued';
    statusCell.textContent = 'Queued';
    const resultCell = row.insertCell();
    jobs.set(id, { file, statusCell, resultCell, pdfUrl: null, pdfName: null, failed: false });
    queue.push(id);
  }
  table.hidden = false;
  toolbar.hidden = false;
  updateSummary();
  pump();
}

function updateSummary() {
  const all = [...jobs.values()];
  const done = all.filter(j => j.pdfUrl).length;
  const failed = all.filter(j => j.failed).length;
  summaryEl.textContent = `${done}/${jobs.size} converted` + (failed ? `, ${failed} failed` : '');
  downloadAllBtn.disabled = done === 0;
}

async function pump() {
  if (pumping) return;
  pumping = true;
  while (queue.length) {
    const id = queue.shift();
    const job = jobs.get(id);
    job.statusCell.className = 'convert-status working';
    job.statusCell.textContent = 'Converting…';
    try {
      const { bytes, kind } = await convertOne(job.file);
      finishJob(id, { bytes, kind });
    } catch (err) {
      finishJob(id, { error: err && err.message || String(err) });
    }
  }
  pumping = false;
}

async function convertOne(file) {
  if (isPdf(file)) {
    return { bytes: new Uint8Array(await file.arrayBuffer()), kind: 'passthrough' };
  }
  if (isImage(file.name)) {
    return { bytes: await imageToPdf(file), kind: 'pdf-lib' };
  }
  if (needsLibreOffice(file.name)) {
    return { bytes: await convertToPdf(file), kind: 'libreoffice' };
  }
  throw new Error('Unsupported file type');
}

// Single-page PDF from an image, page sized to the image (96 dpi → 72 pt).
async function imageToPdf(file) {
  const PDFLib = window.PDFLib;
  if (!PDFLib) throw new Error('pdf-lib is not loaded');
  let bytes = new Uint8Array(await file.arrayBuffer());
  let isJpeg = /\.jpe?g$/i.test(file.name);
  if (/\.webp$/i.test(file.name)) {
    bytes = await webpToPng(file);
    isJpeg = false;
  }
  const doc = await PDFLib.PDFDocument.create();
  const image = isJpeg ? await doc.embedJpg(bytes) : await doc.embedPng(bytes);
  const width = image.width * 72 / 96;
  const height = image.height * 72 / 96;
  const page = doc.addPage([width, height]);
  page.drawImage(image, { x: 0, y: 0, width, height });
  return doc.save();
}

async function webpToPng(file) {
  const bitmap = await createImageBitmap(file);
  const canvas = document.createElement('canvas');
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  canvas.getContext('2d').drawImage(bitmap, 0, 0);
  const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}

function finishJob(id, { bytes, kind, error }) {
  const job = jobs.get(id);
  const name = job.file.name;
  const stem = name.lastIndexOf('.') > 0 ? name.substring(0, name.lastIndexOf('.')) : name;

  if (error) {
    job.failed = true;
    job.statusCell.className = 'convert-status failed';
    job.statusCell.textContent = 'Failed';
    job.resultCell.textContent = error;
    window.__results.push({ name, ok: false, error: String(error) });
  } else {
    const blob = new Blob([bytes], { type: 'application/pdf' });
    job.pdfUrl = URL.createObjectURL(blob);
    job.pdfName = stem + '.pdf';
    job.statusCell.className = 'convert-status done';
    job.statusCell.textContent = kind === 'passthrough' ? 'Already PDF' : 'Done';

    const a = document.createElement('a');
    a.href = job.pdfUrl;
    a.download = job.pdfName;
    a.textContent = 'Download';
    a.className = 'mini-button dl';
    const p = document.createElement('button');
    p.type = 'button';
    p.textContent = 'Preview';
    p.className = 'mini-button';
    p.style.marginLeft = '0.4rem';
    p.onclick = () => showPreview(job);
    job.resultCell.append(a, p);

    showPreview(job);
    window.__results.push({
      name, ok: true, filter: kind,
      size: bytes.length,
      head: String.fromCharCode(...bytes.slice(0, 5)),
    });
  }
  updateSummary();
}

function showPreview(job) {
  previewWrap.hidden = false;
  previewTitle.textContent = job.pdfName;
  previewFrame.src = job.pdfUrl;
}

downloadAllBtn.onclick = () => {
  for (const job of jobs.values()) {
    if (!job.pdfUrl) continue;
    const a = document.createElement('a');
    a.href = job.pdfUrl;
    a.download = job.pdfName;
    a.click();
  }
};

fileInput.onchange = () => { addFiles(fileInput.files); fileInput.value = ''; };
dropZone.ondragover = e => { e.preventDefault(); dropZone.classList.add('dragging'); };
dropZone.ondragleave = () => dropZone.classList.remove('dragging');
dropZone.ondrop = e => {
  e.preventDefault();
  dropZone.classList.remove('dragging');
  addFiles(e.dataTransfer.files);
};

if (!window.crossOriginIsolated) {
  // Images and PDFs still work; only the office-document path needs isolation.
  console.warn('Not cross-origin isolated: office document conversion unavailable.');
}
