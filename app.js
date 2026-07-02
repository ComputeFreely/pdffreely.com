(function () {
  "use strict";

  var $ = function (selector) {
    return document.querySelector(selector);
  };

  var $$ = function (selector) {
    return Array.prototype.slice.call(document.querySelectorAll(selector));
  };

  var state = {
    sources: {},
    pages: [],
    activePageId: "",
    dragPageId: "",
    sourceCounter: 1,
    pageCounter: 1,
    totalInputBytes: 0,
    busy: false
  };

  var libreOfficeConverter = null;
  var libreOfficeReady = null;
  var libreOfficeModuleReady = null;
  var libreOfficeFontsReady = null;
  var libreOfficeUsesFonts = true;
  var libreOfficeLastProgress = null;
  var assetVersion = "2026-07-02-9";
  var libreOfficeRuntimeBaseUrl = "https://data.pdffreely.com/libreoffice/2.6.0/";
  var libreOfficeFontBundleUrl = libreOfficeRuntimeBaseUrl + "fonts/freely-fonts.zip";
  var libreOfficeFontTimeoutMs = 30000;
  var libreOfficeInitializeTimeoutMs = 120000;
  var libreOfficeConvertBaseTimeoutMs = 120000;
  var libreOfficeConvertPerMbTimeoutMs = 2000;
  var libreOfficePresentationConvertBaseTimeoutMs = 300000;
  var libreOfficePresentationConvertPerMbTimeoutMs = 8000;
  var libreOfficeConvertMaxTimeoutMs = 480000;
  var libreOfficePresentationConvertMaxTimeoutMs = 900000;

  var documentFormats = new Set([
    "doc", "docx", "xls", "xlsx", "ppt", "pptx", "odt", "ods", "odp",
    "rtf", "txt", "html", "htm", "csv", "epub"
  ]);

  var slowDocumentLoadFormats = new Set(["ppt", "pptx", "odp"]);

  var pageSizes = {
    letter: [612, 792],
    a4: [595.28, 841.89]
  };

  var crcTable = null;

  document.addEventListener("DOMContentLoaded", function () {
    if (!window.PDFLib || !window.pdfjsLib) {
      setStatus("PDF engine did not load", "danger");
      return;
    }

    window.pdfjsLib.GlobalWorkerOptions.workerSrc = versionedAssetUrl("assets/vendor/pdf.worker.min.js");
    bindEvents();
    syncControls();
    renderPages();
  });

  function bindEvents() {
    $("#fileInput").addEventListener("change", function (event) {
      handleFiles(event.target.files);
      event.target.value = "";
    });

    $("#dropZone").addEventListener("dragover", function (event) {
      event.preventDefault();
      $("#dropZone").classList.add("dragging");
    });

    $("#dropZone").addEventListener("dragleave", function () {
      $("#dropZone").classList.remove("dragging");
    });

    $("#dropZone").addEventListener("drop", function (event) {
      event.preventDefault();
      $("#dropZone").classList.remove("dragging");
      handleFiles(event.dataTransfer.files);
    });

    $("#pageGrid").addEventListener("click", handlePageClick);
    $("#pageGrid").addEventListener("change", handlePageChange);
    $("#pageGrid").addEventListener("dragstart", handleDragStart);
    $("#pageGrid").addEventListener("dragover", handleDragOver);
    $("#pageGrid").addEventListener("drop", handleDrop);
    $("#pageGrid").addEventListener("dragend", handleDragEnd);

    $$(".batch-bar [data-batch]").forEach(function (button) {
      button.addEventListener("click", function () {
        handleBatch(button.dataset.batch);
      });
    });

    $("#downloadPdf").addEventListener("click", function () {
      downloadCombined(false);
    });
    $("#downloadSelected").addEventListener("click", function () {
      downloadCombined(true);
    });
    $("#downloadSplit").addEventListener("click", downloadSplitZip);
    $("#resetAll").addEventListener("click", resetAll);

    $("#pdfForm").addEventListener("input", syncControls);
    $("#pdfForm").addEventListener("change", syncControls);
  }

  async function handleFiles(fileList) {
    if (state.busy) {
      setStatus("Finish the current file before adding more", "warn");
      return;
    }

    var files = Array.prototype.slice.call(fileList || []);
    var usableFiles = files.filter(isSupportedInputFile);

    if (!usableFiles.length) {
      setStatus("Use PDF, image, Office, OpenDocument, text, CSV, HTML, or EPUB files", "warn");
      return;
    }

    setBusy(true, "Reading " + usableFiles.length + " file" + (usableFiles.length === 1 ? "" : "s"));

    var lastError = "";
    for (var index = 0; index < usableFiles.length; index += 1) {
      var file = usableFiles[index];
      try {
        if (isPdfFile(file)) {
          await addPdfFile(file);
        } else if (isImageFile(file)) {
          await addImageFile(file);
        } else {
          await addDocumentFile(file);
        }
      } catch (error) {
        lastError = file.name + ": " + getErrorMessage(error);
        setStatus(lastError, "danger");
      }
    }

    if (!state.activePageId && state.pages.length) {
      state.activePageId = state.pages[0].id;
    }

    setBusy(false, state.pages.length ? "Ready" : (lastError || "No pages"), lastError && !state.pages.length ? "danger" : undefined);
    renderPages();
  }

  async function addPdfFile(file) {
    var buffer = await file.arrayBuffer();
    var bytes = new Uint8Array(buffer);
    await addPdfBytes(bytes, file.name, file.size);
  }

  async function addPdfBytes(bytes, name, inputSize, convertedFrom) {
    var pdfDoc = await window.PDFLib.PDFDocument.load(bytes, {
      ignoreEncryption: true,
      updateMetadata: false
    });
    var sourceId = "source-" + state.sourceCounter;
    state.sourceCounter += 1;

    var source = {
      id: sourceId,
      type: "pdf",
      name: name,
      size: inputSize || bytes.length,
      bytes: bytes,
      pdfDoc: pdfDoc,
      pageCount: pdfDoc.getPageCount(),
      pdfjsDoc: null,
      convertedFrom: convertedFrom || ""
    };

    state.sources[sourceId] = source;
    state.totalInputBytes += source.size;

    for (var pageIndex = 0; pageIndex < source.pageCount; pageIndex += 1) {
      state.pages.push({
        id: "page-" + state.pageCounter,
        sourceId: sourceId,
        type: "pdf",
        pageIndex: pageIndex,
        selected: true,
        rotation: 0,
        thumb: "",
        thumbState: "pending"
      });
      state.pageCounter += 1;
    }

    try {
      source.pdfjsDoc = await window.pdfjsLib.getDocument({
        data: bytes.slice(0),
        disableFontFace: true
      }).promise;
      renderPdfThumbnails(source);
    } catch (error) {
      markSourceThumbsFailed(sourceId);
    }
  }

  async function addDocumentFile(file) {
    var pdfBytes = await convertDocumentFileToPdf(file);
    await addPdfBytes(pdfBytes, file.name, file.size, getExtension(file.name).toUpperCase());
  }

  async function convertDocumentFileToPdf(file) {
    var bytes = new Uint8Array(await file.arrayBuffer());
    return await convertWithLibreOffice(bytes, file.name);
  }

  async function convertWithLibreOffice(inputBytes, filename) {
    try {
      return await convertWithLibreOfficeAttempt(inputBytes, filename, true);
    } catch (error) {
      if (!shouldRetryLibreOfficeConversion(error)) {
        throw error;
      }

      setStatus("Document converter reset; retrying", "warn");
      await resetLibreOffice();
      try {
        return await convertWithLibreOfficeAttempt(inputBytes, filename, true);
      } catch (retryError) {
        if (shouldRetryLibreOfficeConversion(retryError)) {
          await resetLibreOffice();
        }
        throw retryError;
      }
    }
  }

  async function convertWithLibreOfficeAttempt(inputBytes, filename, useFonts) {
    await ensureLibreOffice(useFonts);
    setStatus("Converting " + filename, "warn");
    var timeoutMs = getLibreOfficeConversionTimeoutMs(inputBytes, filename);
    var stopHeartbeat = startLibreOfficeConversionHeartbeat(timeoutMs);
    var options = { outputFormat: "pdf" };
    var inputFormat = getExtension(filename);
    if (inputFormat) {
      options.inputFormat = inputFormat;
    }
    try {
      var result = await withTimeout(
        libreOfficeConverter.convert(inputBytes, options, filename),
        timeoutMs,
        "Document conversion timed out."
      );
      return toUint8Array(result.data);
    } finally {
      stopHeartbeat();
    }
  }

  async function ensureLibreOffice(useFonts) {
    useFonts = useFonts !== false;

    if (libreOfficeReady && libreOfficeUsesFonts === useFonts) {
      return libreOfficeReady;
    }

    if (libreOfficeReady && libreOfficeUsesFonts !== useFonts) {
      await resetLibreOffice();
    }

    if (!window.crossOriginIsolated || !window.SharedArrayBuffer) {
      throw new Error("Document conversion needs cross-origin isolation headers. Use the deployed site or the local isolated dev server.");
    }

    setStatus("Loading document converter", "warn");
    libreOfficeUsesFonts = useFonts;
    libreOfficeReady = initializeLibreOffice(useFonts).catch(function (error) {
      libreOfficeReady = null;
      libreOfficeConverter = null;
      throw error;
    });
    return libreOfficeReady;
  }

  async function initializeLibreOffice(useFonts) {
    var module = await loadLibreOfficeModule();
    var converterOptions = {
      sofficeJs: versionedAssetUrl("/assets/vendor/libreoffice/wasm/soffice.js"),
      sofficeWasm: libreOfficeRuntimeBaseUrl + "wasm/soffice.wasm",
      sofficeData: libreOfficeRuntimeBaseUrl + "wasm/soffice.data",
      sofficeWorkerJs: versionedAssetUrl("/assets/vendor/libreoffice/wasm/soffice.worker.js"),
      browserWorkerJs: versionedAssetUrl("/assets/vendor/libreoffice/dist/browser.worker.global.js"),
      verbose: false,
      onProgress: function (info) {
        libreOfficeLastProgress = {
          message: info && info.message ? info.message : "",
          phase: info && info.phase ? info.phase : "",
          at: Date.now()
        };
        var percent = Number.isFinite(info && info.percent) ? Math.round(info.percent) + "% " : "";
        var message = info && info.message ? info.message : "Working";
        setStatus(percent + message, "warn");
      }
    };

    if (useFonts) {
      setStatus("Loading document fonts", "warn");
      var fonts = await ensureLibreOfficeFonts(module);
      if (fonts.length) {
        converterOptions.fonts = fonts;
      }
    }

    libreOfficeConverter = new module.WorkerBrowserConverter(converterOptions);
    try {
      await withTimeout(
        libreOfficeConverter.initialize(),
        libreOfficeInitializeTimeoutMs,
        "Document converter timed out while starting."
      );
    } catch (error) {
      await destroyLibreOfficeConverter(libreOfficeConverter);
      libreOfficeConverter = null;
      throw error;
    }
    setStatus("Document converter ready");
  }

  async function resetLibreOffice() {
    var converter = libreOfficeConverter;
    libreOfficeReady = null;
    libreOfficeConverter = null;
    libreOfficeLastProgress = null;
    if (converter) {
      await destroyLibreOfficeConverter(converter);
    }
  }

  async function destroyLibreOfficeConverter(converter) {
    try {
      if (converter && typeof converter.destroy === "function") {
        await converter.destroy();
      }
    } catch (error) {
      console.warn("Could not destroy LibreOffice converter", error);
    }
  }

  function getLibreOfficeConversionTimeoutMs(inputBytes, filename) {
    var sizeMb = Math.ceil((inputBytes && inputBytes.length ? inputBytes.length : 0) / 1048576);
    if (slowDocumentLoadFormats.has(getExtension(filename))) {
      var presentationTimeout = libreOfficePresentationConvertBaseTimeoutMs + sizeMb * libreOfficePresentationConvertPerMbTimeoutMs;
      return Math.min(libreOfficePresentationConvertMaxTimeoutMs, presentationTimeout);
    }
    var timeout = libreOfficeConvertBaseTimeoutMs + sizeMb * libreOfficeConvertPerMbTimeoutMs;
    return Math.min(libreOfficeConvertMaxTimeoutMs, timeout);
  }

  function startLibreOfficeConversionHeartbeat(timeoutMs) {
    var startedAt = Date.now();
    var lastStatusAt = 0;
    var timer = window.setInterval(function () {
      var elapsedMs = Date.now() - startedAt;
      if (elapsedMs < 45000 || elapsedMs - lastStatusAt < 30000) {
        return;
      }
      if (!libreOfficeLastProgress || !/loading document/i.test(libreOfficeLastProgress.message)) {
        return;
      }
      lastStatusAt = elapsedMs;
      var elapsedText = formatDuration(Math.round(elapsedMs / 1000));
      var limitText = formatDuration(Math.round(timeoutMs / 1000));
      setStatus("Still loading document (" + elapsedText + " / " + limitText + ")", "warn");
    }, 5000);

    return function () {
      window.clearInterval(timer);
    };
  }

  function loadLibreOfficeModule() {
    if (!libreOfficeModuleReady) {
      libreOfficeModuleReady = import(versionedAssetUrl("/assets/vendor/libreoffice/dist/browser.js"));
    }
    return libreOfficeModuleReady;
  }

  function ensureLibreOfficeFonts(module) {
    if (!libreOfficeFontsReady) {
      var fontLoad = module.loadFontsFromUrl(libreOfficeFontBundleUrl);
      fontLoad.catch(function () {});
      libreOfficeFontsReady = withTimeout(
        fontLoad,
        libreOfficeFontTimeoutMs,
        "Timed out after " + Math.round(libreOfficeFontTimeoutMs / 1000) + " seconds."
      ).catch(function (error) {
        setStatus("Font bundle unavailable; continuing with built-in fonts", "warn");
        console.warn("Could not load LibreOffice font bundle", error);
        return [];
      });
    }
    return libreOfficeFontsReady;
  }
  function isSupportedInputFile(file) {
    return isPdfFile(file) || isImageFile(file) || isDocumentFile(file);
  }

  function isPdfFile(file) {
    return file.type === "application/pdf" || /\.pdf$/i.test(file.name);
  }

  function isImageFile(file) {
    return /^image\/(png|jpeg|webp)$/.test(file.type) || /\.(png|jpe?g|webp)$/i.test(file.name);
  }

  function isDocumentFile(file) {
    return documentFormats.has(getExtension(file.name));
  }

  async function addImageFile(file) {
    var buffer = await file.arrayBuffer();
    var bytes = new Uint8Array(buffer);
    var objectUrl = URL.createObjectURL(file);
    var dimensions = await getImageDimensions(objectUrl);
    var sourceId = "source-" + state.sourceCounter;
    state.sourceCounter += 1;

    state.sources[sourceId] = {
      id: sourceId,
      type: "image",
      name: file.name,
      size: file.size,
      bytes: bytes,
      mime: file.type,
      objectUrl: objectUrl,
      width: dimensions.width,
      height: dimensions.height
    };
    state.totalInputBytes += file.size;

    state.pages.push({
      id: "page-" + state.pageCounter,
      sourceId: sourceId,
      type: "image",
      pageIndex: 0,
      selected: true,
      rotation: 0,
      thumb: objectUrl,
      thumbState: "ready"
    });
    state.pageCounter += 1;
  }

  function getImageDimensions(url) {
    return new Promise(function (resolve, reject) {
      var image = new Image();
      image.onload = function () {
        resolve({ width: image.naturalWidth || image.width, height: image.naturalHeight || image.height });
      };
      image.onerror = function () {
        reject(new Error("Unable to read image"));
      };
      image.src = url;
    });
  }

  function renderPdfThumbnails(source) {
    state.pages.filter(function (item) {
      return item.sourceId === source.id;
    }).forEach(function (item) {
      renderPdfThumbnail(source, item);
    });
  }

  async function renderPdfThumbnail(source, item) {
    try {
      var page = await source.pdfjsDoc.getPage(item.pageIndex + 1);
      var baseViewport = page.getViewport({ scale: 1 });
      var scale = Math.min(1.2, 190 / Math.max(1, baseViewport.width));
      var viewport = page.getViewport({ scale: scale });
      var canvas = document.createElement("canvas");
      var context = canvas.getContext("2d");
      if (!context) {
        throw new Error("Canvas unavailable");
      }
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      await page.render({ canvasContext: context, viewport: viewport }).promise;
      item.thumb = canvas.toDataURL("image/jpeg", 0.82);
      item.thumbState = "ready";
    } catch (error) {
      item.thumbState = "failed";
    }
    updatePageCardThumb(item);
  }

  function markSourceThumbsFailed(sourceId) {
    state.pages.forEach(function (item) {
      if (item.sourceId === sourceId) {
        item.thumbState = "failed";
      }
    });
  }

  function renderPages() {
    var grid = $("#pageGrid");
    grid.innerHTML = state.pages.map(renderPageCard).join("");
    $("#emptyState").hidden = state.pages.length > 0;
    updateStats();
    syncControls();
  }

  function renderPageCard(item, index) {
    var source = state.sources[item.sourceId];
    var title = item.type === "pdf" ? "Page " + (item.pageIndex + 1) : "Image";
    var pageLabel = String(index + 1).padStart(2, "0");
    var activeClass = item.id === state.activePageId ? " active" : "";
    var selectedText = item.selected ? " checked" : "";
    var thumb = renderThumb(item);

    return '<article class="page-card' + activeClass + '" draggable="true" data-page-id="' + escapeAttr(item.id) + '">' +
      '<div class="page-topline">' +
        '<label><input type="checkbox" data-action="select"' + selectedText + '> <span class="page-title">' + escapeText(pageLabel + " " + title) + "</span></label>" +
        '<button class="icon-action danger" type="button" data-action="remove" aria-label="Remove page">x</button>' +
      "</div>" +
      '<div class="thumb-wrap">' + thumb + "</div>" +
      '<div class="page-meta">' +
        '<strong>' + escapeText(source ? source.name : "Unknown file") + "</strong><br>" +
        escapeText(getPageMetaText(item, source)) +
        (item.rotation ? "<br>Rotation " + item.rotation + " deg" : "") +
      "</div>" +
      '<div class="page-actions" aria-label="Page actions">' +
        '<button class="icon-action" type="button" data-action="move-up" aria-label="Move page up">up</button>' +
        '<button class="icon-action" type="button" data-action="move-down" aria-label="Move page down">down</button>' +
        '<button class="icon-action" type="button" data-action="rotate-left" aria-label="Rotate page left">-90</button>' +
        '<button class="icon-action" type="button" data-action="rotate-right" aria-label="Rotate page right">+90</button>' +
        '<button class="icon-action" type="button" data-action="duplicate" aria-label="Duplicate page">copy</button>' +
      "</div>" +
    "</article>";
  }

  function getPageMetaText(item, source) {
    if (!source) {
      return "Unknown source";
    }
    if (item.type === "pdf") {
      return (source.convertedFrom ? source.convertedFrom + " converted page " : "PDF page ") + (item.pageIndex + 1);
    }
    return formatPixels(source.width, source.height);
  }

  function renderThumb(item) {
    if (item.thumb) {
      return '<img alt="" src="' + escapeAttr(item.thumb) + '">';
    }
    if (item.thumbState === "failed") {
      return '<div class="thumb-placeholder">Preview off</div>';
    }
    return '<div class="thumb-placeholder">Loading</div>';
  }

  function updatePageCardThumb(item) {
    var card = $('[data-page-id="' + cssEscape(item.id) + '"]');
    if (!card) {
      return;
    }
    var thumb = card.querySelector(".thumb-wrap");
    if (thumb) {
      thumb.innerHTML = renderThumb(item);
    }
  }

  function handlePageClick(event) {
    var card = event.target.closest(".page-card");
    if (!card) {
      return;
    }

    state.activePageId = card.dataset.pageId;

    var button = event.target.closest("[data-action]");
    if (!button || button.dataset.action === "select") {
      updateActiveCardClass();
      return;
    }

    var action = button.dataset.action;
    if (action === "move-up") {
      movePage(state.activePageId, -1);
    } else if (action === "move-down") {
      movePage(state.activePageId, 1);
    } else if (action === "rotate-left") {
      rotatePages([state.activePageId], -90);
    } else if (action === "rotate-right") {
      rotatePages([state.activePageId], 90);
    } else if (action === "remove") {
      removePages([state.activePageId]);
    } else if (action === "duplicate") {
      duplicatePage(state.activePageId);
    }
    renderPages();
  }

  function updateActiveCardClass() {
    $$(".page-card").forEach(function (card) {
      card.classList.toggle("active", card.dataset.pageId === state.activePageId);
    });
  }

  function handlePageChange(event) {
    var checkbox = event.target.closest('[data-action="select"]');
    if (!checkbox) {
      return;
    }
    var card = event.target.closest(".page-card");
    var item = getPageById(card.dataset.pageId);
    if (item) {
      item.selected = checkbox.checked;
      updateStats();
      syncControls();
    }
  }

  function handleDragStart(event) {
    var card = event.target.closest(".page-card");
    if (!card) {
      return;
    }
    state.dragPageId = card.dataset.pageId;
    card.classList.add("dragging");
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", state.dragPageId);
  }

  function handleDragOver(event) {
    if (event.target.closest(".page-card")) {
      event.preventDefault();
    }
  }

  function handleDrop(event) {
    var targetCard = event.target.closest(".page-card");
    if (!targetCard || !state.dragPageId || targetCard.dataset.pageId === state.dragPageId) {
      return;
    }
    event.preventDefault();
    var fromIndex = findPageIndex(state.dragPageId);
    var toIndex = findPageIndex(targetCard.dataset.pageId);
    if (fromIndex === -1 || toIndex === -1) {
      return;
    }
    var moved = state.pages.splice(fromIndex, 1)[0];
    state.pages.splice(toIndex, 0, moved);
    renderPages();
  }

  function handleDragEnd() {
    state.dragPageId = "";
    $$(".page-card.dragging").forEach(function (card) {
      card.classList.remove("dragging");
    });
  }

  function handleBatch(action) {
    if (action === "select-all") {
      state.pages.forEach(function (item) {
        item.selected = true;
      });
    } else if (action === "select-none") {
      state.pages.forEach(function (item) {
        item.selected = false;
      });
    } else if (action === "invert") {
      state.pages.forEach(function (item) {
        item.selected = !item.selected;
      });
    } else if (action === "rotate-left") {
      rotatePages(getSelectedPageIds(), -90);
    } else if (action === "rotate-right") {
      rotatePages(getSelectedPageIds(), 90);
    } else if (action === "remove") {
      removePages(getSelectedPageIds());
    }
    renderPages();
  }

  function movePage(id, delta) {
    var index = findPageIndex(id);
    var nextIndex = index + delta;
    if (index < 0 || nextIndex < 0 || nextIndex >= state.pages.length) {
      return;
    }
    var tmp = state.pages[index];
    state.pages[index] = state.pages[nextIndex];
    state.pages[nextIndex] = tmp;
  }

  function duplicatePage(id) {
    var index = findPageIndex(id);
    if (index < 0) {
      return;
    }
    var copy = Object.assign({}, state.pages[index], {
      id: "page-" + state.pageCounter
    });
    state.pageCounter += 1;
    state.pages.splice(index + 1, 0, copy);
    state.activePageId = copy.id;
  }

  function rotatePages(ids, delta) {
    var idMap = {};
    ids.forEach(function (id) {
      idMap[id] = true;
    });
    state.pages.forEach(function (item) {
      if (idMap[item.id]) {
        item.rotation = normalizeDegrees(item.rotation + delta);
      }
    });
  }

  function removePages(ids) {
    var idMap = {};
    ids.forEach(function (id) {
      idMap[id] = true;
    });
    state.pages = state.pages.filter(function (item) {
      return !idMap[item.id];
    });
    if (!getPageById(state.activePageId)) {
      state.activePageId = state.pages.length ? state.pages[0].id : "";
    }
  }

  async function downloadCombined(selectedOnly) {
    var items = selectedOnly ? getSelectedPages() : state.pages.slice();
    if (!items.length) {
      setStatus(selectedOnly ? "No selected pages" : "Add pages first", "warn");
      return;
    }

    setBusy(true, selectedOnly ? "Building selected PDF" : "Building PDF");
    try {
      var bytes = await buildPdf(items);
      downloadBlob(new Blob([bytes], { type: "application/pdf" }), getOutputFilename(selectedOnly ? "-selected.pdf" : ".pdf"));
      setBusy(false, formatBytes(bytes.length) + " PDF ready");
    } catch (error) {
      setBusy(false, getErrorMessage(error), "danger");
    }
  }

  async function downloadSplitZip() {
    var items = getSelectedPages();
    if (!items.length) {
      items = state.pages.slice();
    }
    if (!items.length) {
      setStatus("Add pages first", "warn");
      return;
    }

    setBusy(true, "Building split ZIP");
    try {
      var files = [];
      for (var index = 0; index < items.length; index += 1) {
        var bytes = await buildPdf([items[index]]);
        files.push({
          name: sanitizeBase(getOutputFilename("")) + "-page-" + String(index + 1).padStart(3, "0") + ".pdf",
          data: bytes
        });
      }
      var zipBytes = createZip(files);
      downloadBlob(new Blob([zipBytes], { type: "application/zip" }), sanitizeBase(getOutputFilename("")) + "-split.zip");
      setBusy(false, formatBytes(zipBytes.length) + " ZIP ready");
    } catch (error) {
      setBusy(false, getErrorMessage(error), "danger");
    }
  }

  async function buildPdf(items) {
    var PDFLib = window.PDFLib;
    var out = await PDFLib.PDFDocument.create();
    var options = getExportOptions();
    var font = await out.embedFont(PDFLib.StandardFonts.Helvetica);
    var boldFont = await out.embedFont(PDFLib.StandardFonts.HelveticaBold);

    if (options.title) {
      out.setTitle(options.title);
    }
    if (options.author) {
      out.setAuthor(options.author);
    }
    out.setCreator("PDF Freely");
    out.setProducer("PDF Freely");
    out.setModificationDate(new Date());
    out.setCreationDate(new Date());

    for (var index = 0; index < items.length; index += 1) {
      var item = items[index];
      var page;
      if (item.type === "pdf") {
        page = await copyPdfPage(out, item);
      } else {
        page = await addImagePage(out, item, options);
      }
      drawOverlays(page, item, index + 1, items.length, options, font, boldFont);
    }

    return await out.save({
      useObjectStreams: options.compact,
      addDefaultPage: false
    });
  }

  async function copyPdfPage(out, item) {
    var PDFLib = window.PDFLib;
    var source = state.sources[item.sourceId];
    var copiedPages = await out.copyPages(source.pdfDoc, [item.pageIndex]);
    var copied = copiedPages[0];
    var existingRotation = copied.getRotation() && copied.getRotation().angle ? copied.getRotation().angle : 0;
    copied.setRotation(PDFLib.degrees(normalizeDegrees(existingRotation + item.rotation)));
    out.addPage(copied);
    return copied;
  }

  async function addImagePage(out, item, options) {
    var PDFLib = window.PDFLib;
    var source = state.sources[item.sourceId];
    var prepared = await prepareImageForPdf(source, options.imageQuality);
    var embedded = prepared.mime === "image/png" ? await out.embedPng(prepared.bytes) : await out.embedJpg(prepared.bytes);
    var size = getImagePageSize(source, options);
    var page = out.addPage(size);
    var box = fitBox(embedded.width, embedded.height, size[0], size[1], options.imageMargin, options.imageFit);

    page.drawImage(embedded, {
      x: box.x,
      y: box.y,
      width: box.width,
      height: box.height
    });
    page.setRotation(PDFLib.degrees(item.rotation));
    return page;
  }

  async function prepareImageForPdf(source, quality) {
    var preserveOriginal = quality >= 100 && (source.mime === "image/png" || source.mime === "image/jpeg");
    if (preserveOriginal) {
      return {
        bytes: source.bytes,
        mime: source.mime
      };
    }

    var image = await loadImage(source.objectUrl);
    var canvas = document.createElement("canvas");
    var context = canvas.getContext("2d");
    if (!context) {
      throw new Error("Canvas unavailable");
    }
    canvas.width = image.naturalWidth || source.width;
    canvas.height = image.naturalHeight || source.height;

    var outputType = quality >= 100 && source.mime === "image/png" ? "image/png" : "image/jpeg";
    if (outputType === "image/jpeg") {
      context.fillStyle = "#ffffff";
      context.fillRect(0, 0, canvas.width, canvas.height);
    }
    context.drawImage(image, 0, 0);
    var blob = await canvasToBlob(canvas, outputType, quality / 100);
    return {
      bytes: new Uint8Array(await blob.arrayBuffer()),
      mime: outputType
    };
  }

  function loadImage(url) {
    return new Promise(function (resolve, reject) {
      var image = new Image();
      image.onload = function () {
        resolve(image);
      };
      image.onerror = function () {
        reject(new Error("Unable to prepare image"));
      };
      image.src = url;
    });
  }

  function canvasToBlob(canvas, type, quality) {
    return new Promise(function (resolve, reject) {
      canvas.toBlob(function (blob) {
        if (blob) {
          resolve(blob);
        } else {
          reject(new Error("Image conversion failed"));
        }
      }, type, quality);
    });
  }

  function getImagePageSize(source, options) {
    if (options.imagePageSize === "original") {
      return [Math.max(1, source.width), Math.max(1, source.height)];
    }
    return pageSizes[options.imagePageSize] || pageSizes.letter;
  }

  function fitBox(imageWidth, imageHeight, pageWidth, pageHeight, margin, fit) {
    var boxWidth = Math.max(1, pageWidth - margin * 2);
    var boxHeight = Math.max(1, pageHeight - margin * 2);
    if (fit === "stretch") {
      return { x: margin, y: margin, width: boxWidth, height: boxHeight };
    }
    var scale = fit === "cover" ?
      Math.max(boxWidth / imageWidth, boxHeight / imageHeight) :
      Math.min(boxWidth / imageWidth, boxHeight / imageHeight);
    var width = imageWidth * scale;
    var height = imageHeight * scale;
    return {
      x: margin + (boxWidth - width) / 2,
      y: margin + (boxHeight - height) / 2,
      width: width,
      height: height
    };
  }

  function drawOverlays(page, item, pageNumber, totalPages, options, font, boldFont) {
    var size = page.getSize();
    var pageWidth = size.width;
    var pageHeight = size.height;
    var color = hexToRgb(options.stampColor);

    if (options.watermarkText && scopeApplies(options.watermarkScope, item)) {
      drawWatermark(page, options.watermarkText, pageWidth, pageHeight, options, boldFont);
    }

    if (options.stampText && scopeApplies(options.stampScope, item)) {
      drawPositionedText(page, options.stampText, pageWidth, pageHeight, {
        font: boldFont,
        size: options.stampSize,
        color: color,
        position: options.stampPosition,
        opacity: 1
      });
    }

    if (options.pageNumbers) {
      drawPositionedText(page, "Page " + pageNumber + " of " + totalPages, pageWidth, pageHeight, {
        font: font,
        size: 10,
        color: color,
        position: options.pageNumberPosition,
        opacity: 0.9
      });
    }

  }

  function drawWatermark(page, text, pageWidth, pageHeight, options, font) {
    var PDFLib = window.PDFLib;
    var fontSize = fitTextSize(text, font, options.watermarkSize, pageWidth * 0.92);
    var width = font.widthOfTextAtSize(text, fontSize);
    page.drawText(text, {
      x: pageWidth / 2 - width / 2,
      y: pageHeight / 2,
      size: fontSize,
      font: font,
      color: hexToRgb(options.stampColor),
      opacity: options.watermarkOpacity / 100,
      rotate: PDFLib.degrees(options.watermarkAngle)
    });
  }

  function drawPositionedText(page, text, pageWidth, pageHeight, options) {
    var margin = 34;
    var fontSize = fitTextSize(text, options.font, options.size, pageWidth - margin * 2);
    var width = options.font.widthOfTextAtSize(text, fontSize);
    var height = fontSize;
    var point = getPosition(pageWidth, pageHeight, width, height, options.position, margin);
    page.drawText(text, {
      x: point.x,
      y: point.y,
      size: fontSize,
      font: options.font,
      color: options.color,
      opacity: options.opacity
    });
  }

  function getPosition(pageWidth, pageHeight, boxWidth, boxHeight, position, margin) {
    var x = margin;
    var y = margin;

    if (position.indexOf("center") !== -1) {
      x = (pageWidth - boxWidth) / 2;
    } else if (position.indexOf("right") !== -1) {
      x = pageWidth - margin - boxWidth;
    }

    if (position.indexOf("top") !== -1) {
      y = pageHeight - margin - boxHeight;
    } else if (position === "center") {
      y = (pageHeight - boxHeight) / 2;
    }

    return {
      x: Math.max(margin / 2, x),
      y: Math.max(margin / 2, y)
    };
  }

  function scopeApplies(scope, item) {
    if (scope === "selected") {
      return item.selected;
    }
    if (scope === "current") {
      return item.id === state.activePageId;
    }
    return true;
  }

  function fitTextSize(text, font, requestedSize, maxWidth) {
    var width = font.widthOfTextAtSize(text, requestedSize);
    if (width <= maxWidth) {
      return requestedSize;
    }
    return Math.max(7, requestedSize * (maxWidth / Math.max(1, width)));
  }

  function getExportOptions() {
    return {
      outputName: $("#outputName").value.trim(),
      title: $("#metadataTitle").value.trim(),
      author: $("#metadataAuthor").value.trim(),
      compact: $("#compactPdf").checked,
      imagePageSize: $("#imagePageSize").value,
      imageFit: $("#imageFit").value,
      imageMargin: clampNumber($("#imageMargin").value, 0, 96, 36),
      imageQuality: clampNumber($("#imageQuality").value, 50, 100, 92),
      pageNumbers: $("#pageNumbers").checked,
      pageNumberPosition: $("#pageNumberPosition").value,
      stampText: $("#stampText").value.trim(),
      stampPosition: $("#stampPosition").value,
      stampScope: $("#stampScope").value,
      stampSize: clampNumber($("#stampSize").value, 8, 72, 18),
      stampColor: $("#stampColor").value,
      watermarkText: $("#watermarkText").value.trim(),
      watermarkScope: $("#watermarkScope").value,
      watermarkAngle: clampNumber($("#watermarkAngle").value, -90, 90, -35),
      watermarkSize: clampNumber($("#watermarkSize").value, 24, 160, 64),
      watermarkOpacity: clampNumber($("#watermarkOpacity").value, 5, 60, 18)
    };
  }

  function syncControls() {
    $("#imageMargin").value = clampNumber($("#imageMargin").value, 0, 96, 36);
    $("#imageQuality").value = clampNumber($("#imageQuality").value, 50, 100, 92);
    $("#stampSize").value = clampNumber($("#stampSize").value, 8, 72, 18);
    $("#watermarkSize").value = clampNumber($("#watermarkSize").value, 24, 160, 64);
    $("#watermarkOpacity").value = clampNumber($("#watermarkOpacity").value, 5, 60, 18);

    $("#imageMarginValue").textContent = $("#imageMargin").value + " pt";
    $("#imageQualityValue").textContent = $("#imageQuality").value + "%";
    $("#stampSizeValue").textContent = $("#stampSize").value + " pt";
    $("#watermarkSizeValue").textContent = $("#watermarkSize").value + " pt";
    $("#watermarkOpacityValue").textContent = $("#watermarkOpacity").value + "%";

    var selectedCount = getSelectedPages().length;
    var hasPages = state.pages.length > 0 && !state.busy;
    $("#downloadPdf").disabled = !hasPages;
    $("#downloadSelected").disabled = !hasPages || selectedCount === 0;
    $("#downloadSplit").disabled = !hasPages;

    $$(".batch-bar [data-batch]").forEach(function (button) {
      var action = button.dataset.batch;
      var needsSelection = action === "rotate-left" || action === "rotate-right" || action === "remove";
      button.disabled = !hasPages || (needsSelection && selectedCount === 0);
    });
  }

  function updateStats() {
    $("#fileCount").textContent = Object.keys(state.sources).length;
    $("#pageCount").textContent = state.pages.length;
    $("#selectedCount").textContent = getSelectedPages().length;
    $("#totalSize").textContent = formatBytes(state.totalInputBytes);
  }

  function resetAll() {
    Object.keys(state.sources).forEach(function (id) {
      var source = state.sources[id];
      if (source.objectUrl) {
        URL.revokeObjectURL(source.objectUrl);
      }
      if (source.pdfjsDoc && source.pdfjsDoc.destroy) {
        source.pdfjsDoc.destroy();
      }
    });
    state.sources = {};
    state.pages = [];
    state.activePageId = "";
    state.totalInputBytes = 0;
    $("#pdfForm").reset();
    $("#outputName").value = "pdffreely-document.pdf";
    setStatus("Ready");
    renderPages();
  }

  function createZip(files) {
    var localParts = [];
    var centralParts = [];
    var offset = 0;
    var encoder = new TextEncoder();

    files.forEach(function (file) {
      var nameBytes = encoder.encode(file.name);
      var data = file.data instanceof Uint8Array ? file.data : new Uint8Array(file.data);
      var crc = crc32(data);
      var local = new Uint8Array(30 + nameBytes.length);
      writeUint32(local, 0, 0x04034b50);
      writeUint16(local, 4, 20);
      writeUint16(local, 6, 0x0800);
      writeUint16(local, 8, 0);
      writeDosTimeDate(local, 10, new Date());
      writeUint32(local, 14, crc);
      writeUint32(local, 18, data.length);
      writeUint32(local, 22, data.length);
      writeUint16(local, 26, nameBytes.length);
      writeUint16(local, 28, 0);
      local.set(nameBytes, 30);
      localParts.push(local, data);

      var central = new Uint8Array(46 + nameBytes.length);
      writeUint32(central, 0, 0x02014b50);
      writeUint16(central, 4, 20);
      writeUint16(central, 6, 20);
      writeUint16(central, 8, 0x0800);
      writeUint16(central, 10, 0);
      writeDosTimeDate(central, 12, new Date());
      writeUint32(central, 16, crc);
      writeUint32(central, 20, data.length);
      writeUint32(central, 24, data.length);
      writeUint16(central, 28, nameBytes.length);
      writeUint16(central, 30, 0);
      writeUint16(central, 32, 0);
      writeUint16(central, 34, 0);
      writeUint16(central, 36, 0);
      writeUint32(central, 38, 0);
      writeUint32(central, 42, offset);
      central.set(nameBytes, 46);
      centralParts.push(central);

      offset += local.length + data.length;
    });

    var centralSize = centralParts.reduce(function (sum, part) {
      return sum + part.length;
    }, 0);
    var end = new Uint8Array(22);
    writeUint32(end, 0, 0x06054b50);
    writeUint16(end, 8, files.length);
    writeUint16(end, 10, files.length);
    writeUint32(end, 12, centralSize);
    writeUint32(end, 16, offset);
    writeUint16(end, 20, 0);

    return concatUint8(localParts.concat(centralParts, [end]));
  }

  function crc32(data) {
    if (!crcTable) {
      crcTable = [];
      for (var n = 0; n < 256; n += 1) {
        var c = n;
        for (var k = 0; k < 8; k += 1) {
          c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        }
        crcTable[n] = c >>> 0;
      }
    }
    var crc = 0xffffffff;
    for (var index = 0; index < data.length; index += 1) {
      crc = crcTable[(crc ^ data[index]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function writeDosTimeDate(target, offset, date) {
    var time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
    var day = date.getDate();
    var month = date.getMonth() + 1;
    var year = Math.max(1980, date.getFullYear()) - 1980;
    var dosDate = (year << 9) | (month << 5) | day;
    writeUint16(target, offset, time);
    writeUint16(target, offset + 2, dosDate);
  }

  function writeUint16(target, offset, value) {
    target[offset] = value & 0xff;
    target[offset + 1] = (value >>> 8) & 0xff;
  }

  function writeUint32(target, offset, value) {
    target[offset] = value & 0xff;
    target[offset + 1] = (value >>> 8) & 0xff;
    target[offset + 2] = (value >>> 16) & 0xff;
    target[offset + 3] = (value >>> 24) & 0xff;
  }

  function concatUint8(parts) {
    var length = parts.reduce(function (sum, part) {
      return sum + part.length;
    }, 0);
    var output = new Uint8Array(length);
    var offset = 0;
    parts.forEach(function (part) {
      output.set(part, offset);
      offset += part.length;
    });
    return output;
  }

  function withTimeout(promise, timeoutMs, message) {
    return new Promise(function (resolve, reject) {
      var timeout = window.setTimeout(function () {
        reject(createTimeoutError(message));
      }, timeoutMs);

      promise.then(function (value) {
        window.clearTimeout(timeout);
        resolve(value);
      }, function (error) {
        window.clearTimeout(timeout);
        reject(error);
      });
    });
  }

  function createTimeoutError(message) {
    var error = new Error(message);
    error.name = "TimeoutError";
    return error;
  }

  function isTimeoutError(error) {
    return error && error.name === "TimeoutError";
  }

  function shouldRetryLibreOfficeConversion(error) {
    return isTimeoutError(error) || isTransientLibreOfficeError(error);
  }

  function isTransientLibreOfficeError(error) {
    var message = getErrorMessage(error).toLowerCase();
    return /call_indirect|signature|webassembly|runtimeerror|wasm|memory access|table index|worker|not initialized|drawviewshell/.test(message);
  }

  function versionedAssetUrl(url) {
    return url + (url.indexOf("?") === -1 ? "?" : "&") + "v=" + assetVersion;
  }

  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 500);
  }

  function getOutputFilename(suffix) {
    var raw = $("#outputName").value.trim() || "pdffreely-document.pdf";
    if (!suffix) {
      return raw.replace(/\.pdf$/i, "");
    }
    return raw.replace(/\.pdf$/i, "") + suffix;
  }

  function sanitizeBase(name) {
    return String(name || "pdffreely-document")
      .replace(/\.pdf$/i, "")
      .replace(/[^a-z0-9._-]+/gi, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 80) || "pdffreely-document";
  }

  function getExtension(name) {
    var match = /\.([a-z0-9]+)$/i.exec(name || "");
    return match ? match[1].toLowerCase() : "";
  }

  function toUint8Array(value) {
    if (value instanceof Uint8Array) {
      return value;
    }
    if (value instanceof ArrayBuffer) {
      return new Uint8Array(value);
    }
    return new Uint8Array(value);
  }

  function getSelectedPages() {
    return state.pages.filter(function (item) {
      return item.selected;
    });
  }

  function getSelectedPageIds() {
    return getSelectedPages().map(function (item) {
      return item.id;
    });
  }

  function getPageById(id) {
    return state.pages.find(function (item) {
      return item.id === id;
    });
  }

  function findPageIndex(id) {
    return state.pages.findIndex(function (item) {
      return item.id === id;
    });
  }

  function setBusy(isBusy, text, level) {
    state.busy = isBusy;
    setStatus(text, level);
    syncControls();
  }

  function setStatus(text, level) {
    var badge = $("#engineBadge");
    badge.textContent = text || "Ready";
    badge.classList.toggle("warn", level === "warn");
    badge.classList.toggle("danger", level === "danger");
  }

  function getErrorMessage(error) {
    var message = error && error.message ? error.message : String(error || "");
    if (/encrypted|password/i.test(message)) {
      return "Encrypted PDFs need a password before they can be edited.";
    }
    if (/invalid|parse|PDF/i.test(message)) {
      return message;
    }
    return message || "Unable to process this file.";
  }

  function hexToRgb(hex) {
    var value = String(hex || "#111827").replace("#", "");
    if (value.length === 3) {
      value = value.split("").map(function (char) {
        return char + char;
      }).join("");
    }
    var number = parseInt(value, 16);
    var red = (number >> 16) & 255;
    var green = (number >> 8) & 255;
    var blue = number & 255;
    return window.PDFLib.rgb(red / 255, green / 255, blue / 255);
  }

  function normalizeDegrees(value) {
    return ((value % 360) + 360) % 360;
  }

  function clampNumber(value, min, max, fallback) {
    var number = Number(value);
    if (!Number.isFinite(number)) {
      number = fallback;
    }
    return Math.min(max, Math.max(min, number));
  }

  function formatBytes(bytes) {
    if (!bytes) {
      return "0 KB";
    }
    var units = ["B", "KB", "MB", "GB"];
    var value = bytes;
    var unit = 0;
    while (value >= 1024 && unit < units.length - 1) {
      value /= 1024;
      unit += 1;
    }
    return (unit === 0 ? value : value.toFixed(value >= 10 ? 1 : 2)) + " " + units[unit];
  }

  function formatDuration(seconds) {
    var totalSeconds = Math.max(0, Number(seconds) || 0);
    var minutes = Math.floor(totalSeconds / 60);
    var remainder = totalSeconds % 60;
    if (!minutes) {
      return remainder + "s";
    }
    return minutes + "m " + String(remainder).padStart(2, "0") + "s";
  }

  function formatPixels(width, height) {
    return Math.round(width) + " x " + Math.round(height) + " px";
  }

  function cssEscape(value) {
    if (window.CSS && window.CSS.escape) {
      return window.CSS.escape(value);
    }
    return String(value).replace(/"/g, '\\"');
  }

  function escapeAttr(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function escapeText(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }
})();
