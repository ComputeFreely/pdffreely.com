if (!window.JSZip) {
  await import("./jszip.min.js");
}

if (!window.JSZip) {
  throw new Error("JSZip did not load.");
}

export default window.JSZip;
