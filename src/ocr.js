// In-browser OCR (Tesseract.js, traditional Chinese + English). Worker, wasm core and language
// data are served from /ocr (public/ocr), so nothing is sent to a third party and no API key is needed.
import { createWorker } from "tesseract.js";

let workerPromise = null;
let progressCb = null;

function getWorker() {
  if (!workerPromise) {
    workerPromise = createWorker(["chi_tra", "eng"], 1, {
      workerPath: "/ocr/worker.min.js",
      corePath: "/ocr/core",
      langPath: "/ocr/lang",
      logger: (m) => { if (progressCb) progressCb(m); },
    }).catch((e) => {
      workerPromise = null;
      throw e;
    });
  }
  return workerPromise;
}

export async function terminateOcr() {
  if (!workerPromise) return;
  try { (await workerPromise).terminate(); } catch (e) { /* ignore */ }
  workerPromise = null;
}

// Phone photos are big, skewed in contrast and often rotated by EXIF: normalise size, orientation and contrast.
async function loadBitmap(file) {
  try {
    return await createImageBitmap(file, { imageOrientation: "from-image" });
  } catch (e) {
    try {
      return await createImageBitmap(file);
    } catch (e2) {
      const url = URL.createObjectURL(file);
      try {
        const img = new Image();
        await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = reject; img.src = url; });
        return img;
      } finally {
        URL.revokeObjectURL(url);
      }
    }
  }
}

export async function preprocessImage(file, targetLongSide = 2000) {
  const bmp = await loadBitmap(file);
  const srcW = bmp.naturalWidth || bmp.width;
  const srcH = bmp.naturalHeight || bmp.height;
  const longSide = Math.max(srcW, srcH);
  const scale = Math.min(2, targetLongSide / longSide);
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(srcW * scale);
  canvas.height = Math.round(srcH * scale);
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height);
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const d = img.data;
  const hist = new Uint32Array(256);
  for (let i = 0; i < d.length; i += 4) {
    const g = (d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) | 0;
    d[i] = d[i + 1] = d[i + 2] = g;
    hist[g]++;
  }
  const total = canvas.width * canvas.height;
  let acc = 0, lo = 0, hi = 255;
  for (let v = 0; v < 256; v++) { acc += hist[v]; if (acc >= total * 0.02) { lo = v; break; } }
  acc = 0;
  for (let v = 255; v >= 0; v--) { acc += hist[v]; if (acc >= total * 0.02) { hi = v; break; } }
  const range = Math.max(40, hi - lo);
  for (let i = 0; i < d.length; i += 4) {
    const v = Math.max(0, Math.min(255, ((d[i] - lo) * 255) / range));
    d[i] = d[i + 1] = d[i + 2] = v;
  }
  ctx.putImageData(img, 0, 0);
  return canvas;
}

/** @returns {Promise<string>} recognised text, one visual line per line */
export async function recognizeText(file, onProgress) {
  progressCb = onProgress || null;
  try {
    const canvas = await preprocessImage(file);
    const worker = await getWorker();
    const { data } = await worker.recognize(canvas);
    return data.text || "";
  } finally {
    progressCb = null;
  }
}
