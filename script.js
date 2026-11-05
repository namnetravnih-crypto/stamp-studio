(() => {
'use strict';

const $ = id => document.getElementById(id);
const MAX_SIDE = 1600;                       // максимальная сторона рабочего изображения, px
const DPI = 600;                             // разрешение выходного PNG
const TARGET_MM = 40;                        // диаметр печати в готовом файле
const TARGET_PX = Math.round(TARGET_MM / 25.4 * DPI); // 945 px

const el = {
  file: $('fileInput'), label: $('uploadLabel'), drop: $('dropZone'),
  workspace: document.querySelector('.workspace'),
  container: $('canvasContainer'), canvas: $('mainCanvas'), cursor: $('brushCursor'),
  status: $('statusBar').firstElementChild, fname: $('fileName'),
  removeText: $('removeTextCheck'), pText: $('paramText'),
  pBright: $('paramBrightness'), pChroma: $('paramChroma'),
  pShadow: $('paramShadows'), pSmooth: $('paramSmoothing'),
  btnMask: $('btnMagicCursor'), btnEraser: $('btnEraser'),
  brush: $('brushSize'), brushVal: $('valBrushSize'),
  undo: $('btnUndo'), round: $('btnMakeRound'),
  aging: $('agingCheck'), pAging: $('paramAging'),
  dl: $('btnDownload'), reset: $('btnReset')
};
const ctx = el.canvas.getContext('2d', { willReadFrequently: true });

/* ---------- состояние ---------- */
let origCanvas = null;       // загруженное изображение (уменьшенное до MAX_SIDE)
let src = null, srcPix = null, W = 0, H = 0;
let grid = null;             // карта фона (для подавления бликов и теней)
let base = null;             // {col, alpha, str} — результат автоматической обработки
let cur = null;              // альфа после ручных правок
let imgOut = null;
let edits = [], history = [];
let tool = null;
const sizes = { mask: 40, eraser: 15 };
let dispScale = 1;
let fade = null, fadeKey = '';
let stroke = null, lastPt = null;
let fileLabel = 'stamp';

const clamp = (v, a, b) => v < a ? a : (v > b ? b : v);
const setStatus = t => { el.status.textContent = t; };

/* ---------- параметры ---------- */
function P() {
  return {
    rm: el.removeText.checked, text: +el.pText.value,
    bright: +el.pBright.value, chroma: +el.pChroma.value,
    shadow: +el.pShadow.value, smooth: +el.pSmooth.value,
    aging: el.aging.checked, ageK: +el.pAging.value / 100
  };
}
function chromaThr(p) {
  return Math.max(8, 75 - p.chroma * 0.8) + (p.rm ? p.text * 0.3 : 0);
}

/* ---------- загрузка файлов ---------- */
let pdfjsPromise = null;
function loadPdfJs() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (!pdfjsPromise) {
    pdfjsPromise = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
      s.onload = () => {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc =
          'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        res(window.pdfjsLib);
      };
      s.onerror = () => { pdfjsPromise = null; rej(new Error('Не удалось загрузить модуль чтения PDF (нужен интернет).')); };
      document.head.appendChild(s);
    });
  }
  return pdfjsPromise;
}

async function pdfToCanvas(file) {
  const lib = await loadPdfJs();
  const data = await file.arrayBuffer();
  const pdf = await lib.getDocument({ data }).promise;
  const page = await pdf.getPage(1);
  const v1 = page.getViewport({ scale: 1 });
  const scale = MAX_SIDE / Math.max(v1.width, v1.height);
  const vp = page.getViewport({ scale });
  const c = document.createElement('canvas');
  c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  const cx = c.getContext('2d');
  cx.fillStyle = '#fff'; cx.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: cx, viewport: vp }).promise;
  return c;
}

function imageToCanvas(file) {
  return new Promise((res, rej) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const k = Math.min(1, MAX_SIDE / Math.max(img.naturalWidth, img.naturalHeight));
      const c = document.createElement('canvas');
      c.width = Math.max(1, Math.round(img.naturalWidth * k));
      c.height = Math.max(1, Math.round(img.naturalHeight * k));
      const cx = c.getContext('2d');
      cx.fillStyle = '#fff'; cx.fillRect(0, 0, c.width, c.height);
      cx.imageSmoothingQuality = 'high';
      cx.drawImage(img, 0, 0, c.width, c.height);
      URL.revokeObjectURL(url);
      res(c);
    };
    img.onerror = () => { URL.revokeObjectURL(url); rej(new Error('Не удалось открыть изображение (формат не поддерживается).')); };
    img.src = url;
  });
}

async function handleFile(file) {
  if (!file) return;
  const isPdf = file.type === 'application/pdf' || /\.pdf$/i.test(file.name);
  const isImg = file.type.startsWith('image/');
  if (!isPdf && !isImg) { setStatus('Нужен файл изображения или PDF.'); return; }
  setStatus('Загрузка…');
  try {
    const c = isPdf ? await pdfToCanvas(file) : await imageToCanvas(file);
    origCanvas = c;
    fileLabel = file.name.replace(/\.[^.]+$/, '') || 'stamp';
    el.fname.textContent = file.name + (isPdf ? ' (1-я страница)' : '');
    resetEverything();
    setStatus('Файл загружен. Настройте параметры справа налево сверху вниз.');
  } catch (err) {
    setStatus('Ошибка: ' + err.message);
  }
}

/* ---------- подготовка источника ---------- */
function setSource(canvas) {
  src = canvas; W = canvas.width; H = canvas.height;
  srcPix = canvas.getContext('2d').getImageData(0, 0, W, H).data;
  el.canvas.width = W; el.canvas.height = H;
  imgOut = ctx.createImageData(W, H);
  fade = null; fadeKey = '';
  buildGrid();
  el.drop.style.display = 'none';
  el.container.style.display = 'block';
  layout();
  process();
}

function layout() {
  if (!src) return;
  const aw = el.workspace.clientWidth - 70, ah = el.workspace.clientHeight - 70;
  dispScale = Math.max(0.05, Math.min(aw / W, ah / H));
  el.canvas.style.width = Math.round(W * dispScale) + 'px';
  el.canvas.style.height = Math.round(H * dispScale) + 'px';
}

/* Карта фона: по блокам берём цвет яркой части (80-й перцентиль) — это бумага с бликами/тенями */
function buildGrid() {
  const B = Math.max(24, Math.round(Math.min(W, H) / 14));
  const gw = Math.ceil(W / B), gh = Math.ceil(H / B);
  let g = new Float32Array(gw * gh * 3);
  const hist = new Uint32Array(256);
  for (let gy = 0; gy < gh; gy++) {
    for (let gx = 0; gx < gw; gx++) {
      const x0 = gx * B, y0 = gy * B, x1 = Math.min(W, x0 + B), y1 = Math.min(H, y0 + B);
      hist.fill(0); let cnt = 0;
      for (let y = y0; y < y1; y += 2) for (let x = x0; x < x1; x += 2) {
        const i = (y * W + x) * 4;
        hist[(srcPix[i] * 77 + srcPix[i + 1] * 150 + srcPix[i + 2] * 29) >> 8]++; cnt++;
      }
      let acc = 0, p = 255; const target = cnt * 0.2;
      for (let l = 255; l >= 0; l--) { acc += hist[l]; if (acc >= target) { p = l; break; } }
      let sr = 0, sg = 0, sb = 0, n = 0;
      for (let y = y0; y < y1; y += 2) for (let x = x0; x < x1; x += 2) {
        const i = (y * W + x) * 4;
        if (((srcPix[i] * 77 + srcPix[i + 1] * 150 + srcPix[i + 2] * 29) >> 8) >= p) {
          sr += srcPix[i]; sg += srcPix[i + 1]; sb += srcPix[i + 2]; n++;
        }
      }
      const o = (gy * gw + gx) * 3;
      if (n) { g[o] = sr / n; g[o + 1] = sg / n; g[o + 2] = sb / n; } else { g[o] = g[o + 1] = g[o + 2] = 255; }
    }
  }
  for (let pass = 0; pass < 2; pass++) {
    const t = new Float32Array(g.length);
    for (let gy = 0; gy < gh; gy++) for (let gx = 0; gx < gw; gx++) {
      for (let k = 0; k < 3; k++) {
        let s = 0, c = 0;
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          const yy = gy + dy, xx = gx + dx;
          if (yy < 0 || yy >= gh || xx < 0 || xx >= gw) continue;
          s += g[(yy * gw + xx) * 3 + k]; c++;
        }
        t[(gy * gw + gx) * 3 + k] = s / c;
      }
    }
    g = t;
  }
  grid = { B, gw, gh, g };
}

/* ---------- основная обработка ---------- */
function process() {
  if (!src) return;
  const p = P(), N = W * H;
  const col = new Uint8ClampedArray(N * 4);
  const a = new Uint8ClampedArray(N), st = new Uint8ClampedArray(N);
  const m = Math.min(1, p.shadow / 50);
  const thr = chromaThr(p);
  const den = thr * 0.6 + 10;
  const satCut = 0.1 + p.text / 100 * 0.3;
  const { g, gw, gh, B } = grid;

  const xa = new Int32Array(W), xb = new Int32Array(W), xf = new Float32Array(W);
  for (let x = 0; x < W; x++) {
    const f = clamp((x + 0.5) / B - 0.5, 0, gw - 1);
    const i0 = Math.min(Math.floor(f), gw - 1);
    xa[x] = i0; xb[x] = Math.min(i0 + 1, gw - 1); xf[x] = f - i0;
  }

  for (let y = 0; y < H; y++) {
    const f = clamp((y + 0.5) / B - 0.5, 0, gh - 1);
    const y0 = Math.min(Math.floor(f), gh - 1), y1 = Math.min(y0 + 1, gh - 1), yf = f - y0;
    const rA = y0 * gw, rB = y1 * gw;
    for (let x = 0; x < W; x++) {
      const p00 = (rA + xa[x]) * 3, p01 = (rA + xb[x]) * 3;
      const p10 = (rB + xa[x]) * 3, p11 = (rB + xb[x]) * 3;
      const fx = xf[x];
      const idx = y * W + x, i = idx * 4;
      let r = srcPix[i], gg = srcPix[i + 1], b = srcPix[i + 2];
      if (m > 0) {
        const bgR = Math.max(20, (g[p00] * (1 - fx) + g[p01] * fx) * (1 - yf) + (g[p10] * (1 - fx) + g[p11] * fx) * yf);
        const bgG = Math.max(20, (g[p00 + 1] * (1 - fx) + g[p01 + 1] * fx) * (1 - yf) + (g[p10 + 1] * (1 - fx) + g[p11 + 1] * fx) * yf);
        const bgB = Math.max(20, (g[p00 + 2] * (1 - fx) + g[p01 + 2] * fx) * (1 - yf) + (g[p10 + 2] * (1 - fx) + g[p11 + 2] * fx) * yf);
        r += (Math.min(255, r * 255 / bgR) - r) * m;
        gg += (Math.min(255, gg * 255 / bgG) - gg) * m;
        b += (Math.min(255, b * 255 / bgB) - b) * m;
      }
      const mx = Math.max(r, gg, b), mn = Math.min(r, gg, b), ch = mx - mn;
      const lum = 0.299 * r + 0.587 * gg + 0.114 * b;
      let al = clamp((ch - thr) / den, 0, 1);
      let s = ch;
      if (p.rm) {
        const sat = mx > 0 ? ch / mx : 0;
        if (sat < satCut) al *= sat / satCut;
      } else {
        const d = clamp((p.bright - lum) / 30, 0, 1);
        if (d > al) al = d;
        s += Math.max(0, p.bright - lum);
      }
      a[idx] = al * 255;
      st[idx] = s;
      col[i] = r; col[i + 1] = gg; col[i + 2] = b; col[i + 3] = 255;
    }
  }

  removeSpecks(a, Math.max(8, Math.round(N * 0.000008)));
  base = { col, alpha: boxBlur(a, p.smooth), str: st };
  render();
}

function removeSpecks(a, minSize) {
  const N = W * H, vis = new Uint8Array(N), q = new Int32Array(N);
  for (let s = 0; s < N; s++) {
    if (vis[s] || a[s] < 40) continue;
    let head = 0, tail = 0; q[tail++] = s; vis[s] = 1;
    while (head < tail) {
      const i = q[head++], x = i % W, y = (i - x) / W;
      for (let dy = -1; dy <= 1; dy++) {
        const ny = y + dy; if (ny < 0 || ny >= H) continue;
        for (let dx = -1; dx <= 1; dx++) {
          const nx = x + dx; if (nx < 0 || nx >= W) continue;
          const j = ny * W + nx;
          if (!vis[j] && a[j] >= 40) { vis[j] = 1; q[tail++] = j; }
        }
      }
    }
    if (tail < minSize) for (let k = 0; k < tail; k++) a[q[k]] = 0;
  }
}

function boxBlur(a, r) {
  if (!r) return a;
  const win = 2 * r + 1, t = new Uint8ClampedArray(a.length), o = new Uint8ClampedArray(a.length);
  for (let y = 0; y < H; y++) {
    const row = y * W; let sum = 0;
    for (let x = -r; x <= r; x++) sum += a[row + clamp(x, 0, W - 1)];
    for (let x = 0; x < W; x++) {
      t[row + x] = sum / win;
      sum += a[row + Math.min(x + r + 1, W - 1)] - a[row + Math.max(x - r, 0)];
    }
  }
  for (let x = 0; x < W; x++) {
    let sum = 0;
    for (let y = -r; y <= r; y++) sum += t[clamp(y, 0, H - 1) * W + x];
    for (let y = 0; y < H; y++) {
      o[y * W + x] = sum / win;
      sum += t[Math.min(y + r + 1, H - 1) * W + x] - t[Math.max(y - r, 0) * W + x];
    }
  }
  return o;
}

/* ---------- ручные правки ---------- */
function applyCircle(a, s, x, y, r, type, mThr) {
  const x0 = Math.max(0, Math.floor(x - r)), x1 = Math.min(W - 1, Math.ceil(x + r));
  const y0 = Math.max(0, Math.floor(y - r)), y1 = Math.min(H - 1, Math.ceil(y + r));
  const r2 = r * r;
  for (let yy = y0; yy <= y1; yy++) {
    const dy = yy - y;
    for (let xx = x0; xx <= x1; xx++) {
      const dx = xx - x;
      if (dx * dx + dy * dy <= r2) {
        const i = yy * W + xx;
        if (type === 'eraser') a[i] = 0;
        else if (s[i] < mThr) a[i] = 0;   // круг-маска: остаётся только насыщенный оттиск
      }
    }
  }
}

function render() {
  if (!base) return;
  const p = P();
  cur = base.alpha.slice();
  const mThr = chromaThr(p) * 1.8 + 10;
  for (const e of edits) for (const pt of e.pts) applyCircle(cur, base.str, pt[0], pt[1], e.r, e.type, mThr);
  paint(p);
}

/* ---------- состаривание ---------- */
function makeFade() {
  const N = W * H, z = new Float32Array(N), rnd = new Uint8Array(N);
  let s = 123456789;
  const rng = () => { s ^= s << 13; s ^= s >>> 17; s ^= s << 5; return (s >>> 0) / 4294967296; };
  for (const [n, wt] of [[5, 0.5], [13, 0.3], [37, 0.2]]) {
    const gw = n + 2, gh = Math.ceil(n * H / W) + 2;
    const gr = new Float32Array(gw * gh);
    for (let i = 0; i < gr.length; i++) gr[i] = rng();
    for (let y = 0; y < H; y++) {
      const fy = y * n / W, y0 = Math.floor(fy); let ty = fy - y0; ty = ty * ty * (3 - 2 * ty);
      for (let x = 0; x < W; x++) {
        const fx = x * n / W, x0 = Math.floor(fx); let tx = fx - x0; tx = tx * tx * (3 - 2 * tx);
        const a0 = gr[y0 * gw + x0], a1 = gr[y0 * gw + x0 + 1];
        const b0 = gr[(y0 + 1) * gw + x0], b1 = gr[(y0 + 1) * gw + x0 + 1];
        z[y * W + x] += wt * ((a0 + (a1 - a0) * tx) * (1 - ty) + (b0 + (b1 - b0) * tx) * ty);
      }
    }
  }
  for (let i = 0; i < N; i++) {
    const t = clamp((z[i] - 0.42) / 0.30, 0, 1);
    z[i] = t * t * (3 - 2 * t);
    rnd[i] = (rng() * 255) | 0;
  }
  return { z, rnd };
}

function paint(p) {
  const out = imgOut.data, N = W * H;
  out.set(base.col);
  if (p.aging) {
    const key = W + 'x' + H;
    if (!fade || fadeKey !== key) { fade = makeFade(); fadeKey = key; }
    const { z, rnd } = fade, k = p.ageK;
    for (let i = 0; i < N; i++) {
      let a = cur[i];
      if (a) {
        let pale = k * (0.85 * z[i] + 0.12 * rnd[i] / 255) + 0.2 * k;
        if (rnd[i] < k * 14 * (0.3 + z[i])) pale += 0.7;
        a *= 1 - Math.min(1, pale);
      }
      out[i * 4 + 3] = a;
    }
  } else {
    for (let i = 0; i < N; i++) out[i * 4 + 3] = cur[i];
  }
  ctx.putImageData(imgOut, 0, 0);
}

let paintQueued = false;
function schedulePaint() {
  if (paintQueued) return; paintQueued = true;
  requestAnimationFrame(() => { paintQueued = false; if (base) paint(P()); });
}
let procQueued = false;
function scheduleProcess() {
  if (procQueued) return; procQueued = true;
  requestAnimationFrame(() => { procQueued = false; process(); });
}
let renderQueued = false;
function scheduleRender() {
  if (renderQueued) return; renderQueued = true;
  requestAnimationFrame(() => { renderQueued = false; render(); });
}

/* ---------- круглая печать (выравнивание эллипса в круг) ---------- */
function makeRound() {
  if (!base) { setStatus('Сначала загрузите файл.'); return; }
  let n = 0, sx = 0, sy = 0, sxx = 0, syy = 0, sxy = 0;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    if (cur[y * W + x] > 128) { n++; sx += x; sy += y; sxx += x * x; syy += y * y; sxy += x * y; }
  }
  if (n < 50) { setStatus('Печать на изображении не найдена — подстройте параметры.'); return; }
  const mx = sx / n, my = sy / n;
  const cxx = sxx / n - mx * mx, cyy = syy / n - my * my, cxy = sxy / n - mx * my;
  const th = 0.5 * Math.atan2(2 * cxy, cxx - cyy);
  const c = Math.cos(th), s = Math.sin(th);
  let umin = 1e9, umax = -1e9, vmin = 1e9, vmax = -1e9;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    if (cur[y * W + x] > 128) {
      const dx = x - mx, dy = y - my;
      const u = dx * c + dy * s, v = -dx * s + dy * c;
      if (u < umin) umin = u; if (u > umax) umax = u;
      if (v < vmin) vmin = v; if (v > vmax) vmax = v;
    }
  }
  const eu = umax - umin, ev = vmax - vmin;
  const cu = (umax + umin) / 2, cv = (vmax + vmin) / 2;
  const cx = mx + cu * c - cv * s, cy = my + cu * s + cv * c;
  const R = (eu + ev) / 4;
  const q = Math.min(1, 2400 / (2 * R * 1.08));
  const S = Math.ceil(2 * R * 1.08 * q), Rout = R * q;
  const out = new ImageData(S, S), od = out.data;
  for (let y = 0; y < S; y++) for (let x = 0; x < S; x++) {
    const px = x - S / 2 + 0.5, py = y - S / 2 + 0.5;
    const u = px * (eu / 2) / Rout, v = py * (ev / 2) / Rout;
    const sx0 = clamp(cx + u * c - v * s, 0, W - 1.001);
    const sy0 = clamp(cy + u * s + v * c, 0, H - 1.001);
    const ix = Math.floor(sx0), iy = Math.floor(sy0), fx = sx0 - ix, fy = sy0 - iy;
    const i00 = (iy * W + ix) * 4, i01 = i00 + 4, i10 = i00 + W * 4, i11 = i10 + 4;
    const o = (y * S + x) * 4;
    for (let k = 0; k < 3; k++) {
      od[o + k] = (srcPix[i00 + k] * (1 - fx) + srcPix[i01 + k] * fx) * (1 - fy) +
                  (srcPix[i10 + k] * (1 - fx) + srcPix[i11 + k] * fx) * fy;
    }
    od[o + 3] = 255;
  }
  const c2 = document.createElement('canvas'); c2.width = S; c2.height = S;
  c2.getContext('2d').putImageData(out, 0, 0);
  history.push({ type: 'round', prevSrc: src, prevEdits: edits });
  edits = [];
  setSource(c2);
  setStatus('Печать приведена к кругу. Ручные правки сброшены — лучше применять эту функцию до ручной чистки.');
}

/* ---------- отмена ---------- */
function undo() {
  const h = history.pop();
  if (!h) { setStatus('Нечего отменять.'); return; }
  if (h.type === 'edit') { edits.pop(); render(); }
  else if (h.type === 'round') { edits = h.prevEdits; setSource(h.prevSrc); }
  setStatus('Последнее действие отменено.');
}

function resetEverything() {
  if (!origCanvas) return;
  edits = []; history = [];
  el.removeText.checked = false; el.aging.checked = false;
  el.pText.value = 50; el.pBright.value = 150; el.pChroma.value = 35;
  el.pShadow.value = 40; el.pSmooth.value = 2; el.pAging.value = 60;
  updateLabels();
  const c = document.createElement('canvas');
  c.width = origCanvas.width; c.height = origCanvas.height;
  c.getContext('2d').drawImage(origCanvas, 0, 0);
  setSource(c);
}

/* ---------- скачивание ---------- */
const crcT = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
function crc32(buf) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) c = crcT[(c ^ buf[i]) & 255] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
function addDpi(bytes, dpi) {            // вставка блока pHYs сразу после IHDR (смещение 33)
  const ppm = Math.round(dpi / 0.0254);
  const chunk = new Uint8Array(21), dv = new DataView(chunk.buffer);
  dv.setUint32(0, 9); chunk.set([0x70, 0x48, 0x59, 0x73], 4);
  dv.setUint32(8, ppm); dv.setUint32(12, ppm); chunk[16] = 1;
  dv.setUint32(17, crc32(chunk.subarray(4, 17)));
  const out = new Uint8Array(bytes.length + 21);
  out.set(bytes.subarray(0, 33)); out.set(chunk, 33); out.set(bytes.subarray(33), 54);
  return out;
}

function download() {
  if (!base) { setStatus('Сначала загрузите файл.'); return; }
  const d = imgOut.data;
  let x0 = W, y0 = H, x1 = -1, y1 = -1;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    if (d[(y * W + x) * 4 + 3] > 8) {
      if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
    }
  }
  if (x1 < 0) { setStatus('Результат пуст — нечего скачивать.'); return; }
  const bw = x1 - x0 + 1, bh = y1 - y0 + 1;
  const k = TARGET_PX / Math.max(bw, bh);
  const out = document.createElement('canvas'); out.width = TARGET_PX; out.height = TARGET_PX;
  const oc = out.getContext('2d');
  oc.imageSmoothingEnabled = true; oc.imageSmoothingQuality = 'high';
  const dw = bw * k, dh = bh * k;
  oc.drawImage(el.canvas, x0, y0, bw, bh, (TARGET_PX - dw) / 2, (TARGET_PX - dh) / 2, dw, dh);
  out.toBlob(async blob => {
    const bytes = addDpi(new Uint8Array(await blob.arrayBuffer()), DPI);
    const url = URL.createObjectURL(new Blob([bytes], { type: 'image/png' }));
    const a = document.createElement('a');
    a.href = url; a.download = fileLabel + '_stamp_40mm.png';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    const ratio = Math.max(bw, bh) / Math.min(bw, bh);
    setStatus('PNG сохранён: ' + TARGET_PX + '×' + TARGET_PX + ' px, ' + DPI + ' dpi, диаметр 40 мм.' +
      (ratio > 1.05 ? ' Печать вытянута — нажмите «Круглая печать» для точного круга.' : ''));
  }, 'image/png');
}

/* ---------- интерфейс ---------- */
function updateLabels() {
  $('val1').textContent = el.pBright.value;
  $('val2').textContent = el.pChroma.value;
  $('val3').textContent = el.pShadow.value;
  $('val4').textContent = el.pSmooth.value;
  $('valText').textContent = el.pText.value;
  $('valAging').textContent = el.pAging.value;
  el.brushVal.textContent = el.brush.value + ' px';
}

function setTool(t) {
  tool = tool === t ? null : t;
  el.btnMask.classList.toggle('tool-active', tool === 'mask');
  el.btnEraser.classList.toggle('tool-active', tool === 'eraser');
  if (tool === 'eraser') { el.brush.min = 1; el.brush.max = 30; }
  else { el.brush.min = 5; el.brush.max = 180; }
  if (tool) el.brush.value = sizes[tool];
  updateLabels();
  el.canvas.style.cursor = tool ? 'none' : 'default';
  if (!tool) el.cursor.style.display = 'none';
}

function moveCursor(e) {
  if (!tool) return;
  const cr = el.container.getBoundingClientRect();
  const d = Math.max(4, sizes[tool] * dispScale);
  el.cursor.style.display = 'block';
  el.cursor.style.width = d + 'px'; el.cursor.style.height = d + 'px';
  el.cursor.style.left = (e.clientX - cr.left - el.container.clientLeft) + 'px';
  el.cursor.style.top = (e.clientY - cr.top - el.container.clientTop) + 'px';
  el.cursor.style.borderColor = tool === 'eraser' ? 'rgba(255,120,120,.95)' : 'rgba(255,255,255,.85)';
}

function imgPos(e) {
  const r = el.canvas.getBoundingClientRect();
  return [(e.clientX - r.left) * W / r.width, (e.clientY - r.top) * H / r.height];
}

function addPoint(e) {
  const [x, y] = imgPos(e);
  const mThr = chromaThr(P()) * 1.8 + 10;
  const put = (px, py) => { stroke.pts.push([px, py]); applyCircle(cur, base.str, px, py, stroke.r, stroke.type, mThr); };
  if (lastPt) {
    const dx = x - lastPt[0], dy = y - lastPt[1], dist = Math.hypot(dx, dy);
    const step = Math.max(1, stroke.r * 0.5), n = Math.floor(dist / step);
    for (let i = 1; i <= n; i++) put(lastPt[0] + dx * i / n, lastPt[1] + dy * i / n);
    if (n === 0) put(x, y);
  } else put(x, y);
  lastPt = [x, y];
  schedulePaint();
}

el.canvas.addEventListener('pointerdown', e => {
  if (!tool || !base) return;
  e.preventDefault();
  el.canvas.setPointerCapture(e.pointerId);
  stroke = { type: tool, r: Math.max(0.75, sizes[tool] / 2), pts: [] };
  lastPt = null;
  addPoint(e);
});
el.canvas.addEventListener('pointermove', e => {
  moveCursor(e);
  if (stroke) addPoint(e);
});
const endStroke = () => {
  if (!stroke) return;
  if (stroke.pts.length) { edits.push(stroke); history.push({ type: 'edit' }); }
  stroke = null; lastPt = null;
};
el.canvas.addEventListener('pointerup', endStroke);
el.canvas.addEventListener('pointercancel', endStroke);
el.canvas.addEventListener('pointerleave', () => { if (!stroke) el.cursor.style.display = 'none'; });
el.canvas.addEventListener('pointerenter', moveCursor);

el.brush.addEventListener('input', () => {
  if (tool) sizes[tool] = +el.brush.value;
  updateLabels();
});

[el.pBright, el.pChroma, el.pShadow, el.pSmooth, el.pText].forEach(s =>
  s.addEventListener('input', () => { updateLabels(); scheduleProcess(); }));
el.removeText.addEventListener('change', scheduleProcess);
el.aging.addEventListener('change', () => { if (base) scheduleRender(); });
el.pAging.addEventListener('input', () => { updateLabels(); if (base) scheduleRender(); });

el.btnMask.addEventListener('click', () => setTool('mask'));
el.btnEraser.addEventListener('click', () => setTool('eraser'));
el.undo.addEventListener('click', undo);
el.round.addEventListener('click', makeRound);
el.dl.addEventListener('click', download);
el.reset.addEventListener('click', () => {
  if (!origCanvas) { setStatus('Файл ещё не загружен.'); return; }
  resetEverything(); setStatus('Всё сброшено к исходному изображению.');
});

document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); undo(); }
});

/* загрузка: кнопка */
el.file.addEventListener('change', () => {
  const f = el.file.files[0];
  handleFile(f).finally(() => { el.file.value = ''; });
});

/* загрузка: перетаскивание (кнопка, окно-заглушка и вся рабочая область) */
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => e.preventDefault());
function bindDrop(node, hoverNode) {
  node.addEventListener('dragover', e => { e.preventDefault(); hoverNode.classList.add('dragover'); });
  node.addEventListener('dragleave', () => hoverNode.classList.remove('dragover'));
  node.addEventListener('drop', e => {
    e.preventDefault(); e.stopPropagation();
    hoverNode.classList.remove('dragover');
    handleFile(e.dataTransfer.files && e.dataTransfer.files[0]);
  });
}
bindDrop(el.label, el.label);
bindDrop(el.drop, el.drop);
bindDrop(el.workspace, el.drop);
el.drop.addEventListener('click', () => el.file.click());

window.addEventListener('resize', layout);
updateLabels();
setStatus('Готов к работе. Загрузите изображение или PDF.');
})();
