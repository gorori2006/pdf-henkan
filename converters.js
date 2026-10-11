// 形式の判定と、各形式を「PDFのページ」に変換する処理。
// すべてブラウザの中で行い、ファイルを外部に送ることはない。
//
// 変換結果のページは次の2種類:
//   { type: 'image', format: 'jpg'|'png', bytes, pageW, pageH, x, y, w, h }  … 画像を貼ったページ（単位は pt、y は上から）
//   { type: 'pdf', doc }                                                      … 元から PDF のページ（pdf-lib の文書）
(function (global) {
'use strict';

const PX_PER_PT = 96 / 72;
const A4 = { w: 595.28, h: 841.89 };          // pt
const MARGIN_PT = 36;                          // ページの余白（約 12.7mm）
const RENDER_SCALE = 2;                        // 文書を画像にするときの解像度（2 = 約 192dpi）
const MAX_IMAGE_PX = 3000;                     // 写真はこれより大きければ縮小する（メモリ節約）
const MAX_PAGES = 300;                         // 1ファイルあたりのページ数の上限
const MIN_SCALE = 0.6;                         // 横に広い文書はここまで縮小し、それでも入らなければ横にもページを分ける
const EMU_PER_PX = 9525;                       // PowerPoint の座標単位

// ---------- 形式の判定 ----------

const KINDS = {
  jpg: 'image', jpeg: 'image', png: 'image', gif: 'image', webp: 'image', bmp: 'image',
  heic: 'heic', heif: 'heic',
  pdf: 'pdf',
  txt: 'text', text: 'text', log: 'text',
  md: 'markdown', markdown: 'markdown',
  csv: 'csv', tsv: 'csv',
  xlsx: 'excel', xlsm: 'excel', xls: 'excel', ods: 'excel',
  docx: 'word',
  pptx: 'powerpoint',
  html: 'html', htm: 'html',
};

const LABELS = {
  image: '画像', heic: '画像（HEIC）', pdf: 'PDF', text: 'テキスト', markdown: 'Markdown',
  csv: 'CSV', excel: 'Excel', word: 'Word', powerpoint: 'PowerPoint', html: 'Webページ',
};

const APPLE = 'Mac の Pages・Numbers・Keynote のファイルは、そのアプリの「ファイル → 書き出す → PDF」でPDFにできます';
const UNSUPPORTED = {
  doc: '古い Word 形式（.doc）には対応していません。Word で開いて「.docx」形式で保存し直してください',
  ppt: '古い PowerPoint 形式（.ppt）には対応していません。PowerPoint で開いて「.pptx」形式で保存し直してください',
  pages: APPLE, numbers: APPLE, key: APPLE,
  odt: 'この形式（.odt）には対応していません。「.docx」形式で保存し直してください',
  odp: 'この形式（.odp）には対応していません。「.pptx」形式で保存し直してください',
  rtf: 'この形式（.rtf）には対応していません。Word などで「.docx」形式で保存し直してください',
  zip: 'ZIP ファイルは、先に展開（解凍）してから中のファイルを選んでください',
};

function extOf(name) {
  const m = /\.([^./\\]+)$/.exec(name || '');
  return m ? m[1].toLowerCase() : '';
}

function detectKind(name, mime) {
  const ext = extOf(name);
  let kind = KINDS[ext];
  if (!kind && UNSUPPORTED[ext]) return { kind: null, reason: UNSUPPORTED[ext] };
  if (!kind) {
    // 拡張子が無い・分からないときは、ファイルの種類情報（MIME）で判断する
    if (/^image\/(jpeg|png|gif|webp|bmp)$/.test(mime)) kind = 'image';
    else if (/^image\/hei[cf]$/.test(mime)) kind = 'heic';
    else if (mime === 'application/pdf') kind = 'pdf';
    else if (mime === 'text/plain') kind = 'text';
  }
  if (!kind) return { kind: null, reason: 'この形式には対応していません' };
  return { kind, label: LABELS[kind] };
}

// ---------- 文字コード ----------

// 文字コードを自動で判定して文字列にする。
// UTF-8 として読めなければ Shift_JIS（Windows や Excel で作ったファイルに多い）とみなす。
function decodeText(bytes) {
  if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(bytes.subarray(3));
  if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('shift_jis').decode(bytes);
  }
}

// ---------- ページ割り（純粋な計算） ----------

// 長い文書をページに分ける位置を決める。
//   rects  : 途中で切ってはいけない範囲 [{top, bottom}]（文字の1行、画像、表の1行など）
//   pageH  : 1ページに入る高さ
//   total  : 文書全体の高さ
//   forced : 必ずここで改ページする位置（Excel のシートの区切りなど）
// 戻り値は [[開始, 終了], ...]
function computeBreaks(rects, pageH, total, forced = [], minFill = 0.3) {
  const pages = [];
  const forcedSorted = forced.filter(f => f > 0.5 && f < total - 0.5).sort((a, b) => a - b);
  let start = 0;
  while (start < total - 0.5) {
    const limit = start + pageH;
    let end;
    const f = forcedSorted.find(v => v > start + 0.5 && v <= limit);
    if (f !== undefined) end = f;
    else if (limit >= total) end = total;
    else end = safeBreak(rects, start + pageH * minFill, limit);
    pages.push([start, end]);
    start = end;
  }
  if (!pages.length) pages.push([0, Math.max(total, 0)]);   // 空のファイルでも1ページ作る
  return pages;
}

// lo〜hi の間で、どの行にもかからない一番下の位置を探す。無ければ hi で切る。
function safeBreak(rects, lo, hi) {
  const near = rects.filter(r => r.bottom > lo && r.top < hi);
  const cands = [hi];
  for (const r of near) {
    if (r.top > lo && r.top < hi) cands.push(r.top);
    if (r.bottom > lo && r.bottom < hi) cands.push(r.bottom);
  }
  cands.sort((a, b) => b - a);
  for (const y of cands) {
    if (!near.some(r => r.top < y - 0.5 && r.bottom > y + 0.5)) return y;
  }
  return hi;
}

// 画像1枚を A4 ページに収める配置（横長の画像は横向きページ）。
// 小さい画像は引き伸ばさない。
function fitImage(imgW, imgH) {
  const landscape = imgW > imgH;
  const pageW = landscape ? A4.h : A4.w;
  const pageH = landscape ? A4.w : A4.h;
  const s = Math.min((pageW - 2 * MARGIN_PT) / imgW, (pageH - 2 * MARGIN_PT) / imgH, 72 / 96);
  const w = imgW * s, h = imgH * s;
  return { pageW, pageH, x: (pageW - w) / 2, y: (pageH - h) / 2, w, h };
}

// ---------- 小さな道具 ----------

function esc(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

const loadedLibs = {};
// ライブラリは、その形式を変換するときに初めて読み込む（最初の表示を軽くするため）
function loadLib(file) {
  if (!loadedLibs[file]) {
    loadedLibs[file] = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'lib/' + file;
      s.onload = resolve;
      s.onerror = () => { delete loadedLibs[file]; reject(new Error('部品（' + file + '）を読み込めませんでした。ページを再読み込みしてください')); };
      document.head.appendChild(s);
    });
  }
  return loadedLibs[file];
}

async function readBytes(file) {
  return new Uint8Array(await file.arrayBuffer());
}

function makeCanvas(w, h) {
  const c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(w));
  c.height = Math.max(1, Math.round(h));
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, c.width, c.height);
  return { canvas: c, ctx };
}

function canvasBytes(canvas, type, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(async blob => {
      if (!blob) { reject(new Error('画像の作成に失敗しました（メモリが足りない可能性があります）')); return; }
      resolve(new Uint8Array(await blob.arrayBuffer()));
    }, type, quality);
  });
}

// ---------- 画像 ----------

async function imageToPage(blob, asPng) {
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.src = url;
    try { await img.decode(); } catch { throw new Error('画像を読み込めませんでした（壊れているか、このブラウザが対応していない形式です）'); }
    const k = Math.min(1, MAX_IMAGE_PX / Math.max(img.naturalWidth, img.naturalHeight));
    const { canvas, ctx } = makeCanvas(img.naturalWidth * k, img.naturalHeight * k);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);   // ここでスマホ写真の向き（回転情報）も反映される
    const bytes = asPng ? await canvasBytes(canvas, 'image/png') : await canvasBytes(canvas, 'image/jpeg', 0.92);
    canvas.width = canvas.height = 0;
    return { type: 'image', format: asPng ? 'png' : 'jpg', bytes, ...fitImage(img.naturalWidth, img.naturalHeight) };
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function heicToPage(file) {
  try {
    return await imageToPage(file, false);              // Safari はそのまま読める
  } catch {
    await loadLib('heic2any.min.js');
    let jpeg;
    try {
      const timeout = new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 120000));
      jpeg = await Promise.race([global.heic2any({ blob: file, toType: 'image/jpeg', quality: 0.92 }), timeout]);
    } catch {
      throw new Error('HEIC 画像を読み込めませんでした');
    }
    return imageToPage(Array.isArray(jpeg) ? jpeg[0] : jpeg, false);
  }
}

// ---------- PDF ----------

async function loadPdf(file) {
  await loadLib('pdf-lib.min.js');
  try {
    return { type: 'pdf', doc: await global.PDFLib.PDFDocument.load(await readBytes(file)) };
  } catch (e) {
    if (/encrypt/i.test(e && e.message)) throw new Error('パスワードや保護がかかったPDFは使えません');
    throw new Error('PDFを読み込めませんでした（壊れている可能性があります）');
  }
}

// ---------- 文書（HTML）をページ画像にする ----------

const FONT_STACK = '-apple-system, BlinkMacSystemFont, "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic UI", "Yu Gothic", Meiryo, "Noto Sans JP", sans-serif';
const SERIF_STACK = '"Hiragino Mincho ProN", "Yu Mincho", YuMincho, "MS Mincho", "Noto Serif JP", serif';
const MONO_STACK = 'ui-monospace, Menlo, Consolas, "Osaka-Mono", "MS Gothic", monospace';

// 変換用の枠の中では、外部への通信（画像・フォント・スクリプトの読み込み）をすべて禁止する
const FRAME_CSP = `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: blob:; style-src 'unsafe-inline'; font-src data:">`;

const DOC_CSS = `
html { background: #fff; color: #111; }
body { margin: 0; font-family: ${FONT_STACK}; font-size: 15px; line-height: 1.7; overflow-wrap: break-word; }
img { max-width: 100%; height: auto; }
h1, h2, h3, h4 { line-height: 1.4; margin: 1em 0 0.5em; }
p { margin: 0.6em 0; }
table { border-collapse: collapse; margin: 0.6em 0; }
td, th { border: 1px solid #999; padding: 3px 6px; vertical-align: top; }
th { background: #f0f0f0; }
pre, code { font-family: ${MONO_STACK}; font-size: 13px; }
pre { white-space: pre-wrap; background: #f4f4f4; padding: 8px 10px; border-radius: 4px; }
blockquote { margin: 0.6em 0; padding-left: 12px; border-left: 4px solid #ccc; color: #444; }
.plain .line { white-space: pre-wrap; tab-size: 4; min-height: 1.7em; }
.sheet-name { font-size: 16px; margin: 0 0 8px; }
.sheet td { white-space: nowrap; font-size: 13px; line-height: 1.4; }
.page-break { height: 0; }
.doc-title { font-size: 28px; text-align: center; margin-top: 0; }
.doc-subtitle { text-align: center; color: #555; }
`;

// 画面の外に見えない枠（iframe）を作り、その中に文書を表示して処理する
async function withFrame(srcdoc, width, fn) {
  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
  iframe.tabIndex = -1;
  iframe.style.cssText = `position:absolute;left:-30000px;top:0;width:${width}px;height:200px;border:0;`;
  const loaded = new Promise(r => { iframe.onload = r; });
  iframe.srcdoc = srcdoc;
  document.body.appendChild(iframe);
  try {
    await loaded;
    const doc = iframe.contentDocument;
    await Promise.all([...doc.images].map(img => img.decode().catch(() => {})));
    if (doc.fonts) await doc.fonts.ready;
    return await fn(doc, iframe);
  } finally {
    iframe.remove();
  }
}

function wrapBody(bodyHtml, css) {
  return `<!doctype html><html><head><meta charset="utf-8">${FRAME_CSP}<style>${css}</style></head><body>${bodyHtml}</body></html>`;
}

// 他人が作った Web ページは、スクリプトや外部読み込みを取り除いてから表示する
function sanitizeHtml(html) {
  const d = new DOMParser().parseFromString(html, 'text/html');
  d.querySelectorAll('script, iframe, frame, frameset, object, embed, link, meta, base, noscript').forEach(e => e.remove());
  // 外部の画像などは読みに行かない（data: / blob: 以外の参照は外す）
  const external = v => !/^\s*(data|blob):/i.test(v);
  for (const el of d.querySelectorAll('*')) {
    for (const a of [...el.attributes]) {
      const n = a.name.toLowerCase();
      if (/^on/.test(n) || /^\s*javascript:/i.test(a.value)) el.removeAttribute(a.name);
      else if ((n === 'src' || n === 'poster' || n === 'background' || (n === 'href' && el.localName !== 'a')) && external(a.value)) el.removeAttribute(a.name);
      else if (n === 'srcset') el.removeAttribute(a.name);
      else if (n === 'style') el.setAttribute('style', stripCssUrls(a.value));
    }
  }
  d.querySelectorAll('style').forEach(st => { st.textContent = stripCssUrls(st.textContent).replace(/@import[^;]*;/gi, ''); });
  const style = d.createElement('style');
  style.textContent = 'html { background: #fff; }';
  d.head.prepend(style);
  d.head.insertAdjacentHTML('afterbegin', '<meta charset="utf-8">' + FRAME_CSP);
  return '<!doctype html>' + d.documentElement.outerHTML;
}

function stripCssUrls(css) {
  return css.replace(/url\(\s*(?!['"]?\s*(data|blob):)[^)]*\)/gi, 'none');
}

// 途中で切ってはいけない範囲（文字の各行・画像・表の行）と、強制改ページの位置を測る
function measureBlocks(doc) {
  const win = doc.defaultView;
  const sy = win.scrollY;
  const rects = [];
  const range = doc.createRange();
  const walker = doc.createTreeWalker(doc.body, win.NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!n.nodeValue.trim()) continue;
    range.selectNodeContents(n);
    for (const r of range.getClientRects()) {
      if (r.height > 0) rects.push({ top: r.top + sy, bottom: r.bottom + sy });
    }
  }
  for (const el of doc.body.querySelectorAll('img, svg, canvas, video, hr, tr')) {
    const r = el.getBoundingClientRect();
    if (r.height <= 0) continue;
    if (el.tagName === 'TR' && r.height > 400) continue;   // とても高い表の行は、行の中で切ってよい
    rects.push({ top: r.top + sy, bottom: r.bottom + sy });
  }
  const forced = [...doc.querySelectorAll('.page-break')].map(el => el.getBoundingClientRect().top + sy);
  // 横方向に分けるときに、列の途中で切らないための範囲（表の最初の30行のセル）
  const cols = [];
  for (const tr of doc.querySelectorAll('tr')) {
    if (tr.rowIndex >= 30) continue;
    for (const cell of tr.cells) {
      const r = cell.getBoundingClientRect();
      cols.push({ top: r.left + win.scrollX, bottom: r.right + win.scrollX });
    }
  }
  return { rects, forced, cols };
}

// HTML を A4 のページ画像にする。横に広い文書（大きな表など）は横向き・縮小で収める。
async function renderHtmlPages(html, { fullDocument = false, onPage } = {}) {
  await loadLib('html2canvas.min.js');
  const marginPx = MARGIN_PT * PX_PER_PT;
  const portraitW = Math.floor(A4.w * PX_PER_PT - 2 * marginPx);
  const landscapeW = Math.floor(A4.h * PX_PER_PT - 2 * marginPx);
  const srcdoc = fullDocument ? sanitizeHtml(html) : wrapBody(html, DOC_CSS);

  return withFrame(srcdoc, portraitW, async (doc, iframe) => {
    const root = doc.documentElement;
    let landscape = false;
    if (root.scrollWidth > portraitW + 1) {
      landscape = true;
      iframe.style.width = Math.max(landscapeW, root.scrollWidth) + 'px';
    }
    const W = Math.max(root.scrollWidth, landscape ? landscapeW : portraitW);
    const pageWpx = (landscape ? A4.h : A4.w) * PX_PER_PT;
    const pageHpx = (landscape ? A4.w : A4.h) * PX_PER_PT;
    const areaW = pageWpx - 2 * marginPx;
    const areaH = pageHpx - 2 * marginPx;
    const f = Math.max(MIN_SCALE, Math.min(1, areaW / W));    // 縮小率
    iframe.style.height = root.scrollHeight + 'px';
    const total = root.scrollHeight;

    const { rects, forced, cols } = measureBlocks(doc);
    const rows = computeBreaks(rects, areaH / f, total, forced);
    const columns = W * f > areaW + 1 ? computeBreaks(cols, areaW / f, W) : [[0, W]];
    const count = rows.length * columns.length;
    if (count > MAX_PAGES) throw new Error(`長すぎます（${count}ページ）。${MAX_PAGES}ページ以内になるよう分けてください`);

    const pages = [];
    // 横にも分けるときは、左→右、上→下の順に並べる
    for (const [s, e] of rows) {
      for (const [xs, xe] of columns) {
        const { canvas, ctx } = makeCanvas(pageWpx * RENDER_SCALE, pageHpx * RENDER_SCALE);
        const h = Math.ceil(e - s), w = Math.ceil(xe - xs);
        if (h >= 1 && w >= 1) {
          const shot = await global.html2canvas(root, {
            x: xs, y: s, width: w, height: h, windowWidth: W, windowHeight: total,
            scale: RENDER_SCALE * f, backgroundColor: '#ffffff', logging: false,
          });
          ctx.drawImage(shot, marginPx * RENDER_SCALE, marginPx * RENDER_SCALE, w * f * RENDER_SCALE, h * f * RENDER_SCALE);
          shot.width = shot.height = 0;
        }
        const bytes = await canvasBytes(canvas, 'image/jpeg', 0.85);
        canvas.width = canvas.height = 0;
        const pageW = pageWpx / PX_PER_PT, pageH = pageHpx / PX_PER_PT;
        pages.push({ type: 'image', format: 'jpg', bytes, pageW, pageH, x: 0, y: 0, w: pageW, h: pageH });
        if (onPage) onPage(pages.length, count);
      }
    }
    return pages;
  });
}

// ---------- テキスト・Markdown・表 ----------

function textToHtml(text) {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  if (lines.length && lines[lines.length - 1] === '') lines.pop();
  return '<div class="plain">' + lines.map(l => `<div class="line">${esc(l)}</div>`).join('') + '</div>';
}

// Excel・CSV をシートごとの表の HTML にする
async function sheetsHtml(bytes, csvText, separator) {
  await loadLib('xlsx.full.min.js');
  const XLSX = global.XLSX;
  let wb;
  try {
    wb = csvText != null
      ? XLSX.read(csvText, { type: 'string', raw: true, FS: separator })   // raw: 先頭の0（電話番号など）を消さない
      : XLSX.read(bytes, { type: 'array', cellDates: true });
  } catch (e) {
    if (/password|encrypt/i.test(e && e.message)) throw new Error('パスワード付きのファイルは使えません');
    throw new Error('ファイルを読み込めませんでした（壊れている可能性があります）');
  }
  const meta = (wb.Workbook && wb.Workbook.Sheets) || [];
  const parts = [];
  wb.SheetNames.forEach((name, i) => {
    if (meta[i] && meta[i].Hidden) return;                  // 非表示のシートは出さない
    const ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) return;                         // 空のシート
    const table = XLSX.utils.sheet_to_html(ws, { header: '', footer: '' });
    parts.push((wb.SheetNames.length > 1 ? `<h1 class="sheet-name">${esc(name)}</h1>` : '') + `<div class="sheet">${table}</div>`);
  });
  return parts.length ? parts : ['<p>（空のファイルです）</p>'];
}

// シートごとに用紙の向き・縮小率を決めて、順にページにする
async function renderSheets(parts, onPage) {
  const pages = [];
  for (const html of parts) {
    const done = pages.length;
    pages.push(...await renderHtmlPages(html, { onPage: onPage && ((p, n) => onPage(done + p, done + n)) }));
    if (pages.length > MAX_PAGES) throw new Error(`ページが多すぎます（${MAX_PAGES}ページまで）`);
  }
  return pages;
}

// ---------- PowerPoint ----------

const IMG_MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };

const kids = (el, name) => (el ? [...el.children].filter(c => c.localName === name) : []);
const kid = (el, name) => kids(el, name)[0] || null;
const path = (el, ...names) => names.reduce((e, n) => kid(e, n), el);
const desc = (el, name) => (el ? el.getElementsByTagNameNS('*', name)[0] || null : null);
const descAll = (el, name) => (el ? [...el.getElementsByTagNameNS('*', name)] : []);
const num = (el, attr) => { const v = el && el.getAttribute(attr); return v == null || v === '' ? null : Number(v); };
const pick = (list, fn) => { for (const e of list) { const v = fn(e); if (v != null) return v; } return null; };

// "ppt/slides/" と "../media/image1.png" から "ppt/media/image1.png" を作る
function resolvePath(dir, target) {
  const parts = (target.startsWith('/') ? target.slice(1) : dir + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

// 色の明るさ調整（PowerPoint のテーマ色によく付いている lumMod / lumOff）
function applyLum(hex, lumMod = 100000, lumOff = 0) {
  let r = parseInt(hex.slice(1, 3), 16) / 255, g = parseInt(hex.slice(3, 5), 16) / 255, b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0, l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = Math.min(1, Math.max(0, l * lumMod / 100000 + lumOff / 100000));
  const hue = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  if (s === 0) { r = g = b = l; } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    r = hue(p, q, h + 1 / 3); g = hue(p, q, h); b = hue(p, q, h - 1 / 3);
  }
  return '#' + [r, g, b].map(v => Math.round(v * 255).toString(16).padStart(2, '0')).join('');
}

function readTheme(themeDoc) {
  const map = {};
  const cs = desc(themeDoc, 'clrScheme');
  if (cs) {
    for (const c of cs.children) {
      const s = desc(c, 'srgbClr'), sys = desc(c, 'sysClr');
      const v = s ? s.getAttribute('val') : sys ? sys.getAttribute('lastClr') : null;
      if (v) map[c.localName] = v;
    }
  }
  map.bg1 = map.lt1 || 'FFFFFF'; map.tx1 = map.dk1 || '000000';
  map.bg2 = map.lt2 || map.bg1; map.tx2 = map.dk2 || map.tx1;
  return map;
}

// 色の指定（srgbClr / schemeClr など）を CSS の色にする
function colorOf(holder, theme) {
  if (!holder) return null;
  const c = kid(holder, 'srgbClr') || kid(holder, 'schemeClr') || kid(holder, 'sysClr') || kid(holder, 'prstClr');
  if (!c) return null;
  let hex;
  if (c.localName === 'srgbClr') hex = c.getAttribute('val');
  else if (c.localName === 'schemeClr') hex = theme[c.getAttribute('val')];
  else if (c.localName === 'sysClr') hex = c.getAttribute('lastClr') || '000000';
  else return c.getAttribute('val');
  if (!hex) return null;
  let col = applyLum('#' + hex, num(kid(c, 'lumMod'), 'val') ?? 100000, num(kid(c, 'lumOff'), 'val') ?? 0);
  const alpha = num(kid(c, 'alpha'), 'val');
  if (alpha != null && alpha < 100000) {
    col = `rgba(${parseInt(col.slice(1, 3), 16)},${parseInt(col.slice(3, 5), 16)},${parseInt(col.slice(5, 7), 16)},${(alpha / 100000).toFixed(3)})`;
  }
  return col;
}

// 塗りつぶし: 色 / null（塗りなし）/ undefined（指定なし）
function fillOf(el, theme) {
  if (!el) return undefined;
  if (kid(el, 'noFill')) return null;
  const solid = kid(el, 'solidFill');
  if (solid) return colorOf(solid, theme);
  const grad = kid(el, 'gradFill');
  if (grad) return colorOf(desc(grad, 'gs'), theme);       // グラデーションは最初の色で代用
  return undefined;
}

function readXfrm(x) {
  if (!x) return null;
  const off = kid(x, 'off'), ext = kid(x, 'ext');
  if (!off || !ext) return null;
  return {
    x: num(off, 'x') || 0, y: num(off, 'y') || 0, cx: num(ext, 'cx') || 0, cy: num(ext, 'cy') || 0,
    rot: (num(x, 'rot') || 0) / 60000, flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1',
  };
}

function phOf(node) {
  const nv = kid(node, 'nvSpPr') || kid(node, 'nvPicPr') || kid(node, 'nvGraphicFramePr') || kid(node, 'nvCxnSpPr');
  const ph = desc(nv, 'ph');
  return ph ? { type: ph.getAttribute('type'), idx: ph.getAttribute('idx') } : null;
}

const phGroup = t => (t === 'title' || t === 'ctrTitle' ? 'title' : t === 'dt' || t === 'ftr' || t === 'sldNum' ? t : 'body');

function findPh(doc, ph, type) {
  if (!doc) return null;
  const all = descAll(doc, 'sp').filter(s => phOf(s));
  return (ph.idx != null && all.find(s => phOf(s).idx === ph.idx))
    || all.find(s => (phOf(s).type || 'body') === type)
    || all.find(s => phGroup(phOf(s).type || 'body') === phGroup(type))
    || null;
}

function autoNumber(type, n) {
  const alpha = n => { let s = ''; while (n > 0) { s = String.fromCharCode(97 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); } return s; };
  const roman = n => [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]
    .reduce((s, [v, r]) => { while (n >= v) { s += r; n -= v; } return s; }, '');
  const t = type || 'arabicPeriod';
  let body = t.startsWith('alphaLc') ? alpha(n) : t.startsWith('alphaUc') ? alpha(n).toUpperCase()
    : t.startsWith('romanLc') ? roman(n) : t.startsWith('romanUc') ? roman(n).toUpperCase()
    : t.startsWith('circleNum') ? String.fromCharCode(0x2460 + n - 1) : String(n);
  if (t.endsWith('ParenBoth')) return `(${body})`;
  if (t.endsWith('ParenR')) return body + ')';
  if (t.endsWith('Period')) return body + '.';
  return body;
}

const ptToPx = pt => pt * PX_PER_PT;

// テキスト（段落の並び）を HTML にする。styles は書式を探しに行く順番（自分 → レイアウト → マスター）
function renderParagraphs(txBody, styles, ctx, opt) {
  const theme = ctx.theme;
  const scale = opt.fontScale || 1;
  const counters = [];
  let html = '';
  for (const p of kids(txBody, 'p')) {
    const pPr = kid(p, 'pPr');
    const lvl = num(pPr, 'lvl') || 0;
    const chain = [pPr, ...styles.map(s => kid(s, `lvl${lvl + 1}pPr`))].filter(Boolean);
    const defs = chain.map(e => kid(e, 'defRPr')).filter(Boolean);
    const runs = [...p.children].filter(c => c.localName === 'r' || c.localName === 'br' || c.localName === 'fld');
    const firstRPr = runs.map(r => kid(r, 'rPr')).find(Boolean);
    const endRPr = kid(p, 'endParaRPr');

    const prop = (rPr, attr) => pick([rPr, ...defs].filter(Boolean), e => e.getAttribute(attr));
    const runColor = rPr => pick([rPr, ...defs].filter(Boolean), e => colorOf(kid(e, 'solidFill'), theme)) || opt.color || '#' + theme.tx1;
    const runSize = rPr => { const sz = prop(rPr, 'sz'); return ptToPx((sz ? sz / 100 : opt.defaultPt) * scale); };

    const algn = pick(chain, e => e.getAttribute('algn'));
    const marL = pick(chain, e => num(e, 'marL')) || 0;
    const indent = pick(chain, e => num(e, 'indent')) || 0;
    const lnPct = pick(chain, e => num(path(e, 'lnSpc', 'spcPct'), 'val'));
    const lnPts = pick(chain, e => num(path(e, 'lnSpc', 'spcPts'), 'val'));
    const bef = pick(chain, e => num(path(e, 'spcBef', 'spcPts'), 'val'));
    const bullet = pick(chain, e => kid(e, 'buNone') ? { none: true }
      : kid(e, 'buChar') ? { char: kid(e, 'buChar').getAttribute('char') }
      : kid(e, 'buAutoNum') ? { auto: kid(e, 'buAutoNum').getAttribute('type'), start: num(kid(e, 'buAutoNum'), 'startAt') || 1 }
      : kid(e, 'buBlip') ? { char: '•' } : null);

    const lineH = lnPts != null ? `${ptToPx(lnPts / 100)}px` : String(1.2 * ((lnPct ?? 100000) - (opt.lnReduce || 0)) / 100000);
    const style = [
      `text-align:${{ ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[algn] || 'left'}`,
      `padding-left:${marL / EMU_PER_PX}px`, `text-indent:${indent / EMU_PER_PX}px`, `line-height:${lineH}`,
      bef ? `margin-top:${ptToPx(bef / 100)}px` : '',
    ].filter(Boolean).join(';');

    const span = (text, rPr) => {
      const css = [`font-size:${runSize(rPr).toFixed(2)}px`, `color:${runColor(rPr)}`];
      if (prop(rPr, 'b') === '1') css.push('font-weight:bold');
      if (prop(rPr, 'i') === '1') css.push('font-style:italic');
      const u = prop(rPr, 'u');
      if (u && u !== 'none') css.push('text-decoration:underline');
      const strike = prop(rPr, 'strike');
      if (strike && strike !== 'noStrike') css.push('text-decoration:line-through');
      const base = num(rPr, 'baseline');
      if (base) css.push(`vertical-align:${base > 0 ? 'super' : 'sub'}`, `font-size:${(runSize(rPr) * 0.7).toFixed(2)}px`);
      const fontEl = rPr && (kid(rPr, 'ea') || kid(rPr, 'latin'));
      const face = fontEl && fontEl.getAttribute('typeface');
      if (face && /mincho|明朝|serif|times|georgia/i.test(face)) css.push(`font-family:${SERIF_STACK.replace(/"/g, "'")}`);
      return `<span style="${css.join(';')}">${esc(text).replace(/\t/g, '    ')}</span>`;
    };

    let inner = '';
    for (const r of runs) {
      if (r.localName === 'br') { inner += '<br>'; continue; }
      let text = (kid(r, 't') || {}).textContent || '';
      if (r.localName === 'fld' && r.getAttribute('type') === 'slidenum') text = String(ctx.slideNo);
      inner += span(text, kid(r, 'rPr'));
    }
    const hasText = runs.some(r => r.localName !== 'br' && ((kid(r, 't') || {}).textContent || '').length);

    if (hasText && bullet && !bullet.none && opt.bullets !== false) {
      let mark = bullet.char;
      if (bullet.auto) {
        counters.length = lvl + 1;
        counters[lvl] = (counters[lvl] || bullet.start - 1) + 1;
        mark = autoNumber(bullet.auto, counters[lvl]);
      }
      if (mark) inner = span(mark + ' ', firstRPr) + inner;
    } else if (!bullet || !bullet.auto) {
      counters.length = lvl;
    }
    if (!hasText) inner = span(' ', endRPr || firstRPr);
    html += `<div style="${style}">${inner}</div>`;
  }
  return html;
}

// スライド1枚分の図形を、位置どおりに並べた HTML にする
async function renderShapes(tree, map, ctx, rels, skipPh) {
  let out = '';
  const nodes = [];
  for (const n of tree.children) {
    if (n.localName === 'AlternateContent') {
      const alt = kid(n, 'Fallback') || kid(n, 'Choice');
      if (alt) nodes.push(...alt.children);
    } else nodes.push(n);
  }
  for (const node of nodes) {
    const name = node.localName;
    if (name === 'grpSp') {
      const x = path(node, 'grpSpPr', 'xfrm');
      const xf = readXfrm(x);
      if (!xf) { out += await renderShapes(node, map, ctx, rels, skipPh); continue; }
      const chOff = kid(x, 'chOff'), chExt = kid(x, 'chExt');
      const cx0 = num(chOff, 'x') ?? xf.x, cy0 = num(chOff, 'y') ?? xf.y;
      const sx = num(chExt, 'cx') ? xf.cx / num(chExt, 'cx') : 1, sy = num(chExt, 'cy') ? xf.cy / num(chExt, 'cy') : 1;
      const inner = b => map({ ...b, x: xf.x + (b.x - cx0) * sx, y: xf.y + (b.y - cy0) * sy, cx: b.cx * sx, cy: b.cy * sy });
      out += await renderShapes(node, inner, ctx, rels, skipPh);
    } else if (name === 'sp' || name === 'cxnSp' || name === 'pic' || name === 'graphicFrame') {
      const ph = phOf(node);
      if (ph && skipPh) continue;
      try {
        out += await renderShape(node, map, ctx, rels, ph);
      } catch (e) {
        console.warn('図形を描けませんでした', e);
      }
    }
  }
  return out;
}

async function renderShape(node, map, ctx, rels, ph) {
  const theme = ctx.theme;
  const name = node.localName;
  const layoutPh = ph ? findPh(ctx.layout, ph, ph.type || 'body') : null;
  const type = ph ? (ph.type || (layoutPh && phOf(layoutPh).type) || 'body') : null;
  const masterPh = ph ? findPh(ctx.master, { type }, type) : null;

  const spPr = kid(node, 'spPr');
  const xfEl = name === 'graphicFrame' ? kid(node, 'xfrm') : kid(spPr, 'xfrm');
  const xf = readXfrm(xfEl) || readXfrm(path(layoutPh, 'spPr', 'xfrm')) || readXfrm(path(masterPh, 'spPr', 'xfrm'));
  if (!xf) return '';
  const b = map(xf);
  const box = `position:absolute;left:${b.x.toFixed(2)}px;top:${b.y.toFixed(2)}px;width:${b.cx.toFixed(2)}px;height:${b.cy.toFixed(2)}px;`
    + (xf.rot ? `transform:rotate(${xf.rot}deg);` : '');

  // 画像
  if (name === 'pic') {
    const blip = desc(node, 'blip');
    const r = blip && rels[blip.getAttribute('r:embed')];
    const src = r && await ctx.media(r.path);
    if (!src) return `<div style="${box}${UNSHOWN}">画像（表示できない形式）</div>`;
    const crop = desc(node, 'srcRect');
    const l = (num(crop, 'l') || 0) / 100000, t = (num(crop, 't') || 0) / 100000;
    const rr = (num(crop, 'r') || 0) / 100000, bb = (num(crop, 'b') || 0) / 100000;
    const iw = b.cx / Math.max(0.01, 1 - l - rr), ih = b.cy / Math.max(0.01, 1 - t - bb);
    const flip = (xf.flipH || xf.flipV) ? `transform:scale(${xf.flipH ? -1 : 1},${xf.flipV ? -1 : 1});` : '';
    return `<div style="${box}overflow:hidden;"><img src="${src}" style="position:absolute;left:${(-l * iw).toFixed(2)}px;top:${(-t * ih).toFixed(2)}px;width:${iw.toFixed(2)}px;height:${ih.toFixed(2)}px;max-width:none;${flip}"></div>`;
  }

  // 表・グラフなど
  if (name === 'graphicFrame') {
    const gd = desc(node, 'graphicData');
    const tbl = desc(gd, 'tbl');
    if (tbl) return renderTable(tbl, b, ctx);
    return `<div style="${box}${UNSHOWN}">グラフ・図（SmartArt など）は表示できません</div>`;
  }

  // 図形・テキストボックス
  const style = kid(node, 'style');
  let fill = fillOf(spPr, theme);
  if (fill === undefined) fill = style && num(kid(style, 'fillRef'), 'idx') ? colorOf(kid(style, 'fillRef'), theme) : null;
  const ln = kid(spPr, 'ln');
  let lineColor = ln ? fillOf(ln, theme) : undefined;
  if (lineColor === undefined) lineColor = style && num(kid(style, 'lnRef'), 'idx') ? colorOf(kid(style, 'lnRef'), theme) : null;
  const lineW = Math.max(1, (num(ln, 'w') || 12700) / EMU_PER_PX);
  const prst = kid(spPr, 'prstGeom')?.getAttribute('prst') || '';

  if (name === 'cxnSp' || prst === 'line' || /Connector/.test(prst)) {
    if (!lineColor) return '';
    // 線は「細長い四角を回転させたもの」として描く（html2canvas は SVG の線をうまく描けないため）
    const len = Math.hypot(b.cx, b.cy);
    const angle = Math.atan2(xf.flipV ? -b.cy : b.cy, xf.flipH ? -b.cx : b.cx) * 180 / Math.PI + xf.rot;
    const cx = b.x + b.cx / 2, cy = b.y + b.cy / 2;
    return `<div style="position:absolute;left:${(cx - len / 2).toFixed(2)}px;top:${(cy - lineW / 2).toFixed(2)}px;width:${len.toFixed(2)}px;height:${lineW.toFixed(2)}px;`
      + `background:${lineColor};transform:rotate(${angle.toFixed(3)}deg);"></div>`;
  }

  let deco = '';
  if (fill) deco += `background:${fill};`;
  if (lineColor) deco += `border:${lineW.toFixed(2)}px solid ${lineColor};box-sizing:border-box;`;
  if (prst === 'ellipse') deco += 'border-radius:50%;';
  else if (prst === 'roundRect') deco += `border-radius:${(Math.min(b.cx, b.cy) * 0.1667).toFixed(2)}px;`;

  const txBody = kid(node, 'txBody');
  let text = '';
  if (txBody && descAll(txBody, 't').some(t => t.textContent.trim())) {
    const bodyPrs = [kid(txBody, 'bodyPr'), path(layoutPh, 'txBody', 'bodyPr'), path(masterPh, 'txBody', 'bodyPr')].filter(Boolean);
    const anchor = pick(bodyPrs, e => e.getAttribute('anchor')) || 't';
    const ins = k => (pick(bodyPrs, e => num(e, k)) ?? (k === 'lIns' || k === 'rIns' ? 91440 : 45720)) / EMU_PER_PX;
    const wrapNone = pick(bodyPrs, e => e.getAttribute('wrap')) === 'none';
    const fit = desc(kid(txBody, 'bodyPr'), 'normAutofit');
    const txStyles = desc(ctx.master, 'txStyles');
    const masterStyle = !ph ? null : phGroup(type) === 'title' ? kid(txStyles, 'titleStyle') : phGroup(type) === 'body' ? kid(txStyles, 'bodyStyle') : kid(txStyles, 'otherStyle');
    const styles = [kid(txBody, 'lstStyle'), path(layoutPh, 'txBody', 'lstStyle'), path(masterPh, 'txBody', 'lstStyle'),
      masterStyle, !ph ? ctx.defaultTextStyle : null, !ph ? kid(txStyles, 'otherStyle') : null].filter(Boolean);
    const fontRefColor = style ? colorOf(kid(style, 'fontRef'), theme) : null;
    const body = renderParagraphs(txBody, styles, ctx, {
      defaultPt: phGroup(type) === 'title' ? 40 : ph ? 24 : 18,
      color: fontRefColor,
      fontScale: (num(fit, 'fontScale') ?? 100000) / 100000,
      lnReduce: num(fit, 'lnSpcReduction') || 0,
    });
    text = `<div style="position:absolute;inset:0;display:flex;flex-direction:column;justify-content:${{ ctr: 'center', b: 'flex-end' }[anchor] || 'flex-start'};`
      + `padding:${ins('tIns')}px ${ins('rIns')}px ${ins('bIns')}px ${ins('lIns')}px;box-sizing:border-box;`
      + `${wrapNone ? 'white-space:nowrap;' : ''}"><div>${body}</div></div>`;
  }
  if (!deco && !text) return '';
  return `<div style="${box}${deco}">${text}</div>`;
}

const UNSHOWN = 'display:flex;align-items:center;justify-content:center;background:#f2f2f2;border:1px dashed #aaa;color:#888;font-size:14px;box-sizing:border-box;';

function renderTable(tbl, b, ctx) {
  const cols = descAll(kid(tbl, 'tblGrid'), 'gridCol').map(c => (num(c, 'w') || 0) / EMU_PER_PX);
  const otherStyle = kid(desc(ctx.master, 'txStyles'), 'otherStyle');
  let rows = '';
  for (const tr of kids(tbl, 'tr')) {
    let cells = '';
    for (const tc of kids(tr, 'tc')) {
      if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') continue;
      const tcPr = kid(tc, 'tcPr');
      const bg = fillOf(tcPr, ctx.theme);
      const span = (num(tc, 'gridSpan') > 1 ? ` colspan="${num(tc, 'gridSpan')}"` : '') + (num(tc, 'rowSpan') > 1 ? ` rowspan="${num(tc, 'rowSpan')}"` : '');
      const txt = renderParagraphs(kid(tc, 'txBody'), [kid(kid(tc, 'txBody'), 'lstStyle'), otherStyle].filter(Boolean), ctx, { defaultPt: 18, bullets: false });
      cells += `<td${span} style="border:1px solid #999;padding:4.8px 9.6px;vertical-align:top;${bg ? `background:${bg};` : ''}">${txt}</td>`;
    }
    rows += `<tr style="height:${((num(tr, 'h') || 0) / EMU_PER_PX).toFixed(2)}px">${cells}</tr>`;
  }
  return `<table style="position:absolute;left:${b.x}px;top:${b.y}px;width:${b.cx}px;border-collapse:collapse;table-layout:fixed;">`
    + `<colgroup>${cols.map(w => `<col style="width:${w}px">`).join('')}</colgroup>${rows}</table>`;
}

async function slideBackground(parts, ctx) {
  for (const [doc, rels] of parts) {
    const bg = doc && desc(doc, 'bg');
    if (!bg) continue;
    const bgPr = kid(bg, 'bgPr');
    if (bgPr) {
      const blip = desc(bgPr, 'blip');
      const r = blip && rels[blip.getAttribute('r:embed')];
      const src = r && await ctx.media(r.path);
      if (src) return { color: '#fff', image: src };
      return { color: fillOf(bgPr, ctx.theme) || '#fff' };
    }
    const c = colorOf(kid(bg, 'bgRef'), ctx.theme);
    if (c) return { color: c };
  }
  return { color: '#fff' };
}

async function pptxPages(buffer, onPage) {
  await Promise.all([loadLib('jszip.min.js'), loadLib('html2canvas.min.js')]);
  let zip;
  try { zip = await global.JSZip.loadAsync(buffer); } catch {
    throw new Error('PowerPoint ファイルを開けませんでした（壊れているか、パスワード付きの可能性があります）');
  }
  const cache = {};
  const xml = async p => {
    if (!p) return null;
    if (!(p in cache)) {
      const f = zip.file(p);
      cache[p] = f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null;
    }
    return cache[p];
  };
  const media = async p => {
    const f = zip.file(p), mime = IMG_MIME[extOf(p)];
    return f && mime ? `data:${mime};base64,${await f.async('base64')}` : null;
  };
  const readRels = async part => {
    const dir = part.slice(0, part.lastIndexOf('/') + 1);
    const doc = await xml(dir + '_rels/' + part.slice(dir.length) + '.rels');
    const map = {};
    if (!doc) return map;
    for (const r of doc.documentElement.children) {
      if (r.getAttribute('TargetMode') === 'External') continue;
      map[r.getAttribute('Id')] = { type: r.getAttribute('Type') || '', path: resolvePath(dir, r.getAttribute('Target') || '') };
    }
    return map;
  };
  const relOf = (rels, kind) => Object.values(rels).find(r => r.type.endsWith('/' + kind));

  const pres = await xml('ppt/presentation.xml');
  const size = desc(pres, 'sldSz');
  if (!size) throw new Error('PowerPoint ファイルを読み込めませんでした');
  const W = num(size, 'cx') / EMU_PER_PX, H = num(size, 'cy') / EMU_PER_PX;
  const presRels = await readRels('ppt/presentation.xml');
  const toPx = b => ({ ...b, x: b.x / EMU_PER_PX, y: b.y / EMU_PER_PX, cx: b.cx / EMU_PER_PX, cy: b.cy / EMU_PER_PX });

  const slidesHtml = [];
  let slideNo = 0;
  for (const id of descAll(desc(pres, 'sldIdLst'), 'sldId')) {
    const rel = presRels[id.getAttribute('r:id')];
    const doc = rel && await xml(rel.path);
    if (!doc) continue;
    slideNo++;
    if (doc.documentElement.getAttribute('show') === '0') continue;   // 非表示スライドは出さない
    const rels = await readRels(rel.path);
    const layoutPath = (relOf(rels, 'slideLayout') || {}).path;
    const layout = await xml(layoutPath);
    const layoutRels = layoutPath ? await readRels(layoutPath) : {};
    const masterPath = (relOf(layoutRels, 'slideMaster') || {}).path;
    const master = await xml(masterPath);
    const masterRels = masterPath ? await readRels(masterPath) : {};
    const theme = readTheme(await xml((relOf(masterRels, 'theme') || {}).path));
    const ctx = { theme, layout, master, media, slideNo, defaultTextStyle: desc(pres, 'defaultTextStyle') };

    const bg = await slideBackground([[doc, rels], [layout, layoutRels], [master, masterRels]], ctx);
    let inner = bg.image ? `<img src="${bg.image}" style="position:absolute;left:0;top:0;width:${W}px;height:${H}px;max-width:none">` : '';
    // マスターとレイアウトの飾り（ロゴなど）→ スライド本体の順に重ねる
    const showMaster = doc.documentElement.getAttribute('showMasterSp') !== '0';
    if (showMaster && master && layout && layout.documentElement.getAttribute('showMasterSp') !== '0') {
      inner += await renderShapes(desc(master, 'spTree'), toPx, ctx, masterRels, true);
    }
    if (showMaster && layout) inner += await renderShapes(desc(layout, 'spTree'), toPx, ctx, layoutRels, true);
    inner += await renderShapes(desc(doc, 'spTree'), toPx, ctx, rels, false);
    slidesHtml.push(`<div class="slide" style="background:${bg.color}">${inner}</div>`);
  }
  if (!slidesHtml.length) throw new Error('表示できるスライドがありません');
  if (slidesHtml.length > MAX_PAGES) throw new Error(`スライドが多すぎます（${MAX_PAGES}枚まで）`);

  const css = `html,body{margin:0;background:#fff}body{font-family:${FONT_STACK};}
.slide{position:relative;width:${W}px;height:${H}px;overflow:hidden;white-space:pre-wrap;overflow-wrap:break-word;color:#000}
.slide div{white-space:inherit}`;
  return withFrame(wrapBody(slidesHtml.join(''), css), W, async doc => {
    const pages = [];
    const els = doc.querySelectorAll('.slide');
    const pageW = W / PX_PER_PT, pageH = H / PX_PER_PT;
    for (let i = 0; i < els.length; i++) {
      const shot = await global.html2canvas(els[i], { scale: RENDER_SCALE, backgroundColor: '#ffffff', logging: false, width: W, height: H, windowWidth: W });
      const bytes = await canvasBytes(shot, 'image/jpeg', 0.88);
      shot.width = shot.height = 0;
      pages.push({ type: 'image', format: 'jpg', bytes, pageW, pageH, x: 0, y: 0, w: pageW, h: pageH });
      if (onPage) onPage(i + 1, els.length);
    }
    return pages;
  });
}

// ---------- 入口 ----------

async function convert(file, kind, onPage) {
  const ext = extOf(file.name);
  switch (kind) {
    case 'image':
      return [await imageToPage(file, ext === 'png' || file.type === 'image/png')];
    case 'heic':
      return [await heicToPage(file)];
    case 'pdf':
      return [await loadPdf(file)];
    case 'text':
      return renderHtmlPages(textToHtml(decodeText(await readBytes(file))), { onPage });
    case 'markdown': {
      await loadLib('marked.min.js');
      const html = global.marked.parse(decodeText(await readBytes(file)), { gfm: true, breaks: true });
      return renderHtmlPages(html, { onPage });
    }
    case 'csv':
      return renderSheets(await sheetsHtml(null, decodeText(await readBytes(file)), ext === 'tsv' ? '\t' : undefined), onPage);
    case 'excel':
      return renderSheets(await sheetsHtml(await readBytes(file)), onPage);
    case 'word': {
      await loadLib('mammoth.browser.min.js');
      let result;
      try {
        result = await global.mammoth.convertToHtml({ arrayBuffer: await file.arrayBuffer() }, {
          styleMap: ["p[style-name='Title'] => h1.doc-title:fresh", "p[style-name='Subtitle'] => h2.doc-subtitle:fresh"],
        });
      } catch {
        throw new Error('Word ファイルを開けませんでした（壊れているか、パスワード付きの可能性があります）');
      }
      return renderHtmlPages(result.value || '<p></p>', { onPage });
    }
    case 'powerpoint':
      return pptxPages(await file.arrayBuffer(), onPage);
    case 'html':
      return renderHtmlPages(decodeText(await readBytes(file)), { fullDocument: true, onPage });
    default:
      throw new Error('この形式には対応していません');
  }
}

global.Conv = {
  detectKind, extOf, decodeText, computeBreaks, fitImage, resolvePath, applyLum, autoNumber, textToHtml, esc,
  loadLib, convert, A4,
};
})(typeof window !== 'undefined' ? window : globalThis);
