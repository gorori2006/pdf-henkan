// 画面の操作: ファイルの追加・並べ替え・変換・ダウンロード
(function () {
'use strict';

const $ = id => document.getElementById(id);
const els = {
  drop: $('drop'), input: $('file-input'), pick: $('pick-btn'),
  listSection: $('list-section'), list: $('list'), count: $('count'),
  sort: $('sort-btn'), clear: $('clear-btn'), convert: $('convert-btn'),
  progressSection: $('progress-section'), progress: $('progress'), progressText: $('progress-text'),
  resultSection: $('result-section'), resultText: $('result-text'), download: $('download-link'), share: $('share-btn'),
  errorSection: $('error-section'), errorList: $('error-list'),
};

const ICONS = { image: '🖼️', heic: '🖼️', pdf: '📕', text: '📝', markdown: '📝', csv: '📊', excel: '📊', word: '📘', powerpoint: '📙', html: '🌐' };

let items = [];          // { id, file, kind, label, reason }
let nextId = 1;
let busy = false;
let resultUrl = null;
let resultFile = null;

// ---------- ファイルの追加 ----------

function addFiles(fileList) {
  for (const file of fileList) {
    const { kind, label, reason } = Conv.detectKind(file.name, file.type);
    items.push({ id: nextId++, file, kind, label, reason });
  }
  clearResult();
  render();
}

els.pick.addEventListener('click', () => els.input.click());
els.input.addEventListener('change', () => {
  addFiles(els.input.files);
  els.input.value = '';               // 同じファイルをもう一度選べるようにする
});

// 画面のどこにドロップしても追加できるようにする
let dragDepth = 0;
window.addEventListener('dragenter', e => {
  if (!e.dataTransfer || ![...e.dataTransfer.types].includes('Files')) return;
  dragDepth++;
  els.drop.classList.add('dragging');
});
window.addEventListener('dragleave', () => {
  if (--dragDepth <= 0) { dragDepth = 0; els.drop.classList.remove('dragging'); }
});
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => {
  e.preventDefault();
  dragDepth = 0;
  els.drop.classList.remove('dragging');
  if (busy || !e.dataTransfer) return;
  addFiles(e.dataTransfer.files);
});

// ---------- 一覧 ----------

function formatSize(n) {
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return (n / 1024).toFixed(0) + ' KB';
  return (n / 1024 / 1024).toFixed(1) + ' MB';
}

function button(text, label, onClick, disabled) {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'icon-btn';
  b.textContent = text;
  b.setAttribute('aria-label', label);
  b.title = label;
  b.disabled = !!disabled || busy;
  b.addEventListener('click', onClick);
  return b;
}

function move(index, delta) {
  const to = index + delta;
  if (to < 0 || to >= items.length) return;
  [items[index], items[to]] = [items[to], items[index]];
  clearResult();
  render();
}

function render() {
  els.listSection.hidden = items.length === 0;
  const ok = items.filter(i => i.kind).length;
  els.count.textContent = items.length ? `（${ok}件${items.length > ok ? `・対応していないもの ${items.length - ok}件` : ''}）` : '';
  els.list.textContent = '';
  items.forEach((it, i) => {
    const li = document.createElement('li');
    li.className = it.kind ? 'item' : 'item bad';

    const icon = document.createElement('span');
    icon.className = 'item-icon';
    icon.textContent = it.kind ? ICONS[it.kind] : '⚠️';

    const info = document.createElement('div');
    info.className = 'item-info';
    const name = document.createElement('div');
    name.className = 'item-name';
    name.textContent = it.file.name;
    const sub = document.createElement('div');
    sub.className = 'item-sub';
    sub.textContent = it.kind ? `${it.label}・${formatSize(it.file.size)}` : it.reason;
    info.append(name, sub);

    const tools = document.createElement('div');
    tools.className = 'item-tools';
    tools.append(
      button('↑', '上へ', () => move(i, -1), i === 0),
      button('↓', '下へ', () => move(i, 1), i === items.length - 1),
      button('×', '外す', () => { items.splice(i, 1); clearResult(); render(); }),
    );
    li.append(icon, info, tools);
    els.list.appendChild(li);
  });
  els.convert.disabled = busy || ok === 0;
  els.sort.disabled = els.clear.disabled = busy;
  els.pick.disabled = busy;
  document.querySelectorAll('input[name="mode"]').forEach(r => { r.disabled = busy; });
}

const collator = new Intl.Collator('ja', { numeric: true, sensitivity: 'base' });
els.sort.addEventListener('click', () => {
  items.sort((a, b) => collator.compare(a.file.name, b.file.name));   // IMG_2.jpg が IMG_10.jpg より前になる
  clearResult();
  render();
});
els.clear.addEventListener('click', () => { items = []; clearResult(); render(); });

// ---------- 変換 ----------

function baseName(name) {
  return name.replace(/\.[^.]+$/, '') || 'ファイル';
}

function stamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}

// 同じ名前のPDFが重ならないようにする（報告書.pdf, 報告書 (2).pdf …）
function uniqueName(name, used) {
  let candidate = name, n = 2;
  while (used.has(candidate)) candidate = name.replace(/\.pdf$/, ` (${n++}).pdf`);
  used.add(candidate);
  return candidate;
}

function showProgress(done, total, name, page, pages) {
  els.progress.max = total;
  els.progress.value = done + (pages ? page / pages : 0);
  els.progressText.textContent = `${Math.min(done + 1, total)} / ${total}　${name} を変換中…` + (pages ? `（${page}/${pages}ページ）` : '');
}

function clearResult() {
  if (resultUrl) URL.revokeObjectURL(resultUrl);
  resultUrl = null;
  resultFile = null;
  els.resultSection.hidden = true;
  els.errorSection.hidden = true;
}

async function run() {
  const targets = items.filter(i => i.kind);
  if (!targets.length || busy) return;
  const mode = document.querySelector('input[name="mode"]:checked').value;
  busy = true;
  clearResult();
  render();
  els.progressSection.hidden = false;
  const failures = [];
  const outputs = [];          // { name, bytes }
  const used = new Set();

  try {
    const merged = mode === 'merge' ? await PdfBuilder.create() : null;
    for (let i = 0; i < targets.length; i++) {
      const it = targets[i];
      showProgress(i, targets.length, it.file.name);
      try {
        const pages = await Conv.convert(it.file, it.kind, (p, n) => showProgress(i, targets.length, it.file.name, p, n));
        if (merged) {
          await merged.add(pages);
        } else {
          const b = await PdfBuilder.create();
          await b.add(pages);
          outputs.push({ name: uniqueName(baseName(it.file.name) + '.pdf', used), bytes: await b.save() });
        }
      } catch (e) {
        console.error(it.file.name, e);
        failures.push({ name: it.file.name, message: (e && e.message) || '原因不明のエラーです' });
      }
    }
    if (merged && merged.pageCount) {
      const ok = targets.length - failures.length;
      const name = ok === 1 ? baseName(targets.find(t => !failures.some(f => f.name === t.file.name)).file.name) + '.pdf' : `まとめ_${stamp()}.pdf`;
      outputs.push({ name, bytes: await merged.save(), pages: merged.pageCount });
    }

    if (outputs.length === 1) {
      const o = outputs[0];
      showResult(new File([o.bytes], o.name, { type: 'application/pdf' }), o.pages ? `${o.pages}ページ` : '');
    } else if (outputs.length > 1) {
      // 別々のPDFが複数あるときは、ZIP 1つにまとめる（スマホで何回もダウンロードしなくて済むように）
      els.progressText.textContent = 'ZIP にまとめています…';
      await Conv.loadLib('jszip.min.js');
      const zip = new JSZip();
      outputs.forEach(o => zip.file(o.name, o.bytes));
      const blob = await zip.generateAsync({ type: 'blob' });
      showResult(new File([blob], `PDF_${stamp()}.zip`, { type: 'application/zip' }), `PDF ${outputs.length}個（ZIP にまとめました）`);
    }
  } catch (e) {
    console.error(e);
    failures.push({ name: '全体', message: (e && e.message) || '原因不明のエラーです' });
  } finally {
    busy = false;
    els.progressSection.hidden = true;
    showFailures(failures);
    render();
  }
}

els.convert.addEventListener('click', run);

function showResult(file, detail) {
  resultFile = file;
  resultUrl = URL.createObjectURL(file);
  els.download.href = resultUrl;
  els.download.download = file.name;
  els.resultText.textContent = `✅ できました：${file.name}` + (detail ? `（${detail}）` : '');
  els.share.hidden = !(navigator.canShare && navigator.canShare({ files: [file] }));
  els.resultSection.hidden = false;
  els.resultSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

els.share.addEventListener('click', async () => {
  if (!resultFile) return;
  try {
    await navigator.share({ files: [resultFile], title: resultFile.name });
  } catch (e) {
    if (e.name !== 'AbortError') alert('共有できませんでした。「ダウンロード」を使ってください。');
  }
});

function showFailures(failures) {
  els.errorList.textContent = '';
  for (const f of failures) {
    const li = document.createElement('li');
    const b = document.createElement('b');
    b.textContent = f.name;
    li.append(b, '：' + f.message);
    els.errorList.appendChild(li);
  }
  els.errorSection.hidden = failures.length === 0;
}

// 変換中にページを閉じそうになったら確認する
window.addEventListener('beforeunload', e => {
  if (busy) { e.preventDefault(); e.returnValue = ''; }
});

render();
})();
