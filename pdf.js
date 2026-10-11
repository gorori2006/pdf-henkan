// 変換したページを pdf-lib で1冊の PDF に組み立てる
(function (global) {
'use strict';

async function createBuilder() {
  await global.Conv.loadLib('pdf-lib.min.js');
  const { PDFDocument } = global.PDFLib;
  const out = await PDFDocument.create();
  out.setCreator('PDF変換');
  out.setProducer('PDF変換');

  return {
    get pageCount() { return out.getPageCount(); },

    // 1ファイル分のページを追加する。途中で失敗したら、そのファイルのページは1枚も残さない
    async add(pages) {
      const before = out.getPageCount();
      try {
        for (const p of pages) {
          if (p.type === 'pdf') {
            const copied = await out.copyPages(p.doc, p.doc.getPageIndices());
            copied.forEach(pg => out.addPage(pg));
          } else {
            const img = p.format === 'png' ? await out.embedPng(p.bytes) : await out.embedJpg(p.bytes);
            const page = out.addPage([p.pageW, p.pageH]);
            page.drawImage(img, { x: p.x, y: p.pageH - p.y - p.h, width: p.w, height: p.h });   // PDF は下からの座標
          }
        }
      } catch (e) {
        while (out.getPageCount() > before) out.removePage(out.getPageCount() - 1);
        throw e;
      }
    },

    save() { return out.save(); },
  };
}

global.PdfBuilder = { create: createBuilder };
})(typeof window !== 'undefined' ? window : globalThis);
