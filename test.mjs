// converters.js の計算部分（ブラウザがなくても動く部分）を検証する。
//   実行: node test.mjs
import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const dir = path.dirname(fileURLToPath(import.meta.url));
const ctx = vm.createContext({ TextDecoder, console });
vm.runInContext(fs.readFileSync(path.join(dir, 'converters.js'), 'utf8'), ctx);
const C = ctx.Conv;

let pass = 0, fail = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) { pass++; return; }
  fail++;
  console.log(`NG  ${name}\n    期待: ${e}\n    実際: ${a}`);
}
const near = (a, b) => Math.abs(a - b) < 0.01;

// ---- 形式の判定 ----
const kind = (name, mime = '') => C.detectKind(name, mime).kind;
check('jpg', kind('IMG_0001.JPG'), 'image');
check('heic', kind('写真.heic'), 'heic');
check('pdf', kind('a.pdf'), 'pdf');
check('txt', kind('メモ.txt'), 'text');
check('md', kind('README.md'), 'markdown');
check('csv', kind('名簿.csv'), 'csv');
check('tsv', kind('名簿.tsv'), 'csv');
check('xlsx', kind('売上.xlsx'), 'excel');
check('xls', kind('古い表.xls'), 'excel');
check('docx', kind('報告書.docx'), 'word');
check('pptx', kind('提案.pptx'), 'powerpoint');
check('html', kind('page.htm'), 'html');
check('ドットが複数', kind('2026.10.09.報告.docx'), 'word');
check('拡張子なし・MIMEで判定', kind('scan', 'image/jpeg'), 'image');
check('拡張子なし・不明', kind('noext', ''), null);
check('.doc は非対応', kind('古い.doc'), null);
check('.doc の理由に案内', /docx/.test(C.detectKind('古い.doc').reason), true);
check('.pages の理由に案内', /書き出す/.test(C.detectKind('資料.pages').reason), true);
check('ラベル', C.detectKind('a.pptx').label, 'PowerPoint');

// ---- 文字コード ----
const enc = s => new Uint8Array(Buffer.from(s, 'utf8'));
check('UTF-8', C.decodeText(enc('こんにちは')), 'こんにちは');
check('UTF-8 BOM付き', C.decodeText(new Uint8Array([0xEF, 0xBB, 0xBF, ...enc('あ')])), 'あ');
// 「日本」の Shift_JIS は 93 FA 96 7B、①は 87 40
check('Shift_JIS', C.decodeText(new Uint8Array([0x93, 0xFA, 0x96, 0x7B, 0x87, 0x40])), '日本①');
check('UTF-16LE BOM付き', C.decodeText(new Uint8Array([0xFF, 0xFE, 0x42, 0x30])), 'あ');
check('空', C.decodeText(new Uint8Array([])), '');

// ---- ページ割り ----
const lines = n => Array.from({ length: n }, (_, i) => ({ top: i * 20, bottom: i * 20 + 18 }));
check('1ページに収まる', C.computeBreaks(lines(5), 1000, 98), [[0, 98]]);
check('空の文書でも1ページ', C.computeBreaks([], 1000, 0), [[0, 0]]);
// 行の高さ20、ページ高さ110 → 100 の位置（6行目の上）で切れる
check('行の途中で切らない', C.computeBreaks(lines(10), 110, 198)[0], [0, 100]);
{
  const pages = C.computeBreaks(lines(200), 1000, 3998);
  const bad = pages.some(([, e]) => lines(200).some(r => r.top < e - 0.5 && r.bottom > e + 0.5));
  check('長い文書: どの区切りも行にかからない', bad, false);
  check('長い文書: すきまなく続く', pages.every(([s], i) => i === 0 || s === pages[i - 1][1]), true);
  check('長い文書: 最後まで含む', pages[pages.length - 1][1], 3998);
}
check('1ページより高い画像は途中で切る', C.computeBreaks([{ top: 0, bottom: 2500 }], 1000, 2500), [[0, 1000], [1000, 2000], [2000, 2500]]);
check('強制改ページ', C.computeBreaks(lines(10), 1000, 198, [60]), [[0, 60], [60, 198]]);
check('範囲外の強制改ページは無視', C.computeBreaks(lines(3), 1000, 58, [0, 58, 999]), [[0, 58]]);

// ---- 画像の配置 ----
{
  const f = C.fitImage(4000, 3000);   // 横長の写真
  check('横長 → 横向きページ', f.pageW > f.pageH, true);
  check('余白内に収まる', f.w <= f.pageW - 72 + 0.01 && f.h <= f.pageH - 72 + 0.01, true);
  check('縦横比を保つ', near(f.w / f.h, 4 / 3), true);
  check('中央に置く', near(f.x * 2 + f.w, f.pageW) && near(f.y * 2 + f.h, f.pageH), true);
  const t = C.fitImage(1000, 3000);
  check('縦長 → 縦向きページ', t.pageW < t.pageH, true);
  const s = C.fitImage(100, 50);
  check('小さい画像は引き伸ばさない', near(s.w, 75) && near(s.h, 37.5), true);
}

// ---- PowerPoint 用 ----
check('相対パス', C.resolvePath('ppt/slides/', '../media/image1.png'), 'ppt/media/image1.png');
check('同じフォルダ', C.resolvePath('ppt/', 'slides/slide1.xml'), 'ppt/slides/slide1.xml');
check('絶対パス', C.resolvePath('ppt/slides/', '/ppt/media/a.png'), 'ppt/media/a.png');
check('明るさ: 変更なし', C.applyLum('#336699'), '#336699');
check('明るさ: 黒を25%明るく', C.applyLum('#000000', 75000, 25000), '#404040');
check('明るさ: 白を50%暗く', C.applyLum('#ffffff', 50000, 0), '#808080');
check('番号 1.', C.autoNumber('arabicPeriod', 1), '1.');
check('番号 b)', C.autoNumber('alphaLcParenR', 2), 'b)');
check('番号 (IV)', C.autoNumber('romanUcParenBoth', 4), '(IV)');
check('番号 ③', C.autoNumber('circleNumDbPlain', 3), '③');

// ---- テキスト ----
check('テキストをHTMLに（エスケープ）', C.textToHtml('<b>&\n2行目\n'), '<div class="plain"><div class="line">&lt;b&gt;&amp;</div><div class="line">2行目</div></div>');
check('改行コード CRLF', C.textToHtml('a\r\nb'), '<div class="plain"><div class="line">a</div><div class="line">b</div></div>');

console.log(`\n${pass + fail}項目中 ${pass}項目 OK${fail ? `、${fail}項目 NG` : ''}`);
process.exit(fail ? 1 : 0);
