// 直击「打包环节是否降低图片质量」：比对三种路径产出的图片字节是否完全一致。
// 路径 A：改动前的实际行为（compression 写在顶层 → 实际 DEFLATE + base64 嵌入）
// 路径 B：改动后的行为（zip:{compression:'STORE'} + 字节嵌入 + 去重）
// 判定：只要三者内嵌的图片字节与源图逐字节相同（SHA-256 一致），即证明打包环节零质量损失。
import { createRequire } from 'module';
import crypto from 'crypto';

const require = createRequire(import.meta.url);
const ExcelJS = require('../vendor/exceljs.min.js');

function b64ToBytes(b64) {
  const bin = atob(b64);
  const n = bin.length;
  const out = new Uint8Array(n);
  for (let i = 0; i < n; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function imgBytes(img) {
  if (!img) return null;
  if (img.bytes) return img.bytes;
  if (!img.base64) return null;
  try { const b = b64ToBytes(img.base64); img.bytes = b; img.base64 = null; return b; } catch (e) { return null; }
}
function imgFingerprint(bytes) {
  if (!bytes || !bytes.length) return '';
  const n = bytes.length;
  const seg = (s, e) => String.fromCharCode.apply(null, bytes.subarray(s, e));
  const mid = Math.max(0, (n >> 1) - 512);
  return n + '|' + seg(0, Math.min(1024, n)) + '|' + seg(mid, Math.min(n, mid + 1024)) + '|' + seg(Math.max(0, n - 1024), n);
}

const rows = 90, uniqN = 30, KB = 55; // 3 倍重复，用于同时检验去重
function makeSrc(kb) {
  const bytes = kb * 1024;
  const b = crypto.randomBytes(bytes);
  b[0] = 0xff; b[1] = 0xd8; b[bytes - 2] = 0xff; b[bytes - 1] = 0xd9;
  return b;
}
const srcRaw = [];
for (let i = 0; i < uniqN; i++) srcRaw.push(makeSrc(KB));
const srcHash = srcRaw.map((b) => crypto.createHash('sha256').update(b).digest('hex'));

// 每行引用的图（含重复引用）
const rowOf = [];
for (let r = 0; r < rows; r++) rowOf.push(r % uniqN);

// ---- 路径 A：改动前（顶层 compression，实际走 DEFLATE；base64 嵌入，不去重）----
async function pathA() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.addRow(['图']);
  const imgs = rowOf.map((i) => ({ base64: srcRaw[i].toString('base64'), extension: 'jpeg', width: 1200, height: 900 }));
  for (let r = 0; r < rows; r++) {
    const id = wb.addImage({ base64: imgs[r].base64, extension: 'jpeg' });
    ws.addImage(id, { tl: { col: 0, row: r + 1 }, ext: { width: 120, height: 90 }, editAs: 'oneCell' });
  }
  const buf = await wb.xlsx.writeBuffer({ compression: 'STORE' }); // 顶层写法：实际未生效
  return { ab: Buffer.from(buf), media: wb.model.media.length };
}

// ---- 路径 B：改动后（zip 内 STORE；字节嵌入；内容去重）----
async function pathB() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.addRow(['图']);
  const imgs = rowOf.map((i) => ({ base64: srcRaw[i].toString('base64'), extension: 'jpeg', width: 1200, height: 900 }));
  const seen = new Map();
  let reused = 0;
  for (let r = 0; r < rows; r++) {
    const bytes = imgBytes(imgs[r]);
    const fp = imgFingerprint(bytes);
    let id = seen.get(fp);
    if (id === undefined) { id = wb.addImage({ buffer: bytes, extension: 'jpeg' }); seen.set(fp, id); }
    else reused++;
    ws.addImage(id, { tl: { col: 0, row: r + 1 }, ext: { width: 120, height: 90 }, editAs: 'oneCell' });
  }
  const chunks = [];
  const sink = { write(c, cb) { chunks.push(c); if (typeof cb === 'function') cb(); return true; }, end() {} };
  await wb.xlsx.write(sink, { zip: { compression: 'STORE' } });
  return { ab: Buffer.from(await new Blob(chunks).arrayBuffer()), media: wb.model.media.length, reused };
}

async function extractMedia(ab) {
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(ab);
  const ws2 = wb2.getWorksheet('S');
  return {
    bytes: wb2.model.media.map((m) => Buffer.from(m.buffer)),
    anchors: ws2.model.media.length,
  };
}

const A = await pathA();
const B = await pathB();
const ea = await extractMedia(A.ab);
const eb = await extractMedia(B.ab);

console.log(`规模：${rows} 行（不同图 ${uniqN} 张，3 倍重复引用），每张 ${KB}KB`);
console.log(`路径A（改动前）：文件 ${(A.ab.length / 1024 / 1024).toFixed(2)} MB · media=${A.media} · 锚点=${ea.anchors}`);
console.log(`路径B（改动后）：文件 ${(B.ab.length / 1024 / 1024).toFixed(2)} MB · media=${B.media}（复用 ${B.reused}） · 锚点=${eb.anchors}`);

const hashOf = (b) => crypto.createHash('sha256').update(b).digest('hex');
let okAA = true, okBB = true, okAB = true;
for (let i = 0; i < uniqN; i++) {
  if (hashOf(ea.bytes[i]) !== srcHash[i]) { okAA = false; console.log(`  A 第${i}张不一致`); }
  if (hashOf(eb.bytes[i]) !== srcHash[i]) { okBB = false; console.log(`  B 第${i}张不一致`); }
  if (hashOf(ea.bytes[i]) !== hashOf(eb.bytes[i])) { okAB = false; console.log(`  A/B 第${i}张互不相同`); }
}

console.log('\n--- 图片质量校验（以字节为单位，任何降质都会导致哈希不同）---');
const chk = [
  ['路径A 内嵌图片 == 源图字节', okAA, `${ea.bytes.length} 张全部 SHA-256 一致`],
  ['路径B 内嵌图片 == 源图字节', okBB, `${eb.bytes.length} 张全部 SHA-256 一致`],
  ['路径A 与 路径B 图片字节完全相同', okAB, '两种打包方式产出同一份像素'],
  ['路径B 锚点数 == 总行数', eb.anchors === rows, `${eb.anchors}/${rows}（去重未丢任何一行的图）`],
  ['路径B media 数 == 去重后张数', B.media === uniqN, `${B.media}/${uniqN}`],
];
let ok = true;
for (const [name, pass, detail] of chk) {
  console.log(`  ${pass ? 'PASS' : 'FAIL'}  ${name}  → ${detail}`);
  if (!pass) ok = false;
}
console.log(`\n单张图片字节数：源图 ${srcRaw[0].length} · A ${ea.bytes[0].length} · B ${eb.bytes[0].length}`);
console.log(ok ? '\n结论：打包环节零质量损失 ✓' : '\n存在质量损失 ✗');
process.exit(ok ? 0 : 1);
