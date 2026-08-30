// 等价验证：复刻 app.js 改动后的 Excel 写入路径（字节化 + 释放 base64 + 指纹去重 + 零拷贝 sink），
// 校验 ①图片字节无损 ②media 去重数正确 ③锚点数正确 ④二次导出（缓存已无 base64）仍正常。
import { createRequire } from 'module';
import crypto from 'crypto';

const require = createRequire(import.meta.url);
const ExcelJS = require('../vendor/exceljs.min.js');

// ===== 以下三个函数与 app.js 中的实现逐字一致 =====
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
  try {
    const b = b64ToBytes(img.base64);
    img.bytes = b;
    img.base64 = null;
    return b;
  } catch (e) { return null; }
}
function imgFingerprint(bytes) {
  if (!bytes || !bytes.length) return '';
  const n = bytes.length;
  const seg = (s, e) => String.fromCharCode.apply(null, bytes.subarray(s, e));
  const mid = Math.max(0, (n >> 1) - 512);
  return n + '|' + seg(0, Math.min(1024, n)) + '|' + seg(mid, Math.min(n, mid + 1024)) + '|' + seg(Math.max(0, n - 1024), n);
}
// ================================================

function makeJpeg(kb) {
  const bytes = kb * 1024;
  const b = crypto.randomBytes(bytes);
  b[0] = 0xff; b[1] = 0xd8; b[bytes - 2] = 0xff; b[bytes - 1] = 0xd9;
  return b;
}

const rows = 120, uniqN = 40, KB = 60; // 3 倍重复率
const uniqRaw = [];
for (let i = 0; i < uniqN; i++) uniqRaw.push(makeJpeg(KB));

// 构造 img 对象（模拟 imgCache）：只用 base64 初始化，与取图产物一致
const cache = [];
for (let r = 0; r < rows; r++) {
  const raw = uniqRaw[r % uniqN];
  cache.push({ base64: raw.toString('base64'), extension: 'jpeg', width: 1200, height: 900 });
}
// 保存原始字节用于事后比对
const expectBytes = uniqRaw.map((b) => Buffer.from(b));

async function exportOnce(imgs, label) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('T');
  ws.addRow(['图']);
  const mediaSeen = new Map();
  let mediaReused = 0;
  for (let r = 0; r < rows; r++) {
    const img = imgs[r];
    const bytes = imgBytes(img);
    if (!bytes) throw new Error('字节化失败');
    const fp = imgFingerprint(bytes);
    let imgId = mediaSeen.get(fp);
    if (imgId === undefined) {
      imgId = wb.addImage({ buffer: bytes, extension: img.extension });
      mediaSeen.set(fp, imgId);
    } else mediaReused++;
    ws.addImage(imgId, { tl: { col: 0, row: r + 1 }, ext: { width: 120, height: 90 }, editAs: 'oneCell' });
  }
  const chunks = [];
  const sink = {
    write(chunk, cb) { chunks.push(chunk); if (typeof cb === 'function') cb(); return true; },
    end() {},
  };
  await wb.xlsx.write(sink, { zip: { compression: 'STORE' } });
  return { chunks, media: wb.model.media.length, mediaReused, label };
}

// ---- 第一次导出（img 只有 base64）----
const r1 = await exportOnce(cache, '首次导出');
console.log(`首次导出：media=${r1.media}（期望 ${uniqN}） · 复用=${r1.mediaReused}（期望 ${rows - uniqN}） · chunks=${r1.chunks.length}`);
console.log(`  base64 是否已释放：${cache.every((i) => i.base64 === null) ? '是' : '否'} · 是否都带 bytes：${cache.every((i) => !!i.bytes) ? '是' : '否'}`);

// ---- 第二次导出（模拟重试/再导一次：此时 imgCache 已无 base64，只剩 bytes）----
const r2 = await exportOnce(cache, '二次导出');
console.log(`二次导出：media=${r2.media}（期望 ${uniqN}） · 复用=${r2.mediaReused}（期望 ${rows - uniqN}） · chunks=${r2.chunks.length}`);

// ---- 校验产出的 xlsx ----
const ab = Buffer.from(await new Blob(r1.chunks).arrayBuffer());
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(ab);
const ws2 = wb2.getWorksheet('T');

console.log('\n--- 正确性校验 ---');
const chk = [];
chk.push(['media 数 = 去重后张数', wb2.model.media.length === uniqN, `${wb2.model.media.length}/${uniqN}`]);
chk.push(['锚点数 = 总行数', (ws2.model.media ? ws2.model.media.length : -1) === rows, `${ws2.model.media.length}/${rows}`]);

// 图片字节无损：把读回的 media buffer 与原始字节逐一比对
let allMatch = true;
const got = wb2.model.media.map((m) => Buffer.from(m.buffer));
for (let i = 0; i < uniqN; i++) {
  if (!got[i] || got[i].length !== expectBytes[i].length || !got[i].equals(expectBytes[i])) { allMatch = false; break; }
}
chk.push(['图片字节与原始完全一致（无损）', allMatch, allMatch ? `${uniqN} 张全对` : '存在不一致']);
chk.push(['二次导出产物与首次一致', Buffer.from(await new Blob(r2.chunks).arrayBuffer()).equals(ab), '字节级相同']);

let ok = true;
for (const [name, pass, detail] of chk) {
  console.log(`  ${pass ? 'PASS' : 'FAIL'}  ${name}  → ${detail}`);
  if (!pass) ok = false;
}
const sizeMB = (ab.length / 1024 / 1024).toFixed(2);
console.log(`\n产物大小 ${sizeMB} MB（未去重应为 ${(rows * KB / 1024).toFixed(2)} MB）`);
console.log(ok ? '\n全部通过 ✓' : '\n存在失败项 ✗');
process.exit(ok ? 0 : 1);
