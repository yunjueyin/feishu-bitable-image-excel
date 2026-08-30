// 基准：量化「生成文件阶段」各写法的耗时。只测写文件，不涉及取图/打勾/图片质量。
// 关键：数据用 crypto.randomBytes（真随机、不可压缩），贴近真实已压缩的 JPEG/PNG。
import { createRequire } from 'module';
import crypto from 'crypto';

const require = createRequire(import.meta.url);
const ExcelJS = require('../vendor/exceljs.min.js');

const rows = Number(process.argv[2] || 200);
const uniqImgs = Number(process.argv[3] || rows);
const KB = Number(process.argv[4] || 200);

function makeFakeJpeg(kb) {
  const bytes = kb * 1024;
  const b = crypto.randomBytes(bytes); // 不可压缩，等同真实 JPEG 熵
  b[0] = 0xff; b[1] = 0xd8; b[bytes - 2] = 0xff; b[bytes - 1] = 0xd9;
  return b;
}

console.log(`规模：${rows} 行 · 每张 ${KB}KB · 不同图片 ${uniqImgs} 张 → 图片总量 ≈ ${(rows * KB / 1024).toFixed(1)} MB\n`);

const uniq = [];
for (let i = 0; i < uniqImgs; i++) {
  const buf = makeFakeJpeg(KB);
  uniq.push({ base64: buf.toString('base64'), buffer: Buffer.from(buf) });
}
const rowImgs = [];
for (let r = 0; r < rows; r++) rowImgs.push(uniq[r % uniqImgs]);

function fp(img) {
  const s = img.base64;
  return s.length + ':' + s.slice(0, 256) + ':' + s.slice(-256);
}
const mb = (n) => (n / 1024 / 1024).toFixed(1) + ' MB';

function build({ dedupe = false, asBuffer = false } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.addRow(['图']);
  const seen = new Map();
  let reused = 0;
  for (let r = 0; r < rows; r++) {
    const img = rowImgs[r];
    let id;
    if (dedupe) {
      const k = fp(img);
      id = seen.get(k);
      if (id === undefined) {
        id = wb.addImage(asBuffer ? { buffer: img.buffer, extension: 'jpeg' } : { base64: img.base64, extension: 'jpeg' });
        seen.set(k, id);
      } else reused++;
    } else {
      id = wb.addImage(asBuffer ? { buffer: img.buffer, extension: 'jpeg' } : { base64: img.base64, extension: 'jpeg' });
    }
    ws.addImage(id, { tl: { col: 0, row: r + 1 }, ext: { width: 120, height: 90 }, editAs: 'oneCell' });
  }
  return { wb, reused, media: wb.model.media.length };
}

async function run(label, { dedupe = false, asBuffer = false, zipOpts = {}, useSink = false } = {}) {
  global.gc && global.gc();
  const { wb, reused, media } = build({ dedupe, asBuffer });
  const t0 = Date.now();
  let blob;
  let chunks = 0;
  if (useSink) {
    const arr = [];
    const sink = {
      write(chunk, cb) { arr.push(chunk); if (typeof cb === 'function') cb(); return true; },
      end() {},
    };
    await wb.xlsx.write(sink, zipOpts);
    chunks = arr.length;
    blob = new Blob(arr, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  } else {
    const buf = await wb.xlsx.writeBuffer(zipOpts);
    blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  const ms = Date.now() - t0;
  const r = { label, ms, size: blob.size, media, reused, chunks, blob };
  console.log(`${label.padEnd(38)} ${String(ms).padStart(6)}ms · 产物 ${mb(blob.size).padStart(10)} · media=${String(media).padStart(4)}${reused ? ' 复用' + reused : ''}${chunks ? ' chunks=' + chunks : ''}`);
  return r;
}

const results = [];
// A：现状（compression 写在顶层 → 实际未生效，走默认 DEFLATE）
results.push(await run('A 现状 writeBuffer{STORE顶层}', { zipOpts: { compression: 'STORE' } }));
// B：修复参数层级 zip:{compression:'STORE'}
results.push(await run('B 修复 zip:{STORE}', { zipOpts: { zip: { compression: 'STORE' } } }));
// C：修复 STORE + 零拷贝 sink
results.push(await run('C STORE + 零拷贝sink', { zipOpts: { zip: { compression: 'STORE' } }, useSink: true }));
// D：修复 STORE + 传 buffer（跳过 JSZip base64 解码）
results.push(await run('D STORE + addImage(buffer)', { zipOpts: { zip: { compression: 'STORE' } }, asBuffer: true }));
// E：修复 STORE + 零拷贝 + 传 buffer
results.push(await run('E STORE + 零拷贝 + buffer', { zipOpts: { zip: { compression: 'STORE' } }, useSink: true, asBuffer: true }));
// F：E + 去重
results.push(await run('F E + 图片去重', { zipOpts: { zip: { compression: 'STORE' } }, useSink: true, asBuffer: true, dedupe: true }));

console.log('\n--- 产物校验（用 ExcelJS 重新读回）---');
for (const r of [results[0], results[2], results[5]]) {
  const ab = Buffer.from(await r.blob.arrayBuffer());
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(ab);
  const ws2 = wb2.getWorksheet('S');
  const anchors = ws2.model.media ? ws2.model.media.length : -1;
  console.log(`  ${r.label.padEnd(38)} media=${wb2.model.media.length}（期望${r.media}） 锚点=${anchors}（期望${rows}） ${wb2.model.media.length === r.media && anchors === rows ? 'OK' : 'MISMATCH'}`);
}

console.log('\n--- 相对现状提速 ---');
const base = results[0].ms;
for (const r of results) {
  const d = ((1 - r.ms / base) * 100);
  console.log(`${r.label.padEnd(38)} ${String(r.ms).padStart(6)}ms  ${(d >= 0 ? '-' : '+')}${Math.abs(d).toFixed(1)}% ${d >= 0 ? '（快）' : '（慢）'}`);
}
