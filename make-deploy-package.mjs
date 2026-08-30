// 生成 GitHub Pages 部署包 deploy-package.zip
// 用法：node make-deploy-package.mjs
import { createRequire } from 'module';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const require = createRequire(import.meta.url);
const JSZip = require('./vendor/jszip.min.js');
const root = path.dirname(fileURLToPath(import.meta.url));

// 与历史部署包保持一致的文件清单
const FILES = ['index.html', 'app.js', 'styles.css', 'DEPLOY.md', 'README.md'];
const DIRS = ['vendor'];

const zip = new JSZip();
for (const f of FILES) {
  const p = path.join(root, f);
  if (!fs.existsSync(p)) { console.error('缺少文件：' + f); process.exit(1); }
  zip.file(f, fs.readFileSync(p));
}
for (const d of DIRS) {
  const dp = path.join(root, d);
  for (const f of fs.readdirSync(dp)) {
    if (fs.statSync(path.join(dp, f)).isFile()) zip.file(d + '/' + f, fs.readFileSync(path.join(dp, f)));
  }
}

const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
const out = path.join(root, 'deploy-package.zip');
fs.writeFileSync(out, buf);
console.log('已生成 ' + path.basename(out) + '（' + (buf.length / 1024).toFixed(1) + ' KB）');
for (const k of Object.keys(zip.files)) console.log('  · ' + k);
