// 多维表图片导出 Excel —— 飞书多维表自定义插件（纯前端）
// 依赖：ExcelJS（UMD 全局，已在 index.html 引入，失败回退 CDN）
// SDK：@lark-base-open/js-sdk（ESM，动态 import，带 CDN 兜底）

// ---------- 常量 ----------
const FIELD_TYPE_ATTACHMENT = 17;
const FIELD_TYPE_CHECKBOX = 7;
const SUPPORTED_IMG = ['png', 'jpeg', 'gif', 'bmp'];
const ImageQualityFallback = { Low: 120, Mid: 360, HIGH: 720, MAX: 1280 };
const THUMB_QUALITY_HIGH = 2560; // 尝试高于 SDK MAX(1280) 的缩略图质量，飞书服务端若支持则返回更大图

// ---------- 状态 ----------
const state = {
  bitable: null,
  ImageQuality: ImageQualityFallback,
  table: null,
  tableId: '',
  tableName: '',
  tableMetas: [],    // 当前多维表的所有数据表 [{id, name}]
  fields: [],        // [{id, name, type, isPrimary}]
  records: [],       // [{recordId, fields}]
  maxAttach: {},     // fieldId -> 该字段单格最多附件数
  loaded: false,
  stat: { orig: 0, thumb: 0 }, // 本次导出图片来源统计
  imgQuality: 'thumb', // thumb=缩略图(最快·推荐) / orig=高清原图(本地直连飞书)
};

// ---------- 工具 ----------
const $ = (sel) => document.querySelector(sel);
function log(msg, cls) {
  const el = $('#log');
  const line = document.createElement('div');
  if (cls) line.className = cls;
  line.textContent = msg;
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}
function setStatus(text, kind) {
  $('#statusText').textContent = text;
  $('#status').className = 'status status-' + (kind || 'idle');
}
function setProgress(p) {
  p = Math.max(0, Math.min(100, p));
  const bar = $('#progressBar');
  if (bar) bar.style.width = p + '%';
  const pct = $('#progressPct');
  if (pct) pct.textContent = Math.round(p) + '%';
}
function showProgress() { const b = $('#progressBox'); if (b) b.classList.remove('hidden'); }
function hideProgress() { const b = $('#progressBox'); if (b) b.classList.add('hidden'); }
function setProgressCount(text) { const c = $('#progressCount'); if (c) c.textContent = text || ''; }
function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }
// 带超时的 Promise 包装：超过 ms 即 reject，避免飞书 SDK 调用挂起导致整批导出永久卡死
function withTimeout(promise, ms, label) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error('timeout' + (label ? ' ' + label : '') + ' ' + ms + 'ms')), ms);
  });
  return Promise.race([Promise.resolve(promise), timeout]).finally(() => clearTimeout(timer));
}
// 带超时 + 重试的 SDK 调用：飞书 getCellThumbnailUrls 等服务端生成操作在高并发下会超时，
// 但串行/低并发重试通常能成功。retries 次重试，退避递增（base, base*2, base*3...）。
async function withRetry(fn, { retries = 2, timeoutMs = 8000, label = '', baseDelay = 300 } = {}) {
  let lastErr;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      return await withTimeout(fn(), timeoutMs, label + (attempt > 0 ? ' #' + attempt : ''));
    } catch (e) {
      lastErr = e;
      if (attempt < retries) {
        const d = baseDelay * (attempt + 1);
        log('  ' + label + ' 第' + (attempt + 1) + '次失败（' + (e.message || e) + '），' + d + 'ms 后重试', 'warn');
        await sleep(d);
      }
    }
  }
  throw lastErr;
}

// 全局并发信号量：飞书 getCellThumbnailUrls / getCellAttachmentUrls 是服务端生成/鉴权操作，
// 对并发极敏感；用信号量把「实际在途的 SDK 调用数」硬限到很小的值，避免压垮服务导致整批超时。
function makeLimiter(max) {
  let active = 0;
  const queue = [];
  const next = () => { if (queue.length && active < max) { active++; const run = queue.shift(); run(); } };
  return (fn) => new Promise((resolve, reject) => {
    const run = () => {
      Promise.resolve().then(fn).then(resolve, reject).finally(() => { active--; next(); });
    };
    if (active < max) { active++; run(); } else { queue.push(run); }
  });
}
const thumbLimit = makeLimiter(3);   // 同时在途的缩略图请求 ≤ 3
const attachLimit = makeLimiter(3);  // 同时在途的原图请求 ≤ 3

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src; s.onload = () => resolve(); s.onerror = () => reject(new Error('load fail ' + src));
    document.head.appendChild(s);
  });
}

// ---------- 加载 SDK ----------
async function loadSdk() {
  const urls = [
    'https://cdn.jsdelivr.net/npm/@lark-base-open/js-sdk@1.0.2/dist/index.mjs',
    'https://esm.sh/@lark-base-open/js-sdk@1.0.2',
  ];
  for (const url of urls) {
    try {
      const m = await import(/* @vite-ignore */ url);
      if (m && m.bitable) {
        state.bitable = m.bitable;
        state.ImageQuality = m.ImageQuality || ImageQualityFallback;
        log('飞书 SDK 已加载（' + url + '）');
        return true;
      }
    } catch (e) {
      log('SDK 加载失败：' + url + ' — ' + e.message, 'err');
    }
  }
  return false;
}

// ---------- 加载 ExcelJS（UMD 兜底）----------
async function ensureExcelJS() {
  if (typeof window.ExcelJS !== 'undefined') return true;
  try {
    await loadScript('https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js');
    return typeof window.ExcelJS !== 'undefined';
  } catch (e) {
    log('ExcelJS 加载失败：' + e.message, 'err');
    return false;
  }
}

// ---------- 加载 JSZip（UMD 兜底）----------
async function ensureJSZip() {
  if (typeof window.JSZip !== 'undefined') return true;
  try {
    await loadScript('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js');
    return typeof window.JSZip !== 'undefined';
  } catch (e) {
    log('JSZip 加载失败：' + e.message, 'err');
    return false;
  }
}

// ---------- 初始化 ----------
async function init() {
  const okSdk = await loadSdk();
  if (!okSdk) {
    setStatus('未加载飞书 SDK（请检查网络 / CDN）', 'err');
    log('本插件需作为飞书多维表「自定义插件」打开。', 'err');
    return;
  }
  try {
    // 列出全部数据表，支持用户切换
    try {
      const metas = await state.bitable.base.getTableMetaList();
      state.tableMetas = Array.isArray(metas) ? metas : [];
    } catch (e) {
      log('获取表列表失败（将仅使用当前表）：' + e.message, 'warn');
      state.tableMetas = [];
    }
    const table = await state.bitable.base.getActiveTable();
    state.table = table;
    state.tableId = (table && table.id) || '';
    state.tableName = await table.getName();
    renderTableSelect(state.tableMetas, state.tableId);
    setStatus('已连接：' + state.tableName, 'ok');
    await loadFields();
    $('#btnLoad').disabled = false;
  } catch (e) {
    setStatus('未在飞书多维表环境中，或无法获取当前表：' + e.message, 'err');
    log('若你在普通浏览器打开本页，这是正常的——请作为飞书自定义插件使用。', 'err');
  }
}

async function loadFields() {
  const metas = await state.table.getFieldMetaList();
  state.fields = metas.map((m) => ({
    id: m.id,
    name: m.name,
    type: m.type,
    isPrimary: !!m.isPrimary,
    isAttachment: m.type === FIELD_TYPE_ATTACHMENT,
  }));
  renderFieldList();
  renderMarkOptions();
  renderNamingOptions();
}

function renderFieldList() {
  const box = $('#fieldList');
  box.innerHTML = '';
  if (!state.fields.length) {
    box.innerHTML = '<div class="placeholder">该表没有字段。</div>';
    return;
  }
  const attachments = state.fields.filter((f) => f.isAttachment);
  const others = state.fields.filter((f) => !f.isAttachment);

  const renderGroup = (title, list, defaultChecked) => {
    if (!list.length) return;
    const group = document.createElement('div');
    group.className = 'field-group';
    const head = document.createElement('div');
    head.className = 'field-group-title';
    head.textContent = title;
    group.appendChild(head);
    for (const f of list) {
      const row = document.createElement('label');
      row.className = 'field-item';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = defaultChecked;
      cb.dataset.fieldId = f.id;
      const name = document.createElement('span');
      name.className = 'fname';
      name.textContent = f.name + (f.isPrimary ? '（主键）' : '');
      row.appendChild(cb);
      row.appendChild(name);
      if (f.isAttachment) {
        const tag = document.createElement('span');
        tag.className = 'ftag';
        tag.textContent = '图片';
        row.appendChild(tag);
      }
      group.appendChild(row);
    }
    box.appendChild(group);
  };

  renderGroup('图片列（嵌入单元格）', attachments, true);
  renderGroup('文字列', others, true);
  updateFieldSummary();
}

function selectAllFields(checked) {
  document.querySelectorAll('#fieldList input[type=checkbox]').forEach((cb) => (cb.checked = !!checked));
  updateFieldSummary();
}

function selectImageFields() {
  document.querySelectorAll('#fieldList input[type=checkbox]').forEach((cb) => {
    const fid = cb.dataset.fieldId;
    const f = state.fields.find((x) => x.id === fid);
    cb.checked = !!(f && f.isAttachment);
  });
  updateFieldSummary();
}

// 折叠下拉：切换展开/收起，并同步箭头与无障碍状态
function toggleFieldDropdown() {
  const dd = $('#fieldDropdown');
  if (!dd) return;
  const collapsed = dd.classList.toggle('collapsed');
  const chev = $('#fieldToggle').querySelector('.chev');
  if (chev) chev.classList.toggle('collapsed', collapsed);
  $('#fieldToggle').setAttribute('aria-expanded', String(!collapsed));
}

// 头部摘要：显示「已选 X / N 列」或「共 N 列 · 未选」
function updateFieldSummary() {
  const sum = $('#fieldSummary');
  if (!sum) return;
  const total = state.fields.length;
  if (!total) { sum.textContent = '无字段'; return; }
  const checked = document.querySelectorAll('#fieldList input[type=checkbox]:checked').length;
  sum.textContent = checked === 0 ? ('共 ' + total + ' 列 · 未选') : ('已选 ' + checked + ' / ' + total + ' 列');
}

function renderTableSelect(metas, currentId) {
  const sel = $('#tableSelect');
  if (!sel) return;
  sel.innerHTML = '';
  if (!metas || !metas.length) { sel.classList.add('hidden'); return; }
  for (const m of metas) {
    const o = document.createElement('option');
    o.value = m.id; o.textContent = m.name || m.id;
    sel.appendChild(o);
  }
  if (currentId) sel.value = currentId;
  sel.classList.remove('hidden');
}

async function switchTable(id) {
  if (!state.bitable || !id) return;
  if (id === state.tableId) return;
  try {
    setStatus('正在切换数据表…', 'idle');
    const table = await state.bitable.base.getTableById(id);
    state.table = table;
    state.tableId = id;
    state.tableName = await table.getName();
    state.loaded = false;
    state.records = [];
    $('#btnExport').disabled = true;
    $('#btnExportZip').disabled = true;
    $('#count').textContent = '';
    setStatus('已连接：' + state.tableName, 'ok');
    await loadFields();
    log('已切换数据表：' + state.tableName + '（请重新点「加载数据」）', 'ok');
  } catch (e) {
    setStatus('切换数据表失败：' + e.message, 'err');
    log('切换数据表失败：' + e.message, 'err');
  }
}

function getSelectedFieldIds() {
  // 按表格原始字段顺序返回勾选项（UI 分组展示不影响导出列顺序，保证与表格列序一致）
  const checked = new Set([...document.querySelectorAll('#fieldList input[type=checkbox]:checked')].map((c) => c.dataset.fieldId));
  return state.fields.filter((f) => checked.has(f.id)).map((f) => f.id);
}

// ---------- 读取全部记录 ----------
async function loadData() {
  if (!state.table) return;
  $('#btnLoad').disabled = true;
  $('#btnExport').disabled = true;
  state.loaded = false;
  state.records = [];
  state.maxAttach = {};
  showProgress();
  setProgress(0);
  log('开始读取记录…');
  // 默认按当前视图（含筛选/排序结果）导出全部内容；
  // 仅当用户勾选「忽略筛选/排序 · 导出全部记录」时才拉全部记录
  let viewId;
  if (!$('#ignoreView').checked) {
    try {
      const view = await state.table.getActiveView();
      viewId = view && view.id;
      if (viewId) {
        let vname = viewId;
        try { vname = await view.getName(); } catch (e) {}
        log('按当前视图导出：' + vname + '（筛选/排序结果将一并导出全部内容）');
      }
    } catch (e) {
      log('获取当前视图失败，将导出全部记录：' + e.message, 'warn');
      viewId = null;
    }
  } else {
    log('已忽略视图筛选/排序，导出全部记录');
  }

  try {
    const all = [];
    let pageToken;
    let hasMore = true;
    let page = 0;
    do {
      const resp = await state.table.getRecordsByPage(
        Object.assign({ pageSize: 200, pageToken }, viewId ? { viewId } : {})
      );
      const recs = resp.records || [];
      all.push(...recs);
      setProgress(Math.min(95, Math.round((all.length / (resp.total || all.length)) * 95)));
      hasMore = resp.hasMore;
      pageToken = resp.pageToken;
      page++;
      if (page > 2000) { log('已达分页安全上限，停止读取。', 'warn'); break; }
    } while (hasMore && pageToken);

    state.records = all;
    // 统计每个附件字段单格最多附件数（决定预留几列）
    for (const f of state.fields) {
      if (!f.isAttachment) continue;
      let mx = 0;
      for (const r of all) {
        const cell = r.fields && r.fields[f.id];
        if (Array.isArray(cell)) mx = Math.max(mx, cell.filter((x) => x && x.token).length);
      }
      state.maxAttach[f.id] = mx;
    }
    state.loaded = true;
    $('#count').textContent = '共 ' + all.length + ' 行';
    $('#btnExport').disabled = false;
    $('#btnExportZip').disabled = false;
    setProgress(100);
    log('读取完成，共 ' + all.length + ' 行。', 'ok');
  } catch (e) {
    log('读取失败：' + e.message, 'err');
    setStatus('读取数据失败', 'err');
  } finally {
    $('#btnLoad').disabled = false;
    hideProgress();
  }
}

// ---------- 图片处理 ----------
function mimeToExt(mime) {
  const map = {
    'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg',
    'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp',
    'image/svg+xml': 'png',
  };
  return map[mime] || 'png';
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = () => reject(fr.error);
    fr.readAsDataURL(blob);
  });
}

function canvasToDataUrl(bmp, type) {
  const c = document.createElement('canvas');
  c.width = bmp.width; c.height = bmp.height;
  c.getContext('2d').drawImage(bmp, 0, 0);
  return c.toDataURL(type);
}

// 将 Blob 转成 Excel 可直接使用的 {base64, extension, width, height}（纯转换，不计数）
async function blobToImageResult(blob) {
  const mime = blob.type || 'image/png';
  let ext = mimeToExt(mime);
  let dataUrl;
  if (SUPPORTED_IMG.includes(ext)) {
    dataUrl = await blobToDataUrl(blob);
  } else {
    const bmp = await createImageBitmap(blob);
    dataUrl = canvasToDataUrl(bmp, 'image/png');
    ext = 'png';
  }
  return finalizeExcelImage(dataUrl, ext);
}
// 原图：Blob → Excel 图片（计 orig）
async function blobToExcelImage(blob) {
  const r = await blobToImageResult(blob);
  state.stat.orig++;
  return r;
}

// 原图兜底：Blob → 本地缩放至 targetW 像素 → 转 JPEG（减小体积）→ Excel 图片（计 orig）。
// 用于缩略图服务超时时尽量保住图片；飞书 CDN 偶尔 CORS 拦截导致下载失败，由上层 try/catch 兜底。
async function blobToExcelImageResized(blob, targetW) {
  const mime = blob.type || 'image/png';
  let dataUrl;
  if (SUPPORTED_IMG.includes(mimeToExt(mime))) {
    dataUrl = await blobToDataUrl(blob);
  } else {
    const bmp = await createImageBitmap(blob);
    dataUrl = canvasToDataUrl(bmp, 'image/png');
  }
  const resized = await resizeImage(dataUrl, targetW, 0.85);
  state.stat.orig++;
  return finalizeExcelImage(resized, null);
}

// 缩略图返回可能是：base64 字符串 / data URL / http(s) URL / 对象({url|thumbnail|data|base64})
function detectMime(b64) {
  if (/^iVBORw0KGgo/.test(b64)) return 'image/png';
  if (/^\/9j\//.test(b64)) return 'image/jpeg';
  if (/^R0lGOD/.test(b64)) return 'image/gif';
  if (/^Qk/.test(b64)) return 'image/bmp';
  return 'image/png';
}
function extractThumbStr(item) {
  if (item == null) return '';
  if (typeof item === 'string') return item;
  if (typeof item === 'object') return item.url || item.thumbnail || item.data || item.base64 || item.downloadUrl || '';
  return String(item);
}
async function thumbToExcelImage(item) {
  const s = extractThumbStr(item);
  if (!s) return null;
  // http(s) URL：fetch 取像素（飞书部分版本缩略图返回下载链接而非 base64）
  if (/^https?:\/\//i.test(s)) {
    try {
      const resp = await fetchWithTimeout(s, 8000);
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      const r = await blobToImageResult(await resp.blob());
      state.stat.thumb++;
      return r;
    } catch (e) {
      state.stat.thumb++;
      return null;
    }
  }
  // base64 字符串（可能带 data: 前缀）
  let raw = s;
  if (/^data:/i.test(raw)) raw = raw.split(',')[1];
  if (!raw) return null;
  const mime = detectMime(raw);
  const dataUrl = 'data:' + mime + ';base64,' + raw;
  state.stat.thumb++;
  return finalizeExcelImage(dataUrl, null);
}

async function finalizeExcelImage(dataUrl, forceExt) {
  const m = dataUrl.match(/^data:(image\/[\w+.-]+);base64,(.*)$/);
  const mime = m ? m[1] : 'image/png';
  let b64 = m ? m[2] : dataUrl.split(',')[1];
  let ext = forceExt || mimeToExt(mime);
  if (!SUPPORTED_IMG.includes(ext)) {
    // 不支持的格式，重新编码为 png
    const bmp = await blobToBitmap(dataUrl);
    dataUrl = canvasToDataUrl(bmp, 'image/png');
    const m2 = dataUrl.match(/^data:(image\/png);base64,(.*)$/);
    b64 = m2[2]; ext = 'png';
  }
  const dims = await imageSize(dataUrl);
  if (!dims) return null; // 图坏/无法解码：返回 null，上层跳过，避免 ExcelJS addImage 读 undefined.width 崩溃
  return { base64: b64, extension: ext, width: dims.w, height: dims.h };
}

function blobToBitmap(dataUrl) {
  return fetch(dataUrl).then((r) => r.blob()).then((b) => createImageBitmap(b));
}
function imageDims(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth || 150, h: img.naturalHeight || 120 });
    img.onerror = () => resolve(null); // 图坏：返回 null，让上层跳过该图，避免 ExcelJS 解码失败中断导出
    img.src = dataUrl;
  });
}

// 从 base64 图片头解析宽高（无需解码整张像素，速度远快于 new Image），用于提速
function b64ToBytesAt(b64, maxBytes) {
  const clean = b64.replace(/^data:[^;]+;base64,/, '');
  const need = Math.min(clean.length, Math.ceil((maxBytes + 4) * 4 / 3));
  const bin = atob(clean.slice(0, need));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function imageSizeFromBase64(b64) {
  try {
    const b = b64ToBytesAt(b64, 512);
    // PNG
    if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
      const w = (b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19];
      const h = (b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23];
      if (w > 0 && h > 0) return { w, h };
    }
    // GIF
    if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) {
      const w = b[6] | (b[7] << 8);
      const h = b[8] | (b[9] << 8);
      if (w > 0 && h > 0) return { w, h };
    }
    // BMP
    if (b[0] === 0x42 && b[1] === 0x4d) {
      const w = b[18] | (b[19] << 8) | (b[20] << 16) | (b[21] << 24);
      const h = b[22] | (b[23] << 8) | (b[24] << 16) | (b[25] << 24);
      if (w > 0 && h > 0) return { w, h };
    }
    // JPEG
    if (b[0] === 0xff && b[1] === 0xd8) {
      let i = 2;
      while (i < b.length - 9) {
        if (b[i] !== 0xff) { i++; continue; }
        const m = b[i + 1];
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
          const h = (b[i + 5] << 8) | b[i + 6];
          const w = (b[i + 7] << 8) | b[i + 8];
          if (w > 0 && h > 0) return { w, h };
        }
        const len = (b[i + 2] << 8) | b[i + 3];
        i += 2 + len;
      }
    }
  } catch (e) { /* 解析失败回退 new Image */ }
  return null;
}
// 优先 header 解析拿尺寸（远快），失败再回退加载图片
async function imageSize(dataUrl) {
  const header = imageSizeFromBase64(dataUrl);
  if (header) return header;
  return imageDims(dataUrl);
}

// 把图片缩放（保持比例）到目标宽度，转 JPEG 以减小体积，返回 data URL
function resizeImage(dataUrl, targetW, quality) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const w = img.naturalWidth || 1, h = img.naturalHeight || 1;
      if (w <= targetW) { resolve(dataUrl); return; } // 已足够小，跳过 canvas 重编码，省 CPU
      const c = document.createElement('canvas');
      c.width = targetW;
      c.height = Math.max(1, Math.round(h * targetW / w));
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
      try { resolve(c.toDataURL('image/jpeg', quality)); } catch (e) { reject(e); }
    };
    img.onerror = () => reject(new Error('resize load fail'));
    img.src = dataUrl;
  });
}

// 带超时的一次性 fetch：避免个别原图 URL 卡住拖慢整批导出（CORS 拦截会在超时前立即失败，不浪费时间）
function fetchWithTimeout(url, ms) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
  return fetch(url, { signal: ctrl.signal, mode: 'cors' }).finally(() => clearTimeout(t));
}

// 抓取一个附件字段单元格的全部图片：
//  - 缩略图模式（默认·最快）：直接取飞书缩略图 base64，不跨域、不需要任何代理、不逐个联网
//  - 高清原图模式：本地直连飞书附件 URL（不借助代理），失败再回退缩略图
async function fetchCellImages(fieldId, recordId, cellVal) {
  const tokens = (Array.isArray(cellVal) ? cellVal : []).filter((x) => x && x.token).map((x) => x.token);
  if (!tokens.length) return [];
  const empty = () => new Array(tokens.length).fill(null);

  const run = async () => {
    const out = empty();

    // ① 缩略图（SDK 直接返回 base64，无 CORS，最快）：受信号量限流 + 12s 超时 + 轻度重试。
    //    优先用较快的 HIGH(720)，失败再试 MAX(1280)；720 在飞书侧生成更快，能显著减少超时。
    if (state.imgQuality !== 'orig') {
      const failedIdx = out.map((v, i) => (v ? -1 : i)).filter((i) => i >= 0);
      const failedTokens = failedIdx.map((i) => tokens[i]);
      let thumbs = null;
      for (const q of [state.ImageQuality.HIGH, state.ImageQuality.MAX]) {
        try {
          const r = await thumbLimit(() => withRetry(
            async () => { await sleep(150); return state.table.getCellThumbnailUrls(failedTokens, fieldId, recordId, q); },
            { retries: 2, timeoutMs: 12000, label: 'getCellThumbnailUrls(q' + q + ')', baseDelay: 600 }
          ));
          if (r && r.length) { thumbs = r; break; }
          else log('  缩略图质量 ' + q + ' 返回空（tokens=' + failedTokens.length + '）', 'warn');
        } catch (e) {
          log('  缩略图质量 ' + q + ' 失败：' + (e && e.message ? e.message : e), 'warn');
        }
      }
      if (thumbs) {
        await Promise.all(failedIdx.map(async (idx, k) => {
          if (thumbs[k] != null) { try { out[idx] = await thumbToExcelImage(thumbs[k]); } catch (e) {} }
        }));
        if (!state._thumbLogged) {
          state._thumbLogged = true;
          const sm = thumbs[0];
          const t = typeof sm === 'string' ? sm : (sm && (sm.url || sm.thumbnail || JSON.stringify(sm))) || String(sm);
          log('缩略图返回格式样例（' + (typeof sm) + '）：' + String(t).slice(0, 60));
        }
      }
    }

    // ② 原图兜底（orig 模式主力；或缩略图失败时的兜底）：直连飞书附件 URL 本地下载 + canvas 缩放转 JPEG。
    //    飞书 CDN 偶尔 CORS 拦截导致下载失败，属正常，失败即跳过、不刷屏。
    const stillMissing = out.map((v, i) => (v ? -1 : i)).filter((i) => i >= 0);
    if (stillMissing.length) {
      let urls = [];
      try {
        urls = await attachLimit(() => withRetry(
          () => state.table.getCellAttachmentUrls(stillMissing.map((i) => tokens[i]), fieldId, recordId),
          { retries: 1, timeoutMs: 10000, label: 'getCellAttachmentUrls', baseDelay: 500 }
        ));
      } catch (e) { /* 兜底失败，进入下方空图判断 */ }
      if (urls && urls.length) {
        await Promise.all(stillMissing.map(async (idx, k) => {
          const u = urls[k];
          if (!u) return;
          try {
            const resp = await fetchWithTimeout(u, 6000);
            if (!resp.ok) throw new Error('HTTP ' + resp.status);
            out[idx] = await blobToExcelImageResized(await resp.blob(), 1200);
          } catch (e) { /* 取不到原图：留空 */ }
        }));
      }
    }

    if (!out.some(Boolean)) log('  该单元格原图与缩略图均不可取', 'err');
    return out;
  };

  // 整体超时兜底：单格取图最多 45s，超时则放弃该格图片返回空，避免一个单元格的飞书接口挂起把整批导出拖死。
  return withTimeout(run(), 45000, 'fetchCellImages').catch((e) => {
    log('  单格取图超时已跳过（' + tokens.length + ' 张）', 'warn');
    return empty();
  });
}

// ---------- 文字单元格格式化 ----------
function formatText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (Array.isArray(v)) return v.map((x) => formatText(x)).filter(Boolean).join('，');
  if (typeof v === 'object') {
    if ('text' in v) return String(v.text);
    if ('name' in v) return String(v.name);
    return JSON.stringify(v);
  }
  return String(v);
}

// ---------- 导出 Excel ----------
function buildColumnPlan(fieldIds) {
  // 返回 [{fieldId, isAttachment, imgIndex}]，附件字段按 maxAttach 预留多列
  const plan = [];
  for (const fid of fieldIds) {
    const f = state.fields.find((x) => x.id === fid);
    if (f && f.isAttachment) {
      const n = Math.max(1, state.maxAttach[fid] || 1);
      for (let i = 0; i < n; i++) plan.push({ fieldId: fid, isAttachment: true, imgIndex: i });
    } else {
      plan.push({ fieldId: fid, isAttachment: false, imgIndex: -1 });
    }
  }
  return plan;
}

async function exportExcel() {
  if (!state.loaded || !state.records.length) {
    log('请先「加载数据」。', 'warn');
    return;
  }
  const ExcelJS = window.ExcelJS;
  if (!ExcelJS) { log('ExcelJS 不可用，无法导出。', 'err'); return; }

  try {

  const fieldIds = getSelectedFieldIds();
  if (!fieldIds.length) { log('请至少选择一个字段。', 'warn'); return; }
  const plan = buildColumnPlan(fieldIds);
  const DISPLAY_W = Math.max(40, Math.min(400, parseInt($('#imgWidth').value, 10) || 50));
  // 行级并发：每个 worker 负责若干行；但「真正在途的飞书 SDK 取图调用」已由全局信号量 thumbLimit/attachLimit 硬限到 ≤3，
  // 行并发仅影响取图与写表之间的流水线，不会压垮飞书缩略图服务
  const concurrency = Math.max(1, Math.min(10, parseInt($('#concurrency').value, 10) || 2));
  state.imgMode = ($('#imgMode').value || 'float'); // float=浮动图片(兼容所有) / image=IMAGE 公式(需365)
  state.imgQuality = ($('#imgQuality').value || 'thumb'); // thumb=缩略图(最快) / orig=高清原图(直连飞书)
  state.stat = { orig: 0, thumb: 0, embedded: 0 };

  $('#btnExport').disabled = true;
  $('#btnLoad').disabled = true;
  showProgress();
  setProgress(0);
  log('开始生成 Excel…');

  const total = state.records.length;
  const attachFields = [...new Set(plan.filter((c) => c.isAttachment).map((c) => c.fieldId))];

  // ===== 阶段一：并发取图（仅 IO，绝不碰 ExcelJS），按行号存入内存 =====
  // 关键：先取齐所有图片再写表，避免多 worker 并发写 worksheet 内部状态导致图片错位
  log('阶段一：并发读取图片（' + total + ' 行）…');
  const imgData = new Array(total);
  let cursor = 0;
  async function fetchWorker() {
    while (cursor < total) {
      const idx = cursor++;
      const rec = state.records[idx];
      const cache = {};
      await Promise.all(attachFields.map(async (fid) => {
        const cell = rec.fields && rec.fields[fid];
        cache[fid] = await fetchCellImages(fid, rec.recordId, cell);
      }));
      imgData[idx] = cache;
      const done = idx + 1;
      setProgress(Math.round(done / total * 70));
      setProgressCount(done + ' / ' + total + ' 行取图');
      if (idx % 20 === 0) log('  取图 ' + done + ' / ' + total + ' 行');
    }
  }
  const fw = [];
  for (let i = 0; i < concurrency; i++) fw.push(fetchWorker());
  await Promise.all(fw);

  // ===== 阶段二：单线程顺序写 Excel（彻底消除并发写表竞态） =====
  log('阶段二：写入表格（' + total + ' 行）…');
  const wb = new ExcelJS.Workbook();
  wb.creator = '多维表图片导出';
  wb.created = new Date();
  const ws = wb.addWorksheet(state.tableName || 'Sheet1');

  // 表头
  const headers = plan.map((c) => {
    const f = state.fields.find((x) => x.id === c.fieldId);
    if (c.isAttachment && c.imgIndex > 0) return f.name + ' (' + (c.imgIndex + 1) + ')';
    return f ? f.name : c.fieldId;
  });
  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).alignment = { vertical: 'middle' };

  // 列宽：图片列确保 ≥ DISPLAY_W 像素（字符宽 ≈ 像素/7，+1 字符余量防裁剪）；文字列约 154 像素
  plan.forEach((c, i) => {
    const col = ws.getColumn(i + 1);
    col.width = c.isAttachment ? Math.max(12, Math.ceil(DISPLAY_W / 7) + 1) : 22;
  });

  // 表头行高度固定
  ws.getRow(1).height = 15;

  for (let idx = 0; idx < total; idx++) {
    const rec = state.records[idx];
    const rowNum = idx + 2;
    const row = ws.getRow(rowNum);
    let maxRowPx = 0;
    const attachCache = imgData[idx] || {};
    for (let ci = 0; ci < plan.length; ci++) {
      const c = plan[ci];
      if (c.isAttachment) {
        const imgs = attachCache[c.fieldId] || [];
        const img = imgs[c.imgIndex];
        if (img && img.base64 && img.width && img.height) {
          const mode = state.imgMode || 'float';
          // 精确像素尺寸：宽 = DISPLAY_W，高等比缩放 → 绝不变形
          const Wd = DISPLAY_W;
          const Hd = Math.max(1, Math.round(Wd * img.height / Math.max(1, img.width)));
          let imgId = null;
          try {
            imgId = wb.addImage({ base64: img.base64, extension: img.extension });
            // oneCellAnchor：tl 锚定单元格左上角，ext 为精确像素尺寸，editAs='oneCell' → 图片不随单元格拉伸变形
            ws.addImage(imgId, {
              tl: { col: ci, row: idx + 1 },
              ext: { width: Wd, height: Hd },
              editAs: 'oneCell',
            });
          } catch (e) {
            log('  第 ' + rowNum + ' 行某图嵌入失败已跳过：' + e.message, 'warn');
            imgId = null;
          }
          if (mode === 'image' && imgId) {
            // 叠加 IMAGE 公式（Excel 365/新版 WPS 可复制带走），失败不影响已贴入的浮动图
            try {
              const src = 'data:' + img.extension + ';base64,' + img.base64;
              const r = await resizeImage(src, DISPLAY_W, 0.85);
              if (r.length <= 32000) {
                row.getCell(ci + 1).value = { formula: 'IMAGE("' + r + '",1)' };
                state.stat.embedded++;
              }
            } catch (e) { /* 公式失败无影响 */ }
          }
          // 行高按图片等比缩放后的像素高度来设置
          if (Hd > maxRowPx) maxRowPx = Hd;
        }
      } else {
        const f = state.fields.find((x) => x.id === c.fieldId);
        row.getCell(ci + 1).value = formatText(rec.fields ? rec.fields[c.fieldId] : undefined);
        if (f && f.isPrimary) row.getCell(ci + 1).font = { bold: true };
      }
    }
    // 行高（pt）≈ 像素高 × 0.75（96dpi），+ 4pt 余量确保图片不被裁剪
    row.height = Math.max(18, Math.round(maxRowPx * 0.75) + 4);
    const done = idx + 1;
    setProgress(70 + Math.round(done / total * 30));
    setProgressCount(done + ' / ' + total + ' 行写入');
    if (idx % 50 === 0) log('  写入第 ' + done + ' / ' + total + ' 行');
  }

  log('正在写入文件…');
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const safe = (state.tableName || 'export').replace(/[\\/:*?"<>|]/g, '_');
  const stamp = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = safe + '_图片_' + stamp + '.xlsx';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);

  setProgress(100);
  let imgTip = '';
  if (state.imgMode === 'image' && state.stat.embedded > 0) {
    imgTip = '；图片已贴入单元格（全版本可见），并写入 IMAGE 公式：用 Excel 365 / 新版 WPS 打开时复制单元格可带走图片，旧版自动忽略公式、仍显示贴入图。';
  } else if (state.imgMode === 'image') {
    imgTip = '；图片已贴入单元格（全版本可见）；本次图片较大未写入 IMAGE 公式，复制带走需 Excel 365。';
  } else {
    imgTip = '；图片以浮动方式贴入单元格（全版本可见）；旧版 Excel 复制单元格不会带走图片（Excel 原生限制）。';
  }
  log('导出完成：' + a.download + '（图片：原图 ' + state.stat.orig + ' / 缩略图 ' + state.stat.thumb + imgTip + '）', 'ok');
  if (state.stat.orig === 0 && state.stat.thumb > 0 && state.imgQuality === 'orig') {
    log('诊断：已选「高清原图」但全部回退为缩略图。原图被飞书 CDN 的 CORS 策略拦截，前端无法取到像素。可在「图片设置 → 图片质量」改回「缩略图」（最快最稳）。', 'warn');
  }
  try {
    await markExported(state.records);
  } catch (e) {
    log('标记「已导出」失败（不影响已生成的文件）：' + (e && e.message ? e.message : e), 'warn');
  }
  setProgress(100);
  } catch (e) {
    log('导出异常中断：' + (e && e.message ? e.message : e), 'err');
    setStatus('导出失败（详见运行日志）', 'err');
  } finally {
    $('#btnExport').disabled = false;
    $('#btnLoad').disabled = false;
    hideProgress();
  }
}

// ---------- 导出后标记「已导出」 ----------
async function markExported(records) {
  const sel = $('#markField') ? $('#markField').value : '';
  if (!sel || !records || !records.length) return; // 不标记
  let fieldId = sel;

  // 选项：新建「已导出」复选框字段
  if (sel === '__create__') {
    try {
      const res = await state.table.addField({ type: FIELD_TYPE_CHECKBOX, name: '已导出' });
      fieldId = res && res.id;
      if (!fieldId) throw new Error('addField 未返回字段 id');
      const fname = (res && res.name) || '已导出';
      log('已新建「' + fname + '」复选框字段', 'ok');
      // 加入字段列表并刷新下拉（不重渲染字段选择，避免清掉用户勾选）
      state.fields.push({ id: fieldId, name: fname, type: FIELD_TYPE_CHECKBOX, isPrimary: false, isAttachment: false });
      renderMarkOptions(fieldId);
      const mf = $('#markField'); if (mf) mf.value = fieldId;
    } catch (e) {
      log('新建标记字段失败：' + e.message + '（已跳过标记）', 'err');
      return;
    }
  }

  log('开始标记已导出（共 ' + records.length + ' 行）…');
  let okCount = 0;
  const n = records.length;
  let done = 0, cursor = 0;
  const conc = 4;
  async function worker() {
    while (cursor < n) {
      const i = cursor++;
      const rec = records[i];
      try {
        await withTimeout(state.table.setCellValue(fieldId, rec.recordId, true), 6000, 'setCellValue');
        okCount++;
      } catch (e) {
        log('  标记失败（第 ' + (i + 1) + ' 行）：' + e.message, 'warn');
      }
      done++;
    }
  }
  const workers = [];
  for (let k = 0; k < conc; k++) workers.push(worker());
  await Promise.all(workers);
  log('已标记「已导出」：' + okCount + ' / ' + n + ' 行', okCount === n ? 'ok' : 'warn');
}

// 生成「导出后标记」下拉（复选框字段 + 新建选项）
function renderMarkOptions(selectedId) {
  const sel = $('#markField');
  if (!sel) return;
  sel.innerHTML = '';
  const none = document.createElement('option');
  none.value = ''; none.textContent = '不标记';
  sel.appendChild(none);
  for (const f of state.fields) {
    if (f.type !== FIELD_TYPE_CHECKBOX) continue;
    const o = document.createElement('option');
    o.value = f.id; o.textContent = f.name + '（复选框）';
    sel.appendChild(o);
  }
  const create = document.createElement('option');
  create.value = '__create__'; create.textContent = '＋ 新建「已导出」复选框字段';
  sel.appendChild(create);
  if (selectedId) sel.value = selectedId;
}

// 生成「图片命名列」下拉（所有字段均可格式化文本，默认选主键）
function renderNamingOptions() {
  const sel = $('#namingField');
  if (!sel) return;
  sel.innerHTML = '';
  let defId = '';
  for (const f of state.fields) {
    const o = document.createElement('option');
    o.value = f.id;
    o.textContent = f.name + (f.isPrimary ? '（主键）' : '');
    sel.appendChild(o);
    if (f.isPrimary) defId = f.id;
  }
  if (defId) sel.value = defId;
}

// ---------- 导出图片 ZIP（按文字列命名）----------
async function exportZip() {
  if (!state.loaded || !state.records.length) {
    log('请先「加载数据」。', 'warn');
    return;
  }
  const JSZip = window.JSZip;
  if (!JSZip) { log('JSZip 不可用，无法导出 ZIP。', 'err'); return; }

  try {

  const fieldIds = getSelectedFieldIds();
  let attachFields = state.fields.filter((f) => f.isAttachment && fieldIds.includes(f.id));
  if (!attachFields.length) attachFields = state.fields.filter((f) => f.isAttachment);
  if (!attachFields.length) { log('没有可用的图片字段。', 'warn'); return; }

  const namingId = $('#namingField').value;
  state.stat = { orig: 0, thumb: 0 };
  // 行级并发：每个 worker 负责若干行；但「真正在途的飞书 SDK 取图调用」已由全局信号量 thumbLimit/attachLimit 硬限到 ≤3，
  // 行并发仅影响取图与写表之间的流水线，不会压垮飞书缩略图服务
  const concurrency = Math.max(1, Math.min(10, parseInt($('#concurrency').value, 10) || 2));

  $('#btnExport').disabled = true;
  $('#btnExportZip').disabled = true;
  $('#btnLoad').disabled = true;
  showProgress();
  setProgress(0);
  log('开始生成图片 ZIP…');

  const total = state.records.length;

  // 阶段一：并发取图（仅 IO），按行号存入内存
  log('阶段一：并发读取图片（' + total + ' 行）…');
  const imgData = new Array(total);
  let cursor = 0;
  async function fetchWorker() {
    while (cursor < total) {
      const idx = cursor++;
      const rec = state.records[idx];
      const cache = {};
      await Promise.all(attachFields.map(async (f) => {
        const cell = rec.fields && rec.fields[f.id];
        cache[f.id] = await fetchCellImages(f.id, rec.recordId, cell);
      }));
      imgData[idx] = cache;
      const done = idx + 1;
      setProgress(Math.round(done / total * 80));
      setProgressCount(done + ' / ' + total + ' 行取图');
      if (idx % 20 === 0) log('  取图 ' + done + ' / ' + total + ' 行');
    }
  }
  const fw = [];
  for (let i = 0; i < concurrency; i++) fw.push(fetchWorker());
  await Promise.all(fw);

  // 阶段二：单线程写 ZIP
  log('阶段二：打包 ZIP…');
  const zip = new JSZip();
  const used = new Set();
  const safe = (raw) => {
    let s = (raw == null ? '' : String(raw)).trim().replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '_');
    if (s.length > 80) s = s.slice(0, 80);
    return s;
  };
  const uniq = (name) => {
    if (!used.has(name)) { used.add(name); return name; }
    let i = 2;
    while (used.has(name + '_' + i)) i++;
    const u = name + '_' + i; used.add(u); return u;
  };

  let fileCount = 0;
  for (let idx = 0; idx < total; idx++) {
    const rec = state.records[idx];
    const base = safe(formatText(rec.fields ? rec.fields[namingId] : undefined)) || rec.recordId;
    const cache = imgData[idx] || {};
    for (const f of attachFields) {
      const imgs = cache[f.id] || [];
      for (let k = 0; k < imgs.length; k++) {
        const img = imgs[k];
        if (!img) continue;
        const fname = uniq(base + '__' + safe(f.name) + '_' + (k + 1) + '.' + img.extension);
        zip.file(fname, img.base64, { base64: true });
        fileCount++;
      }
    }
    const done = idx + 1;
    setProgress(80 + Math.round(done / total * 20));
    setProgressCount(done + ' / ' + total + ' 行 · ' + fileCount + ' 张图');
    if (idx % 50 === 0) log('  打包第 ' + done + ' / ' + total + ' 行');
  }

  log('正在打包 ZIP（共 ' + fileCount + ' 张图片）…');
  const blob = await zip.generateAsync({ type: 'blob' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const safeName = (state.tableName || 'export').replace(/[\\/:*?"<>|]/g, '_');
  const stamp = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = safeName + '_图片_' + stamp + '.zip';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);

  setProgress(100);
  log('导出完成：' + a.download + '（' + fileCount + ' 张图片，原图 ' + state.stat.orig + ' / 缩略图 ' + state.stat.thumb + '）', 'ok');
  if (state.stat.orig === 0 && state.stat.thumb > 0) {
    log('诊断：本次图片均为缩略图（最长边≤1280），已是最快路径。', 'warn');
  }
  try {
    await markExported(state.records);
  } catch (e) {
    log('标记「已导出」失败（不影响已生成的文件）：' + (e && e.message ? e.message : e), 'warn');
  }
  setProgress(100);
  } catch (e) {
    log('ZIP 导出异常中断：' + (e && e.message ? e.message : e), 'err');
    setStatus('导出失败（详见运行日志）', 'err');
  } finally {
    $('#btnExport').disabled = false;
    $('#btnExportZip').disabled = false;
    $('#btnLoad').disabled = false;
    hideProgress();
  }
}

// ---------- 设置弹窗 / 日志折叠 ----------
function openSettings() { $('#settingsModal').hidden = false; }
function closeSettings() { $('#settingsModal').hidden = true; }

window.addEventListener('DOMContentLoaded', () => {
  $('#btnLoad').addEventListener('click', loadData);
  $('#btnExport').addEventListener('click', exportExcel);
  $('#btnExportZip').addEventListener('click', exportZip);
  $('#tableSelect').addEventListener('change', (e) => switchTable(e.target.value));
  $('#btnSelectAll').addEventListener('click', () => selectAllFields(true));
  $('#btnClearAll').addEventListener('click', () => selectAllFields(false));
  $('#btnSelectImg').addEventListener('click', selectImageFields);
  // 导出字段折叠下拉
  $('#fieldToggle').addEventListener('click', toggleFieldDropdown);
  $('#fieldToggle').addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleFieldDropdown(); } });
  $('#fieldList').addEventListener('change', updateFieldSummary);
  // 设置弹窗
  $('#btnSettings').addEventListener('click', openSettings);
  $('#btnCloseSettings').addEventListener('click', closeSettings);
  $('#settingsModal').addEventListener('click', (e) => { if (e.target === $('#settingsModal')) closeSettings(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeSettings(); });
  // 日志折叠
  $('#logToggle').addEventListener('click', () => {
    const panel = $('#logPanel');
    const collapsed = panel.classList.toggle('collapsed');
    const chev = $('#logToggle').querySelector('.chev');
    if (chev) chev.classList.toggle('collapsed', collapsed);
  });
  ensureExcelJS().then((ok) => {
    if (!ok) setStatus('ExcelJS 加载失败（Excel 导出将不可用）', 'err');
    init();
  });
  ensureJSZip(); // 后台预载，不阻塞
});
