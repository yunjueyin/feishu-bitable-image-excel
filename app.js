// 多维表图片导出 Excel —— 飞书多维表自定义插件（纯前端）
// 依赖：ExcelJS（UMD 全局，已在 index.html 引入，失败回退 CDN）
// SDK：@lark-base-open/js-sdk（ESM，动态 import，带 CDN 兜底）

// ---------- 常量 ----------
const FIELD_TYPE_ATTACHMENT = 17;
const FIELD_TYPE_CHECKBOX = 7;
const SUPPORTED_IMG = ['png', 'jpeg', 'gif', 'bmp'];
const ImageQualityFallback = { Low: 120, Mid: 360, HIGH: 720, MAX: 1280 };
const THUMB_QUALITY_HIGH = 2560; // 尝试高于 SDK MAX(1280) 的缩略图质量，飞书服务端若支持则返回更大图

// 飞书 SDK 1.0.2 实际并未导出 ImageQuality 枚举（已核对 dist 包），运行时 state.ImageQuality 恒为兜底常量。
// 为兼容「未来 SDK 真导出该枚举且键名大小写不同（High/MAX 等）」的情况，这里做大小写容错解析：
// 优先取 SDK 枚举（多种大小写都试），取不到则回退到已知可用的数字 720/1280。
function resolveQuality() {
  const Q = (state && state.ImageQuality) || ImageQualityFallback;
  const pick = (...keys) => { for (const k of keys) if (Q[k] != null) return Q[k]; return undefined; };
  return {
    HIGH: pick('HIGH', 'High', 'high') != null ? pick('HIGH', 'High', 'high') : ImageQualityFallback.HIGH,
    MAX: pick('MAX', 'Max', 'max') != null ? pick('MAX', 'Max', 'max') : ImageQualityFallback.MAX,
  };
}

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
  fetchStat: { ok: 0, fail: 0, fallback: 0, empty: 0 }, // 取图实时计数（交互3）；empty=空图片单元格(不计失败/不重试)
  aborted: false, // 取消标志（功能5）
  onlyUnmarked: false, // 仅导出未标记行（功能4）
  imgQuality: 'orig', // orig=高清原图(本地直连飞书·默认推荐) / thumb=缩略图(最快·最稳)
  imgCache: {},      // 会话内已取图缓存（断点续传/仅补缺失），键=质量|字段|记录|序号
  failPairs: new Set(),  // 本次导出失败图片键集合（供「仅重试失败项」）
  failRows: new Set(),   // 本次导出存在失败图片的记录 id
  exporting: false,  // 导出进行中（防重复点击）
  lastExport: null,  // 'excel' | 'zip'，记录上次导出类型（供重试复用）
  retryMode: false,  // 当前是否为「仅重试失败项」运行
  fetchBytes: 0,     // 本次导出已取图字节累计（实时体积预估）
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
// 人类可读体积（实时预估导出文件大小）
function humanSize(n) {
  n = n || 0;
  if (n >= 1048576) return (n / 1048576).toFixed(n >= 10485760 ? 0 : 1) + ' MB';
  if (n >= 1024) return Math.round(n / 1024) + ' KB';
  return n + ' B';
}
// 累计已取图字节（base64 长度 × 3/4 ≈ 解码后字节数），用于实时体积预估
function addImgBytes(img) {
  if (img && img.base64) state.fetchBytes = (state.fetchBytes || 0) + Math.ceil(img.base64.length * 3 / 4);
}
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

// ---------- 本地存储（功能8：记住选择与设置）----------
const LS = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
};
const selKey = (tid) => 'fie_sel_' + tid;
// 记住/恢复每表勾选字段
function saveSelection() {
  if (!state.tableId) return;
  LS.set(selKey(state.tableId), getSelectedFieldIds());
}
function applySelection() {
  const saved = LS.get(selKey(state.tableId), null);
  const boxes = document.querySelectorAll('#fieldList input[type=checkbox]');
  if (saved && Array.isArray(saved) && saved.length) {
    const set = new Set(saved);
    boxes.forEach((cb) => { cb.checked = set.has(cb.dataset.fieldId); });
  }
  updateFieldSummary();
}
// 记住/恢复常用导出设置（全局）
function saveSettings() {
  LS.set('fie_settings', {
    imgQuality: getSeg('imgQuality'),
    imgMode: getSeg('imgMode'),
    imgWidth: $('#imgWidth').value,
    concurrency: $('#concurrency').value,
    onlyUnmarked: $('#onlyUnmarked').checked,
    namingField: $('#namingField').value,
    markField: $('#markField').value,
  });
}
function applySettings() {
  const s = LS.get('fie_settings', {});
  if (s.imgQuality) setSeg('imgQuality', s.imgQuality);
  if (s.imgMode) setSeg('imgMode', s.imgMode);
  if (s.imgWidth) $('#imgWidth').value = s.imgWidth;
  if (s.concurrency) $('#concurrency').value = s.concurrency;
  if (typeof s.onlyUnmarked === 'boolean') $('#onlyUnmarked').checked = s.onlyUnmarked;
  if (s.namingField) $('#namingField').value = s.namingField;
  if (s.markField) $('#markField').value = s.markField;
  state.onlyUnmarked = !!s.onlyUnmarked;
}

// ---------- 有效记录集（功能4：仅导出未标记行）----------
function getEffectiveRecords() {
  const markSel = $('#markField') ? $('#markField').value : '';
  const onlyUnmarked = $('#onlyUnmarked') ? $('#onlyUnmarked').checked : false;
  let recs = state.records;
  if (onlyUnmarked && markSel && markSel !== '__create__') {
    recs = recs.filter((r) => {
      const v = r.fields ? r.fields[markSel] : undefined;
      return v !== true; // 只保留未勾选（未标记）的行
    });
  }
  return recs;
}

// 导出规模预估（交互4：大表确认）
function estimateExport() {
  const recs = getEffectiveRecords();
  const fieldIds = getSelectedFieldIds();
  const attachFields = state.fields.filter((f) => f.isAttachment && fieldIds.includes(f.id));
  let imgs = 0;
  for (const r of recs) {
    for (const f of attachFields) {
      const c = r.fields && r.fields[f.id];
      if (Array.isArray(c)) imgs += c.filter((x) => x && x.token).length;
    }
  }
  return { rows: recs.length, imgs, tooLarge: recs.length > 500 || imgs > 800 };
}

// ---------- 分块取图（功能6：降内存峰值）----------
// 仅对传入的 records 取图，返回与 records 对齐的 imgData；尊重取消标志。
async function fetchImagesForRecords(records, attachFieldIds, onProgress) {
  const out = new Array(records.length);
  let cursor = 0;
  let done = 0;
  const conc = Math.max(1, Math.min(12, parseInt($('#concurrency').value, 10) || 2));
  async function worker() {
    while (!state.aborted && cursor < records.length) {
      const i = cursor++;
      const rec = records[i];
      const cache = {};
      await Promise.all(attachFieldIds.map((fid) =>
        fetchCellImages(fid, rec.recordId, rec.fields ? rec.fields[fid] : undefined).then((imgs) => { cache[fid] = imgs; })
      ));
      out[i] = cache;
      done++;
      if (onProgress) onProgress(done, records.length);
    }
  }
  const ws = [];
  for (let i = 0; i < conc; i++) ws.push(worker());
  await Promise.all(ws);
  return out;
}

// ---------- 文件名 / 下载（复用，避免重复）----------
function makeXlsxName() {
  const safe = (state.tableName || 'export').replace(/[\\/:*?"<>|]/g, '_');
  const stamp = new Date().toISOString().slice(0, 10);
  return safe + '_图片_' + stamp + '.xlsx';
}
function makeZipName() {
  const safe = (state.tableName || 'export').replace(/[\\/:*?"<>|]/g, '_');
  const stamp = new Date().toISOString().slice(0, 10);
  return safe + '_图片_' + stamp + '.zip';
}
function triggerDownload(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// ---------- 取消 / 完成 UI（功能5）----------
function showCancel() { const b = $('#btnCancel'); if (b) b.classList.remove('hidden'); }
function hideCancel() { const b = $('#btnCancel'); if (b) b.classList.add('hidden'); }
function finishExportUI() {
  $('#btnExport').disabled = false;
  $('#btnExportZip').disabled = false;
  $('#btnLoad').disabled = false;
  hideProgress();
  hideCancel();
  state.exporting = false;
  updateRetryButton();
}

// ---------- 成功 toast（UI2 / 交互5）----------
function showToast(title, msg) {
  const t = $('#toast');
  if (!t) return;
  $('#toastTitle').textContent = title;
  $('#toastMsg').textContent = msg;
  t.classList.remove('hidden');
  // 触发过渡
  requestAnimationFrame(() => t.classList.add('show'));
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => {
    t.classList.remove('show');
    setTimeout(() => t.classList.add('hidden'), 280);
  }, 5200);
}

// ---------- 通用确认弹窗（交互1：标记副作用确认 / 交互4：大表确认）----------
function showConfirm(message, opts) {
  opts = opts || {};
  return new Promise((resolve) => {
    const modal = $('#confirmModal');
    if (!modal) { resolve(true); return; }
    $('#confirmTitle').textContent = opts.title || '确认';
    $('#confirmMsg').textContent = message;
    const okBtn = $('#confirmOk');
    const cancelBtn = $('#confirmCancel');
    okBtn.textContent = opts.okText || '确定';
    cancelBtn.textContent = opts.cancelText || '取消';
    modal.hidden = false;
    const cleanup = () => {
      modal.hidden = true;
      okBtn.removeEventListener('click', onOk);
      cancelBtn.removeEventListener('click', onCancel);
      modal.removeEventListener('click', onMask);
    };
    const onOk = () => { cleanup(); resolve(true); };
    const onCancel = () => { cleanup(); resolve(false); };
    const onMask = (e) => { if (e.target === modal) onCancel(); };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', onCancel);
    modal.addEventListener('click', onMask);
  });
}

// 标记前二次确认（交互1）。首次确认后记住，避免每次都弹。
async function ensureMarkConsent() {
  const sel = $('#markField') ? $('#markField').value : '';
  if (!sel) return true; // 不标记
  if (LS.get('fie_mark_consent', false)) return true;
  let label = '';
  const opt = $('#markField').selectedOptions && $('#markField').selectedOptions[0];
  if (opt) label = opt.textContent;
  const detail = sel === '__create__'
    ? '新建「已导出」字段并标记已导出的行'
    : '在字段「' + label + '」中标记已导出的行';
  const ok = await showConfirm('导出完成后将' + detail + '。此操作会修改你的多维表（可手动撤销）。确定继续吗？', { okText: '继续并记住', cancelText: '取消' });
  if (ok) LS.set('fie_mark_consent', true);
  return ok;
}

// ---------- 导出前预览（UI1）----------
async function loadPreviewThumbs() {
  const cont = $('#previewThumbs');
  const hint = $('#previewHint');
  if (cont) cont.innerHTML = '';
  if (hint) hint.textContent = '正在取图…';
  const attachFields = state.fields.filter((f) => f.isAttachment);
  if (!attachFields.length) { if (hint) hint.textContent = '本表没有图片字段，无可预览。'; return; }
  const rec = state.records[0];
  if (!rec) { if (hint) hint.textContent = '尚无数据，请先加载数据。'; return; }
  let loaded = 0;
  try {
    for (const f of attachFields) {
      const cell = rec.fields && rec.fields[f.id];
      const tokens = (Array.isArray(cell) ? cell : []).filter((x) => x && x.token).slice(0, 3);
      for (const tk of tokens) {
        if (loaded >= 6) break;
        try {
          const r = await thumbLimit(() => withRetry(
            async () => state.table.getCellThumbnailUrls([tk], f.id, rec.recordId, resolveQuality().MAX),
            { retries: 1, timeoutMs: 10000, label: 'preview' }
          ));
          const im = await thumbToExcelImage(r && r[0]);
          if (im && cont) {
            const img = document.createElement('img');
            img.src = 'data:' + im.extension + ';base64,' + im.base64;
            img.className = 'pv-thumb';
            img.alt = f.name;
            cont.appendChild(img);
            loaded++;
          }
        } catch (e) { /* 忽略单张失败 */ }
      }
      if (loaded >= 6) break;
    }
  } catch (e) { /* 忽略 */ }
  if (hint) hint.textContent = loaded ? ('已预览前 ' + loaded + ' 张（仅示意，实际以导出为准）') : '预览取图失败（不影响导出）。';
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

// ---------- 主题（浅色 + 暗色切换，记忆偏好） ----------
const THEME_KEY = 'fie_theme';
function applyThemeClass(theme) {
  document.documentElement.setAttribute('data-theme', theme === 'dark' ? 'dark' : 'light');
}
function initTheme() {
  let theme = LS.get(THEME_KEY, '');
  if (!theme) theme = (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
  applyThemeClass(theme);
  if (window.matchMedia) {
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
      if (!LS.get(THEME_KEY, '')) applyThemeClass(e.matches ? 'dark' : 'light');
    });
  }
}
function toggleTheme() {
  const cur = document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
  const next = cur === 'dark' ? 'light' : 'dark';
  applyThemeClass(next);
  LS.set(THEME_KEY, next);
}

// ---------- 分段控件（质量 / 嵌入方式）取值与赋值 ----------
function getSeg(id) {
  const el = document.getElementById(id);
  if (!el) return '';
  const active = el.querySelector('.seg-btn[aria-pressed="true"]');
  return active ? active.dataset.val : '';
}
function setSeg(id, val) {
  const el = document.getElementById(id);
  if (!el) return;
  el.querySelectorAll('.seg-btn').forEach((b) => {
    b.setAttribute('aria-pressed', String(b.dataset.val === val));
  });
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
    applySettings(); // 恢复上次导出的常用设置（功能8）
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
  applySelection(); // 恢复本表上次勾选（功能8）
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
  // 空状态引导：未勾选任何列时给出友好提示（交互：空状态引导）
  const eh = $('#fieldEmptyHint');
  if (eh) eh.classList.toggle('hidden', checked !== 0);
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
    log('已切换数据表：' + state.tableName + '，正在自动加载数据…', 'ok');
    await loadData(); // 交互2：切表后自动加载，免去手动再点
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
  state.aborted = false; // 复位取消标志
  $('#btnLoad').disabled = true;
  $('#btnExport').disabled = true;
  state.loaded = false;
  state.records = [];
  state.maxAttach = {};
  showProgress();
  setProgress(0);
  { const deh = $('#dataEmptyHint'); if (deh) deh.classList.add('hidden'); } // 空状态引导：开始时先隐藏
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
    // 预览卡：更新预估并启用预览按钮（UI1 / 交互4）
    const est = estimateExport();
    const ps = $('#previewSummary');
    if (ps) ps.textContent = '预计 ' + est.rows + ' 行 · 约 ' + est.imgs + ' 张图片' + (est.tooLarge ? '（较大，导出前会确认）' : '');
    const pb = $('#btnLoadPreview');
    if (pb) pb.disabled = false;
    setProgress(100);
    log('读取完成，共 ' + all.length + ' 行。', 'ok');
    // 空状态引导：0 行时给出提示
    { const deh = $('#dataEmptyHint'); if (deh) deh.classList.toggle('hidden', all.length !== 0); }
    // 自动应用默认导出方案（若存在）
    applyDefaultScheme();
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
// 缩略图最长边上限：飞书 getCellThumbnailUrls 在部分版本/环境下会忽略质量参数、直接返回原图，
// 而代码此前把取回的图原样嵌入，导致「缩略图」模式实际导出的是高清原图。
// 这里在嵌入前兜底缩放到 THUMB_CAP，保证导出图确定 ≤ 阈值（对齐请求的 High=720）。
const THUMB_CAP = 720;
let _thumbCappedLogged = false;
async function resizeToMaxSide(dataUrl, maxSide, quality) {
  const dims = await imageSize(dataUrl);
  if (!dims) return dataUrl;
  const cur = Math.max(dims.w, dims.h);
  if (cur <= maxSide) return dataUrl; // 已在阈值内，免重编码省 CPU
  const scale = maxSide / cur;
  const w = Math.max(1, Math.round(dims.w * scale));
  const h = Math.max(1, Math.round(dims.h * scale));
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const c = document.createElement('canvas');
      c.width = w; c.height = h;
      c.getContext('2d').drawImage(img, 0, 0, w, h);
      try { resolve(c.toDataURL('image/jpeg', quality)); } catch (e) { reject(e); }
    };
    img.onerror = () => reject(new Error('resize fail'));
    img.src = dataUrl;
  });
}
async function capThumbnail(obj) {
  if (!obj || !obj.base64 || !obj.width || !obj.height) return obj;
  const maxSide = Math.max(obj.width, obj.height);
  if (maxSide <= THUMB_CAP) return obj; // 已在阈值内，原样返回
  if (!_thumbCappedLogged) {
    _thumbCappedLogged = true;
    log('检测到缩略图超过 ' + THUMB_CAP + 'px（实际约 ' + maxSide + 'px），已自动缩放至 ' + THUMB_CAP + 'px 兜底。', 'warn');
  }
  const dataUrl = 'data:' + obj.extension + ';base64,' + obj.base64;
  const r = await resizeToMaxSide(dataUrl, THUMB_CAP, 0.85);
  return finalizeExcelImage(r, null);
}

// 原图超大封顶：高清原图模式保留原始分辨率，但超过 ORIG_CAP 的超大图（如扫描件/全景图）
// 直接嵌 Excel 会占用巨量内存甚至卡死，这里在嵌入前兜底缩放到 ORIG_CAP 最长边（jpeg 0.9）。
const ORIG_CAP = 4096;
let _origCappedLogged = false;
async function capMaxSide(obj, cap) {
  if (!obj || !obj.base64 || !obj.width || !obj.height) return obj;
  const maxSide = Math.max(obj.width, obj.height);
  if (maxSide <= cap) return obj; // 已在阈值内，原样返回省 CPU
  if (!_origCappedLogged) {
    _origCappedLogged = true;
    log('检测到原图超过 ' + cap + 'px（实际约 ' + maxSide + 'px），已自动封顶至 ' + cap + 'px 防内存爆。', 'warn');
  }
  const dataUrl = 'data:' + obj.extension + ';base64,' + obj.base64;
  const r = await resizeToMaxSide(dataUrl, cap, 0.9);
  return finalizeExcelImage(r, null);
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
      return capThumbnail(r);
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
  return capThumbnail(await finalizeExcelImage(dataUrl, null));
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
  const rawItems = Array.isArray(cellVal) ? cellVal : [];
  const tokens = rawItems.filter((x) => x && x.token).map((x) => x.token);
  if (!tokens.length) {
    // 空单元格：无有效图片 token（字段缺失 / 空数组 / 占位无图）。直接返回——不取图、不计失败、不进「仅重试失败项」。
    if (state.fetchStat) state.fetchStat.empty += 1;
    return [];
  }
  const empty = () => new Array(tokens.length).fill(null);
  const out = empty();
  const q = state.imgQuality;

  // 会话缓存（断点续传 / 仅补缺失）：键含 质量|字段|记录|序号，切换质量即失效
  for (let i = 0; i < tokens.length; i++) {
    const k = q + '|' + fieldId + '|' + recordId + '|' + i;
    if (state.imgCache[k]) out[i] = state.imgCache[k];
  }
  const need = [];
  for (let i = 0; i < tokens.length; i++) if (!out[i]) need.push(i);
  if (!need.length) return out; // 全部命中缓存，直接返回（不计入 fetchStat，省一次空跑）

  const run = async () => {
    const beforeCount = out.filter(Boolean).length;

    // ① 缩略图（仅非 orig 模式）：SDK 直接返回 base64，无 CORS，最快。受信号量限流 + 12s 超时 + 轻度重试。
    if (q !== 'orig') {
      const needTokens = need.map((i) => tokens[i]);
      let thumbs = null;
      const Q = resolveQuality();
      for (const qq of [Q.HIGH, Q.MAX]) {
        try {
          const r = await thumbLimit(() => withRetry(
            async () => state.table.getCellThumbnailUrls(needTokens, fieldId, recordId, qq),
            { retries: 2, timeoutMs: 12000, label: 'getCellThumbnailUrls(q' + qq + ')', baseDelay: 300 }
          ));
          if (r && r.length) { thumbs = r; break; }
          else log('  缩略图质量 ' + qq + ' 返回空（tokens=' + needTokens.length + '）', 'warn');
        } catch (e) {
          log('  缩略图质量 ' + qq + ' 失败：' + (e && e.message ? e.message : e), 'warn');
        }
      }
      if (thumbs) {
        await Promise.all(need.map(async (idx, k) => {
          if (thumbs[k] != null) {
            try {
              const im = await thumbToExcelImage(thumbs[k]);
              if (im) { out[idx] = im; addImgBytes(im); }
            } catch (e) {}
          }
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
    const stillMissing = [];
    for (let i = 0; i < tokens.length; i++) if (!out[i]) stillMissing.push(i);
    if (stillMissing.length) {
      let urls = [];
      try {
        const missTokens = stillMissing.map((i) => tokens[i]);
        urls = await attachLimit(() => withRetry(
          () => state.table.getCellAttachmentUrls(missTokens, fieldId, recordId),
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
            let im;
            if (q === 'orig') {
              // 高清原图模式：保留原始分辨率，仅对超过 ORIG_CAP(4096) 的超大图封顶防内存爆；
              // 仅转换 ExcelJS 不支持的格式。
              im = await capMaxSide(await blobToExcelImage(await resp.blob()), ORIG_CAP);
            } else {
              // 缩略图模式的兜底：原图取不到时缩放至 1200px 安全网，避免空图（体积可控）。
              im = await blobToExcelImageResized(await resp.blob(), 1200);
            }
            if (im) { out[idx] = im; addImgBytes(im); }
          } catch (e) { /* 取不到原图：留空 */ }
        }));
      }
    }

    // 写入会话缓存（供断点续传 / 仅补缺失 / 仅重试失败项复用）
    for (let i = 0; i < tokens.length; i++) {
      if (out[i]) {
        const k = q + '|' + fieldId + '|' + recordId + '|' + i;
        state.imgCache[k] = out[i];
      }
    }

    // 取图实时计数（交互3）：成功/缺失/原图回退
    if (state.fetchStat) {
      const got = out.filter(Boolean).length;
      const filled = Math.max(0, got - beforeCount);
      state.fetchStat.ok += got;
      state.fetchStat.fail += (out.length - got);
      if (filled > 0) state.fetchStat.fallback += filled;
    }

    // 记录失败项（供「仅重试失败项」入口）
    if (!state.aborted) {
      for (let i = 0; i < tokens.length; i++) {
        if (!out[i]) {
          const k = q + '|' + fieldId + '|' + recordId + '|' + i;
          state.failPairs.add(k);
          state.failRows.add(recordId);
        }
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

async function exportExcel(options) {
  options = options || {};
  if (state.exporting) { log('已有导出任务进行中，已忽略重复点击。', 'warn'); return; }
  if (!state.loaded || !state.records.length) {
    log('请先「加载数据」。', 'warn');
    return;
  }
  const ExcelJS = window.ExcelJS;
  if (!ExcelJS) { log('ExcelJS 不可用，无法导出。', 'err'); return; }

  // 立即禁用导出/加载按钮，防止重复触发（交互：导出中禁用按钮防重复点击）
  state.exporting = true;
  $('#btnExport').disabled = true;
  $('#btnExportZip').disabled = true;
  $('#btnLoad').disabled = true;

  // 大表预估确认（交互4）——重试时跳过
  const est = estimateExport();
  if (est.tooLarge && !options.skipConfirm) {
    const ok = await showConfirm('本次导出约 ' + est.rows + ' 行、' + est.imgs + ' 张图片，文件可能较大且耗时较久。确定继续吗？', { okText: '继续导出', cancelText: '取消' });
    if (!ok) { log('已取消导出（大表确认）。', 'warn'); return finishExportUI(); }
  }
  // 导出后标记副作用确认（交互1）——重试时跳过
  if (!options.skipConfirm && !(await ensureMarkConsent())) { log('已取消导出（标记确认）。', 'warn'); return finishExportUI(); }

  try {
    const fieldIds = getSelectedFieldIds();
    if (!fieldIds.length) { log('请至少选择一个字段。', 'warn'); return finishExportUI(); }
    const plan = buildColumnPlan(fieldIds);
    const DISPLAY_W = Math.max(40, Math.min(400, parseInt($('#imgWidth').value, 10) || 50));
    state.imgMode = (getSeg('imgMode') || 'float'); // float=浮动图片(兼容所有) / image=IMAGE 公式(需365)
    state.imgQuality = (getSeg('imgQuality') || 'orig'); // thumb=缩略图(最快) / orig=高清原图(直连飞书)
    state.stat = { orig: 0, thumb: 0, embedded: 0 };
    state.fetchStat = { ok: 0, fail: 0, fallback: 0, empty: 0 }; // 交互3；empty=空单元格
    state.aborted = false; // 功能5
    state.lastExport = 'excel';
    state.retryMode = !!options.skipConfirm;
    state.fetchBytes = 0;
    state.failPairs = new Set();
    state.failRows = new Set();

    showCancel(); showProgress(); setProgress(0);
    log('开始生成 Excel…');

    const recs = getEffectiveRecords(); // 功能4：仅导出未标记行
    const total = recs.length;
    if (!total) { log('没有可导出的记录（所选「仅导出未标记行」下已全部标记）。', 'warn'); finishExportUI(); return; }

    const attachFields = [...new Set(plan.filter((c) => c.isAttachment).map((c) => c.fieldId))];

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
    ws.getRow(1).height = 15;

    // 单行写入（内部调用 ExcelJS，保持单线程避免竞态）
    const writeRow = (idx, rec, attachCache) => {
      const rowNum = idx + 2;
      const row = ws.getRow(rowNum);
      let maxRowPx = 0;
      for (let ci = 0; ci < plan.length; ci++) {
        const c = plan[ci];
        if (c.isAttachment) {
          const imgs = attachCache[c.fieldId] || [];
          const img = imgs[c.imgIndex];
          if (img && img.base64 && img.width && img.height) {
            const Wd = DISPLAY_W;
            const Hd = Math.max(1, Math.round(Wd * img.height / Math.max(1, img.width)));
            try {
              const imgId = wb.addImage({ base64: img.base64, extension: img.extension });
              ws.addImage(imgId, { tl: { col: ci, row: idx + 1 }, ext: { width: Wd, height: Hd }, editAs: 'oneCell' });
              if (state.imgMode === 'image') {
                try {
                  const src = 'data:' + img.extension + ';base64,' + img.base64;
                  const r = resizeImage(src, DISPLAY_W, 0.85);
                  Promise.resolve(r).then((rd) => { if (rd && rd.length <= 32000) { row.getCell(ci + 1).value = { formula: 'IMAGE("' + rd + '",1)' }; state.stat.embedded++; } });
                } catch (e) { /* 公式失败无影响 */ }
              }
            } catch (e) {
              log('  第 ' + rowNum + ' 行某图嵌入失败已跳过：' + e.message, 'warn');
            }
            if (Hd > maxRowPx) maxRowPx = Hd;
          }
        } else {
          const f = state.fields.find((x) => x.id === c.fieldId);
          row.getCell(ci + 1).value = formatText(rec.fields ? rec.fields[c.fieldId] : undefined);
          if (f && f.isPrimary) row.getCell(ci + 1).font = { bold: true };
        }
      }
      row.height = Math.max(18, Math.round(maxRowPx * 0.75) + 4);
    };

    // 仅勾选图片列时，图片全空的数据行无内容可导出，直接跳过（紧凑写行，不占 Excel 物理行）
    const onlyImgCols = attachFields.length > 0 && attachFields.length === plan.length;
    let outRow = 0;            // 实际写入 Excel 的紧凑行号（0 基；表头为第 1 行）
    let skippedEmptyRows = 0;

    // 功能6：分块取图→写表→释放，避免全量 base64 驻留内存导致 OOM
    const CHUNK = 50;
    let processed = 0;
    for (let start = 0; start < total; start += CHUNK) {
      if (state.aborted) break; // 功能5
      const end = Math.min(start + CHUNK, total);
      const chunk = recs.slice(start, end);
      const imgData = attachFields.length
        ? await fetchImagesForRecords(chunk, attachFields, (d) => {
            setProgress(Math.round((processed + d) / total * 70));
            setProgressCount((processed + d) + ' / ' + total + ' 行取图 · 成' + state.fetchStat.ok + ' 回退' + state.fetchStat.fallback + ' 空' + state.fetchStat.empty + (state.fetchBytes ? ' · ' + humanSize(state.fetchBytes) : ''));
          })
        : chunk.map(() => ({}));
      for (let k = 0; k < chunk.length; k++) {
        const cache = imgData[k] || {};
        // 仅选图片列且该行所有图片列均空 → 跳过（无内容可导出，且 Excel 无其他列数据）
        const allImgEmpty = attachFields.length > 0 && attachFields.every((fid) => ((cache[fid] || []).length === 0));
        if (onlyImgCols && allImgEmpty) { skippedEmptyRows++; continue; }
        writeRow(outRow, chunk[k], cache);
        outRow++;
      }
      processed += chunk.length;
      setProgress(70 + Math.round(processed / total * 30));
      setProgressCount(processed + ' / ' + total + ' 行写入 · 已写 ' + outRow + ' 行' + (state.fetchBytes ? ' · 已取图 ' + humanSize(state.fetchBytes) : ''));
      if (processed % 200 === 0) log('  写入第 ' + processed + ' / ' + total + ' 行');
      if (state.aborted) break;
    }

    if (state.aborted) { log('已取消导出。', 'warn'); finishExportUI(); return; }

    log('正在写入文件…');
    const buf = await wb.xlsx.writeBuffer();
    const name = makeXlsxName();
    triggerDownload(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
    setProgress(100);
    let imgTip = '';
    if (state.imgMode === 'image' && state.stat.embedded > 0) {
      imgTip = '；图片已贴入单元格（全版本可见），并写入 IMAGE 公式：用 Excel 365 / 新版 WPS 打开时复制单元格可带走图片，旧版自动忽略公式、仍显示贴入图。';
    } else if (state.imgMode === 'image') {
      imgTip = '；图片已贴入单元格（全版本可见）；本次图片较大未写入 IMAGE 公式，复制带走需 Excel 365。';
    } else {
      imgTip = '；图片以浮动方式贴入单元格（全版本可见）；旧版 Excel 复制单元格不会带走图片（Excel 原生限制）。';
    }
    const fileSize = buf.byteLength;
    log('导出完成：' + name + '（图片：原图 ' + state.stat.orig + ' / 缩略图 ' + state.stat.thumb + ' · 文件 ' + humanSize(fileSize) + imgTip + '）', 'ok');
    if (skippedEmptyRows > 0) log('已跳过 ' + skippedEmptyRows + ' 行（仅选图片列且图片全空，无内容可导出）', 'warn');
    if (state.fetchStat.empty > 0) log('空图片单元格 ' + state.fetchStat.empty + ' 个已识别（不计失败、不重试）', 'info');
    if (state.stat.orig === 0 && state.stat.thumb > 0 && state.imgQuality === 'orig') {
      log('诊断：已选「高清原图」但全部回退为缩略图。原图被飞书 CDN 的 CORS 策略拦截，前端无法取到像素。可在「图片设置 → 图片质量」改回「缩略图」（最快最稳）。', 'warn');
    }
    showToast('Excel 导出完成', name + '（' + total + ' 行 · ' + humanSize(fileSize) + ' · 图片 原图' + state.stat.orig + '/缩略图' + state.stat.thumb + '）');
    if (!state.retryMode) {
      try {
        await markExported(recs); // 功能4：只标记本次导出的有效行
      } catch (e) {
        log('标记「已导出」失败（不影响已生成的文件）：' + (e && e.message ? e.message : e), 'warn');
      }
    }
    setProgress(100);
  } catch (e) {
    log('导出异常中断：' + (e && e.message ? e.message : e), 'err');
    setStatus('导出失败（详见运行日志）', 'err');
  } finally {
    finishExportUI();
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
async function exportZip(options) {
  options = options || {};
  if (state.exporting) { log('已有导出任务进行中，已忽略重复点击。', 'warn'); return; }
  if (!state.loaded || !state.records.length) {
    log('请先「加载数据」。', 'warn');
    return;
  }
  const JSZip = window.JSZip;
  if (!JSZip) { log('JSZip 不可用，无法导出 ZIP。', 'err'); return; }

  // 立即禁用导出/加载按钮，防止重复触发（交互：导出中禁用按钮防重复点击）
  state.exporting = true;
  $('#btnExport').disabled = true;
  $('#btnExportZip').disabled = true;
  $('#btnLoad').disabled = true;

  // 大表预估确认（交互4）——重试时跳过
  const est = estimateExport();
  if (est.tooLarge && !options.skipConfirm) {
    const ok = await showConfirm('本次导出约 ' + est.rows + ' 行、' + est.imgs + ' 张图片，文件可能较大且耗时较久。确定继续吗？', { okText: '继续导出', cancelText: '取消' });
    if (!ok) { log('已取消导出（大表确认）。', 'warn'); return finishExportUI(); }
  }
  // 导出后标记副作用确认（交互1）——重试时跳过
  if (!options.skipConfirm && !(await ensureMarkConsent())) { log('已取消导出（标记确认）。', 'warn'); return finishExportUI(); }

  try {
    const fieldIds = getSelectedFieldIds();
    let attachFields = state.fields.filter((f) => f.isAttachment && fieldIds.includes(f.id));
    if (!attachFields.length) attachFields = state.fields.filter((f) => f.isAttachment);
    if (!attachFields.length) { log('没有可用的图片字段。', 'warn'); return finishExportUI(); }

    const namingId = $('#namingField').value;
    state.stat = { orig: 0, thumb: 0 };
    state.fetchStat = { ok: 0, fail: 0, fallback: 0, empty: 0 }; // 交互3；empty=空单元格
    state.aborted = false; // 功能5
    state.lastExport = 'zip';
    state.retryMode = !!options.skipConfirm;
    state.fetchBytes = 0;
    state.failPairs = new Set();
    state.failRows = new Set();
    state.imgQuality = (getSeg('imgQuality') || 'orig'); // thumb=缩略图(最快) / orig=高清原图(直连飞书)

    showCancel(); showProgress(); setProgress(0);
    log('开始生成图片 ZIP…');

    const recs = getEffectiveRecords(); // 功能4：仅导出未标记行
    const total = recs.length;
    if (!total) { log('没有可导出的记录（所选「仅导出未标记行」下已全部标记）。', 'warn'); finishExportUI(); return; }

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

    // ZIP 仅导出图片：图片全空的数据行无文件可写，直接跳过
    let skippedEmptyRows = 0;
    // 功能6：分块取图→写 ZIP→释放，降低内存峰值
    const CHUNK = 50;
    let processed = 0;
    let fileCount = 0;
    for (let start = 0; start < total; start += CHUNK) {
      if (state.aborted) break; // 功能5
      const end = Math.min(start + CHUNK, total);
      const chunk = recs.slice(start, end);
      const imgData = attachFields.length
        ? await fetchImagesForRecords(chunk, attachFields.map((f) => f.id), (d) => {
            setProgress(Math.round((processed + d) / total * 80));
            setProgressCount((processed + d) + ' / ' + total + ' 行取图 · 成' + state.fetchStat.ok + ' 回退' + state.fetchStat.fallback + ' 空' + state.fetchStat.empty + (state.fetchBytes ? ' · ' + humanSize(state.fetchBytes) : ''));
          })
        : chunk.map(() => ({}));
      for (let k = 0; k < chunk.length; k++) {
        const cache = imgData[k] || {};
        const allImgEmpty = attachFields.length > 0 && attachFields.every((fid) => ((cache[fid] || []).length === 0));
        if (allImgEmpty) { skippedEmptyRows++; continue; } // ZIP 仅图片：全空行无内容，跳过
        const rec = chunk[k];
        const base = safe(formatText(rec.fields ? rec.fields[namingId] : undefined)) || rec.recordId;
        for (const f of attachFields) {
          const imgs = cache[f.id] || [];
          for (let m = 0; m < imgs.length; m++) {
            const img = imgs[m];
            if (!img) continue;
            const fname = uniq(base + '__' + safe(f.name) + '_' + (m + 1) + '.' + img.extension);
            zip.file(fname, img.base64, { base64: true });
            fileCount++;
          }
        }
      }
      processed += chunk.length;
      setProgress(80 + Math.round(processed / total * 20));
      setProgressCount(processed + ' / ' + total + ' 行 · ' + fileCount + ' 张图' + (state.fetchBytes ? ' · ' + humanSize(state.fetchBytes) : ''));
      if (state.aborted) break;
    }

    if (state.aborted) { log('已取消导出。', 'warn'); finishExportUI(); return; }

    log('正在打包 ZIP（共 ' + fileCount + ' 张图片）…');
    // 图片本身已压缩，STORE（不重压缩）显著快于默认 DEFLATE，且体积几乎不变（打包提速）
    const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
    const name = makeZipName();
    triggerDownload(blob, name);
    setProgress(100);
    const fileSize = blob.size;
    log('导出完成：' + name + '（' + fileCount + ' 张图片，原图 ' + state.stat.orig + ' / 缩略图 ' + state.stat.thumb + ' · 文件 ' + humanSize(fileSize) + '）', 'ok');
    if (skippedEmptyRows > 0) log('已跳过 ' + skippedEmptyRows + ' 行（图片全空，ZIP 无内容可写）', 'warn');
    if (state.fetchStat.empty > 0) log('空图片单元格 ' + state.fetchStat.empty + ' 个已识别（不计失败、不重试）', 'info');
    if (state.stat.orig === 0 && state.stat.thumb > 0) {
      log('诊断：本次图片均为缩略图（最长边≤1280），已是最快路径。', 'warn');
    }
    showToast('ZIP 导出完成', name + '（' + fileCount + ' 张图片 · ' + humanSize(fileSize) + '）');
    if (!state.retryMode) {
      try {
        await markExported(recs); // 功能4：只标记本次导出的有效行
      } catch (e) {
        log('标记「已导出」失败（不影响已生成的文件）：' + (e && e.message ? e.message : e), 'warn');
      }
    }
    setProgress(100);
  } catch (e) {
    log('ZIP 导出异常中断：' + (e && e.message ? e.message : e), 'err');
    setStatus('导出失败（详见运行日志）', 'err');
  } finally {
    finishExportUI();
  }
}

// ---------- 导出方案（多套命名预设）----------
const SCHEME_KEY = 'fie_schemes';
function getSchemes() {
  const s = LS.get(SCHEME_KEY, { default: null, items: {} });
  if (!s || typeof s !== 'object' || !s.items) return { default: null, items: {} };
  if (!s.items) s.items = {};
  return s;
}
function setSchemes(s) { LS.set(SCHEME_KEY, s); }
function readSettingsFromUI() {
  return {
    imgQuality: getSeg('imgQuality'),
    imgMode: getSeg('imgMode'),
    imgWidth: $('#imgWidth').value,
    concurrency: $('#concurrency').value,
    onlyUnmarked: $('#onlyUnmarked').checked,
    ignoreView: $('#ignoreView').checked,
    namingField: $('#namingField').value,
    markField: $('#markField').value,
  };
}
function applySettingsToUI(s) {
  if (!s) return;
  if (s.imgQuality) setSeg('imgQuality', s.imgQuality);
  if (s.imgMode) setSeg('imgMode', s.imgMode);
  if (s.imgWidth) $('#imgWidth').value = s.imgWidth;
  if (s.concurrency) $('#concurrency').value = s.concurrency;
  if (typeof s.onlyUnmarked === 'boolean') $('#onlyUnmarked').checked = s.onlyUnmarked;
  if (typeof s.ignoreView === 'boolean') $('#ignoreView').checked = s.ignoreView;
  if (s.namingField) $('#namingField').value = s.namingField;
  if (s.markField) $('#markField').value = s.markField;
  state.onlyUnmarked = !!s.onlyUnmarked;
}
function renderSchemeList(selected) {
  const sel = $('#schemeSelect');
  if (!sel) return;
  const schemes = getSchemes();
  sel.innerHTML = '';
  const names = Object.keys(schemes.items || {});
  if (!names.length) {
    const o = document.createElement('option');
    o.value = ''; o.textContent = '（暂无方案，可在上方保存）';
    sel.appendChild(o);
    sel.disabled = true;
    return;
  }
  sel.disabled = false;
  for (const n of names) {
    const o = document.createElement('option');
    o.value = n;
    o.textContent = n + (schemes.default === n ? '（默认）' : '');
    sel.appendChild(o);
  }
  if (selected && schemes.items[selected]) sel.value = selected;
}
function saveCurrentScheme() {
  const name = ($('#schemeName').value || '').trim();
  if (!name) { showToast('请先输入方案名', '方案名不能为空'); return; }
  const schemes = getSchemes();
  const cur = schemes.items[name] || { settings: {}, fields: {} };
  cur.settings = readSettingsFromUI();
  cur.fields = Object.assign({}, cur.fields);
  cur.fields[state.tableId] = getSelectedFieldIds();
  schemes.items[name] = cur;
  setSchemes(schemes);
  renderSchemeList(name);
  log('已保存导出方案：' + name + '（' + cur.fields[state.tableId].length + ' 个字段）', 'ok');
  showToast('方案已保存', name);
}
function applyScheme(name) {
  const schemes = getSchemes();
  const item = schemes.items[name];
  if (!item) return;
  applySettingsToUI(item.settings);
  saveSettings();
  if (item.fields && item.fields[state.tableId]) {
    LS.set(selKey(state.tableId), item.fields[state.tableId]);
    applySelection();
  }
  log('已应用导出方案：' + name, 'ok');
  showToast('方案已应用', name);
}
function deleteScheme(name) {
  if (!name) return;
  const schemes = getSchemes();
  if (!schemes.items[name]) return;
  delete schemes.items[name];
  if (schemes.default === name) schemes.default = null;
  setSchemes(schemes);
  renderSchemeList();
  log('已删除导出方案：' + name);
}
function setDefaultScheme(name) {
  if (!name) return;
  const schemes = getSchemes();
  if (!schemes.items[name]) return;
  schemes.default = name;
  setSchemes(schemes);
  renderSchemeList(name);
  log('已设默认导出方案：' + name, 'ok');
  showToast('已设默认方案', name);
}
// 加载数据时自动应用默认方案（若存在）
function applyDefaultScheme() {
  const schemes = getSchemes();
  if (!schemes.default || !schemes.items[schemes.default]) return;
  const item = schemes.items[schemes.default];
  applySettingsToUI(item.settings);
  saveSettings();
  if (item.fields && item.fields[state.tableId]) {
    LS.set(selKey(state.tableId), item.fields[state.tableId]);
    applySelection();
  }
}
// 「仅重试失败项」入口：复用会话缓存，仅重取失败的图片，跳过确认弹窗
function retryFailed() {
  if (state.exporting) { log('导出进行中，请稍候。', 'warn'); return; }
  if (!state.failPairs || state.failPairs.size === 0) { showToast('没有失败项', '当前没有需要重试的失败项'); return; }
  log('开始重试失败项（复用已取图，仅重取失败图片）…', 'ok');
  if (state.lastExport === 'zip') exportZip({ skipConfirm: true });
  else exportExcel({ skipConfirm: true });
}
// 根据失败项刷新「仅重试失败项」按钮
function updateRetryButton() {
  const b = $('#btnRetryFailed');
  if (!b) return;
  const n = state.failPairs ? state.failPairs.size : 0;
  if (n > 0) {
    b.classList.remove('hidden');
    b.textContent = '仅重试失败项（' + n + '）';
  } else {
    b.classList.add('hidden');
  }
}

// ---------- 设置弹窗 / 日志折叠 ----------
function openSettings() { $('#settingsModal').hidden = false; renderSchemeList(); }
function closeSettings() { $('#settingsModal').hidden = true; }

window.addEventListener('DOMContentLoaded', () => {
  $('#btnLoad').addEventListener('click', loadData);
  $('#btnExport').addEventListener('click', exportExcel);
  $('#btnExportZip').addEventListener('click', exportZip);
  $('#tableSelect').addEventListener('change', (e) => switchTable(e.target.value));
  // 字段选择变化即记忆（功能8）
  $('#btnSelectAll').addEventListener('click', () => { selectAllFields(true); saveSelection(); });
  $('#btnClearAll').addEventListener('click', () => { selectAllFields(false); saveSelection(); });
  $('#btnSelectImg').addEventListener('click', () => { selectImageFields(); saveSelection(); });
  // 导出字段折叠下拉
  $('#fieldToggle').addEventListener('click', toggleFieldDropdown);
  $('#fieldToggle').addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleFieldDropdown(); } });
  $('#fieldList').addEventListener('change', () => { updateFieldSummary(); saveSelection(); });
  // 主题切换
  initTheme();
  $('#btnTheme').addEventListener('click', toggleTheme);
  // 分段控件点击即记忆（功能8）
  ['imgQuality', 'imgMode'].forEach((id) => {
    const seg = document.getElementById(id);
    if (!seg) return;
    seg.querySelectorAll('.seg-btn').forEach((btn) => {
      btn.addEventListener('click', () => { setSeg(id, btn.dataset.val); saveSettings(); });
    });
  });
  // 常用设置变化即记忆（功能8）
  ['imgWidth', 'concurrency', 'onlyUnmarked', 'namingField', 'markField'].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', saveSettings);
  });
  // 取消导出（功能5）
  $('#btnCancel').addEventListener('click', () => { state.aborted = true; log('正在取消，请稍候…', 'warn'); });
  // 导出前预览（UI1）
  $('#previewToggle').addEventListener('click', () => {
    const panel = $('#previewPanel');
    const collapsed = panel.classList.toggle('collapsed');
    const chev = $('#previewToggle').querySelector('.chev');
    if (chev) chev.classList.toggle('collapsed', collapsed);
  });
  $('#btnLoadPreview').addEventListener('click', loadPreviewThumbs);
  // 设置弹窗
  $('#btnSettings').addEventListener('click', openSettings);
  $('#btnCloseSettings').addEventListener('click', closeSettings);
  $('#settingsModal').addEventListener('click', (e) => { if (e.target === $('#settingsModal')) closeSettings(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeSettings(); });
  // 导出方案（多套命名预设）
  $('#btnSchemeSave').addEventListener('click', saveCurrentScheme);
  $('#btnSchemeApply').addEventListener('click', () => { const n = $('#schemeSelect').value; if (n) applyScheme(n); else showToast('请选择方案', '先在上方下拉选择要应用的方案'); });
  $('#btnSchemeDelete').addEventListener('click', () => { const n = $('#schemeSelect').value; if (n) deleteScheme(n); });
  $('#btnSchemeDefault').addEventListener('click', () => { const n = $('#schemeSelect').value; if (n) setDefaultScheme(n); else showToast('请选择方案', '先在下拉选择要设为默认的方案'); });
  // 仅重试失败项
  $('#btnRetryFailed').addEventListener('click', retryFailed);
  // 日志折叠
  $('#logToggle').addEventListener('click', () => {
    const panel = $('#logPanel');
    const collapsed = panel.classList.toggle('collapsed');
    const chev = $('#logToggle').querySelector('.chev');
    if (chev) chev.classList.toggle('collapsed', collapsed);
  });
  // 速度优化：初始化与 ExcelJS 加载并行，先尽快进入可用状态
  init();
  ensureExcelJS().then((ok) => {
    if (!ok) setStatus('ExcelJS 加载失败（Excel 导出将不可用）', 'err');
  });
  ensureJSZip(); // 后台预载，不阻塞
});
