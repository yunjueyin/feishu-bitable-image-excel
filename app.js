// 多维表图片导出 Excel —— 飞书多维表自定义插件（纯前端）
// 依赖：ExcelJS（UMD 全局，已在 index.html 引入，失败回退 CDN）
// SDK：@lark-base-open/js-sdk（ESM，动态 import，带 CDN 兜底）

// ---------- 常量 ----------
const FIELD_TYPE_ATTACHMENT = 17;
const FIELD_TYPE_CHECKBOX = 7;
const SUPPORTED_IMG = ['png', 'jpeg', 'gif', 'bmp'];
const ImageQualityFallback = { Low: 120, Mid: 360, HIGH: 720, MAX: 1280 };

// ---------- 状态 ----------
const state = {
  bitable: null,
  ImageQuality: ImageQualityFallback,
  table: null,
  tableName: '',
  fields: [],        // [{id, name, type, isPrimary}]
  records: [],       // [{recordId, fields}]
  maxAttach: {},     // fieldId -> 该字段单格最多附件数
  loaded: false,
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
  $('#progressBar').style.width = Math.max(0, Math.min(100, p)) + '%';
}
function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

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

// ---------- 初始化 ----------
async function init() {
  const okSdk = await loadSdk();
  if (!okSdk) {
    setStatus('未加载飞书 SDK（请检查网络 / CDN）', 'err');
    $('#envHint').classList.remove('hidden');
    return;
  }
  try {
    const table = await state.bitable.base.getActiveTable();
    state.table = table;
    state.tableName = await table.getName();
    setStatus('已连接：' + state.tableName, 'ok');
    await loadFields();
    $('#btnLoad').disabled = false;
  } catch (e) {
    setStatus('未在飞书多维表环境中，或无法获取当前表：' + e.message, 'err');
    $('#envHint').classList.remove('hidden');
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
}

function renderFieldList() {
  const box = $('#fieldList');
  box.innerHTML = '';
  if (!state.fields.length) {
    box.innerHTML = '<div class="placeholder">该表没有字段。</div>';
    return;
  }
  for (const f of state.fields) {
    const row = document.createElement('label');
    row.className = 'field-item';
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = f.isAttachment; // 附件字段默认勾选
    cb.dataset.fieldId = f.id;
    const name = document.createElement('span');
    name.className = 'fname';
    name.textContent = f.name + (f.isPrimary ? '（主键）' : '');
    row.appendChild(cb);
    row.appendChild(name);
    if (f.isAttachment) {
      const tag = document.createElement('span');
      tag.className = 'ftag';
      tag.textContent = '📎 图片';
      row.appendChild(tag);
    }
    box.appendChild(row);
  }
}

function getSelectedFieldIds() {
  const includeAll = $('#includeAll').checked;
  const checked = new Set(
    [...document.querySelectorAll('#fieldList input[type=checkbox]:checked')].map((c) => c.dataset.fieldId)
  );
  // includeAll：勾选的 + 其它所有字段（作为文字列）
  return state.fields
    .filter((f) => checked.has(f.id) || (includeAll && !checked.has(f.id)))
    .map((f) => f.id);
}

// ---------- 读取全部记录 ----------
async function loadData() {
  if (!state.table) return;
  $('#btnLoad').disabled = true;
  $('#btnExport').disabled = true;
  state.loaded = false;
  state.records = [];
  state.maxAttach = {};
  setProgress(0);
  log('开始读取记录…');
  // 是否仅导出当前视图
  let viewId;
  if ($('#onlyCurrentView').checked) {
    try {
      const view = await state.bitable.base.getActiveView();
      viewId = view && view.id;
      if (viewId) {
        let vname = viewId;
        try { vname = await view.getName(); } catch (e) {}
        log('仅导出当前视图：' + vname);
      }
    } catch (e) {
      log('获取当前视图失败，将导出全部记录：' + e.message, 'warn');
      viewId = null;
    }
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
    setProgress(100);
    log('读取完成，共 ' + all.length + ' 行。', 'ok');
  } catch (e) {
    log('读取失败：' + e.message, 'err');
    setStatus('读取数据失败', 'err');
  } finally {
    $('#btnLoad').disabled = false;
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

// 将 Blob 转成 Excel 可直接使用的 {base64, extension, width, height}
async function blobToExcelImage(blob) {
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

// 缩略图返回的是 base64 字符串（可能带 data: 前缀，也可能不带）
function detectMime(b64) {
  if (/^iVBORw0KGgo/.test(b64)) return 'image/png';
  if (/^\/9j\//.test(b64)) return 'image/jpeg';
  if (/^R0lGOD/.test(b64)) return 'image/gif';
  if (/^Qk/.test(b64)) return 'image/bmp';
  return 'image/png';
}
async function thumbToExcelImage(b64Str) {
  let raw = b64Str;
  if (/^data:/i.test(raw)) raw = raw.split(',')[1];
  const mime = detectMime(raw);
  const dataUrl = 'data:' + mime + ';base64,' + raw;
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
  const dims = await imageDims(dataUrl);
  return { base64: b64, extension: ext, width: dims.w, height: dims.h };
}

function blobToBitmap(dataUrl) {
  return fetch(dataUrl).then((r) => r.blob()).then((b) => createImageBitmap(b));
}
function imageDims(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth || 150, h: img.naturalHeight || 120 });
    img.onerror = () => resolve({ w: 150, h: 120 });
    img.src = dataUrl;
  });
}

// 用 <img crossOrigin> 加载原图再画到 canvas 取 base64（绕开 fetch 的 CORS 限制，拿到的是原图）
function loadImageBytes(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => {
      try {
        const c = document.createElement('canvas');
        c.width = img.naturalWidth; c.height = img.naturalHeight;
        c.getContext('2d').drawImage(img, 0, 0);
        const dataUrl = c.toDataURL('image/png'); // 若跨域未授权会被污染而抛错
        resolve(finalizeExcelImage(dataUrl, 'png'));
      } catch (e) { reject(e); }
    };
    img.onerror = () => reject(new Error('img load fail'));
    img.src = url;
  });
}

// 抓取一个附件字段单元格的全部图片：优先原图（fetch → <img> 跨域），最后才回退缩略图
async function fetchCellImages(fieldId, recordId, cellVal) {
  const tokens = (Array.isArray(cellVal) ? cellVal : []).filter((x) => x && x.token).map((x) => x.token);
  if (!tokens.length) return [];
  const out = new Array(tokens.length).fill(null);
  let urls = [];
  try {
    urls = await state.table.getCellAttachmentUrls(tokens, fieldId, recordId);
  } catch (e) {
    log('  获取附件 URL 失败（将尝试缩略图）：' + e.message, 'warn');
  }
  if (urls.length) {
    const results = await Promise.all(urls.map(async (u) => {
      // 1) fetch 原图（最清晰）
      try {
        const resp = await fetch(u, { mode: 'cors' });
        if (!resp.ok) throw new Error('HTTP ' + resp.status);
        return await blobToExcelImage(await resp.blob());
      } catch (e1) {
        // 2) 退一步用 <img crossOrigin> 取原图（部分环境 fetch 被拦但 img 可用）
        try { return await loadImageBytes(u); } catch (e2) { return null; }
      }
    }));
    results.forEach((r, i) => { if (r) out[i] = r; });
  }
  // 3) 仍失败的，用缩略图兜底（base64，不受 CORS 影响，但分辨率较低）
  const failedIdx = out.map((v, i) => (v ? -1 : i)).filter((i) => i >= 0);
  if (failedIdx.length) {
    try {
      const failedTokens = failedIdx.map((i) => tokens[i]);
      const thumbs = await state.table.getCellThumbnailUrls(
        failedTokens, fieldId, recordId, state.ImageQuality.MAX
      );
      await Promise.all(failedIdx.map(async (idx, k) => {
        if (thumbs[k]) {
          try { out[idx] = await thumbToExcelImage(thumbs[k]); }
          catch (e) { /* 忽略单张 */ }
        }
      }));
    } catch (e) {
      log('  缩略图兜底也失败：' + e.message, 'err');
    }
  }
  return out;
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

  const fieldIds = getSelectedFieldIds();
  if (!fieldIds.length) { log('请至少选择一个字段。', 'warn'); return; }
  const plan = buildColumnPlan(fieldIds);
  const DISPLAY_W = Math.max(60, Math.min(400, parseInt($('#imgWidth').value, 10) || 150));
  const concurrency = Math.max(1, Math.min(20, parseInt($('#concurrency').value, 10) || 6));

  $('#btnExport').disabled = true;
  $('#btnLoad').disabled = true;
  setProgress(0);
  log('开始生成 Excel…');

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

  // 列宽：图片列按显示宽度贴合，让图片正好落在格子里
  plan.forEach((c, i) => {
    const col = ws.getColumn(i + 1);
    col.width = c.isAttachment ? Math.max(12, (DISPLAY_W - 4) / 7) : 22;
  });

  // 限速并发
  let cursor = 0;
  const total = state.records.length;
  async function worker() {
    while (cursor < total) {
      const idx = cursor++;
      const rec = state.records[idx];
      const rowNum = idx + 2;
      const row = ws.getRow(rowNum);
      let maxH = 0;
      // 预取每个附件字段的图片（每行内并行）
      const attachCache = {};
      const attachFields = [...new Set(plan.filter((c) => c.isAttachment).map((c) => c.fieldId))];
      await Promise.all(attachFields.map(async (fid) => {
        const cell = rec.fields && rec.fields[fid];
        attachCache[fid] = await fetchCellImages(fid, rec.recordId, cell);
      }));
      for (let ci = 0; ci < plan.length; ci++) {
        const c = plan[ci];
        if (c.isAttachment) {
          const imgs = attachCache[c.fieldId] || [];
          const img = imgs[c.imgIndex];
          if (img) {
            const ratio = img.height / Math.max(1, img.width);
            const dispH = Math.round(DISPLAY_W * ratio);
            const imgId = wb.addImage({ base64: img.base64, extension: img.extension });
            ws.addImage(imgId, {
              tl: { col: ci, row: rowNum - 1 },
              ext: { width: DISPLAY_W, height: dispH },
            });
            maxH = Math.max(maxH, dispH);
          }
        } else {
          const f = state.fields.find((x) => x.id === c.fieldId);
          row.getCell(ci + 1).value = formatText(rec.fields ? rec.fields[c.fieldId] : undefined);
          if (f && f.isPrimary) row.getCell(ci + 1).font = { bold: true };
        }
      }
      if (maxH > 0) row.height = Math.max(20, maxH * 0.75 + 4);
      setProgress(Math.round(((idx + 1) / total) * 100));
      if (idx % 10 === 0) log('  处理第 ' + (idx + 1) + ' / ' + total + ' 行');
    }
  }

  const workers = [];
  for (let i = 0; i < concurrency; i++) workers.push(worker());
  await Promise.all(workers);

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
  log('导出完成：' + a.download, 'ok');
  await markExported(state.records);
  $('#btnExport').disabled = false;
  $('#btnLoad').disabled = false;
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
        await state.table.setCellValue(fieldId, rec.recordId, true);
        okCount++;
      } catch (e) {
        log('  标记失败（第 ' + (i + 1) + ' 行）：' + e.message, 'warn');
      }
      done++;
      setProgress(Math.round((done / n) * 100));
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

// ---------- 绑定 ----------
window.addEventListener('DOMContentLoaded', () => {
  $('#btnLoad').addEventListener('click', loadData);
  $('#btnExport').addEventListener('click', exportExcel);
  ensureExcelJS().then((ok) => {
    if (!ok) setStatus('ExcelJS 加载失败（导出将不可用）', 'err');
    init();
  });
});
