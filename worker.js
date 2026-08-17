// 飞书多维表图片导出插件 —— 原图 CORS 代理 Worker（Cloudflare Workers）
// 作用：插件把 getCellAttachmentUrls 拿到的「原图下载链接」发给本 Worker，
//      Worker 在服务端（无浏览器 CORS 限制）拉取字节，再带 Access-Control-Allow-Origin 返回，
//      浏览器即可把原图像素画到 canvas 嵌入 Excel。
// 安全：仅允许转发飞书 / 飞书 CDN 域名，拒绝任意其他 URL（避免变成开放代理）。
// 部署：见 README-worker.md（wrangler deploy，免费）。

export default {
  async fetch(request, env, ctx) {
    // 处理 CORS 预检
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: corsHeaders() });
    }
    if (request.method !== 'GET') {
      return json({ error: 'method not allowed' }, 405);
    }

    const url = new URL(request.url);
    const target = url.searchParams.get('u'); // 原图下载链接
    if (!target) {
      return json({ error: 'missing u param' }, 400);
    }

    let t;
    try {
      t = new URL(target);
    } catch (e) {
      return json({ error: 'invalid u' }, 400);
    }
    if (!isAllowedHost(t.hostname)) {
      return json({ error: 'host not allowed: ' + t.hostname }, 403);
    }

    try {
      const upstream = await fetch(target, { redirect: 'follow' });
      if (!upstream.ok) {
        return json({ error: 'upstream http ' + upstream.status }, upstream.status);
      }
      const buf = await upstream.arrayBuffer();
      const ct = upstream.headers.get('content-type') || 'application/octet-stream';
      return new Response(buf, {
        status: 200,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': ct,
          'Cache-Control': 'public, max-age=3600',
        },
      });
    } catch (e) {
      return json({ error: 'fetch failed: ' + e.message }, 502);
    }
  },
};

function isAllowedHost(host) {
  return (
    host === 'open.feishu.cn' ||
    host.endsWith('.feishu.cn') ||
    host === 'open.larksuite.com' ||
    host.endsWith('.larksuite.com') ||
    host.endsWith('.feishucdn.com')
  );
}

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': '*',
  };
}

function json(obj, status) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...corsHeaders(), 'Content-Type': 'application/json' },
  });
}
