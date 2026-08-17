# 原图代理 Worker 部署说明（Cloudflare Workers）

飞书前端 SDK 拿不到附件**原图**的像素（飞书 CDN 跨域不返回 CORS 头，canvas 会被污染）。
本 Worker 只做一件事：**在服务端（无浏览器 CORS 限制）把插件发来的「飞书原图下载链接」拉成字节，
再带 `Access-Control-Allow-Origin: *` 返回给浏览器**，于是插件就能把真·原图嵌入 Excel。

> 关键点：这是**免凭证版**——Worker 只是 CORS 代理，直接转发插件给的下载链接，
> **不需要你的飞书 app_id / app_secret**。安全上只放行飞书 / 飞书 CDN 域名，拒绝其他 URL。

## 一、部署（约 3 分钟，免费）

1. 注册 / 登录 https://dash.cloudflare.com （免费版够用）。
2. 安装 wrangler（Node 环境）：
   ```bash
   npm install -g wrangler
   wrangler login
   ```
3. 在本项目目录执行：
   ```bash
   wrangler deploy
   ```
   成功后终端会输出你的 Worker 地址，形如：
   `https://feishu-img-proxy.<你的子域>.workers.dev`
   复制这个地址。

## 二、在插件里启用

1. 飞书多维表里打开插件（硬刷新一次加载最新版）。
2. 选项区找到「**原图代理（Cloudflare Worker URL）**」，粘贴上面的 Worker 地址。
3. 点「导出 Excel」/「导出图片 ZIP」。
4. 看日志最后一行：`原图 N / 缩略图 M`。
   - 填了 Worker 且 `N > 0` → 已是**真·原图**分辨率。
   - 仍 `N = 0` → Worker 没通（地址错 / 未部署 / 网络），自动回退缩略图。检查 Worker 地址和部署状态。

## 三、常见问题

- **Worker 返回 403 host not allowed**：说明插件发来的不是飞书域名链接，正常防护，忽略即可（此时插件回退缩略图）。
- **Worker 返回 502 fetch failed**：飞书下载链接过期或网路抖动。导出时插件实时取最新链接再发 Worker，一般重试即可；若频繁，可改用进阶版（见下）。
- **想更稳 / 不想依赖前端临时链接**：进阶版让 Worker 用 `tenant_access_token` 自己调
  `GET https://open.feishu.cn/open-apis/drive/v1/medias/{file_token}/download`。
  需要你提供飞书自建应用的 `app_id` / `app_secret`，配成 Worker Secrets，并改 `worker.js` 走 token 模式。需要我给进阶版代码再说。

## 四、本地自测 Worker（可选）

```bash
wrangler dev
# 另开终端：
curl "http://127.0.0.1:8787/?u=<一个飞书附件下载链接>" -o test.jpg
# 能下载到 test.jpg 即代理正常
```
