# 部署指南（国内可访问托管）

> 适用场景：原 GitHub Pages 地址（`yunjueyin.github.io`）在你的网络 / 飞书内被墙，插件打不开。
> 本目录 `deploy-package.zip` 已含全部前端静态文件，解压即可上传到任意静态托管。

---

## 一、先说结论（好消息）

插件**完全不需要任何后端 / 代理 / Worker**，纯前端就能跑。

取图逻辑（见 `app.js` 的 `fetchCellImages`）：
1. **缩略图模式（默认·最快）**：直接调用飞书 SDK `getCellThumbnailUrls`，飞书**直接返回 base64**，
   不跨域、不逐个联网、不需要公网，导出最快最稳，体积也小（正是你要的小图效果）。
2. **高清原图模式**：调用飞书 SDK `getCellAttachmentUrls` 拿到附件直连 URL，本地 `fetch` 取原图；
   若被飞书 CDN 的 CORS 策略拦截，则自动回退到上面的缩略图。整个过程**不借助任何代理**。

所以：只要把前端静态文件放到「飞书能加载的 HTTPS 地址」，插件即可用，没有任何外部依赖。

---

## 二、方案 A：Gitee Pages（推荐，国内快、免费）

Gitee 是国内代码托管，Gitee Pages 生成的 `*.gitee.io` 在国内与飞书内均可访问。

1. 注册 / 登录 https://gitee.com （建议用户名沿用 `yunjueyin`）。
2. 新建**公开**仓库，例如 `feishu-bitable-image-excel`。
3. 把本包解压后的全部文件推上去（或把仓库地址 + 一个 Gitee 私人令牌 PAT 给我，我直接帮你推）。
4. 仓库页 → **服务 → Gitee Pages** → 部署（分支 `main`，部署目录 `/`）→ 得到地址
   `https://yunjueyin.gitee.io/feishu-bitable-image-excel`
5. 飞书「多维表格 → 自定义插件」里，把网页地址改成上面的 Gitee Pages 地址，重新加载即可。

> 注意：Gitee Pages 免费版部署后会有约 1 分钟 CDN 刷新延迟；若打开是旧内容，清缓存或等一会。

---

## 三、方案 B：腾讯云 COS 静态网站（稳，需有腾讯云账号）

1. 对象存储 COS 建一个**公有读**存储桶，开启「静态网站托管」。
2. 把本包文件全部上传到桶根目录（`index.html` 设为默认首页）。
3. COS 会给出一个形如 `https://<bucket>-<appid>.cos-website.ap-guangzhou.myqcloud.com` 的访问地址；
   如需自定义域名再绑定。该地址国内 / 飞书内均可访问。

---

## 四、方案 C：阿里云 OSS 静态网站（稳，需有阿里云账号）

1. OSS 建**公共读** Bucket，开启「静态页面」，`默认首页` 设为 `index.html`。
2. 上传本包全部文件。
3. 使用 OSS 提供的 `*.oss-cn-<region>.aliyuncs.com` 访问地址（或绑定自定义域名）。

---

## 五、方案 D：Cloudflare Pages（免费，但 Cloudflare 国内偶尔抽风）

1. 本机把代码推到 GitHub（你本地网络若可达 GitHub）。
2. Cloudflare Pages 关联该仓库，Framework 选 `None`，构建命令留空，输出目录 `/`。
3. 得到 `*.pages.dev` 地址。国内访问 Cloudflare 有时不稳定，作备选。

---

## 六、验证

- 浏览器直接打开托管地址，应看到插件界面（字段列表、数据表下拉、导出按钮）。
- 飞书内重新加载自定义插件，确认能正常加载、能导出带图 Excel。
- 若图片为空：确认当前多维表有附件字段；若选了「高清原图」却全为缩略图，通常是飞书 CDN 的
  CORS 拦截所致（正常现象），在「图片设置 → 图片质量」改回「缩略图（最快）」即可。

---

## 七、安全提醒（重要）

- 此前构建环境曾把 GitHub 令牌明文写进 `.git/config` 并经对话暴露，**请立即到 GitHub → Settings → Developer settings → Personal access tokens 吊销旧令牌并重新生成**。
- 若提供 Gitee PAT，请仅授予 `projects` / 仓库写权限，用完后可在 Gitee → 设置 → 私人令牌 吊销。
