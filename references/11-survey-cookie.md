# Cookie 处理

## 核心原则

**Cookie 失效时，直接自动弹出浏览器让用户登录，不要询问用户选择哪种方式。**
用户不需要知道 Cookie 是什么，不需要打开 F12，不需要做任何技术操作。

## 自动刷新机制

脚本内部已集成全自动刷新。当 `check`、`download`、`clean` 检测到 Cookie 失效时，
**按优先级尝试两条路径，任一成功即继续，全程无需用户手动复制 Cookie**：

### 路径 1（优先）：CDP 复用日常浏览器登录态

若用户日常浏览器（Edge/Chrome）已开启远程调试、且已登录问卷平台，脚本会经内嵌的
`scripts/webaccess/grab_cookie.mjs`（Node）通过 CDP `Storage.getCookies` 直接读取登录
cookie（含 HttpOnly 的 `JSESSIONID`/`SURVEY_TOKEN`）。

- **优点**：免独立 profile、免首次手动登录，直接复用用户已有的登录态。
- **前提**：① 本机装了 Node（`node -v`，需 22+，原生 WebSocket）；② 日常浏览器以远程调试模式运行
  （地址栏进 `chrome://inspect` 或 `edge://inspect`，开启 "Discover network targets" / 用
  `--remote-debugging-port=9222` 启动）。
- **多浏览器**：运行中途检测到多个可用浏览器且未设偏好时**不打断流程**，直接回退路径 2；
  可先做一次性设置（见下）固定浏览器后即可启用。

### 路径 2（回退）：Playwright 独立浏览器登录

CDP 不可用（无 Node / 未开远程调试 / 未登录 / 多浏览器无偏好 / 任何异常）时，
**静默回退**到原有 `refresh_cookie.py` 流程：

1. **自动调用** `refresh_cookie.py`，弹出浏览器窗口（使用持久化 profile，存放于 `scripts/.browser_profile/`）
2. 如果**首次使用**：浏览器显示登录页，用户输入账号密码完成登录（仅需一次）
3. 如果**之前已登录过**：浏览器进入网易 SSO「确认登录」页（显示已记住的账号）。**脚本自动点击"确认登录"按钮**，无需用户手动操作
4. 检测到目标 cookies 后，等待 SSO 重定向链上的所有 cookies 落盘（约 2 秒），保存后自动关闭浏览器

**对用户来说就是：(理想) CDP 秒级复用登录态；(回退) 弹出浏览器 → 首次登录一次 → 后续全自动。**

> 💡 **持久化机制**：Playwright 登录会话保存在 `scripts/.browser_profile/` 与 `scripts/.browser_profile_global/`，
> 长期复用直到网易 SSO 主动失效。该目录已在 `.gitignore` 中排除，不会泄露账号信息。

## CDP 一次性设置（可选，用于启用免登录）

多浏览器场景需先固定一个浏览器；单浏览器则自动使用，无需设置。

```bash
# 1. 列出已开启远程调试、且真正说 DevTools 协议的浏览器（空列表=没开或没装 Node）
python {SKILL_DIR}/scripts/survey_download.py --platform cn cookie-cdp --list

# 2. 抓取并持久化偏好（多浏览器时用 --browser 指定 id：chrome/edge/chromium）
python {SKILL_DIR}/scripts/survey_download.py --platform cn cookie-cdp --browser edge --save-pref
```

- `cookie-cdp --list` 返回 `{"status":"list","browsers":[...]}`。**空列表**表示当前无可用 DevTools 端点
  （浏览器未开远程调试、或端口被占用/残留、或未装 Node）——此时不必强推 CDP，回退 Playwright 即可。
- `cookie-cdp`（不带 `--list`）status 语义：`ok`（已写入 config，含 `auth_valid`）/ `ambiguous`（多浏览器，
  需用 `--browser` 指定）/ `no_login`（浏览器已连但未登录问卷平台）/ `empty`（无可用端点）/
  `unavailable`（无 Node 或脚本缺失）/ `mismatch`（指定的浏览器没检测到）。
- **偏好持久化**：`--save-pref` 成功后把浏览器 id 写入 config 的 `web_access_browser`，
  后续自动刷新会优先用它走 CDP；刷新 cookie 时该偏好不会被覆盖丢失。

> ⚠️ **AI 端**：只在用户想启用/排查免登录时才主动跑 `cookie-cdp`。日常下载/清洗遇到 Cookie 失效
> **不要**先问用户，脚本会自己按「CDP → Playwright」顺序处理。仅当两条路径都失败时才告知用户。

## AI 端行为规范

⚠️ **严禁**出现以下行为：
- ❌ 询问用户"选择哪种登录方式"
- ❌ 让用户去 F12 控制台复制 Cookie
- ❌ 让用户手动提供 SURVEY_TOKEN 或 JSESSIONID
- ❌ 给用户展示 `init --survey_token` 命令

✅ **正确做法**：
- Cookie 失效 → 直接告知用户"正在为您打开浏览器"
- 用户首次使用提示："首次需要您登录一次，登录态会自动保存，下次起会自动跳过登录"
- 登录成功后 → 继续执行原来的操作（下载/清洗等），不中断流程
- 如果刷新失败（超时/Playwright 未安装）→ 告知用户安装命令后重试

## 依赖安装

**Playwright（回退路径必需）**——首次使用需一次性安装：
```bash
pip install playwright
playwright install chromium
```

如果用户环境没有 Playwright，**直接帮用户执行安装命令**，不要让用户自己去搞。

**Node（CDP 优先路径可选）**——用于复用日常浏览器登录态：需 Node 22+（原生 WebSocket）。
没装 Node 不影响使用，脚本会自动回退 Playwright。

## 手动触发（仅调试用）

```bash
python {SKILL_DIR}/scripts/refresh_cookie.py --platform cn
python {SKILL_DIR}/scripts/refresh_cookie.py --platform global
```

## 故障排查

| 现象 | 排查方向 |
|------|----------|
| 每次都要重新输入账号密码 | 检查 `scripts/.browser_profile/` 是否存在且包含 `Default/` 目录；profile 被误删时需重新登录一次 |
| SSO 「确认登录」页未自动点击 | 网络慢导致页面未加载完成；脚本会持续等待，可继续手动点击；或 SSO 页面 DOM 结构变更需更新选择器 |
| Playwright 报错 "Executable doesn't exist" | 执行 `playwright install chromium` 或 `playwright install msedge` |
| CDP `cookie-cdp --list` 返回空列表 | 浏览器没开远程调试（进 `edge://inspect` 开启，或用 `--remote-debugging-port=9222` 启动）；或该端口被别的进程占用（脚本已用 `/json/version` 校验，非真 DevTools 端点会被过滤）；或未装 Node。无需强推，回退 Playwright 即可 |
| CDP 返回 `no_login` | 浏览器已连上但没登录问卷平台；先在该浏览器手动登录一次问卷后台再重试，或直接走 Playwright |
