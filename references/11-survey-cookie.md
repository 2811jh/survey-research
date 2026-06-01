# Cookie 处理

## 核心原则

**Cookie 失效时，直接自动弹出浏览器让用户登录，不要询问用户选择哪种方式。**
用户不需要知道 Cookie 是什么，不需要打开 F12，不需要做任何技术操作。

## 自动刷新机制

脚本内部已集成全自动刷新。当 `check`、`download`、`clean` 检测到 Cookie 失效时：

1. **自动调用** `refresh_cookie.py`，弹出浏览器窗口（使用持久化 profile，存放于 `scripts/.browser_profile/`）
2. 如果**首次使用**：浏览器显示登录页，用户输入账号密码完成登录（仅需一次）
3. 如果**之前已登录过**：浏览器进入网易 SSO「确认登录」页（显示已记住的账号）。**脚本自动点击"确认登录"按钮**，无需用户手动操作
4. 检测到目标 cookies 后，等待 SSO 重定向链上的所有 cookies 落盘（约 2 秒），保存后自动关闭浏览器

**整个过程对用户来说就是：弹出浏览器 → (首次) 登录一次 → (后续) 全自动完成。**

> 💡 **持久化机制**：登录会话保存在 `scripts/.browser_profile/` 与 `scripts/.browser_profile_global/`，
> 长期复用直到网易 SSO 主动失效。该目录已在 `.gitignore` 中排除，不会泄露账号信息。

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

首次使用需一次性安装 Playwright：
```bash
pip install playwright
playwright install chromium
```

如果用户环境没有 Playwright，**直接帮用户执行安装命令**，不要让用户自己去搞。

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
