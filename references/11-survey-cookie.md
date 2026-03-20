# Cookie 处理

## 核心原则

**Cookie 失效时，直接自动弹出浏览器让用户登录，不要询问用户选择哪种方式。**
用户不需要知道 Cookie 是什么，不需要打开 F12，不需要做任何技术操作。

## 自动刷新机制

脚本内部已集成全自动刷新。当 `check`、`download`、`clean` 检测到 Cookie 失效时：

1. **自动调用** `refresh_cookie.py`，弹出浏览器窗口
2. 如果已有登录态（session 未过期），浏览器自动获取 Cookie，**全程无需人工操作**
3. 如果 session 也过期了，浏览器会显示登录页面，**用户只需在弹出的浏览器中正常登录即可**
4. 登录成功后脚本自动检测、保存 Cookie，浏览器自动关闭

**整个过程对用户来说就是：弹出一个浏览器 → 登录 → 自动继续。**

## AI 端行为规范

⚠️ **严禁**出现以下行为：
- ❌ 询问用户"选择哪种登录方式"
- ❌ 让用户去 F12 控制台复制 Cookie
- ❌ 让用户手动提供 SURVEY_TOKEN 或 JSESSIONID
- ❌ 给用户展示 `init --survey_token` 命令

✅ **正确做法**：
- Cookie 失效 → 直接告知用户"正在为您打开浏览器，请在弹出的页面中登录"
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
python {SKILL_DIR}/scripts/refresh_cookie.py --platform intl
```