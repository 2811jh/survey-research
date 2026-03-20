# Cookie 处理

脚本内部已集成自动刷新机制。正常情况下无需手动处理。

## 自动刷新

当 `check`、`download`、`clean` 检测到 Cookie 失效时，会自动调用 `refresh_cookie.py`：
- 打开浏览器访问问卷系统
- 首次运行：用户在浏览器中登录，session 保存到 `.browser_profile`
- 后续运行：复用 session，全自动，无需人工操作

依赖 Playwright（一次性安装）：
```bash
pip install playwright
playwright install chromium
```

## 手动刷新

也可以手动触发：
```bash
python {SKILL_DIR}/scripts/refresh_cookie.py
```

## 手动备用方案

如果 Playwright 不可用（未安装），引导用户：

1. 浏览器访问对应平台：
   - 国内：`https://survey-game.163.com/index.html#/surveylist`
   - 国外：`https://survey-game.easebar.com/index.html#/surveylist`
2. `Ctrl+Shift+I` → 「应用程序」→「Cookie」→ 找到 `SURVEY_TOKEN` 和 `JSESSIONID`
3. 执行：

```bash
python {SKILL_DIR}/scripts/survey_download.py init --survey_token "xxx" --jsessionid "xxx"
```
