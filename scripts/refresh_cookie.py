#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
自动刷新网易问卷系统 Cookie

使用 Playwright 打开浏览器，自动获取登录态 Cookie 并保存。
- 首次运行：需要手动登录（登录后自动保存 session）
- 后续运行：复用已保存的 session，自动获取新 Cookie（无需重新登录）

使用前需安装 Playwright：
  pip install playwright
  playwright install chromium
"""

import json
import os
import sys
import time


# 双平台支持
PLATFORM_URLS = {
    "cn": "https://survey-game.163.com/index.html#/surveylist",
    "intl": "https://survey-game.easebar.com/index.html#/surveylist",
}
PLATFORM_DOMAINS = {
    "cn": "survey-game.163.com",
    "intl": "survey-game.easebar.com",
}
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")


def _log(msg):
    print(f"[refresh_cookie] {msg}", flush=True)


def refresh_cookie(timeout=300, platform="cn"):
    """
    自动刷新 Cookie。
    1. 打开浏览器访问问卷系统
    2. 如果已有登录态，自动获取 Cookie
    3. 如果没有，等待用户手动登录
    4. 通过页面内 API 调用验证登录成功后保存所有 Cookie

    platform: "cn"=国内, "intl"=国外
    返回: True=成功, False=失败
    """
    survey_url = PLATFORM_URLS.get(platform, PLATFORM_URLS["cn"])
    target_domain = PLATFORM_DOMAINS.get(platform, PLATFORM_DOMAINS["cn"])
    # 每个平台使用独立的浏览器 profile，避免 session 冲突
    # 放在用户主目录下，避免路径含特殊字符（如 &）导致浏览器启动失败
    profile_dir = os.path.join(os.path.expanduser("~"), ".survey_download", f"browser_profile_{platform}")
    os.makedirs(profile_dir, exist_ok=True)
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        _log("ERROR: Playwright not installed.")
        _log("  pip install playwright")
        _log("  playwright install chromium")
        return False

    _log(f"Launching browser for {platform} platform...")
    with sync_playwright() as p:
        # 使用持久化上下文，保留登录 session（每个平台独立）
        # 优先用系统 Edge，如果失败则回退到 Playwright 内置 Chromium
        launch_kwargs = {
            "user_data_dir": profile_dir,
            "headless": False,
            "args": ["--disable-blink-features=AutomationControlled"],
        }
        try:
            context = p.chromium.launch_persistent_context(channel="msedge", **launch_kwargs)
        except Exception as e:
            _log(f"Edge launch failed ({e}), falling back to built-in Chromium...")
            context = p.chromium.launch_persistent_context(**launch_kwargs)

        page = context.pages[0] if context.pages else context.new_page()
        _log(f"Navigating to {survey_url}")
        page.goto(survey_url, wait_until="domcontentloaded")

        _log("Waiting for login...")
        _log("(If you see the login page, please log in manually. The script will auto-detect.)")

        start_time = time.time()
        while time.time() - start_time < timeout:
            # 策略：不靠 Cookie 名判断，而是直接在页面内调 API 看是否返回问卷列表
            try:
                # 先确保页面在正确的 URL 上（可能被重定向到登录页）
                current_url = page.url
                on_login_page = (
                    "login.netease.com" in current_url
                    or target_domain not in current_url
                )
                if on_login_page:
                    # 还在登录页，等待用户操作
                    time.sleep(3)
                    elapsed = int(time.time() - start_time)
                    if elapsed % 30 == 0 and elapsed > 0:
                        _log(f"Still waiting for login... ({elapsed}s / {timeout}s)")
                    continue

                # 页面已在问卷系统域名上，尝试通过页面内 fetch 验证
                _log(f"Page on target domain: {current_url[:60]}")
                resp = page.evaluate("""async () => {
                    try {
                        const r = await fetch('/view/survey/list', {
                            method: 'POST',
                            headers: {'Content-Type': 'application/json', 'X-Requested-With': 'XMLHttpRequest'},
                            body: JSON.stringify({pageNo:1,surveyName:"",status:"-1",deliveryRange:-1,type:-1,groupId:-1,groupUser:-1,gameName:""})
                        });
                        const text = await r.text();
                        try { return JSON.parse(text); } catch(e) { return {"_raw": text.substring(0, 200)}; }
                    } catch(e) {
                        return {"_error": e.message};
                    }
                }""")

                if resp and resp.get("resultCode") == 100:
                    _log("Login verified! Extracting cookies...")
                    # 收集该平台域名下的所有 Cookie
                    all_cookies = context.cookies()
                    cookie_dict = {}
                    for c in all_cookies:
                        if target_domain in c.get("domain", ""):
                            cookie_dict[c["name"]] = c["value"]

                    if not cookie_dict:
                        _log("Warning: no cookies found for domain, collecting all...")
                        for c in all_cookies:
                            cookie_dict[c["name"]] = c["value"]

                    _log(f"Collected {len(cookie_dict)} cookies")
                    config = {
                        "platform": platform,
                        "cookies": cookie_dict,
                        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                    }
                    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                        json.dump(config, f, ensure_ascii=False, indent=2)
                    _log(f"Cookies saved to {CONFIG_PATH}")
                    context.close()
                    return True
                else:
                    _log(f"API response: {str(resp)[:100]}, retrying...")

            except Exception as e:
                _log(f"Check error: {e}")

            time.sleep(3)

        _log(f"Timeout after {timeout}s. Failed to detect valid login.")
        context.close()
        return False


def main():
    import argparse
    parser = argparse.ArgumentParser(description="自动刷新网易问卷系统 Cookie")
    parser.add_argument("--timeout", type=int, default=300, help="等待登录超时（秒，默认300）")
    parser.add_argument("--platform", choices=["cn", "intl"], default="cn",
                        help="平台: cn=国内, intl=国外（默认 cn）")
    args = parser.parse_args()

    success = refresh_cookie(timeout=args.timeout, platform=args.platform)
    if success:
        _log("✓ Cookie refresh completed!")
        print(json.dumps({"status": "success", "message": "Cookie 已自动刷新"}, ensure_ascii=False))
    else:
        _log("× Cookie refresh failed.")
        print(json.dumps({"status": "error", "message": "Cookie 刷新失败"}, ensure_ascii=False))
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()