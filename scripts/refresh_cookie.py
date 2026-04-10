#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
自动刷新网易问卷系统 Cookie（认证体系对齐 survey-checker）

使用 Playwright + Edge 打开浏览器，通过检测 required_cookies 判断登录状态。
- 首次运行：需要手动登录（登录后自动保存 session）
- 后续运行：复用 .browser_profile 保留的 session，自动获取新 Cookie（无需重新登录）

平台支持:
  cn     → survey-game.163.com      → config.json          → .browser_profile/
  global → survey-game.easebar.com  → config_global.json   → .browser_profile_global/
"""

import json
import os
import sys
import time

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 平台配置（与 survey-checker/core/constants.py 完全一致）
PLATFORMS = {
    "cn": {
        "label": "国内",
        "base_url": "https://survey-game.163.com",
        "cookie_domain": "survey-game.163.com",
        "target_cookies": {"SURVEY_TOKEN", "JSESSIONID", "P_INFO"},
        "required_cookies": {"SURVEY_TOKEN", "JSESSIONID"},
    },
    "global": {
        "label": "国外",
        "base_url": "https://survey-game.easebar.com",
        "cookie_domain": "survey-game.easebar.com",
        "target_cookies": {"oversea-online_SURVEY_TOKEN", "SURVEY_TOKEN", "JSESSIONID", "P_INFO"},
        "required_cookies": {"oversea-online_SURVEY_TOKEN"},
    },
}


def _log(msg):
    print(f"[refresh_cookie] {msg}", flush=True)


def _config_file(platform="cn"):
    """返回对应平台的 config 文件路径（与 survey-checker 一致）"""
    if platform == "cn":
        return os.path.join(SCRIPT_DIR, "config.json")
    return os.path.join(SCRIPT_DIR, f"config_{platform}.json")


def _profile_dir(platform="cn"):
    """返回对应平台的浏览器 profile 目录（与 survey-checker 一致）"""
    if platform == "cn":
        return os.path.join(SCRIPT_DIR, ".browser_profile")
    return os.path.join(SCRIPT_DIR, f".browser_profile_{platform}")


def save_cookies(platform, cookie_dict):
    """将 Cookie dict 保存到对应平台的 config 文件"""
    cfg = _config_file(platform)
    config = {
        "cookies": cookie_dict,
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }
    with open(cfg, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    _log(f"Cookies saved to {cfg}")


def refresh_cookie(platform="cn", timeout=300):
    """
    用 Playwright 打开浏览器，等待登录后自动保存 Cookie。
    登录检测策略：检测 required_cookies 是否存在（对齐 survey-checker/core/auth.py）。
    返回 True=成功，False=失败
    """
    plat = PLATFORMS[platform]
    base_url = plat["base_url"]
    target_cookies = plat["target_cookies"]
    required_cookies = plat["required_cookies"]
    profile_dir_path = _profile_dir(platform)
    survey_url = f"{base_url}/index.html#/surveylist"

    os.makedirs(profile_dir_path, exist_ok=True)

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        _log("ERROR: Playwright not installed. Run: pip install playwright && playwright install chromium")
        return False

    _log(f"Platform: {plat['label']} ({base_url})")
    _log("Launching browser...")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir_path,
            channel="msedge",
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        _log(f"Navigating to {survey_url}")
        page.goto(survey_url, wait_until="domcontentloaded")
        _log("Waiting for login cookies...")
        _log("(If you see the login page, please log in manually.)")

        start_time = time.time()
        while time.time() - start_time < timeout:
            cookies = context.cookies()
            cookie_dict = {
                c["name"]: c["value"]
                for c in cookies
                if c["name"] in target_cookies
            }
            if required_cookies.issubset(cookie_dict.keys()):
                _log("Detected required cookies, saving...")
                save_cookies(platform, cookie_dict)
                context.close()
                return True
            time.sleep(2)
            elapsed = int(time.time() - start_time)
            if elapsed % 30 == 0 and elapsed > 0:
                _log(f"Still waiting... ({elapsed}s / {timeout}s)")

        _log(f"Timeout after {timeout}s.")
        context.close()
        return False


def main():
    import argparse
    parser = argparse.ArgumentParser(description="自动刷新网易问卷系统 Cookie")
    parser.add_argument("--timeout", type=int, default=300, help="等待登录超时（秒，默认300）")
    parser.add_argument(
        "--platform", choices=["cn", "global"], default="cn",
        help="平台: cn=国内(163.com), global=国外(easebar.com)（默认 cn）",
    )
    args = parser.parse_args()

    success = refresh_cookie(platform=args.platform, timeout=args.timeout)
    if success:
        _log("✓ Cookie refresh completed!")
        print(json.dumps({"status": "success", "message": "Cookie 已自动刷新"}, ensure_ascii=False))
    else:
        _log("× Cookie refresh failed.")
        print(json.dumps({"status": "error", "message": "Cookie 刷新失败"}, ensure_ascii=False))
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()