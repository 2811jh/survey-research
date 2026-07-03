#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
自动刷新网易问卷系统 Cookie（认证体系对齐 survey-checker）

使用 Playwright 打开浏览器（多浏览器：按 Chrome→Edge→内置 Chromium 自动挑可用的，
也可用 --browser 指定），通过检测 required_cookies 判断登录状态。
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


def _channel_candidates(preferred=None):
    """
    构造 Playwright 启动通道候选（有序，去重）。
    channel=None 表示用 Playwright 自带的 chromium（需 `playwright install chromium`）。

    preferred: 浏览器偏好 id（chrome/edge/msedge/chromium），来自 CDP 抓取时保存的
    web_access_browser，或命令行 --browser。优先尝试。

    不写死单一浏览器：按「偏好 → Chrome → Edge → 自带 Chromium」顺序逐个尝试启动，
    用第一个装了的。这样 Mac 没装 Edge 也能用 Chrome / 自带 Chromium，行为对齐 web-access 的多浏览器。
    """
    pref_map = {
        "chrome": ("chrome", "Chrome"),
        "edge": ("msedge", "Microsoft Edge"),
        "msedge": ("msedge", "Microsoft Edge"),
        "chromium": (None, "Chromium(内置)"),
    }
    order = []

    def add(channel, label):
        # 按 channel 去重（而非整个 tuple），避免 label 差异导致同一浏览器被重复尝试
        if not any(c == channel for c, _ in order):
            order.append((channel, label))

    if preferred and preferred in pref_map:
        add(*pref_map[preferred])
    add("chrome", "Chrome")
    add("msedge", "Microsoft Edge")
    add(None, "Chromium(内置)")
    return order


def _launch_any_browser(p, profile_dir_path, preferred=None):
    """
    依次尝试候选浏览器通道，返回 (context, label)；全部失败返回 (None, None)。
    仅当某通道对应的浏览器未安装（Executable doesn't exist 等）时才跳到下一个。
    """
    last_err = None
    for channel, label in _channel_candidates(preferred):
        try:
            kwargs = dict(
                user_data_dir=profile_dir_path,
                headless=False,
                args=["--disable-blink-features=AutomationControlled"],
            )
            if channel:
                kwargs["channel"] = channel
            context = p.chromium.launch_persistent_context(**kwargs)
            _log(f"Launched browser: {label}" + (f" (channel={channel})" if channel else " (bundled chromium)"))
            return context, label
        except Exception as e:
            last_err = e
            _log(f"  {label} 不可用，尝试下一个：{str(e).splitlines()[0]}")
    _log(f"ERROR: 没有可用的浏览器（已尝试 Chrome/Edge/内置 Chromium）。最后错误：{last_err}")
    _log("  可安装其一：Chrome 或 Edge；或执行 `playwright install chromium` 使用内置浏览器。")
    return None, None


def refresh_cookie(platform="cn", timeout=300, preferred_browser=None):
    """
    用 Playwright 打开浏览器，等待登录后自动保存 Cookie。
    登录检测策略：检测 required_cookies 是否存在（对齐 survey-checker/core/auth.py）。

    自动化登录流程：
    1. 打开问卷系统主页（自动重定向到网易 SSO）
    2. 如检测到 SSO 「确认登录」页面（已记住账号），自动点击确认登录按钮
    3. 检测到 required_cookies 后，等待 1 秒确保所有 cookies 落盘，再保存

    .browser_profile 持久化：首次手动登录后，SSO cookies 会保留在 user_data_dir，
    后续运行直接进入「确认登录」页，自动点击即可，全程无需手动输入。

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
    _log(f"Browser profile: {profile_dir_path}")
    _log("Launching browser...")

    context = None
    try:
        with sync_playwright() as p:
            context, _ = _launch_any_browser(p, profile_dir_path, preferred=preferred_browser)
            if context is None:
                return False
            page = context.pages[0] if context.pages else context.new_page()
            _log(f"Navigating to {survey_url}")
            page.goto(survey_url, wait_until="domcontentloaded")
            _log("Waiting for login cookies...")
            _log("(If you see the login page, please log in manually.)")
            _log("(If you see the SSO confirm page, the script will auto-click for you.)")

            start_time = time.time()
            sso_clicked = False
            while time.time() - start_time < timeout:
                # 自动点击 SSO「确认登录」按钮（仅尝试一次）
                if not sso_clicked:
                    try:
                        current_url = page.url or ""
                        if "login.netease.com" in current_url and "redirect-to-affirm" in current_url:
                            # 这是 SSO 确认登录页：尝试点击「确认登录」按钮
                            confirm_btn = page.locator(
                                "button:has-text('确认登录'), input[type='submit'][value='确认登录'], a:has-text('确认登录')"
                            ).first
                            if confirm_btn.count() > 0:
                                _log("Detected SSO confirm page, auto-clicking 「确认登录」...")
                                confirm_btn.click(timeout=5000)
                                sso_clicked = True
                                page.wait_for_load_state("domcontentloaded", timeout=15000)
                    except Exception as e:
                        _log(f"  Auto-click failed (will continue waiting for manual login): {e}")
                        sso_clicked = True  # 不再重试，避免反复报错

                cookies = context.cookies()
                cookie_dict = {
                    c["name"]: c["value"]
                    for c in cookies
                    if c["name"] in target_cookies
                }
                if required_cookies.issubset(cookie_dict.keys()):
                    _log("Detected required cookies, waiting 2s for SSO cookies to settle...")
                    time.sleep(2)  # 给 SSO 重定向链上的所有 cookie 落盘时间
                    # 重新读取以获取最新值
                    cookies = context.cookies()
                    cookie_dict = {
                        c["name"]: c["value"]
                        for c in cookies
                        if c["name"] in target_cookies
                    }
                    _log("Saving cookies...")
                    save_cookies(platform, cookie_dict)
                    return True
                time.sleep(2)
                elapsed = int(time.time() - start_time)
                if elapsed % 30 == 0 and elapsed > 0:
                    _log(f"Still waiting... ({elapsed}s / {timeout}s)")

            _log(f"Timeout after {timeout}s.")
            return False
    finally:
        # 显式关闭以确保 user_data_dir 中所有 cookies 落盘
        if context is not None:
            try:
                context.close()
            except Exception:
                pass


def main():
    import argparse
    parser = argparse.ArgumentParser(description="自动刷新网易问卷系统 Cookie")
    parser.add_argument("--timeout", type=int, default=300, help="等待登录超时（秒，默认300）")
    parser.add_argument(
        "--platform", choices=["cn", "global"], default="cn",
        help="平台: cn=国内(163.com), global=国外(easebar.com)（默认 cn）",
    )
    parser.add_argument(
        "--browser", default=None,
        help="优先使用的浏览器: chrome/edge/chromium（缺省则按 Chrome→Edge→内置 Chromium 顺序自动挑可用的）",
    )
    args = parser.parse_args()

    success = refresh_cookie(platform=args.platform, timeout=args.timeout, preferred_browser=args.browser)
    if success:
        _log("✓ Cookie refresh completed!")
        print(json.dumps({"status": "success", "message": "Cookie 已自动刷新"}, ensure_ascii=False))
    else:
        _log("× Cookie refresh failed.")
        print(json.dumps({"status": "error", "message": "Cookie 刷新失败"}, ensure_ascii=False))
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()