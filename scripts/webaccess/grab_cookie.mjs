// grab_cookie.mjs —— 通过 CDP 从用户日常浏览器抓取网易问卷登录 cookie（含 HttpOnly）。
//
// 复用 web-access 的浏览器发现机制（browser-discovery.mjs），连接用户已登录的浏览器，
// 用 CDP `Storage.getCookies` 读取 cookie 库（可读 HttpOnly，document.cookie 读不到），
// 过滤出目标平台 cookie，以 JSON 打印到 stdout 供 Python 侧写入 config.json。
//
// 用法：
//   node grab_cookie.mjs --platform cn|global [--browser chrome|edge|...] [--list]
//
// stdout JSON status：
//   ok        —— 抓到全部必需 cookie，含 cookies:{name:value}
//   list      —— 仅 --list：列出检测到的浏览器
//   ambiguous —— 检测到多个浏览器且未指定 --browser，需上层询问用户
//   mismatch  —— 指定的 --browser 未检测到
//   empty     —— 没有任何浏览器开启远程调试
//   no_login  —— 浏览器已连上，但没有目标平台的必需 cookie（未登录问卷平台）
//   error     —— 连接/CDP 出错
// 退出码：ok/list=0，其余非 0（便于 Python 侧快速判定）。

import http from 'node:http';
import { detectAll, findFallbackPort } from './browser-discovery.mjs';

const PLATFORMS = {
  cn: {
    domainSuffixes: ['163.com'],
    target: ['SURVEY_TOKEN', 'JSESSIONID', 'P_INFO'],
    required: ['SURVEY_TOKEN', 'JSESSIONID'],
  },
  global: {
    domainSuffixes: ['easebar.com'],
    target: ['oversea-online_SURVEY_TOKEN', 'SURVEY_TOKEN', 'JSESSIONID', 'P_INFO'],
    required: ['oversea-online_SURVEY_TOKEN'],
  },
};

function out(obj, code = 0) {
  process.stdout.write(JSON.stringify(obj));
  process.exit(code);
}

function parseArgs(argv) {
  const a = { platform: 'cn', browser: null, list: false };
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    if (k === '--platform') a.platform = argv[++i];
    else if (k === '--browser') a.browser = argv[++i];
    else if (k === '--list') a.list = true;
  }
  return a;
}

// 无 wsPath 时（如手动指定端口），从 /json/version 拿 browser 级 webSocketDebuggerUrl
function fetchBrowserWsUrl(port) {
  return new Promise((resolve) => {
    const req = http.get({ host: '127.0.0.1', port, path: '/json/version', timeout: 3000 }, (res) => {
      let data = '';
      res.on('data', (c) => (data += c));
      res.on('end', () => {
        try { resolve(JSON.parse(data).webSocketDebuggerUrl || null); }
        catch { resolve(null); }
      });
    });
    req.on('error', () => resolve(null));
    req.on('timeout', () => { req.destroy(); resolve(null); });
  });
}

// 连接 CDP browser 端点，调用 Storage.getCookies，返回 cookie 数组
function getAllCookies(wsUrl, timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const done = (fn, arg) => { if (!settled) { settled = true; try { ws.close(); } catch {} fn(arg); } };
    const ws = new WebSocket(wsUrl);
    const timer = setTimeout(() => done(reject, new Error('CDP timeout')), timeoutMs);
    ws.onopen = () => ws.send(JSON.stringify({ id: 1, method: 'Storage.getCookies', params: {} }));
    ws.onerror = (e) => { clearTimeout(timer); done(reject, new Error('WS error: ' + (e?.message || 'unknown'))); };
    ws.onmessage = (ev) => {
      let msg;
      try { msg = JSON.parse(ev.data); } catch { return; }
      if (msg.id === 1) {
        clearTimeout(timer);
        if (msg.error) return done(reject, new Error(msg.error.message || 'CDP error'));
        done(resolve, (msg.result && msg.result.cookies) || []);
      }
    };
  });
}

async function grabFrom(browser, plat) {
  // 优先用 main() 校验阶段缓存的实时 ws 地址；缺失时再探测一次。
  // /json/version 是权威存活探测：拿不到 webSocketDebuggerUrl 即非真正 DevTools 端点
  // （DevToolsActivePort 里的 wsPath 可能残留 / 端口被其他进程占用，连它只会挂起超时）。
  const wsUrl = browser.wsUrl || await fetchBrowserWsUrl(browser.port);
  if (!wsUrl) {
    const err = new Error('not_devtools_endpoint');
    err.code = 'not_devtools';
    throw err;
  }
  const cookies = await getAllCookies(wsUrl);
  return pick(cookies, plat);
}

function pick(cookies, plat) {
  const picked = {};
  const found = [];
  for (const c of cookies) {
    const dom = (c.domain || '').replace(/^\./, '');
    const domOk = plat.domainSuffixes.some((s) => dom === s || dom.endsWith('.' + s));
    if (domOk && plat.target.includes(c.name)) {
      picked[c.name] = c.value;
      found.push(c.name);
    }
  }
  return { picked, found };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const plat = PLATFORMS[args.platform];
  if (!plat) out({ status: 'error', message: `unknown platform: ${args.platform}` }, 2);

  // detectAll 仅做 DevToolsActivePort 文件 + TCP 探测，可能假阳性（端口被别的进程占用）。
  // 这里再用 /json/version 逐个校验，只保留真正说 DevTools 协议的端点，并缓存实时 ws 地址。
  const candidates = await detectAll();
  const detected = [];
  for (const b of candidates) {
    const wsUrl = await fetchBrowserWsUrl(b.port);
    if (wsUrl) detected.push({ ...b, wsUrl });
  }

  if (args.list) {
    out({ status: 'list', browsers: detected.map((b) => ({ id: b.id, label: b.label, port: b.port })) }, 0);
  }

  // 选定浏览器
  let target = null;
  if (args.browser) {
    target = detected.find((b) => b.id === args.browser);
    if (!target) out({ status: 'mismatch', requested: args.browser, detected: detected.map((b) => b.id) }, 3);
  } else if (detected.length === 0) {
    const fp = await findFallbackPort();
    const wsUrl = fp ? await fetchBrowserWsUrl(fp) : null;
    if (wsUrl) target = { id: 'fallback', label: `port ${fp}`, port: fp, wsUrl };
    else out({ status: 'empty' }, 4);
  } else if (detected.length === 1) {
    target = detected[0];
  } else {
    out({ status: 'ambiguous', browsers: detected.map((b) => ({ id: b.id, label: b.label, port: b.port })) }, 5);
  }

  try {
    const { picked, found } = await grabFrom(target, plat);
    const missing = plat.required.filter((n) => !(n in picked));
    if (missing.length > 0) {
      out({ status: 'no_login', browser: { id: target.id, label: target.label }, found, missing }, 6);
    }
    out({ status: 'ok', browser: { id: target.id, label: target.label }, cookies: picked, count: found.length }, 0);
  } catch (e) {
    // 端口活着但不是真正的 DevTools 端点（端口被占用/残留），视为"无可用浏览器"，让上层回退 Playwright
    if (e && e.code === 'not_devtools') {
      out({ status: 'empty', reason: 'stale_or_occupied_port', port: target.port }, 4);
    }
    out({ status: 'error', browser: { id: target.id, label: target.label }, message: String(e?.message || e) }, 2);
  }
}

main();
