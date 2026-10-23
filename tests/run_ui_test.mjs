/**
 * 用最小 DOM 桩（stub）在 Node 里真正跑一遍用户脚本的「面板 + 导出」全流程。
 * 目的：验证 UI 层没有拼写错误 / 未定义引用，导出流程能走到“完成”。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const src = fs.readFileSync(path.join(here, '..', 'deepseek-archiver.user.js'), 'utf8');

const downloads = [];
const errors = [];

class El {
  constructor(tag, id) {
    this.tagName = String(tag || 'div').toUpperCase();
    this.id = id || '';
    this.children = [];
    this.style = {};
    this.checked = false;
    this.value = '';
    this.disabled = false;
    this.scrollTop = 0;
    this.scrollHeight = 0;
    this.offsetWidth = 340;
    this.offsetHeight = 420;
    this._text = '';
    this._html = '';
    this._listeners = {};
    this.classList = {
      _s: new Set(),
      add(c) { this._s.add(c); },
      remove(c) { this._s.delete(c); },
      toggle(c) { this._s.has(c) ? this._s.delete(c) : this._s.add(c); },
    };
  }
  addEventListener(t, fn) { (this._listeners[t] = this._listeners[t] || []).push(fn); }
  appendChild(c) { this.children.push(c); return c; }
  removeChild(c) { const i = this.children.indexOf(c); if (i >= 0) this.children.splice(i, 1); }
  remove() { this.removed = true; }
  get childNodes() { return this.children; }
  get firstChild() { return this.children[0]; }
  set innerHTML(v) { this._html = String(v); this.children = []; }
  get innerHTML() { return this._html; }
  set textContent(v) { this._text = String(v); this.children = []; }
  get textContent() { return this._text || this.children.map((c) => c.textContent).join(''); }
  set className(v) { this._class = v; }
  get className() { return this._class || ''; }
  click() {
    if (this.tagName === 'A') {
      downloads.push({ name: this.download, href: String(this.href || '').slice(0, 32) });
      return;
    }
    (this._listeners.click || []).forEach((f) => f({ target: this }));
  }
  setPointerCapture() {}
  getBoundingClientRect() { return { left: 100, top: 100, width: 340, height: 420 }; }
  closest() { return null; }
  attachShadow() {
    const els = new Map();
    const root = {
      _html: '',
      set innerHTML(v) {
        this._html = String(v);
        for (const m of String(v).matchAll(/id="([^"]+)"/g)) els.set(m[1], new El('div', m[1]));
      },
      get innerHTML() { return this._html; },
      getElementById: (id) => els.get(id) || null,
      querySelector: () => new El('div'),
    };
    this.shadowRoot = root;
    return root;
  }
}

const body = new El('body');
const byId = new Map();
globalThis.window = globalThis;
globalThis.document = {
  readyState: 'complete',
  title: 'DeepSeek - 测试',
  body,
  createElement: (t) => new El(t),
  getElementById: (id) => byId.get(id) || null,
  addEventListener() {},
};
const realAppend = body.appendChild.bind(body);
body.appendChild = (el) => { if (el.id) byId.set(el.id, el); return realAppend(el); };

globalThis.localStorage = (() => {
  const store = { userToken: JSON.stringify({ value: 'fake-token-0123456789abcdef0123456789abcdef' }) };
  return {
    getItem: (k) => (k in store ? store[k] : null),
    setItem: (k, v) => { store[k] = String(v); },
    key: (i) => Object.keys(store)[i],
    get length() { return Object.keys(store).length; },
  };
})();

globalThis.location = { origin: 'https://chat.deepseek.com', pathname: '/a/chat/s/sess-0001', hash: '', search: '' };

const sessions = [
  { id: 'sess-0001', title: '模拟会话 A / 含特殊字符', pinned: false, inserted_at: 1759000000, updated_at: 1759600000 },
  { id: 'sess-0002', title: '模拟会话 B', pinned: true, inserted_at: 1759200000, updated_at: 1759700000, model_type: 'expert' },
];
const msgs = [
  { role: 'USER', inserted_at: 1759000000, status: 'FINISHED', fragments: [{ type: 'REQUEST', content: '你好，给我 **加粗** 和 `行内代码`。' }] },
  { role: 'ASSISTANT', inserted_at: 1759000030, status: 'FINISHED', fragments: [
    { type: 'THINK', content: '用户想要格式示例。' },
    { type: 'RESPONSE', content: '## 标题二\n\n1. 一\n2. 二\n\n```js\nconst a = 1;\n```\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n> 引用\n\n见 [文档](https://example.com)。' },
  ] },
];

const requests = [];
globalThis.fetch = async (input) => {
  const url = String(input);
  requests.push(url);
  const mk = (obj) => ({ ok: true, status: 200, text: async () => JSON.stringify(obj) });
  if (url.includes('/chat_session/fetch_page')) {
    const pinned = /pinned=true/.test(url);
    return mk({ code: 0, data: { biz_code: 0, biz_data: { chat_sessions: sessions.filter((s) => (pinned ? s.pinned : !s.pinned)) } } });
  }
  if (url.includes('/chat/history_messages')) {
    const id = /chat_session_id=([^&]+)/.exec(url)[1];
    const info = sessions.find((s) => s.id === id) || { id };
    return mk({ code: 0, data: { biz_code: 0, biz_data: { chat_session: info, chat_messages: msgs } } });
  }
  return { ok: false, status: 404, text: async () => '{}' };
};

process.on('unhandledRejection', (e) => errors.push('REJECT: ' + (e && e.stack || e)));
process.on('uncaughtException', (e) => errors.push('THROW: ' + (e && e.stack || e)));

try {
  (0, eval)(src);
} catch (e) {
  errors.push('EVAL: ' + e.stack);
}

const host = byId.get('ds-archive-panel');
if (!host) {
  console.log('❌ 面板没有挂到 body 上');
  console.log(errors.join('\n'));
  process.exit(1);
}

const sr = host.shadowRoot;
const ui = {
  btnAll: sr.getElementById('btnAll'),
  status: sr.getElementById('status'),
  log: sr.getElementById('log'),
  fill: sr.getElementById('fill'),
  hint: sr.getElementById('hint'),
  optMerge: sr.getElementById('optMerge'),
  optThink: sr.getElementById('optThink'),
  optJson: sr.getElementById('optJson'),
  optLimit: sr.getElementById('optLimit'),
  btnCancel: sr.getElementById('btnCancel'),
};

console.log('面板元素齐全:', Object.values(ui).every(Boolean));
console.log('登录提示:', ui.hint.textContent);
console.log('初始状态:', ui.status.textContent);

ui.optMerge.checked = true;
ui.optThink.checked = true;
ui.btnAll.click();

const started = Date.now();
await new Promise((resolve) => {
  const timer = setInterval(() => {
    const s = ui.status.textContent || '';
    if (/完成|出错|错误/.test(s) || Date.now() - started > 15000) {
      clearInterval(timer);
      resolve();
    }
  }, 60);
});

console.log('最终状态:', ui.status.textContent);
console.log('进度条:', ui.fill.style.width);
console.log('下载:', JSON.stringify(downloads));
console.log('请求数:', requests.length);
console.log('日志:\n' + (ui.log.textContent || '(空)').split('\n').map((l) => '  ' + l).join('\n'));
if (errors.length) {
  console.log('运行错误:\n' + errors.join('\n'));
}
const ok = !errors.length && downloads.length > 0 && /完成/.test(ui.status.textContent);
console.log(ok ? '\n✅ UI 全流程通过' : '\n❌ UI 全流程失败');
process.exit(ok ? 0 : 1);
