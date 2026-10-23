// ==UserScript==
// @name         DeepSeek 对话归档 · Word 导出
// @namespace    https://chat.deepseek.com/
// @version      1.0.0
// @description  把 DeepSeek 网页端的会话（你问的 + AI 答的）完整导出为 Word (.docx)，支持全部会话 / 当前会话 / 合并总文档
// @match        https://chat.deepseek.com/*
// @run-at       document-idle
// @grant        none
// @noframes
// ==/UserScript==

/*
 * 说明：
 * 1) 本文件既是油猴脚本，也可以整份复制到 chat.deepseek.com 的浏览器控制台里直接运行。
 * 2) 不依赖任何第三方库：ZIP 打包与 .docx（OOXML）文件都是在本地逐字节拼出来的。
 * 3) 所有数据只在本机浏览器里流转，不会上传到任何服务器。
 */
(function () {
  'use strict';

  const VERSION = '1.0.0';
  const PAGE_SIZE = 50;
  const API = '/api/v0';

  /* =========================================================================
   * 一、基础工具
   * ====================================================================== */

  const pad2 = (n) => String(n).padStart(2, '0');

  function fmtTime(sec) {
    if (!sec && sec !== 0) return '';
    let ms = Number(sec);
    if (!isFinite(ms) || ms <= 0) return '';
    if (ms < 1e11) ms *= 1000;                 // 秒 → 毫秒
    const d = new Date(ms);
    if (isNaN(d.getTime())) return '';
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
  }

  function fmtDay(sec) {
    const t = fmtTime(sec);
    return t ? t.slice(0, 10) : '未知日期';
  }

  function stamp(d) {
    d = d || new Date();
    return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}_${pad2(d.getHours())}${pad2(d.getMinutes())}`;
  }

  function safeName(s, max) {
    max = max || 60;
    let t = String(s == null ? '' : s)
      .replace(/[\u0000-\u001f]/g, ' ')
      .replace(/[\\/:*?"<>|]/g, '_')
      .replace(/[\u200b-\u200f\ufeff]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/^[.\s]+/, '');
    if (t.length > max) t = t.slice(0, max).trim();
    if (!t) t = '未命名会话';
    if (/^(CON|PRN|AUX|NUL|COM\d|LPT\d)$/i.test(t)) t = '_' + t;
    return t;
  }

  const isCJK = (ch) => /[\u2e80-\u9fff\u3000-\u303f\uff00-\uffef]/.test(ch);

  function joinSoftLines(lines) {
    let out = '';
    for (const raw of lines) {
      const ln = raw.replace(/\s+$/, '');
      if (!out) { out = ln; continue; }
      const a = out.slice(-1), b = ln.charAt(0);
      out += (isCJK(a) || isCJK(b)) ? ln : ' ' + ln;
    }
    return out;
  }

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  /* =========================================================================
   * 二、ZIP 打包（store 模式，无需压缩库）
   * ====================================================================== */

  const enc = new TextEncoder();
  const utf8 = (s) => enc.encode(s);

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(buf) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function dosTime(d) {
    const time = (d.getHours() << 11) | (d.getMinutes() << 5) | (Math.floor(d.getSeconds() / 2) & 0x1F);
    const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    return { time: time & 0xFFFF, date: date & 0xFFFF };
  }

  /** entries: [{name:string, bytes:Uint8Array}] → Uint8Array（合法 ZIP） */
  function buildZip(entries) {
    const now = new Date();
    const dt = dosTime(now);
    const items = entries.map((e) => ({
      name: utf8(e.name),
      data: e.bytes instanceof Uint8Array ? e.bytes : utf8(String(e.bytes)),
      crc: crc32(e.bytes instanceof Uint8Array ? e.bytes : utf8(String(e.bytes))),
    }));

    let total = 22;
    for (const it of items) total += 30 + it.name.length + it.data.length + 46 + it.name.length;

    const out = new Uint8Array(total);
    const dv = new DataView(out.buffer);
    let p = 0;
    const offsets = [];

    for (const it of items) {
      offsets.push(p);
      dv.setUint32(p, 0x04034b50, true);
      dv.setUint16(p + 4, 20, true);        // version needed
      dv.setUint16(p + 6, 0x0800, true);    // UTF-8 名称标志
      dv.setUint16(p + 8, 0, true);         // store
      dv.setUint16(p + 10, dt.time, true);
      dv.setUint16(p + 12, dt.date, true);
      dv.setUint32(p + 14, it.crc, true);
      dv.setUint32(p + 18, it.data.length, true);
      dv.setUint32(p + 22, it.data.length, true);
      dv.setUint16(p + 26, it.name.length, true);
      dv.setUint16(p + 28, 0, true);
      p += 30;
      out.set(it.name, p); p += it.name.length;
      out.set(it.data, p); p += it.data.length;
    }

    const cdStart = p;
    items.forEach((it, i) => {
      dv.setUint32(p, 0x02014b50, true);
      dv.setUint16(p + 4, 20, true);
      dv.setUint16(p + 6, 20, true);
      dv.setUint16(p + 8, 0x0800, true);
      dv.setUint16(p + 10, 0, true);
      dv.setUint16(p + 12, dt.time, true);
      dv.setUint16(p + 14, dt.date, true);
      dv.setUint32(p + 16, it.crc, true);
      dv.setUint32(p + 20, it.data.length, true);
      dv.setUint32(p + 24, it.data.length, true);
      dv.setUint16(p + 28, it.name.length, true);
      dv.setUint16(p + 30, 0, true);
      dv.setUint16(p + 32, 0, true);
      dv.setUint16(p + 34, 0, true);
      dv.setUint16(p + 36, 0, true);
      dv.setUint32(p + 38, 0, true);
      dv.setUint32(p + 42, offsets[i], true);
      p += 46;
      out.set(it.name, p); p += it.name.length;
    });

    dv.setUint32(p, 0x06054b50, true);
    dv.setUint16(p + 4, 0, true);
    dv.setUint16(p + 6, 0, true);
    dv.setUint16(p + 8, items.length, true);
    dv.setUint16(p + 10, items.length, true);
    dv.setUint32(p + 12, p - cdStart, true);
    dv.setUint32(p + 16, cdStart, true);
    dv.setUint16(p + 20, 0, true);
    return out;
  }

  /* =========================================================================
   * 三、OOXML 小工具
   * ====================================================================== */

  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /** XML 转义 + 清洗非法控制字符 */
  function x(s) {
    return String(s == null ? '' : s)
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  const PPR_SEQ = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'numPr', 'pBdr', 'shd', 'tabs',
    'spacing', 'ind', 'jc', 'outlineLvl', 'rPr'];
  const RPR_SEQ = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'strike', 'color', 'spacing',
    'sz', 'szCs', 'highlight', 'u', 'shd', 'vertAlign'];

  function ordered(obj, seq) {
    let s = '';
    for (const k of seq) if (obj[k]) s += obj[k];
    return s;
  }

  /**
   * 生成 <w:pPr>。
   * o: {style, keepNext, pageBreakBefore, border, shd, tabs, spacing, ind, jc, outline, rPr}
   */
  function pPr(o) {
    o = o || {};
    const map = {
      pStyle: o.style ? `<w:pStyle w:val="${o.style}"/>` : '',
      keepNext: o.keepNext ? '<w:keepNext/>' : '',
      keepLines: o.keepLines ? '<w:keepLines/>' : '',
      pageBreakBefore: o.pageBreakBefore ? '<w:pageBreakBefore/>' : '',
      pBdr: o.border || '',
      shd: o.shd || '',
      tabs: o.tabs || '',
      spacing: o.spacing ? `<w:spacing ${o.spacing}/>` : '',
      ind: o.ind ? `<w:ind ${o.ind}/>` : '',
      jc: o.jc ? `<w:jc w:val="${o.jc}"/>` : '',
      outlineLvl: o.outline != null ? `<w:outlineLvl w:val="${o.outline}"/>` : '',
      rPr: o.rPr || '',
    };
    const body = ordered(map, PPR_SEQ);
    return body ? `<w:pPr>${body}</w:pPr>` : '';
  }

  /**
   * 生成 <w:rPr>。
   * o: {b, i, strike, code, color, sz, underline, rStyle, shd, font}
   */
  function rPr(o) {
    o = o || {};
    const mono = o.code || o.font === 'mono';
    const map = {
      rStyle: o.rStyle ? `<w:rStyle w:val="${o.rStyle}"/>` : '',
      rFonts: mono
        ? '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="等线" w:cs="Consolas"/>'
        : (o.font ? `<w:rFonts w:ascii="${o.font}" w:hAnsi="${o.font}" w:eastAsia="微软雅黑"/>` : ''),
      b: o.b ? '<w:b/><w:bCs/>' : '',
      i: o.i ? '<w:i/><w:iCs/>' : '',
      strike: o.strike ? '<w:strike/>' : '',
      color: o.color ? `<w:color w:val="${o.color}"/>` : '',
      sz: o.sz ? `<w:sz w:val="${o.sz}"/><w:szCs w:val="${o.sz}"/>` : '',
      u: o.underline ? '<w:u w:val="single"/>' : '',
      shd: o.shd ? `<w:shd w:val="clear" w:color="auto" w:fill="${o.shd}"/>` : '',
    };
    const body = ordered(map, RPR_SEQ);
    return body ? `<w:rPr>${body}</w:rPr>` : '';
  }

  /** 生成一个文字 run（自动处理换行符与制表符） */
  function run(text, o) {
    if (text === '' || text == null) return '';
    const parts = String(text).replace(/\t/g, '    ').split('\n');
    let out = '';
    parts.forEach((part, idx) => {
      if (idx) out += '<w:br/>';
      if (part) out += `<w:t xml:space="preserve">${x(part)}</w:t>`;
    });
    return `<w:r>${rPr(o)}${out}</w:r>`;
  }

  /** 段落 */
  function para(children, o) {
    const body = Array.isArray(children) ? children.join('') : (children || '');
    return `<w:p>${pPr(o || {})}${body}</w:p>`;
  }

  function hrPara(opts) {
    opts = opts || {};
    return `<w:p>${pPr({
      border: `<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="6" w:color="${opts.color || 'D8DEE4'}"/></w:pBdr>`,
      spacing: opts.spacing || 'before="0" after="140" line="240" lineRule="auto"',
    })}</w:p>`;
  }

  function pageBreak() {
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  }

  function bookmark(id, name) {
    return `<w:bookmarkStart w:id="${id}" w:name="${name}"/><w:bookmarkEnd w:id="${id}"/>`;
  }

  /* =========================================================================
   * 四、Markdown → 块结构 → OOXML
   * ====================================================================== */

  const RE_HEADING = /^(#{1,6})\s+(.+?)\s*#*\s*$/;
  const RE_FENCE = /^\s{0,3}(```|~~~)\s*([\w+#.\-]*)\s*$/;
  const RE_HR = /^\s{0,3}([-*_])(?:\s*\1){2,}\s*$/;
  const RE_LIST = /^(\s*)([-*+]|\d{1,3}[.)])\s+(.*)$/;
  const RE_QUOTE = /^\s{0,3}>\s?/;
  const RE_TABLE_SEP = /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/;

  function splitTableRow(line) {
    let s = line.trim();
    if (s.startsWith('|')) s = s.slice(1);
    if (s.endsWith('|')) s = s.slice(0, -1);
    const cells = [];
    let cur = '';
    for (let i = 0; i < s.length; i++) {
      if (s[i] === '\\' && s[i + 1] === '|') { cur += '|'; i++; continue; }
      if (s[i] === '|') { cells.push(cur.trim()); cur = ''; continue; }
      cur += s[i];
    }
    cells.push(cur.trim());
    return cells;
  }

  /** 把 Markdown 文本解析成块数组 */
  function parseBlocks(md) {
    const lines = String(md == null ? '' : md).replace(/\r\n?/g, '\n').split('\n');
    const blocks = [];
    let i = 0;
    let buf = [];
    const indentStack = [];

    const flush = () => {
      if (buf.length) {
        blocks.push({ type: 'p', text: joinSoftLines(buf) });
        buf = [];
      }
    };

    while (i < lines.length) {
      const line = lines[i];

      // 代码块
      let m = RE_FENCE.exec(line);
      if (m) {
        flush();
        const fence = m[1];
        const lang = m[2] || '';
        const code = [];
        i++;
        while (i < lines.length && !new RegExp('^\\s{0,3}' + fence).test(lines[i])) {
          code.push(lines[i]);
          i++;
        }
        i++; // 跳过结束围栏
        blocks.push({ type: 'code', lang, text: code.join('\n') });
        continue;
      }

      // 空行
      if (!line.trim()) {
        flush();
        indentStack.length = 0;
        i++;
        continue;
      }

      // 标题
      m = RE_HEADING.exec(line);
      if (m) {
        flush();
        indentStack.length = 0;
        blocks.push({ type: 'h', level: m[1].length, text: m[2].trim() });
        i++;
        continue;
      }

      // 分隔线
      if (RE_HR.test(line)) {
        flush();
        indentStack.length = 0;
        blocks.push({ type: 'hr' });
        i++;
        continue;
      }

      // 表格
      if (line.includes('|') && i + 1 < lines.length && RE_TABLE_SEP.test(lines[i + 1]) && lines[i + 1].includes('-')) {
        flush();
        indentStack.length = 0;
        const header = splitTableRow(line);
        i += 2;
        const rows = [];
        while (i < lines.length && lines[i].includes('|') && lines[i].trim()) {
          rows.push(splitTableRow(lines[i]));
          i++;
        }
        blocks.push({ type: 'table', header, rows });
        continue;
      }

      // 引用
      if (RE_QUOTE.test(line)) {
        flush();
        indentStack.length = 0;
        const q = [];
        while (i < lines.length && (RE_QUOTE.test(lines[i]) || (lines[i].trim() && q.length && !RE_HEADING.test(lines[i]) && !RE_FENCE.test(lines[i])))) {
          q.push(lines[i].replace(RE_QUOTE, ''));
          i++;
        }
        blocks.push({ type: 'quote', text: joinSoftLines(q).trim() });
        continue;
      }

      // 列表
      m = RE_LIST.exec(line);
      if (m) {
        flush();
        const indent = m[1].replace(/\t/g, '  ').length;
        // 用缩进栈推算层级，兼容 2 空格 / 4 空格两种写法
        while (indentStack.length && indent < indentStack[indentStack.length - 1]) indentStack.pop();
        if (!indentStack.length || indent > indentStack[indentStack.length - 1]) indentStack.push(indent);
        const level = Math.min(4, Math.max(0, indentStack.length - 1));
        const marker = m[2];
        let text = m[3];
        const ordered = /\d/.test(marker);
        const task = /^\[([ xX])\]\s+/.exec(text);
        if (task) text = text.replace(/^\[[ xX]\]\s+/, '');

        // 续行（缩进的后续行并入本条）
        i++;
        const cont = [];
        while (i < lines.length && lines[i].trim() && /^\s{2,}\S/.test(lines[i]) && !RE_LIST.test(lines[i])) {
          cont.push(lines[i].trim());
          i++;
        }
        blocks.push({
          type: 'li', ordered, level, marker,
          task: task ? (task[1].toLowerCase() === 'x') : null,
          text: joinSoftLines([text].concat(cont)),
        });
        continue;
      }

      // 数学公式块
      if (/^\s*\$\$/.test(line)) {
        flush();
        const math = [];
        let closed = false;
        let first = line.replace(/^\s*\$\$/, '');
        if (/\$\$\s*$/.test(first)) { math.push(first.replace(/\$\$\s*$/, '')); i++; closed = true; }
        if (!closed) {
          if (first.trim()) math.push(first);
          i++;
          while (i < lines.length) {
            if (/\$\$\s*$/.test(lines[i])) { math.push(lines[i].replace(/\$\$\s*$/, '')); i++; closed = true; break; }
            math.push(lines[i]); i++;
          }
        }
        blocks.push({ type: 'math', text: math.join('\n').trim() });
        continue;
      }

      // 普通段落
      buf.push(line.replace(/^\s+/, ''));
      i++;
    }
    flush();
    return blocks;
  }

  /**
   * 行内解析：**粗体** *斜体* ~~删除~~ `代码` [链接](url) 裸链接 $公式$
   * 返回 [{text, b, i, strike, code, link, math}]
   */
  function parseInline(text, st) {
    st = st || {};
    const out = [];
    let buf = '';
    let i = 0;
    const flush = () => {
      if (buf) { out.push(Object.assign({}, st, { text: buf })); buf = ''; }
    };
    const push = (t, extra) => out.push(Object.assign({}, st, extra || {}, { text: t }));

    while (i < text.length) {
      const rest = text.slice(i);
      const ch = text[i];

      // 转义
      if (ch === '\\' && i + 1 < text.length && /[\\`*_{}\[\]()#+\-.!~>|$]/.test(text[i + 1])) {
        buf += text[i + 1]; i += 2; continue;
      }

      // 行内公式 $$..$$ / $..$
      if (ch === '$') {
        const mm = /^\$\$?([^$]+?)\$\$?/.exec(rest);
        if (mm && mm[1].trim()) { flush(); push(mm[1], { code: true, math: true }); i += mm[0].length; continue; }
      }

      // 链接 [文字](地址)
      const lm = /^\[([^\]\n]*)\]\(\s*<?([^)>\s]+)>?(?:\s+["'][^"']*["'])?\s*\)/.exec(rest);
      if (lm) {
        flush();
        const inner = parseInline(lm[1], Object.assign({}, st, { link: lm[2] }));
        if (inner.length) out.push(...inner); else push(lm[1], { link: lm[2] });
        i += lm[0].length;
        continue;
      }

      // 裸链接
      const am = /^https?:\/\/[^\s<>"'）】\]]+/.exec(rest);
      if (am && !st.link) { flush(); push(am[0].replace(/[.,;:]+$/, ''), { link: am[0].replace(/[.,;:]+$/, '') }); i += am[0].replace(/[.,;:]+$/, '').length; continue; }

      // 粗体
      if (rest.startsWith('**') || rest.startsWith('__')) {
        const mk = rest.slice(0, 2);
        const end = text.indexOf(mk, i + 2);
        if (end > i + 2) {
          flush();
          out.push(...parseInline(text.slice(i + 2, end), Object.assign({}, st, { b: true })));
          i = end + 2;
          continue;
        }
      }

      // 删除线
      if (rest.startsWith('~~')) {
        const end = text.indexOf('~~', i + 2);
        if (end > i + 2) {
          flush();
          out.push(...parseInline(text.slice(i + 2, end), Object.assign({}, st, { strike: true })));
          i = end + 2;
          continue;
        }
      }

      // 行内代码
      if (ch === '`') {
        const mm = /^(`+)([\s\S]*?[^`])\1(?!`)/.exec(rest);
        if (mm) { flush(); push(mm[2].trim() || mm[2], { code: true }); i += mm[0].length; continue; }
      }

      // 斜体
      if (ch === '*' || ch === '_') {
        if (ch === '*' && text[i + 1] === '*') { buf += ch; i++; continue; }
        const end = text.indexOf(ch, i + 1);
        if (end > i + 1 && end - i < 400) {
          const inner = text.slice(i + 1, end);
          if (inner.trim() && !/^\s/.test(inner) && !/\s$/.test(inner)) {
            flush();
            out.push(...parseInline(inner, Object.assign({}, st, { i: true })));
            i = end + 1;
            continue;
          }
        }
      }

      buf += ch;
      i++;
    }
    flush();
    return out;
  }

  const TEX_SYMBOLS = {
    times: '×', div: '÷', cdot: '·', pm: '±', mp: '∓',
    leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠', approx: '≈', equiv: '≡',
    infty: '∞', to: '→', rightarrow: '→', Rightarrow: '⇒', leftarrow: '←',
    sum: 'Σ', prod: '∏', int: '∫', partial: '∂', nabla: '∇',
    in: '∈', notin: '∉', subset: '⊂', cup: '∪', cap: '∩', forall: '∀', exists: '∃', emptyset: '∅',
    sqrt: '√', angle: '∠', degree: '°',
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', theta: 'θ', lambda: 'λ',
    mu: 'μ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', phi: 'φ', omega: 'ω',
    Delta: 'Δ', Sigma: 'Σ', Omega: 'Ω', Gamma: 'Γ', Theta: 'Θ', Lambda: 'Λ',
    log: 'log', ln: 'ln', exp: 'exp', sin: 'sin', cos: 'cos', tan: 'tan',
    max: 'max', min: 'min', lim: 'lim',
  };

  const TEX_FUNCS = new Set(['log', 'ln', 'exp', 'sin', 'cos', 'tan', 'max', 'min', 'lim']);

  /** 把常见 LaTeX 记号转成可读的 Unicode 文本（Word 里不排真公式，只求能读） */
  function deTeX(s) {
    return String(s)
      .replace(/\\frac\s*\{([^{}]*)\}\s*\{([^{}]*)\}/g, '($1)/($2)')
      .replace(/\\sqrt\s*\{([^{}]*)\}/g, '√($1)')
      .replace(/\\left\s*|\\right\s*/g, '')
      .replace(/\\[,;:!]/g, ' ')
      .replace(/\\quad|\\qquad/g, '  ')
      .replace(/\\([A-Za-z]+)/g, (m, name, offset, whole) => {
        const sym = TEX_SYMBOLS[name] !== undefined ? TEX_SYMBOLS[name] : name;
        // 函数名前面是字母/数字/右括号时补一个空格，避免 "nlog n"
        if (TEX_FUNCS.has(name) && offset > 0 && /[A-Za-z0-9)\]]/.test(whole[offset - 1])) return ' ' + sym;
        return sym;
      })
      .replace(/\\([{}%_])/g, '$1')
      .replace(/[{}]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /** 把行内 token 渲染成 run 串，处理超链接 */
  function renderInline(text, ctx, base) {
    const tokens = parseInline(String(text == null ? '' : text));
    let out = '';
    for (const t of tokens) {
      const o = Object.assign({}, base || {});
      if (t.b) o.b = true;
      if (t.i) o.i = true;
      if (t.strike) o.strike = true;
      if (t.code) { o.code = true; o.shd = 'EFF2F5'; o.sz = (base && base.sz) || 19; }
      if (t.link) {
        o.color = '0B5CAD';
        o.underline = true;
        out += hyperlinkRun(t.text, t.link, ctx, o);
        continue;
      }
      out += run(t.math ? deTeX(t.text) : t.text, o);
    }
    return out;
  }

  /** 超链接：注册关系 id 并输出 <w:hyperlink> */
  function hyperlinkRun(text, url, ctx, o) {
    const clean = String(url).trim().replace(/[\s"'<>]/g, (c) => encodeURI(c));
    if (!clean) return run(text, o);
    const id = ctx.linkFor(clean);
    return `<w:hyperlink r:id="${id}" w:history="1">${run(text, o)}</w:hyperlink>`;
  }

  function codeBlockPara(text) {
    const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
    const out = [];
    const CHUNK = 200;
    for (let start = 0; start < lines.length; start += CHUNK) {
      const slice = lines.slice(start, start + CHUNK);
      let body = '';
      slice.forEach((ln, idx) => {
        if (idx) body += '<w:br/>';
        const s = ln.replace(/\t/g, '    ');
        body += s ? `<w:t xml:space="preserve">${x(s)}</w:t>` : '<w:t xml:space="preserve"> </w:t>';
      });
      out.push(`<w:p>${pPr({ style: 'CodeBlock' })}<w:r>${rPr({})}${body}</w:r></w:p>`);
    }
    return out.join('');
  }

  function mathBlockPara(text) {
    return `<w:p>${pPr({ style: 'MathBlock' })}${run(deTeX(String(text).replace(/\n/g, ' ')), { code: true, sz: 19 })}</w:p>`;
  }

  function tableXml(header, rows, ctx) {
    const n = Math.max(header.length, ...rows.map((r) => r.length), 1);
    const totalW = 9026;
    const colW = Math.floor(totalW / n);
    const grid = Array.from({ length: n }, () => `<w:gridCol w:w="${colW}"/>`).join('');
    const border = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="D5DBE1"/>`;

    const cell = (txt, isHead) => {
      const inner = isHead
        ? run(String(txt || ''), { b: true, sz: 19 })
        : renderInline(txt || '', ctx, { sz: 19 });
      return `<w:tc><w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/>` +
        (isHead ? '<w:shd w:val="clear" w:color="auto" w:fill="F2F4F7"/>' : '') +
        `<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="110" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="110" w:type="dxa"/></w:tcMar>` +
        `<w:vAlign w:val="center"/></w:tcPr>` +
        `<w:p>${pPr({ style: 'TableText' })}${inner}</w:p></w:tc>`;
    };

    const headRow = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${header.concat(Array(Math.max(0, n - header.length)).fill('')).map((c) => cell(c, true)).join('')}</w:tr>`;
    const bodyRows = rows.map((r) => `<w:tr>${r.concat(Array(Math.max(0, n - r.length)).fill('')).map((c) => cell(c, false)).join('')}</w:tr>`).join('');

    return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>` +
      `<w:tblBorders>${border('top')}${border('left')}${border('bottom')}${border('right')}${border('insideH')}${border('insideV')}</w:tblBorders>` +
      `<w:tblLayout w:type="autofit"/><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblGrid>${grid}</w:tblGrid>${headRow}${bodyRows}</w:tbl>` +
      `<w:p>${pPr({ spacing: 'before="0" after="80" line="120" lineRule="auto"' })}</w:p>`;
  }

  const HEADING_STYLE = { 1: 'Heading1', 2: 'Heading2', 3: 'Heading3', 4: 'Heading4', 5: 'Heading4', 6: 'Heading4' };

  function renderBlocks(md, ctx, opts) {
    opts = opts || {};
    const blocks = parseBlocks(md);
    const base = opts.base || { sz: 21 };
    let out = '';

    for (const b of blocks) {
      switch (b.type) {
        case 'h': {
          const lvl = Math.min(6, Math.max(1, b.level));
          const bmId = ctx.nextBm();
          out += `<w:p>${pPr({ style: HEADING_STYLE[lvl], outline: lvl - 1, keepNext: true })}` +
            `<w:bookmarkStart w:id="${bmId}" w:name="_Toc${bmId}"/>` +
            renderInline(b.text, ctx, { b: true }) +
            `<w:bookmarkEnd w:id="${bmId}"/></w:p>`;
          break;
        }
        case 'p':
          out += para(renderInline(b.text, ctx, base), { spacing: 'before="0" after="120" line="330" lineRule="auto"' });
          break;
        case 'hr':
          out += hrPara({ color: 'E4E8ED' });
          break;
        case 'quote':
          out += para(renderInline(b.text, ctx, base), { style: 'Quote' });
          break;
        case 'code':
          out += codeBlockPara(b.text);
          break;
        case 'math':
          out += mathBlockPara(b.text);
          break;
        case 'table':
          out += tableXml(b.header, b.rows, ctx);
          break;
        case 'li': {
          const lvl = b.level || 0;
          const bullet = ['•', '–', '▪'][lvl % 3];
          let prefix;
          if (b.task != null) prefix = b.task ? '☑ ' : '☐ ';
          else if (b.ordered) prefix = String(b.marker).replace(/[.)]$/, '') + '. ';
          else prefix = bullet + ' ';
          const left = 360 + lvl * 360;
          out += `<w:p>${pPr({
            spacing: 'before="0" after="60" line="320" lineRule="auto"',
            ind: `w:left="${left}" w:hanging="300"`,
          })}${run(prefix, { b: b.ordered, color: b.ordered ? '30363D' : '57606A' })}${renderInline(b.text, ctx, base)}</w:p>`;
          break;
        }
        default:
          break;
      }
    }
    return out;
  }

  /* =========================================================================
   * 五、文档模板（styles / settings / header / footer …）
   * ====================================================================== */

  const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  const CONTENT_TYPES = XML_HEAD + `
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

  const ROOT_RELS = XML_HEAD + `
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

  const APP_XML = XML_HEAD + `
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>DeepSeek 对话归档</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>
<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion></Properties>`;

  const SETTINGS_XML = XML_HEAD + `
<w:settings xmlns:w="${W_NS}">
<w:zoom w:percent="100"/>
<w:defaultTabStop w:val="420"/>
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/>
<w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/>
</w:compat>
</w:settings>`;

  const STYLES_XML = XML_HEAD + `
<w:styles xmlns:w="${W_NS}">
<w:docDefaults>
<w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="微软雅黑" w:cs="Calibri"/>
<w:color w:val="1F2328"/><w:sz w:val="21"/><w:szCs w:val="21"/>
<w:lang w:val="en-US" w:eastAsia="zh-CN"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:line="320" w:lineRule="auto" w:after="120"/><w:jc w:val="left"/></w:pPr></w:pPrDefault>
</w:docDefaults>

<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>

<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:spacing w:before="0" w:after="80" w:line="280" w:lineRule="auto"/><w:outlineLvl w:val="0"/></w:pPr>
<w:rPr><w:b/><w:color w:val="111827"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:keepNext/><w:spacing w:before="320" w:after="140" w:line="300" w:lineRule="auto"/><w:outlineLvl w:val="0"/></w:pPr>
<w:rPr><w:b/><w:color w:val="111827"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:keepNext/><w:spacing w:before="280" w:after="100" w:line="300" w:lineRule="auto"/><w:outlineLvl w:val="1"/></w:pPr>
<w:rPr><w:b/><w:color w:val="1F2328"/><w:sz w:val="27"/><w:szCs w:val="27"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:keepNext/><w:spacing w:before="220" w:after="80" w:line="300" w:lineRule="auto"/><w:outlineLvl w:val="2"/></w:pPr>
<w:rPr><w:b/><w:color w:val="1F2328"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:keepNext/><w:spacing w:before="180" w:after="60" w:line="300" w:lineRule="auto"/><w:outlineLvl w:val="3"/></w:pPr>
<w:rPr><w:b/><w:color w:val="30363D"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="Meta"><w:name w:val="Meta"/><w:basedOn w:val="Normal"/>
<w:pPr><w:spacing w:before="0" w:after="40" w:line="260" w:lineRule="auto"/></w:pPr>
<w:rPr><w:color w:val="6E7781"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>
<w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="D0D7DE"/></w:pBdr>
<w:shd w:val="clear" w:color="auto" w:fill="F7F8FA"/>
<w:spacing w:before="80" w:after="140" w:line="300" w:lineRule="auto"/>
<w:ind w:left="360" w:right="180"/></w:pPr>
<w:rPr><w:color w:val="4B5563"/><w:sz w:val="20"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="Think"><w:name w:val="Think"/><w:basedOn w:val="Normal"/>
<w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="E3B341"/></w:pBdr>
<w:shd w:val="clear" w:color="auto" w:fill="FCFAF4"/>
<w:spacing w:before="100" w:after="160" w:line="280" w:lineRule="auto"/>
<w:ind w:left="300" w:right="180"/></w:pPr>
<w:rPr><w:color w:val="6E7781"/><w:sz w:val="19"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code Block"/><w:basedOn w:val="Normal"/>
<w:pPr><w:pBdr>
<w:top w:val="single" w:sz="4" w:space="6" w:color="E2E6EA"/>
<w:left w:val="single" w:sz="4" w:space="6" w:color="E2E6EA"/>
<w:bottom w:val="single" w:sz="4" w:space="6" w:color="E2E6EA"/>
<w:right w:val="single" w:sz="4" w:space="6" w:color="E2E6EA"/></w:pBdr>
<w:shd w:val="clear" w:color="auto" w:fill="F6F8FA"/>
<w:spacing w:before="80" w:after="180" w:line="250" w:lineRule="auto"/>
<w:ind w:left="120" w:right="120"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="等线" w:cs="Consolas"/>
<w:color w:val="24292F"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="MathBlock"><w:name w:val="Math Block"/><w:basedOn w:val="Normal"/>
<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F8F8FA"/><w:spacing w:before="80" w:after="160" w:line="280" w:lineRule="auto"/><w:ind w:left="240"/></w:pPr>
<w:rPr><w:i/><w:color w:val="1F2328"/><w:sz w:val="19"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="Table Text"/><w:basedOn w:val="Normal"/>
<w:pPr><w:spacing w:before="0" w:after="0" w:line="260" w:lineRule="auto"/></w:pPr>
<w:rPr><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/>
<w:pPr><w:spacing w:before="0" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr>
<w:rPr><w:color w:val="8B949E"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="TocEntry"><w:name w:val="TOC Entry"/><w:basedOn w:val="Normal"/>
<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9026"/></w:tabs>
<w:spacing w:before="0" w:after="80" w:line="280" w:lineRule="auto"/></w:pPr>
<w:rPr><w:sz w:val="20"/></w:rPr></w:style>

<w:style w:type="character" w:styleId="CodeInline"><w:name w:val="Code Inline"/>
<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="等线"/>
<w:color w:val="24292F"/><w:sz w:val="19"/><w:shd w:val="clear" w:color="auto" w:fill="EFF2F5"/></w:rPr></w:style>
</w:styles>`;

  function coreXml(title) {
    const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
    return XML_HEAD + `
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>${x(title)}</dc:title>
<dc:subject>DeepSeek 对话归档</dc:subject>
<dc:creator>DeepSeek 对话归档工具 v${VERSION}</dc:creator>
<cp:lastModifiedBy>DeepSeek 对话归档工具 v${VERSION}</cp:lastModifiedBy>
<cp:revision>1</cp:revision>
<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
  }

  function headerXml(left, right) {
    const tail = right
      ? `<w:r><w:tab/></w:r><w:r><w:rPr><w:color w:val="9AA4AF"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">${x(right)}</w:t></w:r>`
      : '';
    return XML_HEAD + `
<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}">
<w:p><w:pPr><w:pStyle w:val="Caption"/>
<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="4" w:color="D8DEE4"/></w:pBdr>
<w:tabs><w:tab w:val="right" w:pos="9026"/></w:tabs>
<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
<w:r><w:rPr><w:color w:val="9AA4AF"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">${x(left)}</w:t></w:r>${tail}
</w:p></w:hdr>`;
  }

  const FOOTER_XML = XML_HEAD + `
<w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}">
<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
<w:r><w:rPr><w:color w:val="9AA4AF"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">第 </w:t></w:r>
<w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:color w:val="9AA4AF"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>
<w:r><w:rPr><w:color w:val="9AA4AF"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve"> 页 / 共 </w:t></w:r>
<w:fldSimple w:instr=" NUMPAGES "><w:r><w:rPr><w:color w:val="9AA4AF"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>
<w:r><w:rPr><w:color w:val="9AA4AF"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve"> 页</w:t></w:r>
</w:p></w:ftr>`;

  function documentXml(bodyXml, ctx) {
    return XML_HEAD + `
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<w:body>${bodyXml}
<w:sectPr>
<w:headerReference w:type="default" r:id="rIdHeader1"/>
<w:footerReference w:type="default" r:id="rIdFooter1"/>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1304" w:right="1247" w:bottom="1247" w:left="1247" w:header="851" w:footer="794" w:gutter="0"/>
<w:cols w:space="425"/>
</w:sectPr>
</w:body></w:document>`;
  }

  function documentRels(ctx) {
    let rels = `<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>` +
      `<Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`;
    for (const [url, id] of ctx.links) {
      rels += `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${x(url)}" TargetMode="External"/>`;
    }
    return XML_HEAD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
  }

  /** 组装 docx 字节流 */
  function packDocx(opts) {
    const ctx = opts.ctx;
    const entries = [
      { name: '[Content_Types].xml', bytes: utf8(CONTENT_TYPES) },
      { name: '_rels/.rels', bytes: utf8(ROOT_RELS) },
      { name: 'docProps/core.xml', bytes: utf8(coreXml(opts.title)) },
      { name: 'docProps/app.xml', bytes: utf8(APP_XML) },
      { name: 'word/document.xml', bytes: utf8(documentXml(opts.body, ctx)) },
      { name: 'word/styles.xml', bytes: utf8(STYLES_XML) },
      { name: 'word/settings.xml', bytes: utf8(SETTINGS_XML) },
      { name: 'word/header1.xml', bytes: utf8(headerXml(opts.headerLeft, opts.headerRight)) },
      { name: 'word/footer1.xml', bytes: utf8(FOOTER_XML) },
      { name: 'word/_rels/document.xml.rels', bytes: utf8(documentRels(ctx)) },
    ];
    return buildZip(entries);
  }

  /* =========================================================================
   * 六、会话 → Word 文档
   * ====================================================================== */

  function makeCtx() {
    const links = new Map();
    let bm = 1000;
    return {
      links,
      get bmId() { return bm; },
      nextBm() { return ++bm; },
      linkFor(url) {
        if (links.has(url)) return links.get(url);
        const id = 'rIdLink' + (links.size + 1);
        links.set(url, id);
        return id;
      },
    };
  }

  function roleBand(msg) {
    const isUser = msg.role === 'user';
    const label = isUser ? '你' : 'DeepSeek';
    const fill = isUser ? 'EAF2FE' : 'F1F2F4';
    const color = isUser ? '1A56B0' : '24292F';
    return `<w:p>${pPr({
      keepNext: true,
      shd: `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`,
      spacing: 'before="300" after="140" line="280" lineRule="auto"',
    })}${run(label, { b: true, color, sz: 21 })}${run('    ' + fmtTime(msg.time), { color: '8B949E', sz: 17 })}</w:p>`;
  }

  /** 单条消息：角色条 → 思考过程 → 正文 → 附件 */
  function renderMessage(msg, ctx, opts) {
    let out = roleBand(msg);

    if (opts.think && msg.think) {
      out += `<w:p>${pPr({ style: 'Think' })}${run('思考过程', { b: true, sz: 17, color: '9A6700' })}</w:p>`;
      out += `<w:p>${pPr({ style: 'Think' })}${run(msg.think, { sz: 19 })}</w:p>`;
    }

    const body = (msg.content || '').trim();
    out += body
      ? renderBlocks(body, ctx, opts)
      : para(run('（此条内容为空）', { sz: 19, color: '9AA4AF', i: true }), { spacing: 'before="0" after="120"' });

    if (msg.files && msg.files.length) {
      out += para(run('附件：' + msg.files.join('、'), { sz: 17, color: '6E7781' }),
        { style: 'Meta', spacing: 'before="100" after="220" line="260" lineRule="auto"' });
    }
    return out;
  }

  function sessionMeta(sess) {
    const parts = [];
    if (sess.created_at) parts.push('创建 ' + fmtTime(sess.created_at));
    if (sess.updated_at) parts.push('更新 ' + fmtTime(sess.updated_at));
    parts.push(`共 ${sess.messages.length} 条消息`);
    if (sess.model_type) parts.push(sess.model_type === 'expert' ? '深度思考模式' : '快速模式');
    let out = para(run(parts.join('　·　'), { sz: 17, color: '6E7781' }), { style: 'Meta' });
    out += para(run('会话 ID：' + sess.id, { sz: 15, color: '9AA4AF' }), { style: 'Caption' });
    return out;
  }

  function renderSessionBody(sess, ctx, opts) {
    let out = '';
    out += para(run(sess.title, { b: true }), { style: 'Title' });
    out += sessionMeta(sess);
    out += hrPara({ color: 'D8DEE4' });

    for (const msg of sess.messages) out += renderMessage(msg, ctx, opts);
    return out;
  }

  /** 单个会话 → docx 字节 */
  function buildDocx(session, opts) {
    opts = opts || {};
    const ctx = makeCtx();
    const body = renderSessionBody(session, ctx, opts);
    return packDocx({
      title: session.title,
      headerLeft: `DeepSeek 对话归档 · ${session.title}`,
      headerRight: '导出于 ' + fmtTime(Date.now()),
      body,
      ctx,
    });
  }

  /** 多个会话 → 一份总文档 */
  function buildMergedDocx(sessions, opts) {
    opts = opts || {};
    const ctx = makeCtx();
    const now = Date.now();
    const totalMsgs = sessions.reduce((n, s) => n + s.messages.length, 0);

    let cover = '';
    // 空段落撑高（Word 会忽略页面顶端的段前距，所以用固定行高占位）
    cover += `<w:p>${pPr({ spacing: 'before="0" after="0" line="2600" lineRule="exact"' })}</w:p>`;
    cover += para(run('DeepSeek 对话归档', { b: true }), { style: 'Title', spacing: 'before="0" after="120" line="300" lineRule="auto"' });
    cover += para(run('C O N V E R S A T I O N   A R C H I V E', { sz: 16, color: '8B949E' }), { style: 'Meta', spacing: 'before="0" after="240" line="240" lineRule="auto"' });
    cover += hrPara({ color: 'D8DEE4', spacing: 'before="120" after="320" line="240" lineRule="auto"' });
    cover += para(run('导出时间', { sz: 17, color: '6E7781' }), { style: 'Meta' });
    cover += para(run(fmtTime(now), { sz: 30, b: true, color: '1F2328' }), { style: 'Meta', spacing: 'before="0" after="280" line="280" lineRule="auto"' });
    cover += para(run('会话数量', { sz: 17, color: '6E7781' }), { style: 'Meta' });
    cover += para(run(`${sessions.length} 个`, { sz: 30, b: true, color: '1F2328' }), { style: 'Meta', spacing: 'before="0" after="280" line="280" lineRule="auto"' });
    cover += para(run('消息总数', { sz: 17, color: '6E7781' }), { style: 'Meta' });
    cover += para(run(`${totalMsgs} 条`, { sz: 30, b: true, color: '1F2328' }), { style: 'Meta', spacing: 'before="0" after="280" line="280" lineRule="auto"' });
    if (opts.sourceUrl) cover += para(run(opts.sourceUrl, { sz: 16, color: '9AA4AF' }), { style: 'Caption', spacing: 'before="120" after="0" line="240" lineRule="auto"' });

    // 目录
    let toc = pageBreak();
    toc += para(run('目录', { b: true }), { style: 'Heading1', pageBreakBefore: true });
    sessions.forEach((s, idx) => {
      const anchor = '_TocS' + (idx + 1);
      ctx.nextBm();
      toc += `<w:p>${pPr({ style: 'TocEntry' })}` +
        `<w:hyperlink w:anchor="${anchor}" w:history="1">${run(String(idx + 1).padStart(2, '0') + '　' + s.title, { color: '0B5CAD' })}</w:hyperlink>` +
        `<w:r>${rPr({ sz: 17, color: '8B949E' })}<w:tab/></w:r>` +
        run(fmtDay(s.updated_at || s.created_at) + '　·　' + s.messages.length + ' 条', { sz: 17, color: '8B949E' }) +
        `</w:p>`;
    });

    let body = cover + toc;
    sessions.forEach((s, idx) => {
      const anchor = '_TocS' + (idx + 1);
      const id = ctx.nextBm();
      body += `<w:p>${pPr({ style: 'Heading1', pageBreakBefore: true })}` +
        `<w:bookmarkStart w:id="${id}" w:name="${anchor}"/>` +
        run(String(idx + 1).padStart(2, '0') + '　' + s.title, { b: true }) +
        `<w:bookmarkEnd w:id="${id}"/></w:p>`;
      body += sessionMeta(s);
      body += hrPara({ color: 'D8DEE4' });
      for (const msg of s.messages) body += renderMessage(msg, ctx, opts);
    });

    return packDocx({
      title: 'DeepSeek 对话归档',
      headerLeft: 'DeepSeek 对话归档 · 全部会话',
      headerRight: '导出于 ' + fmtTime(now),
      body,
      ctx,
    });
  }

  /* =========================================================================
   * 七、数据获取与归一化
   * ====================================================================== */

  function pick(obj, path) {
    let cur = obj;
    for (const k of path.split('.')) {
      if (cur == null) return undefined;
      cur = cur[k];
    }
    return cur;
  }

  function readToken() {
    const extract = (raw) => {
      if (!raw) return '';
      let v = String(raw).trim();
      if (v.startsWith('"') && v.endsWith('"')) v = v.slice(1, -1);
      if (v.startsWith('{')) {
        try {
          const j = JSON.parse(v);
          return j.value || j.token || j.access_token || '';
        } catch (e) { return ''; }
      }
      return v;
    };
    const keys = ['userToken', 'user_token', 'token', 'access_token'];
    for (const k of keys) {
      const t = extract(localStorage.getItem(k));
      if (t && t.length > 20) return t;
    }
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i) || '';
      if (/token/i.test(k)) {
        const t = extract(localStorage.getItem(k));
        if (t && t.length > 40) return t;
      }
    }
    return '';
  }

  function authHeaders() {
    const h = {
      'x-client-locale': 'zh_CN',
      'x-client-platform': 'web',
      'x-app-version': '2.0.0',
      'x-client-version': '2.0.0',
    };
    const t = readToken();
    if (t) h['Authorization'] = /^Bearer\s/i.test(t) ? t : 'Bearer ' + t;
    return h;
  }

  async function apiGet(path, params) {
    const url = new URL(API + path, location.origin);
    if (params) {
      Object.keys(params).forEach((k) => {
        if (params[k] !== undefined && params[k] !== null && params[k] !== '') url.searchParams.set(k, params[k]);
      });
    }
    let res;
    try {
      res = await fetch(url.toString(), {
        method: 'GET',
        headers: authHeaders(),
        credentials: 'include',
        cache: 'no-store',
      });
    } catch (e) {
      throw new Error('网络请求失败：' + e.message);
    }
    const text = await res.text();
    let json = null;
    try { json = JSON.parse(text); } catch (e) { /* 非 JSON */ }

    if (res.status === 401 || res.status === 403) {
      throw new Error('登录状态失效（HTTP ' + res.status + '），请刷新页面重新登录后再试');
    }
    if (!res.ok) {
      throw new Error('接口返回 HTTP ' + res.status + (text ? '：' + text.slice(0, 120) : ''));
    }
    if (!json) {
      throw new Error('接口返回的不是 JSON，可能被风控拦截。请刷新页面后再试');
    }
    const code = json.code != null ? json.code : (json.data && json.data.biz_code);
    if (code !== undefined && code !== 0 && code !== null) {
      const msg = json.msg || pick(json, 'data.biz_msg') || ('code ' + code);
      throw new Error('接口错误：' + msg);
    }
    return json;
  }

  /** 拉取全部会话（含置顶），按更新时间倒序 */
  async function listSessions(onProgress) {
    const out = [];
    const seen = new Set();

    for (const pinned of [true, false]) {
      let cursor = null;
      let idle = 0;
      for (let guard = 0; guard < 400; guard++) {
        const params = { 'lte_cursor.pinned': String(pinned) };
        if (cursor != null) params['lte_cursor.updated_at'] = String(cursor);
        let json;
        try {
          json = await apiGet('/chat_session/fetch_page', params);
        } catch (e) {
          // 置顶那一路失败不影响主流程；已有数据时也先继续
          if (pinned || out.length) break;
          throw e;
        }
        const items = pick(json, 'data.biz_data.chat_sessions') || pick(json, 'data.biz_data.sessions') || [];
        if (!items.length) break;

        let added = 0;
        for (const it of items) {
          if (it && it.id && !seen.has(it.id)) { seen.add(it.id); out.push(it); added++; }
        }
        if (onProgress) onProgress(out.length);
        if (!added) { if (++idle >= 2) break; } else idle = 0;

        const last = items[items.length - 1];
        const next = last && last.updated_at;
        if (next == null || next === cursor) break;
        cursor = next;
      }
    }

    out.sort((a, b) => (b.updated_at || b.inserted_at || 0) - (a.updated_at || a.inserted_at || 0));
    return out;
  }

  async function fetchSessionMessages(id) {
    const json = await apiGet('/chat/history_messages', { chat_session_id: id });
    return pick(json, 'data.biz_data') || null;
  }

  function normalizeMessage(m) {
    if (!m) return null;
    const roleRaw = String(m.role || m.author || '').toUpperCase();
    const role = roleRaw === 'USER' ? 'user' : (roleRaw === 'ASSISTANT' || roleRaw === 'BOT' ? 'assistant' : '');
    if (!role) return null;

    const frags = Array.isArray(m.fragments) ? m.fragments : [];
    let contentParts = [];
    const thinks = [];
    const files = [];

    for (const f of frags) {
      if (!f) continue;
      const t = String(f.type || '').toUpperCase();
      let c = f.content;
      if (c && typeof c === 'object') c = c.content != null ? c.content : JSON.stringify(c);
      c = typeof c === 'string' ? c : '';
      if (t === 'REQUEST' || t === 'RESPONSE' || t === 'TEXT' || t === 'CONTENT') contentParts.push(c);
      else if (t === 'THINK' || t === 'REASONING') thinks.push(c);
      else if (t === 'FILE' || t === 'DOCUMENT') {
        let name = '';
        try {
          const j = JSON.parse(c);
          name = j.name || j.file_name || j.filename || (j.file && j.file.name) || '';
        } catch (e) { /* ignore */ }
        if (name) files.push(name);
      }
    }

    let content = contentParts.filter((s) => s && s.trim()).join('\n\n');
    if (!content && typeof m.content === 'string') content = m.content;
    if (!content && m.content && typeof m.content === 'object') content = String(m.content.content || '');
    if (!content && typeof m.text === 'string') content = m.text;

    return {
      role,
      time: m.inserted_at || m.created_at || m.updated_at || null,
      content: content || '',
      think: thinks.filter((s) => s && s.trim()).join('\n\n'),
      files,
      status: m.status || '',
    };
  }

  function normalizeSession(session, bizData) {
    const info = Object.assign({}, session || {}, (bizData && bizData.chat_session) || {});
    const raw = (bizData && (bizData.chat_messages || bizData.messages)) || [];
    const messages = [];
    for (const m of raw) {
      const nm = normalizeMessage(m);
      if (!nm) continue;
      if (nm.status === 'WIP' && !nm.content && !nm.think) continue;
      if (!nm.content && !nm.think && !nm.files.length) continue;
      messages.push(nm);
    }
    return {
      id: info.id || (session && session.id) || '',
      title: String(info.title || '').trim() || '未命名会话',
      created_at: info.inserted_at || (session && session.inserted_at) || null,
      updated_at: info.updated_at || (session && session.updated_at) || null,
      model_type: info.model_type || '',
      pinned: !!(session && session.pinned),
      messages,
    };
  }

  /** 会话 → 可序列化的 JSON 结构（原始备份用） */
  function toJsonPayload(sessions, sourceUrl) {
    return {
      tool: 'deepseek-archiver',
      version: VERSION,
      source: sourceUrl || 'https://chat.deepseek.com',
      exported_at: new Date().toISOString(),
      count: sessions.length,
      message_count: sessions.reduce((n, s) => n + s.messages.length, 0),
      sessions,
    };
  }

  if (typeof document === 'undefined') {
    /* Node 测试入口：无 DOM 时只导出核心函数 */
    globalThis.__DSARCHIVE__ = {
      VERSION, buildDocx, buildMergedDocx, buildZip, packDocx,
      normalizeSession, normalizeMessage, parseBlocks, parseInline,
      renderBlocks, safeName, fmtTime, toJsonPayload, makeCtx,
    };
    return;
  }

  /* =========================================================================
   * 八、界面
   * ====================================================================== */

  const PANEL_ID = 'ds-archive-panel';
  const state = { busy: false, cancel: false, ui: null };

  const CSS = `
  :host { all: initial; }
  .card {
    position: fixed; right: 18px; bottom: 18px; width: 346px;
    font: 13px/1.55 "Segoe UI", "Microsoft YaHei", system-ui, -apple-system, sans-serif;
    color: #e6edf3; background: #161b22;
    border: 1px solid #30363d; border-radius: 12px;
    box-shadow: 0 12px 40px rgba(0,0,0,.45);
    z-index: 2147483000; overflow: hidden; user-select: none;
  }
  .head { display: flex; align-items: center; gap: 8px; padding: 10px 12px;
    background: linear-gradient(180deg,#1c2330,#161b22); cursor: move;
    border-bottom: 1px solid #21262d; }
  .head b { font-size: 13px; font-weight: 600; letter-spacing: .3px; }
  .dot { width: 8px; height: 8px; border-radius: 50%; background: #3fb950; box-shadow: 0 0 8px #3fb95088; }
  .ver { color: #6e7781; font-size: 11px; margin-left: auto; }
  .mini { background: transparent; border: 0; color: #8b949e; cursor: pointer; font-size: 14px;
    padding: 2px 6px; border-radius: 6px; line-height: 1; }
  .mini:hover { background: #21262d; color: #e6edf3; }
  .body { padding: 12px; display: flex; flex-direction: column; gap: 10px; }
  .hint { font-size: 12px; color: #8b949e; }
  .hint.ok { color: #3fb950; }
  .hint.warn { color: #d29922; }
  .hint.err { color: #f85149; }
  .row { display: flex; gap: 8px; }
  button.act { flex: 1; border: 0; border-radius: 8px; padding: 8px 10px; font-size: 12.5px;
    font-weight: 600; cursor: pointer; font-family: inherit; transition: background .15s, opacity .15s; }
  button.primary { background: #1f6feb; color: #fff; }
  button.primary:hover { background: #388bfd; }
  button.ghost { background: #21262d; color: #c9d1d9; border: 1px solid #30363d; }
  button.ghost:hover { background: #30363d; }
  button.act:disabled { opacity: .45; cursor: not-allowed; }
  button.danger { background: #b62324; color: #fff; flex: 0 0 auto; padding: 8px 12px; }
  .opts { display: flex; flex-direction: column; gap: 6px; font-size: 12px; color: #adbac7;
    background: #0d1117; border: 1px solid #21262d; border-radius: 8px; padding: 9px 10px; }
  .opts label { display: flex; align-items: center; gap: 6px; cursor: pointer; }
  .opts input[type=checkbox] { accent-color: #1f6feb; cursor: pointer; margin: 0; }
  .opts input[type=number] { width: 58px; background: #161b22; color: #e6edf3; border: 1px solid #30363d;
    border-radius: 5px; padding: 2px 6px; font-size: 12px; font-family: inherit; }
  .bar { height: 5px; background: #21262d; border-radius: 3px; overflow: hidden; }
  .fill { height: 100%; width: 0; background: linear-gradient(90deg,#1f6feb,#3fb950); transition: width .18s; }
  .status { font-size: 12px; color: #8b949e; min-height: 16px; word-break: break-all; }
  .log { max-height: 92px; overflow-y: auto; font-size: 11.5px; line-height: 1.6; color: #7d8590;
    background: #0d1117; border: 1px solid #21262d; border-radius: 8px; padding: 7px 9px;
    font-family: "Cascadia Mono", Consolas, monospace; white-space: pre-wrap; }
  .log:empty { display: none; }
  .log .ok { color: #3fb950; }
  .log .er { color: #f85149; }
  .foot { font-size: 11px; color: #565f6b; text-align: right; }
  .foot .link { color: #58a6ff; cursor: pointer; text-decoration: underline dotted; }
  .foot .link:hover { color: #79c0ff; }
  .collapsed .body { display: none; }
  `;

  function panelHtml() {
    return `
  <div class="card" id="card">
    <div class="head">
      <span class="dot"></span><b>DeepSeek 归档 · Word</b>
      <span class="ver">v${VERSION}</span>
      <button class="mini" id="mini" title="收起/展开">—</button>
    </div>
    <div class="body">
      <div class="hint" id="hint">正在检查登录状态…</div>
      <div class="row">
        <button class="act primary" id="btnAll">导出全部会话</button>
        <button class="act ghost" id="btnCur">当前会话</button>
        <button class="act danger" id="btnCancel" style="display:none">取消</button>
      </div>
      <div class="opts">
        <label><input type="checkbox" id="optMerge"> 额外生成一份「合并总文档」</label>
        <label><input type="checkbox" id="optThink"> 保留 AI 的思考过程</label>
        <label><input type="checkbox" id="optJson" checked> 附带 JSON 原始备份</label>
        <label>只导出最近 <input type="number" id="optLimit" min="0" step="10" placeholder="全部"> 个会话</label>
      </div>
      <div class="bar"><div class="fill" id="fill"></div></div>
      <div class="status" id="status">准备就绪</div>
      <div class="log" id="log"></div>
      <div class="foot"><span class="link" id="diag">接口诊断</span> · 数据仅在本机处理</div>
    </div>
  </div>`;
  }

  function buildPanel() {
    if (document.getElementById(PANEL_ID)) return state.ui;
    const host = document.createElement('div');
    host.id = PANEL_ID;
    const root = host.attachShadow({ mode: 'open' });
    root.innerHTML = `<style>${CSS}</style>${panelHtml()}`;
    document.body.appendChild(host);

    const $ = (id) => root.getElementById(id);
    const ui = {
      host, root,
      hint: $('hint'), status: $('status'), log: $('log'), fill: $('fill'),
      btnAll: $('btnAll'), btnCur: $('btnCur'), btnCancel: $('btnCancel'),
      optMerge: $('optMerge'), optThink: $('optThink'), optJson: $('optJson'), optLimit: $('optLimit'),
    };
    state.ui = ui;

    ui.btnAll.addEventListener('click', () => runExport('all'));
    ui.btnCur.addEventListener('click', () => runExport('current'));
    ui.btnCancel.addEventListener('click', () => {
      state.cancel = true;
      ui.status.textContent = '正在取消…';
    });
    $('mini').addEventListener('click', () => {
      root.getElementById('card').classList.toggle('collapsed');
    });
    $('diag').addEventListener('click', diagnose);

    // 拖动
    const card = root.getElementById('card');
    const head = root.querySelector('.head');
    let drag = null;
    head.addEventListener('pointerdown', (e) => {
      if (e.target.closest('.mini')) return;
      const r = card.getBoundingClientRect();
      drag = { x: e.clientX, y: e.clientY, left: r.left, top: r.top };
      card.style.right = 'auto'; card.style.bottom = 'auto';
      card.style.left = r.left + 'px'; card.style.top = r.top + 'px';
      head.setPointerCapture(e.pointerId);
    });
    head.addEventListener('pointermove', (e) => {
      if (!drag) return;
      const w = card.offsetWidth, h = card.offsetHeight;
      let left = drag.left + e.clientX - drag.x;
      let top = drag.top + e.clientY - drag.y;
      left = Math.max(4, Math.min(window.innerWidth - w - 4, left));
      top = Math.max(4, Math.min(window.innerHeight - h - 4, top));
      card.style.left = left + 'px';
      card.style.top = top + 'px';
    });
    head.addEventListener('pointerup', () => { drag = null; });

    refreshLoginHint();
    return ui;
  }

  function refreshLoginHint() {
    const ui = state.ui;
    if (!ui) return;
    const token = readToken();
    ui.hint.className = 'hint ' + (token ? 'ok' : 'warn');
    ui.hint.textContent = token ? '已检测到登录信息，可以开始导出' : '未检测到登录令牌；若导出失败请刷新页面重新登录';
  }

  function logLine(msg, kind) {
    const ui = state.ui;
    if (!ui) return;
    const line = document.createElement('div');
    if (kind) line.className = kind;
    line.textContent = msg;
    ui.log.appendChild(line);
    while (ui.log.childNodes.length > 200) ui.log.removeChild(ui.log.firstChild);
    ui.log.scrollTop = ui.log.scrollHeight;
  }

  const setStatus = (t) => { if (state.ui) state.ui.status.textContent = t; };

  function setProgress(done, total) {
    if (!state.ui) return;
    const pct = total ? Math.round((done / total) * 100) : 0;
    state.ui.fill.style.width = pct + '%';
    if (total) state.ui.status.textContent = `${done} / ${total}（${pct}%）`;
  }

  function busyUI(busy) {
    const ui = state.ui;
    if (!ui) return;
    state.busy = busy;
    ui.btnAll.disabled = busy;
    ui.btnCur.disabled = busy;
    ui.btnCancel.style.display = busy ? '' : 'none';
    if (!busy) ui.btnCancel.disabled = false;
  }

  function readOpts() {
    const ui = state.ui;
    const raw = ui.optLimit.value.trim();
    const limit = raw ? Math.max(0, parseInt(raw, 10) || 0) : 0;
    return {
      merge: ui.optMerge.checked,
      think: ui.optThink.checked,
      json: ui.optJson.checked,
      limit,
      sourceUrl: 'https://chat.deepseek.com',
    };
  }

  function currentSessionId() {
    const m = /\/s\/([0-9a-zA-Z-]{6,})/.exec(location.pathname + location.hash) ||
      /[?&]session_id=([0-9a-zA-Z-]{6,})/.exec(location.search);
    return m ? m[1] : '';
  }

  function download(bytes, filename, mime) {
    const blob = new Blob([bytes], { type: mime || 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 8000);
  }

  function humanSize(n) {
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
    return (n / 1048576).toFixed(1) + ' MB';
  }

  const README_TXT = (sessions, opts) =>
    [
      'DeepSeek 对话归档',
      '==================',
      '',
      `导出时间：${fmtTime(Date.now())}`,
      `会话数量：${sessions.length}`,
      `消息条数：${sessions.reduce((n, s) => n + s.messages.length, 0)}`,
      `工具版本：v${VERSION}`,
      '',
      '目录说明：',
      '  会话/              每个会话一份 Word 文档（.docx），按时间从早到晚编号',
      '  全部会话合集.docx   所有会话合并成的一份文档（含目录，可点击跳转）',
      '  原始数据.json      结构化原始数据，可用于二次处理或重新生成',
      '',
      '说明：Word 文档可直接用 Microsoft Word / WPS / LibreOffice 打开。',
      '      代码块、表格、标题层级、超链接均已尽量保留。',
      '      数学公式按 LaTeX 源码形式保留（未做公式排版）。',
    ].join('\r\n');

  /* ---------------------------- 接口诊断 ---------------------------- */

  async function diagnose() {
    if (state.busy) return;
    const ui = state.ui;
    ui.log.innerHTML = '';
    const token = readToken();
    logLine('开始诊断 v' + VERSION);
    logLine('token：' + (token ? token.slice(0, 10) + '…（长度 ' + token.length + '）' : '未找到'));
    logLine('页面：' + location.pathname);

    try {
      const r1 = await fetch(new URL(API + '/chat_session/fetch_page?lte_cursor.pinned=false', location.origin), {
        headers: authHeaders(), credentials: 'include', cache: 'no-store',
      });
      logLine('会话列表接口 HTTP ' + r1.status);
      const t1 = await r1.text();
      logLine('响应长度 ' + t1.length);
      let j1 = null;
      try { j1 = JSON.parse(t1); } catch (e) { logLine('返回的不是 JSON：' + t1.slice(0, 180), 'er'); return; }
      logLine('顶层字段：' + Object.keys(j1).join(', '));
      const bd = pick(j1, 'data.biz_data') || {};
      logLine('biz_data 字段：' + Object.keys(bd).join(', '));
      const list = bd.chat_sessions || [];
      logLine('会话数：' + list.length, 'ok');
      if (!list.length) { logLine('列表为空，诊断结束'); return; }
      logLine('会话字段：' + Object.keys(list[0]).join(', '));
      logLine('示例：' + (list[0].title || '(无标题)') + ' / ' + list[0].id);

      const r2 = await fetch(new URL(API + '/chat/history_messages?chat_session_id=' + list[0].id, location.origin), {
        headers: authHeaders(), credentials: 'include', cache: 'no-store',
      });
      logLine('消息接口 HTTP ' + r2.status);
      const j2 = await r2.json().catch(() => null);
      if (!j2) { logLine('消息接口返回非 JSON', 'er'); return; }
      const bd2 = pick(j2, 'data.biz_data') || {};
      logLine('biz_data 字段：' + Object.keys(bd2).join(', '));
      const msgs = bd2.chat_messages || [];
      logLine('消息数：' + msgs.length, 'ok');
      if (msgs[0]) {
        logLine('消息字段：' + Object.keys(msgs[0]).join(', '));
        logLine('role：' + msgs[0].role + ' / fragments：' +
          (msgs[0].fragments || []).map((f) => f.type).join(' | '));
        const nm = normalizeMessage(msgs[0]);
        logLine('归一化结果：' + (nm ? nm.role + ' / ' + (nm.content || '').length + ' 字' : '失败'));
      }
      logLine('诊断结束', 'ok');
    } catch (e) {
      logLine('诊断失败：' + e.message, 'er');
      console.error('[DeepSeek 归档]', e);
    }
  }

  /* ---------------------------- 主流程 ---------------------------- */

  async function runExport(mode) {
    if (state.busy) return;
    const ui = state.ui;
    const opts = readOpts();
    state.cancel = false;
    ui.log.innerHTML = '';
    busyUI(true);
    setProgress(0, 0);

    const t0 = Date.now();
    try {
      let sessions = [];

      if (mode === 'current') {
        const id = currentSessionId();
        if (!id) throw new Error('当前页面不是某个具体会话，请先点开一个会话');
        setStatus('读取当前会话…');
        logLine('读取会话 ' + id);
        const data = await fetchSessionMessages(id);
        const s = normalizeSession({ id }, data);
        if (!s.messages.length) throw new Error('这个会话里没有可导出的消息');
        s.title = s.title === '未命名会话' ? (document.title || '').replace(/\s*-\s*DeepSeek.*$/, '') || '未命名会话' : s.title;
        sessions = [s];
        logLine(`已读取 ${s.messages.length} 条消息`, 'ok');
        setProgress(1, 1);
      } else {
        setStatus('读取会话列表…');
        logLine('正在读取会话列表…');
        const list = await listSessions((n) => {
          if (state.ui) state.ui.status.textContent = `已发现 ${n} 个会话…`;
        });
        logLine(`会话列表：共 ${list.length} 个`, 'ok');
        if (!list.length) throw new Error('没有获取到任何会话，请确认账号里确实有对话记录');

        let picked = list;
        if (opts.limit > 0) picked = list.slice(0, opts.limit);

        let okCount = 0, failCount = 0;
        for (let i = 0; i < picked.length; i++) {
          if (state.cancel) throw new Error('已取消');
          const info = picked[i];
          setStatus(`正在导出 ${i + 1}/${picked.length}：${info.title || info.id}`);
          try {
            const data = await fetchSessionMessages(info.id);
            const s = normalizeSession(info, data);
            if (s.messages.length) {
              sessions.push(s);
              okCount++;
            } else {
              logLine(`跳过（无消息）：${info.title || info.id}`);
            }
          } catch (e) {
            failCount++;
            logLine(`失败：${info.title || info.id} — ${e.message}`, 'er');
          }
          setProgress(i + 1, picked.length);
        }
        logLine(`读取完成：成功 ${okCount}，失败 ${failCount}`, failCount ? 'er' : 'ok');
        if (!sessions.length) throw new Error('所有会话都读取失败了，请刷新页面重试');
      }

      // 按时间正序，方便编号与阅读
      sessions.sort((a, b) => (a.created_at || 0) - (b.created_at || 0));

      // 生成 Word
      const entries = [];
      const sessFiles = [];
      for (let i = 0; i < sessions.length; i++) {
        if (state.cancel) throw new Error('已取消');
        const s = sessions[i];
        const name = `${String(i + 1).padStart(3, '0')}_${safeName(s.title)}.docx`;
        setStatus(`正在生成 Word ${i + 1}/${sessions.length}：${s.title}`);
        const bytes = buildDocx(s, { think: opts.think, sourceUrl: opts.sourceUrl });
        sessFiles.push({ name, bytes });
        setProgress(i + 1, sessions.length);
      }

      let merged = null;
      if (opts.merge && sessions.length > 1) {
        setStatus('正在生成合并总文档…');
        merged = buildMergedDocx(sessions, { think: opts.think, sourceUrl: opts.sourceUrl });
      }

      let jsonBytes = null;
      if (opts.json) {
        const payload = toJsonPayload(sessions, opts.sourceUrl);
        jsonBytes = utf8(JSON.stringify(payload, null, 2));
      }

      const stampStr = stamp(new Date());

      if (sessions.length === 1 && !opts.merge) {
        // 单个会话直接给 .docx
        const f = sessFiles[0];
        download(f.bytes, `${safeName(sessions[0].title)}.docx`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        logLine(`已下载：${f.name}（${humanSize(f.bytes.length)}）`, 'ok');
        setStatus(`完成 · 1 个会话 · ${humanSize(f.bytes.length)}`);
      } else {
        entries.push({ name: '归档说明.txt', bytes: utf8(README_TXT(sessions, opts)) });
        if (merged) entries.push({ name: '全部会话合集.docx', bytes: merged });
        for (const f of sessFiles) entries.push({ name: '会话/' + f.name, bytes: f.bytes });
        if (jsonBytes) entries.push({ name: '原始数据.json', bytes: jsonBytes });

        setStatus('正在打包…');
        const zip = buildZip(entries);
        const zipName = `DeepSeek归档_${stampStr}.zip`;
        download(zip, zipName, 'application/zip');
        logLine(`已下载：${zipName}（${humanSize(zip.length)}）`, 'ok');
        setStatus(`完成 · ${sessions.length} 个会话 · ${humanSize(zip.length)}`);
      }

      const secs = ((Date.now() - t0) / 1000).toFixed(1);
      logLine(`全部完成，用时 ${secs}s`, 'ok');
      if (state.ui) state.ui.fill.style.width = '100%';
    } catch (e) {
      setStatus('出错：' + e.message);
      logLine('错误：' + e.message, 'er');
      console.error('[DeepSeek 归档]', e);
    } finally {
      busyUI(false);
    }
  }

  /* ---------------------------- 启动 ---------------------------- */

  function boot() {
    if (!document.body) return setTimeout(boot, 300);
    buildPanel();
    let lastPath = location.pathname;
    setInterval(() => {
      if (location.pathname !== lastPath) {
        lastPath = location.pathname;
        refreshLoginHint();
      }
    }, 1500);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }

  window.__deepseekArchiver = { version: VERSION, runExport, buildDocx, buildMergedDocx };
})();
