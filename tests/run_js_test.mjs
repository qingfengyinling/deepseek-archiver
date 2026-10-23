// 在 Node 里直接跑用户脚本的核心部分（无 DOM 时会自动走测试导出分支）
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const src = fs.readFileSync(path.join(here, '..', 'deepseek-archiver.user.js'), 'utf8');

// 用间接 eval 在全局作用域执行
(0, eval)(src);

const DS = globalThis.__DSARCHIVE__;
if (!DS) {
  console.error('核心函数未导出，脚本结构可能被改坏了');
  process.exit(1);
}

const LONG = `这是一段很长的中文回答，用来测试段落换行与中英文混排的效果。The quick brown fox jumps over the lazy dog, and the paragraph should wrap naturally inside the page margins without any manual line breaks. 中文与 English 混排时不应该被插入多余的空格，也不应该在词中间断开。`;

function md() {
  return [
    '## 一、先说结论',
    '',
    '**核心结论**：可以。`deepseek-chat` 与 `deepseek-reasoner` 都支持 *上下文缓存*，命中的部分按更低的单价计费。',
    '',
    LONG,
    '',
    '> 注意：缓存命中需要前缀完全一致，因此把固定的系统提示词放在最前面效果最好。',
    '',
    '### 1.1 计费对比',
    '',
    '| 项目 | 缓存命中 | 缓存未命中 | 说明 |',
    '| --- | ---: | ---: | :--- |',
    '| 输入（每百万 tokens） | 0.5 元 | 4 元 | 命中率越高越省 |',
    '| 输出（每百万 tokens） | 8 元 | 8 元 | 输出不参与缓存 |',
    '',
    '### 1.2 示例代码',
    '',
    '```python',
    'import openai',
    '',
    'client = openai.OpenAI(api_key="sk-xxx", base_url="https://api.deepseek.com")',
    '',
    'def ask(q: str) -> str:',
    '    resp = client.chat.completions.create(',
    '        model="deepseek-chat",',
    '        messages=[{"role": "user", "content": q}],',
    '    )',
    '    return resp.choices[0].message.content  # 返回文本',
    '```',
    '',
    '### 1.3 操作步骤',
    '',
    '1. 准备固定的前缀提示词；',
    '2. 把变动的部分放到最后；',
    '3. 观察返回里的 `prompt_cache_hit_tokens`。',
    '',
    '其中第一点最关键：',
    '',
    '- 不要在前面插入时间戳',
    '  - 时间戳会破坏前缀一致性',
    '  - 真的需要就放到末尾',
    '- 不要拼接随机 ID',
    '',
    '- [ ] 待办：统计一周的命中率',
    '- [x] 已完成：接入缓存统计',
    '',
    '---',
    '',
    '公式示例：$O(n\\log n)$，块级公式：',
    '',
    '$$',
    'T(n) = 2T(n/2) + O(n)',
    '$$',
    '',
    '参考链接：[官方文档](https://api-docs.deepseek.com/zh-cn/guides/kv_cache)，以及 https://platform.deepseek.com/usage 。',
    '',
    '#### 脚注式补充',
    '',
    '> 引用块第二段：~~旧的结论~~ 已经过时，请以新版为准。',
  ].join('\n');
}

const session = {
  id: 'a1b2c3d4-0000-4000-8000-000000000001',
  title: 'DeepSeek 上下文缓存怎么计费？（示例）',
  created_at: 1759000000,
  updated_at: 1759600000,
  model_type: 'default',
  pinned: false,
  messages: [
    { role: 'user', time: 1759000000, content: '帮我讲清楚 DeepSeek 的**上下文缓存**是怎么计费的，最好给个代码例子。', think: '', files: [] },
    { role: 'assistant', time: 1759000060, content: md(), think: '用户想要计费规则 + 代码示例，先给结论再展开。', files: [] },
    { role: 'user', time: 1759000200, content: '那我把系统提示词放在最前面就行了对吧？\n还有别的前提条件吗？', think: '', files: ['说明文档.pdf'] },
    { role: 'assistant', time: 1759000260, content: '对，但有三点前提：\n\n1. 前缀必须**逐字节一致**；\n2. 缓存有效期约几小时到几天，不保证；\n3. 只在同一个账号、同一个模型下生效。', think: '', files: [] },
  ],
};

const session2 = {
  id: 'a1b2c3d4-0000-4000-8000-000000000002',
  title: '帮我写一份周报 / 含斜杠与特殊字符',
  created_at: 1759500000,
  updated_at: 1759600000,
  model_type: 'expert',
  pinned: false,
  messages: [
    { role: 'user', time: 1759500000, content: '写一份周报，突出性能优化。', think: '', files: [] },
    { role: 'assistant', time: 1759500060, content: '## 本周工作总结\n\n- 接口 P99 从 820ms 降到 210ms；\n- 数据库慢查询减少 62%；\n\n> 下周计划：补齐监控告警。', think: '', files: [] },
  ],
};

const outDir = path.join(here, 'out');
fs.mkdirSync(outDir, { recursive: true });

const a = DS.buildDocx(session, { think: true, sourceUrl: 'https://chat.deepseek.com/a/chat/s/' + session.id });
fs.writeFileSync(path.join(outDir, 'session1.docx'), a);

const b = DS.buildDocx(session2, { think: false });
fs.writeFileSync(path.join(outDir, 'session2.docx'), b);

const merged = DS.buildMergedDocx([session, session2], { think: true });
fs.writeFileSync(path.join(outDir, 'merged.docx'), merged);

const zip = DS.buildZip([
  { name: '会话/001_a.docx', bytes: a },
  { name: '会话/002_b.docx', bytes: b },
  { name: '全部会话合集.docx', bytes: merged },
  { name: '原始数据.json', bytes: new TextEncoder().encode(JSON.stringify(DS.toJsonPayload([session, session2]))) },
]);
fs.writeFileSync(path.join(outDir, 'bundle.zip'), zip);

// 顺便校验 Markdown 解析结果
const blocks = DS.parseBlocks(md());
console.log('blocks:', blocks.map((x) => x.type).join(','));
console.log('inline:', JSON.stringify(DS.parseInline('**粗** *斜* `code` [链](https://a.com) 普通')));
console.log('files:', ['session1.docx', 'session2.docx', 'merged.docx', 'bundle.zip'].map((f) => {
  const p = path.join(outDir, f);
  return f + '=' + fs.statSync(p).size;
}).join(' '));
