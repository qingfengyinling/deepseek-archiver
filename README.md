# DeepSeek 对话归档 → Word

把 DeepSeek 网页端（chat.deepseek.com）里的对话，**按会话逐个导出成 Word 文档**：
每个会话一份 `.docx`，里面完整保留「你说的每一句」和「AI 回的每一段」，
包括标题层级、列表、代码块、表格、引用、粗体、链接。

不依赖任何第三方库，不上传任何数据，全部在本机完成。

---

## 一、30 秒上手（推荐）

1. 浏览器安装 **Tampermonkey**（油猴）扩展；
2. 打开 Tampermonkey → 添加新脚本 → 把 `deepseek-archiver.user.js` 的**全部内容**粘贴进去 → `Ctrl+S` 保存；
3. 打开并登录 <https://chat.deepseek.com>，右下角会出现一个深色小面板 **「DeepSeek 归档 · Word」**；
4. 点 **「导出全部会话」**，等进度条走完，浏览器会下载一个 zip 包。

解压后的结构：

```
DeepSeek归档_20261022_0900.zip
├── 归档说明.txt
├── 会话/
│   ├── 001_某个会话标题.docx
│   ├── 002_另一个会话.docx
│   └── ...
├── 全部会话合集.docx      ← 勾选「合并总文档」才有，带目录、可点击跳转
└── 原始数据.json          ← 结构化备份，可用于二次处理
```

> 只想要某一个会话：打开那个会话，点 **「当前会话」**，直接下载单个 `.docx`。

### 面板上的选项

| 选项 | 说明 |
| --- | --- |
| 额外生成一份「合并总文档」 | 所有会话合成一本书，带封面、目录、页眉页脚 |
| 保留 AI 的思考过程 | 默认关闭。开启后会把 R1 / 深度思考的推理过程也写进去 |
| 附带 JSON 原始备份 | 默认开启。以后想换排版可以拿它重新生成，不必再抓一遍 |
| 只导出最近 N 个会话 | 留空 = 全部 |

### 不想装油猴？

也能用：在 `chat.deepseek.com` 页面按 `F12` → 切到 **Console（控制台）** →
把 `deepseek-archiver.user.js` 整份粘贴进去回车，同样会弹出面板。
（文件开头的 `// ==UserScript==` 只是注释，粘进控制台不影响运行。）

---

## 二、命令行版（Python）

适合：想放到脚本里批量跑、或者只拿到了一份 `原始数据.json` / 下载的 zip 想重新生成。

需要 Python 3.9+ 和 `python-docx`：

```powershell
pip install python-docx
```

### 方式 A：直接抓取

token 的取法：登录 chat.deepseek.com → `F12` → **Application / 应用** →
**Local Storage** → `https://chat.deepseek.com` → 找到 `userToken`，复制其中的 `value`。

```powershell
python deepseek_archive.py fetch --token eyJhbGciOi... --merge
```

### 方式 B：把已导出的数据重新转成 Word

```powershell
# 直接拖 zip 进去也行
python deepseek_archive.py convert DeepSeek归档_20261022_0900.zip --merge --think
python deepseek_archive.py convert 原始数据.json -o 我的归档
```

Windows 上也可以把 zip / json **直接拖到 `运行转换.cmd` 上**，双击等结果即可。

参数：

```
fetch   --token T  [-o 输出目录] [--merge] [--think] [--limit N] [--id 会话ID]
convert <输入文件>  [-o 输出目录] [--merge] [--think]
```

`--merge` 生成合并总文档，`--think` 保留思考过程，`--limit` 只导最近 N 个。

---

## 三、导出的 Word 长什么样

- **A4 版面**，微软雅黑 + Calibri，1.33 倍行距，页边距 2.2 cm；
- **页眉**：`DeepSeek 对话归档 · 会话标题`　右侧：导出时间；
- **页脚**：`第 X 页 / 共 Y 页`（真正的域，翻页自动更新）；
- **每条消息一个色带**：蓝色 = 你，灰色 = DeepSeek，带时间戳；
- 正文按 Markdown 还原：`#` 标题 → Word 标题样式（会进导航窗格）、
  列表 → 项目符号/编号（缩进层级保留）、```` ``` ```` 代码块 → 等宽字体 + 灰底边框、
  表格 → 真表格（表头加粗、隔行浅灰）、`>` 引用 → 左侧竖线 + 浅灰底、链接可点击；
- 合并总文档额外有封面（导出时间 / 会话数 / 消息数）和可点击的目录。

### 已知取舍

| 内容 | 处理方式 |
| --- | --- |
| 数学公式 | 按 LaTeX 转可读文本（`\frac{a}{b}` → `(a)/(b)`、`\le` → `≤`），不做公式排版 |
| 图片 / 文件附件 | 只记录文件名（`附件：说明文档.pdf`），不下载文件本体 |
| 代码块语法高亮 | 不做，统一等宽灰底 |
| 已删除 / 被截断的消息 | 按接口返回的内容导出，没有内容的消息会跳过并记录 |

---

## 四、常见问题

**Q：点导出后提示「登录状态失效」**
刷新页面重新登录，再点一次。脚本读的是浏览器里现有的登录态，不碰账号密码。

**Q：导出了 0 个会话？**
点面板右下角的 **「接口诊断」**，它会把接口返回的字段打到日志里。
把日志内容发出来即可定位——多半是 DeepSeek 改了接口路径或字段名。

**Q：会话很多，会不会很慢 / 卡死？**
300 个会话大约 1~3 分钟（受网络影响），页面不会卡死，可以随时点「取消」。
合并总文档在会话特别多时体积会比较大（几百 MB），建议关掉它，只留单会话文档。

**Q：文档里的中文变成方框？**
说明打开文档的电脑缺少「微软雅黑」。Windows 自带；macOS 可装 Office 字体包。

**Q：脚本会把我聊天记录发到别处吗？**
不会。全部逻辑在页面内执行：读取接口 → 本地拼装 `.docx` → 浏览器下载。没有任何外发请求。

---

## 五、目录说明

```
deepseek-archiver/
├── deepseek-archiver.user.js   主程序（油猴脚本 / 控制台脚本二合一）
├── deepseek_archive.py         命令行版（fetch 抓取 / convert 转换）
├── 运行转换.cmd                 Windows 拖放式转换入口
├── 示例输出/                    用假数据生成的效果预览
│   ├── 示例_单个会话.docx
│   └── 示例_合并总文档.docx
└── tests/                      开发自测（构造假数据 → 生成 → 校验 → 渲染）
    ├── run_js_test.mjs         生成三个 .docx / 一个 .zip（纯 JS 层）
    ├── run_ui_test.mjs         用 DOM 桩跑一遍「面板 + 导出」全流程
    └── check_docx.py           用 python-docx 打开并校验结构
```

自测方法：

```powershell
node tests/run_js_test.mjs          # 用假会话生成 tests/out/*.docx
node tests/run_ui_test.mjs          # 走完整导出流程（mock 接口）
python tests/check_docx.py          # 校验生成的 docx 结构
```

        └─ fragments[]: REQUEST(你的话) / RESPONSE(AI 的话) / THINK(思考) / FILE(附件)
```

认证走 `Authorization: Bearer <localStorage.userToken.value>`。
接口一旦变动，改 `deepseek-archiver.user.js` 里 `apiGet` 的两个路径即可。
