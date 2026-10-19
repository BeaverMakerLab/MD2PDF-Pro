# MD2PDF Pro · 专业级 Markdown 转 PDF 排版工作台

<p align="center">
  <strong>面向 AI 时代交付物的开源专业排版工具 · 让大模型输出的 Markdown 一秒蜕变为出版级高管汇报 PDF</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10%2B-blue?logo=python" alt="Python Version">
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows" alt="Platform">
  <img src="https://img.shields.io/badge/Markdown-CommonMark%20%2B%20GFM-orange" alt="Markdown Support">
  <img src="https://img.shields.io/badge/Math-KaTeX%20Offline-9cf" alt="Math Support">
  <img src="https://img.shields.io/badge/Diagrams-Mermaid%20Offline-ff69b4" alt="Diagrams Support">
  <img src="https://img.shields.io/badge/Privacy-100%25%20Offline-success" alt="Privacy First">
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License">
</p>

<p align="center">
  <img src="assets/ui_preview.png" alt="MD2PDF Pro 工作台界面预览" width="850">
</p>

---

## 💡 为什么开发这个工具？

随着 **ChatGPT、Claude、DeepSeek、Kimi** 等大模型的普及，**Markdown** 已成为事实上的大模型内容输出标准。然而在日常协同与汇报中，普遍存在一个巨大痛点：

> **“Markdown 是给机器读的，而优雅排版的 PDF 才是给领导、客户和业务方看的。”**

- ❌ **直接发送 .md 文件**：非技术人员缺少专用编辑器，满屏都是 `#`、`|---|`、`$$` 和原始代码，阅读体验极差，显得极其不专业；
- ❌ **直接复制粘贴到 Word**：版式瞬间崩溃，表格严重错位变形，LaTeX 数学公式变成无意义文本，Mermaid 流程图直接丢失；
- ❌ **使用公网在线转换工具**：企业核心业务材料、提示词方案、科研成果等涉及商业秘密与敏感隐私，上传公网平台面临不可控的泄密风险；
- ❌ **使用专业排版工具（如 Pandoc / LaTeX）**：配置门槛过高，动辄需要安装几个 G 的编译环境，难以在职场与办公电脑上快速普及。

**MD2PDF Pro** 为解决这“最后一公里交付”痛点而生：**100% 本地离线运行、零第三方云端依赖、一键输出出版级精美 PDF 报告**。

---

## 📊 渲染效果实测对比 (Markdown vs PDF)

本项目自带完整的样例演示文档 [`example.md`](example.md)，在默认「政企科技蓝」预设下转换生成的出版级白皮书 [`example.pdf`](example.pdf) 效果如下：

<p align="center">
  <img src="assets/pdf_preview_showcase.png" alt="MD2PDF Pro 渲染效果实测展示" width="850">
  <br>
  <em>图：左侧为 Mermaid 矢量架构图与 LaTeX 数学公式渲染；右侧为 GFM 专业表格、任务清单与代码块渲染</em>
</p>

### 核心要素渲染对照

| 语法特性 | 大模型原始输出 ([`example.md`](example.md)) | MD2PDF Pro 出版级导出 ([`example.pdf`](example.pdf)) |
| :--- | :--- | :--- |
| **架构流程图** | ```` ```mermaid graph TD ... ``` ```` | **离线矢量流程图**：自动布局排版，自适应纸张宽度，文字清晰无畸变 |
| **数学公式** | `$$P(C \mid X) = \frac{1}{1 + e^{-(\dots)}}$$` | **KaTeX 矢量级排版**：分数线、根号、希腊字母完美高保真打印级呈现 |
| **数据表格** | `\| 序号 \| 功能模块 \| 交付状态 \|`<br>`\| :---: \| :--- \| :---: \|` | **政企科技蓝高定表格**：浅蓝质感表头背景、细灰边框、自动换行与斑马纹隔行换色 |
| **任务清单** | `- [x] 已完成集成`<br>`- [ ] 进行中` | **印刷级矢量复选框**：自动替换为科技蓝勾选框与方形待办框 |
| **代码块** | ```` ```python def render(): ... ``` ```` | **代码卡片卡槽**：等宽 Consolas 字体、科技蓝左饰条、轻灰底色卡片 |
| **删除线** | `~~原计划预算 ￥120,000~~` | **标准灰色横贯划线**：不遮挡字形，对比直观清晰 |
| **注解尾注** | `[^1]` 与 `[^1]: 技术白皮书说明` | **正文上标锚点**：文末自动生成浅灰分隔线与详细说明专区 |

---

## ✨ 核心特性

### 1. 🛡️ 100% 本地离线 & 隐私安全
- **数据永不出机**：直接调度本地已有 Chromium 内核（自动识别 Windows 自带的 Microsoft Edge 或 Chrome）静默渲染，**绝不进行任何云端数据上传与网络遥测**，满足政府、金融及大型企业严苛的保密审计标准。

### 2. 🎨 经深度调优的出版级排版审美
- **三大预设风格一键切换**：
  - **政企科技蓝（默认推荐）**：专为数字化转型方案、技术白皮书设计，表头浅蓝衬底、紫调行内代码、典雅层级标题。
  - **严谨公文风（仿宋·红黑·首行缩进）**：严格遵循公文规范排版，仿宋体正文、红黑配色、标准首行两字符缩进。
  - **GitHub 经典清爽**：经典开发者极简开源风格，灰度细线卡片、干净利落。
- **Word 经典字号体系**：采用熟悉的初号、小初、一号、小四（12pt）、五号等规范字号，告别反人类的无单位数字滑块。
- **版式与防孤行优化**：纸张纵横向自由切换、页边距三档适配，一级标题支持居中/居左并内置 `page-break-after: avoid` 防跨页落单。

### 3. 📐 CommonMark + GFM 规范 100% 深度支持
- **专业级 GFM 表格**：自动计算列宽防溢出、科技蓝专属表头、斑马纹隔行换色、单元格换行防截断。
- **离线 LaTeX 数学公式（内置 KaTeX 0.16）**：行内公式 `$E=mc^2$` 与块级大公式 `$$\int f(x)dx$$` 全本地矢量渲染，打印放大无锯齿。
- **离线 Mermaid 架构图（内置 Mermaid 10.9）**：直接在 Markdown 内使用 ````mermaid 编写流程图、时序图、架构图、甘特图，无需联网直接出图。
- **丰富扩展语法**：任务清单复选框（`- [ ]` / `- [x]`）、删除线（`~~text~~`）、文档目录（`[TOC]`）、脚注尾注（`[^1]`）、原生 HTML 标签混排。

### 4. 📑 批量并行处理与长篇报告合并
- **批量独立转换**：拖入几十篇文档或整个文件夹，一次性自动输出对应独立 PDF。
- **多章合并为单本大报告**：支持自由上下拖拽调整次序，将多篇散落的 Markdown 章节自动拼接，合成为一本装订成册的单一长篇白皮书。

### 5. 🔒 企业级正式要素支持
- **防泄密倾斜平铺水印**：支持自定义文字、字号自适应调节（14pt ~ 40pt），动态 SVG 瓦片中心旋转算法，杜绝文字溢出截断。
- **公文级页眉页脚与页码**：支持页眉密级/组织标示，页脚自动计算 `第 X 页 / 共 Y 页`，首页支持智能免除页码。

### 6. 🖥️ 现代化高分屏 DPI 自适应体验
- **原生 TkDND 拖拽**：基于 Windows 原生 OLE 拖放接口，无论单文件、多文件还是整个文件夹，直接拖入窗口即完成识别，零底层崩溃风险。
- **高分屏大字号适配**：自动识别 2.5K / 4K 150%~200% 缩放比，左侧大工作区与超大文件列表文字，清晰易读。

---

## 🚀 快速上手

### 方式一：快捷运行（推荐）

1. 下载或克隆本项目至本地：
   ```bash
   git clone https://github.com/BeaverMakerLab/MD2PDF-Pro.git
   ```
2. 双击项目根目录下的 **`【启动 MD2PDF Pro】.bat`** 即可立即启动工作台。

---

### 方式二：源码运行

确保本地已安装 Python 3.10 或更高版本：

```bash
# 1. 进入项目根目录
cd MD2PDF-Pro

# 2. 安装极简核心依赖
pip install -r requirements.txt

# 3. 启动应用
python md2pdf_pro.py
```

---

## 🛠️ Markdown 语法支持对照表

| 规范类别 | 支持语法 | 渲染效果 |
| :--- | :--- | :--- |
| **CommonMark** | 多级标题 (`#` ~ `######`) | 自动段前段后间距，防页末落单，可选居中/居左 |
| **CommonMark** | 粗体 (`**`) / 斜体 (`*`) / 引用 (`>`) | 科技蓝粗装饰左边线，浅灰高雅背景 |
| **CommonMark** | 有序/无序列表、超链接、图片 | 自动计算缩进，图片支持相对路径/绝对路径/网络图片且限宽居中 |
| **GFM 扩展** | 表格 (`\| 列1 \| 列2 \|`) | 自动表头着色、隔行斑马纹、对齐语法 (`:---`, `:---:`, `---:`) |
| **GFM 扩展** | 围栏代码块 (```` ```python ````) | 等宽 Consolas 字体、优雅圆角卡片、柔和行间距 |
| **GFM 扩展** | 任务列表 (`- [ ]`, `- [x]`) | 矢量印刷级只读复选框，高亮区分完成状态 |
| **GFM 扩展** | 删除线 (`~~已作废~~`) | 居中规范横贯线与柔和灰度显示 |
| **GFM 扩展** | 脚注与尾注 (`[^1]`) | 正文上标锚点，文末隔离线与详细说明回跳 |
| **GFM 扩展** | 自动目录 (`[TOC]`) | 扫描全文各级标题自动生成树形层级目录 |
| **高阶出版级** | LaTeX 数学公式 (`$...$`, `$$...$$`) | 内置本地离线 KaTeX 引擎，完美呈现复杂高数公式 |
| **高阶出版级** | Mermaid 矢量图 (```` ```mermaid ````) | 内置本地离线 Mermaid 引擎，流程图/时序图全离线矢量出图 |
| **高阶出版级** | 原生 HTML 标签混排 | 支持在 Markdown 中嵌入 `<div>`、`<span style="...">`、`<center>` 等 |

---

## 📂 项目结构

```text
md2pdf-pro/
├── assets/
│   ├── js/
│   │   ├── auto-render.min.js   # 离线 KaTeX 自动渲染引擎
│   │   ├── katex.min.css        # 离线 KaTeX 字体与样式
│   │   ├── katex.min.js         # 离线 KaTeX 核心库
│   │   └── mermaid.min.js       # 离线 Mermaid 图表渲染引擎
│   └── themes/                  # 预设主题样式表
│       ├── github_clean.css
│       ├── gov_academic.css
│       └── tech_blue.css
├── 【启动 MD2PDF Pro】.bat      # Windows 一键无控制台静默启动脚本
├── config.json                  # 个性化排版配置自动持久化存档
├── md2pdf_pro.py                # 工作台核心主程序（UI 与渲染引擎）
├── requirements.txt             # 核心 Python 依赖清单
├── .gitignore                   # Git 版本控制忽略文件
├── LICENSE                      # MIT 开源许可证
└── README.md                    # 项目中文说明文档
```

---

## ❓ 常见问题 (FAQ)

#### Q1：我需要额外安装 Node.js、Pandoc 或 wkhtmltopdf 吗？
> **完全不需要。** 本工具不依赖任何厚重的外部环境或编译器，仅需极简的几个 Python 模块，底层直接调度 Windows 10/11 自带的 Microsoft Edge（或 Google Chrome）提供的 Headless 打印管道，开箱即用。

#### Q2：在完全断开外网的政企内网环境下可以使用吗？
> **完全可以。** 包括所有公式引擎（KaTeX）、图表引擎（Mermaid）、界面组件以及字体样式，均已 100% 完整打包在本地 `assets/` 目录中，在拔掉网线或内网专用机上运行体验毫无差异。

#### Q3：导出的 PDF 文件保存在哪里？
> 默认保存在被转换 Markdown 文件的**同级目录下**；若在左栏勾选并选择了“指定保存目录”，则统一导出至您指定的文件夹中，转换完毕后可直接点击「打开输出目录」直达。

---

## 📄 开源许可证

本项目遵循 [MIT License](LICENSE) 开源许可证。允许自由用于商业与非商业场景，请保留原版权声明与许可证说明。
