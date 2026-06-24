# 未来方舟品牌文档系统 · Claude Code 指令手册

> **优先级最高。** 每次生成任何视觉文档（DOCX / PPTX / HTML / PDF）前，必须先读本文件。
> 本文件位于根目录。所有路径均相对于根目录（`brand 3/`）。

---

## 目录结构总览

```
brand 3/                          ← 根目录（_BRAND）
│
├── projects/                     ← 所有项目在这里，一个文件夹 = 一个项目
│   └── {project-name}/
│       ├── context.md            ← 项目提示词（AI 角色 / 输出风格 / 领域知识）★
│       ├── docs/                 ← 输出交付物（HTML / PDF / DOCX / PPTX）
│       │   └── img/              ← 该项目专属图片（截图、裁剪图等）
│       ├── references/           ← 该项目的参考材料（PDF、竞品图、内容简报）
│       ├── media/                ← 该项目确认要用的视觉素材
│       │   ├── MANIFEST.md       ← 素材清单，每个文件的用途说明（人工维护）
│       │   ├── covers/           ← 封面图 / Hero 大图
│       │   ├── illustrations/    ← 章节插图
│       │   ├── photos/           ← 实景照片
│       │   └── icons/            ← 自定义图标
│       └── jobs/                 ← 该项目的生成脚本（gen_*.py / render.mjs 等）
│
├── assets/                       ← 品牌 Logo（所有项目共用，禁止修改）
├── scripts/                      ← Python 品牌构建器（所有项目共用）
├── templates/html/               ← HTML/CSS 模板（所有项目共用）
├── templates/pdf/                ← PDF 输出 CSS（所有项目共用）
├── tests/                        ← 合规测试脚本（所有项目共用）
├── tokens.json                   ← 设计 Token 唯一来源（所有项目共用）
├── brand_system.md               ← 设计原则（所有项目共用）
├── project.json                  ← 项目索引 + 当前激活项目
└── CLAUDE.md                     ← 本文件
```

**核心原则：** 根目录下只放共用的品牌基础设施，所有项目特有内容放进 `projects/{name}/`。删除一个项目 = 删除整个 `projects/{name}/` 文件夹，其余不受影响。

---

## 项目管理

### 切换项目

修改 `project.json` 的 `active` 字段：

```json
{ "active": "another-project-name" }
```

### 新建项目

1. 在 `project.json` 添加记录并设为 active
2. 在 `projects/` 下建立标准目录结构（脚本会自动 `makedirs`，但 `media/` 子目录和 `MANIFEST.md` 需要手动建或让 AI 建）

```json
{
  "active": "new-project",
  "projects": {
    "new-project": {
      "name": "项目中文名",
      "description": "一句话描述",
      "created": "YYYY-MM-DD"
    }
  }
}
```

标准目录结构：
```bash
mkdir -p projects/new-project/{docs/img,references,media/{covers,illustrations,photos,icons},jobs}
# 然后复制 media/MANIFEST.md 模板过去，并创建 context.md
```

**必须同步创建 `context.md`**（见下方"项目提示词"章节的模板），否则 AI 无项目上下文，会退回通用风格。

---

## 每次生成前必做：上下文加载（三步）

> 接到任何文档生成任务后，**按顺序完成这三步，再写脚本或生成内容。**

### Step 0 — 读取 context.md（加载项目提示词）★ 最先执行

```bash
cat projects/{active}/context.md
```

- 读取并内化【角色定义】【目标受众】【输出风格】【领域知识】【输出约束】五个维度
- 这些设定在整个任务会话中持续生效，**不需要用户重复说明**
- 若 `context.md` 不存在：**立即告知用户并停止**，要求先创建（参见"项目提示词"章节）
- 若 user prompt 与 context.md 中的某项设定冲突：**user prompt 优先**，但须记录差异

### Step 1 — 扫描 references/（理解内容与背景）

```bash
ls projects/{active}/references/
```

- **PDF**：用 `pymupdf`（`fitz`）逐页提取文字和图片，理解产品功能、内容结构
- **图片**（PNG/JPG）：视觉阅读，理解风格、场景、用色倾向
- **文本文件**：直接读取，提取关键信息和内容方向

这些文件**不直接嵌入**输出，用来指导内容策划、章节结构、文案方向。

### Step 2 — 读取 media/MANIFEST.md（确定必用素材）

```bash
cat projects/{active}/media/MANIFEST.md
ls projects/{active}/media/covers/ projects/{active}/media/photos/ ...
```

- 按 `MANIFEST.md` 的「用于」字段确定每个素材出现在输出中的位置
- **不允许遗漏任何已登记素材**
- 若 `media/` 下有文件但未在 `MANIFEST.md` 登记，**主动询问用途**，不要猜测

---

## 生成脚本标准写法

脚本放在 `projects/{name}/jobs/` 下，开头固定模板：

```python
import sys, os, json

# projects/{name}/jobs/script.py 往上三级到根目录
_BRAND   = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
_PROJECT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, _BRAND)

# 路径常量
_DOCS  = os.path.join(_PROJECT, "docs")
_MEDIA = os.path.join(_PROJECT, "media")
_REFS  = os.path.join(_PROJECT, "references")
_GEN   = os.path.join(_BRAND, ".generated", os.path.basename(_PROJECT))

os.makedirs(_DOCS, exist_ok=True)
os.makedirs(os.path.join(_DOCS, "img"), exist_ok=True)
os.makedirs(os.path.join(_GEN, "reports"), exist_ok=True)
```

路径使用示例：

```python
# 输出交付物
doc.save(os.path.join(_DOCS, "output.docx"))

# 项目图片
img_path = os.path.join(_DOCS, "img", "screenshot.png")

# 确认素材
cover  = os.path.join(_MEDIA, "covers",        "cover.jpg")
photo  = os.path.join(_MEDIA, "photos",        "scene.jpg")
icon   = os.path.join(_MEDIA, "icons",         "icon.svg")
illust = os.path.join(_MEDIA, "illustrations", "diagram.png")

# 共用品牌资源
logo_svg = os.path.join(_BRAND, "assets", "logo-horizontal-primary.svg")
logo_png = os.path.join(_BRAND, "assets", "logo-horizontal-primary.png")

# 系统产物
report = os.path.join(_GEN, "reports", "compliance_report.html")
```

---

## 核心数值（禁止硬编码其他值）

### 颜色
```
主绿    #3EC99E  — 标题、强调、Logo、表格 Header
主绿深  #4B9E31  — Hover / 深色变体
主绿浅  #EAFAF5  — 表格交替行、卡片背景
辅色    #C8E13C  — 次级强调、Badge
中性900 #0E1216  — 正文标题
中性700 #3D444A  — 正文文字
中性400 #8A9199  — 说明文字、Footer
中性200 #D0D5DD  — 边框、分隔线
中性100 #F2F3F5  — 表格斑马纹、Note 背景
白色    #FFFFFF  — 页面背景、反白文字
```

### 字体
```
中文：Alibaba PuHuiTi 2.0（首选）/ PingFang SC / Microsoft YaHei / Noto Sans SC
英文：Inter / SF Pro
代码：JetBrains Mono
```
字体文件统一放在 `fonts/alibaba-puhuiti/`（随项目分发）。
PPTX/DOCX 写入字体名 `"Alibaba PuHuiTi 2.0"`；HTML/PDF 通过 `@font-face` 加载本地文件。

### 页面（DOCX A4）
```
上下边距：25.4mm   左右边距：31.7mm
H1：28pt #0E1216   H2：22pt #0E1216   H3：18pt #3EC99E
```

---

## 各格式强制规则

### DOCX

```python
from scripts.docx_builder import BrandDocx

doc = BrandDocx(doc_type="文档类型标签")
doc.add_title("主标题", subtitle="SUBTITLE")
doc.add_heading("一、章节标题", level=1)
doc.add_body("正文内容 ...")
doc.add_note("提示内容", label="注：")
doc.save(os.path.join(_DOCS, "output.docx"))
```
- 表格 Header：背景 `#3EC99E`，文字白色
- 表格斑马纹：奇数行 `#F2F3F5`
- Note 块：左侧 3pt `#3EC99E` 竖线 + `#F2F3F5` 底色

### PPTX

```python
# ── API 速查（不是结构模板，顺序由内容决定）──────────────────────
from scripts.pptx_builder import BrandPptx
import scripts.brand_tokens as BT

prs = BrandPptx()
prs.add_cover(...)            # 深色封面
prs.add_cover_light(...)      # 浅色封面（报告风）
prs.add_toc(...)              # 目录（可选，内容超过 5 页且有章节结构时才需要）
prs.add_divider_rich(...)     # 章节分隔页（章节数 ≥ 3 时才用，章节数少时直接省略）
prs.add_body_slide(...)       # 纯文字 / 要点列表
prs.add_two_col_slide(...)    # 两列对比
prs.add_three_cards(...)      # 三并列（≤3 项）
prs.add_six_cards(...)        # 六并列（4-6 项）
prs.add_big_stats(...)        # 核心指标（≤8 项）
prs.add_timeline(...)         # 时间轴 / 步骤流程（≤5 水平，6+ 蛇形）
prs.add_quote(...)            # 引用 / 佐证（深色背景）
prs.add_table_slide(...)      # 数据表格
prs.add_module_grid(...)      # 功能/模块矩阵（≤8 项）
prs.add_about_slide(...)      # 公司介绍
prs.add_closing(...)          # 结尾
prs.save(os.path.join(_DOCS, "output.pptx"))
```

> **结构原则（重要）：PPT 结构必须从内容逻辑出发，不得套用任何现有 PPT 的章节结构。**
>
> - 每次新建 PPT 前，先问：**这份 PPT 的核心叙事逻辑是什么？** 再根据内容选版式、定顺序
> - 章节数、章节名、每章页数——完全由内容决定，不存在"标准几章"
> - 同一个版式（如 Three Cards）在不同 PPT 里位置可以完全不同
> - 如果内容不需要 TOC 或章节分隔页，就不加——不要为了结构感而填充
> - 参考现有 PPT 只看**版式效果**，不看**章节结构**

#### 布局网格

| 常量 | 值 | 含义 |
|---|---|---|
| 画布 | 16:9，Inches(13.33) × Inches(7.5) | 标准宽屏 |
| 左/右边距 ML/MR | 21 mm | 所有内容元素的水平起点 |
| 可用内容宽 CW | 画布宽 − ML − MR ≈ 297 mm | |
| 页眉区 HEADER_H | 36 mm（从顶部起） | label + 大标题 + subtitle |
| 页脚区 FOOTER_H | 13 mm（从底部起） | logo + 辅助标签 |
| 内容区 CONTENT_Y/CONTENT_H | 36 mm 起，高约 92 mm | 卡片、文字、图表的有效区域 |
| 两列间距 C2_GAP | 6 mm | two_col_slide 用 |
| 三列间距 C3_GAP | 5 mm | three_cards 用 |

#### 标题区（Header）解剖

```
[label]          ← 9pt, PRIMARY_500 (#3EC99E)，大写英文，顶部 4.5mm
[大标题]         ← 26pt bold, NEUTRAL_900 (#0E1216)，有 label 时 11mm 处，否则 5.5mm 处
[装饰元素]       ← title_deco（见下方装饰规则）
[副标题 / 分隔线] ← subtitle 12pt NEUTRAL_400；无 subtitle 时画 0.3mm 灰色分隔线
```

- label 是**英文大写**辅助标签，用于标注章节归属（如 `PLATFORM CAPABILITIES`）
- 大标题只显示**单行**（≤15 字为宜），超长会压缩字体
- subtitle 用于补充说明，12pt，不超过一行

#### 页脚（Footer）

- 右下角：横向 logo PNG，高 7mm（所有内容页统一）
- 封面/分隔页/结尾页：logo 单独处理，不调用 `_footer()`

#### 卡片颜色语义（全系统统一）

适用于 `add_three_cards`、`add_six_cards`、`add_big_stats`、`add_module_grid` 等所有卡片类版式。

| 底色 | Accent 色 | 语义 | 优先级 |
|---|---|---|---|
| PRIMARY_100 `#EAFAF5` | PRIMARY_500 `#3EC99E` | 标准/主要特性 | ★★★ |
| NEUTRAL_100 `#F2F3F5` | SUCCESS_HEX `#5CC13C` | 安全/补充/次要 | ★★★ |
| SECONDARY_100 `#F8FBE7` | SECONDARY_500 `#C8E13C` | 创新/机遇 | ★★★ |
| CARD_ORANGE_BG `#FFF1DF` | WARNING_HEX `#FFB928` | 高风险/注意 | ★★★ |
| CARD_TEAL_BG `#E0F7FA` | TEAL_HEX `#3CC5CF` | 扩展/生态 | ★★ |
| CARD_PURPLE_BG `#F0E8FF` | PURPLE_HEX `#8255E1` | 战略/特殊（颜色不够时用） | ★★ |
| NEUTRAL_900 `#0E1216` | SECONDARY_500 `#C8E13C` + 白字 | **突出/亮点/强调**（非危险） | ★★★ 深色卡 |
| CARD_DANGER_BG `#FFF2F2` | DANGER_HEX `#F12D2D` | **危险/风险**（确实是危险才用） | ★ 慎用 |

**深色卡规则（hard rule）：**
- 卡片数 < 5（三卡版式）：**禁止**深色卡，三卡已自动屏蔽
- 卡片数 ≥ 5：最多 **1 张**深色卡，在 `cards` dict 里加 `"dark": True`
- 深色卡数量由 `brand_config.json` 的 `card_rules.max_dark_per_slide` 控制，**不硬编码**

#### 各版式选型速查

| 场景 | 推荐版式 | 关键参数 |
|---|---|---|
| 演讲开场 | `add_cover` | title/subtitle/tagline |
| 浅色开场（报告风） | `add_cover_light` | title/subtitle/date_or_meta |
| 章节过渡 | `add_divider_rich` | chapter_num/chapter_title/chapter_items/current_item |
| 目录页 | `add_toc` | chapters 列表（num/title/subtitle/state） |
| 文字要点/纯文字 | `add_body_slide` | bullets 列表 或 body_text 字符串 |
| 两侧对比 | `add_two_col_slide` | left/right title + content，会自动加竖向分隔线 |
| 3个并列特性 | `add_three_cards` | cards ≤3，无深色卡 |
| 6个并列特性 | `add_six_cards` | cards ≤6，最多1张深色卡 |
| 核心指标 ≤4 | `add_big_stats` | stats 4项，colorful 自动=True（两列彩色） |
| 核心指标 5-8 | `add_big_stats` | stats 5-8项，colorful 自动=False（中性白底） |
| 功能模块矩阵 ≤8 | `add_module_grid` | modules 含 num/icon/title/en/bullets |
| 时间轴 ≤5步 | `add_timeline` | 自动渲染为**水平**单行 |
| 时间轴 6步+ | `add_timeline` | 自动渲染为**双行蛇形**（wrap 模式） |
| 数据表格 | `add_table_slide` | headers/rows/note；Header 行绿底白字，斑马纹 |
| 引用/佐证/数据背书 | `add_quote` | 深色背景，大引号，作者署名；**不加 title_deco** |
| 公司介绍 | `add_about_slide` | body_text + callout_items + right_panel |
| 结尾/Call to Action | `add_closing` | slogan_parts（多色分段）+ slogan_sub |

#### 正文排版数值

| 元素 | 字号 | 颜色 | 行距 |
|---|---|---|---|
| 大标题（header） | 26pt bold | NEUTRAL_900 | — |
| 副标题（header subtitle） | 12pt | NEUTRAL_400 | — |
| label（section tag） | 9pt | PRIMARY_500 | — |
| 两列列标题 | 13pt bold | PRIMARY_500 | — |
| 两列正文 | 14pt | NEUTRAL_700 | 22pt |
| 子弹点（body_slide） | 17pt | NEUTRAL_700 | 21pt + 8pt 段前 |
| 正文段落（body_text） | 16pt | NEUTRAL_700 | 26pt |
| 卡片标题 | 16pt bold | NEUTRAL_900 | — |
| 卡片 tag（eyebrow） | 9pt bold | accent 色 | — |
| 卡片正文 | 13pt | NEUTRAL_700 | 18pt |
| 统计数值（big_stats） | 44pt bold | accent 色 | — |
| 统计标签 | 14pt bold | NEUTRAL_900 | — |
| 统计说明 | 12pt | NEUTRAL_700 | 18pt |

#### 封面/分隔/结尾页 渐变标题色

封面（`add_cover`）、章节页（`add_divider*`）、结尾页（`add_closing`）的主标题均使用三色渐变：

```
PRIMARY_500 #3EC99E → SUCCESS_HEX #5CC13C → SECONDARY_500 #C8E13C
```

普通内容页标题（`_header`）**不使用渐变**，用纯色 NEUTRAL_900。

#### PPTX 装饰性元素（Decorations）

装饰图片统一存放于 `assets/装饰性元素/`，分三个系列：

| 系列 | 使用场景 | 颜色 |
|---|---|---|
| **A 系**（`A*.png`） | 白色背景的普通内容页（最常用） | Primary `#3EC99E` |
| **B 系**（`B*.png`） | 灰色底或整页图片的页面 | Success `#5CC13C` |
| **C 系**（`C*.png`） | 深色背景页（如封面风格内容页） | Secondary `#C8E13C` |

每个系列有四种装饰类型：

| 类型 | 文件名规律 | 用途 |
|---|---|---|
| `underline` | `{系列}下划*.png` | 通用：任何页面标题关键词强调 |
| `circle` | `{系列}画圈*.png` | 突出**核心答案词**（陈述句）或圈出问句答案（设问句） |
| `arrow1` | `{系列}箭头1*.png` | 文字 + 图片**顶部对齐**时，从文字指向图片 |
| `arrow2` | `{系列}箭头2*.png` | 文字 + 图片**底部对齐**时，从文字指向图片 |

**使用规则（硬性约束）：**

1. **每页最多一个装饰**，禁止同一页出现两种及以上装饰（页面会显得花哨）
2. 有图文排版（`text_image` / `image_left_text` 类布局）的页面，优先考虑 arrow1 / arrow2
3. 无箭头的内容页（含图文但不符合对齐条件、纯文字页），**必须**有 underline 或 circle
4. 封面（`add_cover`）、章节分隔页（`add_divider*`）、TOC、结尾页（`add_closing`）**不加装饰**
5. circle 须置于标题文字**下层**（z-order behind），底部与标题文字基线对齐

**最小宽度规则：** 装饰视觉宽度始终保持约 **4 个字宽**（避免看起来太小太轻）。
即使目标词只有 2 字（如"八大"、"承诺"），也按 4 字宽渲染，并以目标词为中心居中。

**`title_deco` 参数说明**（所有 `add_*` 方法均支持）：

```python
title_deco = {
    "type":       "underline",  # "underline" | "circle" | "arrow1" | "arrow2"
    "char_start": 7,            # 目标词在标题中的起始字符偏移（从 0 开始，按视觉字宽估算）
    "char_count": 4,            # 目标词的字符数（不足 4 时系统自动扩展到 4 字宽）
    "char_w_mm":  8.1,          # 可选，覆盖默认字宽（默认 8.1mm，26pt bold CJK）
    "series":     "a",          # 可选，默认 "a"；深色页用 "c"，图片背景页用 "b"
}
```

**circle 字符宽度估算（26pt bold）：**
- 全宽 CJK 字符：≈ 8.1 mm/字
- 半宽数字/字母（如"8"）：≈ 4.0 mm/字
- 间隔符" · "：≈ 2 字宽（约 16 mm）

**典型用法示例：**

```python
# 下划线：强调末尾4字（最常见）
prs.add_three_cards(title="工厂培训管理的三大困局", ...,
    title_deco={"type": "underline", "char_start": 7, "char_count": 4})  # 三大困局

# 画圈：陈述句突出核心答案（全员覆盖）
prs.add_body_slide(title="我们的建议：方案 B 全员覆盖", ...,
    title_deco={"type": "circle", "char_start": 11, "char_count": 4})   # 全员覆盖

# 画圈：设问句答案（8周，半宽"8"约0.5字宽）
prs.add_timeline(title="8 周标准交付，无需 IT 团队介入", ...,
    title_deco={"type": "circle", "char_start": 0, "char_count": 2})    # 8 周

# " · " 分隔符处理（约占2字宽偏移）
prs.add_table_slide(title="两种方案 · 清晰对比", ...,
    title_deco={"type": "underline", "char_start": 6, "char_count": 4}) # 清晰对比
```

### HTML

输出文件在 `projects/{name}/docs/` 时，路径为：

```html
<link rel="stylesheet" href="../../../templates/html/brand.css">
<img src="../../../assets/logo-horizontal-primary.svg" height="32">
```

- 颜色全部用 `var(--color-primary-500)` 等 CSS 变量，禁止内联写死十六进制

### PDF（Puppeteer 方案）

```javascript
// projects/{name}/jobs/render.mjs
import puppeteer from 'puppeteer';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// 往上三级到根目录，再定位字体目录
const FONTS_DIR = path.resolve(__dirname, '..', '..', '..', 'fonts', 'alibaba-puhuiti');

const browser = await puppeteer.launch({
  headless: true,
  executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  args: ['--no-sandbox', '--allow-file-access-from-files'],
});
const page = await browser.newPage();

// 注入本地字体（必须在 setContent 之前，确保 @font-face 生效）
await page.addStyleTag({ content: `
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:100;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_35_Thin_35_Thin.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:300;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_45_Light_45_Light.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:400;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_55_Regular_55_Regular.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:700;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_55_Regular_85_Bold.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:500;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_65_Medium_65_Medium.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:600;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_75_SemiBold_75_SemiBold.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:800;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_95_ExtraBold_95_ExtraBold.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:900;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_105_Heavy_105_Heavy.ttf") format("truetype"); }
` });

await page.setContent(readFileSync(SRC, 'utf8'), { waitUntil: 'networkidle0' });
await page.evaluateHandle('document.fonts.ready');
await new Promise(r => setTimeout(r, 1500));
await page.pdf({
  path: DEST, width: '297mm', height: '210mm',
  printBackground: true,
  margin: { top: 0, bottom: 0, left: 0, right: 0 },
  preferCSSPageSize: true,
});
await browser.close();
```

Puppeteer 依赖安装（仅首次）：
```bash
cd /tmp/pdf-gen && npm install puppeteer
```

> **字体注入说明**：用 `file://` 绝对路径注入 `@font-face`，在 `setContent` 之前执行，
> 避免 Puppeteer 无 baseURL 时相对路径失效。不再依赖 Google Fonts 网络请求。

---

## 项目提示词（context.md）

### 三层优先级

```
user prompt   ← 用户单次任务输入，最高优先级，仅影响当次任务
context.md    ← 项目级 AI 上下文，本章节说明的文件
CLAUDE.md     ← 系统硬约束（品牌色值、字体、路径、合规），永远生效，不可被覆盖
```

**context.md 能覆盖的范围**：写作角色、语气、领域术语偏好、受众假设、内容禁忌。
**context.md 不能覆盖的范围**：CLAUDE.md 中的颜色值、字体名、路径规范、禁止事项表。

### 标准格式（新建项目时复制此模板）

```markdown
# {项目名} · 项目提示词（Project Context）

> **自动加载**：本文件由系统在每次任务开始前读取，无需用户重复说明。
> 优先级：高于 CLAUDE.md 风格判断，低于 user prompt 单次指令，低于 CLAUDE.md 技术硬约束。

---

## 角色定义（Role）
<!-- 描述 AI 在本项目中的专家身份和职责范围 -->
你是一名...

## 目标受众（Audience）
<!-- 文档的主要和次要阅读对象，以及他们的背景假设 -->
- 主要受众：...
- 次要受众：...

## 输出风格（Style）
<!-- 语言、语气、结构偏好、文案密度要求 -->
- 语言：中文为主...
- 语气：...
- 结构：...

## 领域知识（Domain）
<!-- 本项目涉及的行业背景、核心概念、标准规范、竞品认知 -->

## 输出约束（Constraints）
<!-- 禁止事项、特殊限制、数据引用规范 -->
- 不虚构...
```

### 维护规则

- context.md 由**项目负责人（人工）**维护，AI 不得自行修改
- 每次项目目标、受众或文档类型发生重大变化时，**人工更新** context.md
- 若同一项目需要面向不同受众出不同版本文档，在 context.md 的受众章节中用子标题区分，并在 user prompt 中指定"按受众A版本"或"按受众B版本"

---

## 每次输出必做：合规测试

```bash
python3 tests/test_compliance.py projects/{name}/docs/output.docx --format docx
python3 tests/test_compliance.py projects/{name}/docs/output.html --format html
```

**通过标准：总分 ≥ 80 / 100。低于 80 必须修复后重新测试。**

---

## 品牌信息自定义指南

### 修改公司信息（适用于其他公司部署）

品牌信息集中管理在 `scripts/brand_tokens.py`，修改时按以下顺序：

#### 1. 修改核心配置文件
```python
# scripts/brand_tokens.py 第 21-34 行
BRAND_NAME_CN  = "你的公司简称"
BRAND_NAME_EN  = "YourCompany"
BRAND_FULL_CN  = "你的公司全称有限公司"
BRAND_FULL_EN  = "Your Company Full Name Co., Ltd."

# 地址信息
COMPANY_ADDRESS_CN = "你的公司地址"
COMPANY_ADDRESS_EN = "Your Company Address"
```

#### 2. 同步更新模板文件
修改 `brand_tokens.py` 后，需手动更新以下模板中的硬编码文本：

- `templates/html/base.html`：第 78、100 行的公司名称
- `templates/pdf/brand.css`：第 22 行的页眉公司名称
- `templates/ppt/gen_brand_master.py`：第 250、259 行的公司名称

#### 3. Logo 文件替换
将新公司的 Logo 文件放入 `assets/` 目录，保持文件名一致：
- `logo-horizontal-primary.png` / `.svg`
- `logo-stacked-primary.png`
- `logo-mark-primary.png`

#### 4. 验证一致性
```bash
# 生成测试文档检查
python3 projects/sample/jobs/test_brand_consistency.py

# 确保所有格式显示相同的公司名称
ls projects/sample/docs/test_*
```

### 配置检查清单

使用此清单确保品牌信息在所有输出格式中保持一致：

- [ ] `brand_tokens.py` 中的公司信息已更新
- [ ] DOCX 文档页眉页脚显示新公司名称
- [ ] HTML 模板 Footer 显示新公司名称
- [ ] PDF 页眉显示新公司名称
- [ ] PPT 封面页显示新公司名称
- [ ] 证明文件签名块显示新公司名称
- [ ] 所有 Logo 文件已替换为新公司 Logo

---

## 禁止事项

| 禁止 | 原因 |
|---|---|
| 硬编码非品牌颜色（`#FF0000`、`#333` 等） | 破坏视觉一致性 |
| 不从 `scripts/docx_builder.py` 起步直接写 DOCX | 容易遗漏页眉/页脚/字体 |
| 修改 `tests/golden/` 下的文件 | 黄金参考不可变 |
| 在 DOCX Header 用 stacked logo | 空间不足，应用 horizontal |
| 在封面 / 大幅区域用 horizontal logo | 应用 stacked 以增强品牌感 |
| CSS 中写死十六进制颜色 | 应使用 CSS 变量 |
| 在 `projects/` 之外存放项目专属文件 | 破坏结构一致性，无法按项目清理 |
| 生成脚本未加正确的 sys.path.insert | 会导致 scripts 模块找不到 |
| AI 自行修改 context.md | 项目提示词由人工维护，AI 只读不写 |
| 跳过 Step 0 直接开始生成 | 无项目上下文会导致角色/风格/领域偏差 |
| 在 gen_*.py 中硬编码邮箱、官网、电话 | 应使用 `BT.COMPANY_EMAIL` / `BT.COMPANY_WEBSITE` / `BT.COMPANY_PHONE`，修改一处全部生效 |
