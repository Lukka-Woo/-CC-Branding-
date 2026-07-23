# 未来方舟品牌文档系统 · Claude Code 指令手册

> **优先级最高。** 每次生成任何视觉文档（DOCX / PPTX / HTML / PDF）前，必须先读本文件。
> 本文件只包含行为规则；API 参数细节见 `reference/pptx-api.md`，项目初始化见 `reference/setup-guide.md`。

---

## 目录结构

```
brand 3/                          ← 根目录（_BRAND）
├── projects/{name}/
│   ├── context.md                ← 项目提示词 ★ 必须存在
│   ├── docs/                     ← 输出交付物（HTML / PDF / DOCX / PPTX）
│   │   └── img/
│   ├── references/               ← 参考材料（PDF、竞品图、内容简报）
│   ├── media/                    ← 确认要用的视觉素材
│   │   └── MANIFEST.md           ← 素材清单（人工维护）
│   └── jobs/                     ← 生成脚本（gen_*.py / render.mjs）
├── assets/                       ← 品牌 Logo（禁止修改）
├── scripts/                      ← Python 品牌构建器
├── templates/                    ← HTML / PDF / PPT 模板
├── reference/                    ← API 参考 + 操作指南（按需读取）
├── tokens.json                   ← 设计 Token 唯一来源
├── brand_system.md               ← 设计原则
└── project.json                  ← 项目索引 + 当前激活项目（active 字段）
```

根目录只放共用品牌基础设施；项目特有内容全部放 `projects/{name}/`。

---

## 每次生成前必做（三步，按顺序）

### Step 0 — 读取 context.md ★ 最先执行

```bash
cat projects/{active}/context.md
```

内化【角色定义】【目标受众】【输出风格】【领域知识】【输出约束】，整个会话持续生效。
若文件不存在：**立即停止**，告知用户先创建（模板见 `reference/setup-guide.md`）。
若 user prompt 与 context.md 冲突：**user prompt 优先**。

### Step 1 — 扫描 references/ 并建立内容清单

```bash
ls projects/{active}/references/
```

PDF 用 `pymupdf`（`fitz`）提取文字和图片；图片视觉阅读；文本直接读取。
这些文件**不嵌入**输出，用于指导内容策划和文案方向。

**扫描后必须建立内容清单**——在开始任何版式选型前，逐主题记录：

```
主题 A：N 个内容点
  1. [完整标题] — [完整说明]
  2. [完整标题] — [完整说明]
  ...N. [完整标题] — [完整说明]

主题 B：M 个内容点
  ...
```

**红线：内容清单在版式选型之前锁定，不允许事后因"版式只有3个槽"而回头删减内容点。**

### Step 1.5 — 内容优先的版式选型（hard rule）

内容清单建立后，按以下流程为每个主题选版式：

```
确定内容点数量 N（来自 references，不受版式影响）
  ↓
N == 3  → add_three_cards / add_numbered_rows / add_accent_rows
N == 4  → add_four_cards / add_two_col_slide / add_numbered_rows / add_accent_rows
N == 5  → add_six_cards + intro_text/intro_flow（均衡布局：intro+2 / 3）
N == 6  → add_six_cards
N == 7  → add_module_grid + intro_text/intro_flow（均衡布局）
N >= 8  → add_module_grid / add_table_slide / add_body_slide（或分页）
N 不确定 → add_body_slide + pill bullets
```

**三条禁止：**

| 禁止 | 原因 |
|---|---|
| 用 preset slot 数反推"这里有 N 个点" | 应从原文提取 N，再选匹配 preset |
| 删减原文内容点以适配 preset slot 数量 | 内容完整性优先于版式整洁 |
| 将不同概念"合并"进一个 slot 以凑数 | 可精简语言，不得合并不同概念 |

### Step 2 — 读取 media/MANIFEST.md

```bash
cat projects/{active}/media/MANIFEST.md
```

按「用于」字段确定每个素材的输出位置。**不允许遗漏已登记素材。**
`media/` 下有文件但未登记时，主动询问用途，不猜测。

---

## 生成脚本固定头部

```python
import sys, os

_BRAND   = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
_PROJECT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, _BRAND)

_DOCS  = os.path.join(_PROJECT, "docs")
_MEDIA = os.path.join(_PROJECT, "media")
_REFS  = os.path.join(_PROJECT, "references")

os.makedirs(_DOCS, exist_ok=True)
os.makedirs(os.path.join(_DOCS, "img"), exist_ok=True)
```

路径常量用法：

```python
doc.save(os.path.join(_DOCS, "output.docx"))
cover  = os.path.join(_MEDIA, "covers", "cover.jpg")
logo_png = os.path.join(_BRAND, "assets", "logo-horizontal-primary.png")
```

---

## 内容与布局分离（content.json）

新建或重构 job 脚本时，**将文案内容与布局代码分离**：

```
projects/{name}/jobs/
  content.json    ← 所有幻灯片的文字内容（标题、卡片文案、数据、bullet）
  gen_*.py        ← 读取 content.json，负责版式调用和布局参数
```

### content.json 结构

```json
{
  "_meta": { "project": "项目名", "deck": "文件名", "lang": "cn", "version": "v1" },
  "slides": {
    "cover":    { "title": "...", "subtitle": "...", "tagline": "..." },
    "toc":      { "title": "...", "description": "...", "chapters": [...] },
    "div_01":   { "chapter_num": "01", "chapter_title": "...", "subtitle": "..." },
    "cards_01": { "title": "...", "subtitle": "...", "cards": [...] },
    "stats":    { "title": "...", "subtitle": "...", "stats": [["值", "标签", "说明"], ...] },
    "timeline": { "title": "...", "subtitle": "...", "milestones": [["期间", "标题", "说明"], ...] },
    "closing":  { "slogan_sub": "..." }
  }
}
```

### gen_*.py 读取方式

```python
import json

with open(os.path.join(os.path.dirname(__file__), "content.json"), encoding="utf-8") as f:
    S = json.load(f)["slides"]

TOC_ITEMS = [("01", "章节一"), ("02", "章节二")]   # 分隔页共用常量

prs.add_cover(**S["cover"])
prs.add_toc(**S["toc"], label="TABLE OF CONTENTS")
prs.add_divider_rich(**S["div_01"], chapter_items=TOC_ITEMS, current_item=0)
prs.add_three_cards(**S["cards_01"], label="LABEL", title_deco={...})

# stats / milestones 需从 list 转 tuple：
prs.add_big_stats(
    title=S["stats"]["title"], subtitle=S["stats"]["subtitle"],
    stats=[tuple(st) for st in S["stats"]["stats"]],
    label="...", title_deco={...},
)
prs.add_timeline(
    title=S["timeline"]["title"], subtitle=S["timeline"]["subtitle"],
    milestones=[tuple(m) for m in S["timeline"]["milestones"]],
    label="...", title_deco={...},
)
```

### 分离原则

| 放入 content.json | 留在 gen_*.py |
|---|---|
| 所有中文文案（title / subtitle / body / bullets / tag） | `label`（英文大写，版式身份） |
| 卡片 / 模块 / 统计数据结构 | `title_deco` 参数 |
| `dark` 布尔标记（内容语义） | `chapter_items`（TOC 常量） |
| `featured` 布尔标记 | BT 颜色引用（`BT.COMPANY_WEBSITE` 等） |
| 双语 `{"cn": ..., "en": ...}` dict | `right_panel` / `slogan_parts`（含 BT 引用） |

**不对旧项目强制回填；新项目和改版从第一个 job 脚本开始遵守。**

---

## 核心数值（禁止硬编码其他值）

具体色值 / 字体族 / DOCX 页面数值**不在本文件重复罗列**——唯一数据源是 `tokens.json`
（经 `scripts/brand_tokens.py` 加载）。当前生效的完整数值表见自动生成文件：

**`reference/brand-values.md`**

修改 `tokens.json` 后运行一次 `python3 scripts/gen_brand_reference.py` 重新生成，
确保文档数值与实际渲染值始终一致，不会像手工维护的表格那样悄悄漂移。

所有颜色常量通过 `import scripts.brand_tokens as BT` 使用，禁止在脚本中内联十六进制。
PPTX/DOCX 写字体名固定引用 `BT.FONT_CN` / `BT.FONT_EN`；HTML/PDF 通过 `@font-face` 加载本地文件（`fonts/alibaba-puhuiti/`）。

---

## DOCX 格式规则

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
- DOCX Header 使用 horizontal logo（PNG），禁用 SVG 注入

---

## PPTX 格式规则

### 结构原则

> **PPT 结构必须从内容逻辑出发，不得套用现有 PPT 的章节结构。**

- 每次新建 PPT 前先问：**这份 PPT 的核心叙事逻辑是什么？**
- 章节数、章节名、每章页数——完全由内容决定
- 不需要 TOC 或章节分隔页时，不加——不要为了结构感填充
- 参考现有 PPT 只看**版式效果**，不看**章节结构**

### TOC subtitle 完整性规则（hard rule）

`add_toc` 每个 chapter 的 `subtitle` 字段必须覆盖该章内**所有不重复的逻辑主题**：

- **一个主题 = 一个条目**，不论该主题跨几张幻灯片（如"核心优势与资质"占3页，在 subtitle 里仍只写一次）
- **不允许遗漏**：章内每张幻灯片的核心主题，都必须能在 subtitle 里找到对应条目
- **允许合并**：内容高度相关的连续子场景（如"场景一~三"都属于安全合规），可合并为一个条目
- **写法**：各条目用 ` · ` 分隔，保持简短（2–4字/条）

**正确示例（01 关于我们，含P4–P7）：**
```
subtitle="运营挑战 · 行业趋势 · 核心优势与资质 · 解决方案全景"
```

**错误示例（遗漏P7）：**
```
subtitle="运营挑战 · 行业趋势 · 核心优势与资质"  ← ❌ 缺解决方案全景
```

**实施步骤**：生成脚本写完后，对照每章的 `prs.add_*()` 调用列表，逐章核查 subtitle 是否完整覆盖。

### 布局网格

| 常量 | 值 | 含义 |
|---|---|---|
| 画布 | Inches(13.33) × Inches(7.5)，16:9 | 标准宽屏 |
| 左/右边距 ML/MR | 21 mm | 所有内容元素水平起点 |
| 可用内容宽 CW | ≈ 297 mm | 画布宽 − ML − MR |
| 页眉区 HEADER_H | 36 mm（从顶部起） | label + 标题 + subtitle |
| 页脚区 FOOTER_H | 13 mm（从底部起） | logo + 辅助标签 |
| 内容区 CONTENT_Y | 36 mm 起，高约 92 mm | 卡片、文字、图表有效区域 |

### Header 规则

```
[label]       ← 9pt PRIMARY_500，大写英文，顶部 4.5mm
[大标题]      ← 26pt bold NEUTRAL_900，有 label 时 11mm 处，否则 5.5mm 处
[subtitle]    ← 12pt NEUTRAL_400；无 subtitle 时画 0.3mm 灰色分隔线
```

- label 是**英文大写**辅助标签（如 `PLATFORM CAPABILITIES`）
- 大标题只显示**单行**（≤15 字为宜）

### Footer 规则

- 右下角：横向 logo PNG，高 7mm（所有内容页统一）
- 封面/分隔页/结尾页：logo 单独处理，不调用 `_footer()`

### 卡片颜色语义（全系统统一）

适用于 `add_three_cards`、`add_six_cards`、`add_big_stats`、`add_module_grid`。
本表只定义**角色 → 语义 → 优先级**的映射规则（模型/品牌无关，换品牌也不用改这张表）；
具体色值见 `reference/brand-values.md`。

| 底色角色 | Accent 角色 | 语义 | 优先级 |
|---|---|---|---|
| PRIMARY_100 | PRIMARY_500 | 标准/主要特性 | ★★★ |
| NEUTRAL_100 | SUCCESS | 安全/补充/次要 | ★★★ |
| SECONDARY_100 | SECONDARY_500 | 创新/机遇 | ★★★ |
| CARD_ORANGE_BG | WARNING | 高风险/注意 | ★★★ |
| CARD_TEAL_BG | TEAL | 扩展/生态 | ★★ |
| CARD_PURPLE_BG | PURPLE | 战略/特殊 | ★★ |
| NEUTRAL_900 | SECONDARY_500 + 白字 | **突出/亮点**（非危险） | ★★★ 深色卡 |
| CARD_DANGER_BG | DANGER | **危险/风险**（确实危险才用） | ★ 慎用 |

**深色卡 hard rule：**
- 卡片数 < 5（三卡版式）：**禁止**深色卡（三卡已自动屏蔽）
- 卡片数 ≥ 5：最多 **1 张**深色卡，在 `cards` dict 加 `"dark": True`

**卡片边框 hard rule：**
- 有颜色背景（所有 `*_100` 浅色系、深色卡、任何非白底）：**禁止加边框**
- 白色背景卡片：可加 `BORDER_DEFAULT_HEX` 灰色细边框

### 版式多样性规则（hard rule）

**同一版式在一套 PPT 中最多出现 2 次。** 超出部分须换用视觉逻辑相近但结构不同的替代版式。

| 超出的版式 | 推荐替代 |
|---|---|
| `add_three_cards`（第3次+） | `add_numbered_rows`（编号问题/原因列表）或 `add_accent_rows`（技术规格/数据类型对比） |
| `add_two_col_slide`（第3次+） | `add_two_col_pills`（分类标签）或 `add_body_slide` |
| `add_body_slide`（第3次+） | `add_module_grid` 或 `add_big_stats` |
| `add_big_stats`（第3次+） | `add_timeline` 或 `add_six_cards` |

**例外：** 内容逻辑上要求严格并列对齐（如3个业务模块各自介绍页），同一版式可酌情保留。封面/分隔/结尾页不计入次数。

**实施建议：** 脚本完成后统计各版式调用次数；如有违规，参照替代表修改，并同步调整内容表达粒度。

---

### 版式选型速查

| 场景 | 推荐版式 |
|---|---|
| 演讲开场 | `add_cover` |
| 浅色开场（报告风） | `add_cover_light` |
| 章节过渡（章节数 ≥ 3） | `add_divider_rich` |
| 目录（内容 > 5 页且有章节结构） | `add_toc` |
| 文字要点 / 区块导航 | `add_body_slide`（支持 dict pill bullets） |
| 两侧对比 | `add_two_col_slide` |
| 两侧分类标签 | `add_two_col_pills` |
| 3 个并列 | `add_three_cards` |
| **3 个编号问题/原因/缺陷**（替代 three_cards） | `add_numbered_rows` |
| **3 个数据类型/规格对比**（替代 three_cards） | `add_accent_rows` |
| **4 个并列（富文本特性/壁垒/对比）** | `add_four_cards`（单行横排，禁用 `add_six_cards`） |
| 6 个并列 | `add_six_cards` |
| **5 个并列** | `add_six_cards` + `intro_text`/`intro_flow` |
| 核心指标 ≤4 | `add_big_stats`（自动彩色两列） |
| 核心指标 5-8 | `add_big_stats`（自动中性白底） |
| 功能模块 4/6/8 项 | `add_module_grid` |
| **功能模块 5 项** | `add_module_grid` + `intro_text`/`intro_flow` |
| **功能模块 7 项** | `add_module_grid` + `intro_text`/`intro_flow` |
| 有方向性流程 / 步骤 | `intro_flow=["A","B","C"]`（pill→箭头→pill） |
| 时间轴 ≤5 步 | `add_timeline`（自动水平） |
| 时间轴 6 步+ | `add_timeline`（自动蛇形双行） |
| 数据表格 | `add_table_slide` |
| 引用 / 数据背书 | `add_quote`（深色背景，**不加 title_deco**） |
| 公司介绍 | `add_about_slide` |
| 结尾 / Call to Action | `add_closing` |

### 均衡布局规则（hard rule）

**任何卡片/模块布局，禁止出现孤立行（最后一行只有 1 个）。**

| 数量 | 禁止做法 | 正确做法 |
|---|---|---|
| 5 张卡片 | 直接传 5 → 3+2 行 | 传 `intro_text`/`intro_flow` → text+2 / 3 均衡 |
| 5 个模块 | 直接传 5 → 4+1 行 | 传 `intro_text`/`intro_flow` → 3 列 text+2 / 3 |
| 7 个模块 | 直接传 7 → 4+3 行 | 传 `intro_text`/`intro_flow` → 4 列 text+3 / 4 |
| **4 张卡片** | `add_six_cards`（产生 3+1 孤立行）| **改用 `add_four_cards`（横排单行）** |
| 6 卡片，4/6/8 模块 | — | 直接传，无需 intro |

**intro slot 内容参数（可并用）：**
- `intro_label`：小标题，8pt 绿色，≤8 字
- `intro_text`：2-3 句概括，不重复标题
- `intro_flow`：字符串列表，渲染为 pill→箭头→pill

**intro slot 视觉规则（hard rule）：** 无底色、无边框、无左侧装饰矩形，直接在白色背景上渲染。

---

## 卡片内容布局类型（Card Inner Layout）

卡片背景形状由版式方法控制（颜色、圆角、边框）；**卡片内部内容**可通过每张卡片 dict 的 `"layout"` 字段选择以下类型。不传 `"layout"` 时默认为 `"vertical_stack"`（向下兼容）。

| 布局类型 | 适用场景 | 关键 data 字段 |
|---|---|---|
| `"vertical_stack"` | **默认**，tag→title→bar→body 竖排 | `tag`, `title`, `body` |
| `"horizontal_split"` | 有明确数值指标，数字放右侧，说明文字放左侧 | `metric`（右侧大字），`title`，`body` |
| `"icon_left"` | 功能/模块类卡片，圆形 badge 在左上，正文在右侧及下方 | `icon`（1-2字母），`title`，`body` |
| `"quote_card"` | 用户引语、数据背书、权威声明 | `quote`（引用正文），`attribution`（来源） |

### `horizontal_split` 使用示例

```python
# 在 add_three_cards / add_six_cards / add_four_cards 中
{"layout": "horizontal_split", "metric": "99.9%", "title": "数据连续性", "body": "断网自动触发插值，历史数据无缺口"}

# 在 add_big_stats 中（整个方法级别控制）
prs.add_big_stats(..., card_layout="horizontal_split")
# 此时 stats 的 (val, label, desc) 分别对应 metric右侧 / title左上 / body左下
```

### `icon_left` 使用示例

```python
{"layout": "icon_left", "icon": "IoT", "title": "实时数据采集", "body": "秒级采集，15 分钟时间桶自动拼接"}
```

### `quote_card` 使用示例

```python
{"layout": "quote_card", "quote": "上线三个月，峰时用电占比从 78% 降到 31%。", "attribution": "某头部汽车 Tier-1 厂长"}
```

### 文本溢出防护（hard rule）

所有卡片 body 文本框统一设置 `auto_size = TEXT_TO_FIT_SHAPE`，文字自动收缩适配文本框，**不允许文字溢出卡片边界**。生成脚本中的 body 文字不需要手动插入 `\n`，让系统自动换行。

---

## Pill 使用规范

### 正确使用场景

| 场景 | 位置 |
|---|---|
| 步骤 / 流程 | `intro_flow`、`body_slide` bullets |
| 状态变迁 | `body_slide` bullets `"style": "warning"/"danger"` |
| 正文区块导航 | `body_slide` bullets `{"pill": "区块名", "text": "内容"}` |
| 工作流 | `intro_flow=["A","B","C"]` |

### 禁止用 Pill 的场景

| 禁止 | 原因 |
|---|---|
| `add_three_cards` / `add_six_cards` 的 `"tag"` 字段 | `tag` 是卡片眉题，已有背景色区分；渲染为小号加粗彩色文字 |
| `add_module_grid` 的 `"en"` 字段 | `en` 是模块英文副标签；渲染为小号加粗彩色文字 |
| 段落里的来源/日期属性 | 是正文内容，不是导航标签 |

### Pill 自动渲染说明

| 元素 | 渲染方式 |
|---|---|
| 卡片 `"tag"` 字段 | 小号加粗彩色文字（eyebrow），**不是 pill** |
| `add_two_col_slide` 左列标题 | PRIMARY_500 filled pill |
| `add_two_col_slide` 右列标题 | SECONDARY_500 filled pill（自动深色文字） |
| `add_module_grid` `"en"` 字段 | 小号加粗 accent 彩色文字，**不是 pill** |
| `add_timeline` period 标签 | 彩色 filled pill（系统内置） |

**Pill 样式名和参数详见 `reference/pptx-api.md`。**

---

## Callout / Note 块

callout 是页面底部补充说明区块。所有支持 `note=` 的方法：`add_body_slide` / `add_two_col_slide` / `add_big_stats` / `add_timeline` / `add_table_slide`。

| 场景 | style | 背景 |
|---|---|---|
| 通用注解、限制条件 | `"note"`（默认） | 渐变 `#E8F9F3` → `#F8FBE8` |
| 功能说明 | `"info"` | 同 note |
| 操作建议 | `"tip"` | 渐变 `#F8FBE8` → `#E8F9F3` |
| 需注意的条件/例外 | `"warning"` | 渐变 `#FFF1DF`（浅橙）→ `#F8FBE7`（SECONDARY_100） |
| 数据风险、不可逆操作 | `"danger"` | `#FFF2F2` 红底纯色（慎用） |

note/info/tip 用 primary 绿色胶囊；warning 用橙色胶囊（渐变底）；danger 保留纯色红底。
所有渐变方向：左→右（`angle_deg=0`）。视觉规则：无边框、无左侧装饰线。固定高度 `Mm(12)`，置于内容区底部。

---

## 正文排版数值

| 元素 | 字号 | 颜色 |
|---|---|---|
| 大标题（header） | 26pt bold | NEUTRAL_900 |
| 副标题（subtitle） | 12pt | NEUTRAL_400 |
| label（section tag） | 9pt | PRIMARY_500 |
| 两列正文 | 14pt | NEUTRAL_700 |
| body_slide bullets | 17pt | NEUTRAL_700 |
| body_text 段落 | 16pt | NEUTRAL_700 |
| 卡片标题 | 16pt bold | NEUTRAL_900 |
| 卡片正文 | 13pt | NEUTRAL_700 |
| 统计数值（big_stats） | 44pt bold | accent 色 |
| 统计标签 | 14pt bold | NEUTRAL_900 |

---

## 封面 / 分隔 / 结尾页渐变标题色

`add_cover` / `add_divider*` / `add_closing` 的主标题使用三色渐变：

```
PRIMARY_500 #3EC99E → SUCCESS_HEX #5CC13C → SECONDARY_500 #C8E13C
```

普通内容页标题（`_header`）**不使用渐变**，用纯色 NEUTRAL_900。

---

## 装饰性元素（title_deco）

装饰图片存放于 `assets/装饰性元素/`，分三个系列：

| 系列 | 使用场景 |
|---|---|
| **A 系**（`A*.png`） | 白色背景普通内容页（最常用） |
| **B 系**（`B*.png`） | 灰色底或整页图片的页面 |
| **C 系**（`C*.png`） | 深色背景页 |

每个系列有四种类型：`underline`（下划）/ `circle`（画圈）/ `arrow1`（顶部对齐箭头）/ `arrow2`（底部对齐箭头）。

**使用 hard rules：**

1. **每页最多一个装饰**，禁止同一页出现两种
2. 有图文排版的页面，优先考虑 arrow1 / arrow2
3. 纯文字页或含图文但不符合箭头条件：**必须**有 underline 或 circle
4. `add_cover` / `add_divider*` / TOC / `add_closing`：**不加装饰**
5. circle 须置于标题文字**下层**，底部与基线对齐

**最小宽度规则：** 装饰始终保持约 4 个字宽，即使目标词只有 2 字。

**`title_deco` 参数和典型示例见 `reference/pptx-api.md`。**

---

## HTML 格式规则

```html
<link rel="stylesheet" href="../../../templates/html/brand.css">
<img src="../../../assets/logo-horizontal-primary.svg" height="32">
```

颜色全部用 `var(--color-primary-500)` 等 CSS 变量，**禁止内联十六进制**。

---

## 多语言（i18n）规则

### 语言模式

| 模式 | 说明 |
|---|---|
| `"cn"` | **默认**。所有字段按中文渲染，兼容现有全中文脚本 |
| `"en"` | 全英文 PPT。Header 标题取 `"en"` 字段，文本框更宽（CW×0.92），禁止换行 |
| `"bilingual"` | 中英对照。Header 主标题显示 CN；title dict 的 `"en"` 自动填入 subtitle 位置；卡片/文本块内 CN 在上、EN 在下 |

### 激活方式（job 脚本第一行）

```python
from scripts.components import configure
configure(lang="bilingual")   # "cn" | "en" | "bilingual"
```

### 文本字段格式

所有组件的文本字段接受 `str`（仅中文）或 `{"cn": ..., "en": ...}`（双语）：

```python
# 纯字符串 — cn/en 模式均可用，bilingual 时当作 CN 显示
title = "核心能力矩阵"

# 双语 dict — 三种模式均支持
title = {"cn": "核心能力矩阵", "en": "Core Capability Matrix"}
```

**支持双语字段的组件：**

| 组件 | 可传 dict 的字段 |
|---|---|
| `Card` | `title` / `body` / `tag` / `metric` / `icon` / `quote` / `attribution` |
| `TextBlock` | `content` / `bullets` 列表中每项 |
| `Callout` | `text` |
| `Stat` | `label` / `desc`（`val` 是数值，不翻译） |
| `FlowPills` | `items` 列表中每项 |

### Header 的 i18n 行为（hard rule）

```
lang="cn"        → 渲染传入的字符串或 dict["cn"]，布局不变
lang="bilingual" → 主标题显示 CN；dict["en"] 自动顶替 subtitle 槽位
lang="en"        → 主标题显示 dict["en"]；文本框宽 CW×0.92；wrap=False（单行）
```

**`add_*` 方法的 `title` 参数直接传 dict，无需额外处理：**

```python
prs.add_three_cards(
    title={"cn": "核心能力矩阵", "en": "Core Capability Matrix"},
    label="PLATFORM CAPABILITIES",   # label 始终英文大写，不需要 dict
    cards=[...]
)
```

### 全英文模式标题写法（hard rule）

- 英文标题文本框设定为**禁止换行**，超出会被截断
- **提炼关键词**，控制在 5-8 个英文单词以内
- 避免介词堆叠（`of`, `for`, `to` 等冗余词），用名词短语
- 示例：`"Data Continuity & Recovery"` ✓ / `"How We Ensure Data Is Continuously Available"` ✗

### 双语模式 subtitle 冲突规则

bilingual 模式下，title dict 的 `"en"` **优先**占用 subtitle 槽位。若该页还需要独立 subtitle，需在内容区另行渲染，不通过 `subtitle=` 参数传入。

### 完整 job 脚本示例片段

```python
from scripts.components import configure
configure(lang="bilingual")

prs.add_three_cards(
    title={"cn": "三大核心优势", "en": "Three Core Advantages"},
    label="KEY DIFFERENTIATORS",
    cards=[
        {
            "color": "primary",
            "title": {"cn": "实时数据采集", "en": "Real-time Ingestion"},
            "body":  {"cn": "秒级采集，15 分钟时间桶自动拼接，断网补齐。",
                      "en": "Second-level sampling with auto backfill on disconnect."},
        },
        ...
    ]
)
```

---

## context.md 三层优先级

```
user prompt  ← 最高优先级，仅影响当次任务
context.md   ← 项目级角色/风格/领域设定
CLAUDE.md    ← 系统硬约束（颜色、字体、路径、禁止事项），永远生效
```

- context.md 能覆盖：写作角色、语气、领域术语、受众假设、内容禁忌
- context.md **不能覆盖**：颜色值、字体名、路径规范、禁止事项表
- **AI 不得自行修改 context.md**，由项目负责人人工维护

---

## 每次输出必做：合规测试

```bash
python3 tests/test_compliance.py projects/{name}/docs/output.docx --format docx
python3 tests/test_compliance.py projects/{name}/docs/output.html --format html
```

**通过标准：总分 ≥ 80 / 100。低于 80 必须修复后重新测试。**

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
| 在 `projects/` 之外存放项目专属文件 | 破坏结构一致性 |
| 生成脚本未加正确的 sys.path.insert | 会导致 scripts 模块找不到 |
| AI 自行修改 context.md | 项目提示词由人工维护，AI 只读不写 |
| 跳过 Step 0 直接开始生成 | 无项目上下文会导致角色/风格/领域偏差 |
| 在 gen_*.py 中硬编码邮箱、官网、电话 | 应使用 `BT.COMPANY_EMAIL` / `BT.COMPANY_WEBSITE` / `BT.COMPANY_PHONE` |
| 全英文模式标题写长句/换行 | 文本框 wrap=False，超出会截断；必须提炼关键词控制在 ≤8 词 |
| bilingual 模式同时传 `subtitle=` 参数 | title dict 的 "en" 已占用 subtitle 槽，两者冲突时 EN 覆盖传入值 |
| 版式 slot 数量决定内容点数量 | 必须先从 references 确定 N，再选适合 N 的 preset；不得反向截断 |
| 为了"版面整洁"合并不同概念进同一卡片/条目 | 可精简语言，不可合并概念；内容完整性是红线 |
| 在 CLAUDE.md / brand_system.md 中手写具体色值、字体名等数值 | 会与 `tokens.json` 产生漂移（已发生过）；规则只写角色名，数值统一见 `reference/brand-values.md` |

---

## 参考材料（按需读取）

- `reference/pptx-api.md` — Pill 样式表、`_pill()` 参数、`intro_flow` 详情、`_callout()` 参数、`title_deco` 参数 + 示例、PDF Puppeteer 脚本
- `reference/setup-guide.md` — 新建项目步骤、context.md 模板、品牌定制指南
- `reference/brand-values.md` — **自动生成**的色值/字体/页面数值表，唯一数据源 `tokens.json`；改值后运行 `python3 scripts/gen_brand_reference.py` 重新生成
