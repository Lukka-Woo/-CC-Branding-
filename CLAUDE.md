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

### Step 1 — 扫描 references/

```bash
ls projects/{active}/references/
```

PDF 用 `pymupdf`（`fitz`）提取文字和图片；图片视觉阅读；文本直接读取。
这些文件**不嵌入**输出，用于指导内容策划和文案方向。

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

## 核心数值（禁止硬编码其他值）

### 颜色

```
主绿    #3EC99E  — 标题、强调、Logo、表格 Header
主绿深  #4B9E31  — Hover / 深色变体
主绿浅  #EAFAF5  — 卡片背景、交替行
辅色    #C8E13C  — 次级强调、Badge
中性900 #0E1216  — 正文标题
中性700 #3D444A  — 正文文字
中性400 #8A9199  — 说明文字、Footer
中性200 #D0D5DD  — 边框、分隔线
中性100 #F2F3F5  — 斑马纹、Note 背景
白色    #FFFFFF  — 页面背景、反白文字
```

所有颜色常量通过 `import scripts.brand_tokens as BT` 使用，禁止在脚本中内联十六进制。

### 字体

```
中文：Alibaba PuHuiTi 2.0 / PingFang SC / Microsoft YaHei / Noto Sans SC
英文：Inter / SF Pro
代码：JetBrains Mono
```

PPTX/DOCX 写字体名 `"Alibaba PuHuiTi 2.0"`；HTML/PDF 通过 `@font-face` 加载本地文件（`fonts/alibaba-puhuiti/`）。

### DOCX 页面（A4）

```
上下边距：25.4mm   左右边距：31.7mm
H1：28pt #0E1216   H2：22pt #0E1216   H3：18pt #3EC99E
```

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

| 底色 | Accent 色 | 语义 | 优先级 |
|---|---|---|---|
| `#EAFAF5` PRIMARY_100 | `#3EC99E` PRIMARY_500 | 标准/主要特性 | ★★★ |
| `#F2F3F5` NEUTRAL_100 | `#5CC13C` SUCCESS | 安全/补充/次要 | ★★★ |
| `#F8FBE7` SECONDARY_100 | `#C8E13C` SECONDARY_500 | 创新/机遇 | ★★★ |
| `#FFF1DF` CARD_ORANGE_BG | `#FFB928` WARNING | 高风险/注意 | ★★★ |
| `#E0F7FA` CARD_TEAL_BG | `#3CC5CF` TEAL | 扩展/生态 | ★★ |
| `#F0E8FF` CARD_PURPLE_BG | `#8255E1` PURPLE | 战略/特殊 | ★★ |
| `#0E1216` NEUTRAL_900 | `#C8E13C` + 白字 | **突出/亮点**（非危险） | ★★★ 深色卡 |
| `#FFF2F2` CARD_DANGER_BG | `#F12D2D` DANGER | **危险/风险**（确实危险才用） | ★ 慎用 |

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
| 4 或 6 个并列 | `add_six_cards` |
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
| 4/6 卡片，4/6/8 模块 | — | 直接传，无需 intro |

**intro slot 内容参数（可并用）：**
- `intro_label`：小标题，8pt 绿色，≤8 字
- `intro_text`：2-3 句概括，不重复标题
- `intro_flow`：字符串列表，渲染为 pill→箭头→pill

**intro slot 视觉规则（hard rule）：** 无底色、无边框、无左侧装饰矩形，直接在白色背景上渲染。

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

---

## 参考材料（按需读取）

- `reference/pptx-api.md` — Pill 样式表、`_pill()` 参数、`intro_flow` 详情、`_callout()` 参数、`title_deco` 参数 + 示例、PDF Puppeteer 脚本
- `reference/setup-guide.md` — 新建项目步骤、context.md 模板、品牌定制指南
