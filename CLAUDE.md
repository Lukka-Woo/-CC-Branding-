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
中文：Alibaba_PuHuiTi_2.0 SC / PingFang SC / Microsoft YaHei / Noto Sans SC
英文：Inter / SF Pro
代码：JetBrains Mono
```

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
from scripts.pptx_builder import BrandPptx

prs = BrandPptx()
prs.add_cover("标题", "副标题")
prs.add_divider("章节标题")
prs.add_body_slide("标题", content)
prs.save(os.path.join(_DOCS, "output.pptx"))
```
- 比例：16:9 | 封面背景：`#0E1216` | 分隔页：`#3EC99E`

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

const browser = await puppeteer.launch({
  headless: true,
  executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  args: ['--no-sandbox'],
});
const page = await browser.newPage();
await page.setContent(readFileSync(SRC, 'utf8'), { waitUntil: 'networkidle0' });
await page.addStyleTag({
  url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600;700&display=swap'
});
await page.evaluateHandle('document.fonts.ready');
await new Promise(r => setTimeout(r, 2500));
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
