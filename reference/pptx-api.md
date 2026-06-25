# PPTX API 参考手册

> 写 PPTX 生成脚本前读此文件。本文件记录参数细节和代码示例——行为规则和选型逻辑在 `CLAUDE.md`。

---

## Pill 设计系统

### `_PILL_STYLES` 样式表

| 样式名 | 背景色 | 文字色 | 适用语义 |
|---|---|---|---|
| `primary` | `#3EC99E` | 白色 | 主要/标准/确认 |
| `secondary` | `#C8E13C` | 深色 | 创新/机遇/次级 |
| `success` | `#5CC13C` | 白色 | 成功/通过/安全 |
| `warning` | `#FFB928` | 深色 | 注意/警告/待处理 |
| `teal` | `#3CC5CF` | 白色 | 扩展/生态 |
| `purple` | `#8255E1` | 白色 | 战略/特殊 |
| `danger` | `#F12D2D` | 白色 | 危险/严重错误 |
| `dark` | `#0E1216` | 白色 | 强调/亮点 |
| `primary-soft` | `#EAFAF5` | `#3EC99E` | 流程步骤（默认） |
| `secondary-soft` | `#F8FBE7` | `#C8E13C` | 次级步骤 |
| `neutral` | `#F2F3F5` | `#3D444A` | 中性标签/分类 |
| `neutral-dark` | `#D0D5DD` | `#3D444A` | 更深中性 |

### Pill API 用法

```python
# 1. 单个 pill（module-level helper，在 job 脚本中通常不直接调用）
_pill(slide, "标签文字", l=x, t=y, bg=BT.PRIMARY_500_HEX, fg=BT.WHITE_HEX, font_sz=8)
# 返回 pill 宽度（EMU），用于布局链接：next_x = l + _pill(...) + Mm(3)

# 2. 有名样式（推荐）
bg, fg = _PILL_STYLES["primary-soft"]
_pill(slide, "状态A", l=x, t=y, bg=bg, fg=fg, font_sz=9)

# 3. 带箭头流（intro_flow 参数）— 最常用
prs.add_module_grid(..., intro_flow=["步骤A", "步骤B", "步骤C"])

# 4. body_slide 区块导航 bullets
prs.add_body_slide(
    title="...",
    bullets=[
        {"pill": "已完成", "text": "PDF/HTML 解析 + SHA-256 存证", "style": "success"},
        {"pill": "进行中", "text": "AI 三层摘要调优，目标准确率 ≥90%", "style": "warning"},
        {"pill": "计划中", "text": "条款结构化抽取 + 版本比对引擎", "style": "neutral"},
        "普通文字 bullet（纯字符串，无 pill）",
    ]
)
```

### `_pill()` 参数

| 参数 | 默认值 | 说明 |
|---|---|---|
| `text` | 必填 | 标签文字（建议 ≤6 字，最多 12 字） |
| `l`, `t` | 必填 | 左上角坐标（EMU） |
| `bg` | PRIMARY_500 | pill 背景色 hex |
| `fg` | 自动 | 文字色（省略时自动选白/深） |
| `font_sz` | `8` | 字号（pt），建议 7–11 pt |
| `bold` | `True` | 是否加粗 |
| `h` | `Mm(7)` | pill 高度 |
| `max_w` | `None` | 最大宽度（防止溢出） |

### `intro_flow` / `_flow_pills` 详细说明

**触发方式：** 在 `add_six_cards` / `add_module_grid` 传 `intro_flow=["步骤A", ...]`，自动渲染为 pill→箭头→pill 流。

**内部参数（`_flow_pills` helper）：**

| 参数 | 默认值 | 说明 |
|---|---|---|
| `items` | 必填 | 步骤名称字符串列表 |
| `pill_h` | `Mm(8)` | pill 高度 |
| `font_sz` | `9` | 字号（pt） |
| `pill_bg` | `#EAFAF5` | pill 背景色（primary-soft） |
| `pill_fg` | `#3EC99E` | pill 文字色 |
| `arr_c` | `#8A9199` | 箭头颜色 |

**布局行为：**
- 每个 item 渲染为圆角 pill，item 间插入"→"文字
- 自动换行：超出 `max_w` 时折到下一行（行首不加箭头）
- pill 宽度随文字长度自适应（CJK ≈ 4.0mm/字 × font_sz/9）

---

## Callout API

### `_callout()` 参数

| 参数 | 默认值 | 说明 |
|---|---|---|
| `text` | 必填 | 注解文字（建议 ≤40 字，1行） |
| `l`, `t`, `w` | 必填 | 位置和宽度（EMU） |
| `label` | `None` | 覆盖默认 pill 标签文字；`""` 不显示 pill |
| `style` | `"note"` | `"note"` / `"info"` / `"tip"` / `"warning"` / `"danger"` |
| `font_sz` | `10` | 文字字号（pt） |

### 使用示例

```python
# 所有支持 note= 的方法（add_body_slide / add_two_col_slide / add_big_stats / add_timeline / add_table_slide）：
prs.add_table_slide(
    ...,
    note="Phase 2 / Phase 3 启动条件：上一期验收通过后方可立项。",
    note_style="warning",  # 省略时默认 "note"
)

prs.add_timeline(
    ...,
    note="数据截止 2025-12，实际影响以官方最新版本为准。",
)
```

**视觉规则：** 固定高度 `CALLOUT_H = Mm(12)`，圆角矩形背景 + 左侧 pill + 文字。不加边框、不加左侧装饰线。

---

## `title_deco` 装饰参数

所有 `add_*` 方法均支持 `title_deco=` 参数。

```python
title_deco = {
    "type":       "underline",  # "underline" | "circle" | "arrow1" | "arrow2"
    "char_start": 7,            # 目标词在标题中的起始字符偏移（从 0 开始）
    "char_count": 4,            # 目标词字符数（不足 4 时系统自动扩展到 4 字宽）
    "char_w_mm":  8.1,          # 可选，覆盖默认字宽（默认 8.1mm，26pt bold CJK）
    "series":     "a",          # 可选，默认 "a"；深色页 "c"，图片背景页 "b"
}
```

**字符宽度估算（26pt bold）：**
- 全宽 CJK 字符：≈ 8.1 mm/字
- 半宽数字/字母（如"8"）：≈ 4.0 mm/字
- 间隔符" · "：≈ 2 字宽（约 16 mm）

**典型示例：**

```python
# 下划线：强调末尾4字（最常见）
prs.add_three_cards(title="工厂培训管理的三大困局", ...,
    title_deco={"type": "underline", "char_start": 7, "char_count": 4})  # 三大困局

# 画圈：陈述句突出核心答案
prs.add_body_slide(title="我们的建议：方案 B 全员覆盖", ...,
    title_deco={"type": "circle", "char_start": 11, "char_count": 4})   # 全员覆盖

# 画圈：设问句答案（半宽"8"约0.5字宽）
prs.add_timeline(title="8 周标准交付，无需 IT 团队介入", ...,
    title_deco={"type": "circle", "char_start": 0, "char_count": 2})    # 8 周

# " · " 分隔符处理（约占2字宽偏移）
prs.add_table_slide(title="两种方案 · 清晰对比", ...,
    title_deco={"type": "underline", "char_start": 6, "char_count": 4}) # 清晰对比
```

---

## PDF 输出（Puppeteer 方案）

```javascript
// projects/{name}/jobs/render.mjs
import puppeteer from 'puppeteer';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const FONTS_DIR = path.resolve(__dirname, '..', '..', '..', 'fonts', 'alibaba-puhuiti');

const browser = await puppeteer.launch({
  headless: true,
  executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  args: ['--no-sandbox', '--allow-file-access-from-files'],
});
const page = await browser.newPage();

// 注入本地字体（必须在 setContent 之前）
await page.addStyleTag({ content: `
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:100;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_35_Thin_35_Thin.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:300;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_45_Light_45_Light.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:400;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_55_Regular_55_Regular.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:500;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_65_Medium_65_Medium.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:600;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_75_SemiBold_75_SemiBold.ttf") format("truetype"); }
  @font-face { font-family:"Alibaba PuHuiTi 2.0"; font-weight:700;
    src:url("file://${FONTS_DIR}/Alibaba_PuHuiTi_2.0_55_Regular_85_Bold.ttf") format("truetype"); }
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

**依赖安装（仅首次）：**
```bash
cd /tmp/pdf-gen && npm install puppeteer
```

**字体注入说明：** 用 `file://` 绝对路径注入 `@font-face`，在 `setContent` 之前执行，避免 Puppeteer 无 baseURL 时相对路径失效。
