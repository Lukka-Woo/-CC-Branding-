# Brand Design System

## 定位

公司名：未来方舟（上海）科技有限公司
公司英文名：Arksus

面向企业客户的 AI 智能化解决方案。视觉传达专业、可信、可持续——像咨询公司与科技公司的交集。

| 维度 | 定位 |
|------|------|
| 情感 | 冷静、理性、有温度 |
| 视觉 | 极简、克制、几何秩序感 |
| 隐喻 | 数据流、低碳循环、安全边界 |

### 关键词

```
智能进化  绿色计算  零信任架构  碳智能  数据流动性
边缘推理  模型安全  能耗优化  决策智能  合规可视化
```

## 设计原则

1. **克制即高级** —— 留白是设计的一部分，颜色遵循 60–30–10 法则
2. **一致性压倒一切** —— 圆角、字阶、间距严格统一
3. **数据不说谎** —— 可视化清晰、严谨、无 3D 装饰
4. **可持续** —— 深色模式减少能耗，字体兼顾性能与美观

## 文件映射

| 文件 | 职责 |
|------|------|
| `tokens.json` | 色彩 / 字体 / 间距 / 圆角 / 阴影等全部设计令牌 |
| `assets/logo-horizontal-primary.svg` | 横向 Logo，浅色背景使用 |
| `assets/logo-horizontal-reverse.svg` | 横向 Logo，深色背景使用 |
| `assets/logo-stacked-primary.svg` |  stacked Logo，浅色背景使用 |
| `assets/logo-stacked-reverse.svg` |  stacked Logo，深色背景使用 |
| `templates/pptx/brand-master.pptx` | PPT 封面/内页/过渡页版式 |
| `templates/docx/proposal-template.docx` | 提案文档标题/正文/表格样式 |
| `templates/pdf/brand.css` | PDF 输出的页眉页脚/分页/排版 |

## Logo System

We have two logo layouts and two color variants.

### Layout Variants

#### Horizontal Logo

Files:
- `logo-horizontal-primary.svg`
- `logo-horizontal-reverse.svg`

Use the horizontal logo in compact horizontal spaces:
- page headers
- page footers
- PPT slide footer
- PDF footer
- Word document header
- small brand placement areas

Do not use the horizontal logo as the main hero logo on covers or section divider pages unless space is limited.

#### Stacked Logo

Files:
- `logo-stacked-primary.svg`
- `logo-stacked-reverse.svg`

Use the stacked logo in large brand emphasis areas:
- cover pages
- section divider pages
- closing pages
- large empty brand areas

Do not use the stacked logo in footers, headers, or tight horizontal spaces.

### Color Variants

#### Primary

Use `primary` logos on:
- white background
- light background
- light gray background
- pale brand background

#### Reverse

Use `reverse` logos on:
- dark background
- brand primary color background
- black background
- navy background
- dark gradient background

### Selection Rule

First choose layout by space:
- compact horizontal area → horizontal
- large brand emphasis area → stacked

Then choose color variant by background:
- light background → primary
- dark background → reverse

## 审美规则速查

- **主色**：智慧绿（详见 `tokens.json`）——象征 AI 与可持续发展的交集
- **辅助色**：安全蓝 —— 传递信任与数据安全
- **字体**：Inter（英文）+ Noto Sans SC（中文）+ JetBrains Mono（代码）
- **图形**：仅使用圆角矩形（rounded corners 一致）、极细线框、几何噪声背景
- **插图**：抽象几何构成，不使用具象人物/物体插画


> 各模板的具体实现以对应文件为准，本规范不做重复详述。
