# 品牌数值参考表（自动生成，请勿手动编辑）

> 本文件由 `scripts/gen_brand_reference.py` 从 `scripts/brand_tokens.py`
> （唯一数据源，追溯至 `tokens.json`）生成。
>
> 修改颜色 / 字体 / 页面数值 → 编辑 `tokens.json` → 重新运行：
> `python3 scripts/gen_brand_reference.py`
>
> `CLAUDE.md` / `brand_system.md` 中涉及具体数值的规则一律指向本文件，
> 不在多处重复硬编码同一个值。

## 颜色

| 角色 | 色值 | 用途 |
|---|---|---|
| 主绿 | `#3EC99E` | 标题、强调、Logo、表格 Header |
| 主绿深 | `#319E7C` | Hover / 深色变体 |
| 主绿浅 | `#EAFAF5` | 卡片背景、交替行 |
| 辅色 | `#C8E13C` | 次级强调、Badge |
| 中性900 | `#0E1216` | 正文标题 |
| 中性700 | `#3D444A` | 正文文字 |
| 中性400 | `#8A9199` | 说明文字、Footer |
| 中性200 | `#D0D5DD` | 边框、分隔线 |
| 中性100 | `#F2F3F5` | 斑马纹、Note 背景 |
| 白色 | `#FFFFFF` | 页面背景、反白文字 |

所有颜色常量通过 `import scripts.brand_tokens as BT` 使用，禁止在脚本中内联十六进制。

## 字体

| 用途 | 字体族 |
|---|---|
| 中文（主 / 降级） | Alibaba PuHuiTi 2.0 / PingFang SC / Microsoft YaHei / Noto Sans SC（Web） |
| 英文 | Arial |
| 代码 | JetBrains Mono |

PPTX/DOCX 写字体名固定引用 `BT.FONT_CN` / `BT.FONT_EN`；
HTML/PDF 通过 `@font-face` 加载本地文件（`fonts/alibaba-puhuiti/`）。

## DOCX 页面（A4）

| 项 | 值 |
|---|---|
| 上下边距 | 25.4mm |
| 左右边距 | 31.7mm |
| H1 | 28pt `#0E1216` |
| H2 | 22pt `#0E1216` |
| H3 | 18pt `#3EC99E` |

## 卡片颜色语义（全系统统一）

适用于 `add_three_cards`、`add_six_cards`、`add_big_stats`、`add_module_grid`。

| 底色 | Accent 色 | 语义 | 优先级 |
|---|---|---|---|
| `#EAFAF5` PRIMARY_100 | `#3EC99E` PRIMARY_500 | 标准/主要特性 | ★★★ |
| `#F2F3F5` NEUTRAL_100 | `#5CC13C` SUCCESS | 安全/补充/次要 | ★★★ |
| `#F8FBE7` SECONDARY_100 | `#C8E13C` SECONDARY_500 | 创新/机遇 | ★★★ |
| `#FFF1DF` CARD_ORANGE_BG | `#FFB928` WARNING | 高风险/注意 | ★★★ |
| `#E0F7FA` CARD_TEAL_BG | `#3CC5CF` TEAL | 扩展/生态 | ★★ |
| `#F0E8FF` CARD_PURPLE_BG | `#8255E1` PURPLE | 战略/特殊 | ★★ |
| `#0E1216` NEUTRAL_900 | `#C8E13C` SECONDARY_500 + 白字 | 突出/亮点（非危险） | ★★★ 深色卡 |
| `#FFF2F2` CARD_DANGER_BG | `#F12D2D` DANGER | 危险/风险（确实危险才用） | ★ 慎用 |

## 圆角（PPTX / DOCX）

| 常量 | 值 | 用途 |
|---|---|---|
| RADIUS_XS_MM | 2mm | 细节圆角（角标、小装饰点） |
| RADIUS_SM_MM | 4mm | ★ 默认：所有普通卡片、按钮 |
| RADIUS_MD_MM | 6mm | 大卡片、面板 |
| RADIUS_LG_MM | 8mm | 外层大容器 / 图片占位区块 |
| RADIUS_PILL_MM | pill | 胶囊标签 / 序号徽章 / 圆形点 |
