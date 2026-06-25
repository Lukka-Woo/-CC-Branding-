# 项目初始化 & 品牌定制指南

> 新建项目或替换品牌时读此文件。日常生成任务不需要。

---

## 新建项目

### 1. 更新 project.json

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

### 2. 创建目录结构

```bash
mkdir -p projects/new-project/{docs/img,references,media/{covers,illustrations,photos,icons},jobs}
```

### 3. 创建 context.md（必须）

复制以下模板到 `projects/new-project/context.md`：

```markdown
# {项目名} · 项目提示词（Project Context）

> **自动加载**：本文件由系统在每次任务开始前读取，无需用户重复说明。
> 优先级：高于 CLAUDE.md 风格判断，低于 user prompt，低于 CLAUDE.md 技术硬约束。

---

## 角色定义（Role）
你是一名...

## 目标受众（Audience）
- 主要受众：...
- 次要受众：...

## 输出风格（Style）
- 语言：中文为主...
- 语气：...
- 结构：...

## 领域知识（Domain）
（本项目涉及的行业背景、核心概念、标准规范、竞品认知）

## 输出约束（Constraints）
- 不虚构...
```

### 4. 创建 media/MANIFEST.md

记录每个素材文件的用途（人工维护）：

```markdown
# 素材清单

| 文件路径 | 用途 | 出现位置 |
|---|---|---|
| covers/hero.jpg | 封面大图 | 封面页右侧 |
| photos/scene.jpg | 实景照片 | P5 右侧图片区 |
```

---

## 修改品牌信息（部署到其他公司）

### 1. 修改 brand_tokens.py

```python
# scripts/brand_tokens.py 第 54-67 行
BRAND_NAME_CN  = "你的公司简称"
BRAND_NAME_EN  = "YourCompany"
BRAND_FULL_CN  = "你的公司全称有限公司"
BRAND_FULL_EN  = "Your Company Full Name Co., Ltd."
COMPANY_ADDRESS_CN = "你的公司地址"
COMPANY_ADDRESS_EN = "Your Company Address"
COMPANY_EMAIL   = "info@yourcompany.com"
COMPANY_WEBSITE = "https://www.yourcompany.com"
COMPANY_PHONE   = "13800000000"
```

### 2. 同步更新模板文件中的硬编码文本

| 文件 | 位置 |
|---|---|
| `templates/html/base.html` | 第 78、100 行的公司名称 |
| `templates/pdf/brand.css` | 第 22 行的页眉公司名称 |
| `templates/ppt/gen_brand_master.py` | 第 250、259 行的公司名称 |

### 3. 替换 Logo 文件

放入 `assets/` 目录，保持文件名：
- `logo-horizontal-primary.png` / `.svg`
- `logo-stacked-primary.png`
- `logo-mark-primary.png`

### 4. 配置检查清单

- [ ] `brand_tokens.py` 公司信息已更新
- [ ] DOCX 文档页眉页脚显示新公司名称
- [ ] HTML 模板 Footer 显示新公司名称
- [ ] PDF 页眉显示新公司名称
- [ ] PPT 封面页显示新公司名称
- [ ] 所有 Logo 文件已替换
