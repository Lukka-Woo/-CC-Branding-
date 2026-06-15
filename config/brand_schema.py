"""
品牌配置数据结构定义
用于统一管理所有品牌相关的配置参数
"""

from dataclasses import dataclass, field
from typing import Dict, Any
import json
import os

@dataclass
class BrandConfig:
    """统一品牌配置数据结构"""

    # ── 公司信息 ──────────────────────────────────────────────────────────────
    company_name_cn: str = "上海未来方舟智能科技有限公司"
    company_name_en: str = "Shanghai ArktechX Intelligent Technology Co., Ltd."
    company_short_cn: str = "未来方舟"
    company_short_en: str = "ArktechX"

    # ── 联系信息 ──────────────────────────────────────────────────────────────
    address_cn: str = "上海市浦河创业中心（桂平路）302栋"
    address_en: str = "Building 302, Puhe Entrepreneurship Center (Guiping Road), Shanghai"
    phone: str = "+86 400-123-4567"
    email: str = "contact@arktechx.com"
    website: str = "https://www.arktechx.com"
    icp_license: str = "沪ICP备12345678号"

    # ── 品牌色彩 ──────────────────────────────────────────────────────────────
    primary_color: str = "#3EC99E"      # 主品牌色
    primary_dark: str = "#4B9E31"       # 主品牌色深色变体
    primary_light: str = "#EAFAF5"      # 主品牌色浅色变体
    secondary_color: str = "#C8E13C"    # 辅助色
    secondary_light: str = "#F8FBE7"    # 辅助色浅色变体

    # ── 中性色系 ──────────────────────────────────────────────────────────────
    neutral_900: str = "#0E1216"        # 深色文字
    neutral_700: str = "#3D444A"        # 标准文字
    neutral_400: str = "#8A9199"        # 次要文字
    neutral_200: str = "#D0D5DD"        # 边框
    neutral_100: str = "#F2F3F5"        # 背景
    neutral_50: str = "#F8FAFC"         # 浅色背景
    white: str = "#FFFFFF"              # 白色

    # ── 字体设置 ──────────────────────────────────────────────────────────────
    font_cn_primary: str = "Alibaba PuHuiTi 2.0"
    font_cn_fallback: str = "PingFang SC, Microsoft YaHei, Noto Sans SC"
    font_en_primary: str = "Inter"
    font_en_fallback: str = "SF Pro, Segoe UI, Arial"
    font_mono: str = "JetBrains Mono, Consolas, monospace"

    # ── 基础字号 (pt) ─────────────────────────────────────────────────────────
    font_size_base: int = 11            # 基础字号
    font_size_small: int = 9            # 小字号
    font_size_large: int = 14           # 大字号
    font_size_h1: int = 28              # 一级标题
    font_size_h2: int = 22              # 二级标题
    font_size_h3: int = 18              # 三级标题

    # ── 间距设置 (mm) ─────────────────────────────────────────────────────────
    spacing_xs: float = 2.0             # 超小间距
    spacing_sm: float = 4.0             # 小间距
    spacing_md: float = 8.0             # 中等间距
    spacing_lg: float = 16.0            # 大间距
    spacing_xl: float = 24.0            # 超大间距

    # ── 页面布局 ──────────────────────────────────────────────────────────────
    page_margin_top: float = 25.4       # 上边距 (mm)
    page_margin_bottom: float = 25.4    # 下边距 (mm)
    page_margin_left: float = 31.7      # 左边距 (mm)
    page_margin_right: float = 31.7     # 右边距 (mm)

    # ── 圆角设置 (mm) ─────────────────────────────────────────────────────────
    border_radius_sm: float = 1.0       # 小圆角
    border_radius_md: float = 2.0       # 中圆角
    border_radius_lg: float = 4.0       # 大圆角

    # ── Logo 路径 ─────────────────────────────────────────────────────────────
    logo_horizontal_path: str = "assets/logo-horizontal-primary.png"
    logo_stacked_path: str = "assets/logo-stacked-primary.png"
    logo_mark_path: str = "assets/logo-mark-primary.png"

    # ── 装饰元素 ──────────────────────────────────────────────────────────────
    decoration_path: str = "assets/装饰性元素/3.png"

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式，便于 JSON 序列化"""
        return self.__dict__

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'BrandConfig':
        """从字典创建配置对象"""
        return cls(**data)

    @classmethod
    def from_json_file(cls, file_path: str) -> 'BrandConfig':
        """从 JSON 文件加载配置"""
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return cls.from_dict(data)

    def save_to_json(self, file_path: str):
        """保存配置到 JSON 文件"""
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    def get_css_variables(self) -> str:
        """生成 CSS 自定义属性"""
        return f"""
:root {{
  /* 品牌色彩 */
  --color-primary: {self.primary_color};
  --color-primary-dark: {self.primary_dark};
  --color-primary-light: {self.primary_light};
  --color-secondary: {self.secondary_color};
  --color-secondary-light: {self.secondary_light};

  /* 中性色 */
  --color-neutral-900: {self.neutral_900};
  --color-neutral-700: {self.neutral_700};
  --color-neutral-400: {self.neutral_400};
  --color-neutral-200: {self.neutral_200};
  --color-neutral-100: {self.neutral_100};
  --color-neutral-50: {self.neutral_50};
  --color-white: {self.white};

  /* 字体 */
  --font-cn: "{self.font_cn_primary}", {self.font_cn_fallback}, sans-serif;
  --font-en: "{self.font_en_primary}", {self.font_en_fallback}, sans-serif;
  --font-mono: {self.font_mono};

  /* 字号 */
  --font-size-base: {self.font_size_base}pt;
  --font-size-small: {self.font_size_small}pt;
  --font-size-large: {self.font_size_large}pt;
  --font-size-h1: {self.font_size_h1}pt;
  --font-size-h2: {self.font_size_h2}pt;
  --font-size-h3: {self.font_size_h3}pt;

  /* 间距 */
  --spacing-xs: {self.spacing_xs}mm;
  --spacing-sm: {self.spacing_sm}mm;
  --spacing-md: {self.spacing_md}mm;
  --spacing-lg: {self.spacing_lg}mm;
  --spacing-xl: {self.spacing_xl}mm;

  /* 圆角 */
  --radius-sm: {self.border_radius_sm}mm;
  --radius-md: {self.border_radius_md}mm;
  --radius-lg: {self.border_radius_lg}mm;
}}"""


# ── 预设品牌配置 ──────────────────────────────────────────────────────────────

def get_arktechx_config() -> BrandConfig:
    """未来方舟默认品牌配置"""
    return BrandConfig()

def get_demo_config() -> BrandConfig:
    """演示用的品牌配置"""
    return BrandConfig(
        company_name_cn="演示科技有限公司",
        company_name_en="Demo Technology Co., Ltd.",
        company_short_cn="演示科技",
        company_short_en="DemoTech",
        primary_color="#2563EB",
        secondary_color="#F59E0B",
        phone="+86 123-456-7890",
        email="hello@demo.com",
        website="https://demo.com"
    )