#!/usr/bin/env python3
"""
品牌模板系统概念演示
展示"一次配置，全局生效"的核心思路
"""

import os
import sys
import json
from datetime import datetime

# 添加当前目录到 Python 路径
sys.path.insert(0, os.path.dirname(__file__))

from config.brand_schema import BrandConfig


def simple_template_render(template_content: str, brand: BrandConfig) -> str:
    """简化的模板渲染函数（不依赖 Jinja2）"""
    result = template_content

    # 替换品牌变量
    replacements = {
        '{{ brand.company_name_cn }}': brand.company_name_cn,
        '{{ brand.company_name_en }}': brand.company_name_en,
        '{{ brand.company_short_cn }}': brand.company_short_cn,
        '{{ brand.primary_color }}': brand.primary_color,
        '{{ brand.secondary_color }}': brand.secondary_color,
        '{{ brand.phone }}': brand.phone,
        '{{ brand.email }}': brand.email,
        '{{ brand.website }}': brand.website,
        '{{ brand.neutral_900 }}': brand.neutral_900,
        '{{ brand.neutral_700 }}': brand.neutral_700,
        '{{ brand.font_cn_primary }}': brand.font_cn_primary,
        '{{ brand.font_en_primary }}': brand.font_en_primary,
        '{{ generation_time }}': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    for placeholder, value in replacements.items():
        result = result.replace(placeholder, value)

    return result


def demonstrate_unified_config():
    """演示统一配置系统"""
    print("🎯 品牌模板系统演示")
    print("=" * 50)

    # 1. 创建三种不同的品牌配置
    configs = {
        "arktechx": BrandConfig(
            company_name_cn="上海未来方舟智能科技有限公司",
            company_short_cn="未来方舟",
            primary_color="#3EC99E",
            secondary_color="#C8E13C",
            phone="+86 400-123-4567",
            email="contact@arktechx.com",
            website="https://www.arktechx.com"
        ),
        "demo": BrandConfig(
            company_name_cn="演示科技有限公司",
            company_short_cn="演示科技",
            primary_color="#2563EB",
            secondary_color="#F59E0B",
            phone="+86 400-999-8888",
            email="hello@demo.tech",
            website="https://demo.tech"
        ),
        "creative": BrandConfig(
            company_name_cn="创意设计工作室",
            company_short_cn="创意设计",
            primary_color="#EC4899",
            secondary_color="#8B5CF6",
            phone="+86 400-777-6666",
            email="hi@creative.studio",
            website="https://creative.studio"
        )
    }

    # 2. 定义模板内容
    brand_tokens_template = '''
# 品牌配置 (自动生成于 {{ generation_time }})
BRAND_FULL_CN = "{{ brand.company_name_cn }}"
BRAND_NAME_CN = "{{ brand.company_short_cn }}"
PRIMARY_500_HEX = "{{ brand.primary_color }}"
SECONDARY_500_HEX = "{{ brand.secondary_color }}"
COMPANY_PHONE = "{{ brand.phone }}"
COMPANY_EMAIL = "{{ brand.email }}"
COMPANY_WEBSITE = "{{ brand.website }}"
'''

    html_template = '''
<footer class="brand-footer">
    <span>{{ brand.company_name_cn }}</span>
    <span>{{ brand.website }} | {{ brand.email }}</span>
</footer>
<style>
    :root {
        --color-primary: {{ brand.primary_color }};
        --color-secondary: {{ brand.secondary_color }};
    }
</style>
'''

    # 3. 演示每个配置的效果
    for config_name, brand_config in configs.items():
        print(f"\n🔄 应用 {config_name.upper()} 配置...")
        print(f"  公司名称: {brand_config.company_name_cn}")
        print(f"  主品牌色: {brand_config.primary_color}")
        print(f"  辅助色: {brand_config.secondary_color}")

        # 生成 brand_tokens.py 内容
        tokens_content = simple_template_render(brand_tokens_template, brand_config)
        print(f"\n  📄 生成的 brand_tokens.py:")
        print("  " + "\n  ".join(tokens_content.strip().split('\n')))

        # 生成 HTML 内容
        html_content = simple_template_render(html_template, brand_config)
        print(f"\n  🌐 生成的 HTML/CSS:")
        print("  " + "\n  ".join(html_content.strip().split('\n')))

        print("\n" + "-" * 40)

    print("\n✨ 演示完成！")
    print("\n💡 核心优势:")
    print("  ✓ 一次修改配置，所有文件自动更新")
    print("  ✓ 支持多套品牌配置快速切换")
    print("  ✓ 消除硬编码，降低维护成本")
    print("  ✓ 支持产品化部署，多租户隔离")


def show_config_structure():
    """展示配置数据结构"""
    print("\n📋 品牌配置数据结构:")
    print("=" * 30)

    config = BrandConfig()
    config_dict = config.to_dict()

    categories = {
        "公司信息": ["company_name_cn", "company_name_en", "company_short_cn", "company_short_en"],
        "联系信息": ["address_cn", "phone", "email", "website", "icp_license"],
        "品牌色彩": ["primary_color", "primary_dark", "secondary_color"],
        "中性色系": ["neutral_900", "neutral_700", "neutral_400", "neutral_200"],
        "字体设置": ["font_cn_primary", "font_en_primary", "font_mono"],
        "字号大小": ["font_size_h1", "font_size_h2", "font_size_base"],
        "间距设置": ["spacing_xs", "spacing_sm", "spacing_md", "spacing_lg"],
        "页面布局": ["page_margin_top", "page_margin_left"],
        "Logo路径": ["logo_horizontal_path", "logo_stacked_path"]
    }

    for category, fields in categories.items():
        print(f"\n  {category}:")
        for field in fields:
            if field in config_dict:
                value = config_dict[field]
                print(f"    {field}: {value}")


def save_example_configs():
    """保存示例配置文件"""
    print("\n💾 保存示例配置文件...")

    # 创建 config 目录
    config_dir = os.path.join(os.path.dirname(__file__), "config")
    os.makedirs(config_dir, exist_ok=True)

    # 保存不同的配置示例
    configs = {
        "default": BrandConfig(),
        "demo": BrandConfig(
            company_name_cn="演示科技有限公司",
            company_short_cn="演示科技",
            primary_color="#2563EB",
            secondary_color="#F59E0B"
        )
    }

    for name, config in configs.items():
        config_file = os.path.join(config_dir, f"{name}.json")
        config.save_to_json(config_file)
        print(f"  ✓ {config_file}")


if __name__ == "__main__":
    # 运行演示
    demonstrate_unified_config()

    # 显示配置结构
    show_config_structure()

    # 保存配置示例
    save_example_configs()

    print(f"\n🎉 演示完成！")
    print(f"📁 查看生成的配置文件: config/")
    print(f"📁 查看模板文件: templates_src/")
    print(f"\n下一步: 安装 Jinja2 后运行完整的模板系统:")
    print(f"  pip install jinja2")
    print(f"  python test_template_system.py")