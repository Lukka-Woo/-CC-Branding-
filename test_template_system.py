#!/usr/bin/env python3
"""
测试新的品牌模板系统
演示"一次配置，全局生效"的功能
"""

import os
import sys
from datetime import datetime

# 添加当前目录到 Python 路径
sys.path.insert(0, os.path.dirname(__file__))

from config.brand_schema import BrandConfig, get_demo_config
from config.brand_manager import BrandManager


def test_brand_template_system():
    """测试品牌模板系统"""
    print("🚀 测试品牌模板系统")
    print("=" * 50)

    # 1. 创建品牌管理器
    brand_root = os.path.dirname(__file__)
    manager = BrandManager(brand_root)

    # 创建模板目录结构
    manager.create_template_structure()

    # 2. 测试默认配置
    print("\n📋 测试默认配置...")
    default_config = manager.load_config("default")
    print(f"  公司名称: {default_config.company_name_cn}")
    print(f"  主品牌色: {default_config.primary_color}")

    # 应用默认配置
    manager.apply_brand_config(default_config)

    # 3. 创建演示配置
    print("\n🎨 创建演示配置...")
    demo_config = BrandConfig(
        company_name_cn="演示科技有限公司",
        company_name_en="Demo Technology Co., Ltd.",
        company_short_cn="演示科技",
        company_short_en="DemoTech",
        address_cn="北京市朝阳区演示大厦1001室",
        address_en="Room 1001, Demo Building, Chaoyang District, Beijing",
        phone="+86 400-999-8888",
        email="hello@demo.tech",
        website="https://demo.tech",
        icp_license="京ICP备87654321号",
        primary_color="#2563EB",      # 蓝色主题
        primary_dark="#1E40AF",
        primary_light="#DBEAFE",
        secondary_color="#F59E0B",    # 橙色辅助
        secondary_light="#FEF3C7",
        font_cn_primary="Microsoft YaHei",
        font_en_primary="Arial",
    )

    # 保存演示配置
    manager.save_config(demo_config, "demo")

    # 4. 切换到演示配置并应用
    print("\n🔄 切换到演示配置...")
    manager.switch_config("demo")

    # 5. 验证生成的文件
    print("\n✅ 验证生成的文件...")

    # 检查 brand_tokens.py
    tokens_file = os.path.join(brand_root, "scripts", "brand_tokens.py")
    if os.path.exists(tokens_file):
        print(f"  ✓ {tokens_file}")
        # 验证内容
        with open(tokens_file, 'r', encoding='utf-8') as f:
            content = f.read()
            if "演示科技有限公司" in content:
                print("    ✓ 公司名称已正确更新")
            if "#2563EB" in content:
                print("    ✓ 主品牌色已正确更新")
    else:
        print(f"  ❌ {tokens_file} 不存在")

    # 检查 HTML 模板
    html_file = os.path.join(brand_root, "templates", "html", "base.html")
    if os.path.exists(html_file):
        print(f"  ✓ {html_file}")
        with open(html_file, 'r', encoding='utf-8') as f:
            content = f.read()
            if "演示科技有限公司" in content:
                print("    ✓ HTML 模板中的公司名称已正确更新")
    else:
        print(f"  ❌ {html_file} 不存在")

    # 检查 CSS 文件
    css_file = os.path.join(brand_root, "templates", "html", "brand.css")
    if os.path.exists(css_file):
        print(f"  ✓ {css_file}")
        with open(css_file, 'r', encoding='utf-8') as f:
            content = f.read()
            if "#2563EB" in content:
                print("    ✓ CSS 中的主品牌色已正确更新")
    else:
        print(f"  ❌ {css_file} 不存在")

    # 6. 显示配置对比
    print("\n📊 配置对比:")
    print(f"  默认配置 → 演示配置")
    print(f"  {default_config.company_name_cn} → {demo_config.company_name_cn}")
    print(f"  {default_config.primary_color} → {demo_config.primary_color}")
    print(f"  {default_config.phone} → {demo_config.phone}")

    # 7. 测试导入生成的 tokens
    print("\n🔧 测试导入生成的 brand_tokens...")
    try:
        sys.path.insert(0, os.path.join(brand_root, "scripts"))
        import brand_tokens as BT
        print(f"  ✓ 导入成功")
        print(f"    公司名称: {BT.BRAND_FULL_CN}")
        print(f"    主品牌色: {BT.PRIMARY_500_HEX}")
        print(f"    联系电话: {BT.COMPANY_PHONE}")
        print(f"    网站地址: {BT.COMPANY_WEBSITE}")
    except Exception as e:
        print(f"  ❌ 导入失败: {e}")

    print("\n🎉 测试完成！")
    print("\n💡 下一步:")
    print("  1. 修改 config/demo.json 中的配置")
    print("  2. 运行 manager.switch_config('demo') 应用更改")
    print("  3. 所有文档将自动使用新的品牌配置")


def demonstrate_one_config_global_effect():
    """演示一次配置全局生效的效果"""
    print("\n🎯 演示: 一次配置，全局生效")
    print("=" * 40)

    brand_root = os.path.dirname(__file__)
    manager = BrandManager(brand_root)

    # 创建三种不同的品牌配置
    configs = {
        "tech": BrandConfig(
            company_name_cn="科技创新有限公司",
            company_short_cn="科技创新",
            primary_color="#10B981",  # 绿色
            secondary_color="#F59E0B",
            font_cn_primary="Microsoft YaHei"
        ),
        "finance": BrandConfig(
            company_name_cn="金融服务股份有限公司",
            company_short_cn="金融服务",
            primary_color="#3B82F6",  # 蓝色
            secondary_color="#8B5CF6",
            font_cn_primary="SimHei"
        ),
        "creative": BrandConfig(
            company_name_cn="创意设计工作室",
            company_short_cn="创意设计",
            primary_color="#EC4899",  # 粉色
            secondary_color="#F59E0B",
            font_cn_primary="Microsoft YaHei UI"
        )
    }

    for name, config in configs.items():
        print(f"\n🔄 应用 {name} 配置...")
        manager.save_config(config, name)
        manager.switch_config(name)

        # 验证效果
        try:
            # 重新导入 brand_tokens 以获取最新值
            if 'brand_tokens' in sys.modules:
                del sys.modules['brand_tokens']

            sys.path.insert(0, os.path.join(brand_root, "scripts"))
            import brand_tokens as BT

            print(f"  ✓ 公司名称: {BT.BRAND_FULL_CN}")
            print(f"  ✓ 主品牌色: {BT.PRIMARY_500_HEX}")
            print(f"  ✓ 联系电话: {BT.COMPANY_PHONE}")

        except Exception as e:
            print(f"  ❌ 验证失败: {e}")

    print("\n✨ 演示完成！成功实现一次配置，全局生效！")


if __name__ == "__main__":
    try:
        # 添加 Jinja2 过滤器
        import jinja2
        from config.brand_manager import BrandManager

        # 为模板添加自定义过滤器
        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip('#')
            return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

        # 测试主流程
        test_brand_template_system()

        # 演示配置切换效果
        demonstrate_one_config_global_effect()

    except ImportError as e:
        print(f"❌ 缺少依赖: {e}")
        print("💡 请安装: pip install jinja2")
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()