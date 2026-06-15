"""
品牌配置管理器
负责加载、应用和渲染品牌配置到整个系统
"""

import os
import json
from pathlib import Path
from typing import Optional, Dict, Any
from jinja2 import Environment, FileSystemLoader, Template

from .brand_schema import BrandConfig


class BrandManager:
    """品牌配置管理器"""

    def __init__(self, brand_root: str):
        self.brand_root = Path(brand_root)
        self.config_dir = self.brand_root / "config"
        self.templates_dir = self.brand_root / "templates_src"  # 源模板目录
        self.output_dir = self.brand_root  # 输出到品牌根目录

        # 确保目录存在
        self.config_dir.mkdir(exist_ok=True)
        self.templates_dir.mkdir(exist_ok=True)

        # 设置 Jinja2 环境
        self.jinja_env = Environment(
            loader=FileSystemLoader(str(self.templates_dir)),
            trim_blocks=True,
            lstrip_blocks=True
        )

        # 当前激活的品牌配置
        self._current_config: Optional[BrandConfig] = None

    def load_config(self, config_name: str = "default") -> BrandConfig:
        """加载品牌配置"""
        config_file = self.config_dir / f"{config_name}.json"

        if config_file.exists():
            self._current_config = BrandConfig.from_json_file(str(config_file))
        else:
            # 使用默认配置
            self._current_config = BrandConfig()
            # 保存默认配置
            self.save_config(self._current_config, config_name)

        return self._current_config

    def save_config(self, config: BrandConfig, config_name: str = "default"):
        """保存品牌配置"""
        config_file = self.config_dir / f"{config_name}.json"
        config.save_to_json(str(config_file))
        print(f"✅ 品牌配置已保存: {config_file}")

    def get_current_config(self) -> BrandConfig:
        """获取当前激活的配置"""
        if self._current_config is None:
            return self.load_config()
        return self._current_config

    def apply_brand_config(self, config: Optional[BrandConfig] = None):
        """应用品牌配置到整个系统"""
        if config is None:
            config = self.get_current_config()

        print("🎨 开始应用品牌配置...")

        # 渲染所有模板文件
        self._render_brand_tokens(config)
        self._render_html_templates(config)
        self._render_css_templates(config)
        self._render_python_templates(config)

        print("✅ 品牌配置应用完成！")

    def _render_brand_tokens(self, config: BrandConfig):
        """渲染 brand_tokens.py 文件"""
        template_file = "brand_tokens.py.j2"
        output_file = self.output_dir / "scripts" / "brand_tokens.py"

        if self._render_template(template_file, output_file, {"brand": config}):
            print(f"  ✓ 已生成: {output_file}")

    def _render_html_templates(self, config: BrandConfig):
        """渲染 HTML 模板文件"""
        templates = [
            ("html/base.html.j2", "templates/html/base.html"),
        ]

        for template_file, output_path in templates:
            output_file = self.output_dir / output_path
            if self._render_template(template_file, output_file, {"brand": config}):
                print(f"  ✓ 已生成: {output_file}")

    def _render_css_templates(self, config: BrandConfig):
        """渲染 CSS 模板文件"""
        templates = [
            ("css/brand.css.j2", "templates/html/brand.css"),
            ("css/pdf.css.j2", "templates/pdf/brand.css"),
        ]

        for template_file, output_path in templates:
            output_file = self.output_dir / output_path
            if self._render_template(template_file, output_file, {"brand": config}):
                print(f"  ✓ 已生成: {output_file}")

    def _render_python_templates(self, config: BrandConfig):
        """渲染 Python 脚本模板"""
        templates = [
            ("python/ppt_gen_brand_master.py.j2", "templates/ppt/gen_brand_master.py"),
        ]

        for template_file, output_path in templates:
            output_file = self.output_dir / output_path
            if self._render_template(template_file, output_file, {"brand": config}):
                print(f"  ✓ 已生成: {output_file}")

    def _render_template(self, template_path: str, output_path: Path, context: Dict[str, Any]) -> bool:
        """渲染单个模板文件"""
        try:
            template = self.jinja_env.get_template(template_path)
            rendered_content = template.render(**context)

            # 确保输出目录存在
            output_path.parent.mkdir(parents=True, exist_ok=True)

            # 写入文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(rendered_content)

            return True

        except Exception as e:
            print(f"  ❌ 渲染失败 {template_path}: {e}")
            return False

    def list_configs(self) -> list:
        """列出所有可用的品牌配置"""
        config_files = list(self.config_dir.glob("*.json"))
        return [f.stem for f in config_files]

    def switch_config(self, config_name: str):
        """切换品牌配置并应用"""
        print(f"🔄 切换品牌配置: {config_name}")
        config = self.load_config(config_name)
        self.apply_brand_config(config)
        return config

    def create_template_structure(self):
        """创建模板文件结构"""
        template_dirs = [
            "html",
            "css",
            "python",
        ]

        for dir_name in template_dirs:
            (self.templates_dir / dir_name).mkdir(exist_ok=True)

        print(f"📁 模板目录结构已创建: {self.templates_dir}")


# ── 便捷函数 ──────────────────────────────────────────────────────────────────

def setup_brand_system(brand_root: str, config_name: str = "default") -> BrandManager:
    """快速设置品牌系统"""
    manager = BrandManager(brand_root)
    manager.create_template_structure()
    manager.switch_config(config_name)
    return manager


def quick_apply_config(brand_root: str, config: BrandConfig):
    """快速应用品牌配置"""
    manager = BrandManager(brand_root)
    manager.apply_brand_config(config)
    return manager


if __name__ == "__main__":
    # 测试代码
    import sys

    if len(sys.argv) > 1:
        brand_root = sys.argv[1]
    else:
        brand_root = os.path.dirname(os.path.dirname(__file__))

    print(f"🎯 初始化品牌管理系统: {brand_root}")
    manager = setup_brand_system(brand_root)

    print("\n📋 当前品牌配置:")
    config = manager.get_current_config()
    print(f"  公司名称: {config.company_name_cn}")
    print(f"  主品牌色: {config.primary_color}")
    print(f"  辅助色: {config.secondary_color}")