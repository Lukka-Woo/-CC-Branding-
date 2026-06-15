import sys, os

_BRAND   = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
_PROJECT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, _BRAND)

_DOCS  = os.path.join(_PROJECT, 'docs')
_MEDIA = os.path.join(_PROJECT, 'media')

os.makedirs(_DOCS, exist_ok=True)

from scripts.docx_builder import BrandDocx
from scripts.cover_builder import CoverConfig, create_docx_cover, create_html_cover


def test_docx_layouts():
    """测试不同布局类型的 DOCX 封面"""

    test_configs = [
        {
            'title': '企业培训平台功能增强建议书',
            'subtitle': 'Enterprise Training Platform Enhancement Proposal',
            'doc_type': 'SOLUTION PROPOSAL',
            'layout_type': 'business',
            'version': 'V1.0',
            'date': '2026年6月',
            'client': '维苏威高级陶瓷（中国）有限公司',
            'classification': '机密文件，仅供内部参考'
        },
        {
            'title': 'AI 智能培训系统技术架构设计文档',
            'subtitle': 'Technical Architecture Design for AI-Powered Training System',
            'doc_type': 'TECHNICAL SPECIFICATION',
            'layout_type': 'technical',
            'version': 'V2.0',
            'date': '2026-06-15',
            'client': '技术团队',
            'classification': 'INTERNAL'
        },
        {
            'title': '智慧培训平台解决方案',
            'subtitle': 'Smart Training Platform Solution Overview',
            'doc_type': 'PRESENTATION',
            'layout_type': 'presentation',
            'version': 'V1.2',
            'date': '2026年6月15日',
            'client': '潜在客户',
            'classification': 'CONFIDENTIAL'
        }
    ]

    for i, config_data in enumerate(test_configs):
        print(f"生成测试封面 {i+1}: {config_data['layout_type']} 布局")

        # 创建配置
        config = CoverConfig(**config_data)

        # 生成 DOCX
        doc = BrandDocx(doc_type=config_data['doc_type'])
        create_docx_cover(doc, config)

        # 添加一些测试内容
        doc.add_heading("一、项目背景", level=1)
        doc.add_body("这是测试内容，用于验证封面后的正常页眉页脚是否正确显示。")

        output_path = os.path.join(_DOCS, f'test_cover_{config_data["layout_type"]}.docx')
        doc.save(output_path)
        print(f"  → 已保存: {output_path}")


def test_html_covers():
    """测试 HTML 封面生成"""

    config = CoverConfig(
        title='企业培训平台增强建议书',
        subtitle='Enterprise Training Platform Enhancement Proposal',
        doc_type='SOLUTION PROPOSAL',
        layout_type='business',
        version='V1.0',
        date='2026年6月',
        client='维苏威高级陶瓷（中国）有限公司',
        classification='机密文件，仅供内部参考'
    )

    print("生成 HTML 封面...")
    html_content = create_html_cover(config)

    output_path = os.path.join(_DOCS, 'test_cover.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"  → 已保存: {output_path}")


def demonstrate_adaptive_features():
    """演示自适应功能"""

    print("\n=== 布局自适应功能演示 ===")

    # 测试不同标题长度的自适应
    title_tests = [
        "短标题",
        "中等长度的项目建议书标题",
        "这是一个很长很长的项目标题用来测试字体大小自动调整功能是否正常工作",
    ]

    for i, title in enumerate(title_tests):
        config = CoverConfig(
            title=title,
            subtitle="自适应测试副标题",
            doc_type="TEST",
            layout_type="business"
        )

        doc = BrandDocx()
        create_docx_cover(doc, config)
        doc.add_body("测试内容")

        output_path = os.path.join(_DOCS, f'adaptive_test_{i+1}.docx')
        doc.save(output_path)
        print(f"标题长度 {len(title)} 字符 → {output_path}")


if __name__ == "__main__":
    print("开始测试自适应封面生成...")

    test_docx_layouts()
    print()
    test_html_covers()
    print()
    demonstrate_adaptive_features()

    print("\n✅ 所有测试完成！")
    print(f"📁 测试文件保存在: {_DOCS}")