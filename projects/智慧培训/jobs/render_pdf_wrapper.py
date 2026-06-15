import sys, os, subprocess
import tempfile
import shutil

_BRAND   = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
_PROJECT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, _BRAND)

_DOCS  = os.path.join(_PROJECT, 'docs')
PDF_GEN_DIR = '/tmp/pdf-gen'

def render_html_to_pdf(html_path, output_path):
    """
    使用系统 Puppeteer 将 HTML 渲染为 PDF
    """
    if not os.path.exists(html_path):
        raise FileNotFoundError(f"HTML 文件不存在: {html_path}")

    if not os.path.exists(PDF_GEN_DIR):
        raise RuntimeError(f"PDF 生成环境不存在: {PDF_GEN_DIR}")

    print(f"📄 开始渲染: {html_path} → {output_path}")

    # 创建临时渲染脚本
    render_script = f"""
import puppeteer from '/tmp/pdf-gen/node_modules/puppeteer/lib/puppeteer/puppeteer.js';
import {{ readFileSync }} from 'fs';

(async () => {{
    const browser = await puppeteer.launch({{
        headless: true,
        executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
        args: ['--no-sandbox', '--disable-setuid-sandbox'],
    }});

    try {{
        const page = await browser.newPage();
        const htmlContent = readFileSync('{html_path}', 'utf8');

        await page.setContent(htmlContent, {{
            waitUntil: 'networkidle0',
            timeout: 30000
        }});

        // 加载品牌字体
        await page.addStyleTag({{
            url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600;700&display=swap'
        }});

        await page.evaluateHandle('document.fonts.ready');
        await new Promise(resolve => setTimeout(resolve, 2000));

        await page.pdf({{
            path: '{output_path}',
            format: 'A4',
            printBackground: true,
            margin: {{ top: 0, bottom: 0, left: 0, right: 0 }},
            preferCSSPageSize: true,
        }});

        console.log('✅ PDF 生成完成');

    }} catch (error) {{
        console.error('❌ PDF 生成失败:', error.message);
        process.exit(1);
    }} finally {{
        await browser.close();
    }}
}})();
"""

    # 写入临时脚本
    script_path = os.path.join(PDF_GEN_DIR, 'temp_render.mjs')
    with open(script_path, 'w', encoding='utf-8') as f:
        f.write(render_script)

    try:
        # 执行渲染
        result = subprocess.run(
            ['node', script_path],
            cwd=PDF_GEN_DIR,
            capture_output=True,
            text=True,
            timeout=120  # 2分钟超时
        )

        if result.returncode != 0:
            print(f"❌ 渲染失败: {result.stderr}")
            raise RuntimeError(f"PDF 渲染失败: {result.stderr}")

        print(f"✅ PDF 生成成功: {output_path}")

    finally:
        # 清理临时脚本
        if os.path.exists(script_path):
            os.remove(script_path)


def test_pdf_generation():
    """测试 PDF 生成功能"""

    html_file = os.path.join(_DOCS, 'test_cover.html')
    pdf_file = os.path.join(_DOCS, 'test_cover_rendered.pdf')

    if not os.path.exists(html_file):
        print(f"⚠️  HTML 文件不存在: {html_file}")
        return False

    try:
        render_html_to_pdf(html_file, pdf_file)
        return True
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False


if __name__ == "__main__":
    print("🧪 测试 PDF 生成功能...")

    success = test_pdf_generation()

    if success:
        print("🎉 PDF 生成功能测试成功！")
    else:
        print("💥 PDF 生成功能测试失败！")
        sys.exit(1)