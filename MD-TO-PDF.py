import os
import markdown
import pdfkit
from pathlib import Path
import logging
import shutil


class MarkdownToPdfConverter:
    def __init__(self):
        """
        初始化转换器，使用固定的输入输出路径
        """
        self.input_dir = Path(r"E:\md")
        self.output_dir = Path(r"E:\pdf")
        self.temp_dir = Path('temp_html')

        # 设置日志
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)

        # 创建输出目录
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.temp_dir.mkdir(exist_ok=True)

        # 配置wkhtmltopdf路径
        self.wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
        if not os.path.exists(self.wkhtmltopdf_path):
            print("错误: 未找到wkhtmltopdf，请先安装wkhtmltopdf！")
            print("下载地址: https://wkhtmltopdf.org/downloads.html")
            raise FileNotFoundError("wkhtmltopdf not found")

        # pdfkit配置
        self.config = pdfkit.configuration(wkhtmltopdf=self.wkhtmltopdf_path)
        self.options = {
            'encoding': 'UTF-8',
            'page-size': 'A4',
            'margin-top': '20mm',
            'margin-right': '20mm',
            'margin-bottom': '20mm',
            'margin-left': '20mm',
            'enable-local-file-access': None,
            'custom-header': [
                ('Accept-Encoding', 'gzip')
            ]
        }

    def convert_md_to_html(self, md_file):
        """
        将Markdown转换为HTML
        :param md_file: Markdown文件路径
        :return: HTML内容
        """
        try:
            with open(md_file, 'r', encoding='utf-8') as f:
                md_content = f.read()

            # 配置Markdown扩展
            extensions = [
                'markdown.extensions.tables',
                'markdown.extensions.fenced_code',
                'markdown.extensions.codehilite',
                'markdown.extensions.toc',
                'markdown.extensions.meta',
                'markdown.extensions.nl2br'
            ]

            # 转换Markdown为HTML
            html_content = markdown.markdown(
                md_content,
                extensions=extensions,
                output_format='html5'
            )

            # 添加CSS样式
            html_template = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body {{
                        font-family: "Microsoft YaHei", Arial, sans-serif;
                        line-height: 1.6;
                        padding: 20px;
                        max-width: 21cm;
                        margin: 0 auto;
                    }}
                    img {{
                        max-width: 100%;
                        height: auto;
                    }}
                    pre {{
                        background-color: #f5f5f5;
                        padding: 10px;
                        border-radius: 5px;
                        overflow-x: auto;
                    }}
                    code {{
                        font-family: Consolas, Monaco, 'Courier New', monospace;
                    }}
                    table {{
                        border-collapse: collapse;
                        width: 100%;
                        margin: 15px 0;
                    }}
                    th, td {{
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                    }}
                    th {{
                        background-color: #f5f5f5;
                    }}
                </style>
            </head>
            <body>
                {html_content}
            </body>
            </html>
            """

            return html_template
        except Exception as e:
            self.logger.error(f"转换{md_file}时出错: {str(e)}")
            raise

    def convert_file(self, md_file):
        """
        转换单个文件
        :param md_file: Markdown文件路径
        """
        try:
            # 生成输出文件名
            pdf_file = self.output_dir / f"{md_file.stem}.pdf"
            temp_html_file = self.temp_dir / f"{md_file.stem}.html"

            self.logger.info(f"正在转换: {md_file.name}")

            # 转换为HTML
            html_content = self.convert_md_to_html(md_file)

            # 保存临时HTML文件
            with open(temp_html_file, 'w', encoding='utf-8') as f:
                f.write(html_content)

            # 使用pdfkit转换为PDF
            pdfkit.from_file(
                str(temp_html_file),
                str(pdf_file),
                options=self.options,
                configuration=self.config
            )

            self.logger.info(f"转换完成: {pdf_file}")

            # 验证PDF文件
            if not pdf_file.exists() or pdf_file.stat().st_size == 0:
                raise Exception("PDF文件生成失败或为空")

            return True
        except Exception as e:
            self.logger.error(f"转换失败: {str(e)}")
            return False
        finally:
            # 删除临时HTML文件
            if temp_html_file.exists():
                temp_html_file.unlink()

    def convert_all(self):
        """
        转换目录下所有Markdown文件
        """
        try:
            # 检查输入目录是否存在
            if not self.input_dir.exists():
                print(f"错误: 输入目录 '{self.input_dir}' 不存在!")
                return

            md_files = list(self.input_dir.glob('*.md'))
            if not md_files:
                print(f"在 {self.input_dir} 中没有找到Markdown文件")
                return

            print(f"找到 {len(md_files)} 个Markdown文件")
            success_count = 0
            fail_count = 0

            for md_file in md_files:
                if self.convert_file(md_file):
                    success_count += 1
                else:
                    fail_count += 1

            print(f"\n转换完成:")
            print(f"成功: {success_count} 个")
            print(f"失败: {fail_count} 个")
            print(f"\nPDF文件已保存到: {self.output_dir}")

        finally:
            # 清理临时目录
            if self.temp_dir.exists():
                shutil.rmtree(self.temp_dir)


if __name__ == '__main__':
    try:
        converter = MarkdownToPdfConverter()
        converter.convert_all()
        input("\n按回车键退出...")
    except Exception as e:
        print(f"发生错误: {str(e)}")
        input("\n按回车键退出...")
