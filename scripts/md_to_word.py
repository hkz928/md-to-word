"""
Markdown 转 Word (中文公文格式)
支持交互式样式选择和静默转换
支持 Microsoft Word 和 WPS Office
"""

import re
import sys
import os
import json
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import Optional, Tuple

try:
    import win32com.client
    from win32com.client import constants
except ImportError:
    print("错误: 需要安装 pywin32")
    print("请运行: pip install pywin32")
    sys.exit(1)


# ==================== 样式配置 ====================

@dataclass
class FontConfig:
    """字体配置"""
    family: str  # 字体名称
    size: float  # 字号(磅)


@dataclass
class ParagraphConfig:
    """段落配置"""
    line_spacing: float  # 行距(磅)
    space_before: float  # 段前(磅)
    space_after: float   # 段后(磅)
    align: str           # 对齐方式: left/center/right/justify
    first_line_indent: float  # 首行缩进(字符)


@dataclass
class ElementStyle:
    """元素样式"""
    name: str
    font: FontConfig
    paragraph: ParagraphConfig


# 默认样式配置
DEFAULT_STYLES = {
    "h1": ElementStyle(
        name="一级标题",
        font=FontConfig(family="黑体", size=16),
        paragraph=ParagraphConfig(
            line_spacing=30,
            space_before=0,
            space_after=12,  # 约0.5行
            align="center",
            first_line_indent=0
        )
    ),
    "h2": ElementStyle(
        name="二级标题",
        font=FontConfig(family="黑体", size=16),
        paragraph=ParagraphConfig(
            line_spacing=30,
            space_before=12,
            space_after=12,
            align="left",
            first_line_indent=0
        )
    ),
    "h3": ElementStyle(
        name="三级标题",
        font=FontConfig(family="黑体", size=16),
        paragraph=ParagraphConfig(
            line_spacing=28,
            space_before=6,
            space_after=6,
            align="left",
            first_line_indent=0
        )
    ),
    "p": ElementStyle(
        name="正文",
        font=FontConfig(family="仿宋_GB2312", size=16),
        paragraph=ParagraphConfig(
            line_spacing=28,
            space_before=0,
            space_after=0,
            align="justify",
            first_line_indent=2  # 首行缩进2字符
        )
    ),
    "li": ElementStyle(
        name="列表",
        font=FontConfig(family="仿宋_GB2312", size=16),
        paragraph=ParagraphConfig(
            line_spacing=28,
            space_before=0,
            space_after=0,
            align="justify",
            first_line_indent=0
        )
    ),
}


# ==================== 应用程序检测 ====================

class AppDetector:
    """检测系统安装的文字处理软件"""

    # Word 的可能 ProgID
    WORD_PROG_IDS = [
        "Word.Application",      # 英文版 Word
        "Word.Application.16",   # Word 2016/2019/2021
        "Word.Application.15",   # Word 2013
        "Word.Application.14",   # Word 2010
    ]

    # WPS 的可能 ProgID
    WPS_PROG_IDS = [
        "WPS.Application",       # WPS 文字
        "KWps.Application",      # WPS 文字（旧版）
    ]

    @classmethod
    def detect_app(cls) -> Tuple[Optional[str], str]:
        """
        检测可用的文字处理软件

        返回: (prog_id, app_name)
            - prog_id: COM ProgID，如果没有找到则为 None
            - app_name: 应用名称 ("Word", "WPS", 或 "None")
        """
        # 优先检测 Word
        for prog_id in cls.WORD_PROG_IDS:
            try:
                win32com.client.Dispatch(prog_id)
                return prog_id, "Word"
            except Exception:
                continue

        # 如果没有 Word，检测 WPS
        for prog_id in cls.WPS_PROG_IDS:
            try:
                win32com.client.Dispatch(prog_id)
                return prog_id, "WPS"
            except Exception:
                continue

        return None, "None"


# ==================== Markdown 解析 ====================

class MarkdownParser:
    """Markdown 解析器"""

    def __init__(self, content: str):
        self.content = content
        self.lines = content.split('\n')

    def parse(self) -> list:
        """解析为元素列表"""
        elements = []
        i = 0

        while i < len(self.lines):
            line = self.lines[i].rstrip()

            # 空行
            if not line:
                elements.append({"type": "empty"})
                i += 1
                continue

            # 一级标题
            if line.startswith('# '):
                elements.append({"type": "h1", "text": line[2:].strip()})
                i += 1
                continue

            # 二级标题
            if line.startswith('## '):
                elements.append({"type": "h2", "text": line[3:].strip()})
                i += 1
                continue

            # 三级标题
            if line.startswith('### '):
                elements.append({"type": "h3", "text": line[4:].strip()})
                i += 1
                continue

            # 无序列表
            if re.match(r'^[\*\-]\s+', line):
                elements.append({
                    "type": "li",
                    "text": re.sub(r'^[\*\-]\s+', '', line),
                    "list_type": "ul"
                })
                i += 1
                continue

            # 有序列表
            if re.match(r'^\d+\.\s+', line):
                elements.append({
                    "type": "li",
                    "text": re.sub(r'^\d+\.\s+', '', line),
                    "list_type": "ol"
                })
                i += 1
                continue

            # 正文段落
            elements.append({"type": "p", "text": line})
            i += 1

        return elements

    @staticmethod
    def process_inline(text: str) -> str:
        """处理行内格式"""
        text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
        text = re.sub(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', r'<em>\1</em>', text)
        text = re.sub(r'~~(.+?)~~', r'<s>\1</s>', text)
        text = re.sub(r'\[([^\]]+)\]\(([^\)]+)\)', r'<a href="\2">\1</a>', text)
        text = re.sub(r'`([^`]+)`', r'<code>\1</code>', text)
        return text


# ==================== HTML 生成 ====================

class HTMLGenerator:
    """HTML 生成器"""

    def __init__(self, styles: dict):
        self.styles = styles

    def generate_css(self) -> str:
        """生成 CSS 样式"""
        css_lines = []

        for elem_type, style in self.styles.items():
            font = style.font
            para = style.paragraph

            if elem_type == "li":
                css_lines.append("""
ul, ol {
    font-family: "%s", "仿宋", "FangSong", serif;
    font-size: %gpt;
    text-align: %s;
    line-height: %gpt;
    margin-top: 0;
    margin-bottom: 0;
    padding-left: 2em;
}
""" % (font.family, font.size, para.align, para.line_spacing))
            else:
                selector = elem_type
                css_lines.append("""
%s {
    font-family: "%s", serif;
    font-size: %gpt;
    font-weight: %s;
    text-align: %s;
    line-height: %gpt;
    margin-top: %gpt;
    margin-bottom: %gpt;
    %s
}
""" % (
                    selector,
                    font.family,
                    font.size,
                    "bold" if elem_type.startswith("h") else "normal",
                    para.align,
                    para.line_spacing,
                    para.space_before,
                    para.space_after,
                    f"text-indent: {para.first_line_indent}em;" if para.first_line_indent > 0 else ""
                ))

        return "\n".join(css_lines)

    def generate(self, elements: list) -> str:
        """生成 HTML"""
        css = self.generate_css()

        html_body = []
        for elem in elements:
            elem_type = elem["type"]

            if elem_type == "empty":
                html_body.append('<p>&nbsp;</p>')
            elif elem_type == "li":
                text = MarkdownParser.process_inline(elem["text"])
                html_body.append(f'<{elem["list_type"]}><li>{text}</li></{elem["list_type"]}>')
            else:
                text = MarkdownParser.process_inline(elem.get("text", ""))
                html_body.append(f'<{elem_type}>{text}</{elem_type}>')

        return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {{
            margin: 2cm 2cm 2cm 2cm;
        }}
{css}
    </style>
</head>
<body>
{''.join(html_body)}
</body>
</html>"""


# ==================== Word/WPS 转换 ====================

class DocumentConverter:
    """
    文档转换器
    支持 Microsoft Word 和 WPS Office
    """

    def __init__(self, silent: bool = True):
        self.silent = silent
        self.app = None
        self.app_name = None

    def _log(self, message: str):
        """输出日志"""
        if not self.silent:
            print(message)

    def _get_app(self) -> Tuple[Optional[object], Optional[str]]:
        """
        获取可用的文字处理应用
        优先使用 Word，如果没有则使用 WPS

        返回: (app_object, app_name)
        """
        prog_id, app_name = AppDetector.detect_app()

        if prog_id is None:
            print("错误: 未找到 Microsoft Word 或 WPS Office")
            print("请安装以下任一软件：")
            print("  - Microsoft Office Word")
            print("  - WPS Office")
            return None, None

        self._log(f"使用 {app_name} 进行转换")
        app = win32com.client.Dispatch(prog_id)
        return app, app_name

    def html_to_docx(self, html_path: str, docx_path: str, open_file: bool = False) -> bool:
        """
        将 HTML 转换为 DOCX
        自动检测并使用 Word 或 WPS
        """
        html_path_abs = os.path.abspath(html_path)
        docx_path_abs = os.path.abspath(docx_path)

        try:
            # 获取应用
            self.app, self.app_name = self._get_app()
            if self.app is None:
                return False

            self.app.Visible = False

            self._log(f"正在打开: {html_path_abs}")
            doc = self.app.Documents.Open(html_path_abs)

            # 设置页边距 (2cm ≈ 56.7 磅)
            margin = 56.7
            doc.PageSetup.TopMargin = margin
            doc.PageSetup.BottomMargin = margin
            doc.PageSetup.LeftMargin = margin
            doc.PageSetup.RightMargin = margin

            self._log(f"正在保存: {docx_path_abs}")
            # wdFormatXMLDocument = 12
            # WPS 也支持相同的格式代码
            doc.SaveAs2(docx_path_abs, FileFormat=12)

            doc.Close(SaveChanges=False)

            if open_file:
                self.app.Visible = True
            else:
                self.app.Quit()

            return True

        except Exception as e:
            print(f"转换失败: {e}")
            if self.app:
                try:
                    self.app.Quit()
                except:
                    pass
            return False


# ==================== 交互式输入 ====================

def ask_yes_no(prompt: str, default: bool = True) -> bool:
    """询问是/否"""
    suffix = " [Y/n]" if default else " [y/N]"
    while True:
        try:
            response = input(prompt + suffix + ": ").strip().lower()
            if not response:
                return default
            if response in ['y', 'yes', '是', 'y']:
                return True
            if response in ['n', 'no', '否', 'n']:
                return False
            print("请输入 y/yes/是 或 n/no/否")
        except EOFError:
            # 非交互模式，返回默认值
            return default


def ask_style_config() -> dict:
    """询问样式配置"""
    print("\n=== 样式配置 ===")
    print("请输入各元素的样式参数（直接回车使用默认值）\n")

    styles = {}

    # 一级标题
    print("【一级标题】")
    try:
        h1_font = input("  字体 (默认: 黑体): ").strip() or "黑体"
        h1_size = input("  字号/磅 (默认: 16): ").strip() or "16"
        h1_align = input("  对齐 (默认: center): ").strip() or "center"
        h1_ls = input("  行距/磅 (默认: 30): ").strip() or "30"
        h1_sa = input("  段后/磅 (默认: 12): ").strip() or "12"

        styles["h1"] = ElementStyle(
            name="一级标题",
            font=FontConfig(family=h1_font, size=float(h1_size)),
            paragraph=ParagraphConfig(
                line_spacing=float(h1_ls),
                space_before=0,
                space_after=float(h1_sa),
                align=h1_align,
                first_line_indent=0
            )
        )

        # 二级标题
        print("\n【二级标题】")
        h2_font = input("  字体 (默认: 黑体): ").strip() or "黑体"
        h2_size = input("  字号/磅 (默认: 16): ").strip() or "16"
        h2_align = input("  对齐 (默认: left): ").strip() or "left"
        h2_ls = input("  行距/磅 (默认: 30): ").strip() or "30"
        h2_sa = input("  段后/磅 (默认: 12): ").strip() or "12"

        styles["h2"] = ElementStyle(
            name="二级标题",
            font=FontConfig(family=h2_font, size=float(h2_size)),
            paragraph=ParagraphConfig(
                line_spacing=float(h2_ls),
                space_before=12,
                space_after=float(h2_sa),
                align=h2_align,
                first_line_indent=0
            )
        )

        # 正文
        print("\n【正文】")
        p_font = input("  字体 (默认: 仿宋_GB2312): ").strip() or "仿宋_GB2312"
        p_size = input("  字号/磅 (默认: 16): ").strip() or "16"
        p_align = input("  对齐 (默认: justify): ").strip() or "justify"
        p_ls = input("  行距/磅 (默认: 28): ").strip() or "28"
        p_indent = input("  首行缩进/字符 (默认: 2): ").strip() or "2"

        styles["p"] = ElementStyle(
            name="正文",
            font=FontConfig(family=p_font, size=float(p_size)),
            paragraph=ParagraphConfig(
                line_spacing=float(p_ls),
                space_before=0,
                space_after=0,
                align=p_align,
                first_line_indent=float(p_indent)
            )
        )

        # 继承默认的三级标题和列表样式
        styles["h3"] = DEFAULT_STYLES["h3"]
        styles["li"] = DEFAULT_STYLES["li"]

    except EOFError:
        # 非交互模式，使用默认样式
        return DEFAULT_STYLES

    return styles


def print_default_styles():
    """打印默认样式"""
    print("\n=== 默认样式预览 ===")
    for key, style in DEFAULT_STYLES.items():
        print(f"\n【{style.name}】")
        print(f"  字体: {style.font.family}, {style.font.size}磅")
        print(f"  行距: {style.paragraph.line_spacing}磅")
        if style.paragraph.space_after > 0:
            print(f"  段后: {style.paragraph.space_after}磅")
        if style.paragraph.first_line_indent > 0:
            print(f"  首行缩进: {style.paragraph.first_line_indent}字符")
        print(f"  对齐: {style.paragraph.align}")


# ==================== 主流程 ====================

def convert_md_to_word(md_path: str, styles: dict, silent: bool = True) -> Optional[str]:
    """转换 MD 到 Word"""
    md_path_obj = Path(md_path)

    if not md_path_obj.exists():
        print(f"错误: 文件不存在: {md_path}")
        return None

    if not silent:
        print(f"\n正在读取: {md_path}")

    # 读取 MD
    with open(md_path, 'r', encoding='utf-8') as f:
        md_content = f.read()

    # 解析
    parser = MarkdownParser(md_content)
    elements = parser.parse()

    # 生成 HTML
    generator = HTMLGenerator(styles)
    html_content = generator.generate(elements)

    # 保存 HTML
    html_path = str(md_path_obj.parent / f"{md_path_obj.stem}.html")
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    # 转换为 Word
    docx_path = str(md_path_obj.parent / f"{md_path_obj.stem}.docx")
    converter = DocumentConverter(silent=silent)

    success = converter.html_to_docx(html_path, docx_path, open_file=False)

    if success:
        return os.path.abspath(docx_path)
    return None


def interactive_mode(md_path: str):
    """交互模式"""
    print(f"\n{'='*40}")
    print(f"  Markdown 转 Word (中文公文格式)")
    print(f"{'='*40}")

    # 检测可用应用
    _, app_name = AppDetector.detect_app()
    print(f"  检测到: {app_name}")
    print(f"{'='*40}")

    print(f"\n待转换文件: {md_path}\n")

    # 显示默认样式
    print_default_styles()

    # 询问是否使用默认样式
    use_default = ask_yes_no("\n是否使用默认样式？", default=True)

    if use_default:
        styles = DEFAULT_STYLES
        print("\n使用默认样式，正在转换...")
    else:
        styles = ask_style_config()
        print("\n正在转换...")

    # 执行转换
    result_path = convert_md_to_word(md_path, styles, silent=True)

    print("\n" + "="*40)
    if result_path:
        print(f"转换成功！")
        print(f"\n输出文件: {result_path}")
        print(f"HTML文件: {result_path.replace('.docx', '.html')}")

        # 询问是否打开
        if ask_yes_no("\n是否打开文档？", default=False):
            os.startfile(result_path)
    else:
        print("转换失败")

    print("="*40)


def main():
    """命令行入口"""
    import argparse

    parser = argparse.ArgumentParser(
        description='Markdown 转 Word (中文公文格式)',
        epilog='支持 Microsoft Word 和 WPS Office'
    )
    parser.add_argument('md_file', nargs='?', help='Markdown 文件路径')
    parser.add_argument('-o', '--output', help='输出 Word 文件路径')
    parser.add_argument('--no-prompt', action='store_true', help='跳过交互，使用默认样式')
    parser.add_argument('--open', action='store_true', help='转换后打开文件')
    parser.add_argument('--check-app', action='store_true', help='仅检测已安装的应用')

    args = parser.parse_args()

    # 检测应用模式
    if args.check_app:
        _, app_name = AppDetector.detect_app()
        print(f"已安装: {app_name}")
        return

    # 检查是否提供了文件路径
    if not args.md_file:
        parser.print_help()
        print("\n错误: 请指定 Markdown 文件路径")
        sys.exit(1)

    md_path = args.md_file

    if args.no_prompt:
        # 静默模式，使用默认样式
        result_path = convert_md_to_word(md_path, DEFAULT_STYLES, silent=True)
        if result_path:
            print(f"转换成功: {result_path}")
            if args.open:
                os.startfile(result_path)
        else:
            print("转换失败")
            sys.exit(1)
    else:
        # 交互模式
        interactive_mode(md_path)


if __name__ == '__main__':
    main()
