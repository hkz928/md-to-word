# MD to Word Converter

[English](#english) | [简体中文](#简体中文)

---

## English

A Python tool that converts Markdown files to Word documents with Chinese official document formatting standards (中文公文格式). Supports both Microsoft Word and WPS Office.

## Features

- **Interactive Mode**: Preview default styles and choose whether to use them or customize
- **Silent Mode**: Convert directly with default styles without prompts
- **Dual Application Support**: Automatically detects and uses Microsoft Word or WPS Office
- **Chinese Official Document Format**: Pre-configured styles compliant with GB/T 9704-2012 standard
- **HTML Intermediate**: Generates HTML files with embedded CSS for quality control

## Requirements

- Python 3.8+
- Microsoft Word (2010+) or WPS Office
- Windows OS

## Installation

1. Clone this repository:
```bash
git clone https://github.com/hkz928/md-to-word.git
cd md-to-word
```

2. Install dependencies:
```bash
pip install pywin32
```

## Usage

### Interactive Mode (Recommended)

```bash
python scripts/md_to_word.py document.md
```

This will:
1. Detect installed Word/WPS application
2. Show default style preview
3. Ask whether to use default styles
4. Convert silently in background
5. Output the file path

### Silent Mode

```bash
python scripts/md_to_word.py document.md --no-prompt
```

Converts directly using default styles without any prompts.

### Convert and Open

```bash
python scripts/md_to_word.py document.md --no-prompt --open
```

Converts and opens the resulting document automatically.

### Check Available Application

```bash
python scripts/md_to_word.py --check-app
```

Outputs: `已安装: Word` or `已安装: WPS`

## Default Styles (Chinese Official Document Format)

| Element | Font | Size | Alignment | Line spacing | Notes |
|---------|------|------|-----------|--------------|-------|
| Heading 1 (#) | SimHei (黑体) | 16pt | Center | 30pt | 0.5 line after |
| Heading 2 (##) | SimHei (黑体) | 16pt | Left | 30pt | 0.5 line after |
| Heading 3 (###) | SimHei (黑体) | 16pt | Left | 28pt | - |
| Body | FangSong_GB2312 (仿宋_GB2312) | 16pt | Justify | 28pt | 2-char indent |
| Page margins | - | - | - | - | 2cm all sides |

## Supported Markdown Syntax

- `# Heading` - Level 1 heading
- `## Heading` - Level 2 heading
- `### Heading` - Level 3 heading
- `**bold**` - Bold text
- `*italic*` - Italic text
- `~~strike~~` - Strikethrough
- `- item` - Unordered list
- `1. item` - Ordered list
- `[link](url)` - Hyperlink
- `` `code` `` - Inline code

## Project Structure

```
md-to-word/
├── README.md             # This file (bilingual)
├── LICENSE               # MIT License
├── .gitignore            # Git ignore file
├── SKILL.md              # Claude Skill definition
├── scripts/
│   └── md_to_word.py     # Main conversion script
├── reference/
│   ├── styles.md         # Style customization guide
│   ├── sizes.md          # Character size reference
│   └── technical.md      # Technical details
└── examples/
    └── 示例文档.md       # Example document
```

## Example

### Input Markdown

```markdown
# 关于启用办公自动化系统的通知

## 一、背景说明

为进一步提高办公效率，推进无纸化办公进程，经研究决定，自即日起启用新的办公自动化系统。

## 二、使用要求

请各单位务必于本月底前完成系统上线工作：

- 组织相关人员参加系统培训
- 完成历史数据迁移
- 建立配套管理制度

特此通知。
```

### Output

- `document.html` - Intermediate HTML file with embedded CSS
- `document.docx` - Final Word document with proper formatting

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Based on Chinese official document standard GB/T 9704-2012
- Supports both Microsoft Word and WPS Office via COM automation

---

## 简体中文

一个将 Markdown 文件转换为符合中文公文格式标准 (GB/T 9704-2012) 的 Word 文档的 Python 工具。支持 Microsoft Word 和 WPS Office。

## 功能特点

- **交互模式**：预览默认样式并选择使用或自定义
- **静默模式**：直接使用默认样式转换，无需确认
- **双应用支持**：自动检测并使用 Microsoft Word 或 WPS Office
- **中文公文格式**：预配置符合 GB/T 9704-2012 标准的样式
- **HTML 中间文件**：生成嵌入 CSS 的 HTML 文件便于质量控制

## 系统要求

- Python 3.8+
- Microsoft Word (2010+) 或 WPS Office
- Windows 操作系统

## 安装

1. 克隆此仓库：
```bash
git clone https://github.com/hkz928/md-to-word.git
cd md-to-word
```

2. 安装依赖：
```bash
pip install pywin32
```

## 使用方法

### 交互模式（推荐）

```bash
python scripts/md_to_word.py 文档.md
```

流程：
1. 检测已安装的 Word/WPS 应用
2. 显示默认样式预览
3. 询问是否使用默认样式
4. 后台静默转换
5. 输出文件路径

### 静默模式

```bash
python scripts/md_to_word.py 文档.md --no-prompt
```

直接使用默认样式转换，无需任何确认。

### 转换后打开

```bash
python scripts/md_to_word.py 文档.md --no-prompt --open
```

转换完成后自动打开生成的文档。

### 检测可用应用

```bash
python scripts/md_to_word.py --check-app
```

输出：`已安装: Word` 或 `已安装: WPS`

## 默认样式（中文公文格式）

| 元素 | 字体 | 字号 | 对齐 | 行距 | 其他 |
|------|------|------|------|------|------|
| 一级标题 (#) | 黑体 | 三号(16pt) | 居中 | 30磅 | 段后0.5行 |
| 二级标题 (##) | 黑体 | 三号(16pt) | 左对齐 | 30磅 | 段后0.5行 |
| 三级标题 (###) | 黑体 | 三号(16pt) | 左对齐 | 28磅 | - |
| 正文 | 仿宋_GB2312 | 三号(16pt) | 两端对齐 | 28磅 | 首行缩进2字符 |
| 页边距 | - | - | - | - | 上下左右各2cm |

## 支持的 Markdown 语法

- `# 标题` - 一级标题
- `## 标题` - 二级标题
- `### 标题` - 三级标题
- `**粗体**` - 粗体文本
- `*斜体*` - 斜体文本
- `~~删除线~~` - 删除线
- `- 项目` - 无序列表
- `1. 项目` - 有序列表
- `[链接](url)` - 超链接
- `` `代码` `` - 行内代码

## 项目结构

```
md-to-word/
├── README.md             # 本文档（中英双语）
├── LICENSE               # MIT 许可证
├── .gitignore            # Git 忽略文件
├── SKILL.md              # Claude Skill 定义
├── scripts/
│   └── md_to_word.py     # 主转换脚本
├── reference/
│   ├── styles.md         # 样式自定义指南
│   ├── sizes.md          # 字号参考
│   └── technical.md      # 技术细节
└── examples/
    └── 示例文档.md       # 示例文档
```

## 示例

### 输入 Markdown

```markdown
# 关于启用办公自动化系统的通知

## 一、背景说明

为进一步提高办公效率，推进无纸化办公进程，经研究决定，自即日起启用新的办公自动化系统。

## 二、使用要求

请各单位务必于本月底前完成系统上线工作：

- 组织相关人员参加系统培训
- 完成历史数据迁移
- 建立配套管理制度

特此通知。
```

### 输出

- `文档.html` - 嵌入 CSS 的 HTML 中间文件
- `文档.docx` - 格式正确的最终 Word 文档

## 贡献

欢迎贡献！请随时提交 Pull Request。

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件。

## 致谢

- 基于中文公文标准 GB/T 9704-2012
- 通过 COM 自动化支持 Microsoft Word 和 WPS Office
