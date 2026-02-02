---
name: md-to-word-converter
description: Convert Markdown files to Word documents with Chinese official document formatting. Supports Microsoft Word and WPS Office. Use when converting .md files to .docx with Chinese typography requirements (黑体, 仿宋_GB2312, specific line spacing, indentation) or when user mentions Chinese official documents,公文格式, or document conversion.
---

# Document Conversion (Markdown to Word)

Converts Markdown files to Word documents with Chinese official document formatting standards.

## Quick start

```bash
# Interactive mode (prompts for style confirmation)
python scripts/md_to_word.py "examples/示例文档.md"

# Silent mode (uses default Chinese official document style)
python scripts/md_to_word.py "examples/示例文档.md" --no-prompt

# Convert and open
python scripts/md_to_word.py "examples/示例文档.md" --no-prompt --open

# Check which application is available
python scripts/md_to_word.py --check-app
```

## Default Chinese official document style

| Element | Font | Size | Alignment | Line spacing | Notes |
|---------|------|------|-----------|--------------|-------|
| h1 (#) | SimHei | 16pt | center | 30pt | 0.5 line after |
| h2 (##) | SimHei | 16pt | left | 30pt | 0.5 line after |
| h3 (###) | SimHei | 16pt | left | 28pt | - |
| Body | FangSong_GB2312 | 16pt | justify | 28pt | 2-char indent |
| Page margins | - | - | - | - | 2cm all sides |

## Application detection

Automatically detects and uses available word processor:

1. Microsoft Word (Word.Application) - preferred
2. WPS Office (WPS.Application) - fallback when Word unavailable

## Conversion workflow

```
.md file -> Parse Markdown -> Generate HTML -> Save .html
-> Open in Word/WPS -> Save as .docx -> Output file path
```

## Supported Markdown syntax

- # Heading (Level 1)
- ## Heading (Level 2)
- ### Heading (Level 3)
- **bold** (Bold text)
- *italic* (Italic text)
- ~~strike~~ (Strikethrough)
- - item (Unordered list)
- 1. item (Ordered list)
- [link](url) (Hyperlink)
- `code` (Inline code)

## Requirements

- Python 3.x
- pip install pywin32
- Microsoft Word or WPS Office installed

## Reference

- Style customization: See reference/styles.md
- Character sizes: See reference/sizes.md
- Technical details: See reference/technical.md
