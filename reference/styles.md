# Style Customization Guide

Custom styles can be configured interactively or by modifying the `DEFAULT_STYLES` dictionary in `md_to_word.py`.

## Style structure

```python
{
    "font": {
        "family": "字体名称",
        "size": 磅值
    },
    "paragraph": {
        "align": "对齐方式",        # left/center/right/justify
        "line_spacing": 行距磅值,
        "space_before": 段前磅值,
        "space_after": 段后磅值,
        "first_line_indent": 首行缩进字符数
    }
}
```

## Common Chinese font families

| Use case | Font family |
|----------|-------------|
| Headings | 黑体, 方正小标宋, 宋体 |
| Body | 仿宋_GB2312, 仿宋, 楷体, 宋体 |
| Emphasis | 黑体, 方正黑体 |

## Common alignment values

| Value | Description |
|-------|-------------|
| left | 左对齐 |
| center | 居中对齐 |
| right | 右对齐 |
| justify | 两端对齐 |
