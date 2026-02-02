# Technical Details

## COM Interface

### Microsoft Word

```python
import win32com.client

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

doc = word.Documents.Open(html_path)
doc.SaveAs2(docx_path, FileFormat=12)  # wdFormatXMLDocument (.docx)
doc.Close(SaveChanges=False)
word.Quit()
```

### WPS Office

```python
import win32com.client

wps = win32com.client.Dispatch("WPS.Application")
# Same API as Word
doc.SaveAs2(docx_path, FileFormat=12)
```

## FileFormat values

| Value | Format | Extension |
|-------|--------|-----------|
| 0 | Word 97-2003 | .doc |
| 12 | Word 2007+ (XML) | .docx |
| 16 | PDF | .pdf |

## Error handling

| Error | Cause | Solution |
|-------|-------|----------|
| `(-2147221164, '调用方未调用'` | Application not found | Install Word or WPS |
| `PermissionError` | File in use | Close the file and retry |
| `FileNotFoundError` | Invalid path | Check file path |

## Page setup in points

```python
# Margins (2cm = 56.7 points)
doc.PageSetup.TopMargin = 56.7
doc.PageSetup.BottomMargin = 56.7
doc.PageSetup.LeftMargin = 56.7
doc.PageSetup.RightMargin = 56.7
```
