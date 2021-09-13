# msoffice2pdf

This package aims to convert Microsoft Office file types to PDF. This lib uses the comtypes package which makes it easy to access and implement custom, dispatch-based COM interfaces or LibreOffice software.

For use in `Windows` environment and `soft=0` Microsoft Office must be installed.

For use in `Windows` environment and `soft=1` you need the latest version of LibreOffice(soffice) installed.

For `Ubuntu(linux)` environment it is only possible to use `soft=1`, that is, LibreOffice(soffice).

Supported files: [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".xml"]


### Installation

Step 1:

`pip3 install msoffice2pdf`

Step 2(Windons Only):

Delete all cache files from the folder below in case there is any error with Microsoft Office conversion: `C:\Users\<User>\AppData\Local\Programs\Python\Python39\Lib\site-packages\comtypes\gen`

### Example:

```python
from msoffice2pdf import convert

output = convert(source="C:/Users/<User>/Downloads/file.docx", output_dir="C:/Users/<User>/Downloads", soft=0)

print(output)
```

