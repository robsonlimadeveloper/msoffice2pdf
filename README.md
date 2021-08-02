# msoffice2pdf

This package aims to convert Microsoft Office file types to PDF. This lib uses the comtypes package which makes it easy to access and implement custom, dispatch-based COM interfaces or LibreOffice software.

For use in `Windows` environment and `soft="msoffice"` Microsoft Office must be installed.

For use in `Windows` environment and `soft="libreoffice"` you need the latest version of LibreOffice(soffice) installed.

For `Ubuntu(linux)` environment it is only possible to use `soft="libreoffice"`, that is, LibreOffice(soffice).

Supported files: [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".xml"]


### Installation

Step 1:

`pip3 install msoffice2pdf`

Step 2

### (Windons Only):

For Microsoft Office use:

*   Delete all cache files from the folder below in case there is any error with Microsoft Office conversion:
 `C:\Users\<User>\AppData\Local\Programs\Python\Python39\Lib\site-packages\comtypes\gen`

*   Efetuar a configuração abaixo(Windows Server):

> 1. Start -> dcomcnfg.exe
> 1. Computers -> My Computer
> 1. DCOM Config 
> 1. Select the Microsoft Word 97-2003 Documents -> Properties
> 1. Tab Identity, change from Launching User to Interactive User

For LibreOffice use:

Install LibreOffice last version:

https://www.libreoffice.org/download/download/

###  (Ubuntu Only):

Install LibreOffice:

`sudo add-apt-repository -y ppa:libreoffice/ppa`

`sudo apt-get update`

`sudo apt-get install libreoffice libreoffice-style-breeze`

### Example:

```python
from msoffice2pdf import convert

output = convert(source="C:/Users/<User>/Downloads/file.docx", output_dir="C:/Users/<User>/Downloads", soft="msoffice")

print(output)
```

