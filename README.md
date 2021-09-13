# msoffice2pdf 0.0.8

This package aims to convert Microsoft Office file types to PDF. This lib uses the comtypes package which makes it easy to access and implement custom, dispatch-based COM interfaces or LibreOffice software.

For use in `Windows` environment and `soft=0` Microsoft Office must be installed.

For use in `Windows` environment and `soft=1` you need the latest version of LibreOffice(soffice) installed.

For `Ubuntu(linux)` environment it is only possible to use `soft=1`, that is, LibreOffice(soffice).

Supported files: [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".xml"]


### Installation

Step 1:

`pip3 install msoffice2pdf`

Step 2

### (Windons Only):

For Microsoft Office use:

*   Delete all cache files from the folder below in case there is any error with Microsoft Office conversion:
 `C:\Users\<User>\AppData\Local\Programs\Python\Python39\Lib\site-packages\comtypes\gen`

*   For Windows Server

> **Step 1:**
>
> Start > Run > dcomcnfg.exe
>
> **Step 2:**
>
>  Select: Computers -> My Computer -> Config DCOM -> Microsoft Word 97-2003 Documents -> Properties
>  Tab general select level authentication to None
>  Tab security select customize and add All
>  Tab identify select this user and add Admin user and password
>  
> **Step 3:**
>  
> Select: Computers -> My Computer -> Config DCOM -> Microsoft Excel Application -> Properties
> Tab general select level authentication to None
> Tab security select customize and add All
> Tab identify select this user and add Admin user and password
>  
> **Step 4:**
>  
> Select: Computers -> My Computer -> Config DCOM -> Microsoft PowerPoint Application -> Properties
> Tab general select level authentication to None
> Tab security select customize and add All
> Tab identify select this user and add Admin user and password


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

output = convert(source="C:/Users/<User>/Downloads/file.docx", output_dir="C:/Users/<User>/Downloads", soft=0)

print(output)
```

