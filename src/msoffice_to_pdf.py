import os
import subprocess
import re
import datetime
import platform

if platform.system() == "Windows":
    from comtypes import client

def remove_files(temp_files_attach):
    """Remove temporary files"""
    for file_temp in temp_files_attach:
        if os.path.isfile(file_temp):
            os.remove(file_temp)

def convert_to_pdf_libreoffice(source: str, timeout=None)-> str:
    """Convert MS Office files using LibreOffice"""

    temp_filename = os.path.dirname(source)+datetime.now().strftime("%Y%m%d%H%M%S%f")+source
    subprocess.run(['cp', os.path.dirname(source)+source, temp_filename],\
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,\
        timeout=timeout, check=True)
    try:
        process = subprocess.run(['soffice', '--headless', '--convert-to',\
            'pdf', '--outdir', os.path.dirname(source), temp_filename],\
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,\
                    timeout=timeout, check=True)
        filename = re.search('-> (.*?) using filter', process.stdout.decode("latin-1"))

    except Exception as exception:
        return {"output": None, "error": exception.__dict__["details"]}

    remove_files([temp_filename])

    return filename.group(1)

def convert_doc_to_pdf_msoffice(source: str, output: str)-> str:
    '''This fuction convert *.doc/*.docx files to pdf'''

    ws_pdf_format: int = 17
    app = client.CreateObject("Word.Application")
    try:
        doc = app.Documents.Open(source)
        doc.ExportAsFixedFormat(output, ws_pdf_format, Item=7, CreateBookmarks=0)
        app.Quit()

    except Exception as exception:
        app.Quit()
        return {"output": None, "error": exception.__dict__["details"]}

    return {"output": output, "error": None}

def convert_xls_to_pdf_msoffice(source: str, output: str)-> str:
    '''This fuction convert *.xls/*.xlsx files to pdf'''
    app = client.CreateObject("Excel.Application")
    try:
        sheets = app.Workbooks.Open(source)
        sheets.ExportAsFixedFormat(0, output)
        app.Quit()
    except Exception as exception:
        app.Quit()
        return {"output": None, "error": exception.__dict__["details"]}
    return {"output": output, "error": None}

def convert_ppt_to_pdf_msoffice(source: str, output: str)-> str:
    '''This fuction convert *.ppt/*.pptx files to pdf'''
    app = client.CreateObject("PowerPoint.Application")
    try:
        obj = app.Presentations.Open(source, False, False, False)
        obj.ExportAsFixedFormat(output, 2, PrintRange=None)
        app.Quit()
    except Exception as exception:
        app.Quit()
        return {"output": None, "error": exception.__dict__["details"]}
    return {"output": output, "error": None}
