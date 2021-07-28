from os import path, remove
from subprocess import run, PIPE
from re import search
from shutil import copy2
from pathlib import Path, PurePosixPath
from datetime import datetime
from sys import platform

if platform == "win32":
    from comtypes import client

def __remove_files(temp_files_attach):
    """Remove temporary files"""
    for file_temp in temp_files_attach:
        if path.isfile(file_temp):
            remove(file_temp)

def __convert_to_pdf_libreoffice(source, output_dir, timeout=None)-> dict:
    """Convert MS Office files using LibreOffice"""
    output = None

    temp_filename = output_dir+"/"+datetime.now().\
        strftime("%Y%m%d%H%M%S%f")+path.basename(source)

    copy2(source, temp_filename)

    try:
        process = run(['soffice', '--headless', '--convert-to',\
            'pdf', '--outdir', path.dirname(source), temp_filename],\
                stdout=PIPE, stderr=PIPE,\
                    timeout=timeout, check=True)
        filename = search('-> (.*?) using filter', process.stdout.decode("latin-1"))
        __remove_files([temp_filename])
        output = filename.group(1).replace("\\", "/")

    except Exception as exception:
        return None

    return output

def __convert_doc_to_pdf_msoffice(source, output_dir):
    '''This fuction convert *.doc/*.docx files to pdf'''
    output = output_dir+"/"+datetime.now().\
        strftime("%Y%m%d%H%M%S%f")+Path(source).stem+".pdf"

    ws_pdf_format: int = 17
    app = client.CreateObject("Word.Application")
    try:
        doc = app.Documents.Open(source)
        doc.ExportAsFixedFormat(output, ws_pdf_format, Item=7, CreateBookmarks=0)
        app.Quit()

    except Exception as exception:
        app.Quit()
        return None

    return output

def __convert_xls_to_pdf_msoffice(source, output_dir):
    '''This fuction convert *.xls/*.xlsx files to pdf'''
    output = output_dir+"/"+datetime.now().\
        strftime("%Y%m%d%H%M%S%f")+Path(source).stem+".pdf"
    app = client.CreateObject("Excel.Application")
    try:
        sheets = app.Workbooks.Open(source)
        sheets.ExportAsFixedFormat(0, output)
        app.Quit()
    except Exception as exception:
        app.Quit()
        return None
    return output

def __convert_ppt_to_pdf_msoffice(source, output_dir):
    '''This fuction convert *.ppt/*.pptx files to pdf'''
    output = output_dir+"/"+datetime.now().\
        strftime("%Y%m%d%H%M%S%f")+Path(source).stem+".pdf"
    app = client.CreateObject("PowerPoint.Application")
    try:
        obj = app.Presentations.Open(source, False, False, False)
        obj.ExportAsFixedFormat(output, 2, PrintRange=None)
        app.Quit()
    except Exception as exception:
        app.Quit()
        return None
    return output

def __verify_source_is_supported_extension(file_extension):
    """This function very if source is supported extension"""
    supported_extensions = [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".xml"]
    return file_extension in supported_extensions

def convert(source, output_dir, soft="msoffice"):
    file_extension = PurePosixPath(source).suffix

    if __verify_source_is_supported_extension(file_extension) and path.isdir(output_dir):

        if platform == "win32" and soft == "msoffice":
            if file_extension in [".doc", ".docx", ".txt", ".xml"]:
                return __convert_doc_to_pdf_msoffice(source, output_dir)
            elif file_extension in [".xls", ".xlsx"]:
                return __convert_xls_to_pdf_msoffice(source, output_dir)
            elif file_extension in [".ppt", ".pptx"]:
                return __convert_ppt_to_pdf_msoffice(source, output_dir)

        elif platform == "win32" and soft == "libreoffice":
            return __convert_to_pdf_libreoffice(source, output_dir)
        else:
            return __convert_to_pdf_libreoffice(source, output_dir)
    else:
        return None
