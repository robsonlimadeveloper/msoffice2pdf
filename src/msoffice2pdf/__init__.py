import sys
from os import path, remove
from subprocess import run, PIPE
from re import search
from shutil import copy2
from pathlib import Path, PurePosixPath
from datetime import datetime
from sys import platform

try:
    #3.8+
    from importlib.metadata import version
except ImportError:
    from importlib_metadata import version

__version__ = version(__package__)
platforms_supported = ["linux", "win32"]

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

def __convert_using_msoffice(source, output_dir, file_extension):
    """"""
    if file_extension in [".doc", ".docx", ".txt", ".xml"]:
                return __convert_doc_to_pdf_msoffice(source, output_dir)
    elif file_extension in [".xls", ".xlsx"]:
        return __convert_xls_to_pdf_msoffice(source, output_dir)
    elif file_extension in [".ppt", ".pptx"]:
        return __convert_ppt_to_pdf_msoffice(source, output_dir)

def convert(source, output_dir, soft=0):
    """This function convert file by software selected"""
    file_extension = PurePosixPath(source).suffix

    if __verify_source_is_supported_extension(file_extension) and path.isdir(output_dir):

        if platform == "win32" and soft == 0:
            return __convert_using_msoffice(source, output_dir, file_extension)
        elif platform in platforms_supported and soft == 1:
            return __convert_to_pdf_libreoffice(source, output_dir)
        elif platform in platforms_supported:
            return __convert_to_pdf_libreoffice(source, output_dir)
        else:
            raise Exception("Platform or conversion software not supported.")
    else:
        raise NotImplementedError("File extension not supported")

def cli():
    """This function receive params to use the convertion"""
    import textwrap
    import argparse

    if "--version" in sys.argv:
        print(__version__)
        sys.exit(0)

    description = textwrap.dedent(
        """"""
    )

    formatter_class = lambda prog: argparse.RawDescriptionHelpFormatter(\
        prog, max_help_position=32)
    parser = argparse.ArgumentParser(description=description,\
        formatter_class=formatter_class)
    parser.add_argument("input",help="input file")
    parser.add_argument("output_dir", nargs="?", help="output file or folder")
    parser.add_argument("--soft", action="store_true", default="msoffice",\
        help="choose software to conversion")
    parser.add_argument("--version", action="store_true", default=False,\
        help="display version and exit")

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)
    else:
        args = parser.parse_args()

    convert(args.input, args.output_dir, args.soft)
