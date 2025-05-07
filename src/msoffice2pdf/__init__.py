import sys
from os import path, remove
from subprocess import run, PIPE
from re import search
from shutil import copy2
from pathlib import Path, PurePosixPath
from datetime import datetime
from sys import platform

try:
    # Python 3.8+
    from importlib.metadata import version
except ImportError:
    from importlib_metadata import version

__version__ = version(__package__)
PLATFORMS_SUPPORTED = ["linux", "win32"]

if platform == "win32":
    from comtypes import client


def remove_files(temp_files_attach):
    """Remove temporary files."""
    for file_temp in temp_files_attach:
        if path.isfile(file_temp):
            remove(file_temp)


def convert_to_pdf_libreoffice(source, output_dir, timeout=None) -> str:
    """Convert MS Office files to PDF using LibreOffice."""
    output = None
    temp_filename = path.join(output_dir, datetime.now().strftime("%Y%m%d%H%M%S%f") + path.basename(source))
    copy2(source, temp_filename)

    try:
        process = run(
            ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, temp_filename],
            stdout=PIPE, stderr=PIPE, timeout=timeout, check=True
        )
        filename = search(r'-> (.*?) using filter', process.stdout.decode("utf-8"))
        remove_files([temp_filename])
        output = filename.group(1).replace("\\", "/") if filename else None
    except Exception as e:
        print(f"Error converting with LibreOffice: {e}")
        return None

    return output


def convert_doc_to_pdf_msoffice(source, output_dir):
    """Convert .doc/.docx files to PDF using MS Office."""
    output = path.join(output_dir, datetime.now().strftime("%Y%m%d%H%M%S%f") + Path(source).stem + ".pdf")
    ws_pdf_format = 17
    app = client.CreateObject("Word.Application")

    try:
        doc = app.Documents.Open(source)
        doc.ExportAsFixedFormat(output, ws_pdf_format, Item=7, CreateBookmarks=0)
        doc.Close()
    except Exception as e:
        print(f"Error converting Word document: {e}")
        return None
    finally:
        app.Quit()

    return output


def convert_xls_to_pdf_msoffice(source, output_dir):
    """Convert .xls/.xlsx files to PDF using MS Office."""
    output = path.join(output_dir, datetime.now().strftime("%Y%m%d%H%M%S%f") + Path(source).stem + ".pdf")
    app = client.CreateObject("Excel.Application")

    try:
        sheets = app.Workbooks.Open(source)
        sheets.ExportAsFixedFormat(0, output)
        sheets.Close()
    except Exception as e:
        print(f"Error converting Excel document: {e}")
        return None
    finally:
        app.Quit()

    return output


def convert_ppt_to_pdf_msoffice(source, output_dir):
    """Convert .ppt/.pptx files to PDF using MS Office."""
    output = path.join(output_dir, datetime.now().strftime("%Y%m%d%H%M%S%f") + Path(source).stem + ".pdf")
    app = client.CreateObject("PowerPoint.Application")

    try:
        presentation = app.Presentations.Open(source, False, False, False)
        presentation.ExportAsFixedFormat(output, 2, PrintRange=None)
        presentation.Close()
    except Exception as e:
        print(f"Error converting PowerPoint document: {e}")
        return None
    finally:
        app.Quit()

    return output


def verify_source_is_supported_extension(file_extension):
    """Verify if the source file extension is supported."""
    supported_extensions = [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".xml"]
    return file_extension in supported_extensions


def convert_using_msoffice(source, output_dir, file_extension):
    """Convert files to PDF using MS Office based on their extension."""
    if file_extension in [".doc", ".docx", ".txt", ".xml"]:
        return convert_doc_to_pdf_msoffice(source, output_dir)
    elif file_extension in [".xls", ".xlsx"]:
        return convert_xls_to_pdf_msoffice(source, output_dir)
    elif file_extension in [".ppt", ".pptx"]:
        return convert_ppt_to_pdf_msoffice(source, output_dir)
    else:
        return None


def convert(source, output_dir, soft=0):
    """Convert file to PDF using the selected software."""
    file_extension = PurePosixPath(source).suffix

    if verify_source_is_supported_extension(file_extension) and path.isdir(output_dir):
        if platform == "win32" and soft == 0:
            return convert_using_msoffice(source, output_dir, file_extension)
        elif platform in PLATFORMS_SUPPORTED and soft == 1:
            return convert_to_pdf_libreoffice(source, output_dir)
        elif platform in PLATFORMS_SUPPORTED:
            return convert_to_pdf_libreoffice(source, output_dir)
        else:
            raise Exception("Platform or conversion software not supported.")
    else:
        raise NotImplementedError("File extension not supported")


def cli():
    """CLI for file conversion."""
    import argparse
    import textwrap

    if "--version" in sys.argv:
        print(__version__)
        sys.exit(0)

    description = textwrap.dedent(
        """
        File Converter
        Convert MS Office files to PDF using MS Office or LibreOffice.
        """
    )

    formatter_class = lambda prog: argparse.RawDescriptionHelpFormatter(prog, max_help_position=32)
    parser = argparse.ArgumentParser(description=description, formatter_class=formatter_class)
    parser.add_argument("input", help="input file")
    parser.add_argument("output_dir", nargs="?", help="output file or folder")
    parser.add_argument("--soft", type=int, choices=[0, 1], default=0, help="software to use for conversion (0 for MS Office, 1 for LibreOffice)")
    parser.add_argument("--version", action="store_true", default=False, help="display version and exit")

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)
    else:
        args = parser.parse_args()

    result = convert(args.input, args.output_dir, args.soft)
    if result:
        print(f"Conversion successful: {result}")
    else:
        print("Conversion failed.")


if __name__ == "__main__":
    cli()
