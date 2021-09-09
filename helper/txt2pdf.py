"""
https://michalzalecki.com/converting-docx-to-pdf-using-python/
"""
import sys
import subprocess
import re


def convert_to(folder, source, timeout=None):
    args = [libreoffice_exec(), '--headless', '--convert-to', 'pdf', '--outdir', folder, source]

    process = subprocess.run(args, timeout=timeout)


def libreoffice_exec():
    if sys.platform == 'darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    if sys.platform == 'win32':
        return 'soffice'
    return 'libreoffice'
