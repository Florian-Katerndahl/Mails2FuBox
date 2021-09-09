#!/home/ubuntu/.local/share/virtualenvs/mailproj-CiqbyxcH/bin/python
"""
https://requests.readthedocs.io/en/latest/user/quickstart/
https://elektrubadur.se/nextcloud-to-do-automation/
"""
import imaplib
import email
from email.message import EmailMessage
import pathlib
import requests
import argparse
import os
import re
from typing import Union
from datetime import datetime
from glob import glob
from helper import txt2pdf

long_termination_string = "-- \r\nfachschaft-geographie mailing list\r\nfachschaft-geographie@lists.fu-berlin.de\r\nhttps://lists.fu-berlin.de/listinfo/fachschaft-geographie\r\n"

def clean_up_path(filepath: Union[str, pathlib.Path]) -> str:
    filepath = str(pathlib.Path(filepath).absolute())
    filepath = re.subn("\\\\", "/", filepath, 100)[0]
    return filepath


def get_date_from_mail(in_string: str) -> datetime:
    part_of_interest = re.search(r"(?<=,\s)[0-9]{2}\s[A-Z][a-z]{2}\s[0-9]{4}", in_string).group(0)
    part_of_interest = datetime.strptime(part_of_interest, "%d %b %Y")
    return part_of_interest


def get_date_from_logfile(in_string: str) -> datetime:
    part_of_interest = re.search(r"[0-9]{2}-[0-9]{2}-[0-9]{4}", in_string).group(0)
    part_of_interest = datetime.strptime(part_of_interest, "%d-%m-%Y")
    return part_of_interest


def naive_filename(fn: str) -> bool:
    if re.match(r".*\..*", fn):
        return True
    else:
        return False


def is_protocol(fn: str) -> bool:
    if re.search(r"protokol[l]?", fn.lower()):
        return True
    else:
        return False


def get_newest_date(fn: str) -> datetime:
    try:
        log = open(fn, "rt")
        last_date = get_date_from_logfile(log.readlines()[-1])
        log.close()
    except FileNotFoundError:
        last_date = datetime.strptime("01-01-1970", "%d-%m-%Y")
    return last_date


def write_log_file(fn: str, mail_delivery: datetime) -> None:
    mail_delivery = mail_delivery.strftime("%d-%m-%Y")
    log = open(fn, "a+")
    log.write(f"{mail_delivery}\n")
    log.close()


def sanitize_files(fn: str) -> str:
    file_without_extension = re.search(r".*(?=\.docx|\.pdf|\.doc|\.txt|\.odf|\.odt)", fn).group(0)
    file_extension = re.search(r"(?<=\.).{3,4}$", fn).group(0)
    sanitized_file = re.subn(r"\.", "-", file_without_extension, 100)[0]
    return sanitized_file + "." + file_extension

def get_attachment(fn: str, dp: str, lf: str, dd: datetime) -> None:
    try:
        if filename is not None and naive_filename(filename) and is_protocol(filename):
            fp = dir_path / filename
            file = open(clean_up_path(fp), "wb")
            file.write(part.get_payload(decode=True))
            file.close()
            write_log_file(log_file, delivery_date)
    except OSError:
        pass

class C:
    pass


c = C()

parser = argparse.ArgumentParser()
parser.add_argument("--host", type=str, help="IMAP URI der Zedat")
parser.add_argument("-u", "--username", type=str, help="Zedat Benutzername")
parser.add_argument("-p", "--password", type=str, help="Zedat Passwort")
parser.add_argument("-ap", "--app_password", type=str, help="FU Box App-Passwort")
parser.add_argument("--WebDAV", type=str, dest="webdav_url",
                    help="https://box.fu-berlin.de/remote.php/dav/files/USERNAME/")
parser.add_argument("--box-destination", type=str, dest="dest", help="path/to/folder/", nargs="+")
# das könnte man auch weglassen, wenn man den Cronjob im entsprechenden Ordner startet (cd ... && ..python main.py)
parser.add_argument("-wd", type=str, dest="working_directory", help="/path/to/local/folder/")

parser.parse_args(namespace=c)

imap_conn = imaplib.IMAP4_SSL(host=c.host)

imap_conn.login(c.username, c.password)

imap_conn.select("INBOX", readonly=True)

typ, data = imap_conn.search(None, "SUBJECT", "[Fachschaft-Geographie]")

# eigentlich nur für erstes Setup
if not (dir_path := pathlib.Path(c.working_directory + "attachments")).exists():
    dir_path.mkdir()

log_file = clean_up_path(pathlib.Path(c.working_directory + "logfile.txt"))

for num in data[0].split():
    typ_2, data_2 = imap_conn.fetch(num, '(RFC822)')
    m = email.message_from_bytes(data_2[0][1], _class=EmailMessage)
    if m.is_multipart():
        date_as_string = m.get("delivery-date")
        delivery_date = get_date_from_mail(date_as_string)
        #if delivery_date >= get_newest_date(log_file):
        if delivery_date > get_newest_date(log_file):
            for part in m.walk():
                if part.is_multipart():  # Anscheinend interessiert irgendwas verschachteltes nicht
                    continue
                elif not part.get("content-disposition"):
                    continue
                elif part.get_content_maintype() == "text":
                    # haesslich, aber ich will's auch nicht mneu schreiben
                    if not (msg_text := part.get_payload(decode=True)) == long_termination_string:
                        filename = part.get_filename()
                        get_attachment(filename, dir_path, log_file, delivery_date)
#                        try:
#                            if filename is not None and naive_filename(filename) and is_protocol(filename):
#                                fp = dir_path / filename
#                                file = open(clean_up_path(fp), "wb")
#                                file.write(part.get_payload(decode=True))
#                                file.close()
#                                write_log_file(log_file, delivery_date)
#                            else:
#                                continue
#                        except OSError:
#                            continue
                    else:
                        continue
                filename = part.get_filename()
                try:
                    if filename is not None and naive_filename(filename) and is_protocol(filename):
                        fp = dir_path / filename
                        file = open(clean_up_path(fp), "wb")
                        file.write(part.get_payload(decode=True))
                        file.close()
                        write_log_file(log_file, delivery_date)
                    else:
                        continue
                except OSError:
                    continue

imap_conn.close()
imap_conn.logout()

# Dateikonvertierung
downloaded = glob(clean_up_path(dir_path) + "/*")

# Wenn keine neuen Dateien heruntergeladen wurde, kann hier bereits aufgehört werden
if not downloaded:
    exit()

# eigentlich nur für erstes Setup
if not (dir_path := pathlib.Path(c.working_directory + "converted")).exists():
    dir_path.mkdir()

for file_path in downloaded:
    file_path = clean_up_path(file_path)
    if not re.search(r"\.pdf$", file_path):
        new_name = re.sub(r"(?<=\.)docx|doc|odf|txt", "pdf", file_path).replace("attachments", "converted")
        txt2pdf.convert_to(clean_up_path(dir_path), file_path, 60)
        if not pathlib.Path(dir_path.absolute(), new_name).exists():
            print(f"Konnte Datei '{file_path}' nicht in PDF umwandeln!")
            continue
        else:
            os.unlink(file_path)
    else:
        filename = re.search(r"(?<=attachments/).*$", file_path).group(0)
        pathlib.Path(file_path).replace(pathlib.Path(dir_path.absolute(), filename))

# Dateien hochladen
converted = glob(clean_up_path(dir_path) + "/*")

session = requests.Session()
session.auth = (c.username, c.app_password)

dest_url = c.webdav_url + " ".join(c.dest)

upload_date = datetime.today().strftime("%d%m%Y")

for file in converted:
    file_path = clean_up_path(file)

    filename = re.search(r"(?<=converted/).*$", file_path).group(0)

    fd = open(file_path, "rb")

    payload = {
        "file": fd
    }

    tries = 0
    put_request = None

    while put_request is None or tries < 5:
        put_request = session.put(dest_url + upload_date + "__" + filename, files=payload)
        if 200 <= put_request.status_code < 300:
            fd.close()
            os.unlink(file_path)
            break
        else:
            tries += 1
    else:
        print(f"Konnte Datei {file_path} nicht hochladen.")

session.close()
