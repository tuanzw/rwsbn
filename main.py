from dotenv import dotenv_values
import shutil
from pathlib import PurePath
from base64 import b64decode
from glob import glob
import os
from mailapi import MailDispatcher
from xlsxapi import ExcelDispatcher
from o365api import SharePoint
from logapi import logger
from jinja2 import Environment, FileSystemLoader

from datetime import datetime, timedelta

env = dotenv_values(".env")
# sharepoint
username = env.get("email")
password = b64decode(env.get("password").encode("utf-8")).decode("utf-8")
site = env.get("url_site")
site_name = env.get("site_name")
doc_lib = env.get("doc_library")

# mail
s_split = env.get("s_split")
attachment_folder = env.get("attachment_folder")
attachement_move_to_folder = env.get("attachement_move_to_folder")
sp_download_folder = env.get("sp_download_folder")
worksheet_name = env.get("worksheet_name")
trip_col_idx = int(env.get("trip_col_idx"))
tn_col_idx = int(env.get("tn_col_idx"))
bn_col_idx = int(env.get("bn_col_idx"))
first_row_idx = int(env.get("first_row_idx"))

# excel
company_name = env.get("company_name")
trucking_vendor = env.get("trucking_vendor")
project = env.get("project")
wdate = env.get("wdate")
site_id = env.get("site_id")
pattern = env.get("pattern")
email_type = env.get("email_type")
email_folder = env.get("email_folder")
email_move_to_folder = env.get("email_move_to_folder")
cc_list = env.get("cc_list")
valid_sites = env.get("valid_sites")

default_recipient = env.get("default_recipient")
forbidden_char = env.get("forbidden_char")

sharepoint_folder = env.get("sharepoint_folder")


def get_datemonth_str() -> str:
    p_date = datetime.now() + timedelta(days=1)
    return p_date.strftime("%m%d")


def build_mail_body(bn_dict: dict) -> str:
    fileloader = FileSystemLoader("templates")
    jinja_env = Environment(loader=fileloader)

    bookings = []
    for key, value in bn_dict.items():
        trip, truck, *_ = key.split(s_split)
        booking = (trip, truck, value)
        bookings.append(booking)
    htmlBody = jinja_env.get_template("mailbody.html").render(bookings=bookings)
    return htmlBody


# read files and return the content of files
def get_file_content(file_path):
    with open(file_path, "rb") as f:
        return f.read()


def not_well_prepared():
    if not os.path.isdir(PurePath(attachment_folder, attachement_move_to_folder)):
        logger.debug(
            f"__STOP__Not existed [{PurePath(attachment_folder, attachement_move_to_folder)}]"
        )
        return True


def run_app():
    try:
        logger.info("START")
        if not_well_prepared():
            exit()

        sp = SharePoint(username, password, site, site_name, doc_lib)
        mail_d = MailDispatcher(
            company_name,
            trucking_vendor,
            project,
            wdate,
            site_id,
            pattern,
            email_type,
            email_folder,
            email_move_to_folder,
            attachment_folder,
            s_split,
            cc_list,
            forbidden_char,
        )
        excel_d = ExcelDispatcher(
            s_split,
            attachment_folder,
            attachement_move_to_folder,
            worksheet_name,
            trip_col_idx,
            tn_col_idx,
            bn_col_idx,
            first_row_idx,
        )
        # check and create daily basic folder if not existed
        add_sp_folder(sp)

        mail_d.proceed_mail()
        excel_d.proceed_excel()

        # list of .xlsx file paths excluding [~$] temporiraly opening xlxs file
        file_paths = glob(f"{attachment_folder}[!~$]*.xlsx")
        for file_path in file_paths:
            file_name = os.path.basename(file_path)
            subject_dict = mail_d.extract_mail_subject(subject=file_name)

            pickup_site = subject_dict.get(site_id)
            sp_dest_folder = f"{sharepoint_folder}/{get_datemonth_str()}/{pickup_site}"
            sp.upload_file(file_name, sp_dest_folder, get_file_content(file_path))
            bn_dict = excel_d.get_bn_dict_from_file(file_path)

            to_list = env.get(subject_dict.get(trucking_vendor).lower())
            #
            if to_list is None:
                to_list = default_recipient
            subject = f"[BookingNumber]_{file_name}"
            try:
                body = build_mail_body(bn_dict)
            except ValueError as v:
                logger.debug(
                    f"__INVALID__Trip/TruckNumber in build_mail_body: {file_name}"
                )
                logger.debug(bn_dict)
                logger.debug(v)
                continue

            mail_d.send_booking_number_mail(body, subject, to_list, cc_list)

            # all good, move file to archive folder
            dest_path = PurePath(
                attachment_folder, attachement_move_to_folder, file_name
            )
            shutil.move(file_path, dest_path)
    except Exception as e:
        logger.exception(e)
    finally:
        logger.info("END")


def add_sp_folder(sp: SharePoint):
    datemonth_str = get_datemonth_str()
    current_day_folder = f"{sharepoint_folder}/{datemonth_str}"
    if sp.folder_existed(current_day_folder) is False:
        sp.add_folder(sharepoint_folder, datemonth_str)
        logger.info(f"__ADD folder: {current_day_folder}")
    for site in valid_sites.split(","):
        if sp.folder_existed(f"{current_day_folder}/{site}") is False:
            sp.add_folder(current_day_folder, site)
            logger.info(f"__ADD folder: {current_day_folder}/{site}")


# create directory if it doesn't exist
def create_dir(path):
    dir_path = PurePath(attachment_folder, sp_download_folder, path)
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)


# get back a list of subfolders from specific folder
def get_folders(sp: SharePoint, folder):
    l = []
    folder_obj = sp.get_folder_list(folder)
    for subfolder_obj in folder_obj:
        subfolder = "/".join([folder, subfolder_obj.name])
        l.append(subfolder)
    return l


# save the file to locate or remote location
def save_file(file_n, file_obj, subfolder):
    last_in_folder_name = subfolder.split("/")[-1]
    dir_path = PurePath(
        attachment_folder,
        sp_download_folder,
        f"{get_datemonth_str()}/{last_in_folder_name}",
    )
    file_dir_path = PurePath(dir_path, file_n)
    with open(file_dir_path, "wb") as f:
        f.write(file_obj)


def get_file(sp: SharePoint, file_n, folder):
    file_obj = sp.download_file(file_n, folder)
    save_file(file_n, file_obj, folder)


def get_files(sp: SharePoint, folder):
    files_list = sp._get_files_list(folder)
    for file in files_list:
        get_file(sp, file.name, folder)


def download_files(sp: SharePoint, sp_folder: str):
    folder_list = get_folders(sp, sp_folder)
    for folder in folder_list:
        for subfolder in get_folders(sp, folder):
            folder_list.append(subfolder)
    for folder in folder_list:
        last_in_folder_name = folder.split("/")[-1]
        create_dir(f"{get_datemonth_str()}/{last_in_folder_name}")
        get_files(sp, folder)


if __name__ == "__main__":
    run_app()
