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

default_recipient = env.get("default_recipient")
forbidden_char = env.get("forbidden_char")

sharepoint_folder = env.get("sharepoint_folder")


def build_mail_body(bn_dict: dict) -> str:
    body = "Trip\tTruck Number \t Booking Number\n"
    for key, value in bn_dict.items():
        trip, truck, *_ = key.split(s_split)
        body = body + f"{trip}\t{truck}\t{value}\n"
    return body


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


if __name__ == "__main__":
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
        mail_d.proceed_mail()
        excel_d.proceed_excel()

        # list of .xlsx file paths excluding [~$] temporiraly opening xlxs file
        file_paths = glob(f"{attachment_folder}[!~$]*.xlsx")
        for file_path in file_paths:
            file_name = os.path.basename(file_path)
            sp.upload_file(file_name, sharepoint_folder, get_file_content(file_path))
            bn_dict = excel_d.get_bn_dict_from_file(file_path)

            subject_dict = mail_d.extract_mail_subject(subject=file_name)
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
