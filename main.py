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
attchement_move_to_folder = env.get("attchement_move_to_folder")
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


sharepoint_folder = env.get("sharepoint_folder")


def build_mail_body(bn_dict: dict) -> str:
    body = "Trip\tTruck Number \t Booking Number\n"
    for key, value in bn_dict.items():
        trip, truck = key.split(s_split)
        body = body + f"{trip}\t{truck}\t{value}\n"
    return body


# read files and return the content of files
def get_file_content(file_path):
    with open(file_path, "rb") as f:
        return f.read()


if __name__ == "__main__":
    try:
        logger.info("START")
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
        )
        excel_d = ExcelDispatcher(
            s_split,
            attachment_folder,
            attchement_move_to_folder,
            worksheet_name,
            trip_col_idx,
            tn_col_idx,
            bn_col_idx,
            first_row_idx,
        )
        sp = SharePoint(username, password, site, site_name, doc_lib)
        mail_d.proceed_mail()
        excel_d.proceed_excel()

        file_paths = glob(f"{attachment_folder}[!~$]*.xlsx")
        for file_path in file_paths:
            file_name = os.path.basename(file_path)
            sp.upload_file(file_name, sharepoint_folder, get_file_content(file_path))
            bn_dict = excel_d.get_bn_dict_from_file(file_path)
            body = build_mail_body(bn_dict)
            mail_d.send_booking_number_mail(file_name, body)

            # all good, move file to archive folder
            dest_path = PurePath(
                attachment_folder, attchement_move_to_folder, file_name
            )
            shutil.move(file_path, dest_path)
    except Exception as e:
        logger.debug(e)
    finally:
        logger.info("END")
