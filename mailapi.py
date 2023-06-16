from dotenv import dotenv_values

import re
from pathlib import PurePath
import win32com.client as win32

from logapi import logger

env = dotenv_values(".env")


class MailDispatcher:
    def __init__(
        self,
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
    ):
        self._company_name = company_name
        self._trucking_vendor = trucking_vendor
        self._project = project
        self._wdate = wdate
        self._site_id = site_id
        self._pattern = pattern
        self._email_type = email_type
        self._email_folder = email_folder
        self._email_move_to_folder = email_move_to_folder
        self._attachment_folder = attachment_folder
        self._s_split = s_split
        self._cc_list = cc_list
        self._forbidden_char = forbidden_char

    def proceed_mail(self):
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Here the name of the folders
        # 3 = Drafts
        # 4 = Outbox
        # 5 = Sent items
        # 6 = Inbox
        inbox = outlook.GetDefaultFolder(6)

        # if email_folder found, work on it else work on Inbox
        wfolder = inbox
        for folder in inbox.Folders:
            if folder.name == self._email_folder:
                wfolder = folder
                break
        logger.info(f"Working on folder {wfolder.Name}")

        # Stop if there is no email_move_to_folder existed
        try:
            to_folder = wfolder.Folders(self._email_move_to_folder)
        except Exception as e:
            print(
                f"Please create {self._email_move_to_folder} as subfolder of {wfolder.Name}"
            )
            logger(
                f"Please create {self._email_move_to_folder} as subfolder of {wfolder.Name}"
            )
            exit()

        # Sound goood, working on emails received on today
        # https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/filtering-items-using-a-date-time-comparison
        messages = wfolder.Items
        messages = messages.Restrict(
            '@SQL=%today("urn:schemas:httpmail:datereceived")%'
        )

        # messages.Sort("[ReceivedTime]", Descending=True)
        to_move_messages = []
        for message in messages:
            # if not a email or subject pattern is not valid, skip
            if message.Class != 43:
                continue
            if re.search(self._pattern, message.Subject) is None:
                logger.info(f"__SKIP__Not well-formed Mail Subject__{message.Subject}")
                continue
            # extract sender from eamil, if not in potentail_senders, skip
            subject_dict = self.extract_mail_subject(message.Subject)
            potential_senders = env.get(subject_dict.get(self._trucking_vendor).lower())
            if potential_senders is None:
                logger.info(
                    f"__SKIP__Senders not defined__{subject_dict.get(self._trucking_vendor)}__{message.Subject}"
                )
                continue
            if self.get_sender_email_string(message) not in potential_senders.lower():
                logger.info(
                    f"__SKIP__{self.get_sender_email_string(message)} not in list__[{potential_senders}]__{message.Subject}"
                )
                continue

            # proceed to save .xlsx files in attachment of valid mail
            to_move_messages.append(message)
            attachments = message.Attachments
            if len(attachments) == 0:
                logger.info(f"__No attachment__{message.Subject}")
            for attachment in attachments:
                if attachment.FileName.split(".")[-1] == "xlsx":
                    file_path = PurePath(
                        self._attachment_folder,
                        f"[{self.build_key(subject_dict)}]_{attachment.FileName}",
                    )
                    attachment.SaveAsFile(file_path)
        # all good, move proceeded email to email_move_to_folder
        for message in to_move_messages:
            message.Move(to_folder)

    def extract_mail_subject(self, subject) -> dict:
        result = {}
        # remove illegal filename character
        subject = "".join(c for c in subject if c not in self._forbidden_char)
        l = subject.split("-")
        r, t, p, w, s, *_ = l
        result[self._company_name] = r.split("[")[-1].strip()
        result[self._wdate] = w.strip()
        result[self._site_id] = s.split("]")[0].strip()
        result[self._project] = p.strip()
        result[self._trucking_vendor] = t.split("]")[0].strip()
        return result

    def build_key(self, k_dict):
        return f"{k_dict.get(self._company_name)}{self._s_split}{k_dict.get(self._trucking_vendor)}{self._s_split}{k_dict.get(self._project)}{self._s_split}{k_dict.get(self._wdate)}{self._s_split}{k_dict.get(self._site_id)}"

    def is_proceeded_mail_subject(self, subject, p_dict):
        k_dict = self.extract_mail_subject(subject)
        p_key = self.build_key(k_dict)
        return p_dict.get(p_key)

    def build_proceeded_dict(self, k_dict, p_dict=None):
        if p_dict is None:
            p_dict = {}
        p_key = self.build_key(k_dict)
        if p_dict.get(p_key) is None:
            p_dict[p_key] = 1
        return p_dict

    def get_sender_email_string(self, message):
        sender_email: str = ""
        if message.SenderEmailType == self._email_type:
            sender_email = message.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            sender_email = message.SenderEmailAddress

        return sender_email.lower()

    def send_booking_number_mail(self, body, subject, to_list, cc_list):
        # https://www.codeforests.com/2020/06/05/how-to-send-email-from-outlook/
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)  # 0 means mail
        mail.To = to_list
        mail.Subject = subject
        mail.HTMLBody = body
        mail.CC = cc_list
        mail.Send()
