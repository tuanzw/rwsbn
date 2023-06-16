from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File


class SharePoint:
    def __init__(self, username, password, site, site_name, doc_lib):
        self._username = username
        self._password = password
        self._site = site
        self._site_name = site_name
        self._doc_lib = doc_lib

    def _auth(self):
        ctx = ClientContext(self._site).with_credentials(
            UserCredential(self._username, self._password)
        )
        return ctx

    def _get_files_list(self, folder_name):
        ctx = self._auth()
        target_folder_url = f"{self._doc_lib}/{folder_name}"
        root_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files

    def get_folder_list(self, folder_name):
        ctx = self._auth()
        target_folder_url = f"{self._doc_lib}/{folder_name}"
        root_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Folders"]).get().execute_query()
        return root_folder.folders

    def download_file(self, file_name, folder_name):
        ctx = self._auth()
        file_url = f"/sites/{self._site_name}/{self._doc_lib}/{folder_name}/{file_name}"
        file = File.open_binary(ctx, file_url)
        return file.content

    def upload_file(self, file_name, folder_name, content):
        ctx = self._auth()
        target_folder_url = f"/sites/{self._site_name}/{self._doc_lib}/{folder_name}"
        target_folder = ctx.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.upload_file(file_name, content).execute_query()
        return response

    def folder_existed(self, folder_name):
        ctx = self._auth()
        target_folder_url = f"/sites/{self._site_name}/{self._doc_lib}/{folder_name}"
        target_folder = (
            ctx.web.get_folder_by_server_relative_path(target_folder_url)
            .select(["Exists"])
            .get()
            .execute_query()
        )
        return target_folder.exists

    def add_folder(self, folder_name, child_folder_name):
        ctx = self._auth()
        target_folder_url = f"/sites/{self._site_name}/{self._doc_lib}/{folder_name}"
        target_folder = ctx.web.get_folder_by_server_relative_path(target_folder_url)
        child_folder = target_folder.add(child_folder_name)
        child_folder.execute_query()
        return child_folder
