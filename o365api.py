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
        conn = ClientContext(self._site).with_credentials(
            UserCredential(self._username, self._password)
        )
        return conn

    def _get_files_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f"{self._doc_lib}/{folder_name}"
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files

    def get_folder_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f"{self._doc_lib}/{folder_name}"
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Folders"]).get().execute_query()
        return root_folder.folders

    def download_file(self, file_name, folder_name):
        conn = self._auth()
        file_url = f"/sites/{self._site_name}/{self._doc_lib}/{folder_name}/{file_name}"
        file = File.open_binary(conn, file_url)
        return file.content

    def upload_file(self, file_name, folder_name, content):
        conn = self._auth()
        target_folder_url = f"/sites/{self._site_name}/{self._doc_lib}/{folder_name}"
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.upload_file(file_name, content).execute_query()
        return response
