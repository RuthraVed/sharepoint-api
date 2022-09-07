#!/usr/bin/python3
# ----------------------------------------------------------------------------
# Created By  : Abhishek Dev
# Created Date: 7th Sept 2022
# version = '1.0'
# ---------------------------------------------------------------------------
import io
from pathlib import Path

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File


class SharepointAutomation:
    """
    This is a class for Sharepoint Automation.

    Attributes
    -----------
    redirect_url : str
    client_id : str
    client_secret: str
    sharepoint_ctx: office365.sharepoint.client_context.ClientContext
        A sharepoint.client_context object
    """

    def __init__(self, settings) -> None:
        self.redirect_url = settings["redirect_url"]
        self.client_id = settings["client_id"]
        self.client_secret = settings["client_secret"]

        # Connecting to sharepoint
        context_auth = AuthenticationContext(self.redirect_url)
        context_auth.acquire_token_for_app(client_id=self.client_id, client_secret=self.client_secret)
        self.sharepoint_ctx = ClientContext(self.redirect_url, context_auth)


    def test_sharepoint_conection(self):
        """
        A static method to test sharepoint connection.
        """
        ctx = self.sharepoint_ctx
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print("Web site title: {0}".format(web.properties["Title"]))


    def list_files(self, sharepoint_folder_url):
        """
        A class method that returns a lists all files, in a specified sharepoint folder url

        :param object parent_folder: <class 'office365.sharepoint.folders.folder.Folder'>
        :return: A dict with filename as "key" and properties as "value" of files and their properties
        """
        sharepoint_folder = self.sharepoint_ctx.web.get_folder_by_server_relative_url(sharepoint_folder_url)
        files_properties_dict = {}
        sharepoint_folder.expand(["Files", "Folders"]).get().execute_query()
        for file in sharepoint_folder.files:
            files_properties_dict[file.properties["Name"]] = file.properties

        return files_properties_dict
    

    def upload_file(self, soure_file_path, sharepoint_folder_url):
        """
        A class method to upload any local file, to a specified sharepoint folder

        :param str soure_file_path:
        :param str sharepoint_folder_url:
        :param object sharepoint_ctx <class 'office365.sharepoint.client_context.ClientContext'>:
        :return: serverRelativeUrl
        """
        # A check to separate filename and filepath
        temp = sharepoint_folder_url.split("/")[-1]
        if ".xlsx" in temp:
            name = temp
            sharepoint_folder_url = sharepoint_folder_url.replace(name, "")

        with open(soure_file_path, "rb") as content_file:
            file_content = content_file.read()

        sharepoint_folder = self.sharepoint_ctx.web.get_folder_by_server_relative_url(sharepoint_folder_url)
        name = Path(soure_file_path).name
        target_file = sharepoint_folder.upload_file(name, file_content).execute_query()
        print("File uploaded to sharepoint at: {0}".format(target_file.serverRelativeUrl))
        return target_file.serverRelativeUrl

    def download_file(self, sharepoint_file_url, download_dir):
        """
        A class method to download any sharepoint_file, to a specified local folder

        :param str sharepoint_file_url --> Relative SP URL:
            Example: '/sites/Smart-Connected-Factory-Small-Ag-and-Turf-Sandbox/Shared Documents/notes.txt'
        :param str download_dir:
        :return: download_file_path
        """
        download_file_path = Path(download_dir) / Path(sharepoint_file_url).name
        with open(download_file_path, "wb") as local_file:
            file = self.sharepoint_ctx.web.get_file_by_server_relative_path(sharepoint_file_url).download(local_file).execute_query()
        
        print("[Ok] file has been downloaded into: {0}".format(download_file_path))
        return download_file_path

    def delete_file(self, sharepoint_file_url):
        """
        A class method to delete any sharepoint_file.

        :param str sharepoint_file_url --> Relative SP URL:
        Example: '/sites/Smart-Connected-Factory-Small-Ag-and-Turf-Sandbox/Shared Documents/notes.txt'
        :return: None
        """            
        file = self.sharepoint_ctx.web.get_file_by_server_relative_url(sharepoint_file_url)
        file.delete_object().execute_query()

    def copy_file(self, source_sp_file_url, destination_sp_folder_url):
        """
        A class method to copy a Sharepoint file from one folder to another Sharepoint folder

        :param str soure_file_path:
        :param str destination_folder_url:
        :return: None
        """
        sharepoint_file = File.open_binary(self.sharepoint_ctx, source_sp_file_url)
        sharepoint_folder = self.sharepoint_ctx.web.get_folder_by_server_relative_url(destination_sp_folder_url)
        target_file = sharepoint_folder.upload_file(Path(source_sp_file_url).name, sharepoint_file.content).execute_query()
        return target_file.serverRelativeUrl

    def move_file(self, source_sp_file_url, destination_sp_folder_url):
        """
        A class method to move file from one folder to another Sharepoint folder

        :param str soure_file_path:
        :param str destination_folder_url:
        :return: None
        """
        self.copy_file(source_sp_file_url, destination_sp_folder_url)
        self.delete_file(source_sp_file_url)



    def read_sp_file_in_memory(self, sharepoint_file_url):
        """
        A class method to read a sharepoint file, in-memory (without downloading)

        :param str sharepoint_file_url: A realative sharepoint file url
        :return: In-memory loaded file object
        """
        sharepoint_file_obj = File.open_binary(self.sharepoint_ctx, sharepoint_file_url)

        # Create an in-memory copy of the sharepoint file
        bytes_file_obj = io.BytesIO()
        bytes_file_obj.write(sharepoint_file_obj.content)

        return bytes_file_obj


    

    

    