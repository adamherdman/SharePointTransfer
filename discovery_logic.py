import os
import re
import json
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

def discover_data_folders(sharepoint_url, queue, config_path):
    """
    Connects to SharePoint, lists TOP-LEVEL folders in 'Shared Documents',
    and sends the list and site properties back to the GUI via a queue.
    """
    try:
        queue.put(("status", "Loading credentials..."))
        with open(config_path, 'r') as f:
            config = json.load(f)
        APP_USERNAME = config["APP_USERNAME"]
        APP_PASSWORD = config["APP_PASSWORD"]

        queue.put(("status", "Connecting to SharePoint for discovery..."))
        user_credentials = UserCredential(APP_USERNAME, APP_PASSWORD)
        ctx = ClientContext(sharepoint_url).with_credentials(user_credentials)
        web = ctx.web
        ctx.load(web, ["Title", "ServerRelativeUrl"])
        ctx.execute_query()
        queue.put(("status", f"Connected to site: {web.properties['Title']}"))
        
        queue.put(("web_props", web.properties))

        queue.put(("status", "Listing folders in 'Shared Documents'..."))
        site_relative_url = web.properties['ServerRelativeUrl']
        doc_library_url = f"{site_relative_url.rstrip('/')}/Shared Documents"

        root_folder = ctx.web.get_folder_by_server_relative_url(doc_library_url)
        sub_folders = root_folder.folders
        ctx.load(sub_folders)
        ctx.execute_query()

        all_folders = [f.name for f in sub_folders if f.name.lower() != "forms"]

        if not all_folders:
            raise FileNotFoundError("No folders were found in 'Shared Documents'. Please check the SharePoint site and permissions.")

        queue.put(("folders_found", all_folders))

    except Exception as e:
        queue.put(("error", str(e)))

def discover_sub_folders(sharepoint_url, parent_folder_url, queue, config_path):
    """
    Connects to SharePoint and lists sub-folders within a specific parent folder.
    This is intended for on-demand loading in the folder explorer dialog.
    """
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        APP_USERNAME = config["APP_USERNAME"]
        APP_PASSWORD = config["APP_PASSWORD"]

        user_credentials = UserCredential(APP_USERNAME, APP_PASSWORD)
        ctx = ClientContext(sharepoint_url).with_credentials(user_credentials)

        target_folder = ctx.web.get_folder_by_server_relative_url(parent_folder_url)
        sub_folders = target_folder.folders
        ctx.load(sub_folders)
        ctx.execute_query()

        all_sub_folders = [f.name for f in sub_folders]
        queue.put(("sub_folders_found", all_sub_folders))

    except Exception as e:
        queue.put(("error", str(e)))