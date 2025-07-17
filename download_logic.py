import os
import re
from io import BytesIO
import json
import pandas as pd
from datetime import datetime

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File as SPFile

def perform_download(sharepoint_url, sharepoint_folder_relative_path, manifest_filename, local_folder_id, data_folder_path, queue, stop_event, config_path, output_dir):
    """
    Performs the download process for a specific SharePoint folder using a specified manifest file.
    """
    error_log_file = None
    error_count = 0
    local_base_dir = None 
    try:
        # Use the provided output_dir for the error log
        error_log_file = os.path.join(output_dir, "download_errors.txt")

        if os.path.exists(error_log_file):
            os.remove(error_log_file)

        queue.put(("status", "Loading credentials..."))
        with open(config_path, 'r') as f:
            config = json.load(f)
        APP_USERNAME = config["APP_USERNAME"]
        APP_PASSWORD = config["APP_PASSWORD"]

        queue.put(("status", "Re-connecting to SharePoint..."))
        user_credentials = UserCredential(APP_USERNAME, APP_PASSWORD)
        ctx = ClientContext(sharepoint_url).with_credentials(user_credentials)
        web = ctx.web
        ctx.load(web, ["ServerRelativeUrl"])
        ctx.execute_query()

        site_relative_url = web.properties['ServerRelativeUrl']
        data_folder_url = f"{site_relative_url.rstrip('/')}/Shared Documents/{sharepoint_folder_relative_path}"

        queue.put(("status", f"Downloading '{manifest_filename}' from SharePoint folder '{sharepoint_folder_relative_path}'..."))
        index_file_url = f"{data_folder_url}/{manifest_filename}"
        response = SPFile.open_binary(ctx, index_file_url)
        response.raise_for_status()

        local_base_dir = os.path.join(data_folder_path, local_folder_id)
        if not os.path.exists(local_base_dir):
            os.makedirs(local_base_dir)

        local_index_path = os.path.join(local_base_dir, manifest_filename)
        with open(local_index_path, "wb") as f:
            f.write(response.content)
        queue.put(("file_info", f"Saved a local copy of '{manifest_filename}' to '{local_base_dir}'."))

        df = pd.read_csv(BytesIO(response.content), encoding='utf-8-sig')

        file_column_name = None
        if 'File' in df.columns:
            file_column_name = 'File'
            queue.put(("file_info", f"Using 'File' column from {manifest_filename}."))
        elif df.shape[1] > 0:
            file_column_name = df.columns[0]
            queue.put(("file_info", f"Warning: 'File' column not found in {manifest_filename}. Using the first column ('{file_column_name}') for file paths."))
        else:
            raise ValueError(f"'{manifest_filename}' in '{sharepoint_folder_relative_path}' is empty or does not contain any columns.")

        if file_column_name is None:
             raise ValueError(f"Could not determine a file path column in '{manifest_filename}' from '{sharepoint_folder_relative_path}'.")

        total_files = len(df)
        queue.put(("status", f"Found {total_files} files to download listed in '{manifest_filename}'."))

        for index, row in df.iterrows():
            if stop_event.is_set():
                queue.put(("status", "Download stopped by user."))
                queue.put(("stopped", (local_base_dir, error_count)))
                return

            queue.put(("progress", (index + 1, total_files)))
            relative_file_path = row[file_column_name]
            if not isinstance(relative_file_path, str) or not relative_file_path.strip():
                queue.put(("file_info", f"Skipping empty or invalid file path in row {index + 2} of {manifest_filename} (column '{file_column_name}')."))
                continue
            file_basename = os.path.basename(relative_file_path)
            queue.put(("filename", f"Processing: {file_basename}"))

            download_successful = False
            local_file_path = os.path.join(local_base_dir, relative_file_path.lstrip('\\/'))
            local_dir = os.path.dirname(local_file_path)
            if not os.path.exists(local_dir): os.makedirs(local_dir)

            try:
                full_path_suffix = relative_file_path.replace('\\', '/').lstrip('/')
                url_attempt_1 = f"{data_folder_url}/{full_path_suffix}"
                file_response = SPFile.open_binary(ctx, url_attempt_1)
                with open(local_file_path, "wb") as f: f.write(file_response.content)
                download_successful = True
            except Exception as e1:
                if "404" in str(e1) or "File Not Found" in str(e1) or "Cannot find" in str(e1):
                    queue.put(("file_info", f"Path '{relative_file_path}' not found for '{file_basename}' in '{sharepoint_folder_relative_path}'. Trying root of this folder..."))
                    try:
                        url_attempt_2 = f"{data_folder_url}/{file_basename}"
                        file_response = SPFile.open_binary(ctx, url_attempt_2)
                        with open(local_file_path, "wb") as f: f.write(file_response.content)
                        download_successful = True
                        queue.put(("file_info", f"Success! Found '{file_basename}' at the root of '{sharepoint_folder_relative_path}'."))
                    except Exception: pass
                else:
                    error_message = f"Failed to download '{relative_file_path}'. Non-404 Error: {type(e1).__name__} - {e1}"
                    queue.put(("file_error", error_message))
                    with open(error_log_file, "a", encoding='utf-8') as f: f.write(f"{datetime.now().isoformat()} - {error_message}\n")
                    error_count += 1

            if not download_successful:
                error_message = f"Failed to find or download '{relative_file_path}' (tried primary path and root of '{sharepoint_folder_relative_path}')."
                queue.put(("file_error", error_message))
                with open(error_log_file, "a", encoding='utf-8') as f: f.write(f"{datetime.now().isoformat()} - {error_message}\n")
                error_count += 1

        if not stop_event.is_set():
            queue.put(("progress", (total_files, total_files)))
            queue.put(("filename", "All files processed."))
            if error_count == 0:
                queue.put(("status", f"Download from '{sharepoint_folder_relative_path}' completed successfully."))
            else:
                queue.put(("status", f"Download from '{sharepoint_folder_relative_path}' completed with {error_count} errors."))
            queue.put(("done", (local_base_dir, error_count)))
    except Exception as e:
        detailed_error = f"Error during download from '{sharepoint_folder_relative_path}': {type(e).__name__} - {e}"
        queue.put(("error", detailed_error))
        if error_log_file:
            with open(error_log_file, "a", encoding='utf-8') as f: f.write(f"{datetime.now().isoformat()} - CRITICAL: {detailed_error}\n")
        fallback_dir = local_base_dir if local_base_dir else data_folder_path
        queue.put(("stopped", (fallback_dir, error_count + 1)))