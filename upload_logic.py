import os
import json
import paramiko
from datetime import datetime

def perform_upload(local_source_path, queue, stop_event, passphrase, config_path, output_dir):
    """
    Connects to SFTP and uploads a directory, sending progress to the GUI queue.
    """
    error_log_file = None
    error_count = 0
    client = None
    sftp = None
    try:
        # Use the provided output_dir for the error log
        error_log_file = os.path.join(output_dir, "upload_errors.txt")

        if os.path.exists(error_log_file):
            os.remove(error_log_file)

        queue.put(("status", "Loading SFTP configuration..."))
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        queue.put(("status", f"Connecting to {config['SFTP_HOSTNAME']}..."))
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(
            hostname=config["SFTP_HOSTNAME"],
            port=int(config.get("SFTP_PORT", 22)),
            username=config["SFTP_USERNAME"],
            key_filename=config["SFTP_PRIVATE_KEY_PATH"],
            passphrase=passphrase,
            timeout=15
        )
        sftp = client.open_sftp()
        queue.put(("status", "SFTP Connection successful."))

        remote_base_dir = os.path.basename(local_source_path)
        
        total_files = sum([len(files) for r, d, files in os.walk(local_source_path)])
        queue.put(("status", f"Found {total_files} files to upload."))
        
        try:
            sftp.stat(remote_base_dir)
        except FileNotFoundError:
            queue.put(("file_info", f"Creating remote directory: {remote_base_dir}"))
            sftp.mkdir(remote_base_dir)

        files_processed = 0
        for root, dirs, files in os.walk(local_source_path):
            if stop_event.is_set():
                queue.put(("status", "Upload stopped by user."))
                queue.put(("stopped", (remote_base_dir, error_count)))
                return

            for dir_name in dirs:
                local_dir = os.path.join(root, dir_name)
                relative_dir = os.path.relpath(local_dir, local_source_path)
                remote_dir = f"{remote_base_dir}/{relative_dir.replace(os.path.sep, '/')}"
                try:
                    sftp.stat(remote_dir)
                except FileNotFoundError:
                    queue.put(("file_info", f"Creating remote subdirectory: {remote_dir}"))
                    sftp.mkdir(remote_dir)
            
            for file_name in files:
                if stop_event.is_set():
                    queue.put(("status", "Upload stopped by user."))
                    queue.put(("stopped", (remote_base_dir, error_count)))
                    return
                
                files_processed += 1
                queue.put(("progress", (files_processed, total_files)))
                queue.put(("filename", f"Uploading: {file_name}"))

                local_file = os.path.join(root, file_name)
                relative_file = os.path.relpath(local_file, local_source_path)
                remote_file = f"{remote_base_dir}/{relative_file.replace(os.path.sep, '/')}"
                
                try:
                    sftp.put(local_file, remote_file)
                except Exception as e:
                    error_message = f"Failed to upload '{local_file}'. Reason: {e}"
                    queue.put(("file_error", error_message))
                    with open(error_log_file, "a", encoding='utf-8') as f:
                        f.write(f"{datetime.now().isoformat()} - {error_message}\n")
                    error_count += 1
        
        if not stop_event.is_set():
            queue.put(("progress", (total_files, total_files)))
            queue.put(("filename", "Upload complete."))
            if error_count == 0:
                queue.put(("status", "Upload completed successfully."))
            else:
                queue.put(("status", f"Upload completed with {error_count} errors."))
            queue.put(("done", (remote_base_dir, error_count)))

    except Exception as e:
        queue.put(("error", str(e)))
    finally:
        if sftp:
            sftp.close()
        if client:
            client.close()