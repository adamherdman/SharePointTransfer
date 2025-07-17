# Data Transfer Hub

![Python Version](https://img.shields.io/badge/python-3.9+-blue.svg)

A desktop application designed to streamline the transfer of large datasets from a SharePoint site to an SFTP server. The tool features a user-friendly graphical interface, manifest-driven downloads, and secure uploads, making complex data migration tasks simple and repeatable.

Main reason for this project was due to the 20GB download limitation within sharepoint.

<!-- Replace with a more current screenshot or GIF of your application -->

---

## Table of Contents

-   [Key Features](#key-features)
-   [Workflow](#workflow)
-   [Prerequisites](#prerequisites)
-   [Installation](#installation)
-   [Configuration](#configuration)
-   [How to Use](#how-to-use)
-   [Building an Executable](#building-an-executable)
-   [Dependencies](#dependencies)
-   [License](#license)
-   [Contact](#contact)

## Key Features

*   **Graphical User Interface**: A clean and modern UI built with `customtkinter` for intuitive operation.
*   **Manifest-Driven Downloads**: Reads a `.csv` manifest file from SharePoint to determine exactly which files to download, including their subdirectory structure.
*   **Interactive SharePoint Browser**: A built-in dialog to navigate SharePoint's "Shared Documents" library and select data folders on the fly.
*   **Secure SFTP Uploads**: Uses `paramiko` for secure, key-based authentication (with passphrase support) to upload the data to an SFTP server.
*   **External Configuration**: All sensitive credentials and paths are managed in an external `config.json` file, keeping them separate from the source code.
*   **In-App Config Editor**: A built-in dialog to easily view and modify the application's configuration without manually editing the JSON file.
*   **Real-time Progress**: Provides live feedback on status, progress bars for downloads/uploads, and a detailed logging window.
*   **Error Handling & Logging**: Generates `download_errors.txt` and `upload_errors.txt` to capture any issues during the transfer process for easy debugging.
*   **Standalone Executable Support**: Designed to be bundled into a single executable file using PyInstaller for easy distribution to non-technical users.

## Workflow

The application simplifies a two-stage data transfer process:

1.  **Download from SharePoint:**
    *   The user provides the URL of the SharePoint site.
    *   The application connects and allows the user to browse for a specific folder within the "Shared Documents" library.
    *   Inside this folder, a `.csv` manifest file must exist. The user provides its name.
    *   The application reads the manifest and downloads all listed files to a local `Data` directory, preserving the folder structure.

2.  **Upload to SFTP:**
    *   After a successful download, the user can start the upload process.
    *   The application lists the available local data folders (e.g., `D12345`).
    *   The user selects a folder and provides the passphrase for their SFTP private key.
    *   The application securely connects to the SFTP server and uploads the entire contents of the selected folder.

## Prerequisites

*   Python 3.9 or newer.
*   Access credentials for a SharePoint site. (App account required, ie. an account without MFA.)
*   Access credentials and a private key for an SFTP server.

## Installation

1.  **Clone the repository:**
    ```sh
    git clone https://github.com/your-username/data-transfer-hub.git
    cd data-transfer-hub
    ```

2.  **Create a virtual environment (recommended):**
    ```sh
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3.  **Install the required dependencies:**
    ```sh
    pip install -r requirements.txt
    ```

## Configuration

Before running the application, you must set up the `config.json` file. You can create this file in the same directory as the application or use the built-in configuration editor.

1.  Launch the application and click the **Edit Configuration** button.
2.  Fill in all the required fields.
3.  Click **Save**. This will create a `config.json` file for you.

Alternatively, you can create the file manually:

**`config.json`**
```json
{
  "DATA_FOLDER_PATH": "C:/Path/To/Your/Local/Data/Folder",
  "APP_USERNAME": "your-email@your-tenant.com",
  "APP_PASSWORD": "your-sharepoint-password",
  "SFTP_HOSTNAME": "sftp.yourserver.com",
  "SFTP_PORT": "22",
  "SFTP_USERNAME": "your-sftp-username",
  "SFTP_PRIVATE_KEY_PATH": "C:/Path/To/Your/id_rsa"
}
```

## How to Use

1.  **Run the application:**
    ```sh
    python main.py
    ```
2.  **Enter the SharePoint Site URL** in the main window.
3.  **Click "Start Download"**:
    *   The app will connect to SharePoint and open a folder explorer.
    *   Navigate to and select the parent folder containing your data and manifest file.
    *   Enter the exact filename of your `.csv` manifest file when prompted.
    *   If the SharePoint folder name doesn't contain a `Dxxxx` style ID, you will be asked to provide one. This ID is used to name the local data folder.
    *   The download will begin, with progress shown in the UI.
4.  **Click "Start Upload"**:
    *   A dialog will show all locally downloaded data folders.
    *   Select the folder you wish to upload.
    *   Enter the passphrase for your SFTP private key (if it has one).
    *   The upload will begin, with progress shown in the UI.


## Dependencies

*   [customtkinter](https://github.com/TomSchimansky/CustomTkinter): For the user interface.
*   [Office365-REST-Python-Client](https://github.com/vgrem/Office365-REST-Python-Client): For connecting to and interacting with SharePoint.
*   [pandas](https://pandas.pydata.org/): For reading and parsing the `.csv` manifest files.
*   [paramiko](http://www.paramiko.org/): For handling the SFTP connection and file transfers.
*   [Pillow](https://python-pillow.org/): For handling images in the UI.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Contact

Created by Adam Herdman - [adam.herdman@nhs.net](mailto:adam.herdman@nhs.net)





