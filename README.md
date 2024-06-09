# Gmail Attachment Downloader

This script allows you to connect to your Gmail account and download email attachments, specifically focusing on `.xlsx` files but easily adaptable for other file types. The attachments are saved to a specified folder, ensuring they are organized and accessible.

## Features

- **IMAP Connection**: Securely connects to your Gmail account using IMAP.
- **Folder Selection**: Searches for emails in a specified folder (e.g., "Daily_Excel").
- **Subject-Based Search**: Filters emails based on the subject line (default: "Daily Excel File").
- **Attachment Handling**: Downloads attachments, with a focus on `.xlsx` files, but can be modified for other types.
- **Configurable Paths**: Saves attachments to a configurable directory, ensuring easy organization and access.
- **Automatic Directory Creation**: Ensures the download directory exists, creating it if necessary.

## Prerequisites

- Python 3.12
- `imaplib` and `email` modules (part of Python's standard library)
- Gmail account with IMAP enabled
- `config.json` file for storing configuration details

## Setup

1. **Enable IMAP in Gmail**:
   - Go to Gmail settings.
   - Under the "Forwarding and POP/IMAP" tab, enable IMAP.

2. **Create `config.json`**:
   - Create a `config.json` file in the same directory as your script with the following content:
   ```json
   {
       "username": "your_email@gmail.com",
       "password": "your_app_specific_password",
       "folder": "Daily_Excel",
       "server": "imap.gmail.com",
       "download_path": "data"
   }
   ```
  - Replace the placeholder values with your actual email, app-specific password, and desired folder paths.

3.  Directory Structure:
  - Ensure the script and config.json are in the same directory.
  - The script will create the data folder (or any other specified folder) if it doesn't already exist.

## Usage
- Run the script to connect to your Gmail account, search for emails with the specified subject, and download any .xlsx attachments to the specified directory.

## Customization
- To adapt the script for other file types, modify the file extension check in the following line:
  ```
  if filename.lower().endswith('.xlsx'):
  ```
  For example, to download PDFs, change it to:
  ```
  if filename.lower().endswith('.pdf'):
  ```
