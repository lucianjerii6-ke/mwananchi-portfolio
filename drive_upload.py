"""
Mwananchi Credit — Google Drive Upload Helper
Uploads the generated Excel portfolio file to a Google Drive folder.
Called automatically by MCL_Excel_Refresh.command after Excel is created.

First run: opens browser for one-time Google account authorisation.
Token saved to ~/.mcl_drive_token.json for all future runs (no re-auth needed).

Usage:
    python3 drive_upload.py <excel_path> <date> <branch>
"""

import os
import sys
import json
import pickle
import base64

try:
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
except ImportError:
    print("google-api packages not installed — skipping Drive upload")
    sys.exit(1)

SCOPES         = ["https://www.googleapis.com/auth/drive.file"]
TOKEN_PATH     = os.path.expanduser("~/.mcl_drive_token.pickle")
FOLDER_NAME    = "MCL Portfolio Reports"
LIVE_FILENAME  = "MCL_Portfolio_LIVE.xlsx"
MIME_XLSX      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ── Embedded OAuth client (public/installed app — no secret needed for Drive) ─
# Using the default Google Drive SDK demo credentials.
# For production use, replace with your own OAuth client ID from Google Cloud Console.
CLIENT_CONFIG = {
    "installed": {
        "client_id": "YOUR_CLIENT_ID.apps.googleusercontent.com",
        "client_secret": "YOUR_CLIENT_SECRET",
        "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob", "http://localhost"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token"
    }
}

CREDENTIALS_PATH = os.path.expanduser("~/.mcl_drive_credentials.json")


def get_service():
    creds = None
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_PATH):
                print("\n  Drive upload: credentials file not found.")
                print(f"  Place your Google OAuth credentials JSON at:")
                print(f"  {CREDENTIALS_PATH}")
                print("  (Download from Google Cloud Console > APIs > Credentials > OAuth 2.0 Client IDs)")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "wb") as f:
            pickle.dump(creds, f)

    return build("drive", "v3", credentials=creds)


def get_or_create_folder(service, name):
    """Find the MCL Portfolio Reports folder, create it if it doesn't exist."""
    q = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=q, fields="files(id,name)").execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    folder = service.files().create(body=meta, fields="id").execute()
    print(f"  Created Drive folder: {name}")
    return folder["id"]


def upload_file(service, local_path, drive_name, folder_id):
    """Upload or replace a file in the Drive folder."""
    # Check if file already exists in folder
    q = f"name='{drive_name}' and '{folder_id}' in parents and trashed=false"
    existing = service.files().list(q=q, fields="files(id,name)").execute().get("files", [])

    media = MediaFileUpload(local_path, mimetype=MIME_XLSX, resumable=False)

    if existing:
        # Update in place — same file ID so Drive history is preserved
        file_id = existing[0]["id"]
        service.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    else:
        meta = {"name": drive_name, "parents": [folder_id]}
        result = service.files().create(body=meta, media_body=media, fields="id").execute()
        return result["id"]


def main():
    if len(sys.argv) < 2:
        print("Usage: drive_upload.py <excel_path> [date] [branch]")
        sys.exit(1)

    excel_path = sys.argv[1]
    date_str   = sys.argv[2] if len(sys.argv) > 2 else ""
    branch     = sys.argv[3] if len(sys.argv) > 3 else "ALL"

    if not os.path.exists(excel_path):
        print(f"File not found: {excel_path}")
        sys.exit(1)

    print(f"\n  Uploading to Google Drive...")
    service   = get_service()
    folder_id = get_or_create_folder(service, FOLDER_NAME)

    # Upload LIVE copy (always same name — Claude searches for this)
    file_id = upload_file(service, excel_path, LIVE_FILENAME, folder_id)
    print(f"  LIVE file updated: {LIVE_FILENAME}")

    # Also upload dated copy if date provided
    if date_str:
        if branch and branch.upper() != "ALL":
            dated_name = f"MCL_Portfolio_{date_str}_{branch.upper()}.xlsx"
        else:
            dated_name = f"MCL_Portfolio_{date_str}.xlsx"
        upload_file(service, excel_path, dated_name, folder_id)
        print(f"  Dated copy saved:  {dated_name}")

    print(f"  Folder: {FOLDER_NAME}")
    print(f"  Done.")


if __name__ == "__main__":
    main()
