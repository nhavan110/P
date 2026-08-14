import json
import os
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


SCOPES = ["https://www.googleapis.com/auth/drive"]
FILE_NAME = "MyPortfolio.xlsx"
FILE_PATH = Path(__file__).resolve().parent / "data" / FILE_NAME


def main():
    service_account_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID")

    if not service_account_json:
        raise RuntimeError("Missing GOOGLE_SERVICE_ACCOUNT_JSON environment variable")
    if not folder_id:
        raise RuntimeError("Missing GOOGLE_DRIVE_FOLDER_ID environment variable")
    if not FILE_PATH.exists():
        raise FileNotFoundError(f"File not found: {FILE_PATH}")

    credentials_info = json.loads(service_account_json)
    credentials = service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=SCOPES,
    )

    drive = build("drive", "v3", credentials=credentials)

    query = (
        f"'{folder_id}' in parents "
        f"and name = '{FILE_NAME}' "
        "and trashed = false"
    )

    response = drive.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name)",
        pageSize=10,
    ).execute()
    files = response.get("files", [])

    media = MediaFileUpload(
        str(FILE_PATH),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )

    if files:
        file_id = files[0]["id"]
        drive.files().update(
            fileId=file_id,
            media_body=media,
        ).execute()
        print(f"Updated Google Drive file: {FILE_NAME} ({file_id})")
    else:
        metadata = {
            "name": FILE_NAME,
            "parents": [folder_id],
        }
        created = drive.files().create(
            body=metadata,
            media_body=media,
            fields="id, name",
        ).execute()
        print(f"Created Google Drive file: {created['name']} ({created['id']})")


if __name__ == "__main__":
    main()
