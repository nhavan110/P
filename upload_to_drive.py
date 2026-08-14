import json
import os
import sys
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload


SCOPES = ["https://www.googleapis.com/auth/drive"]
FILE_NAME = "MyPortfolio.xlsx"
FILE_PATH = Path(__file__).resolve().parent / "data" / FILE_NAME


def load_credentials_info(raw: str) -> dict:
    """Parse GOOGLE_SERVICE_ACCOUNT_JSON, với thông báo lỗi rõ ràng nếu secret
    bị dán sai (thiếu ký tự, escape sai \\n trong private_key, v.v.)."""
    try:
        info = json.loads(raw)
    except json.JSONDecodeError as e:
        raise RuntimeError(
            "GOOGLE_SERVICE_ACCOUNT_JSON không phải JSON hợp lệ "
            f"(lỗi tại dòng {e.lineno}, cột {e.colno}: {e.msg}). "
            "Kiểm tra lại bạn đã copy TOÀN BỘ nội dung file service-account "
            "JSON (kể cả dấu { } ở đầu/cuối) vào GitHub Secret, không bị "
            "GitHub tự động xóa xuống dòng."
        ) from e

    required = {"client_email", "private_key", "token_uri"}
    missing = required - info.keys()
    if missing:
        raise RuntimeError(
            f"GOOGLE_SERVICE_ACCOUNT_JSON thiếu field: {sorted(missing)}. "
            "Có vẻ đây không phải file service-account key hợp lệ."
        )
    return info


def main():
    service_account_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID")

    if not service_account_json:
        raise RuntimeError("Missing GOOGLE_SERVICE_ACCOUNT_JSON environment variable")
    if not folder_id:
        raise RuntimeError("Missing GOOGLE_DRIVE_FOLDER_ID environment variable")
    if not FILE_PATH.exists():
        raise FileNotFoundError(f"File not found: {FILE_PATH}")

    credentials_info = load_credentials_info(service_account_json)
    print(f"ℹ️  Đăng nhập bằng service account: {credentials_info.get('client_email')}")

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

    try:
        response = drive.files().list(
            q=query,
            spaces="drive",
            fields="files(id, name)",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
    except HttpError as e:
        print(f"❌ Lỗi khi tìm file trong Drive (folder_id={folder_id}):", file=sys.stderr)
        print(e.error_details if hasattr(e, "error_details") else e, file=sys.stderr)
        print(
            "→ Kiểm tra: (1) GOOGLE_DRIVE_FOLDER_ID đúng chưa (chỉ lấy phần "
            "ID trong URL, không phải cả link), (2) folder đó đã được share "
            f"quyền 'Editor' cho {credentials_info.get('client_email')} chưa.",
            file=sys.stderr,
        )
        raise
    files = response.get("files", [])

    media = MediaFileUpload(
        str(FILE_PATH),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )

    try:
        if files:
            file_id = files[0]["id"]
            drive.files().update(
                fileId=file_id,
                media_body=media,
                supportsAllDrives=True,
            ).execute()
            print(f"✅ Updated Google Drive file: {FILE_NAME} ({file_id})")
        else:
            metadata = {
                "name": FILE_NAME,
                "parents": [folder_id],
            }
            created = drive.files().create(
                body=metadata,
                media_body=media,
                fields="id, name",
                supportsAllDrives=True,
            ).execute()
            print(f"✅ Created Google Drive file: {created['name']} ({created['id']})")
    except HttpError as e:
        print("❌ Lỗi khi upload file lên Drive:", file=sys.stderr)
        print(e.error_details if hasattr(e, "error_details") else e, file=sys.stderr)
        if e.resp is not None and e.resp.status == 403:
            print(
                "→ Lỗi 403 thường do: folder Drive đích không share quyền "
                f"Editor cho {credentials_info.get('client_email')}, "
                "hoặc Google Drive API chưa được bật (enable) cho project "
                "chứa service account này trong Google Cloud Console.",
                file=sys.stderr,
            )
        raise


if __name__ == "__main__":
    main()
