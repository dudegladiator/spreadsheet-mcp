"""Google Service Account Authentication for Sheets API."""

import os
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build, Resource

# Scopes required for full Sheets API access
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",  # Full read/write access
    "https://www.googleapis.com/auth/drive",  # Full Drive access for creating spreadsheets
]


def get_credentials_path() -> Path:
    """Get the path to the service account credentials file."""
    # Check environment variable first
    env_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE")
    if env_path:
        path = Path(env_path)
        if path.exists():
            return path

    # Check common locations
    common_paths = [
        Path("./credentials/service-account.json"),
        Path("./credentials/service-cred.json"),  # Support existing file
        Path("./service-account.json"),
        Path("./service-cred.json"),
        Path.home() / ".config" / "spreadsheet-mcp" / "service-account.json",
    ]

    for path in common_paths:
        if path.exists():
            return path

    raise FileNotFoundError(
        "Service account credentials not found. "
        "Set GOOGLE_SERVICE_ACCOUNT_FILE environment variable or place "
        "service-account.json in ./credentials/ directory."
    )


def get_credentials() -> service_account.Credentials:
    """Load and return Google service account credentials."""
    credentials_path = get_credentials_path()
    credentials = service_account.Credentials.from_service_account_file(
        str(credentials_path), scopes=SCOPES
    )
    return credentials


def get_sheets_service() -> Resource:
    """Build and return the Google Sheets API service."""
    credentials = get_credentials()
    service = build("sheets", "v4", credentials=credentials)
    return service


def get_drive_service() -> Resource:
    """Build and return the Google Drive API service (for creating spreadsheets)."""
    credentials = get_credentials()
    service = build("drive", "v3", credentials=credentials)
    return service
