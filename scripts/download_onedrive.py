
"""Download the SNL Master workbook from OneDrive."""

import io
import os
import zipfile
from pathlib import Path
from urllib.parse import quote

import msal
import requests


# Select the canonical workbook even when other copies use the same name.
ONEDRIVE_PATH = "/SNL/SNL Master.xlsx"
LOCAL_FILENAME = "SNL Master.xlsx"


# Microsoft authentication
cache = msal.SerializableTokenCache()
cache.deserialize(os.environ["MSAL_CACHE"])

app = msal.PublicClientApplication(
    os.environ["MS_CLIENT_ID"],
    authority="https://login.microsoftonline.com/consumers",
    token_cache=cache,
)

accounts = app.get_accounts()

if not accounts:
    raise RuntimeError(
        "No Microsoft account found in token cache"
    )

result = app.acquire_token_silent_with_error(
    ["Files.Read"],
    account=accounts[0],
)

if not result or "access_token" not in result:
    raise RuntimeError(
        "Microsoft authentication failed: "
        + str((result or {}).get("error", "No token"))
    )


session = requests.Session()
session.headers.update({
    "Authorization": "Bearer " + result["access_token"]
})

BASE = "https://graph.microsoft.com/v1.0/me/drive"


def get_items(url):
    """Read all items from a OneDrive folder."""

    while url:
        response = session.get(url, timeout=90)
        response.raise_for_status()

        data = response.json()

        for item in data.get("value", []):
            yield item

        url = data.get("@odata.nextLink")


# Search OneDrive folders and subfolders
folders = [(BASE + "/root/children", "")]
matches = []
excel_files = []

while folders:
    url, folder_path = folders.pop(0)

    for item in get_items(url):
        name = item.get("name", "")
        path = f"{folder_path}/{name}"

        if "folder" in item:
            folder_id = quote(item["id"], safe="")

            folders.append((
                f"{BASE}/items/{folder_id}/children",
                path,
            ))

        elif "file" in item:
            if name.lower().endswith(".xlsx"):
                excel_files.append(path)

            if path.casefold() == ONEDRIVE_PATH.casefold():
                matches.append((item, path))


# Verify exactly one workbook was found
if len(matches) != 1:
    print("Excel workbooks found in OneDrive:")

    for path in sorted(excel_files)[:60]:
        print(" ", path)

    raise RuntimeError(
        f"Expected exactly one workbook at "
        f"{ONEDRIVE_PATH!r}, "
        f"but found {len(matches)}."
    )


item, path = matches[0]
print("Found workbook:", path)


# Download the workbook
file_id = quote(item["id"], safe="")
url = f"{BASE}/items/{file_id}/content"

response = session.get(url, timeout=90)
response.raise_for_status()


# Verify it is a valid Excel workbook
with zipfile.ZipFile(
    io.BytesIO(response.content)
) as z:
    if "xl/workbook.xml" not in z.namelist():
        raise ValueError(
            "Downloaded file is not a valid Excel workbook"
        )


# Save with permanent filename
Path(LOCAL_FILENAME).write_bytes(response.content)

print(
    "Workbook downloaded successfully:",
    LOCAL_FILENAME,
    len(response.content),
    "bytes",
)
