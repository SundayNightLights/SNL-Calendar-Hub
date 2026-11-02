
"""Download the SNL master workbook from OneDrive."""

import io
import os
import zipfile
from pathlib import Path
from urllib.parse import quote

import msal
import requests

# Name of the workbook stored in OneDrive
ONEDRIVE_FILENAME = "SNL Master.xlsx"

# Keep this name for the existing GitHub workflow
LOCAL_FILENAME = "SNL Fall 26.xlsx"

# Authenticate using existing GitHub secrets
cache = msal.SerializableTokenCache()
cache.deserialize(os.environ["MSAL_CACHE"])

app = msal.PublicClientApplication(
    os.environ["MS_CLIENT_ID"],
    authority="https://login.microsoftonline.com/consumers",
    token_cache=cache,
)

accounts = app.get_accounts()

if not accounts:
    raise RuntimeError("No Microsoft account in token cache")

result = app.acquire_token_silent_with_error(
    ["Files.Read"], account=accounts[0]
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
    """Read every page of a OneDrive folder."""
    while url:
        response = session.get(url, timeout=90)
        response.raise_for_status()
        data = response.json()

        for item in data.get("value", []):
            yield item

        url = data.get("@odata.nextLink")


# Search actual folders rather than the search index
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
                path
            ))

        elif "file" in item:
            if name.lower().endswith(".xlsx"):
                excel_files.append(path)

            if name.casefold() == ONEDRIVE_FILENAME.casefold():
                matches.append((item, path))


# Stop safely if file cannot be identified
if len(matches) != 1:
    print("\nExcel workbooks visible in OneDrive:")

    for path in sorted(excel_files)[:60]:
        print(" ", path)

    raise RuntimeError(
        f"Expected one workbook named "
        f"{ONEDRIVE_FILENAME!r}, "
        f"but found {len(matches)}. "
        "Check the filename and OneDrive account."
    )


item, path = matches[0]
print("Found workbook:", path)

# Download by its OneDrive item ID
file_id = quote(item["id"], safe="")

url = f"{BASE}/items/{file_id}/content"

response = session.get(url, timeout=90)
response.raise_for_status()

# Verify the Excel file
with zipfile.ZipFile(io.BytesIO(response.content)) as z:
    if "xl/workbook.xml" not in z.namelist():
        raise ValueError(
            "Downloaded file is not a valid Excel workbook"
        )

# Save using the name expected by the workflow
Path(LOCAL_FILENAME).write_bytes(response.content)

print(
    "Workbook downloaded successfully:",
    LOCAL_FILENAME,
    len(response.content),
    "bytes"
)
