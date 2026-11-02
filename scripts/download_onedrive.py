
"""Download the SNL master workbook from OneDrive."""

import io
import os
import zipfile
from pathlib import Path
from urllib.parse import quote

import msal
import requests

# Actual filename in OneDrive
ONEDRIVE_FILENAME = "SNL Master.xlsx"

# Local filename expected by GitHub workflow
LOCAL_FILENAME = "SNL Fall 26.xlsx"

# Existing Microsoft authentication
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
        "Microsoft login failed: "
        + str((result or {}).get("error", "no token"))
    )

headers = {
    "Authorization": "Bearer " + result["access_token"]
}

# Find the workbook anywhere in OneDrive
encoded = quote(
    ONEDRIVE_FILENAME.replace("'", "''"),
    safe=""
)

url = (
    "https://graph.microsoft.com/v1.0/"
    f"me/drive/root/search(q='{encoded}')"
)

matches = []

while url:
    response = requests.get(
        url, headers=headers, timeout=90
    )
    response.raise_for_status()
    data = response.json()

    matches.extend(
        item
        for item in data.get("value", [])
        if item.get("name", "").casefold()
        == ONEDRIVE_FILENAME.casefold()
        and "file" in item
    )

    url = data.get("@odata.nextLink")

if len(matches) != 1:
    raise RuntimeError(
        f"Expected exactly one file named "
        f"{ONEDRIVE_FILENAME!r}; found {len(matches)}. "
        "Check the filename or duplicate copies."
    )

# Download the matching Excel workbook
file_id = quote(matches[0]["id"], safe="")

url = (
    "https://graph.microsoft.com/v1.0/"
    f"me/drive/items/{file_id}/content"
)

response = requests.get(
    url, headers=headers, timeout=90
)
response.raise_for_status()

# Confirm that the downloaded file is Excel
with zipfile.ZipFile(io.BytesIO(response.content)) as workbook:
    if "xl/workbook.xml" not in workbook.namelist():
        raise ValueError(
            "Downloaded file is not a valid Excel workbook"
        )

# Save under the filename expected by GitHub
Path(LOCAL_FILENAME).write_bytes(response.content)

print(
    f"Downloaded {ONEDRIVE_FILENAME} "
    f"as {LOCAL_FILENAME} "
    f"({len(response.content)} bytes)"
)
