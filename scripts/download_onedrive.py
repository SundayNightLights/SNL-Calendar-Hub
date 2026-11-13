
"""Download the SNL Master workbook from OneDrive."""

import io
import os
import zipfile
from pathlib import Path
from urllib.parse import parse_qs, quote, urlsplit

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


# Download via a fresh Microsoft Graph redirect. Microsoft Graph returns a
# short-lived, preauthorized URL. For some OneDrive Personal accounts the
# my.microsoftpersonalcontent.com endpoint requires the URL's *tempauth*
# credential as a Bearer header (not the Graph OAuth token).
#
# Keep OAuth tokens confined to graph.microsoft.com. Never log the signed URL,
# tempauth query value, or response exception (which may include the signed URL).
file_id = quote(item["id"], safe="")
graph_content_url = f"{BASE}/items/{file_id}/content"
downloaded = None

for attempt in range(1, 4):
    graph_response = session.get(
        graph_content_url, timeout=90, allow_redirects=False
    )
    if graph_response.status_code == 200:
        downloaded = graph_response.content
        break
    if graph_response.status_code not in (301, 302, 303, 307, 308):
        raise RuntimeError(
            "Graph could not provide the workbook (HTTP "
            f"{graph_response.status_code})."
        )

    signed_url = graph_response.headers.get("Location", "")
    parts = urlsplit(signed_url)
    if parts.scheme != "https" or not parts.hostname:
        raise RuntimeError("Graph returned an invalid download redirect")

    status = "network error"
    try:
        # A separate request prevents forwarding the Graph OAuth bearer token.
        result = requests.get(signed_url, timeout=90)
        status = str(result.status_code)
        if result.ok:
            downloaded = result.content
            break

        if (
            result.status_code in (401, 403)
            and parts.hostname.lower() == "my.microsoftpersonalcontent.com"
        ):
            temp_auth = parse_qs(parts.query).get("tempauth", [None])[0]
            if temp_auth:
                # Workaround for OneDrive Personal's 401 on download.aspx.
                result = requests.get(
                    signed_url,
                    headers={"Authorization": "Bearer " + temp_auth},
                    timeout=90,
                )
                status = str(result.status_code)
                if result.ok:
                    downloaded = result.content
                    break
    except requests.RequestException:
        # Requests exceptions can contain the secret query string.
        pass

    print(
        f"OneDrive download attempt {attempt}/3 failed "
        f"(HTTP {status}); requesting a fresh link."
    )

if downloaded is None:
    raise RuntimeError(
        "OneDrive returned unusable download links after three attempts. "
        "The workbook was not published; verify OneDrive permissions."
    )


# Verify it is a valid Excel workbook
with zipfile.ZipFile(io.BytesIO(downloaded)) as z:
    if "xl/workbook.xml" not in z.namelist():
        raise ValueError("Downloaded file is not a valid Excel workbook")


# Save with permanent filename
Path(LOCAL_FILENAME).write_bytes(downloaded)

print(
    "Workbook downloaded successfully:",
    LOCAL_FILENAME,
    len(downloaded),
    "bytes",
)
