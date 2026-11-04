"""Download all season-scoped SNL GameChanger CSV exports from OneDrive.

Expected OneDrive folder: SNL Stats
Expected file names: F26 Stats.csv, W27 Stats.csv, and later F/W seasons.

Reuses the existing SNL calendar workflow's MS_CLIENT_ID and MSAL_CACHE secrets.
Downloads into stats-input/ for subsequent processing, NOT public/.
Do not modify the existing calendar downloader.
"""

import csv
import io
import os
import re
from collections import deque
from pathlib import Path
from urllib.parse import quote

import msal
import requests

FOLDER_NAME = "SNL Stats"
FILENAME_PATTERN = re.compile(r"^[FW]\d{2} Stats\.csv$", re.IGNORECASE)
OUTPUT_DIR = Path("stats-input")
BASE = "https://graph.microsoft.com/v1.0/me/drive"


def authenticated_session():
    cache = msal.SerializableTokenCache()
    cache.deserialize(os.environ["MSAL_CACHE"])
    app = msal.PublicClientApplication(
        os.environ["MS_CLIENT_ID"],
        authority="https://login.microsoftonline.com/consumers",
        token_cache=cache,
    )
    accounts = app.get_accounts()
    if not accounts:
        raise RuntimeError("No Microsoft account found in token cache")
    result = app.acquire_token_silent_with_error(
        ["Files.Read"], account=accounts[0]
    )
    if not result or "access_token" not in result:
        raise RuntimeError(
            "Microsoft authentication failed: "
            + str((result or {}).get("error", "No token"))
        )
    session = requests.Session()
    session.headers.update({"Authorization": "Bearer " + result["access_token"]})
    return session


def list_items(session, url):
    """Yield all items, including paginated Microsoft Graph results."""
    while url:
        response = session.get(url, timeout=90)
        response.raise_for_status()
        body = response.json()
        yield from body.get("value", [])
        url = body.get("@odata.nextLink")


def locate_stats_folder(session):
    """Find exactly one folder called SNL Stats anywhere in OneDrive."""
    queue = deque([(BASE + "/root/children", "")])
    matches = []
    while queue:
        url, parent = queue.popleft()
        for item in list_items(session, url):
            if "folder" not in item:
                continue
            name = item.get("name", "")
            path = f"{parent}/{name}"
            if name.casefold() == FOLDER_NAME.casefold():
                matches.append((item["id"], path))
                # Still search remaining folders to detect ambiguous matches.
            else:
                folder_id = quote(item["id"], safe="")
                queue.append((f"{BASE}/items/{folder_id}/children", path))
    if len(matches) != 1:
        raise RuntimeError(
            f"Expected exactly one OneDrive folder named {FOLDER_NAME!r}; "
            f"found {len(matches)}."
        )
    return matches[0]


def validate_csv(data, filename):
    """Reject empty, malformed, or obviously wrong exports before deployment."""
    if not data:
        raise ValueError(f"{filename}: file is empty")
    try:
        decoded = data.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError(f"{filename}: file must be UTF-8 CSV") from exc
    try:
        rows = csv.DictReader(io.StringIO(decoded, newline=""))
        headers = set(rows.fieldnames or [])
        required = {"First Name", "Last Name", "Team", "Batting Average"}
        missing = required - headers
        if missing:
            raise ValueError(f"{filename}: not a GameChanger Leaders report; missing columns {sorted(missing)}")
        count = sum(1 for _ in rows)
    except csv.Error as exc:
        raise ValueError(f"{filename}: invalid CSV") from exc
    if count == 0:
        raise ValueError(f"{filename}: no player records")
    return count


def main():
    session = authenticated_session()
    folder_id, path = locate_stats_folder(session)
    print("Found OneDrive stats folder:", path)
    folder_url = f"{BASE}/items/{quote(folder_id, safe='')}/children"
    matches = []
    for item in list_items(session, folder_url):
        if "file" not in item:
            continue
        name = item.get("name", "")
        if FILENAME_PATTERN.fullmatch(name):
            matches.append(item)
    if not matches:
        raise RuntimeError(
            "No season CSV files found in SNL Stats. "
            "Expected a file such as 'F26 Stats.csv'."
        )

    names = [item["name"].casefold() for item in matches]
    if len(names) != len(set(names)):
        raise RuntimeError("Duplicate season filenames found in OneDrive")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for item in sorted(matches, key=lambda item: item["name"].casefold()):
        file_name = item["name"]
        file_id = quote(item["id"], safe="")
        response = session.get(f"{BASE}/items/{file_id}/content", timeout=90)
        response.raise_for_status()
        player_count = validate_csv(response.content, file_name)
        destination = OUTPUT_DIR / file_name
        destination.write_bytes(response.content)
        print(f"Downloaded {file_name}: {player_count} player records")
    print(f"Downloaded {len(matches)} SNL season CSV file(s) to {OUTPUT_DIR}/")
    print("Raw CSV files were NOT placed in the public website folder.")


if __name__ == "__main__":
    main()
