"""Prepare SNL batting-only website stats from GameChanger Leaders CSV exports.

OneDrive download_stats.py puts season files such as F26 Stats.csv and
W27 Stats.csv in stats-input/. This script reads ALL recognized season files,
so historical seasons stay available when the current season changes.

Default: write a local, player-name-free audit to stats-audit.json.
Publishing named player data requires all three explicit flags:
  --publish --confirm-export-scope --confirm-publication-rights

Do not upload raw GameChanger CSVs or stats-audit.json to a public site.
"""

import argparse
import csv
import hashlib
import json
import math
import re
import unicodedata
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

FILE_RE = re.compile(r"^([FW])(\d{2}) Stats\.csv$", re.IGNORECASE)

TEAMS = (
    ("buffaloes", "Buffaloes", ("Buffaloes",)),
    ("iht", "IHT", ("IHT",)),
    ("morsl", "MORSL", ("MORSL", "Dome ball - MORSL", "Dome Ball - MORSL")),
    ("outcasts", "Outcasts", ("Outcasts",)),
    ("pacific-maul", "Pacific Maul", ("Pacific Maul", "Pacific Mall")),
    ("toronto-raccoons", "Toronto Raccoons", ("Toronto Raccoons", "Raccoons")),
)

STATS = {
    "GP": "Games Played",
    "PA": "Plate Appearances",
    "AVG": "Batting Average",
    "QAB": "Quality At-Bats",
    "RBI": "Runs Batted In",
    "H": "Hits",
    "OBP": "On-base Percentage",
    "SLG": "Slugging Percentage",
    "OPS": "On-base Percentage + Slugging Percentage",
    "2B": "Doubles",
    "3B": "Triples",
    "HR": "Home Runs",
}
COUNTS = {"GP", "PA", "QAB", "RBI", "H", "2B", "3B", "HR"}
BATTING_STATS = set(STATS) - {"GP"}
REQUIRED = {"First Name", "Last Name", "Team", "Jersey Number", *STATS.values()}
NOTICE = "Batting statistics are voluntarily submitted to GameChanger by participating teams; non-reporting teams are not included in leaderboards."


def norm(text):
    text = unicodedata.normalize("NFKD", str(text or ""))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(text.strip().casefold().split())


TEAM_LOOKUP = {}
for team_id, display_name, aliases in TEAMS:
    for alias in aliases:
        TEAM_LOOKUP[norm(alias)] = (team_id, display_name)


def season_info(path):
    match = FILE_RE.fullmatch(path.name)
    if not match:
        raise ValueError(f"Invalid stats filename: {path.name}")
    initial, year = match.group(1).upper(), int(match.group(2))
    # Chronological order: F26, W27, F27, W28, etc.
    return {
        "id": f"{initial}{year:02d}",
        "label": f"{'Fall' if initial == 'F' else 'Winter'} {2000 + year}",
        "sort_key": (year, 0 if initial == "W" else 1),
    }


def parse_stat(value, code, source):
    value = str(value or "").strip()
    if not value:
        return None
    if code in COUNTS:
        if not re.fullmatch(r"\d+", value):
            raise ValueError(f"{source}: invalid {code} count {value!r}")
        return int(value)
    try:
        number = float(value)
    except ValueError as exc:
        raise ValueError(f"{source}: invalid {code} rate {value!r}") from exc
    if not math.isfinite(number) or number < 0:
        raise ValueError(f"{source}: invalid {code} rate {value!r}")
    if code in ("AVG", "OBP") and number > 1:
        raise ValueError(f"{source}: {code} outside 0–1 range")
    return number


def read_session(path):
    details = season_info(path)
    players = []
    reported = Counter()
    encountered = set()
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = set(reader.fieldnames or [])
        missing = REQUIRED - headers
        if missing:
            raise ValueError(f"{path.name}: missing GameChanger fields {sorted(missing)}")
        for n, row in enumerate(reader, start=2):
            if None in row:
                raise ValueError(f"{path.name} row {n}: excess CSV columns")
            name = " ".join(filter(None, [row["First Name"].strip(), row["Last Name"].strip()]))
            if not name:
                raise ValueError(f"{path.name} row {n}: blank player name")
            raw_team = row["Team"].strip()
            if norm(raw_team) not in TEAM_LOOKUP:
                raise ValueError(f"{path.name} row {n}: unknown team {raw_team!r}")
            team_id, team_name = TEAM_LOOKUP[norm(raw_team)]
            identity = (team_id, norm(name))
            if identity in encountered:
                raise ValueError(f"{path.name}: duplicate name on team: {name}")
            encountered.add(identity)
            values = {}
            for code, header in STATS.items():
                value = parse_stat(row[header], code, f"{path.name} row {n}")
                if value is not None:
                    values[code] = value
            if BATTING_STATS.intersection(values):
                reported[team_id] += 1
            # A per-session identity; DO NOT infer that same-name players across
            # sessions are the same person. Career aggregation needs ID review.
            provisional_id = hashlib.sha256((details["id"] + "|" + team_id + "|" + norm(name)).encode()).hexdigest()[:14]
            players.append({
                "player_id": "session-" + provisional_id,
                "name": name,
                "team_id": team_id,
                "team": team_name,
                "jersey_number": row["Jersey Number"].strip() or None,
                "stats": values,
            })
    if not players:
        raise ValueError(f"{path.name}: no player rows")
    teams = [
        {
            "team_id": team_id,
            "name": display_name,
            "reporting_status": "stats_available" if reported[team_id] else "not_submitted",
            "players_with_batting_stats": reported[team_id],
        }
        for team_id, display_name, _ in TEAMS
    ]
    return {
        "season_id": details["id"],
        "season_label": details["label"],
        "teams": teams,
        "player_count": len(players),
        "reporting_team_count": sum(bool(reported[team_id]) for team_id, _, _ in TEAMS),
        "players": sorted(players, key=lambda x: (x["team"], x["name"])),
        "_sort_key": details["sort_key"],
    }


def save_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp.replace(path)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input-dir", type=Path, default=Path("stats-input"))
    p.add_argument("--audit-output", type=Path, default=Path("stats-audit.json"))
    p.add_argument("--output", type=Path, default=Path("public/stats-data.json"))
    p.add_argument("--publish", action="store_true")
    p.add_argument("--confirm-export-scope", action="store_true")
    p.add_argument("--confirm-publication-rights", action="store_true")
    a = p.parse_args()
    if a.publish and not (a.confirm_export_scope and a.confirm_publication_rights):
        p.error("Publishing player names requires confirmation of report scope AND publication rights")
    inputs = sorted((x for x in a.input_dir.iterdir() if x.is_file() and FILE_RE.fullmatch(x.name)), key=lambda x: x.name.casefold()) if a.input_dir.is_dir() else []
    if not inputs:
        p.error(f"No F26 Stats.csv / W27 Stats.csv-style files found in {a.input_dir}")
    seasons = [read_session(f) for f in inputs]
    ids = [s["season_id"] for s in seasons]
    if len(set(ids)) != len(ids):
        p.error("Duplicate season identifiers (for example F26 Stats.csv and f26 stats.csv)")
    seasons.sort(key=lambda s: s.pop("_sort_key"))
    generated = datetime.now(timezone.utc).isoformat()
    summary = {
        "generated_at_utc": generated,
        "current_season": seasons[-1]["season_id"],
        "notice": NOTICE,
        "seasons": [
            {
                "season_id": s["season_id"],
                "season_label": s["season_label"],
                "player_count": s["player_count"],
                "reporting_team_count": s["reporting_team_count"],
                "nonreporting_teams": [t["name"] for t in s["teams"] if t["reporting_status"] == "not_submitted"],
            }
            for s in seasons
        ],
        "limitations": "Seasonal rates cannot simply be averaged into career rates. Cross-season player identity must be reviewed before career totals are shown.",
    }
    save_json(a.audit_output, summary)
    print(f"Validated {len(seasons)} season file(s): " + ", ".join(ids))
    for s in seasons:
        print(f"  {s['season_id']}: {s['player_count']} players; {s['reporting_team_count']}/6 teams reporting")
    print(f"Audit saved locally to {a.audit_output} (no player names)")
    if not a.publish:
        print("PREVIEW ONLY: no player statistics were published")
        return
    output = {"generated_at_utc": generated, "current_season": seasons[-1]["season_id"], "notice": NOTICE, "seasons": seasons}
    save_json(a.output, output)
    print(f"Website data generated: {a.output}")


if __name__ == "__main__":
    main()
