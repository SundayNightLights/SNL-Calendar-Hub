#!/usr/bin/env python3
"""Export public SNL schedule/results/standings from SNL Master.xlsx.

This is a separate read-only companion to generate_calendars.py. It does not
modify the workbook, calendar generation, or existing ICS event identifiers.

Usage:
    python scripts/generate_website_data.py 'SNL Master.xlsx' \
        --output public/site-data.json

Output intentionally contains no player rosters or personal information.
Standings are PROVISIONAL; final official seeds require the league's separate
seeding engine and any required manual draw.
"""

import argparse
import datetime as dt
import json
import pathlib
import re
from collections import defaultdict

# Reuse the same XLSX reader, local-time converter and team names as calendars.
from generate_calendars import TEAMS, game_date, read_sheet, times

TEAM_IDS = {
    "Buffaloes": "buffaloes",
    "IHT": "iht",
    "MORSL": "morsl",
    "Outcasts": "outcasts",
    "Pacific Maul": "maul",
    "Toronto Raccoons": "raccoons",
}
VENUE = "Markham Sports Dome"
SEASON = "Fall 2026–27"
SESSION = "Session One"


def string(record, col):
    return str(record.get(col, "") or "").strip()


def nonnegative_score(value, game_id):
    if value == "":
        return None
    if not re.fullmatch(r"\d+", value):
        raise ValueError(f"{game_id}: invalid score {value!r}")
    return int(value)


def integer(value, label, game_id):
    try:
        number = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{game_id}: invalid {label} {value!r}") from exc
    return number


def date_times(record, game_id):
    date = game_date(string(record, "E"))
    if date.weekday() != 6:
        raise ValueError(f"{game_id}: expected Sunday; found {date}")
    slot = integer(string(record, "G"), "slot", game_id)
    if slot not in (1, 2, 3):
        raise ValueError(f"{game_id}: invalid slot {slot}")
    value = string(record, "F")
    # The existing calendar generator uses slots for playoff times.
    is_regular = not game_id.startswith("F26-P")
    start, end = times(date, value if is_regular else None, slot)
    starts = dt.datetime.strptime(start, "%Y%m%dT%H%M%S")
    ends = dt.datetime.strptime(end, "%Y%m%dT%H%M%S")
    diamond_name = string(record, "H")
    match = re.match(r"^(?:Diamond\s+)?([12])\b", diamond_name, re.I)
    if not match:
        raise ValueError(f"{game_id}: invalid diamond {diamond_name!r}")
    diamond = int(match.group(1))
    return {
        "date": date.isoformat(),
        "time": value,
        "timeStart": starts.strftime("%H:%M"),
        "timeEnd": ends.strftime("%H:%M"),
        "startLabel": starts.strftime("%I:%M %p").lstrip("0"),
        "slot": slot,
        "diamond": diamond,
        "diamondLabel": "Near" if diamond == 1 else "Far",
    }


def fetch_games(sheet, kind):
    regex = r"F26-(\d{3})" if kind == "regular" else r"F26-P(\d{2})"
    expected = 60 if kind == "regular" else 15
    games = []
    seen = set()
    for record in sheet.values():
        game_id = string(record, "A")
        match = re.fullmatch(regex, game_id)
        if not match:
            continue
        number = int(match.group(1))
        if not 1 <= number <= expected or game_id in seen:
            raise ValueError(f"Unexpected or duplicate {kind} game: {game_id}")
        seen.add(game_id)
        week = integer(string(record, "B"), "week", game_id)
        away_name = string(record, "C")
        home_name = string(record, "D")
        time_info = date_times(record, game_id)
        away_raw = string(record, "I")
        home_raw = string(record, "J")
        game_3_optional = kind == "playoff" and number in (5, 6, 11, 12, 15)
        not_needed = kind == "playoff" and (
            "not necessary" in string(record, "M").lower()
            or "not necessary" in away_raw.lower()
            or "not necessary" in home_raw.lower()
        )
        if not_needed:
            away_score = home_score = None
            status = "Not Necessary"
        else:
            away_score = nonnegative_score(away_raw, game_id)
            home_score = nonnegative_score(home_raw, game_id)
            status = "Final" if away_score is not None and home_score is not None else "Unplayed"
            if (away_score is None) != (home_score is None):
                # A partially entered score is never published as a final result.
                away_score = home_score = None
        if kind == "regular":
            if away_name not in TEAM_IDS or home_name not in TEAM_IDS or away_name == home_name:
                raise ValueError(f"{game_id}: unrecognized regular-season teams")
            away, home = TEAM_IDS[away_name], TEAM_IDS[home_name]
        else:
            # Preserve "Seed 3" placeholders; never claim real playoff teams before seeding.
            away = TEAM_IDS.get(away_name, away_name)
            home = TEAM_IDS.get(home_name, home_name)
        game = {
            "id": game_id,
            "week": week,
            **time_info,
            "away": away,
            "home": home,
            "awayScore": away_score,
            "homeScore": home_score,
            "status": status,
        }
        if kind == "regular":
            game["awayPoints"] = (
                2 if away_score > home_score else 1 if away_score == home_score else 0
            ) if status == "Final" else None
            game["homePoints"] = (
                2 if home_score > away_score else 1 if away_score == home_score else 0
            ) if status == "Final" else None
            game["awayCappedRD"] = (
                integer(string(record, "N"), "away capped RD", game_id)
                if status == "Final" and string(record, "N") != "" else None
            )
            game["homeCappedRD"] = (
                integer(string(record, "O"), "home capped RD", game_id)
                if status == "Final" and string(record, "O") != "" else None
            )
        else:
            game["round"] = (
                "Quarterfinals" if number <= 6 else
                "Semifinals" if number <= 12 else "Finals"
            )
            game["ifNecessary"] = game_3_optional
            game["scoreEntryEligible"] = string(record, "T") == "1"
        games.append((number, game))
    if len(games) != expected:
        raise ValueError(f"Expected {expected} {kind} games; got {len(games)}")
    return [game for number, game in sorted(games)]



def apply_master_scores(games, master_rows):
    """Use actual yellow Master Schedule score inputs, not formula caches.

    Excel Online can save formulas in Regular Season Data without their cached
    values. The Master Schedule G/H cells are the authoritative score entries.
    Match by week and both team names, not a fragile worksheet row number.
    """
    scores = {}
    for row_number, record in master_rows.items():
        raw_week = string(record, "A")
        away_name = string(record, "E")
        home_name = string(record, "F")
        if not re.fullmatch(r"\d+", raw_week):
            continue
        week = int(raw_week)
        if not 1 <= week <= 10 or away_name not in TEAM_IDS or home_name not in TEAM_IDS:
            continue
        key = (week, TEAM_IDS[away_name], TEAM_IDS[home_name])
        if key in scores:
            raise ValueError(f"Duplicate Master Schedule matchup: {key}")
        label = f"Master Schedule row {row_number}"
        away_score = nonnegative_score(string(record, "G"), label)
        home_score = nonnegative_score(string(record, "H"), label)
        if (away_score is None) != (home_score is None):
            # Only publish a result when both scores have been entered.
            away_score = home_score = None
        scores[key] = (away_score, home_score)

    expected = {(g["week"], g["away"], g["home"]) for g in games}
    if len(expected) != len(games) or set(scores) != expected:
        missing = sorted(expected - set(scores))
        extra = sorted(set(scores) - expected)
        raise ValueError(
            f"Master Schedule does not match regular-season game data: "
            f"{len(missing)} missing, {len(extra)} unexpected; "
            f"sample missing: {missing[:3]}, sample unexpected: {extra[:3]}"
        )

    for game in games:
        key = (game["week"], game["away"], game["home"])
        away_score, home_score = scores[key]
        finished = away_score is not None and home_score is not None
        game["awayScore"] = away_score
        game["homeScore"] = home_score
        game["status"] = "Final" if finished else "Unplayed"
        if finished:
            game["awayPoints"] = 2 if away_score > home_score else 1 if away_score == home_score else 0
            game["homePoints"] = 2 if home_score > away_score else 1 if away_score == home_score else 0
            diff = max(-7, min(7, away_score - home_score))
            game["awayCappedRD"] = diff
            game["homeCappedRD"] = -diff
        else:
            game["awayPoints"] = None
            game["homePoints"] = None
            game["awayCappedRD"] = None
            game["homeCappedRD"] = None
    return games

def calculate_standings(games):
    rows = {
        team: {
            "team": team_id, "name": team, "gp": 0,
            "wins": 0, "losses": 0, "ties": 0, "points": 0,
            "runsFor": 0, "runsAgainst": 0,
            "runDiff": 0, "cappedRunDiff": 0,
        }
        for team, team_id in TEAM_IDS.items()
    }
    # Aggregate only fully entered games. Never depend on Excel formula caches.
    final_games = [g for g in games if g["status"] == "Final"]
    for g in final_games:
        away_name = next(name for name, ident in TEAM_IDS.items() if ident == g["away"])
        home_name = next(name for name, ident in TEAM_IDS.items() if ident == g["home"])
        a, h = rows[away_name], rows[home_name]
        a["gp"] += 1
        h["gp"] += 1
        a["runsFor"] += g["awayScore"]
        a["runsAgainst"] += g["homeScore"]
        h["runsFor"] += g["homeScore"]
        h["runsAgainst"] += g["awayScore"]
        a["points"] += g["awayPoints"]
        h["points"] += g["homePoints"]
        if g["awayScore"] > g["homeScore"]:
            a["wins"] += 1
            h["losses"] += 1
        elif g["awayScore"] < g["homeScore"]:
            h["wins"] += 1
            a["losses"] += 1
        else:
            a["ties"] += 1
            h["ties"] += 1
        if g["awayCappedRD"] is not None:
            a["cappedRunDiff"] += g["awayCappedRD"]
        if g["homeCappedRD"] is not None:
            h["cappedRunDiff"] += g["homeCappedRD"]
    for row in rows.values():
        row["runDiff"] = row["runsFor"] - row["runsAgainst"]
    # Preliminary order: points > wins > H2H pts > capped RD > fewest RA > most RF.
    # H2H points compare teams tied on points AND wins, as in the SNL rules.
    tie_groups = defaultdict(list)
    for name, row in rows.items():
        tie_groups[(row["points"], row["wins"])].append(name)
    h2h = dict.fromkeys(TEAMS, 0)
    for (_, _), group in tie_groups.items():
        if len(group) < 2:
            continue
        tied_ids = {TEAM_IDS[name] for name in group}
        for g in final_games:
            if g["away"] in tied_ids and g["home"] in tied_ids:
                away = next(name for name, ident in TEAM_IDS.items() if ident == g["away"])
                home = next(name for name, ident in TEAM_IDS.items() if ident == g["home"])
                h2h[away] += g["awayPoints"]
                h2h[home] += g["homePoints"]
    for name, row in rows.items():
        row["h2hPoints"] = h2h[name]
    ordered = sorted(rows.values(), key=lambda r: (
        -r["points"], -r["wins"], -r["h2hPoints"],
        -r["cappedRunDiff"], r["runsAgainst"], -r["runsFor"], r["name"]
    ))
    for index, row in enumerate(ordered, 1):
        row["rank"] = index
    return ordered


def export(workbook, output):
    regular = read_sheet(workbook, "Regular Season Data")
    playoffs = read_sheet(workbook, "Playoff Data")
    regular_games = fetch_games(regular, "regular")
    master = read_sheet(workbook, "Master Schedule")
    apply_master_scores(regular_games, master)
    playoff_games = fetch_games(playoffs, "playoff")
    finals = sum(g["status"] == "Final" for g in regular_games)
    payload = {
        "schemaVersion": 1,
        "generatedAt": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "source": pathlib.Path(workbook).name,
        "season": SEASON,
        "session": SESSION,
        "venue": VENUE,
        "timezone": "America/Toronto",
        "teams": [{"id": ident, "name": name} for name, ident in TEAM_IDS.items()],
        "regularSeason": regular_games,
        "standings": calculate_standings(regular_games),
        "standingsStatus": "provisional",
        "playoffs": playoff_games,
        "summary": {
            "regularGames": len(regular_games),
            "completedRegularGames": finals,
            "playoffGames": len(playoff_games),
        },
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return payload


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=pathlib.Path)
    parser.add_argument("--output", type=pathlib.Path, default=pathlib.Path("public/site-data.json"))
    args = parser.parse_args()
    result = export(args.workbook, args.output)
    print(f"Website data written to {args.output}")
    print(f"{result['summary']['completedRegularGames']}/60 regular-season games complete; "
          f"{len(result['standings'])} teams; 15 playoff slots")


if __name__ == "__main__":
    main()
