
"""Generate season-neutral SNL team calendars from SNL Master.xlsx."""

import argparse
import datetime as dt
import json
import pathlib
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile

from standings_engine import official_seeds
from playoff_engine import bracket, SERIES


TEAMS = (
    "Buffaloes",
    "IHT",
    "MORSL",
    "Outcasts",
    "Pacific Maul",
    "Toronto Raccoons",
)

PREFIX = {
    team: team.replace(" ", "_")
    for team in TEAMS
}

TZ = "America/Toronto"
VENUE = "Markham Sports Dome"
ADDRESS = "5300 14th Avenue, Markham, ON L3S 3K8"

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
}

PLAYOFF_ROWS = (
    71, 72, 73, 74, 75, 76,
    78, 79, 80, 81, 82, 83,
    85, 86, 87,
)


def read_sheet(path, name):
    """Read saved Excel values by worksheet name."""

    with zipfile.ZipFile(path) as z:
        shared = []

        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(
                z.read("xl/sharedStrings.xml")
            )

            shared = [
                "".join(
                    t.text or ""
                    for t in item.findall(".//m:t", NS)
                )
                for item in root
            ]

        book = ET.fromstring(
            z.read("xl/workbook.xml")
        )

        sheets = book.findall(
            ".//m:sheets/m:sheet", NS
        )

        matches = [
            s for s in sheets
            if s.get("name") == name
        ]

        if len(matches) != 1:
            raise ValueError(
                f"Worksheet not found: {name}"
            )

        rel_key = (
            "{http://schemas.openxmlformats.org/"
            "officeDocument/2006/relationships}id"
        )

        rel_id = matches[0].get(rel_key)

        relationships = ET.fromstring(
            z.read("xl/_rels/workbook.xml.rels")
        )

        target = next(
            r.get("Target")
            for r in relationships
            if r.get("Id") == rel_id
        )

        if target.startswith("/"):
            location = target.lstrip("/")
        else:
            location = posixpath.normpath(
                posixpath.join("xl", target)
            )

        root = ET.fromstring(
            z.read(location)
        )

        result = {}

        for row in root.findall(
            ".//m:sheetData/m:row", NS
        ):
            values = {}

            for cell in row.findall("m:c", NS):
                col = re.match(
                    r"[A-Z]+", cell.get("r")
                ).group()

                val = cell.find("m:v", NS)
                inline = cell.find("m:is", NS)

                if val is not None:
                    data = val.text or ""

                    if cell.get("t") == "s":
                        data = shared[int(data)]

                elif inline is not None:
                    data = "".join(
                        t.text or ""
                        for t in inline.findall(
                            ".//m:t", NS
                        )
                    )
                else:
                    data = ""

                values[col] = data

            result[int(row.get("r"))] = values


        return normalize_game_columns(result, name)


def normalize_game_columns(rows, name):
    """Accept both SNL game-data column layouts without changing callers.

    Calendar generation and website export use the original A:J positions
    internally. Normalize the rearranged workbook at the input boundary so
    existing event UIDs, playoff logic, and website calculations stay intact.
    """
    if name not in ("Regular Season Data", "Playoff Data"):
        return rows

    headers = {key: str(value).strip() for key, value in rows.get(4, {}).items()}
    if (
        headers.get("C") == "Away Team"
        and headers.get("D") == "Home Team"
        and headers.get("E") == "Date"
        and headers.get("I") == "Away Score"
    ):
        return rows

    if not (
        headers.get("C") == "Date"
        and headers.get("D") == "Time"
        and headers.get("E") == "Slot"
        and headers.get("F") == "Diamond"
        and headers.get("G") == "Away Team"
        and headers.get("H") == "Home Team"
        and headers.get("I") == "Game Status"
        and headers.get("J") == "Away Score"
        and headers.get("K") == "Home Score"
    ):
        raise ValueError(f"{name}: unrecognized game-data column layout")

    # Left: legacy/canonical position; right: column in the new workbook.
    legacy_to_new = {
        "A": "A", "B": "B", "C": "G", "D": "H", "E": "C",
        "F": "D", "G": "E", "H": "F", "I": "J", "J": "K",
        "K": "L", "L": "M", "M": "I",
    }

    return {
        row_number: {
            **record,
            **{
                legacy: record.get(current, "")
                for legacy, current in legacy_to_new.items()
            },
        }
        for row_number, record in rows.items()
    }


def legacy_rows(regular, playoffs):
    """Preserve historical game identities and event UIDs."""

    headers = {
        "A": "Game ID",
        "B": "Week",
        "C": "Away Team",
        "D": "Home Team",
        "E": "Date",
        "F": "Time",
        "G": "Slot",
        "H": "Diamond",
        "I": "Away Score",
        "J": "Home Score",
    }

    for name, sheet in (
        ("Regular Season Data", regular),
        ("Playoff Data", playoffs),
    ):
        for col, label in headers.items():
            actual = str(
                sheet.get(4, {}).get(col, "")
            ).strip()

            if actual != label:
                raise ValueError(
                    f"{name} column {col}: "
                    f"expected {label!r}, "
                    f"found {actual!r}"
                )

    def convert(record):
        return {
            "B": record.get("E", ""),
            "C": record.get("F", ""),
            "D": record.get("G", ""),
            "E": record.get("H", ""),
            "F": record.get("C", ""),
            "G": record.get("D", ""),
            "H": record.get("I", ""),
            "I": record.get("J", ""),
        }

    rows = {}
    regular_ids = set()

    for record in regular.values():
        game_id = str(
            record.get("A", "")
        ).strip()

        match = re.fullmatch(
            r"F26-(\d{3})", game_id
        )

        if not match:
            continue

        number = int(match.group(1))

        if (
            not 1 <= number <= 60
            or game_id in regular_ids
        ):
            raise ValueError(
                f"Invalid or duplicate Game ID: {game_id}"
            )

        regular_ids.add(game_id)

        # Preserve the original game identities.
        # The old holiday row followed game 48.
        original_row = (
            number + 4 + (number >= 49)
        )

        rows[original_row] = convert(record)

        away = rows[original_row]["F"]
        home = rows[original_row]["G"]

        if (
            away not in TEAMS
            or home not in TEAMS
            or away == home
        ):
            raise ValueError(
                f"{game_id}: invalid matchup "
                f"{away!r} at {home!r}"
            )

    expected = {
        f"F26-{n:03d}"
        for n in range(1, 61)
    }

    if regular_ids != expected:
        raise ValueError(
            "Missing regular-season Game IDs: "
            + str(sorted(expected - regular_ids))
        )

    playoff_ids = set()

    for record in playoffs.values():
        game_id = str(
            record.get("A", "")
        ).strip()

        match = re.fullmatch(
            r"F26-P(\d{2})", game_id
        )

        if not match:
            continue

        number = int(match.group(1))

        if (
            not 1 <= number <= 15
            or game_id in playoff_ids
        ):
            raise ValueError(
                f"Invalid playoff Game ID: {game_id}"
            )

        playoff_ids.add(game_id)

        original_row = PLAYOFF_ROWS[number - 1]

        rows[original_row] = convert(record)

    expected = {
        f"F26-P{n:02d}"
        for n in range(1, 16)
    }

    if playoff_ids != expected:
        raise ValueError(
            "Missing playoff Game IDs: "
            + str(sorted(expected - playoff_ids))
        )

    return rows


def read_draw(sheet):
    """Read manually entered final tie-break orders."""

    draw = {}

    for row in range(4, 10):
        team = str(
            sheet.get(row, {}).get("A", "")
        ).strip()

        if team not in TEAMS:
            raise ValueError(
                f"Invalid Playoff Seeding team "
                f"on row {row}: {team!r}"
            )

        value = str(
            sheet[row].get("K", "")
        ).strip()

        if value:
            draw[team] = (
                int(value)
                if value.isdecimal()
                else None
            )

    return draw


def game_date(value):
    return (
        dt.date(1899, 12, 30)
        + dt.timedelta(days=int(float(value)))
    )


def escape(value):
    return (
        str(value)
        .replace("\\", "\\\\")
        .replace(";", "\\;")
        .replace(",", "\\,")
        .replace("\n", "\\n")
    )


def fold(line):
    """Fold iCalendar lines safely."""

    chunks = []
    buf = ""
    size = 0

    for char in line:
        width = len(char.encode("utf-8"))

        if size + width > (
            75 if not chunks else 74
        ):
            chunks.append(buf)
            buf = " "
            size = 1

        buf += char
        size += width

    return "\r\n".join(chunks + [buf])


def times(date, time_range, slot):
    """Convert workbook times to calendar times."""

    if (
        time_range
        and str(time_range).strip()
        .lower().startswith("game 3")
    ):
        time_range = None

    if time_range:
        pattern = (
            r"^(\d{1,2}:\d{2})"
            r"(?:\s*(AM|PM))?"
            r"\s*[-–—]\s*"
            r"(\d{1,2}:\d{2})"
            r"\s*(AM|PM)$"
        )

        match = re.fullmatch(
            pattern,
            str(time_range).strip(),
            re.I,
        )

        if not match:
            raise ValueError(
                f"Unexpected game time: "
                f"{time_range!r}"
            )

        a, ap, b, bp = match.groups()
        ap = ap or bp

        first = dt.datetime.strptime(
            f"{a} {ap.upper()}",
            "%I:%M %p",
        ).time()

        last = dt.datetime.strptime(
            f"{b} {bp.upper()}",
            "%I:%M %p",
        ).time()

        start = dt.datetime.combine(
            date, first
        )

        end = dt.datetime.combine(
            date, last
        )

        if end <= start:
            end += dt.timedelta(days=1)

    else:
        hour, minute = {
            1: (22, 0),
            2: (22, 45),
            3: (23, 30),
        }[slot]

        start = dt.datetime.combine(
            date,
            dt.time(hour, minute),
        )

        end = start + dt.timedelta(
            minutes=45
        )

    return (
        start.strftime("%Y%m%dT%H%M%S"),
        end.strftime("%Y%m%dT%H%M%S"),
    )


def event(
    uid,
    row,
    team,
    opponent,
    side,
    phase,
    tentative=False,
    regular=False,
):
    """Generate one iCalendar event."""

    date = game_date(row["B"])
    slot = int(float(row["D"]))

    diamond_match = re.search(
        r"[12]",
        str(row["E"]),
    )

    if not diamond_match:
        raise ValueError(
            f"Invalid diamond for {team}: "
            f"{row['E']!r}"
        )

    diamond = diamond_match.group()

    field = (
        "Near Field"
        if diamond == "1"
        else "Far Field"
    )

    start, end = times(
        date,
        row["C"] if regular else None,
        slot,
    )

    name = (
        f"{team} vs {opponent} "
        f"({side} | Diamond {diamond})"
    )

    if tentative:
        name += " (If Necessary)"

    description = (
        f"Sunday Night Lights — {phase}\n"
        f"Diamond {diamond} — {field}\n"
        f"Venue: {VENUE}, {ADDRESS}\n"
        f"Schedule subject to change."
    )

    if tentative:
        description += (
            "\nPotential Game 3; played only "
            "if the series is tied after two games."
        )

    location = (
        f"{VENUE} | "
        f"Diamond {diamond} ({field})"
    )

    fields = [
        "BEGIN:VEVENT",
        f"UID:{uid}",
        "DTSTAMP:20261029T000000Z",
        f"DTSTART;TZID={TZ}:{start}",
        f"DTEND;TZID={TZ}:{end}",
        f"SUMMARY:{escape(name)}",
        f"LOCATION:{escape(location)}",
        (
            f"X-SNL-DIAMOND:"
            f"Diamond {diamond} — {field}"
        ),
        f"DESCRIPTION:{escape(description)}",
        (
            "STATUS:TENTATIVE"
            if tentative
            else "STATUS:CONFIRMED"
        ),
        "END:VEVENT",
    ]

    return "\r\n".join(
        fold(line) for line in fields
    )


def reference_calendar(folder, team):
    """Find existing reference calendars by team.

    Supports both the new short filenames and
    previously published reference filenames.
    """

    current = (
        folder / f"{PREFIX[team]}.ics"
    )

    if current.exists():
        return current

    pattern = re.compile(
        r"(?:^|_)"
        + re.escape(PREFIX[team])
        + r"(?:_|$)",
        re.I,
    )

    matches = [
        path
        for path in folder.glob("*.ics")
        if pattern.search(path.stem)
    ]

    if len(matches) > 1:
        raise ValueError(
            f"More than one reference "
            f"calendar for {team}: {matches}"
        )

    return matches[0] if matches else None


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "workbook",
        type=pathlib.Path,
    )

    parser.add_argument(
        "--output",
        type=pathlib.Path,
        default=pathlib.Path("public"),
    )

    parser.add_argument(
        "--reference",
        type=pathlib.Path,
        default=pathlib.Path("reference"),
    )

    args = parser.parse_args()

    regular = read_sheet(
        args.workbook,
        "Regular Season Data",
    )

    playoffs = read_sheet(
        args.workbook,
        "Playoff Data",
    )

    seeding = read_sheet(
        args.workbook,
        "Playoff Seeding",
    )

    rows = legacy_rows(
        regular,
        playoffs,
    )

    # Preserve the original published game UIDs.
    uid_path = (
        args.reference / "game_uids.json"
    )

    if not uid_path.exists():
        raise FileNotFoundError(
            f"Missing original UID registry: "
            f"{uid_path}"
        )

    original_uids = {
        int(k): v
        for k, v in json.loads(
            uid_path.read_text()
        ).items()
    }

    if (
        len(original_uids) != 60
        or len(set(original_uids.values())) != 60
    ):
        raise ValueError(
            "Original UID registry must "
            "contain 60 unique games"
        )

    games = {
        team: []
        for team in TEAMS
    }

    # Generate all 60 regular-season games.
    for row_num in (
        list(range(5, 53))
        + list(range(54, 66))
    ):
        row = rows[row_num]

        away = row["F"]
        home = row["G"]
        uid = original_uids[row_num]

        games[away].append(
            event(
                uid,
                row,
                away,
                home,
                "Away",
                "Regular Season",
                regular=True,
            )
        )

        games[home].append(
            event(
                uid,
                row,
                home,
                away,
                "Home",
                "Regular Season",
                regular=True,
            )
        )

    # Calculate playoff seeds from the scores.
    seeds = official_seeds(
        rows,
        read_draw(seeding),
    )

    if seeds is None:
        print(
            "Playoff seeding pending; "
            "using neutral placeholders."
        )

    scores = {}

    for row_num in PLAYOFF_ROWS:
        row = rows[row_num]

        away = str(
            row["H"]
        ).strip()

        home = str(
            row["I"]
        ).strip()

        if away or home:
            if (
                not away.isdecimal()
                or not home.isdecimal()
            ):
                raise ValueError(
                    f"Invalid playoff score "
                    f"for game row {row_num}"
                )

            scores[row_num] = (
                int(away),
                int(home),
            )

    rounds = bracket(
        seeds,
        scores,
    )

    for series, result in rounds.items():
        phase = {
            "QF": "Quarterfinals",
            "SF": "Semifinals",
            "FINAL": "Finals",
        }[series.split("_")[0]]

        for index, row_num in enumerate(
            SERIES[series]
        ):
            if (
                index == 2
                and result.game3 == "not_needed"
            ):
                continue

            row = rows[row_num]

            if index == 1:
                home = result.lower
                away = result.higher
            else:
                home = result.higher
                away = result.lower

            uid = (
                f"snl-playoff-row-"
                f"{row_num - 1}"
                f"@sundaynightlights"
            )

            tentative = (
                index == 2
                and result.game3 == "potential"
            )

            games[home].append(
                event(
                    uid,
                    row,
                    home,
                    away,
                    "Home",
                    phase,
                    tentative,
                )
            )

            games[away].append(
                event(
                    uid,
                    row,
                    away,
                    home,
                    "Away",
                    phase,
                    tentative,
                )
            )

    # Preserve playoff placeholders from
    # the existing reference calendar files.
    for team in TEAMS:
        prior = reference_calendar(
            args.reference,
            team,
        )

        if prior is None:
            continue

        original = prior.read_text(
            encoding="utf-8-sig"
        )

        for chunk in original.split(
            "BEGIN:VEVENT"
        )[1:]:
            block = (
                "BEGIN:VEVENT"
                + chunk.split(
                    "END:VEVENT"
                )[0]
                + "END:VEVENT"
            )

            if "(Placeholder)" not in block:
                continue

            phase = next(
                (
                    p for p in (
                        "Quarterfinals",
                        "Semifinals",
                        "Finals",
                    )
                    if (
                        f"SUMMARY:{p} "
                        f"(Placeholder)"
                    ) in block
                ),
                None,
            )

            if phase is None:
                continue

            if any(
                f"Sunday Night Lights — {phase}"
                in existing
                for existing in games[team]
            ):
                continue

            games[team].append(
                block.replace(
                    "\r\n", "\n"
                ).replace(
                    "\n", "\r\n"
                )
            )

    # Write six season-neutral calendar files.
    args.output.mkdir(
        parents=True,
        exist_ok=True,
    )

    for team, entries in games.items():

        # Example: Buffaloes.ics
        filename = f"{PREFIX[team]}.ics"

        # Calendar displays just the team name.
        header = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            (
                "PRODID:-//Sunday Night Lights//"
                "Team Calendar//EN"
            ),
            "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH",
            f"NAME:{escape(team)}",
            f"X-WR-CALNAME:{escape(team)}",
            f"X-WR-TIMEZONE:{TZ}",
        ]

        content = "\r\n".join(
            header
            + entries
            + ["END:VCALENDAR", ""]
        )

        (
            args.output / filename
        ).write_bytes(
            content.encode("utf-8")
        )

        print(
            f"{filename}: "
            f"{len(entries)} events"
        )

    # Create a basic calendar index.
    links = "\n".join(
        (
            f'<li><a href="{PREFIX[team]}.ics">'
            f'{team}</a></li>'
        )
        for team in TEAMS
    )

    index = (
        '<!doctype html>'
        '<html lang="en">'
        '<meta charset="utf-8">'
        '<title>Sunday Night Lights Calendars</title>'
        '<h1>Sunday Night Lights — Team Calendars</h1>'
        '<ul>'
        + links
        + '</ul></html>'
    )

    (
        args.output / "index.html"
    ).write_text(
        index,
        encoding="utf-8",
    )

    print(
        "Finished generating six permanent "
        "SNL team calendars."
    )


if __name__ == "__main__":
    main()
