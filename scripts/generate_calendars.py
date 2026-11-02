
"""Generate SNL calendars from Regular Season Data and Playoff Data."""
import argparse
import datetime as dt
import hashlib
import json
import pathlib
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile

from standings_engine import official_seeds
from playoff_engine import bracket, SERIES

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

TEAMS = [
    'Buffaloes', 'IHT', 'MORSL', 'Outcasts',
    'Pacific Maul', 'Toronto Raccoons'
]

VENUE = 'Markham Sports Dome'
VENUE_ADDRESS = (
    'Markham Sports Dome, 5300 14th Avenue, '
    'Markham, ON L3S 3K8'
)

PREFIX = {
    'Buffaloes': 'Buffaloes',
    'IHT': 'IHT',
    'MORSL': 'MORSL',
    'Outcasts': 'Outcasts',
    'Pacific Maul': 'Pacific_Maul',
    'Toronto Raccoons': 'Toronto_Raccoons'
}

TZ = 'America/Toronto'


def cells_from_xlsx(path, sheet_name):
    """Read a worksheet by NAME without evaluating Excel formulas."""
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            shared = [
                ''.join(t.text or '' for t in item.findall('.//m:t', NS))
                for item in root
            ]

        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rel_ns = (
            '{http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships}'
        )

        sheets = wb.findall('.//m:sheets/m:sheet', NS)
        matching = [
            sheet for sheet in sheets
            if sheet.get('name') == sheet_name
        ]

        if len(matching) != 1:
            raise ValueError(
                f'Missing or duplicate worksheet {sheet_name!r}'
            )

        rel_id = matching[0].get(rel_ns + 'id')
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        targets = {
            rel.get('Id'): rel.get('Target')
            for rel in rels
        }
        target = targets[rel_id]

        sheet_path = (
            target.lstrip('/')
            if target.startswith('/')
            else posixpath.normpath(
                posixpath.join('xl', target)
            )
        )

        root = ET.fromstring(z.read(sheet_path))
        rows = {}

        for row in root.findall('.//m:sheetData/m:row', NS):
            result = {}

            for cell in row.findall('m:c', NS):
                col = re.match(
                    r'[A-Z]+', cell.get('r')
                ).group()

                value = cell.find('m:v', NS)
                inline = cell.find('m:is', NS)

                if value is not None:
                    v = value.text or ''
                    if cell.get('t') == 's':
                        v = shared[int(v)]
                elif inline is not None:
                    v = ''.join(
                        t.text or ''
                        for t in inline.findall('.//m:t', NS)
                    )
                else:
                    v = ''

                if cell.find('m:f', NS) is not None:
                    result[col + '_formula'] = True

                result[col] = v

            rows[int(row.get('r'))] = result

        return rows


def make_legacy_rows(regular, playoff):
    """Map the new data sheets to historical game identifiers.

    The existing calendar UIDs use original Master Schedule rows.
    These mappings preserve those identifiers while reading
    game information from the two new source worksheets.

    The holiday break was on original Master row 53.
    """
    expected = {
        'B': 'Away Team',
        'C': 'Home Team',
        'D': 'Date',
        'E': 'Time',
        'F': 'Slot',
        'G': 'Diamond',
        'H': 'Away Score',
        'I': 'Home Score'
    }

    for title, data in [
        ('Regular Season Data', regular),
        ('Playoff Data', playoff)
    ]:
        header = data.get(4, {})

        for col, name in expected.items():
            if str(header.get(col, '')).strip() != name:
                raise ValueError(
                    f'{title}: column {col} must be {name!r}; '
                    f'got {header.get(col)!r}'
                )

    def convert(record, source, row_number):
        result = {
            'A': record.get('A', ''),
            'B': record.get('D', ''),
            'C': record.get('E', ''),
            'D': record.get('F', ''),
            'E': record.get('G', ''),
            'F': record.get('B', ''),
            'G': record.get('C', ''),
            'H': record.get('H', ''),
            'I': record.get('I', '')
        }

        if not result['B'] or not result['C'] or not result['D']:
            raise ValueError(
                f'{source} row {row_number}: '
                'missing date, time or slot'
            )

        return result

    legacy = {}

    # 60 regular-season games
    for source_row in range(5, 65):
        record = regular.get(source_row, {})

        legacy_row = (
            source_row
            if source_row <= 52
            else source_row + 1
        )

        legacy[legacy_row] = convert(
            record, 'Regular Season Data', source_row
        )

        if (
            legacy[legacy_row]['F'] not in TEAMS
            or legacy[legacy_row]['G'] not in TEAMS
        ):
            raise ValueError(
                f'Invalid regular-season team '
                f'on source row {source_row}'
            )

    playoff_master_rows = [
        71, 72, 73, 74, 75, 76,
        78, 79, 80, 81, 82, 83,
        85, 86, 87
    ]

    # 15 potential playoff games
    for source_row, legacy_row in enumerate(
        playoff_master_rows, start=5
    ):
        legacy[legacy_row] = convert(
            playoff.get(source_row, {}),
            'Playoff Data',
            source_row
        )

    if len([
        row for row in legacy
        if 5 <= row <= 65
    ]) != 60:
        raise ValueError(
            'Regular Season Data must contain exactly 60 games'
        )

    return legacy


def read_draw_order(seeding):
    """Read manually selected random draws from Playoff Seeding."""
    draw = {}
    seen = set()

    for row_number in range(4, 10):
        record = seeding.get(row_number, {})
        team = str(record.get('A', '')).strip()

        if team not in TEAMS or team in seen:
            raise ValueError(
                f'Invalid Playoff Seeding team '
                f'at row {row_number}: {team!r}'
            )

        seen.add(team)
        value = str(record.get('K', '')).strip()

        if value:
            draw[team] = (
                int(value) if value.isdecimal() else None
            )

    if seen != set(TEAMS):
        raise ValueError(
            'Playoff Seeding must list all six SNL teams'
        )

    return draw


def date_from_serial(serial):
    return (
        dt.date(1899, 12, 30)
        + dt.timedelta(days=int(float(serial)))
    )


def ics_escape(s):
    return (
        str(s)
        .replace('\\', '\\\\')
        .replace(';', '\\;')
        .replace(',', '\\,')
        .replace('\n', '\\n')
    )


def fold(line):
    """Keep iCalendar lines within the UTF-8 length limit."""
    chunks, buf, size = [], '', 0

    for char in line:
        n = len(char.encode('utf-8'))

        if size + n > (75 if not chunks else 74):
            chunks.append(buf)
            buf, size = ' ', 1

        buf += char
        size += n

    chunks.append(buf)
    return '\r\n'.join(chunks)


def uid_for(row, date, slot, diamond):
    return (
        hashlib.sha1(
            f'SNL|master-row|{row}'.encode()
        ).hexdigest()[:20]
        + '@sundaynightlights'
    )


def read_reference_uids(folder):
    index = {}

    if not folder.exists():
        return index

    for path in folder.glob('*.ics'):
        content = (
            path.read_text(encoding='utf-8-sig')
            .replace('\r\n ', '')
        )

        for part in content.split('BEGIN:VEVENT')[1:]:
            fields = {}

            for line in part.splitlines():
                if ':' in line:
                    k, v = line.split(':', 1)
                    fields[k] = v

            start = fields.get(
                'DTSTART;TZID=America/Toronto'
            )
            summary = fields.get('SUMMARY', '')
            diamond = re.search(
                r'Diamond (\d)', summary
            )

            if start and diamond and fields.get('UID'):
                index[
                    (start, diamond.group(1))
                ] = fields['UID']

    return index


def parse_game_times(date, time_range):
    pattern = (
        r'^(\d{1,2}:\d{2})'
        r'(?:\s*(AM|PM))?'
        r'\s*[-–—]\s*'
        r'(\d{1,2}:\d{2})'
        r'\s*(AM|PM)$'
    )

    match = re.match(
        pattern, str(time_range).strip(), re.I
    )

    if not match:
        raise ValueError(
            f'Invalid game time range: {time_range!r}'
        )

    first, first_period, last, last_period = match.groups()
    first_period = first_period or last_period

    def parse(clock, period):
        return dt.datetime.strptime(
            clock + ' ' + period.upper(),
            '%I:%M %p'
        ).time()

    start = dt.datetime.combine(
        date, parse(first, first_period)
    )
    end = dt.datetime.combine(
        date, parse(last, last_period)
    )

    if end <= start:
        end += dt.timedelta(days=1)

    return start, end


def emit_event(
    uid, date, slot, diamond,
    team, opponent, side, phase,
    tentative=False, time_range=None
):
    if time_range is not None:
        start, end = parse_game_times(
            date, time_range
        )
    else:
        start_hours = {
            1: (22, 0),
            2: (22, 45),
            3: (23, 30)
        }

        hour, minute = start_hours[slot]
        start = dt.datetime.combine(
            date, dt.time(hour, minute)
        )
        end = start + dt.timedelta(minutes=45)

    start_s = start.strftime('%Y%m%dT%H%M%S')
    end_s = end.strftime('%Y%m%dT%H%M%S')

    field = (
        'Near Field'
        if diamond == '1'
        else 'Far Field'
    )

    location = (
        f'{VENUE} | Diamond {diamond} ({field})'
    )

    suffix = ' (If Necessary)' if tentative else ''
    title = (
        f'{team} vs {opponent} '
        f'({side} | Diamond {diamond}){suffix}'
    )

    desc = (
        f'Sunday Night Lights — {phase}'
        f'\\nDiamond {diamond} — {field}'
        f'\\nVenue: {VENUE_ADDRESS}'
        f'\\nSchedule subject to change.'
    )

    if tentative:
        desc += (
            '\\nPotential Game 3; played only '
            'if the series is tied after two games.'
        )

    lines = [
        'BEGIN:VEVENT',
        f'UID:{uid}',
        'DTSTAMP:20261029T000000Z',
        f'DTSTART;TZID={TZ}:{start_s}',
        f'DTEND;TZID={TZ}:{end_s}',
        f'SUMMARY:{ics_escape(title)}',
        f'LOCATION:{ics_escape(location)}',
        f'X-SNL-DIAMOND:Diamond {diamond} — {field}',
        f'DESCRIPTION:{desc}',
        f'STATUS:{"TENTATIVE" if tentative else "CONFIRMED"}',
        'END:VEVENT'
    ]

    return '\r\n'.join(
        fold(s) for s in lines
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        'workbook', type=pathlib.Path
    )
    parser.add_argument(
        '--output',
        type=pathlib.Path,
        default=pathlib.Path('site')
    )
    parser.add_argument(
        '--reference',
        type=pathlib.Path,
        default=pathlib.Path('reference')
    )

    args = parser.parse_args()

    # Read games from the centralized worksheets
    regular = cells_from_xlsx(
        args.workbook, 'Regular Season Data'
    )
    playoff = cells_from_xlsx(
        args.workbook, 'Playoff Data'
    )

    # Read only the organizer's random-draw inputs
    draw_sheet = cells_from_xlsx(
        args.workbook, 'Playoff Seeding'
    )

    rows = make_legacy_rows(regular, playoff)
    draw_order = read_draw_order(draw_sheet)

    ref = read_reference_uids(args.reference)

    # Preserve game IDs from published calendars
    uid_path = args.reference / 'game_uids.json'

    if not uid_path.exists():
        raise FileNotFoundError(
            f'Missing permanent UID registry: {uid_path}'
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
            'Permanent UID registry must contain '
            '60 unique original games'
        )

    games = {
        team: [] for team in TEAMS
    }

    count = 0

    for r in range(5, 66):
        row = rows.get(r, {})

        if not (
            row.get('B')
            and row.get('D')
            and row.get('E')
            and row.get('F')
            and row.get('G')
        ):
            continue

        date = date_from_serial(row['B'])
        slot = int(float(row['D']))

        diamond = re.search(
            r'\d', row['E']
        ).group()

        away, home = row['F'], row['G']

        if away not in games or home not in games:
            raise ValueError(
                f'Unknown team on row {r}: {away}, {home}'
            )

        stamp = dt.datetime.combine(
            date,
            dt.time(
                *{
                    1: (22, 0),
                    2: (22, 45),
                    3: (23, 30)
                }[slot]
            )
        ).strftime('%Y%m%dT%H%M%S')

        uid = original_uids[r]

        games[away].append(
            emit_event(
                uid, date, slot, diamond,
                away, home, 'Away',
                'Regular Season',
                time_range=row['C']
            )
        )

        games[home].append(
            emit_event(
                uid, date, slot, diamond,
                home, away, 'Home',
                'Regular Season',
                time_range=row['C']
            )
        )

        count += 1

    if count != 60:
        raise ValueError(
            f'Expected 60 regular-season games; '
            f'found {count}. Check source workbook.'
        )

    # Calculate playoff seeds from completed results
    seeds = official_seeds(rows, draw_order)

    if seeds is None:
        print(
            'Official playoff seeds not finalized; '
            'retaining neutral playoff placeholders.'
        )

    scores = {}

    for r in range(71, 88):
        if r in (
            SERIES['QF_3v6']
            + SERIES['QF_4v5']
            + SERIES['SF_1']
            + SERIES['SF_2']
            + SERIES['FINAL']
        ):
            raw = rows.get(r, {})
            a = str(raw.get('H', '')).strip()
            h = str(raw.get('I', '')).strip()

            if a or h:
                if not a.isdecimal() or not h.isdecimal():
                    raise ValueError(
                        f'Invalid playoff score row {r}'
                    )

                scores[r] = (int(a), int(h))

    rounds = bracket(seeds, scores)

    for key, result in rounds.items():
        phase = {
            'QF': 'Quarterfinals',
            'SF': 'Semifinals',
            'FINAL': 'Finals'
        }[key.split('_')[0]]

        for game_index, row_num in enumerate(
            SERIES[key]
        ):
            if (
                game_index == 2
                and result.game3 == 'not_needed'
            ):
                continue

            source = rows[row_num]
            date = date_from_serial(source['B'])
            slot = int(float(source['D']))

            diamond = re.search(
                r'\d', source['E']
            ).group()

            if game_index == 1:
                home, away = (
                    result.lower, result.higher
                )
            else:
                home, away = (
                    result.higher, result.lower
                )

            # Keep original playoff event IDs
            uid = (
                f'snl-playoff-row-{row_num - 1}'
                '@sundaynightlights'
            )

            tentative = (
                game_index == 2
                and result.game3 == 'potential'
            )

            games[home].append(
                emit_event(
                    uid, date, slot, diamond,
                    home, away, 'Home',
                    phase, tentative
                )
            )

            games[away].append(
                emit_event(
                    uid, date, slot, diamond,
                    away, home, 'Away',
                    phase, tentative
                )
            )

    # Preserve playoff placeholders until matchups are known
    phase_dates = {
        'Quarterfinals': '20270117',
        'Semifinals': '20270124',
        'Finals': '20270131'
    }

    for team in TEAMS:
        original = (
            args.reference
            / f'SNL_{PREFIX[team]}_Fall_2026-27.ics'
        )

        if not original.exists():
            continue

        data = original.read_text(
            encoding='utf-8-sig'
        )

        for chunk in data.split('BEGIN:VEVENT')[1:]:
            block = (
                'BEGIN:VEVENT'
                + chunk.split('END:VEVENT')[0]
                + 'END:VEVENT'
            )

            if '(Placeholder)' not in block:
                continue

            phase = next(
                (
                    name for name in phase_dates
                    if f'SUMMARY:{name} (Placeholder)' in block
                ),
                None
            )

            if not phase:
                continue

            if any(
                f'Sunday Night Lights — {phase}\\n' in event
                for event in games[team]
            ):
                continue

            games[team].append(
                block.replace('\r\n', '\n')
                .replace('\n', '\r\n')
            )

    args.output.mkdir(
        parents=True, exist_ok=True
    )

    for team, events in games.items():
        name = (
            f'SNL_{PREFIX[team]}_Fall_2026-27.ics'
        )

        header = [
            'BEGIN:VCALENDAR',
            'VERSION:2.0',
            'PRODID:-//SNL//Team Schedules 2026-27//EN',
            'CALSCALE:GREGORIAN',
            'METHOD:PUBLISH',
            f'X-WR-CALNAME:{ics_escape("SNL – " + team + " – Fall 2026–27")}',
            f'X-WR-TIMEZONE:{TZ}'
        ]

        content = '\r\n'.join(
            header + events + ['END:VCALENDAR', '']
        )

        (args.output / name).write_bytes(
            content.encode('utf-8')
        )

        print(
            name, len(events), 'events'
        )

    links = '\n'.join(
        f'<li><a href="{name}">{team}</a></li>'
        for team, name in [
            (
                team,
                f'SNL_{PREFIX[team]}_Fall_2026-27.ics'
            )
            for team in TEAMS
        ]
    )

    (args.output / 'index.html').write_text(
        '<!doctype html><html lang="en">'
        '<meta charset="utf-8">'
        '<title>SNL Team Calendars</title>'
        '<h1>Sunday Night Lights — Fall 2026–27</h1>'
        '<p>Subscribe to your team calendar '
        'using the .ics URL.</p>'
        '<ul>' + links + '</ul></html>',
        encoding='utf-8'
    )


if __name__ == '__main__':
    main()
