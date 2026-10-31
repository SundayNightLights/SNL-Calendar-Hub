
"""Generate team iCalendar subscriptions from SNL Fall 26.xlsx without evaluating Excel formulas."""
import argparse
import datetime as dt
import hashlib
import json
import pathlib
import re
import xml.etree.ElementTree as ET
import zipfile
from standings_engine import official_seeds
from playoff_engine import bracket, SERIES

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
TEAMS = ['Buffaloes', 'IHT', 'MORSL', 'Outcasts', 'Pacific Maul', 'Toronto Raccoons']
VENUE = 'Markham Sports Dome Inc., 5300 14th Avenue, Markham, ON L3S 3K8'
PREFIX = {'Buffaloes':'Buffaloes','IHT':'IHT','MORSL':'MORSL','Outcasts':'Outcasts','Pacific Maul':'Pacific_Maul','Toronto Raccoons':'Toronto_Raccoons'}
TZ = 'America/Toronto'


def cells_from_xlsx(path, sheet_index=1):
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            shared = [''.join(t.text or '' for t in item.findall('.//m:t', NS)) for item in root]
        root = ET.fromstring(z.read(f'xl/worksheets/sheet{sheet_index}.xml'))
        rows = {}
        for row in root.findall('.//m:sheetData/m:row', NS):
            result = {}
            for cell in row.findall('m:c', NS):
                col = re.match(r'[A-Z]+', cell.get('r')).group()
                value = cell.find('m:v', NS)
                inline = cell.find('m:is', NS)
                if value is not None:
                    v = value.text or ''
                    if cell.get('t') == 's':
                        v = shared[int(v)]
                elif inline is not None:
                    v = ''.join(t.text or '' for t in inline.findall('.//m:t', NS))
                else:
                    v = ''
                # Cached formula values are deliberately not trusted for playoff opponents.
                if cell.find('m:f', NS) is not None:
                    result[col + '_formula'] = True
                result[col] = v
            rows[int(row.get('r'))] = result
        return rows


def date_from_serial(serial):
    return dt.date(1899, 12, 30) + dt.timedelta(days=int(float(serial)))


def ics_escape(s):
    return str(s).replace('\\', '\\\\').replace(';', '\\;').replace(',', '\\,').replace('\n', '\\n')


def fold(line):
    # Fold by UTF-8 octets, keeping continuation lines under 75 octets.
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
    # Prefer the UID of the already-distributed ICS files for the same fixture.
    return hashlib.sha1(f'SNL|master-row|{row}'.encode()).hexdigest()[:20] + '@sundaynightlights'


def read_reference_uids(folder):
    index = {}
    if not folder.exists():
        return index
    for path in folder.glob('*.ics'):
        content = path.read_text(encoding='utf-8-sig').replace('\r\n ', '')
        for part in content.split('BEGIN:VEVENT')[1:]:
            fields = {}
            for line in part.splitlines():
                if ':' in line:
                    k, v = line.split(':', 1)
                    fields[k] = v
            start = fields.get('DTSTART;TZID=America/Toronto')
            summary = fields.get('SUMMARY', '')
            diamond = re.search(r'Diamond (\d)', summary)
            if start and diamond and fields.get('UID'):
                index[(start, diamond.group(1))] = fields['UID']
    return index


def parse_game_times(date, time_range):
    pattern = r'^(\d{1,2}:\d{2})(?:\s*(AM|PM))?\s*[-–—]\s*(\d{1,2}:\d{2})\s*(AM|PM)$'
    match = re.match(pattern, str(time_range).strip(), re.I)
    if not match:
        raise ValueError(f'Invalid game time range: {time_range!r}')
    first, first_period, last, last_period = match.groups()
    first_period = first_period or last_period

    def parse(clock, period):
        return dt.datetime.strptime(clock + ' ' + period.upper(), '%I:%M %p').time()
    start = dt.datetime.combine(date, parse(first, first_period))
    end = dt.datetime.combine(date, parse(last, last_period))
    if end <= start:
        end += dt.timedelta(days=1)
    return start, end


def emit_event(uid, date, slot, diamond, team, opponent, side, phase, tentative=False, time_range=None):
    if time_range is not None:
        start, end = parse_game_times(date, time_range)
    else:
        start_hours = {1: (22, 0), 2: (22, 45), 3: (23, 30)}
        hour, minute = start_hours[slot]
        start = dt.datetime.combine(date, dt.time(hour, minute))
        end = start + dt.timedelta(minutes=45)
    start_s, end_s = start.strftime('%Y%m%dT%H%M%S'), end.strftime('%Y%m%dT%H%M%S')
    field = 'Near Field' if diamond == '1' else 'Far Field'
    suffix = ' (If Necessary)' if tentative else ''
    title = f'{team} vs {opponent} ({side} | Diamond {diamond}){suffix}'
    desc = f'Sunday Night Lights — {phase}\\nDiamond {diamond} — {field}\\nVenue: {VENUE}\\nSchedule subject to change.'
    if tentative:
        desc += '\\nPotential Game 3; played only if the series is tied after two games.'
    lines = ['BEGIN:VEVENT', f'UID:{uid}', 'DTSTAMP:20261029T000000Z',
             f'DTSTART;TZID={TZ}:{start_s}', f'DTEND;TZID={TZ}:{end_s}',
             f'SUMMARY:{ics_escape(title)}', f'LOCATION:{ics_escape(VENUE)}',
             f'X-SNL-DIAMOND:Diamond {diamond} — {field}', f'DESCRIPTION:{desc}',
             f'STATUS:{"TENTATIVE" if tentative else "CONFIRMED"}', 'END:VEVENT']
    return '\r\n'.join(fold(s) for s in lines)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('workbook', type=pathlib.Path)
    parser.add_argument('--output', type=pathlib.Path, default=pathlib.Path('site'))
    parser.add_argument('--reference', type=pathlib.Path, default=pathlib.Path('reference'))
    args = parser.parse_args()
    rows = cells_from_xlsx(args.workbook)
    ref = read_reference_uids(args.reference)
    # Permanent row-to-UID registry captured from the original published calendars.
    # Rescheduling a game must not change its UID.
    uid_path = args.reference / 'game_uids.json'
    if not uid_path.exists():
        raise FileNotFoundError(f'Missing permanent UID registry: {uid_path}')
    original_uids = {int(k): v for k, v in json.loads(uid_path.read_text()).items()}
    if len(original_uids) != 60 or len(set(original_uids.values())) != 60:
        raise ValueError('Permanent UID registry must contain 60 unique original games')
    games = {team: [] for team in TEAMS}
    count = 0
    for r in range(5, 66):
        row = rows.get(r, {})
        if not (row.get('B') and row.get('D') and row.get('E') and row.get('F') and row.get('G')):
            continue
        date = date_from_serial(row['B'])
        slot = int(float(row['D']))
        diamond = re.search(r'\d', row['E']).group()
        away, home = row['F'], row['G']
        if away not in games or home not in games:
            raise ValueError(f'Unknown team on row {r}: {away}, {home}')
        stamp = dt.datetime.combine(date, dt.time(*{1:(22,0),2:(22,45),3:(23,30)}[slot])).strftime('%Y%m%dT%H%M%S')
        uid = original_uids[r]
        games[away].append(emit_event(uid, date, slot, diamond, away, home, 'Away', 'Regular Season', time_range=row['C']))
        games[home].append(emit_event(uid, date, slot, diamond, home, away, 'Home', 'Regular Season', time_range=row['C']))
        count += 1
    if count != 60:
        raise ValueError(f'Expected 60 regular-season games; found {count}. Check source workbook.')
    # Score-driven playoff matchups, never stale cached Excel formula outputs.
    seeds = official_seeds(rows)
    scores = {}
    for r in range(71, 88):
        if r in SERIES['QF_3v6'] + SERIES['QF_4v5'] + SERIES['SF_1'] + SERIES['SF_2'] + SERIES['FINAL']:
            raw = rows.get(r,{})
            a, h = str(raw.get('H','')).strip(), str(raw.get('I','')).strip()
            if a or h:
                if not a.isdecimal() or not h.isdecimal():
                    raise ValueError(f'Invalid playoff score row {r}')
                scores[r] = (int(a),int(h))
    rounds = bracket(seeds,scores)
    for key, result in rounds.items():
        phase = {'QF':'Quarterfinals','SF':'Semifinals','FINAL':'Finals'}[key.split('_')[0]]
        for game_index, row_num in enumerate(SERIES[key]):
            if game_index == 2 and result.game3 == 'not_needed':
                continue
            source = rows[row_num]
            date = date_from_serial(source['B'])
            slot = int(float(source['D']))
            diamond = re.search(r'\d',source['E']).group()
            home,away = (result.lower,result.higher) if game_index==1 else (result.higher,result.lower)
            # Preserve original playoff event UIDs after the one-row shift.
            uid = f'snl-playoff-row-{row_num - 1}@sundaynightlights'
            tentative = game_index == 2 and result.game3 == 'potential'
            games[home].append(emit_event(uid,date,slot,diamond,home,away,'Home',phase,tentative))
            games[away].append(emit_event(uid,date,slot,diamond,away,home,'Away',phase,tentative))
    # Until seeds or future matchups are known, preserve original neutral
    # playoff placeholders so every team sees the potential playoff weeks.
    phase_dates = {'Quarterfinals':'20270117','Semifinals':'20270124','Finals':'20270131'}
    for team in TEAMS:
        original = (args.reference / f'SNL_{PREFIX[team]}_Fall_2026-27.ics')
        if not original.exists(): continue
        data = original.read_text(encoding='utf-8-sig')
        for chunk in data.split('BEGIN:VEVENT')[1:]:
            block = 'BEGIN:VEVENT' + chunk.split('END:VEVENT')[0] + 'END:VEVENT'
            if '(Placeholder)' not in block:continue
            phase = next((name for name in phase_dates if f'SUMMARY:{name} (Placeholder)' in block),None)
            if not phase:continue
            # A team's neutral placeholder is replaced by real game entries
            # only when that team's actual matchup for the round is known.
            if any(f'Sunday Night Lights — {phase}\\n' in event for event in games[team]):
                continue
            games[team].append(block.replace('\r\n','\n').replace('\n','\r\n'))
    args.output.mkdir(parents=True, exist_ok=True)
    for team, events in games.items():
        name = f'SNL_{PREFIX[team]}_Fall_2026-27.ics'
        header = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//SNL//Team Schedules 2026-27//EN',
                  'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
                  f'X-WR-CALNAME:{ics_escape("SNL – " + team + " – Fall 2026–27")}',
                  f'X-WR-TIMEZONE:{TZ}']
        content = '\r\n'.join(header + events + ['END:VCALENDAR', ''])
        (args.output / name).write_bytes(content.encode('utf-8'))
        print(name, len(events), 'events')
    links = '\n'.join(f'<li><a href="{name}">{team}</a></li>' for team, name in
                      [(team, f'SNL_{PREFIX[team]}_Fall_2026-27.ics') for team in TEAMS])
    (args.output / 'index.html').write_text('<!doctype html><html lang="en"><meta charset="utf-8"><title>SNL Team Calendars</title><h1>Sunday Night Lights — Fall 2026–27</h1><p>Subscribe to your team calendar using the .ics URL.</p><ul>' + links + '</ul></html>', encoding='utf-8')


if __name__ == '__main__':
    main()
