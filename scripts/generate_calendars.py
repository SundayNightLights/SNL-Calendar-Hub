
def make_legacy_rows(regular, playoff):
    """Map permanent Game IDs to original published calendar IDs.

    Game IDs remain stable even when Excel rows are moved.
    Historical row numbers are retained for compatibility
    with existing calendar UIDs and playoff calculations.
    """
    required = {
        'A': 'Game ID',
        'B': 'Week',
        'C': 'Away Team',
        'D': 'Home Team',
        'E': 'Date',
        'F': 'Time',
        'G': 'Slot',
        'H': 'Diamond',
        'I': 'Away Score',
        'J': 'Home Score'
    }

    for title, data in [
        ('Regular Season Data', regular),
        ('Playoff Data', playoff)
    ]:
        header = data.get(4, {})
        for col, label in required.items():
            if str(header.get(col, '')).strip() != label:
                raise ValueError(
                    f'{title} header {col} expected '
                    f'{label!r}; got {header.get(col)!r}'
                )

    def convert(record, source, source_row):
        converted = {
            'A': record.get('B', ''),
            'B': record.get('E', ''),
            'C': record.get('F', ''),
            'D': record.get('G', ''),
            'E': record.get('H', ''),
            'F': record.get('C', ''),
            'G': record.get('D', ''),
            'H': record.get('I', ''),
            'I': record.get('J', '')
        }

        for col in ('B', 'C', 'D', 'E'):
            if not converted[col]:
                raise ValueError(
                    f'{source} row {source_row}: '
                    'missing date/time/slot/diamond'
                )

        return converted

    legacy = {}
    regular_ids = set()

    for source_row, record in regular.items():
        game_id = str(record.get('A', '')).strip()
        match = re.fullmatch(r'F26-(\d{3})', game_id)

        if not match:
            continue

        number = int(match.group(1))

        if not 1 <= number <= 60 or game_id in regular_ids:
            raise ValueError(
                f'Invalid or duplicate Game ID: {game_id!r}'
            )

        regular_ids.add(game_id)

        # Original holiday row followed game 48.
        legacy_row = number + 4 + (number >= 49)

        legacy[legacy_row] = convert(
            record, 'Regular Season Data', source_row
        )

        away = legacy[legacy_row]['F']
        home = legacy[legacy_row]['G']

        if away not in TEAMS or home not in TEAMS or away == home:
            raise ValueError(
                f'Invalid teams for {game_id}: '
                f'{away!r}, {home!r}'
            )

    expected_regular = {
        f'F26-{n:03d}' for n in range(1, 61)
    }

    if regular_ids != expected_regular:
        raise ValueError(
            'Missing regular-season Game IDs: '
            f'{sorted(expected_regular - regular_ids)}'
        )

    # Original playoff identifiers are preserved.
    playoff_rows = (
        71, 72, 73, 74, 75, 76,
        78, 79, 80, 81, 82, 83,
        85, 86, 87
    )

    playoff_ids = set()

    for source_row, record in playoff.items():
        game_id = str(record.get('A', '')).strip()
        match = re.fullmatch(r'F26-P(\d{2})', game_id)

        if not match:
            continue

        number = int(match.group(1))

        if not 1 <= number <= 15 or game_id in playoff_ids:
            raise ValueError(
                f'Invalid or duplicate playoff ID: {game_id!r}'
            )

        playoff_ids.add(game_id)

        legacy[playoff_rows[number - 1]] = convert(
            record, 'Playoff Data', source_row
        )

    expected_playoffs = {
        f'F26-P{n:02d}' for n in range(1, 16)
    }

    if playoff_ids != expected_playoffs:
        raise ValueError(
            'Missing playoff Game IDs: '
            f'{sorted(expected_playoffs - playoff_ids)}'
        )

    return legacy
