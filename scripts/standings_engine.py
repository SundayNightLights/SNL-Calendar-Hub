
"""Official SNL regular-season ranking without evaluating Excel formulas.

Game records come from Regular Season Data through the calendar adapter.
Random Draw Order comes from Playoff Seeding column K.

Unfinished seasons or unresolved draws cannot produce official seeds.
"""
from collections import defaultdict

TEAMS = (
    'Buffaloes', 'IHT', 'MORSL', 'Outcasts',
    'Pacific Maul', 'Toronto Raccoons'
)


def _score(v):
    if v is None or str(v).strip() == '':
        return None
    s = str(v).strip()
    if not s.isdecimal():
        raise ValueError(f'Invalid score {v!r}')
    return int(s)


def official_seeds(rows, draw_order=None):
    """Return six official seeds or None if unresolved.

    Tiebreakers:
    1. Standings points
    2. Most wins
    3. Head-to-head points
    4. Capped run differential (+/-7 per game)
    5. Fewest runs allowed
    6. Most runs scored
    7. Recorded random draw

    Random draw values are absolute seed positions.
    """
    draw_order = draw_order or {}
    stats = {
        t: dict(
            points=0, wins=0, rd=0,
            allowed=0, scored=0, gp=0
        )
        for t in TEAMS
    }
    played = []

    for i in range(5, 66):
        row = rows.get(i, {})
        if not row.get('F') and not row.get('G'):
            continue
        if not row.get('F') or not row.get('G'):
            raise ValueError(f'Missing team on game row {i}')

        away, home = row['F'], row['G']
        if away not in stats or home not in stats or away == home:
            raise ValueError(
                f'Invalid teams on game row {i}: {away}, {home}'
            )

        x, y = _score(row.get('H')), _score(row.get('I'))
        if x is None and y is None:
            return None
        if x is None or y is None:
            raise ValueError(f'Incomplete score row {i}')

        played.append((away, home, x, y))

        for team, scored, conceded in (
            (away, x, y), (home, y, x)
        ):
            s = stats[team]
            s['gp'] += 1
            s['scored'] += scored
            s['allowed'] += conceded
            s['rd'] += max(-7, min(7, scored - conceded))
            s['points'] += (
                2 if scored > conceded
                else 1 if scored == conceded else 0
            )
            s['wins'] += int(scored > conceded)

    if len(played) != 60:
        return None
    if any(s['gp'] != 20 for s in stats.values()):
        raise ValueError(
            'Each team must have exactly 20 regular-season games'
        )

    def h2h(team, tied):
        pts = 0
        for away, home, x, y in played:
            if away not in tied or home not in tied:
                continue
            if team == away:
                pts += 2 if x > y else 1 if x == y else 0
            if team == home:
                pts += 2 if y > x else 1 if x == y else 0
        return pts

    def rank(group, criterion=0, start_seed=1):
        if len(group) <= 1:
            return group

        if criterion >= 6:
            priorities = [draw_order.get(team) for team in group]
            if any(type(p) is not int for p in priorities):
                return None
            if sorted(priorities) != list(
                range(start_seed, start_seed + len(group))
            ):
                return None
            return sorted(group, key=lambda t: draw_order[t])

        def value(team):
            s = stats[team]
            return (
                s['points'], s['wins'],
                h2h(team, set(group)),
                s['rd'], -s['allowed'], s['scored']
            )[criterion]

        buckets = defaultdict(list)
        for team in group:
            buckets[value(team)].append(team)

        if len(buckets) == 1:
            return rank(group, criterion + 1, start_seed)

        result = []
        next_seed = start_seed

        for val in sorted(buckets, reverse=True):
            subgroup = buckets[val]
            ordered = rank(subgroup, criterion + 1, next_seed)
            if ordered is None:
                return None
            result.extend(ordered)
            next_seed += len(subgroup)

        return result

    return rank(list(TEAMS))
