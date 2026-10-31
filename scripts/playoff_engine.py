
"""SNL playoff progression, independent of Excel formula caches.

Scores are (away, home), matching Master Schedule columns H and I.
None denotes an unplayed game. Tied playoff scores are invalid.
"""
from dataclasses import dataclass

SERIES = {
    'QF_3v6': (71, 73, 75),
    'QF_4v5': (72, 74, 76),
    'SF_1': (78, 80, 82),
    'SF_2': (79, 81, 83),
    'FINAL': (85, 86, 87),
}


@dataclass(frozen=True)
class SeriesResult:
    higher: str
    lower: str
    wins_higher: int
    wins_lower: int
    game3: str  # potential, required, not_needed, completed
    winner: str | None


def series_result(higher, lower, rows, scores):
    """rows=(G1,G2,G3), home sides alternate higher/lower/higher.
    Game 3 is treated as potential before two games are completed.
    """
    wins = {higher: 0, lower: 0}
    for i, row in enumerate(rows):
        pair = scores.get(row)
        if pair is None:
            continue
        if not isinstance(pair, (tuple, list)) or len(pair) != 2 or any(type(x) is not int or x < 0 for x in pair):
            raise ValueError(f'Invalid playoff score on row {row}: {pair!r}')
        away, home = pair
        if away == home:
            raise ValueError(f'Playoff game on row {row} cannot end tied')
        # Game 1: lower away, higher home; game 2: higher away, lower home.
        # Game 3: lower away, higher home.
        winner = (lower if away > home else higher) if i != 1 else (higher if away > home else lower)
        wins[winner] += 1
        if wins[winner] > 2:
            raise ValueError(f'Scores entered after series decided: {rows}')
    first_two = [scores.get(r) for r in rows[:2]]
    third = scores.get(rows[2])
    if third is not None and (None in first_two or max(wins.values()) != 2):
        raise ValueError('Game 3 entered before valid Games 1 and 2')
    if third is not None:
        state = 'completed'
    elif None in first_two:
        state = 'potential'
    elif wins[higher] == 1 and wins[lower] == 1:
        state = 'required'
    else:
        state = 'not_needed'
    champion = next((team for team, n in wins.items() if n == 2), None)
    return SeriesResult(higher, lower, wins[higher], wins[lower], state, champion)


def bracket(seeds, scores):
    """Returns only determined series. seeds is list of six team names in rank order.
    A missing seed list means regular season is not finalized.
    """
    if seeds is None:
        return {}
    if len(seeds) != 6 or len(set(seeds)) != 6:
        raise ValueError('Six distinct playoff seeds required')
    out = {
        'QF_3v6': series_result(seeds[2], seeds[5], SERIES['QF_3v6'], scores),
        'QF_4v5': series_result(seeds[3], seeds[4], SERIES['QF_4v5'], scores),
    }
    a, b = out['QF_3v6'].winner, out['QF_4v5'].winner
    if a is None or b is None:
        return out
    higher_remaining, lower_remaining = sorted((a,b), key=seeds.index)
    out['SF_1'] = series_result(seeds[0], lower_remaining, SERIES['SF_1'], scores)
    out['SF_2'] = series_result(seeds[1], higher_remaining, SERIES['SF_2'], scores)
    a, b = out['SF_1'].winner, out['SF_2'].winner
    if a is None or b is None:
        return out
    high, low = sorted((a,b), key=seeds.index)
    out['FINAL'] = series_result(high, low, SERIES['FINAL'], scores)
    return out
