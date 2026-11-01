"""Order-preserving alignment of two sequences (rows or columns).

The algorithm is a fuzzy extension of "patience diff":

1. Identical items at the start / end of a range are matched directly.
2. Items whose exact key occurs only once on each side become anchors.
   The longest increasing subsequence keeps anchors in order, so moved
   items do not scramble the result. The gaps between anchors are solved
   recursively.
3. If a range has no exact anchors, values that occur in only one item on
   each side "vote" for item pairs (e.g. an ID cell identifies its row even
   after other cells or a whole column changed). A voted pair that is at
   least ``threshold`` similar becomes an anchor.
4. A range that still has no anchors is aligned with dynamic programming
   that maximizes the total similarity of paired items. Ties prefer pairing
   over insertion/deletion, and pairing at the same relative position, so
   fully rewritten items are reported as modifications, not shifts.

Indices are 0-based. The result lists every item exactly once, in order:
``(i, j)`` is a pair, ``(i, None)`` exists only on side A (deleted) and
``(None, j)`` only on side B (inserted).
"""

from bisect import bisect_left
from collections.abc import Callable, Hashable, Sequence

Pair = tuple[int | None, int | None]

# Items at least this similar may anchor a gap; weaker similarities do not
# count as evidence when deciding where insertions/deletions happened.
SAME_THRESHOLD = 0.5
_QUANT = 1000


def align(
    keys_a: Sequence[Hashable],
    keys_b: Sequence[Hashable],
    tokens_a: Sequence[frozenset],
    tokens_b: Sequence[frozenset],
    similarity: Callable[[int, int], float],
    *,
    blank: Hashable,
    budget: int,
    threshold: float = SAME_THRESHOLD,
) -> list[Pair]:
    """Align two sequences. ``budget`` bounds the DP size (n*m) of one gap."""
    matches: list[tuple[int, int]] = []
    stack = [(0, len(keys_a), 0, len(keys_b))]
    while stack:
        a0, a1, b0, b1 = stack.pop()
        while a0 < a1 and b0 < b1 and keys_a[a0] == keys_b[b0]:
            matches.append((a0, b0))
            a0 += 1
            b0 += 1
        while a0 < a1 and b0 < b1 and keys_a[a1 - 1] == keys_b[b1 - 1]:
            a1 -= 1
            b1 -= 1
            matches.append((a1, b1))
        if a0 == a1 or b0 == b1:
            continue
        anchors = _key_anchors(keys_a, keys_b, a0, a1, b0, b1, blank)
        if not anchors:
            anchors = _token_anchors(tokens_a, tokens_b, a0, a1, b0, b1, similarity, threshold)
        if anchors:
            i0, j0 = a0, b0
            for i, j in anchors:
                matches.append((i, j))
                stack.append((i0, i, j0, j))
                i0, j0 = i + 1, j + 1
            stack.append((i0, a1, j0, b1))
        else:
            matches.extend(_dp(a0, a1, b0, b1, similarity, threshold, budget))
    matches.sort()
    result: list[Pair] = []
    i = j = 0
    for mi, mj in matches:
        result.extend((k, None) for k in range(i, mi))
        result.extend((None, k) for k in range(j, mj))
        result.append((mi, mj))
        i, j = mi + 1, mj + 1
    result.extend((k, None) for k in range(i, len(keys_a)))
    result.extend((None, k) for k in range(j, len(keys_b)))
    return result


def _lis(candidates: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Longest subsequence with increasing j. ``candidates`` is sorted by i."""
    tails: list[int] = []
    tail_js: list[int] = []
    previous = [-1] * len(candidates)
    for index, (_, j) in enumerate(candidates):
        position = bisect_left(tail_js, j)
        if position:
            previous[index] = tails[position - 1]
        if position == len(tails):
            tails.append(index)
            tail_js.append(j)
        else:
            tails[position] = index
            tail_js[position] = j
    chain = []
    index = tails[-1] if tails else -1
    while index >= 0:
        chain.append(candidates[index])
        index = previous[index]
    return chain[::-1]


def _unique_positions(items, start, stop):
    positions = {}
    for index in range(start, stop):
        for item in items(index):
            positions[item] = index if item not in positions else -1
    return positions


def _key_anchors(keys_a, keys_b, a0, a1, b0, b1, blank):
    in_a = _unique_positions(lambda i: (keys_a[i],), a0, a1)
    in_b = _unique_positions(lambda j: (keys_b[j],), b0, b1)
    candidates = [
        (i, in_b[key]) for key, i in in_a.items()
        if i >= 0 and key != blank and in_b.get(key, -1) >= 0
    ]
    return _lis(sorted(candidates))


def _token_anchors(tokens_a, tokens_b, a0, a1, b0, b1, similarity, threshold):
    in_a = _unique_positions(tokens_a.__getitem__, a0, a1)
    in_b = _unique_positions(tokens_b.__getitem__, b0, b1)
    votes: dict[int, dict[int, int]] = {}
    for token, i in in_a.items():
        j = in_b.get(token, -1)
        if i >= 0 and j >= 0:
            row = votes.setdefault(i, {})
            row[j] = row.get(j, 0) + 1
    candidates = []
    for i, row in votes.items():
        j = max(row, key=lambda k: (row[k], -k))
        if similarity(i, j) >= threshold:
            candidates.append((i, j))
    return _lis(sorted(candidates))


def _dp(a0, a1, b0, b1, similarity, threshold, budget):
    n, m = a1 - a0, b1 - b0
    if n * m > budget:
        # Too large to search: fall back to position-based pairing.
        return [(a0 + k, b0 + k) for k in range(min(n, m))]
    # Integer scores give exact, deterministic tie-breaking:
    # similarity  >>  number of pairs  >>  pairs at the same relative position.
    unit = min(n, m) + 1
    weight = unit * unit
    width = m + 1
    previous = [0] * width
    moves = bytearray((n + 1) * width)  # 0: pair, 1: only A, 2: only B
    for j in range(1, width):
        moves[j] = 2
    for i in range(1, n + 1):
        current = [0] * width
        moves[i * width] = 1
        for j in range(1, width):
            sim = similarity(a0 + i - 1, b0 + j - 1)
            quality = round(sim * _QUANT) if sim >= threshold else 0
            best = previous[j - 1] + quality * weight + unit + (i == j)
            move = 0
            if previous[j] > best:
                best, move = previous[j], 1
            if current[j - 1] > best:
                best, move = current[j - 1], 2
            current[j] = best
            moves[i * width + j] = move
        previous = current
    pairs = []
    i, j = n, m
    while i and j:
        move = moves[i * width + j]
        if move == 0:
            pairs.append((a0 + i - 1, b0 + j - 1))
            i -= 1
            j -= 1
        elif move == 1:
            i -= 1
        else:
            j -= 1
    return pairs[::-1]
