#
# Copyright (C) 2010 - 2012 Satoru SATOH <satoru.satoh at gmail.com>
# License: MIT
#
from functools import reduce as fold, partial as curry


def zipWith(f, xs=[], ys=[]):
    """
    >>> zipWith(max, [3, 3, 8, 2], [2, 1, 5, 7])
    [3, 3, 8, 7]
    >>> zipWith(lambda x, y: (x,y), [3, 3, 8, 2], [2, 1, 5, 7])
    [(3, 2), (3, 1), (8, 5), (2, 7)]
    >>> assert zipWith(lambda x, y: (x,y), [3, 3, 8], [2, 1, 5]) == zip([3, 3, 8], [2, 1, 5])
    >>>

    @TODO: Handle cases if len(xs) != len(ys), for example, use zip' instead of zip
    where
        zip' [] ys = [(None, y)| y <- ys]  # it works if f is max but not if f is min
        zip' xs [] = [(x, None)| x <- xs]  # because min(None, y) returns nothing (undef).
        zip' xs ys = zip xs ys
    """
    assert callable(f)
    return [f(x, y) for x, y in zip(xs, ys)]


def max_col_widths(xss):
    """
    @return list of max width needed for columns (:: [Int]). see an example below.

    >>> xss = [['aaa', 'bbb', 'cccccccc', 'dd'], ['aa', 'b', 'ccccc', 'ddddddd'], ['aaaa', 'bbbb', 'c', 'dd']]
    >>> max_col_widths(xss)
    [4, 4, 8, 7]
    """
    yss = [[len(x) for x in xs] for xs in xss]
    return fold(curry(zipWith, max), yss[1:], yss[0])


def max_col_widths_2(xss):
    """
    More straight-forward implementation optimized for matric.

    >>> xss = [['aaa', 'bbb', 'cccccccc', 'dd'], ['aa', 'b', 'ccccc', 'ddddddd'], ['aaaa', 'bbbb', 'c', 'dd']]
    >>> max_col_widths(xss)
    [4, 4, 8, 7]
    """
    yss = [[len(x) for x in xs] for xs in xss]
    return [max((yss[i][j] for i in range(0, len(yss)))) for j in range(0, len(yss[0]))]


def adjust_width(width):
    """@FIXME: Tune factor and threashold values.
    """
    factor0 = 200
    factor1 = 10
    threashold0 = 15

    if width < threashold0:
        width += factor1

    return width * factor0


def mergeable_cells(xss, row_start=0, row_end=-1, col_start=0, col_end=-1):
    """
    @return (cell_value, r1, r2, c1, c2) of merge-able cells.

    >>> mergeable_cells([['a', 'b', 'c'], ['a', 'c', 'b']])
    [('a', 0, 1, 0, 0)]
    >>> mergeable_cells([['a', 'b', 'c'], ['a', 'c', 'b'], ['c', 'a', 'b'], ['d', 'e', 'b']])
    [('a', 0, 1, 0, 0), ('b', 1, 3, 2, 2)]
    >>> mergeable_cells([['a', 'b', 'c'], ['a', 'c', 'b'], ['c', 'a', 'b'], ['d', 'e', 'b']], 1)
    [('b', 1, 3, 2, 2)]
    >>> mergeable_cells([['a', 'b', 'c'], ['a', 'c', 'b'], ['c', 'a', 'b'], ['d', 'e', 'b']], col_end=1)
    [('a', 0, 1, 0, 0)]
    """
    rl = (row_end < 0 and len(xss) or row_end)
    cl = (col_end < 0 and max((len(xss[r]) for r in range(0, rl))) or col_end)
    ret = []

    for c in range(col_start, cl):
        yss = [(xss[r][c], r, c) for r in range(row_start, rl) if len(xss[r]) > c]
        mss = [[yss[0]]]  # Possibly mergeable cells

        for ys in yss[1:]:
            ms_last = mss[-1]

            if ys[0] == ms_last[-1][0]: # If matched (maybe mergeable),
                if len(ms_last) > 1:    # and there are more than one cells pushed already
                    mss[-1][-1] = ys    # then replace the last cell pushed with it,
                else:
                    mss[-1] += [ys]     # or append it.
            else:
                mss += [[ys]]           # Append the next candidate ms if not matched.

        #pprint.pprint(mss)
        # remove non-mergeable cells having only one cell and convert each pair
        # [(val, r1, c1), (val, r2, c2)] to (val, r1, r2, c1, c2).
        mcs = [(ms[0][0],ms[0][1],ms[1][1],ms[0][2],ms[1][2]) for ms in mss if len(ms) > 1]
        ret += mcs

    return ret


def normalize_key(key):
    """Normalize key name to be used as SQL key name.

    >>> normalize_key('Key Name')
    key_name
    >>>
    """
    return key.lower().strip().replace(' ', '_')


def rename_if_exists(target, suffix='.bak'):
    """If the file $target (file or dir) exists, it will be renamed and backed
    up as 'TARGET.${suffix}'. (The default suffix is '.bak'.)
    """
    if os.path.exists(target):
        os.rename(target, target + suffix)


# vim:sw=4:ts=4:et:
