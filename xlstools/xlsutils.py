#
# Copyright (C) 2010 - 2012 Red Hat, Inc.
# Red Hat Author(s): Satoru SATOH <ssato@redhat.com>
#
# License: MIT
#
from itertools import groupby

import sys
import xlrd


def show(cell_type, cell_value, datemode):
    """
    @see showable_cell_value in examples/xlrdnameAPIdemo.py in python-xlrd dist.
    """
    if cell_type == xlrd.XL_CELL_EMPTY:
        v = ''
    elif cell_type == xlrd.XL_CELL_DATE:
        try:
            v = xlrd.xldate_as_tuple(cell_value, datemode)
        except xlrd.XLDateError:
            e1, e2 = sys.exc_info()[:2]
            v = "%s:%s" % (e1.__name__, e2)
    elif cell_type == xlrd.XL_CELL_ERROR:
        v = xlrd.error_text_from_code.get(cell_value, '<Unknown error code 0x%02x>' % cell_value)
    else:
        v = cell_value

    return v


def sheet_cell_values_in_the_row_g(sheet, row, datemode=None):
    if datemode is None:
        datemode = sheet.book.datemode

    for y in xrange(0, sheet.ncols):
        ctype = sheet.cell_type(row, y)

        if ctype == xlrd.XL_CELL_EMPTY:
            v = ""
        else:
            v = show(ctype, sheet.cell_value(row, y), datemode)

        yield (y, v)  # row idx, col idx and its value


def sheet_cell_values_g(sheet, row_start):
    datemode = sheet.book.datemode

    for x in xrange(row_start, sheet.nrows):
        for y, v in sheet_cell_values_in_the_row_g(sheet, x, datemode):
            yield (x, y, v)


def fst(tpl_or_list):
    return tpl_or_list[0]


def foreach_sheet_cells_by_row(sheet, row_start=1):
    for k, g in groupby(sheet_cell_values_g(sheet, row_start), fst):
        yield [t[2] for t in g]


def normalize_key(key_str):
    return key_str.lower().replace(" ", "_")


# vim:sw=4:ts=4:et:
