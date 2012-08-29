#
# Generate a series of CSV/JSON files from an Excel (.xls) file.
#
# Copyright (C) 2010, 2011 Red Hat, Inc.
# Red Hat Author(s): Satoru SATOH <ssato@redhat.com>
#
# License: MIT
#
import xlstools.csvx as XC
import xlstools.xlsutils as XU

import logging
import optparse
import os.path
import os
import sys
import xlrd

try:
    from collections import OrderedDict as dict
except ImportError:
    pass

try:
    import json
except ImportError:
    try:
        import simplejson as json
    except ImportError:
        logging.warn("JSON support will be disabled as json module not found.")

        class json(object):

            @staticmethod
            def load(*args, **kwargs):
                raise RuntimeError("Not supported.")

            @staticmethod
            def dump(*args, **kwargs):
                raise RuntimeError("Not supported.")


class DataDumper(object):

    suffix = ".dat"

    def __init__(self, worksheet, name=None, headers=[], outdir=os.curdir):
        self.worksheet = worksheet
        self.name = name is None and self.worksheet.name or name
        self.output = os.path.join(outdir, self.name + self.suffix)

        if headers:
            self.headers = [XU.normalize_key(h) for h in headers]
            self.row_start = 0
        else:
            self.headers = self.get_headers(self.worksheet)
            self.row_start = 1

    def get_headers(self, worksheet):
        return [XU.normalize_key(val) or "-" for idx, val in XU.sheet_cell_values_in_the_row_g(worksheet, 0)]

    def open(self, flag="w"):
        return open(self.output, flag)

    def foreach_sheet_cells_by_row(self):
        return XU.foreach_sheet_cells_by_row(self.worksheet, self.row_start)

    def dump_impl(self):
        raise NotImplementedError("Children classes must implement this!")

    def dump(self):
        logging.info(" Try to dump data in sheet '%s' to '%s'" % (self.worksheet.name, self.output))
        self.dump_impl()
        logging.info(" Done: %s" % self.output)


class CsvDumper(DataDumper):

    suffix = ".csv"

    def open(self):
        return super(CsvDumper, self).open(flag="wb")

    def dump_impl(self):
        out = self.open()
        writer = XC.UnicodeWriter(out)

        writer.writerow(self.headers)

        for rowdata in self.foreach_sheet_cells_by_row():
            writer.writerow(rowdata)

        out.close()


class JsonDumper(DataDumper):

    suffix = ".json"

    def dump_impl(self):
        data = [dict(zip(self.headers, rowdata)) for rowdata in self.foreach_sheet_cells_by_row()]
        json.dump(data, self.open(), ensure_ascii=False, indent=2)


DUMPERS = dict(
    csv=CsvDumper,  # default
    json=JsonDumper,
)


def xls_to(xls_file, dumper, outdir, names=[], headers=[], dumper_map=DUMPERS):
    book = xlrd.open_workbook(xls_file)

    if not os.path.isdir(outdir):
        os.makedirs(outdir)

    for n in range(0, book.nsheets):
        sheet = book.sheet_by_index(n)

        if names and len(names) > n:
            name = names[n]
        else:
            name = sheet.name

        dmpr = dumper_map[dumper](sheet, name, headers, outdir)
        dmpr.dump()


def opts_parser(dumper_map=DUMPERS):
    p = optparse.OptionParser("""%prog [OPTION ...] XLS_FILE

Examples:
  %prog ABC.xls --outdir /tmp/outputs
  %prog ABC.xls --names=aaa,bbb,ccc
""")

    dumper_choices = dumper_map.keys()

    defaults = {
        "names": "",
        "headers": "",
        "dumper": "csv",
        "outdir": os.curdir,
        "verbose": False,
        "quiet": False,
    }
    p.set_defaults(**defaults)

    cog = optparse.OptionGroup(p, "Common Options")
    cog.add_option("", "--dumper", type="choice", choices=dumper_choices,
        help="Select dump format from " + ", ".join(dumper_choices) + " [%default]")
    cog.add_option("", "--names", help="Comma separated filenames")
    cog.add_option("", "--headers",
        help="Comma separated list of headers [default: cell contents in 1st row of input .xls]")
    cog.add_option("-o", "--outdir", help="Specify output dir [%default]")
    cog.add_option("-v", "--verbose", help="Verbose mode", action="store_true")
    cog.add_option("-q", "--quiet", help="Quiet mode", action="store_true")
    #cog.add_option('-T', '--test', help='Test mode - running test suites', default=False, action="store_true")
    p.add_option_group(cog)

    return p


def main(dumper_map=DUMPERS):
    loglevel = logging.WARN
    sheet_names = {}

    parser = opts_parser()
    (options, args) = parser.parse_args()

    if options.verbose:
        loglevel = logging.INFO

    if options.quiet:
        loglevel = logging.ERROR

    logging.basicConfig(level=loglevel)

    if len(args) < 1:
        parser.print_help()
        sys.exit(0)

    names = options.names and options.names.split(',') or []
    headers = options.headers and options.headers.split(',') or []

    xls_file = args[0]

    xls_to(xls_file, options.dumper, options.outdir, names, headers)


if __name__ == '__main__':
    main()

# vim:sw=4:ts=4:et:
