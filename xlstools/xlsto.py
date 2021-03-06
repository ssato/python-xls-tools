#
# xlsto.py - Convert Excel files to other formats, e.g. csv, sql db.
#
# Copyright (C) 2008 - 2012 Satoru SATOH <satoru.satoh at gmail.com>
# License: MIT
#
import xlstools.csvx as XC
import xlstools.utils as U

import copy
import logging
import optparse
import os
import sqlite3
import sys
import xlrd

try:
    import json
except ImportError:
    import simplejson as json


"""xlsto.py

Usage:
 
  ./xlsto.py [Options ...] DATASPEC_FILE

  (-h or --help option will shows usage)


Data spec:

data spec format is JSON. Here is an example.

[
  {
    "location": "http://example.com/pub/xls/",  # optional
    "filepath": "SampleData.xls",  # file path  
    "description": "Sample Data List",
    "sheets": [
      {
        "name": "Sheet1",
        "table_name": "sample_data_list",  # database table name will be used in sql db.
        "description": "Sample Data List",
        "keys": [[3,0],[3,1],[3,2],[3,3],"test_key0"],  # cell or string list of col names.
                                                        # 
        "marker_idx": 0,   # The index of the col to check whether data exists in row or not.
                           # If omitted, 0 will be used.
        "data_range": [[4,-1],[0,12]]  # data range [[row_bein, row_end], [col_bein, col_end]].
                                       # Indices start with 0 and -1 indicates infinite, 
                                       # that is, will be detected automatically.
      },
      ...
    ]
  },
  ...
]
"""


def load_specs(specfile):
    """Loads given data spec and returns it as an internal representation.
    See the spec example above also.
    """
    return json.load(open(specfile, 'r'))


def load_datasets(specfile, filepath):
    """Loads datasets from Excel workbooks (files) according to each file spec
    in specfile and returns these as dict objects.
    """
    for filespec in load_specs(specfile):
        book = xlrd.open_workbook(filepath)  # might throw IndexError, IOError, etc.

        for sheet_idx in range(0, len(filespec['sheets'])):
            sheet = book.sheet_by_index(sheet_idx)
            sheetspec = filespec['sheets'][sheet_idx]
            dataset = copy.copy(sheetspec)

            midx = filespec['sheets'][sheet_idx].get('marker_idx', 0)
            rows,cols = sheetspec['data_range']
            if rows[1] == -1:
                rows[1] = sheet.nrows

            # TODO: exceptions handling. (IndexError, etc.)
            keys = [(isinstance(c, list) and sheet.cell_value(*c) or c) for c in sheetspec['keys']]
            values = [sheet.row(rx)[cols[0]:cols[1]+1] for rx in range(*rows) if sheet.row(rx)[midx].value]

            dataset['keynames'] = [U.normalize_key(k) for k in keys]
            dataset['values'] = values

            yield dataset


# CSV related:
def csv_process_dataset(outdir, dataset, force):
    """Create db tables and insert datasets into it.
    """
    outfile = dataset['table_name']
    keynames = dataset['keynames']
    values = dataset['values']

    outfile = os.path.sep.join((outdir, outfile + '.csv'))
    logging.info("creating csv file '%s'" % outfile)
    if force:
        U.rename_if_exists(outfile)

    writer = XC.UnicodeWriter(open(outfile, 'wb'))
    writer.writerow(keynames)

    for xs in values:
        writer.writerow([(x.value and x.value or "") for x in xs])


def csv_create(specfile, filepath, outdir, force):
    """Create the database (create tables and insert datasets into it).
    """
    if not os.path.exists(outdir):
        os.mkdir(outdir, 0755)

    for dataset in load_datasets(specfile, filepath):
        csv_process_dataset(outdir, dataset, force)


# SQLite DB related:
def db_process_dataset(dbfile, dataset):
    """Create db tables and insert datasets into it.
    """
    conn = sqlite3.connect(dbfile)

    table = dataset['table_name']
    keynames = dataset['keynames']
    values = dataset['values']

    keys = ', '.join(keynames).replace('?','')
    placeholders = ', '.join('?' * len(keynames))

    # 1. create table:
    #sql = "create table %s (%s) if not exists" % (table, keys)
    sql = "create table %s (%s)" % (table, keys)
    logging.info("sql = '%s'" % sql)
    conn.execute(sql)
    conn.commit()

    # 2. insert dataset into the table:
    sql = "insert or replace into %s (%s) values (%s)" % (table, keys, placeholders)
    logging.info("sql = '%s'" % sql)

    for xs in values:
        vs = [(x.value and x.value or "") for x in xs]
        if vs:
            logging.info("value = '%s'" % vs)
            conn.execute(sql, vs)
    conn.commit()

    conn.close()


def db_create(specfile, filepath, dbfile, force):
    """Create the database (create tables and insert datasets into it).
    """
    logging.info("creating db '%s'" % dbfile)
    if force:
        U.rename_if_exists(dbfile)
    for dataset in load_datasets(specfile, filepath):
        db_process_dataset(dbfile, dataset)


def opts_parser():
    parser = optparse.OptionParser('%prog [OPTION ...] INPUT_FILE')
    parser.add_option('-s', '--spec', help='specify "spec" file defines XLS data structure [guessed from input file]')
    parser.add_option('-o', '--output', dest='output', default='output',
        help='specify database file for "sqlite" output or dir for "csv" output. [default: output.db or output/]')
    parser.add_option('-t', '--output-type', dest='type',
        help='Specify the output type, csv or sqlite [default].')
    parser.add_option('-f', '--force', dest='force', action='store_true',
        help='Force overwrite existing file/dir.', default=False)
    parser.add_option('-v', '--verbose', dest='verbose', action='store_true',
        help='Verbose mode.', default=False)
    return parser


def main():
    parser = opts_parser()
    (options, args) = parser.parse_args()

    if options.verbose:
        logging.basicConfig(level=logging.INFO)

    out = options.output

    if options.type == 'csv':
        create_f = csv_create
    else:
        create_f = db_create
        if not out.endswith('.db'):
            out = out + '.db'

    if len(args) < 1:
        parser.print_help()
        sys.exit(-1)

    filepath = args[0]

    if options.spec:
        specfile = options.spec
    else:
        specfile = filepath[:filepath.rfind('.')] + '.spec'

    if not os.path.exists(specfile):
        print >> sys.stderr, "Spec file '%s' does not exists!" % specfile
        sys.exit(-1)

    if not os.path.exists(filepath):
        print >> sys.stderr, "Input file '%s' does not exists!" % filepath
        sys.exit(-1)

    create_f(specfile, filepath, out, options.force)


if __name__ == '__main__':
    main()


# vim:sw=4:ts=4:et:
