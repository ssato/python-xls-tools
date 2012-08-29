#
# Generate an Excel (.xls) file from multiple CSV files.
#
# Copyright (C) 2010 - 2012 Satoru SATOH <satoru.satoh at gmail.com>
# License: MIT
#
"""CSV file format:

label_0, label_1, label_2, ....   => Headers
value_0, value_1, ...             => Dataset
...
"""
import xlstools.csvworkbook as CW

import logging
import optparse
import os.path
import sys


def opts_parser():
    defaults = dict(
        encoding="utf-8", verbose=False, quiet=False,
        vmerge=False, vmerge_col_end=-1,
        auto_col_width=False, header_style="font: name Times New Roman, bold on",
        main_style="font: name Times New Roman",
        merged_style="align: wrap yes, vert center",
    )

    p = optparse.OptionParser(
        """%prog [OPTION ...] CSVFILE_0 [CSVFILE_1 ...] OUTPUT_XLS

Examples:
  %prog aaa.csv bbb.csv ccc.csv ABC.xls
  %prog - output.xls  # read csv data from stdin
  %prog --main-style 'font: name IPAPGothic' --sheet-names "aaa,bbb" A.csv B.csv AB.xls
    """)
    p.set_defaults(**defaults)

    cog = optparse.OptionGroup(p, "Common Options")
    cog.add_option('', '--sheet-names', help='Comma separated worksheet names')
    cog.add_option('-E', '--encoding',
        help='Character set encoding of the CSV files [utf-8]'
    )
    cog.add_option('-v', '--verbose', help='Verbose mode', action="store_true")
    cog.add_option('-q', '--quiet', help='Quiet mode', action="store_true")
    p.add_option_group(cog)

    mog = optparse.OptionGroup(p, "Cell-merging Options")
    mog.add_option('', '--vmerge', action="store_true",
        help='Automatically merge cells having same value'
    )
    mog.add_option('', '--vmerge-col-end', type="int",
        help='Specify the idx of the end column to be merged'
    )
    p.add_option_group(mog)

    sog = optparse.OptionGroup(p, "Style Options")
    sog.add_option('', '--auto-col-width', help='Automatically adjust column widths')
    sog.add_option('', '--header-style', 
        help="Main (content) style. See xlwt's document also, "
            "https://secure.simplistix.co.uk/svn/xlwt/trunk/README.html. "
            "[%default]"
    )
    sog.add_option('', '--main-style', 
        help="Main (content) style, ex. 'font: name IPAPGothic' for "
            "Japanese texts [%default]"
    )
    sog.add_option('', '--merged-style', 
        help="Merged cells' style if --vmerge option is used, "
            "ex. 'vert center', 'horiz center' [%default]")
    p.add_option_group(sog)

    return p


def main():
    loglevel = logging.WARN
    sheet_names = {}

    p = opts_parser()
    (options, args) = p.parse_args()

    if options.verbose:
        loglevel = logging.INFO

    if options.quiet:
        loglevel = logging.ERROR

    logging.basicConfig(level=loglevel)

    if len(args) < 2:
        p.print_help()
        sys.exit(0)

    csvfiles = args[0:-1]
    output = args[-1]

    if options.sheet_names:
        sheet_names = dict(zip(csvfiles, options.sheet_names.split(',')))

    wb = CW.CsvsWorkbook(output, options.header_style, options.main_style)
    for csvf in csvfiles:
        title = sheet_names.get(csvf, os.path.basename(csvf).replace('.csv',''))
        wb.addWorksheetFromCSVFile(
            csvf, csv_encoding=options.encoding, title=title,
            main_style=options.main_style, header_style=options.header_style,
            auto_col_width=options.auto_col_width, 
            vmerge=options.vmerge, vmerge_col_end=options.vmerge_col_end,
            merged_style=options.merged_style,
        )

    wb.save()


if __name__ == '__main__':
    main()

# vim:sw=4:ts=4:et:
