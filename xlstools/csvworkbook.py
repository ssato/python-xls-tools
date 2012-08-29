#
# Copyright (C) 2010 - 2012 Satoru SATOH <satoru.satoh at gmail.com>
# License: MIT
#
from xlstools import adjust_width, max_col_widths, mergeable_cells

import csv
import logging
import sys
import xlwt


class CsvsWorkbook(object):

    default_styles = dict(
        header="font: name Times New Roman, bold on",
        main="font: name Times New Roman",
        merged="align: wrap yes, vert center, horiz center",
    )

    def __init__(self, filename, header_style=False, main_style=False,
            merged_style=False):
        self._filename = filename
        self._workbook = xlwt.Workbook()
        self._sheets = 0
        self.__init_styles(header_style, main_style, merged_style)

    def __init_styles(self, header_style, main_style, merged_style):
        for (sn, ss) in (('header', header_style), ('main', main_style),
                ('merged', merged_style)):
            self.__init_style(sn, ss)

    def __init_style(self, style_name, style):
        if not style:
            style = self.default_styles[style_name]

        setattr(self, "_%s_style" % style_name, \
            self.__to_style(style_name, style))

    def __del__(self):
        self._workbook.save(self._filename)

    def __to_style(self, style_name, style_string):
        try:
            style = xlwt.easyxf(style_string)
            assert isinstance(style, xlwt.Style.XFStyle)
        except:
            logging.warn(
                "Given style '%s'[%s] is not valid. " + \
                "Fall backed to the default." % (style_string, style_name)
            )
            ss = self.default_styles.get(
                style_name, self.default_styles['main']
            )
            style = xlwt.easyxf(ss)

        return style

    def save(self):
        self._workbook.save(self._filename)

    def header_style(self):
        return self._header_style

    def main_style(self):
        return self._main_style

    def merged_style(self):
        return self._merged_style

    def sheets(self):
        return self._sheets + 1

    def addWorksheetFromCSVFile(self, csv_filename, csv_encoding='utf-8',
            title=False, fieldnames=[], header_style=False,
            main_style=False, auto_col_width=False,
            vmerge=False, vmerge_col_end=-1, merged_style=False):
        if not title:
            title = "Sheet %d" % (self.sheets())

        _conv = lambda x: unicode(x, csv_encoding)

        if header_style:
            hstyle = self.__to_style('header', header_style)
        else:
            hstyle = self.header_style()

        if main_style:
            mstyle = self.__to_style('main', main_style)
        else:
            mstyle = self.main_style()

        if merged_style:
            mgstyle = self.__to_style('merged', merged_style)
        else:
            mgstyle = self.merged_style()

        if csv_filename == "-":
            csvf = sys.stdin
        else:
            csvf = open(csv_filename) 
            
        reader = csv.reader(csvf)
        cells = [row for row in reader]
        csvf.close()

        (headers, dataset) = (cells[0], cells)

        worksheet = self._workbook.add_sheet(title)
        self._sheets += 1

        # header fields: given or get from the first line of the csv file.
        if not fieldnames or (fieldnames and len(fieldnames) < len(headers)):
            fieldnames = headers

        for col in range(0, len(fieldnames)):
            logging.info(" col=%d, fieldname=%s" % (col, fieldnames[col]))
            worksheet.write(0, col, _conv(fieldnames[col]), hstyle)

        # @FIXME: Tune factor and threashold values.
        if auto_col_width:
            mcws = max_col_widths(dataset[1:])  # ignore header columns.

            for i in range(0, len(dataset[0])):
                w = adjust_width(mcws[i])
                logging.info(" col[%d].width=%d [%d](adjusted [original])" % (i, w, mcws[i]))
                worksheet.col(i).width = w

        # main data
        rows = len(dataset)

        if vmerge:
            for ms in mergeable_cells(dataset, 1, col_end=vmerge_col_end):
                worksheet.write_merge(ms[1], ms[2], ms[3], ms[4], _conv(ms[0]), mgstyle)

        for row in range(1, rows):
            for col in range(0, len(dataset[row])):
                logging.info(" row=%d, col=%d, data=%s" % (row, col, dataset[row][col]))
                if vmerge:
                    try:
                        worksheet.write(row, col, _conv(dataset[row][col]) or "", mstyle)
                    except:
                        logging.info("The cell (row=%d, col=%d) is a part of merged cells." % (row, col))
                        pass   # skip this cell as it was written as merged cells before.
                else:
                    worksheet.write(row, col, _conv(dataset[row][col]) or "", mstyle)

# vim:sw=4:ts=4:et:
