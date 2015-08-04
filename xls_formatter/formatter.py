__author__ = 'piotr'
import os
import xlwt
from xlwt.Style import default_style
from django.utils.encoding import force_str
from django.utils.formats import number_format
from django.http import HttpResponse


class XlsResponseMixin(object):
    """
    Inheriting classes must implement get_object().
    The object returned must have property tabular_result and optional table_header and table_preface.
    """

    def as_xls(self):
        formatter = XlsFormatter(
            self.object.tabular_result,
            t_header=getattr(self.object, 'table_header', None),
            preface=getattr(self.object, 'table_preface', None)
        )
        return formatter.http_response()


class XlsFormatter(object):
    style_table_header = xlwt.easyxf('font: bold on; align: wrap on; pattern: pattern solid, fore_colour light_turquoise;')
    style_boldheader = xlwt.easyxf('font: bold on; align: wrap on;')
    DEFAULT_SHEET = 'Cover'
    MAX_SHEET_NAME_LEN = 30
    MAX_CELL_LEN = 13000
    compiled_result = None

    def make_str(self, s):
        if hasattr(s, 'as_integer_ratio'):
            return force_str(number_format(s, 2))[:self.MAX_CELL_LEN] + u'%'
        return force_str(s)[:self.MAX_CELL_LEN]

    def non_empty(self, v):
        return v not in [u'', None]

    def dataset_len(self, d):
        dl = 0
        for dx in d:
            dl = max(dl, len([ix for ix in dx if self.non_empty(ix)]))
        return dl

    def is_bold(self, potentiallyboldrow):
        if getattr(potentiallyboldrow, 'is_bold', False):
            return self.style_boldheader
        return default_style

    def handle_row(self, ws, y, row):
        x = 0
        for cell in row:
            ws.write(y, x, self.make_str(cell), self.is_bold(row))
            x += 1
        return y + 1

    def __init__(self, t_body, preface=None, t_header=None, logo_path=None):
        """
        Generate a spreadsheet with a single sheet, containing single table.
        Table header cells get content from optional t_header.
        Table rows cells get contents from t_body.
        If logo_path is present it will be placed before preface, which if present will
        be placed before the table.

        To add more workbooks...

        :preface: an iterable of iterables to place as text before the table
        :t_header: an iterable of table headers
        :t_body: an iterable of iterables representing rows of cells
        :logo_path: path to a 24b rgb bitmap file, relative to current app
        """
        self.style_table_header.borders = xlwt.Borders()
        self.style_table_header.borders.left = xlwt.Borders.MEDIUM
        self.style_table_header.borders.right = xlwt.Borders.MEDIUM
        self.style_table_header.borders.top = xlwt.Borders.MEDIUM
        self.style_table_header.borders.bottom = xlwt.Borders.MEDIUM
        self.wb = xlwt.Workbook(encoding='utf-8')
        cover = self.wb.add_sheet(self.DEFAULT_SHEET)

        y = 0
        if logo_path:
            logofilepath = os.path.join(os.path.abspath(os.path.dirname(__file__)), logo_path)
            cover.insert_bitmap(logofilepath, 0, 0)
            y = 3

        self.write_sheet(cover, t_body, t_header, preface=preface, y=y)

    def add_sheet(self, sheet_name):
        return self.wb.add_sheet(sheet_name[:self.MAX_SHEET_NAME_LEN])

    def write_sheet(self, sheet, t_body, t_header=None, preface=None, y=0):
        x = 0
        if preface:
            for row in preface:
                y = self.handle_row(sheet, y, row)
            y += 3

        if t_header:
            x = 0
            for cell in t_header:
                sheet.write(y, x, self.make_str(cell), self.style_table_header)
                x += 1
            y += 1

        for row in t_body:
            y = self.handle_row(sheet, y, row)
        # return current cursor position for easier
        return y

    def http_response(self):
        resp = HttpResponse(content_type='octet/stream')
        resp['Content-Disposition'] = 'attachment; filename="report.xls"'
        self.wb.save(resp)
        return resp
