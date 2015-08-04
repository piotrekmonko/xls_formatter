#! /usr/bin/env python
from distutils.core import setup
import sys

reload(sys).setdefaultencoding('Utf-8')

setup(
    name='xls_formatter',
    version='0.1',
    def as_xls(self):
        frmtr = XlsFormatter(
            self.object.result,
            t_header=self.object.metrics,
            preface=(
                (u'Nazwa', self.object.cls),
                (u'Start', self.object.start),
                (u'End', self.object.end),
            )
        )
        return frmtr.http_response()


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

        # if preface:
        #     for row in preface:
        #         x = 0
        #         y += 1
        #         for cell in row:
        #             cover.write(y, x, self.make_str(cell))
        #             x += 1

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

    def as_xls(self):
        frmtr = XlsFormatter(
            self.object.result,
            t_header=self.object.metrics,
            preface=(
                (u'Nazwa', self.object.cls),
                (u'Start', self.object.start),
                (u'End', self.object.end),
            )
        )
        return frmtr.http_response()


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

        # if preface:
        #     for row in preface:
        #         x = 0
        #         y += 1
        #         for cell in row:
        #             cover.write(y, x, self.make_str(cell))
        #             x += 1

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

    author='Piotrek Mońko',
    author_email='piotrek.monko@gmail.com',
    description='Utils for quick XLS responses',
    long_description=open('README.md').read(),
    url='https://github.com/piotrekmonko/xls_formatter',
    license='BSD License',
    platforms=['OS Independent'],
    packages=['xls_formatter'],
    include_package_data=True,
    classifiers=[
        'Development Status :: 0.1 - Beta',
        'Environment :: Web Environment',
        'Framework :: Django',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
)
