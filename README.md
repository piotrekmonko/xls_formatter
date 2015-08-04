xls_formatter
=============

Util for quick XLS responses in views.

Requires django and xlwt.

Usage::


    class BitterButter(DetailView):

        def get(self, request, *args, **kwargs):
            obj = self.get_object()
            formatter = XlsFormatter(
                obj.table_contents,
                t_header=obj.table_header,
                preface=(
                    (u'Date', now()),
                    (u'Author', u'Me'),
                )
            )
            return formatter.http_response()


obj.table_contents must be an iterable of tuples representing rows and cells;
each row must have even number of cells.

obj.table_header (optional) must be an iterable of cells to prepend before obj.table_contents;
number of cells must match obj.table_contents[0] tuple length.

preface (optional) must be an iterable of tuples of any-length content to put before the data in obj.table_contents