#!/usr/bin/env python
# -*- coding: utf-8 -*-
from gluon.sqlhtml import ExportClass
from gluon.dal import Row
import cStringIO
from pyheaderfile import Xls, Xlsx

class ExporterPyheaderfile(ExportClass):
    """ Default class to handle xlsx and xls files using pyheaderfile. This
    class don't work alone, need to be called by ExporterXls or ExporterXlsx.
    """
    def __init__(self, rows):
        ExportClass.__init__(self, rows)

    def export(self):  #export Xls and Xlsx with rows.label
        if self.rows:
            s = cStringIO.StringIO()

            db = self.rows.db
            labels = dict()
            for name in self.rows.colnames:
                if name.index('.'):
                    table, column = name.split('.')
                    index = '%s.%s' % (table, column)
                    labels[index] = db[table][column].label.decode('utf8')
                else:
                    column = name
                    index = '%s' % column
                    labels[index] = db[column].label.decode('utf8')


            # instanciate pyheaderfile
            phf = self.phf_class(s, labels.values())

            for row in self.rows:
                add_dict = dict()
                # I hate web2py for it... Ugly boy

                for table in row:
                    # if there is just one table
                    if not isinstance(row[table], Row):
                        column = table
                        table = self.rows.colnames[0].split('.')[0]
                        value = row[column]
                        if isinstance(value, unicode) or isinstance(value, str):
                            value = value.decode('utf8')
                        else:
                            value = str(value)
                        add_dict[labels['%s.%s' % (table, column)]] = value

                    else:
                        for column in row[table]:
                            value = row[table][column]
                            if isinstance(value, unicode) or isinstance(value,
                                                                        str):
                                value = value.decode('utf8')
                            else:
                                value = str(value)
                            add_dict[labels['%s.%s' % (table, column)]] = value

                phf.write(**add_dict)
            return phf.save()
        else:
            return None

class ExporterXls(ExporterPyheaderfile):
    """ Exporter class to be used in grids. Add to exportclasses something like
    this:

    xls=(ExporterXls, 'Excel (XLS)', T('Export file to Xls (Excel)'))

    This class export xls file.
    """
    label = 'Excel (XLS)'
    file_ext = "xls"
    content_type = "application/vnd.ms-excel"

    def __init__(self, rows):
        ExporterPyheaderfile.__init__(self, rows)
        self.phf_class = Xls

class ExporterXlsx(ExporterPyheaderfile):
    """ Exporter class to be used in grids. Add to exportclasses something like
    this:

    xlsx=(ExporterXls, 'New Excel (XLSX)', T('Export file to Xlsx (Excel)'))

    This class export xlsx file.
    """
    label = 'Excel (XLSX)'
    file_ext = "xlsx"
    content_type = "application/vnd.openxmlformats-" \
                   "officedocument.spreadsheetml.sheet"

    def __init__(self, rows):
        ExporterPyheaderfile.__init__(self, rows)
        self.phf_class = Xlsx
