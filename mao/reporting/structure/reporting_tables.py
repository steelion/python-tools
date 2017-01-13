__author__ = 'Ofner Mario'

from mao.reporting.structure.reporting_columns import ReportingColumn


class ReportingTable():

    # Key Column
    # Key Name and 0based ColumnIndex
    kc_name = None
    kc_index = None

    def __init__(self, column_list = None, key_column = None, header_rowheight = None, data_rowheight = None):
        self.column_list = column_list
        self.key_column = key_column
        self.header_rowheight = header_rowheight
        self.data_rowheight = data_rowheight


    def get_key_column(self):
        if self.kc_index is None:
            return None
        else:
            return str("Name: '"+self.kc_name+"' Index: "+str(self.kc_index))

    def set_key_column(self, value):
        if value is None:
            self._key_column = None
            self.kc_name = None
            self.kc_index = None
            return

        if isinstance(value,int):
            if self.column_list is None or value > len(self.column_list):
                # Todo -> Index out of Bounds Error
                raise AttributeError()
            self.kc_index = value
            self.kc_name = self.column_list[value].header_description

        elif isinstance(value,str):
            if self.column_list is None or value not in self.table_header:
                # Todo -> Column not found Error
                raise AttributeError()
            # self.kc_index = self.column_list.index(value)
            self.kc_index = list(x.header_description for x in self.column_list).index(value)
            self.kc_name = value

        else:
            raise AttributeError()


    def get_column_list(self):
        return self._column_list

    def set_column_list(self, value):
        if not all(isinstance(col_entry, ReportingColumn) for col_entry in value):
            raise AttributeError
        else:
            self._column_list = value

    def get_table_header(self):
        return list(x.header_description for x in self.column_list)

    def get_column_widths(self):
        return list(x.column_width for x in self.column_list)

    def get_header_data_format(self):
        return list(x.header_data_format for x in self.column_list)

    def get_column_data_format(self):
        return list(x.column_data_format for x in self.column_list)

    def get_header_text_format(self):
        return list(x.header_text_format for x in self.column_list)

    def get_column_text_format(self):
        return list(x.column_text_format for x in self.column_list)


    key_column = property(fget=get_key_column, fset=set_key_column)
    column_list = property(fget=get_column_list, fset=set_column_list)
    table_header = property(fget=get_table_header)
    column_widths = property(fget=get_column_widths)

    header_data_format = property(fget=get_header_data_format)
    column_data_format = property(fget=get_column_data_format)
    header_text_format = property(fget=get_header_text_format)
    column_text_format = property(fget=get_column_text_format)


