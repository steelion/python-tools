__author__ = 'Ofner Mario'



class ReportingColumn():

    def __init__(self, header_description = None, header_font = None, header_data_format = None, header_text_format = None,
        cell_font = None, column_data_format = None, column_text_format = None, column_width = None ):
        self.header_description = header_description
        self.header_font = header_font
        self.header_data_format = header_data_format
        self.header_text_format = header_text_format
        self.cell_font = cell_font
        self.column_data_format = column_data_format
        self.column_text_format = column_text_format
        self.column_width = column_width


    def get_header_text_format(self):
        return self._header_text_format

    def set_header_text_format(self, value):
        if value is None:
            self._header_text_format = None
            return

        if isinstance(value,list):
            self._header_text_format = value
        elif isinstance(value,str):
            if len(value.strip()) < 1:
                self._header_text_format = None
            else:
                self._header_text_format = str(value).split(';')
                for i in range(0,len(self._header_text_format)):
                    self._header_text_format[i] = str(self._header_text_format[i]).strip()
        else:
            raise AttributeError()

    header_text_format = property(fset=set_header_text_format, fget=get_header_text_format)


    def get_header_data_format(self):
        return self._header_data_format

    def set_header_data_format(self, value):
        default_header_data_format = "General"

        if value is None:
            self._header_data_format = default_header_data_format
            return

        if isinstance(value,str):
            if len(value.strip()) < 1:
                self._header_data_format = default_header_data_format
        else:
            raise AttributeError()

    header_data_format = property(fset=set_header_data_format, fget=get_header_data_format)


    def get_column_data_format(self):
        return self._column_data_format

    def set_column_data_format(self, value):
        default_column_data_format = "General"

        if value is None:
            self._column_data_format = default_column_data_format
            return

        if isinstance(value,str):
            if len(value.strip()) < 1:
                self._column_data_format = default_column_data_format
            else:
                self._column_data_format = value
        else:
            raise AttributeError()

    column_data_format = property(fset=set_column_data_format, fget=get_column_data_format)

