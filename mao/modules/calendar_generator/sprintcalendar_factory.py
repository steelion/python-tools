__author__ = 'Ofner Mario'


from datetime import datetime
import pprint

from logging import getLogger
from sprintcalendar_generator import *


logger = getLogger(__name__)

def hex2dec(hex_val):
    return int(hex_val, 16)

def dec2hex(dec_val):
    # return hex(dec_val)
    return "%X" % dec_val


class SprintCalendarMonthColumnRenderer():

    file_reader_writer_factory = None

    # Report Structure Definitions
    title_font_default = {"name":'Ubuntu Light', "size":15, "bold":True, "italic":False, "vertAlign": None, "underline":'none', "strike":False, "color":'000000'}
    header_font_default = {"name":'Ubuntu Light', "size":12, "bold":True, "italic":False, "vertAlign":None, "underline":'none', "strike":False, "color":'000000'}
    header_font_small = {"name":'Ubuntu Light', "size":8, "bold":False, "italic":False, "vertAlign":None, "underline":'none', "strike":False, "color":'000000'}
    cell_font_default = {"name":'Ubuntu Light', "size":8, "bold":False, "italic":False, "vertAlign":None, "underline":'none', "strike":False, "color":'000000'}
    footer_font_default = {"name":'Ubuntu Light', "size":8, "bold":False, "italic":False, "vertAlign":None, "underline":'none', "strike":False, "color":'000000'}

    cell_font_weekno = {"name":'Ubuntu Light', "size":6, "bold":False, "italic":False, "vertAlign":None, "underline":'none', "strike":False, "color":'000000'}
    cell_font_holiday = {"name":'Ubuntu Light', "size":6, "bold":False, "italic":False, "vertAlign":None, "underline":'none', "strike":False, "color":'FF0000'}
    cell_font_special = {"name":'Ubuntu Light', "size":8, "bold":True, "italic":False, "vertAlign":None, "underline":'none', "strike":False, "color":'000000'}

    bgcolors_template = {
        'blue': {
            'light': [219, 229, 241],
            'dark': [149, 179, 215]
        },
        'red': {
            'light': [242, 221, 220],
            'dark': [217, 151, 149]
        },
        'green': {
            'light': [234, 241, 221],
            'dark': [194, 214, 154]
        },
        'purple': {
            'light': [229, 224, 236],
            'dark': [178, 161, 199]
        },
        'orange': {
            'light': [253, 233, 217],
            'dark': [250, 192, 144]
        }
    }

    bg_colors = {}

    title_rowheight = 45
    footer_rowheight = 10.5
    month_name_rowheight = 18
    month_no_rowheight = 10.5
    data_rowheight = 12.75
    day_rowheight = 18
    release_color_index = 0
    release_color = None


    def __init__(self, _file_reader_writer_factory):
        self.file_reader_writer_factory = _file_reader_writer_factory
        self.calendar_orientation = 'column'
        self.month_width = 4
        self.day_height = 1
        self.month_height = None
        self.month_begin = 'day' # day: firstDayOfMonth; week: Start with Monday
        self.show_month_numbers = True
        self.show_week_numbers = True


    def load_workbook(self, path_to_file, filename):
        # logger.debug("path_to_file: "+ path_to_file + filename)
        self.file_reader_writer_factory.load_workbook( _filename = path_to_file + filename )

    def change_or_create_worksheet(self, worksheet_name):
        if not self.file_reader_writer_factory.change_worksheet(worksheet_name):
            self.file_reader_writer_factory.create_worksheet(worksheet_name)
        return True

    def remove_worksheet(self, worksheet_name):
        return self.file_reader_writer_factory.remove_worksheet(worksheet_name)

    def save_workbook(self):
        self.file_reader_writer_factory.save_workbook()


    # --------------------------------------------------------------------------------
    # - Begin of helpers
    @property
    def sprintcalendar(self):
        return self._sprintcalendar

    @sprintcalendar.setter
    def sprintcalendar(self, calendar):
        if isinstance(calendar, SprintCalendar):
            self._sprintcalendar = calendar
            return True
        return False


    def generate_bg_colors(self, color_list, alpha_count):
        """generates different bgcolors for the given amount of different alpha values"""
        logger.debug(color_list)

        for c in color_list:

            logger.debug(c)

            light_color = self.bgcolors_template[c]['light']
            dark_color = self.bgcolors_template[c]['dark']
            self.bg_colors[c] = []
            logger.debug(str(dec2hex(light_color[0]))+str(dec2hex(light_color[1]))+str(dec2hex(light_color[2])))

            self.bg_colors[c].append(str(dec2hex(light_color[0]))+str(dec2hex(light_color[1]))+str(dec2hex(light_color[2])))

            if alpha_count == 1:
                continue

            if alpha_count > 2:
                tmp_color_shades = []

                for i in range(0, 3):
                    delta = int(light_color[i] - dark_color[i])
                    step = int(delta/(alpha_count-1))
                    logger.debug("delta: %s, step: %s", delta, step)
                    for k in range(0, alpha_count-2):
                        if k >= len(tmp_color_shades):
                            tmp_color_shades.append([])
                        tmp_color_shades[k].append(int(light_color[i] - (step * (k+1))))

                logger.debug(light_color)
                logger.debug(tmp_color_shades)
                logger.debug(dark_color)

                for i in range(0, len(tmp_color_shades)):
                    self.bg_colors[c].append(str(dec2hex(tmp_color_shades[i][0]))+str(dec2hex(tmp_color_shades[i][1]))+str(dec2hex(tmp_color_shades[i][2])))

            self.bg_colors[c].append(str(dec2hex(dark_color[0]))+str(dec2hex(dark_color[1]))+str(dec2hex(dark_color[2])))
            continue
        self.bg_colors['color_list'] = color_list

        self.next_release_color()

    def next_release_color(self):

        release_color_list = self.bg_colors['color_list']

        if self.release_color is None:
            self.release_color_index = 0
        else:
            self.release_color_index += 1
            if self.release_color_index == len(release_color_list):
                self.release_color_index = 0

        self.release_color = release_color_list[self.release_color_index]



    # --------------------------------------------------------------------------------
    # - Begin of helpers
    def get_month_column_width(self):
        column_width_list = [1.8, 2, 6.6, 1.4]
        column_width_list = list(x+0.7109375 for x in column_width_list)
        return column_width_list

    def get_column_widths(self):
        column_width_list = []
        for i in range(0,15):
            column_width_list += self.get_month_column_width()

        return column_width_list

    def get_title_text_format(self):
        return None

    def get_header_text_format(self):
        return None


    # --------------------------------------------------------------------------------
    # - Write defined Table Header initially to output File
    # - The Table Header is defined in this Class
    def render_title(self):

        # table_header = self.get_table_header()

        # # put the header in first row
        # rownum = 1
        # for col in range(1, len(table_header) + 1):
        #     _excel_factory.set_cell_value(rownum, col, table_header[col-1])
        # rownum += 1
        self.file_reader_writer_factory.set_row_values(1, ['CROSS 2 - Release/Sprints 2017'])
        self.file_reader_writer_factory.set_column_width(self.get_column_widths())

        self.file_reader_writer_factory.apply_styles_to_row(1,
            font_description=self.title_font_default,
            dimension_row_height=self.title_rowheight
            )
        self.file_reader_writer_factory.set_cell_text_style(1, 1, "v_top")

    # --------------------------------------------------------------------------------
    # - Write defined Table Header initially to output File
    # - The Table Header is defined in this Class
    def render_footer(self):
        row = self.file_reader_writer_factory.get_last_row() + 1

        self.file_reader_writer_factory.set_row_values(row, ['[Calendar created by Mario Ofner]'])
        self.file_reader_writer_factory.apply_styles_to_row(row,
            font_description=self.footer_font_default,
            dimension_row_height=self.footer_rowheight
            )
        self.file_reader_writer_factory.set_cell_text_style(row, 1, "v_center")


    def render_calendar(self):
        """
        """

        logger.debug("Rendering Title")
        self.render_title()

        logger.debug("Validating Sprintcalendar")
        if self.sprintcalendar is None:
            return false

        logger.debug("Rendering Calendar Header")
        row = 2
        col = 1
        for calendar_month in self.sprintcalendar.month_list:
            self.render_month_header(row_offset=row, col_offset=col, calendar_month=calendar_month)
            self.file_reader_writer_factory.draw_table_border( min_col=col, min_row=row, max_col=(col+self.month_width-1), max_row=(row+1 if self.show_month_numbers else row), border_only = True )
            col += self.month_width

        logger.debug("Set Calendar Header Style")
        self.file_reader_writer_factory.apply_styles_to_row(row, font_description=self.header_font_default, dimension_row_height=self.month_name_rowheight)

        logger.debug("Set Calendar Column Widths")
        self.file_reader_writer_factory.set_column_width(self.get_column_widths())

        if self.show_month_numbers:
            self.file_reader_writer_factory.apply_styles_to_row(row+1, font_description=self.header_font_small, dimension_row_height=self.month_no_rowheight)

        logger.debug("Creating Backgroundcolor List")
        color_list = ['blue', 'green', 'purple', 'orange']
        self.generate_bg_colors(color_list, 3)

        logger.debug("Rendering Month Columns")
        calendar_top_row = 4
        col = None
        self.release_color_index = 0

        for calendar_month in self.sprintcalendar.month_list:
            logger.debug("Rendering days for month: '" + calendar_month.name + "'")
            row = calendar_top_row
            if col is None:
                col = 1
            else:
                col += self.month_width
            for calendar_day in calendar_month.day_list:
                self.render_day(row_offset=row, col_offset=col, calendar_day=calendar_day)
                row += self.day_height

            self.file_reader_writer_factory.draw_table_border( min_col=col, min_row=calendar_top_row, max_col=(col+self.month_width-1), max_row=row-1, border_only = True )


        logger.debug("Adjust height of Rows")
        for r in range(0, 31):
            self.file_reader_writer_factory.apply_styles_to_row(calendar_top_row+r, dimension_row_height=self.day_rowheight)

        logger.debug("Rendering Footer")
        self.render_footer()


    def render_month_header(self, row_offset, col_offset, calendar_month):
        start_column = col_offset
        end_column = start_column + (self.month_width - 1)
        month_name = calendar_month.name
        month_number = calendar_month.month

        self.file_reader_writer_factory.merge_cells(start_row=row_offset, start_column=start_column, end_row=row_offset, end_column=end_column)
        self.file_reader_writer_factory.set_cell_value(row_offset, start_column, month_name)
        self.file_reader_writer_factory.set_cell_text_style(row_offset, start_column, "v_center;h_center")

        if self.show_month_numbers:
            self.file_reader_writer_factory.merge_cells(start_row=row_offset+1, start_column=start_column, end_row=row_offset+1, end_column=end_column)
            self.file_reader_writer_factory.set_cell_value(row_offset+1, start_column, month_number)
            self.file_reader_writer_factory.set_cell_text_style(row_offset+1, start_column, "v_center;h_center")

    def render_day(self, row_offset, col_offset, calendar_day):
        """
        +-------------------------+
        |        month_name       |
        +-------------------------+
        |       month_number      |
        +----+----+----------+----+
        |dnr#| ds |          |wnr#| <= if Monday
        +----+----+----------+----+
        |dnr#| ds | dname    |wnr#| <= if Monday and special day
        +----+----+----------+----+
        |dnr#| ds | dname         | <= if Special day
        +----+----+---------------+
        |dnr#| ds |               |
        +----+----+---------------+
        """

        day_number = calendar_day.day_of_month
        day_name = calendar_day.name
        day_abbr = calendar_day.abbr

        start_column = col_offset+2
        end_column = col_offset+3

        cell_font = self.cell_font_default.copy()
        cell_font_weekno = self.cell_font_weekno.copy()
        cell_font_special = self.cell_font_holiday.copy()
        cell_bg_color = 'FFFFFF'

        self.file_reader_writer_factory.set_cell_value(row_offset, col_offset, day_number)
        self.file_reader_writer_factory.set_cell_text_style(row_offset, col_offset, "v_center;h_center")
        self.file_reader_writer_factory.set_cell_value(row_offset, col_offset+1, day_abbr)
        self.file_reader_writer_factory.set_cell_text_style(row_offset, col_offset+1, "v_center")

        day_special_description = calendar_day.special_description
        self.file_reader_writer_factory.set_cell_value(row_offset, col_offset+2, day_special_description)
        self.file_reader_writer_factory.set_cell_text_style(row_offset, col_offset+2, "v_center")

        if self.show_week_numbers:
            # in case of monday show the week_number
            if calendar_day.isFirstDayOfWeek():
                week_number = calendar_day.week_number
                self.file_reader_writer_factory.set_cell_value(row_offset, col_offset+3, week_number)
                self.file_reader_writer_factory.set_cell_text_style(row_offset, col_offset+3, "v_center;h_right")
            else:
                self.file_reader_writer_factory.merge_cells(start_row=row_offset, start_column=start_column, end_row=row_offset, end_column=end_column)
                self.file_reader_writer_factory.set_cell_text_style(row_offset, start_column, "v_center")


        # hintergrundfarbe vorgeben
        # anzahl der sprints
        # alpha einteilen
        # pro release eine farbe zuteilen oder "durchcyclen" und pro sprint alpha anpassen
        if calendar_day.sprint_type is not None:
            if calendar_day.sprint_type == "RB":
                cell_bg_color = 'FFFFFF'

            if calendar_day.sprint_type[0:2] == "RS":
                shader_index = int(calendar_day.sprint_type[2:3]) - 1
                cell_bg_color = self.bg_colors[self.release_color][shader_index]


        if calendar_day.isHoliday:
            cell_font['color'] = "FF0000"
            cell_font_weekno['color'] = "FF0000"
            cell_font_special['color'] = "FF0000"
            cell_bg_color = 'FFFFFF'

        if calendar_day.isSpecialday:
            cell_font['color'] = "F79646"
            cell_font_weekno['color'] = "F79646"
            cell_font_special = self.cell_font_special.copy()
            cell_font_special['color'] = "F79646"
            cell_bg_color = '000000'

        if calendar_day.isWeedend:
            cell_font['bold'] = True
            cell_font_weekno['bold'] = True
            cell_font_special['bold'] = True
            cell_bg_color = 'FFFFFF'

        if calendar_day.isSpecialday and calendar_day.sprint_type == "RB":
            self.next_release_color()

        self.file_reader_writer_factory.set_cell_font_style(row_offset, col_offset, cell_font)
        self.file_reader_writer_factory.set_cell_font_style(row_offset, col_offset+1, cell_font)
        self.file_reader_writer_factory.set_cell_font_style(row_offset, col_offset+2, cell_font_special)
        self.file_reader_writer_factory.set_cell_font_style(row_offset, col_offset+3, cell_font_weekno)

        # logger.debug(calendar_day.to_string())


        self.file_reader_writer_factory.set_cell_bgcolor(row_offset, col_offset, cell_bg_color)
        self.file_reader_writer_factory.set_cell_bgcolor(row_offset, col_offset+1, cell_bg_color)
        self.file_reader_writer_factory.set_cell_bgcolor(row_offset, col_offset+2, cell_bg_color)
        self.file_reader_writer_factory.set_cell_bgcolor(row_offset, col_offset+3, cell_bg_color)

