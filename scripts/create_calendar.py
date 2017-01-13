__author__ = "Ofner Mario"
import pprint

import locale
from datetime import datetime

from mao.filehandling.excel.excel_factory import ExcelFactory

from mao.modules.calendar_generator.sprintcalendar_factory import SprintCalendarMonthColumnRenderer

import mao.toolbox.default_log_config
import logging

from mao.modules.calendar_generator.sprintcalendar_generator import *

# --------------------------------------------------------------------------------
# - Initializations
# --------------------------------------------------------------------------------

# Activate DEBUG Mode
mao.toolbox.default_log_config.set_rootLogger_log_level(logging.DEBUG)

# Settings for Calendar


locale.setlocale(locale.LC_ALL, 'de')

# Settings for AE-Snapshot
calendar_path = "D:/daten/projects/python/poi/calendar/"
calendar_path = "D:/files/poi/sourcen/python/resources/"
calendar_filename = "sprintcalendar.xlsx"

worksheet_name = "sprintcalendar_2017"


# Object Initializations
logger = logging.getLogger("sprintcalendar")

logger.info("Initializing factories...")
sprintcalendar_excel_factory = ExcelFactory()
sprintcalendar_factory = SprintCalendarMonthColumnRenderer(sprintcalendar_excel_factory)

logger.info("Creating file")
sprintcalendar_factory.load_workbook(calendar_path, calendar_filename)
sprintcalendar_factory.remove_worksheet("Sheet")
sprintcalendar_factory.remove_worksheet(worksheet_name)
sprintcalendar_factory.change_or_create_worksheet(worksheet_name)


logger.info("Creating Calendar Object")
calendar_start = datetime(2016, 12, 1)
calendar_end = datetime(2018, 2, 28)
cal = SprintCalendar(start_date=calendar_start, end_date=calendar_end)

logger.info("Importing Calendar Holidays")
holidays_list = [
    # 2016
    {'name': "Neujahr", 'date': datetime(2016, 1, 1)},
    {'name': "Hl. 3 Könige", 'date': datetime(2016, 1, 6)},
    {'name': "Ostersonntag", 'date': datetime(2016, 3, 27)},
    {'name': "Ostermontag", 'date': datetime(2016, 3, 28)},
    {'name': "Staatsfeiertag", 'date': datetime(2016, 5, 1)},
    {'name': "Chr. Himmelfahrt", 'date': datetime(2016, 5, 5)},
    {'name': "Pfingstsonntag", 'date': datetime(2016, 5, 15)},
    {'name': "Pfingstmontag", 'date': datetime(2016, 5, 16)},
    {'name': "Fronleichnam", 'date': datetime(2016, 5, 26)},
    {'name': "Maria Himmelfahrt", 'date': datetime(2016, 8, 15)},
    {'name': "Nationalfeiertag", 'date': datetime(2016, 10, 26)},
    {'name': "Allerheiligen", 'date': datetime(2016, 11, 1)},
    {'name': "Maria Empfängnis", 'date': datetime(2016, 12, 8)},
    {'name': "Hl. Abend", 'date': datetime(2016, 12, 24)},
    {'name': "Christtag", 'date': datetime(2016, 12, 25)},
    {'name': "Stefanitag", 'date': datetime(2016, 12, 26)},
    {'name': "Silvester", 'date': datetime(2016, 12, 31)},
    # 2017
    {'name': "Neujahr", 'date': datetime(2017, 1, 1)},
    {'name': "Hl. 3 Könige", 'date': datetime(2017, 1, 6)},
    {'name': "Ostersonntag", 'date': datetime(2017, 4, 16)},
    {'name': "Ostermontag", 'date': datetime(2017, 4, 17)},
    {'name': "Staatsfeiertag", 'date': datetime(2017, 5, 1)},
    {'name': "Chr. Himmelfahrt", 'date': datetime(2017, 5, 25)},
    {'name': "Pfingstsonntag", 'date': datetime(2017, 6, 4)},
    {'name': "Pfingstmontag", 'date': datetime(2017, 6, 5)},
    {'name': "Fronleichnam", 'date': datetime(2017, 6, 15)},
    {'name': "Maria Himmelfahrt", 'date': datetime(2017, 8, 15)},
    {'name': "Nationalfeiertag", 'date': datetime(2017, 10, 26)},
    {'name': "Allerheiligen", 'date': datetime(2017, 11, 1)},
    {'name': "Maria Empfängnis", 'date': datetime(2017, 12, 8)},
    {'name': "Hl. Abend", 'date': datetime(2017, 12, 24)},
    {'name': "Christtag", 'date': datetime(2017, 12, 25)},
    {'name': "Stefanitag", 'date': datetime(2017, 12, 26)},
    {'name': "Silvester", 'date': datetime(2017, 12, 31)},
    # 2018
    {'name': "Neujahr", 'date': datetime(2018, 1, 1)},
    {'name': "Hl. 3 Könige", 'date': datetime(2018, 1, 6)},
    {'name': "Ostersonntag", 'date': datetime(2018, 4, 1)},
    {'name': "Ostermontag", 'date': datetime(2018, 4, 2)},
    {'name': "Staatsfeiertag", 'date': datetime(2018, 5, 1)},
    {'name': "Chr. Himmelfahrt", 'date': datetime(2018, 5, 10)},
    {'name': "Pfingstsonntag", 'date': datetime(2018, 5, 20)},
    {'name': "Pfingstmontag", 'date': datetime(2018, 5, 21)},
    {'name': "Fronleichnam", 'date': datetime(2018, 5, 31)},
    {'name': "Maria Himmelfahrt", 'date': datetime(2018, 8, 15)},
    {'name': "Nationalfeiertag", 'date': datetime(2018, 10, 26)},
    {'name': "Allerheiligen", 'date': datetime(2018, 11, 1)},
    {'name': "Maria Empfängnis", 'date': datetime(2018, 12, 8)},
    {'name': "Hl. Abend", 'date': datetime(2018, 12, 24)},
    {'name': "Christtag", 'date': datetime(2018, 12, 25)},
    {'name': "Stefanitag", 'date': datetime(2018, 12, 26)},
    {'name': "Silvester", 'date': datetime(2018, 12, 31)}
    ]
cal.set_holidays(holidays_list)

logger.info("Importing Sprint Information")

# Settings:
# release_length = 3
# sprint_length = '4W'
# release_buffer = '1W'
# Ausnahmen später adaptieren: zb. Dezembersprintende ist der Donnerstag anstatt dem Freitag

# der Vierte Eintrag ist der Style für die Zeilen

sprints = [
    { 'sprint_begin': datetime(2016, 11, 28), 'sprint_end': datetime(2016, 12, 22), 'sprint_name': "2.39.1", 'sprint_type': "RS1" },
    { 'sprint_begin': datetime(2016, 12, 23), 'sprint_end': datetime(2017, 1, 20), 'sprint_name': "2.39.2", 'sprint_type': "RS2" },
    { 'sprint_begin': datetime(2017, 1, 23), 'sprint_end': datetime(2017, 2, 17), 'sprint_name': "2.39.3", 'sprint_type': "RS3" },
    { 'sprint_begin': datetime(2017, 2, 20), 'sprint_end': datetime(2017, 2, 24), 'sprint_name': "Release 2.39", 'sprint_type': "RB"},

    { 'sprint_begin': datetime(2017, 2, 27), 'sprint_end': datetime(2017, 3, 31), 'sprint_name': "2.40.1", 'sprint_type': "RS1" },
    { 'sprint_begin': datetime(2017, 4, 3), 'sprint_end': datetime(2017, 4, 28), 'sprint_name': "2.40.2", 'sprint_type': "RS2" },
    { 'sprint_begin': datetime(2017, 5, 2), 'sprint_end': datetime(2017, 5, 26), 'sprint_name': "2.40.3", 'sprint_type': "RS3" },
    { 'sprint_begin': datetime(2017, 5, 29), 'sprint_end': datetime(2017, 6, 2), 'sprint_name': "Release 2.40", 'sprint_type': "RB"},

    { 'sprint_begin': datetime(2017, 6, 5), 'sprint_end': datetime(2017, 6, 30), 'sprint_name': "2.41.1", 'sprint_type': "RS1" },
    { 'sprint_begin': datetime(2017, 7, 3), 'sprint_end': datetime(2017, 7, 28), 'sprint_name': "2.41.2", 'sprint_type': "RS2" },
    { 'sprint_begin': datetime(2017, 7, 31), 'sprint_end': datetime(2017, 8, 25), 'sprint_name': "2.41.3", 'sprint_type': "RS3" },
    { 'sprint_begin': datetime(2017, 8, 28), 'sprint_end': datetime(2017, 9, 1), 'sprint_name': "Release 2.41", 'sprint_type': "RB"},

    { 'sprint_begin': datetime(2017, 9, 4), 'sprint_end': datetime(2017, 9, 29), 'sprint_name': "2.42.1", 'sprint_type': "RS1" },
    { 'sprint_begin': datetime(2017, 10, 2), 'sprint_end': datetime(2017, 10, 27), 'sprint_name': "2.42.2", 'sprint_type': "RS2" },
    { 'sprint_begin': datetime(2017, 10, 30), 'sprint_end': datetime(2017, 11, 24), 'sprint_name': "2.42.3", 'sprint_type': "RS3" },
    { 'sprint_begin': datetime(2017, 11, 27), 'sprint_end': datetime(2017, 12, 1), 'sprint_name': "Release 2.42", 'sprint_type': "RB"},

    { 'sprint_begin': datetime(2017, 12, 4), 'sprint_end': datetime(2017, 12, 29), 'sprint_name': "2.43.1", 'sprint_type': "RS1" },
    { 'sprint_begin': datetime(2018, 1, 2), 'sprint_end': datetime(2018, 1, 26), 'sprint_name': "2.43.2", 'sprint_type': "RS2" },
    { 'sprint_begin': datetime(2018, 1, 29), 'sprint_end': datetime(2018, 2, 23), 'sprint_name': "2.43.3", 'sprint_type': "RS3" },
    { 'sprint_begin': datetime(2018, 2, 26), 'sprint_end': datetime(2018, 3, 2), 'sprint_name': "Release 2.43", 'sprint_type': "RB"}
    ]

sprint_infos = { 'release_length': 3, 'sprint_length': "4W", 'release_buffer': "1W" }
cal.import_sprints(sprints, sprint_infos)


logger.info("Calendar Object cration finished")

sprintcalendar_factory.sprintcalendar = cal

logger.info("Start rendering Calendar")
sprintcalendar_factory.render_calendar()
logger.info("Rendering Calendar finished")


# --------------------------------------------------------------------------------
# - Sprintcalendar UseIT
# --------------------------------------------------------------------------------

# ts_end = datetime.now()


# day = CalendarDay(2016,1,1)
# day.isHoliday = True
# print( day.isHoliday )

sprintcalendar_factory.save_workbook()




