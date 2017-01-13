__author__ = "Ofner Mario"

import mao.toolbox.default_log_config

from datetime import datetime, date, timedelta
import pprint
import logging
import calendar

# --------------------------------------------------------------------------------
# - Initializations
# --------------------------------------------------------------------------------

# Activate DEBUG Mode
mao.toolbox.default_log_config.set_rootLogger_log_level(logging.DEBUG)
logger = logging.getLogger()

__all__ = [
    'CalendarDay',
    'CalendarMonth',
    'SprintCalendar'
    ]

class CalendarDay():

    def __init__(self, year, month, day):
        self.year = year
        self.month = month
        self.day_of_month = day

        if calendar.weekday(self.year, self.month, self.day_of_month) > 4:
            self._isWeekend = True
        else:
            self._isWeekend = False

        self._isHoliday = False
        self._isSpecialday = False
        self._sprint_type = None
        self._sprint_name = None

    @property
    def isHoliday(self):
        return self._isHoliday

    @isHoliday.setter
    def isHoliday(self, holiday_name):
        self.holiday_name = holiday_name
        self._isHoliday = True

    @property
    def isSpecialday(self):
        return self._isSpecialday

    @isSpecialday.setter
    def isSpecialday(self, specialday_name):
        self.specialday_name = specialday_name
        self._isSpecialday = True

    @property
    def isWeedend(self):
        return self._isWeekend

    def isFirstDayOfWeek(self):
        # means Monday
        first_day_of_week = calendar.MONDAY
        if calendar.weekday(self.year, self.month, self.day_of_month) == first_day_of_week:
            return True
        return False

    @property
    def special_description(self):
        if self.isSpecialday:
            return self.specialday_name
        elif self.isHoliday:
            return self.holiday_name
        else:
            return ""

    @property
    def name(self):
        return calendar.day_name[calendar.weekday(self.year, self.month, self.day_of_month)]

    @property
    def abbr(self):
        return calendar.day_abbr[calendar.weekday(self.year, self.month, self.day_of_month)]

    @property
    def week_number(self):
         return date(self.year, self.month, self.day_of_month).isocalendar()[1]

    @property
    def sprint_type(self):
        return self._sprint_type
    @sprint_type.setter
    def sprint_type(self, sprint_type):
        self._sprint_type = sprint_type

    def to_string(self):
        add_str_we = ""
        add_str_hd = ""
        if self.isWeedend:
            add_str_we = " (W)"
        if self.isHoliday:
            add_str_hd = " (H)"

        return str(self.year)+"/"+str(self.month)+"/"+str(self.day_of_month)+" |"+str(self.week_number)+"|"+add_str_we+add_str_hd


class CalendarMonth():
    def __init__(self, year, month):
        self.year = year
        self.month = month
        self.monthcalendar = calendar.monthcalendar(self.year, self.month)
        (self.first_day_of_month, self.days) = calendar.monthrange(self.year, self.month)
        self._day_list = []

        for week in self.monthcalendar:
            for day in week:
                if day > 0:
                    self._day_list.append(CalendarDay(self.year, self.month, day))

    @property
    def day_list(self):
        return self._day_list

    @property
    def name(self):
        return calendar.month_name[self.month]


class SprintCalendar():

    def __init__(self, start_date = None, end_date = None):

        if start_date is None:
            dt_now = datetime.now()
            start_date = datetime(dt_now.year, dt_now.month, 1)
        if end_date is None:
            end_date = datetime(start_date.year, start_date.month, calendar.monthrange(start_date.year,start_date.month)[1])

        if not isinstance(start_date, datetime):
            logger.info("Format Error %s", "start_date")
        if not isinstance(end_date, datetime):
            logger.info("Format Error %s", "end_date")

        self._start_date = start_date
        self._end_date = end_date
        self._sprint_list = None
        self._release_list = None


        self.generate_base_objects()

    def generate_base_objects(self):
        m1 = (self._start_date.year*12)+(self._start_date.month-1)
        m2 = (self._end_date.year*12)+(self._end_date.month-1)
        self._month_list = []

        for i in range(m1, m2+1):
            self._month_list.append(CalendarMonth(int(i/12), int(i%12)+1))


    @property
    def start_date(self):
        return self._start_date

    @property
    def end_date(self):
        return self._end_date

    @property
    def sprint_start(self):
        return self._sprint_start

    @sprint_start.setter
    def sprint_start(self, sprint_start):
        self._sprint_start = sprint_start

    @property
    def sprint_length(self):
        return self._sprint_length

    @sprint_length.setter
    def sprint_length(self, length):
        """
        The length should be set by 20D or 4W
        todo
        """
        self._sprint_length = length

    @property
    def release_length(self):
        return self._release_length

    @property
    def sprint_list(self):
        return self._sprint_list

    @property
    def release_list(self):
        return self._release_list

    def import_sprints(self, sprints, sprint_infos):
        self._release_length = sprint_infos["release_length"]
        self.sprint_length = sprint_infos["sprint_length"]

        self._sprint_list = []
        self._release_list = []

        for sprint in sprints:
            if self.day_in_calendar(sprint['sprint_begin']) or self.day_in_calendar(sprint['sprint_end']):
                # set all days with specific style and end of sprint with secial_day
                self._sprint_list.append(sprint)
                if sprint['sprint_type'] == "RB":
                    self._release_list.append(sprint['sprint_name'])

                sprint_day_date = sprint['sprint_begin']
                day_delta = timedelta(days=1)
                while sprint_day_date <= sprint['sprint_end']:
                    if self.day_in_calendar(sprint_day_date):
                        sprint_day = self.get_calendar_day_for_date(sprint_day_date)
                        sprint_day.sprint_type = sprint['sprint_type']
                        sprint_day.sprint_name = sprint['sprint_name']

                        if sprint_day_date == sprint['sprint_end']:
                            sprint_day.isSpecialday = sprint['sprint_name']

                    sprint_day_date += day_delta



    def calendar_releases_count(self):
        return len(self.release_list)

    def calendar_sprint_count(self):
        return len(self.sprint_list)



    @property
    def month_list(self):
        return self._month_list

    def get_calendar_day_for_date(self, date):
        for m in self.month_list:
            if m.year == date.year and m.month == date.month:
                return m.day_list[date.day-1]



    def set_holidays(self, holidays_list):
        for holiday_entry in holidays_list:
            if self.day_in_calendar( holiday_entry['date'] ):
                logger.debug("holiday %s, is in range", holiday_entry['name'])
                for m in self.month_list:
                    if m.year == holiday_entry['date'].year and m.month == holiday_entry['date'].month:
                        m.day_list[holiday_entry['date'].day-1].isHoliday = holiday_entry['name']


    def day_in_calendar(self, date):
        if date is not None:
            if (date >= self.start_date) and (date <= self.end_date):
                return True
            else:
                return False

    @property
    def show_week_numbers(self):
        return self._show_week_numbers

    @show_week_numbers.setter
    def show_week_numbers(self, enable_flag):
        self._show_week_numbers = enable_flag


