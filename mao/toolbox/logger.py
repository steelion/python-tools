__author__ = 'Ofner Mario'

from datetime import datetime


class Logger():

    # --------------------------------------------------------------------------------
    # - Initializations
    # --------------------------------------------------------------------------------
    global_log_level = "DEBUG"
    log_microsec_enabled = True

    timestamp_format_str = ""

    def __init__(self):
        self.timestamp_format_str = "%Y-%m-%d %d:%m:%S"
        if self.log_microsec_enabled:
            self.timestamp_format_str += ".%f"


    def current_timestamp(self):
        current_ts = datetime.now()
        return current_ts.strftime(self.timestamp_format_str)


    def severity_check(self, _log_level):
        # todo: implement different levels and dependencies
        return True


    def log_level_shortcut(self, _log_level):
        if _log_level == "DEBUG":
            return "D"


    def msg_constructor(self, _message, _log_level):
        # todo: limit log messages by length or sth.
        log_message = self.current_timestamp() + " " + self.log_level_shortcut(_log_level) + ": " + str(_message)
        return log_message



    # --------------------------------------------------------------------------------
    # - Logging
    # --------------------------------------------------------------------------------
    def print(self, _message, _log_level = None):
        if _log_level is None:
            _log_level = "DEBUG"


        if self.severity_check(_log_level):
            print(self.msg_constructor(_message, _log_level))

