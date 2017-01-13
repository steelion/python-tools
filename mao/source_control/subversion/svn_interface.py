__author__ = 'Ofner Mario'


import pprint
import svn.local
import svn.remote

from datetime import datetime
import re

from logging import getLogger


logger = getLogger(__name__)


class SVNInterface():

    svn_handle = None


    def open_svn_connection(self, _svn_url):
        self.svn_handle = svn.remote.RemoteClient(_svn_url)


    def read_svn_log(self, _timestamp_from, _timestamp_to):
        svn_log = {}
        ts_from = _timestamp_from
        ts_to = _timestamp_to
        # class datetime.datetime(year, month, day[, hour[, minute[, second[, microsecond[, tzinfo]]]]])

        # for log in branch_remote_237.log_default( timestamp_from_dt=ts_from, timestamp_to_dt=ts_to, stop_on_copy=True, changelist=True ):
        for log in self.svn_handle.log_default( timestamp_from_dt=ts_from, timestamp_to_dt=ts_to, stop_on_copy=True ):
            if log.msg != "neue Uebersetzungen":

                activities = ""
                if log.msg is not None:
                    re_match = re.findall( "(CR\d{6})", log.msg )

                if re_match:
                    # set( xx ) -> creates a set from the regex-list
                    #    to make the result unique
                    # after that convert the set to a string
                    #    and join all set-entries with: ', '
                    activities = ", ".join(set(re_match))
                svn_log[log.revision] = {'entry': log, 'ae': activities}
                # logger.debug(log.revision)

        return svn_log


# svn_log = read_svn_log()
# if svn_log is not None:
#     pprint.pprint(svn_log)
