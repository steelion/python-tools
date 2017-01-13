__author__ = 'Ofner Mario'

from mao.reporting.structure.reporting_columns import ReportingColumn
from mao.reporting.structure.reporting_tables import ReportingTable

from datetime import datetime
import pprint

from logging import getLogger


logger = getLogger(__name__)

class ReleaseChangelogFactory():

    file_reader_writer_factory = None

    # Report Structure Definitions
    default_header_font = {"name":'Calibri', "size":10, "bold":True, "italic":False, "vertAlign":None, "underline":'none', "strike":False, "color":'000000'}
    default_cell_font = {"name":'Calibri', "size":10, "bold":False, "italic":False, "vertAlign":None, "underline":'none', "strike":False, "color":'000000'}

    header_rowheight = 40
    data_rowheight = 12.75

    revision_list = None
    release_changelog_table = None

    logging_worksheet_name = "Logging"


    def __init__(self, _file_reader_writer_factory):
        self.file_reader_writer_factory = _file_reader_writer_factory

        release_changelog_column_list = []

        release_changelog_column_list.append( ReportingColumn(header_description="SVN Revision", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="0", column_text_format=None, column_width=9) )

        release_changelog_column_list.append( ReportingColumn(header_description="Commit User", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="General", column_text_format=None, column_width=9) )

        release_changelog_column_list.append( ReportingColumn(header_description="Datetime", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="yyyy-mm-dd hh:mm:ss", column_text_format=None, column_width=19) )

        release_changelog_column_list.append( ReportingColumn(header_description="SVN Log Entry", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="General", column_text_format="text_wrap", column_width=70) )

        release_changelog_column_list.append( ReportingColumn(header_description="CROSS Aktivität", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="General", column_text_format="text_wrap", column_width=19) )

        release_changelog_column_list.append( ReportingColumn(header_description="CROSS Aktivität Planrelease", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="General", column_text_format=None, column_width=15) )

        release_changelog_column_list.append( ReportingColumn(header_description="CROSS Aktivität Status", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="General", column_text_format=None, column_width=9) )

        release_changelog_column_list.append( ReportingColumn(header_description="CROSS Aktivität Testversion", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="General", column_text_format=None, column_width=15) )

        release_changelog_column_list.append( ReportingColumn(header_description="Aus vorhergehendem Release/Branch übernommen", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="General", column_text_format=None, column_width=9) )

        release_changelog_column_list.append( ReportingColumn(header_description="Status gemeinsame abstimmung", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="General", column_text_format=None, column_width=15) )

        release_changelog_column_list.append( ReportingColumn(header_description="Kommentare gemeinsame Abstimmung", header_font=self.default_header_font, header_text_format="text_wrap; rotated",
            cell_font=self.default_cell_font, column_data_format="General", column_text_format="text_wrap", column_width=50) )

        self.release_changelog_table = ReportingTable(column_list=release_changelog_column_list, key_column="SVN Revision", header_rowheight=self.header_rowheight, data_rowheight=self.data_rowheight)

    def prepare_logging_worksheet(self):
        if not self.file_reader_writer_factory.change_worksheet(self.logging_worksheet_name):
            self.file_reader_writer_factory.create_worksheet(self.logging_worksheet_name)
            self.file_reader_writer_factory.change_worksheet(self.logging_worksheet_name)
            self.write_logging_header()
            self.file_reader_writer_factory.save_workbook()


    def get_table_header(self):
        return self.release_changelog_table.table_header

    table_header = property(fget=get_table_header)


    def get_column_data_format(self):
        return self.release_changelog_table.column_data_format


    def get_column_text_format(self):
        return self.release_changelog_table.column_text_format


    def get_header_text_format(self):
        return self.release_changelog_table.header_text_format


    def get_column_widths(self):
        return self.release_changelog_table.column_widths


    def load_changelog_workbook(self, path_to_file, filename):
        self.file_reader_writer_factory.load_workbook( _filename = path_to_file + filename )
        self.prepare_logging_worksheet()

    def change_or_create_worksheet(self, worksheet_name):

        if not self.file_reader_writer_factory.change_worksheet(worksheet_name):
            self.file_reader_writer_factory.create_worksheet(worksheet_name)
            self.file_reader_writer_factory.change_worksheet(worksheet_name)
            self.write_header()
            self.file_reader_writer_factory.save_workbook()

        self.reset_documented_revision_list()


    def save_workbook(self):
        self.file_reader_writer_factory.save_workbook()


    # --------------------------------------------------------------------------------
    # - Write defined Table Header initially to output File
    # - The Table Header is defined in this Class
    def write_header(self):

        # table_header = self.get_table_header()

        # # put the header in first row
        # rownum = 1
        # for col in range(1, len(table_header) + 1):
        #     _excel_factory.set_cell_value(rownum, col, table_header[col-1])
        # rownum += 1
        self.file_reader_writer_factory.set_row_values(1, self.table_header)
        self.file_reader_writer_factory.set_column_width(self.get_column_widths())

        self.file_reader_writer_factory.apply_styles_to_row(1,
            additional_format_list=self.get_header_text_format(),
            font_description=self.default_header_font
            )

    def write_logging_header(self):
        logging_header = [ "Documentation Update", "Log Type", "SVN Revision",
            "OLD: AE Planrelease", "OLD: AE Status", "OLD: AE Testversion",
            "NEW: AE Planrelease", "NEW: AE Status", "NEW: AE Testversion" ]
        logging_header_widths = [25, 10, 11, 23, 23, 23, 23, 23, 23 ]

        self.file_reader_writer_factory.set_row_values(1, logging_header)
        self.file_reader_writer_factory.set_column_width(logging_header_widths)
        self.file_reader_writer_factory.apply_styles_to_row(1, font_description=self.default_header_font )


    def append_revision(self, log_entry):
        new_revision = {}
        new_revision["SVN Revision"] = log_entry["entry"].revision
        new_revision["Commit User"] = log_entry["entry"].author
        new_revision["Datetime"] = log_entry["entry"].date
        new_revision["SVN Log Entry"] = log_entry["entry"].msg
        new_revision["CROSS Aktivität"] = log_entry["ae"]

        data_row = []
        for i in self.get_table_header():
            if i in new_revision:
                data_row.append(new_revision[i])
            else:
                data_row.append(None)

        self.file_reader_writer_factory.append_row(row_data_values=data_row)
        self.file_reader_writer_factory.apply_styles_to_row(
            number_format_list=self.get_column_data_format(),
            additional_format_list=self.get_column_text_format(),
            # dimension_row_height=self.data_rowheight,
            font_description=self.default_cell_font
            )
        self.reset_documented_revision_list()
        self.log_update(log_type="ADD", revision=new_revision["SVN Revision"])


    def reset_documented_revision_list(self):
        self.revision_list = None


    # --------------------------------------------------------------------------------
    # - Returns all already defined Revisions from Excel File
    # - The Column Revision is defined as "SVN Revision"
    def get_documented_revisions(self):
        # get column index
        # fetch all values for column
        if self.revision_list is None or len(self.revision_list) == 0:
            column_index = self.table_header.index("SVN Revision")
            self.revision_list = self.file_reader_writer_factory.get_all_values_for_column(column_index+1)

        return self.revision_list


    # --------------------------------------------------------------------------------
    # - Returns all already defined Revisions from Excel File
    # - The Column Revision is defined as "SVN Revision"
    def get_documented_revisions_with_details(self):
        # get column index
        # fetch all values for column

        col_rev = self.table_header.index("SVN Revision")
        # logger.debug("col_rev: "+str(col_rev))
        col_ae = self.table_header.index("CROSS Aktivität")
        # logger.debug("col_ae: "+str(col_ae))
        col_aes = self.table_header.index("CROSS Aktivität Status")
        # logger.debug("col_aes: "+str(col_aes))
        col_aep = self.table_header.index("CROSS Aktivität Planrelease")
        # logger.debug("col_aep: "+str(col_aep))
        col_aet = self.table_header.index("CROSS Aktivität Testversion")
        # logger.debug("col_aep: "+str(col_aep))

        revs_list = self.file_reader_writer_factory.get_all_values_for_column(col_rev+1)
        ae_list = self.file_reader_writer_factory.get_all_values_for_column(col_ae+1)
        ae_status_list = self.file_reader_writer_factory.get_all_values_for_column(col_aes+1)
        ae_planrel_list = self.file_reader_writer_factory.get_all_values_for_column(col_aep+1)
        ae_testvers_list = self.file_reader_writer_factory.get_all_values_for_column(col_aet+1)

        doc_revs_tmp = [list(revs_list), list(ae_list), list(ae_status_list), list(ae_planrel_list), list(ae_testvers_list)]
        documented_revisions = [list(x) for x in zip(*doc_revs_tmp)]
        return documented_revisions


    def update_revision_documentation(self, revision, ae_details):
        self.get_documented_revisions()
        # logger.debug(len(self.revision_list))

        col_aes = self.table_header.index("CROSS Aktivität Status")
        # logger.debug("col_aes: "+str(col_aes))
        col_aep = self.table_header.index("CROSS Aktivität Planrelease")

        col_aet = self.table_header.index("CROSS Aktivität Testversion")

        if revision in self.revision_list:
            # Read row-Number from Revision List and add 1 for the "1-based-List" and 1 for the Header
            row = self.revision_list.index(revision)+2
            old_data=[self.file_reader_writer_factory.get_cell_value(row, col_aes+1), self.file_reader_writer_factory.get_cell_value(row, col_aep+1), self.file_reader_writer_factory.get_cell_value(row, col_aet+1)]
            new_data=[ae_details["status"], ae_details["planrelease"], ae_details["testversion"]]

            self.file_reader_writer_factory.set_cell_value(row, col_aes+1, ae_details["status"])
            self.file_reader_writer_factory.set_cell_value(row, col_aep+1, ae_details["planrelease"])
            self.file_reader_writer_factory.set_cell_value(row, col_aet+1, ae_details["testversion"])

            self.log_update(log_type="MODIFY", revision=revision, old_data=old_data, new_data=new_data)


    def log_update(self, log_type, revision, old_data=None, new_data=None):
        # current worksheet
        data_row=[]
        data_row.append(str(datetime.now()))
        data_row.append(log_type)
        data_row.append(revision)
        if old_data is not None:
            data_row += old_data
            data_row += new_data

        cws = self.file_reader_writer_factory.ws
        self.file_reader_writer_factory.change_worksheet(self.logging_worksheet_name)
        self.file_reader_writer_factory.append_row(row_data_values=data_row)
        self.file_reader_writer_factory.apply_styles_to_row(font_description=self.default_cell_font)
        self.file_reader_writer_factory.ws = cws
        # logger.debug('Logging Changelog Update')



