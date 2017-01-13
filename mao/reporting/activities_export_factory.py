__author__ = 'Ofner Mario'

from mao.toolbox.logger import Logger

from logging import getLogger


logger = getLogger(__name__)


class ActivitiesExportFactory():

    file_reader_writer_factory = None
    table_header = []

    table_index = "Nummer"


    def __init__(self, _file_reader_writer_factory):
        self.file_reader_writer_factory = _file_reader_writer_factory


    def fetch_table_header_from_file(self):
        self.table_header = self.file_reader_writer_factory.get_table_header()

    def load_workbook(self, path_to_file, filename):
        # logger.debug("path_to_file: "+ path_to_file + filename)
        self.file_reader_writer_factory.load_workbook( _filename = path_to_file + filename )

    def change_or_create_worksheet(self, worksheet_name):
        if not self.file_reader_writer_factory.change_worksheet(worksheet_name):
            return False
        else:
            self.fetch_table_header_from_file()
            return True



    # --------------------------------------------------------------------------------
    # - Returns all already defined Revisions from Excel File
    # - The Column Revision is defined as "SVN Revision"
    def get_ae_details(self):
        # get column index
        # fetch all values for column

        col_ae_number = self.table_header.index("Nummer")
        # logger.debug("col_rev: "+str(col_rev))
        col_ae_status = self.table_header.index("Status")
        # logger.debug("col_ae: "+str(col_ae))
        col_ae_planrelease = self.table_header.index("Planrelease")
        # logger.debug("col_aes: "+str(col_aes))
        col_ae_testversion = self.table_header.index("Testversion")
        # logger.debug("col_aes: "+str(col_aes))

        ae_numbers = self.file_reader_writer_factory.get_all_values_for_column(col_ae_number+1)
        ae_status = self.file_reader_writer_factory.get_all_values_for_column(col_ae_status+1)
        ae_planrel = self.file_reader_writer_factory.get_all_values_for_column(col_ae_planrelease+1)
        ae_testvers = self.file_reader_writer_factory.get_all_values_for_column(col_ae_testversion+1)

        ae_list = {}
        for ae_index in range(0,len(ae_numbers)):
            ae_list[ae_numbers[ae_index]] = { "status": ae_status[ae_index], "planrelease": ae_planrel[ae_index], "testversion": ae_testvers[ae_index] }

        return ae_list






