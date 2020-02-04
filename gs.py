import sys
import logging
import json
import httplib2

import gspread
from oauth2client.service_account import ServiceAccountCredentials

class GoogleSpread:

    def __init__(self, credentials_file, email = 'dammy@mail.com', prefix = '', log_class=None):
        try:
            if log_class is None:
                # logger set
                # Prints logger info to terminal
                self.logger = logging.getLogger(str(prefix) + __name__)
                self.logger.setLevel(logging.INFO)  # Change this to DEBUG if you want a lot more info
                self.ch = logging.StreamHandler()
                # create formatter
                formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
                # add formatter to ch
                self.ch.setFormatter(formatter)
                self.logger.addHandler(self.ch)
            else:
                self.logger = log_class

            self.share_email = email

            feeds = 'https://spreadsheets.google.com/feeds'
            drive = 'https://www.googleapis.com/auth/drive'
            scope = [feeds, drive]

            self.credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)

            http = httplib2.Http()
            http = self.credentials.authorize(http)
            self.credentials.refresh(http)
            self.gc = gspread.authorize(self.credentials)

        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)


    def token_refresh(self):
        try:
            self.logger.info(f'access_token{self.credentials.access_token_expired}')
            if self.credentials.access_token_expired:
#                self.credentials.refresh(httplib2.Http())
                self.gc.login()  # refreshes the token
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def set_cntrol_book(self, ctr_book_key):
        try:
            self.ctr_book = self.gc.open_by_key(ctr_book_key)
            self.ctr_sheet = self.ctr_book.sheet1
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def create_book(self, name):
        try:
            cr = self.gc.create(name)

            cr_list = [cr.title, cr.id]
            self.append_raw(cr_list, True)

            cr.share(self.share_email, perm_type='user', role='writer')

        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def set_workbook(self, spread_key):
        try:
            self.book = self.gc.open_by_key(spread_key)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def add_worksheet(self, name, row, col, control = False):
        try:
            if control:
                self.ctr_book.add_worksheet(title=name, rows=row, cols=col)
            else:
                self.book.add_worksheet(title=name, rows=row, cols=col)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def set_worksheet_1(self):
        try:
            self.sheet = self.book.sheet1
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def set_worksheet(self, sheet_name):
        try:
            self.sheet = self.book.worksheet(sheet_name)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def get_rowvalues(self, row, control = False):
        try:
            if control:
                return self.ctr_sheet.row_values(row)
            else:
                return self.sheet.row_values(row)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def get_colvalues(self, col, control = False):
        try:
            if control:
                return self.ctr_sheet.col_values(col)
            else:
                return self.sheet.col_values(col)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def get_all_values_list(self, control = False):
        try:
            if control:
                return self.ctr_sheet.get_all_values()
            else:
                return self.sheet.get_all_values()
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def get_all_values_dict(self, control = False):
        try:
            if control:
                return self.ctr_sheet.get_all_records(empty2zero=False, head=1, default_blank='')
            else:
                return self.sheet.get_all_records(empty2zero=False, head=1, default_blank='')
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def get_col_count(self, control = False):
        try:
            if control:
                return self.ctr_sheet.col_count
            else:
                return self.sheet.col_count
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def get_row_count(self, control = False):
        try:
            if control:
                return self.ctr_sheet.row_count
            else:
                return self.sheet.row_count
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def get_cell_value(self, row, col, control = False):
        try:
            if control:
                return self.ctr_sheet.cell(row, col).value
            else:
                return self.sheet.cell(row, col).value
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def update_cell(self, row, col, data, control = False):
        try:
            if control:
                self.ctr_sheet.update_cell(row, col, data)
            else:
                self.sheet.update_cell(row, col, data)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def update_col(self, col, list, control = False):
        try:
            for i, value in enumerate(list):
                if control:
                    self.ctr_sheet.update_cell(i + 1, col, value)
                else:
                    self.sheet.update_cell(i + 1, col, value)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def update_raw(self, raw, list, control = False):
        try:
            for i, value in enumerate(list):
                if control:
                    self.ctr_sheet.update_cell(raw, i + 1, value)
                else:
                    self.sheet.update_cell(raw, i + 1, value)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def append_raw(self, list, control = False):
        try:
            if control:
                self.ctr_sheet.append_row(list)
            else:
                self.sheet.append_row(list)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def find_col(self, value):
        try:
            result = self.sheet.find(value)
            return result.col

        except gspread.exceptions.CellNotFound:
            return None

        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def find_cell(self, value):
        try:
            result = self.sheet.find(value)
            return result

        except gspread.exceptions.CellNotFound:
            return None

        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def find_all_cell(self, value):
        try:
            result = self.sheet.findall(value)
            return result

        except gspread.exceptions.CellNotFound:
            return None

        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def find_spread_key(self, book_name):
        try:
            result = self.ctr_sheet.find(book_name)
            return self.get_cell_value(result.row, 2, True)

        except gspread.exceptions.CellNotFound:
            return None

        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None


    def delete_row(self, row):
        try:
            self.sheet.delete_row(row)
        except Exception as e:
            self.logger.error(sys._getframe().f_code.co_name)
            self.logger.error(e)
            return None
