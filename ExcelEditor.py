import openpyxl
import os
import shutil

class ExcelEditor:
    def __init__(self, filename):
        self.temp_file_name = 'temp.xlsx'
        self.origin_file_name = filename
        self.workbook = None

    def load_excel_table(self):
        self.make_temp_file()

        if self.workbook is None:
            self.workbook = openpyxl.open('temp.xlsx')


    def make_temp_file(self):
        if not os.path.exists(self.temp_file_name):
            shutil.copy(self.origin_file_name, self.temp_file_name)

    def delete_make_file(self):
        os.remove(self.temp_file_name)

    def edit_cell(self, cell, value, change_teachers):
        self.load_excel_table()

        sheet = self.workbook[self.workbook.sheetnames[change_teachers]]

        sheet[cell] = value

        self.workbook.save(self.temp_file_name)