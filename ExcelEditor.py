import openpyxl
import os
import shutil

class ExcelEditor:
    def __init__(self, filename):
        self.temp_file_name = 'temp.xlsx'
        self.origin_file_name = filename
        self.workbook = None

    def make_temp_file(self):
        shutil.copyfile(self.origin_file_name, self.temp_file_name)

    def load_excel_table(self):
        if self.workbook is None:
            self.workbook = openpyxl.open('temp.xlsx')


    def delete_make_file(self):
        os.remove(self.temp_file_name)

    def edit_cell(self, cell, value, change_teachers):
        self.load_excel_table()

        sheet = self.workbook._sheets[change_teachers]

        sheet[cell] = value

        self.workbook.save(self.temp_file_name)