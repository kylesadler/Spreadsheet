import os
import csv
import xlrd # reading .xlsx or .xls
import xlwt # writing .xlsx or .xls
import logging

class Spreadsheet:
    """ generalized spreadsheet reader. Works with .xlsx, .xls, and .csv """
    def __init__(self, path):
        self.path = path
        
        ext = os.path.splitext(self.path)[1]
        print(ext)
        if ext == '.csv':
            self.isCSV = True # True if CSV, false if XLSX or XLS
            with open(path, 'r') as f:
                self.sheet = list(csv.reader(f))

        elif ext in ['.xlsx', '.xls']:
            self.isCSV = False
            self.workbook = xlrd.open_workbook(path)
            self.sheet = self.workbook.sheet_by_index(0)
            self.s_names = self.workbook.sheet_names()
        
        else:
            raise Exception

    def get_sheet(self, num): # thows error on CSVs
        if self.isCSV:
            logging.error('CSV does not support sheets')
            return
        elif num >= len(self.s_names):
            logging.error('not enough sheets: ' + str(num) + ' >= ' + str(len(self.s_names)))
            return

        return self.workbook.sheet_by_index(num)

    def change_sheet(self, num): # thows error on CSVs
        if self.isCSV:
            logging.error('CSV does not support sheets')
            return
        elif num >= len(self.s_names):
            logging.error('not enough sheets: ' + str(num) + ' >= ' + str(len(self.s_names)))
            return

        self.sheet = self.workbook.sheet_by_index(num)
    
    def sheet_names(self):
        if self.isCSV:
            logging.error('CSV does not support sheets')
            return
        
        return self.s_names

    def __getitem__(self, r): # get a row
        if self.isCSV:
            return self.sheet[r]
        else:
            return [cell.value for cell in self.sheet.row(r)]

def test():
    files = ['05_2020_real_income_and_outlays.xlsx',  '06_2020_retail_sales_MoM.xls',
    '06_2020_business_inventories.xls', 'business_confidence.csv',   'housing_starts.csv',
    'capacity_utilization.csv',  'industrial_production.csv']

    for f in files:
        s = Spreadsheet(os.path.join('data', f))
        if not s.isCSV:
            print(s.sheet_names())

        print(s[0][0])    

        for cell in s[0]:
            print(cell)