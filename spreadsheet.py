import os
import csv
import xlrd # reading .xlsx or .xls
import logging
import xlsxwriter
from datetime import datetime

class SpreadsheetReader:
    """ generalized spreadsheet reader. Works with .xlsx, .xls, and .csv """
    def __init__(self, path):
        self.path = path
        
        ext = os.path.splitext(self.path)[1]
        # print(ext)
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

    def change_sheet(self, num): # thows error on CSVs
        if self.isCSV:
            # logging.error('CSV does not support sheets')
            return
            
        elif num >= len(self.s_names):
            logging.error('not enough sheets: ' + str(num) + ' >= ' + str(len(self.s_names)))
            return

        self.sheet = self.workbook.sheet_by_index(num)
    
    def sheet_names(self):
        if self.isCSV:
            # logging.error('CSV does not support sheets')
            return
        
        return self.s_names

    def __getitem__(self, r): # get a row
        if self.isCSV:
            return self.sheet[r]
        else:
            return [cell.value for cell in self.sheet.row(r)]

    def numrows(self):
        if self.isCSV:
            return len(self.sheet)
        else:
            return self.sheet.nrows
    
    def numcols(self):
        if self.isCSV:
            logging.error('CSV does not support colums. use len(spreadsheet[index])')
            return
        else:
            return self.sheet.ncols

class SpreadsheetWriter:
    def __init__(self, path):
        self.path = path
        self.workbook = xlsxwriter.Workbook(path)
        self.current_columns = {}
        self.header_style = self.workbook.add_format({'bold': 1})
        self.date_format = self.workbook.add_format({'num_format': 'mm/dd/yy'})

    
    def save(self):
        self.workbook.close()

    def write_data(self, data):
        for i, row in enumerate(data):
            if i == 0:
                self.write_row(i, row, True)
            else:
                self.write_row(i, row)

    def write_col(self, col, data, header_style=None):
        if header_style is None:
            self.sheet.write_column(0, col, data)
        else:
            self.sheet.write_column(0, col, data[0], header_style)
            self.sheet.write_column(1, col, data[1:])
            
        self.sheet.set_column(col, col, max([len(str(x))*250 for x in data]))

    def write_row(self, row, data, is_header=False):
        if is_header:
            self.sheet.write_row(row, 0, data, self.header_style)
        else:
            self.sheet.write_datetime(row, 0, data[0], self.date_format)
            self.sheet.write_row(row, 1, [float(x) if x != '' else '' for x in data[1:]])

    def add_chart(self, name, col, end_row, title, yaxis, cell):
        start_range = '=\''+name+'\'!$'
        chart = self.workbook.add_chart({'type': 'line'})

        if yaxis is not None:
            chart.set_y_axis({'name': yaxis})

        chart.add_series({
            'categories': start_range + 'A$2:$A$'+end_row,
            'values': start_range + col+'$2:$'+col+'$'+end_row,
            })
        chart.set_title({'name': title})
        chart.set_x_axis({'date_axis': True})
        chart.set_legend({'none': True})
        chart.set_style(35)
        self.sheet.insert_chart(cell, chart)

    def get_sheet(self, name):
        self.name = name
        
        self.sheet = self.workbook.get_worksheet_by_name(self.name)
        
        if self.sheet is None:
            self.sheet = self.workbook.add_worksheet(self.name)
            
        if self.name not in self.current_columns:
            self.current_columns[self.name] = 0
 
class SSWriterSummary(SpreadsheetWriter):
    def __init__(self, path):
        super().__init__(path)
        self.get_sheet("Summary")
        self.chart_count = 0

    def add_chart(self, name, col, end_row, title, yaxis, cell):
        super().add_chart(name, col, end_row, title, yaxis, cell)
        
        # add the same chart to summary page
        old_name = self.name
        self.get_sheet("Summary")
        super().add_chart(name, col, end_row, title, yaxis, self._get_chart_cell())
        self.get_sheet(old_name)

    def _get_chart_cell(self):
        vert_spacing = 15
        if self.chart_count % 2 == 0:
            col = 'B'
            row = self.chart_count
        else:
            col = 'J'
            row = self.chart_count - 1
        
        self.chart_count += 1
        return col+str(int(row/2 * vert_spacing + 2))


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
