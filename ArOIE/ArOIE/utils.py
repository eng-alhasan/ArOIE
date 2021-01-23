import xlwt

def create_results_excel_file(file_name,sheet_name):
  workbook = xlwt.Workbook()
  sheet = workbook.add_sheet(sheet_name)
  workbook.save(file_name)