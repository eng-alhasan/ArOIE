import argparse
import xlrd
import tree_analysis_model
import utils

class ArabicOpenIE:
  
  def __init__(self):
    parser = argparse.ArgumentParser()
    parser.add_argument('--dataset_excel_file',dest='dataset_excel_file', help='Dataset excel file')
    parser.add_argument('--excel_sheet_name',dest='excel_sheet_name', help='Excel sheet name')
    parser.add_argument('--results_file_name',dest='results_file_name', help='Excel results file name')
    parser.add_argument('--results_sheet_name',dest='results_sheet_name', help='Excel results sheet name')
    self.args = parser.parse_args()
    if self.args.dataset_excel_file is None or self.args.excel_sheet_name is None:
      parser.print_help()
      exit(-1)
    data = xlrd.open_workbook(self.args.dataset_excel_file)
    self.sheet = data.sheet_by_name(self.args.excel_sheet_name)
    utils.create_results_excel_file(self.args.results_file_name,self.args.results_sheet_name)

  def extract(self):   
    sentence_as_array =[]
    checker = 0
    for row_index in range(5, self.sheet.nrows):
        check_string = self.sheet.cell(row_index, 0).value
        if str(check_string).isnumeric():
            sentence_as_array.append(row_index)
            checker = row_index
        # End of sentence tokens          
        if check_string[0] == "#" and checker == row_index-1:
            sentence_as_array.append(row_index)
            tree_analysis_model.tree_analysis(sentence_as_array,self.args.dataset_excel_file, \
            self.args.excel_sheet_name,self.args.results_file_name,self.args.results_sheet_name)
            sentence_as_array = []
        # Last sentence
        if row_index == self.sheet.nrows-2:
            tree_analysis_model.tree_analysis(sentence_as_array, self.args.dataset_excel_file, \
            self.args.excel_sheet_name,self.args.results_file_name,self.args.results_sheet_name)
            sentence_as_array = []

if __name__ == "__main__":
  arabic_open_ie = ArabicOpenIE()
  arabic_open_ie.extract()

