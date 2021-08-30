from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

ori_wb = load_workbook('original.xlsx')
ori_ws = ori_wb.active

rank_wb = Workbook()

branch = 'AA'

for row in range(1, 14):
  rollno = ori_ws['A' + str(row)].value
  