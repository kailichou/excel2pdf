import win32com.client 
from pywintypes import com_error
import time

start = time.perf_counter()
WB_PATH = r'C:\Users\Administrator\Downloads\Copy of 实践题1：Sample-Superstore-Subset-Excel.xlsx'

PATH_TO_PDF = r'C:\Users\Administrator\Downloads\test.pdf'

# Open Excel Application Successfully
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False 

# Open Excel file
wb = excel.Workbooks.Open(WB_PATH)

# Select the sheets in order to output
wb.WorkSheets([1,2,3]).Select()
wb.ActiveSheet.ExportAsFixedFormat(0,PATH_TO_PDF)
wb.Close()

excel.Quit()
end = time.perf_counter()

print("Total cost time: ", end-start)

# try:
#     wb = excel.Workbooks.Open(WB_PATH)
#     print('workbook is open')
#     ws_index_list = [1,2,3]
#     wb.WorkSheets(ws_index_list).Select()
#     print('workbook is active')
#     wb.ActiveSheet.ExportAsFixedFormat(0,PATH_TO_PDF)

# except com_error as e:
#     print(e)
# else:
#     print('succeeded')
# finally:
#     wb.Close()
#     excel.Quit()