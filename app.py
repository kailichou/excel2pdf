import sys, getopt, os, time
from termcolor import colored
import win32com.client 
from pywintypes import com_error


start = time.perf_counter()

# Get full command-line arguments
full_cmd_arguments = sys.argv 
arg_list = full_cmd_arguments[1:]

short_options = "hi:o:"
long_options = ["help", "input=", "output="]

try:
    arguments, values = getopt.getopt(arg_list, short_options, long_options)
    for current_argument, current_value in arguments:
        if current_argument in ("-h", "--help"):
            print("""
    Display help:

    Options and arguments:
    -h, --help          Displaying help message
    -i, --input         The input file path
    -o, --output        The output file path

    This program requires input and output arguments in order to proceed to the next task.
    """)
        elif current_argument in ("-i","--input"):
            WB_PATH = current_value
        elif current_argument in ("-o","--output"):
            PATH_TO_PDF = current_value.split('.')[0]
            print("Exporting PDF file mode (%s)"%(current_value))
except getopt.error as err:
    # Output errror, and return with an error code
    print(str(err))
    sys.exit(2)



# ========================= main ============================
# use a dynamic proxy
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False 

try:
    # Open Excel file
    wb = excel.Workbooks.Open(WB_PATH)
    sheet_names = [sh.Name for sh in wb.Sheets]
    # sheet_lst = list(range(1, len(sheet_names)+1))

    for sheet_name in sheet_names:
        # Select the sheets in order to output
        wb.WorkSheets(sheet_name).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0,PATH_TO_PDF+sheet_name+'.pdf')        
except NameError as err:
    print(colored('This program requires input and output arguments in order to proceed to the next task.','red'))
    sys.exit(2)
except com_error as e:
    print(e)
finally:
    wb.Close()
    excel.Quit()

end = time.perf_counter()
print("Total running time: ", end-start)
