from Excel_reader import gather_and_add_data
import xlwings as xw

#ENSURE THAT THE FIRST ROW IN MAIN_SHEET IS THE SAMPLE_ID

# EXCEL WORKBOOK CONTAINING SHEET WITH SOURCE DATA
main_wb = xw.Book('All data.xlsx') 

# SHEET CONTAINING SOURCE DATA
main_sheet = main_wb.sheets['Normalized data'] 

# CURRENT_BOOK IS THE BOOK TO WHICH NEW SHEETS WILL BE ADDED WITH 
# THE FILTERED/COMPARISON DATA
current_book = xw.Book('Test.xlsx')

gather_and_add_data(main_sheet, current_book)