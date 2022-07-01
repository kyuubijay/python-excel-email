from openpyxl import load_workbook
from win32com.client import Dispatch
import win32com.client as win32
import os, sys
from pynput import keyboard

args_length = len(sys.argv)
print(args_length)
if (args_length < 2):
    print(f'Usage: ts.py <filename>')
    exit(0)

filename = 'Collabera New Timesheet Template - June 16-30, 2022.xlsx'

fname = os.getcwd() + f"\{filename[:-1]}"
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)
wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()

wb = load_workbook(filename = filename)
sheet = wb['Summary']

start_row=14

attendance=['Normal', 'Rest Day', 'Leave', 'Working Rest Day', 'Special Holiday', 'Special Holiday on RD',
            'Working-Special Holiday on RD', 'Legal Holiday', 'Legal Holiday on RD', 'Working-Legal Holiday on RD']
dates=''.join([n for n in filename.split(',')[0] if n.isdigit()]) #1630
no_of_days=abs(int(dates[:2]) - int(dates[2:4]))

date_string = filename[filename.find('- ') + 2 : filename.find('.')]

sheet['B1'] = 'UnionBank'
sheet['B2'] = 'The Portal - eGobyerno'
sheet['B3'] = '28487'
sheet['B4'] = 'Jaymark Macaranas'

for row in range(start_row, start_row + no_of_days + 1):
    srow = str(row)
    att = sheet['C' + srow].value

    if(att == 'Normal'):
        sheet['D' + srow] = '9:00'
        sheet['E' + srow] = sheet['A' + str(row)].value
        sheet['F' + srow] = '18:00'
        sheet['G' + srow] = '9:00'
        sheet['H' + srow] = '18:00'
        sheet['N' + srow] = 'Strict'

output_file = f'Timesheet - {date_string}.xlsx'

wb.save(output_file)

xl = Dispatch("Excel.Application")
xl.Visible = True
wb = xl.Workbooks.Open(os.getcwd() + f"\{output_file}")
# wb.Close()
# xl.Quit()

print('Send email now? (Y/n)')


# while 1:
#     with keyboard.Events() as events:
#         # Block for as much as possible
#         event = events.get(1e6)
#         if event.key == keyboard.KeyCode.from_char('n'):
#             print("END")
#             break
#         elif event.key == keyboard.KeyCode.from_char('y'):
#             print("CONT")