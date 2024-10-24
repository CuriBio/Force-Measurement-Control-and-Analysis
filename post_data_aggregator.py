import os
import openpyxl as xl
import numpy as np

wb_write = xl.Workbook()
ws_write = wb_write.active
output_name = "wb_fin"

spreadsheet_path = r"C:\Users\kevin\Downloads\RowC&DPostMeasurements\SpreadSheets"
for sheet_num, spreadsheet in enumerate(os.listdir(spreadsheet_path)):
    print(sheet_num, " ", spreadsheet)

    wb_read= xl.load_workbook(spreadsheet_path + "\\" + spreadsheet, data_only=True)
    ws_read = wb_read.active

    to_move = np.array(ws_read['A'])

    for cell_num, cell in enumerate(to_move):
        print(f'A{(cell_num + 1) + len(to_move)*sheet_num} ', cell.value)
        ws_write[f'A{(cell_num + 1) + len(to_move)*sheet_num}'] = cell.value
    ws_write[f'B{6*sheet_num + 1}'] = ws_read['B1'].value

wb_write.save(spreadsheet_path + "\\" + output_name + ".xlsx")