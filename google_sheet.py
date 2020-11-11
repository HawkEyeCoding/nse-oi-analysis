#!/usr/bin/env python3

from datetime import datetime
import time

import gspread
import pandas as pd
import logging
from gspread.models import Cell

logging.basicConfig(level="INFO")
logger = logging.getLogger(__name__)


def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list) + 1)


def get_sheet():
    gc = gspread.service_account(filename='/googleshet_token.json')
    sheet = gc.open_by_key("<SHEET_IT>")
    return sheet


title_row = ['Date Time', 'Spot', 'Spot_Diff', 'Call OI - C', 'Put OI - C', 'Difference']


def insert_record(sheet_name, spot, call_oi_change, put_oi_change):
    sheet = get_sheet()
    ws = ""
    try:
        ws = sheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        create_new_worksheet(sheet_name)

    time.sleep(1)
    new_row_no = int(next_available_row(ws))
    date_time = datetime.now().strftime("%m/%d/%Y, %H:%M:%S")
    # DateTime	Spot	Spot_Diff	Call OI change	Put OI change	Diff (Call - Put)   Call OI Delta Put OI Delta
    record = [date_time, str(spot), f'=MIN(B{new_row_no}-B{new_row_no - 1})', str(call_oi_change), str(put_oi_change),
              str(put_oi_change - call_oi_change)]
    # print(record)
    ws.append_row(record, value_input_option='USER_ENTERED')


def read_row_values(sheet_name, column_name, symbol, row_data):
    ws = get_sheet().worksheet(sheet_name)

    first_row_vals = ws.row_values(1)
    for col in first_row_vals:
        if col.lower() == column_name.lower():
            col_index = first_row_vals.index(col)
            break

    # print(col_values[len(col_values) - 1])  # last value from column
    row_index = int(next_available_row(ws)) - 1
    row_values = ws.row_values(row_index)  # Last row values

    # print(f"Column value: {row_values[col_index]}")
    if row_values[col_index] == row_data[0]:
        if symbol == 'NIFTY' and row_values[col_index + 1] == "":
            cells = [Cell(row=row_index, col=2, value=row_data[1]),
                     Cell(row=row_index, col=3, value=row_data[2]),
                     Cell(row=row_index, col=4, value=row_data[3])]
            ws.update_cells(cells)
            return -1
        elif symbol == 'BANKNIFTY' and len(row_values) < 5:
            cells = [Cell(row=row_index, col=5, value=row_data[1]),
                     Cell(row=row_index, col=6, value=row_data[2]),
                     Cell(row=row_index, col=7, value=row_data[3])]
            ws.update_cells(cells)
            return -1
        return row_values
    else:
        if symbol == 'BANKNIFTY':
            row_data.insert(1, "")
            row_data.insert(2, "")
            row_data.insert(3, "")
        ws.append_row(row_data)
        return -1


def create_new_worksheet(sheet_name):
    ws = get_sheet().add_worksheet(sheet_name, 100, 100, 0)
    ws.append_row(title_row)
    logger.info(f"New sheet :{sheet_name} added")


def read_sheet_into_df(sheet_name):
    ws = get_sheet().worksheet(sheet_name)
    # Read excel worksheet and put into dataframe.
    dataframe = pd.DataFrame(ws.get_all_records())
    return dataframe
