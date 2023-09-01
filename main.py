import openpyxl as xl
import pandas as pd
from datetime import datetime

def main():
    main_wb = xl.load_workbook("resources\\new.xlsx")

    npx_data = pd.read_csv("resources\\ndx_data.csv")
    
    sheet = main_wb["Rolling Returns"]

    # column L
    
    
    # for date, open, high, low, close, volume in zip(npx_data["Date"], npx_data["Open"], npx_data["High"], npx_data["Low"], npx_data["Close"], npx_data["Volume"] ):
    #     npx_date = datetime.strptime(date, "%Y-%m-%d")

    start_date = datetime.strptime(npx_data["Date"][0], "%Y-%m-%d")

    curr_index = 3 
    while(start_date != sheet.cell(curr_index, 11).value):
        curr_index +=1
    
    amount_changed = 0
    for date, close in zip(npx_data["Date"], npx_data["Close"]):
        # date
        sheet.cell(curr_index, 11, datetime.strptime(date, "%Y-%m-%d"))

        # amount
        sheet.cell(curr_index, 12, float(close))

        curr_index += 1
        amount_changed += 1

    print(f'Expected: {len(npx_data["Close"])}\nActual: {amount_changed}')
    main_wb.save("resources\\new.xlsx")
    return 0



if __name__ == "__main__":
    main()