import time
import requests
import pandas as pd
from datetime import datetime, date
from openpyxl import load_workbook

from tools import get_html_to_file

stock_no = "2330"
today_str = datetime.today().strftime("%Y%m%d")
date_str = today_str
url = f"https://www.twse.com.tw/exchangeReport/FMSRFK?response=html&date={date_str}&stockNo={stock_no}"


def get_stock_price_m_from_file(fname):
    f = open(fname, "r")
    html_text = f.read()
    f.close()
    dfs = pd.read_html(html_text)
    return dfs[0]


def record_price_html(stock_no, date_str, fname):
    url = f"https://www.twse.com.tw/exchangeReport/FMSRFK?response=html&date={date_str}&stockNo={stock_no}"
    get_html_to_file(url, fname)


def insert_stock_price_to_excel(stk_df, stk_price_sheet, start_row=1):
    for index, row in stk_df.iterrows():
        row_index = start_row + (index + 1)
        cell_loc = f"A{row_index}"
        year = int(row.iloc[0])
        month = int(row.iloc[1])
        first_day_month = date(year=year + 1911, month=month, day=1)
        stk_price_sheet[cell_loc] = first_day_month
        stk_price_sheet[cell_loc].number_format = 'yyyy/m/d'
        cell_loc = f"B{row_index}"
        cell = stk_price_sheet[cell_loc]
        cell.value = row.iloc[4]
    return row_index


if __name__ == '__main__':
    stock_no = "2330"
    workbook = load_workbook(filename=f"{stock_no}_stock_revenue.xlsx")
    stk_price_sheet = workbook["price_m"]
    this_year = datetime.today().year
    row_index = 1
    for y in range(this_year - 3, this_year + 1):
        if y == this_year:
            date_str = datetime.today().strftime("%Y%m%d")              # today_str
        else:
            date_str = date(year=y, month=1, day=1).strftime("%Y%m%d")  # first_day_of_y

        html_name = f"{stock_no}_{y}.html"
        if 1 == 2:
            record_price_html(stock_no=stock_no, date_str=date_str, fname=html_name)
            time.sleep(30)
        price_m_df = get_stock_price_m_from_file(fname=html_name)
        row_index = insert_stock_price_to_excel(price_m_df, stk_price_sheet, row_index)

    workbook.save(f"{stock_no}_stock_revenue.xlsx")
