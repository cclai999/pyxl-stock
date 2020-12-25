import requests
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from tools import get_html_to_file


def record_revenue_html(stock_no="2330", fname="2330_revenue.html"):
    url = f"http://jsjustweb.jihsun.com.tw/z/zc/zch/zch_{stock_no}.djhtm"
    get_html_to_file(url, fname)


def get_revenue_data(html_name):
    with open(html_name, "r") as f:
        html_text = f.read()
    dfs = pd.read_html(html_text)
    rev_df = dfs[2].iloc[6:, :7]
    return rev_df


def insert_stock_code_to_excel(stk_df, stk_code_sheet):
    for index, row in stk_df.iterrows():
        r = row.tolist()
        # print(r)
        stk_code_sheet.append(r)


def is_float_try(str):
    try:
        float(str)
        return True
    except ValueError:
        return False


def insert_stock_revenue_to_excel(rev_df, stk_revenue_sheet):
    # for row in dataframe_to_rows(rev_df, index=False, header=False):
    #     stk_revenue_sheet.append(row)
    rev_df = rev_df.reset_index().drop('index', axis=1)
    cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    start_row = 2
    for index, row in rev_df.iterrows():
        rlist = row.tolist()
        for col, value in enumerate(rlist):
            if is_float_try(value):
                stk_revenue_sheet[f"{cols[col]}{index + start_row}"] = float(value)
                stk_revenue_sheet[f"{cols[col]}{index + start_row}"].number_format = '#,##0.00'
            else:
                stk_revenue_sheet[f"{cols[col]}{index + start_row}"] = value


if __name__ == '__main__':
    stock_no = "2330"
    html_name = f"{stock_no}_revenue.html"
    if 1 == 2:
        record_revenue_html(stock_no=stock_no, fname=html_name)
    workbook = load_workbook(filename="template_stock_revenue.xlsx")
    stk_revenue_sheet = workbook["revenue"]

    rev_df = get_revenue_data(html_name)
    insert_stock_revenue_to_excel(rev_df, stk_revenue_sheet)

    workbook.save(f"{stock_no}_stock_revenue.xlsx")
