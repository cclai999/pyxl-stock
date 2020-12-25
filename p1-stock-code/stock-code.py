import pandas as pd
from openpyxl import load_workbook

from tools import get_html_to_file

url1 = "https://isin.twse.com.tw/isin/class_main.jsp?owncode=&stockname=&isincode=&market=1&issuetype=1&industry_code=&Page=1&chklike=Y"
url2 = "https://isin.twse.com.tw/isin/class_main.jsp?owncode=&stockname=&isincode=&market=2&issuetype=4&industry_code=&Page=1&chklike=Y"


def get_stock_code(html_fname):
    f = open(html_fname, "r")
    stk_code_html = f.read()
    f.close()
    dfs = pd.read_html(stk_code_html)
    stk = dfs[0].loc[1:, :]
    compact_stk = stk[[2, 3, 7, 4, 6]]
    return compact_stk


def insert_stock_code_to_excel(stk_df, stk_code_sheet):
    for index, row in stk_df.iterrows():
        r = row.tolist()
        # print(r)
        stk_code_sheet.append(r)


def recod_html():
    get_html_to_file(url1, "stk_code1_html.txt")
    get_html_to_file(url2, "stk_code2_html.txt")


if __name__ == '__main__':
    if 1 == 2:
        recod_html()
    workbook = load_workbook(filename="stock_code_blank.xlsx")
    stk_code_sheet = workbook["stk_code"]

    stk_df = get_stock_code("stk_code1_html.txt")
    insert_stock_code_to_excel(stk_df, stk_code_sheet)

    stk_df = get_stock_code("stk_code2_html.txt")
    insert_stock_code_to_excel(stk_df, stk_code_sheet)

    workbook.save("stock_code.xlsx")
