import win32com.client
import xlsxwriter
import xlwings as xw
import pandas as pd
import pythoncom

def read_result():

    win32com.client.Dispatch("Excel.Application",pythoncom.CoInitialize())

    excel_app = xw.App(visible=False)
    excel_book_dados = excel_app.books.open('dados_fatura.xlsx')
    excel_book = excel_app.books.open('TRATAMENTO_DE_DADOS.xlsx')
    excel_app.calculate()
    sheet = excel_book.sheets['Dados Energisa']
    headers = sheet['A1:C1'].value
    results = sheet['A2:C58'].value
    excel_book_dados.close()
    excel_book.close()
    excel_app.quit()

    allRows = []

    for row in results:
        data = {}
        for title, cell in zip(headers, row):
            data[title] = cell

        allRows.append(data)

    usedRows = list(filter(check_used_rows, allRows))

    return usedRows

def check_used_rows(row):
    rowValue = str(row['value'])

    if ('\\|/' in rowValue):
        return False

    if ('/|\\' in rowValue):
        return False

    if ('///' in rowValue):
        return False

    return True