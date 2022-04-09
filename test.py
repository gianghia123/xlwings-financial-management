import datetime as dt

import xlwings as xw
import pandas as ps

wb = xw.Book("test.xlsm")
io = wb.sheets["Nhập-xuất"]
data = wb.sheets["Data"]
table = None
if len(data.tables):
    table = data.tables["data"]
else:
    header_table = xw.Range(data.range("A1"), data.range("F100000")).expand()
    table = data.tables.add(source=header_table, name="data")


def add_data():
    date = io.range("C2").options(dates=dt.date).value
    type_of_transaction = True if io.range("C3").value == "Thu" else False
    amount = io.range("C4").value
    type_of_payment = io.range("C5").value
    note = io.range("C6").value
    datum = {'Ngày': [date.isoformat()],
             'Thu/chi': ["Thu" if type_of_transaction else "Chi"],
             'Số tiền thu': [amount if type_of_transaction else ""],
             'Số tiền chi': [amount if not type_of_transaction else ""],
             'Tiền mặt/ Ngân hàng': [type_of_payment],
             'Ghi chú': [note]}
    ps_datum = ps.DataFrame(data=datum)
    table.resize(xw.Range(data.range("A1"), data.range("F100000")))
    table.update(ps_datum, index=False)