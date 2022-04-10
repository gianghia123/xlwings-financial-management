import datetime as dt

import xlwings as xw
import pandas as ps

import constant as ct

wb = xw.Book("test.xlsm")
io = wb.sheets["Nhập-xuất"]
data = wb.sheets["Data"]
export_field = wb.sheets["Báo cáo"]
table = None
export_table = None

def fetch_top(dict: list) -> set:
    result = []
    for i in dict:
        result.append(i[ct.type_of_payment].lower())
    return set(result)


def header():
    data.range("A1").value = "Ngày"
    data.range("B1").value = "Thu/chi"
    data.range("C1").value = "Số tiền"
    data.range("D1").value = "Tiền mặt/Ngân hàng"
    data.range("E1").value = "Ghi chú"

def ex_header():
    export_field.range("A5").value = "Ngày"
    export_field.range("B5").value = "Thu/chi"
    export_field.range("C5").value = "Số tiền"
    export_field.range("D5").value = "Tiền mặt/Ngân hàng"
    export_field.range("E5").value = "Ghi chú"

def add_data():
    io.range("G7").value = ""
    global table
    if len(data.tables):
        table = data.tables["data"]
    else:
        header_table = xw.Range(data.range("A1"), data.range("E1")).expand()
        table = data.tables.add(source=header_table, name="data")
        header()
    date = io.range("C2").options(dates=dt.date).value
    type_of_transaction =io.range("C3").value
    amount = io.range("C4").value
    type_of_payment = io.range("C5").value
    note = io.range("C6").value
    data.range("2:2").insert("down")
    index = [i + "2" for i in "ABCDE"]
    datum = [date, type_of_transaction, amount, type_of_payment, note]
    if any(j is None for j in datum):
        io.range("G7").value = ct.err
    else:
        for i in range(0,5):
            data.range(index[i]).value = datum[i]


def today():
    io.range("C2").value = dt.date.today()

def export():
    global table
    result = []
    if len(data.tables):
        table = data.tables["data"]
    else:
        header_table = xw.Range(data.range("A1"), data.range("E1")).expand()
        table = data.tables.add(source=header_table, name="data")
        header()
    datum = table.range.options(ps.DataFrame, index=0, header=1).value
    dict_datum = datum.to_dict(orient='records')
    top = list(fetch_top(dict_datum))
    start_date_ref = export_field.range("B2").value if export_field.range("A2").value else dt.date.today()
    end_date_ref = export_field.range("D2").value if export_field.range("D2").value else dt.date.today()
    type_of_transaction_ref = [export_field.range("D3").value.lower()] if export_field.range("D3").value else ["thu", "chi"]
    type_of_payment_ref = [export_field.range("B3").value.lower()] if export_field.range("B3").value else top
    for i in dict_datum:
        if start_date_ref <= i[ct.date].to_pydatetime() and i[ct.date].to_pydatetime() <= end_date_ref and i[ct.type_of_payment].lower() in type_of_payment_ref and i[ct.type_of_transaction].lower() in type_of_transaction_ref:
            result.append(i)
    if not len(result): export_field.range("A5").value = "Không có giá trị thỏa mãn. Vui lòng thử lại."
    else:
        dict_result = {}
        temp = []
        temp2 = []
        temp3 = []
        temp4 = []
        temp5 = []
        for j in result:
            temp.append(j[ct.date])
            temp2.append(j[ct.type_of_transaction])
            temp3.append(j[ct.amount])
            temp4.append(j[ct.type_of_payment])
            temp5.append(j[ct.note])
        dict_result = {"Ngày": temp,
                   "Thu/chi": temp2,
                   "Số tiền": temp3,
                   "Tiền mặt/Ngân hàng": temp4,
                   "Ghi chú": temp5}
        global export_table
        if len(export_field.tables):
            export_table = export_field.tables["baocao"]
        else:
            head = xw.Range("A5:E5")
            export_table = export_field.tables.add(source=head, name="baocao")
            ex_header()
        ps_datum = ps.DataFrame(dict_result)
        export_table.update(ps_datum, index=False)
        last_cell_row = export_table.range.last_cell.row
        total_cell_row = last_cell_row + 2
        last_cell_column = export_table.range.last_cell.column
        total_cell_column  = last_cell_column - 1
        total_cell = xw.Range((total_cell_row, total_cell_column))
        total_cell.value = "Tổng: "
        sum_cell = xw.Range((total_cell_row, total_cell_column+1))
        temp6 = 0
        for i in result:
            if i[ct.type_of_transaction] == "Thu": temp6 += i[ct.amount]
            else: temp6 -= i[ct.amount]
        sum_cell.value = temp6

    for a in result:
        print(a)
    print(type(dict_datum[0]["Ngày"].to_pydatetime().date()))
    print(start_date_ref, end_date_ref, type_of_payment_ref, type_of_transaction_ref, sep="\n")    

def delete_table():
    global table
    if len(data.tables):
        table = data.tables["data"]
    else:
        pass
    if len(data.tables): table.range.value = ""
    else: pass 

def delete_ex_table():
    global export_table
    if len(data.tables):
        export_table = export_field.tables["baocao"]
    else:
        pass
    if len(export_field.tables): 
        export_table.range.value = ""

    else: pass 

def reset():
    index = ["C" + str(i) for i in range(1, 7)]
    for j in index:
        io.range(j).value = ""