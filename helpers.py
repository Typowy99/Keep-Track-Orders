from flask import redirect, render_template, session
from functools import wraps
from openpyxl import load_workbook
from flask import send_file

def login_required(f):
    """
    Decorate routes to require login.

    http://flask.pocoo.org/docs/0.12/patterns/viewdecorators/
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if session.get("user_id") is None:
            return redirect("/login")
        return f(*args, **kwargs)
    return decorated_function


def change_excel_file(all_data, isell_list, AREA):
    #load excel file

    data = all_data[1]
    time = all_data[2]
    reg = all_data[3]
    plomb = all_data[4]
    sent_from = all_data[5]
    sent_to = all_data[6]
    address_from = AREA[sent_from]
    address_to = AREA[sent_to]

    workbook = load_workbook(filename="excel_file/list_przewozowy.xlsx")

    #open workbook
    sheet = workbook.active

    #modify the desired cell
    sheet["P9"] = data
    sheet["P10"] = f"{data} {time}"
    sheet["P11"] = reg
    sheet["P12"] = plomb
    sheet["A3"] = address_from[1]
    sheet["A4"] = address_from[2]
    sheet["A5"] = address_from[3]
    sheet["A10"] = address_to[1]
    sheet["A11"] = address_to[2]
    sheet["A12"] = address_to[3]
    sheet["J3"] = f"{address_from[0][0]}{data}/{address_to[0][0]}"

    counter = 16
    for isell in isell_list:
        sheet[f"B{counter}"] = isell
        counter += 1

    #save the file
    return workbook.save(filename="excel_file/list_przewozowy2.xlsx")



