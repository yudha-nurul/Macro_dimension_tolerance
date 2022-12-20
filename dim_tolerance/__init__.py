# from comtypes.client import GetActiveObject
import sys
import os
from contextlib import redirect_stdout

# import pymsgbox
import win32com.client
ps = win32com.client.Dispatch("PowerShape.Application")
ps.Visible = True

def dim_tolerance_plusmin():
    count = int(ps.evaluate('selection.number'))
    if count == 0:
        print('select at leas 1 dimension')
        sys.exit()
    # print(f'jumlah dimensi yang diselect adalah {count}')

    nama = (ps.evaluate('SELECTION.NAMES'))
    nama = nama.replace(" ", "")
    nama = nama.replace("{", "")
    nama = nama.replace("}", "")
    nama = nama.split(';')
    # print(nama) # this is list of dimension names
    for i in nama:
        print('select clearlist')
        print(f'add dimension "{i}"')
        dim_value = ps.evaluate(f'dimension[{i}].value')
        # print(dim_value)
        # print(type(dim_value))

        if 0 < dim_value <= 3:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.01')
            print('TOLVALUE2 - 0.01')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 3 < dim_value <= 6:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.012')
            print('TOLVALUE2 - 0.012')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 6 < dim_value <= 10:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.015')
            print('TOLVALUE2 - 0.015')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 10 < dim_value <= 18:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.018')
            print('TOLVALUE2 - 0.018')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 18 < dim_value <= 30:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.021')
            print('TOLVALUE2 - 0.021')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 30 < dim_value <= 50:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.025')
            print('TOLVALUE2 - 0.025')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 50 < dim_value <= 80:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.03')
            print('TOLVALUE2 - 0.03')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 80 < dim_value <= 120:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.035')
            print('TOLVALUE2 - 0.035')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 120 < dim_value <= 180:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.04')
            print('TOLVALUE2 - 0.04')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 180 < dim_value <= 250:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.046')
            print('TOLVALUE2 - 0.046')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 250 < dim_value <= 315:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.052')
            print('TOLVALUE2 - 0.052')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 315 < dim_value <= 400:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.057')
            print('TOLVALUE2 - 0.057')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 400 < dim_value <= 500:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.063')
            print('TOLVALUE2 - 0.063')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        print('select clearlist')


# blm selesai yg ini

def dim_tolerance_plusplus():
    count = int(ps.evaluate('selection.number'))
    if count == 0:
        print('select at leas 1 dimension')
        sys.exit()
    # print(f'jumlah dimensi yang diselect adalah {count}')

    nama = (ps.evaluate('SELECTION.NAMES'))
    nama = nama.replace(" ", "")
    nama = nama.replace("{", "")
    nama = nama.replace("}", "")
    nama = nama.split(';')
    # print(nama) # this is list of dimension names
    for i in nama:
        print('select clearlist')
        print(f'add dimension "{i}"')
        dim_value = ps.evaluate(f'dimension[{i}].value')
        # print(dim_value)
        # print(type(dim_value))

        if 0 < dim_value <= 3:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.01')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 3 < dim_value <= 6:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.012')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 6 < dim_value <= 10:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.015')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 10 < dim_value <= 18:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.018')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 18 < dim_value <= 30:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.021')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 30 < dim_value <= 50:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.025')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 50 < dim_value <= 80:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.03')
            print('TOLVALUE2 0.01')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 80 < dim_value <= 120:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.035')
            print('TOLVALUE2 0.01')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 120 < dim_value <= 180:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.04')
            print('TOLVALUE2 0.02')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 180 < dim_value <= 250:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.046')
            print('TOLVALUE2 0.02')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 250 < dim_value <= 315:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.052')
            print('TOLVALUE2 0.02')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 315 < dim_value <= 400:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.057')
            print('TOLVALUE2 0.03')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 400 < dim_value <= 500:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 0.063')
            print('TOLVALUE2 0.03')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        print('select clearlist')
def dim_tolerance_minmin():
    count = int(ps.evaluate('selection.number'))
    if count == 0:
        print('select at leas 1 dimension')
        sys.exit()
    # print(f'jumlah dimensi yang diselect adalah {count}')

    nama = (ps.evaluate('SELECTION.NAMES'))
    nama = nama.replace(" ", "")
    nama = nama.replace("{", "")
    nama = nama.replace("}", "")
    nama = nama.split(';')
    # print(nama) # this is list of dimension names
    for i in nama:
        print('select clearlist')
        print(f'add dimension "{i}"')
        dim_value = ps.evaluate(f'dimension[{i}].value')
        # print(dim_value)
        # print(type(dim_value))

        if 0 < dim_value <= 3:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.01')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 3 < dim_value <= 6:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.012')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 6 < dim_value <= 10:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.015')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 10 < dim_value <= 18:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.018')
            print('TOLVALUE2 0')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 18 < dim_value <= 30:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.021')
            print('TOLVALUE2 -0.01')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 30 < dim_value <= 50:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.025')
            print('TOLVALUE2 -0.01')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 50 < dim_value <= 80:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.03')
            print('TOLVALUE2 -0.01')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 80 < dim_value <= 120:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.035')
            print('TOLVALUE2 -0.02')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 120 < dim_value <= 180:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.04')
            print('TOLVALUE2 -0.02')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 180 < dim_value <= 250:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.046')
            print('TOLVALUE2 -0.02')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 250 < dim_value <= 315:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.052')
            print('TOLVALUE2 -0.02')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 315 < dim_value <= 400:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.057')
            print('TOLVALUE2 -0.03')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        elif 400 < dim_value <= 500:
            print('DIALOG TOLEDIT')
            print('TOLVALUE1 -0.063')
            print('TOLVALUE2 -0.03')
            print('TOOLBAR DIMATTRIBUTES DISMISS')
        print('select clearlist')
def go_plusmin():
    cwd = os.getcwd()
    with open('go.mac', 'w') as file:
        with redirect_stdout(file):
            run = f'{dim_tolerance_plusmin()}'
    assert isinstance(ps, object)
    ps.exec(f'MACRO RUN "{cwd}\go.mac"')
def go_plusplus():
    cwd = os.getcwd()
    with open('go.mac', 'w') as file:
        with redirect_stdout(file):
            run = f'{dim_tolerance_plusplus()}'
    assert isinstance(ps, object)
    ps.exec(f'MACRO RUN "{cwd}\go.mac"')
def go_minmin():
    cwd = os.getcwd()
    with open('go.mac', 'w') as file:
        with redirect_stdout(file):
            run = f'{dim_tolerance_minmin()}'
    assert isinstance(ps, object)
    ps.exec(f'MACRO RUN "{cwd}\go.mac"')
"""
-----------------------------------------------
# this is original format from ps :
add Dimension "8"
select clearlist
DIALOG TOLEDIT
TOLVALUE1 0.01
TOLVALUE2 -0.01
TOOLBAR DIMATTRIBUTES DISMISS
select clearlist
-----------------------------------------------
"""