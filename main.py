"""
cara 1 :
import pymsgbox
# import win32com.client
#powershape = win32com.client.Dispatch ("PowerShape.Application")
ps = win32com.client.GetActiveObject ("PowerShape.Application")
ps.Visible = True
----------------------------
cara 2 :
pshape = GetActiveObject("PowerShape.Application")
pshape.Visible = True
# Connect to PowerShape
# pshape = window.external;
pshape.Exec('LEVEL FILTER USED_OFF')
LEVEL_ON = pshape.Evaluate('level.filtered.number')
pymsgbox.alert(LEVEL_ON)
--------------
# convert str to list
print(a)
print(len(a))
print(type(a))
a = a.replace("{", "")
a = a.replace("}", "")
print(a)
b = a.split(';')
print(b)
print(len(b))
print(type(b))
----------------
"""
import sys

from dim_tolerance import dim_tolerance_plusmin as plusmin, run_plusmin
import json

if __name__ == '__main__':
    plusmin()
    run_plusmin()

# count = (f'selection.name[{count}]')
# print(count)
# get_dim_names = ps.evaluate(f'selection.name[{count}]')
# get_dim_names = int(ps.evaluate(f'{count}'))
# print(get_dim_names)
# dim_value = ps.evaluate(f'dimension[{get_dim_names}].value')
# print(type(dim_value))
# print(dim_value)
# print(f'jumlahnya adalah {count}')
# print(f'namanya adalah {dim_name}')
# print(f'nilainya adalah {dim_value}')
