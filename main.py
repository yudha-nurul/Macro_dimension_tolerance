"""
cara 1 :
import pymsgbox
import win32com.client
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
"""
from comtypes.client import GetActiveObject
# import pymsgbox

ps = GetActiveObject("PowerShape.Application")
ps.Visible = True
count = ps.evaluate('selection.number')
dim_list_names = ps.evaluate('selection.name[0]')
dimdim = int(dim_list_names)


dim_value = ps.evaluate(f'dimension[{dimdim}].value')
# print(type(dim_value))
print(dim_value)
# print(f'jumlahnya adalah {count}')
# print(f'namanya adalah {dim_name}')
# print(f'nilainya adalah {dim_value}')
