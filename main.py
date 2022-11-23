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
jumlah = ps.evaluate('selection.number')
ps.exec(f'print error {jumlah}')

