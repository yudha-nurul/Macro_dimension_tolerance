"""
pshape = GetActiveObject("PowerShape.Application")
pshape.Visible = True
# Connect to PowerShape
# pshape = window.external;
pshape.Exec('LEVEL FILTER USED_OFF')
LEVEL_ON = pshape.Evaluate('level.filtered.number')
pymsgbox.alert(LEVEL_ON)
"""
from comtypes.client import GetActiveObject
import pymsgbox

pshape = GetActiveObject("PowerShape.Application")