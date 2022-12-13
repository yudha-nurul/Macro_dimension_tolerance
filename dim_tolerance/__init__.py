from comtypes.client import GetActiveObject

# import pymsgbox
ps = GetActiveObject("PowerShape.Application")
ps.Visible = True


def dim_tolerance_plusmin():
    # count = int(ps.evaluate('selection.number'))
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
DIALOG TOLEDIT
TOLVALUE1 0.01
TOLVALUE2 -0.01
TOOLBAR DIMATTRIBUTES DISMISS
else


if (3 < dim_value & dim_value <= 6) {
DIALOG TOLEDIT
TOLVALUE1 0.012
TOLVALUE2 -0.012
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (6 < dim_value & dim_value <= 10) {
DIALOG TOLEDIT
TOLVALUE1 0.015
TOLVALUE2 -0.015
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (10 < dim_value & dim_value <= 18) {
DIALOG TOLEDIT
TOLVALUE1 0.018
TOLVALUE2 -0.018
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (18 < dim_value & dim_value <= 30) {
DIALOG TOLEDIT
TOLVALUE1 0.021
TOLVALUE2 -0.021
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (30 < dim_value & dim_value <= 50) {
DIALOG TOLEDIT
TOLVALUE1 0.025
TOLVALUE2 -0.025
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (50 < dim_value & dim_value <= 80) {
DIALOG TOLEDIT
TOLVALUE1 0.03
TOLVALUE2 -0.03
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (80 < dim_value & dim_value <= 120) {
DIALOG TOLEDIT
TOLVALUE1 0.035
TOLVALUE2 -0.035
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (120 < dim_value & dim_value <= 180) {
DIALOG TOLEDIT
TOLVALUE1 0.04
TOLVALUE2 -0.04
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (180 < dim_value & dim_value <= 250) {
DIALOG TOLEDIT
TOLVALUE1 0.046
TOLVALUE2 -0.046
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (250 < dim_value & dim_value <= 315) {
DIALOG TOLEDIT
TOLVALUE1 0.052
TOLVALUE2 -0.052
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (315 < dim_value & dim_value <= 400) {
DIALOG TOLEDIT
TOLVALUE1 0.057
TOLVALUE2 -0.057
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

if (400 < dim_value & dim_value <= 500) {
DIALOG TOLEDIT
TOLVALUE1 0.063
TOLVALUE2 -0.063
TOOLBAR DIMATTRIBUTES DISMISS
} else {
}

DIALOG
TOLEDIT
TOLSTYLE
Line
tolerance
TOOLBAR
DIMATTRIBUTES
DISMISS
"""
