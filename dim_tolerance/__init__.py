def dim_tolerance_plusmin():
    count = int(ps.evaluate('selection.number'))
    print(f'jumlah dimensi yang diselect adalah {count}')

    nama = (ps.evaluate('SELECTION.NAMES'))
    nama = nama.replace("{", "")
    nama = nama.replace("}", "")
    nama = nama.split(';')
    namanya = nama[:]
    print(nama)
    for i in nama:
        print(f'ini adalah nilai dari dimensi {i}')
        dim_value = ps.evaluate(f'dimension[{i}].value')
        print(dim_value)

        if 0 < dim_value & dim_value <= 3:
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
