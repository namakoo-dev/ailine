import sys, openpyxl
path = sys.argv[1]
for mode in ([False, True] if len(sys.argv) > 2 else [False]):
    wb = openpyxl.load_workbook(path, data_only=mode)
    print("=== data_only" if mode else "=== formulas", path)
    for ws in wb.worksheets:
        print("--- sheet", ws.title, ws.dimensions, "merged:", ws.merged_cells.ranges)
        for row in ws.iter_rows():
            vals = [(c.coordinate, c.value, c.number_format if c.number_format != "General" else "") for c in row if c.value is not None]
            if vals:
                print(vals)
