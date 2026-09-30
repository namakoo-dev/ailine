import sys, openpyxl
a = openpyxl.load_workbook(sys.argv[1]); b = openpyxl.load_workbook(sys.argv[2])
for wa in a.worksheets:
    if wa.title not in b.sheetnames:
        print("missing sheet in B:", wa.title); continue
    wb_ = b[wa.title]
    print("sheet", wa.title, "merged A", wa.merged_cells.ranges, "B", wb_.merged_cells.ranges)
    mr = max(wa.max_row, wb_.max_row); mc = max(wa.max_column, wb_.max_column)
    n = 0
    for r in range(1, mr + 1):
        for c in range(1, mc + 1):
            x, y = wa.cell(r, c), wb_.cell(r, c)
            if x.value != y.value or x.number_format != y.number_format or bool(x.font.b) != bool(y.font.b):
                n += 1
                print(x.coordinate, repr(x.value), x.number_format, x.font.b, "->", repr(y.value), y.number_format, y.font.b)
    print("diff cells:", n)
for t in b.sheetnames:
    if t not in a.sheetnames: print("extra sheet in B:", t)
