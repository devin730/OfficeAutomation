import xlrd

# learning some useful APIs

data = xlrd.open_workbook('./data/info.xlsx')
# print(data.sheet_names())
table = data.sheet_by_index(0)
# print(table.name, table.nrows, table.ncols)
# table.cell_value(i,j)
# type(table.cell_value(i, j)
# print(table.row_values(0))