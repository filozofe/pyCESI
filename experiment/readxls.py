#doc: https://pythonhosted.org/xlrd3/


# Import the xlrd module
import xlrd3 as xlrd

# Open the Workbook
book = xlrd.open_workbook(r"C:\Users\phofmann\OneDrive - Cesi\Bureau\ASR1 suivi.xlsx")

# Open the worksheet
worksheet = book.sheet_by_name("ASR1")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))

print("columns:",worksheet.ncols)
print("rows:   ",worksheet.nrows)

# Iterate the rows and columns
for r in range(0, worksheet.nrows):
    for c in (0,3, 4, 6,12):
        # Print the cell values with tab space
        print(worksheet.cell_value(r, c), end='\t\t\t')
    print("")