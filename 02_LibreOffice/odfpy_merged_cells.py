from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell, CoveredTableCell
from odf.text import P

# Create a new spreadsheet document
doc = OpenDocumentSpreadsheet()

# Create a new table (sheet)
table = Table(name="Sheet1")

# Add a row with merged cells
row = TableRow()

# Main cell that spans across 3 columns
# main_cell = TableCell(numbercolumnsrepeated=3)
main_cell = TableCell(numberrowsspanned=1, numbercolumnsspanned=3)
main_cell.addElement(P(text="Merged Cell"))
row.addElement(main_cell)

# Add covered cells (required for merged areas)
for _ in range(2):  # Add two covered cells to complete the merge
    covered_cell = CoveredTableCell()
    row.addElement(covered_cell)

# Add the row to the table
table.addElement(row)

# Add a normal row below for comparison
normal_row = TableRow()
for value in ["A", "B", "C"]:
    cell = TableCell()
    cell.addElement(P(text=value))
    normal_row.addElement(cell)
table.addElement(normal_row)

# Add the table to the spreadsheet
doc.spreadsheet.addElement(table)

# Save the document
doc.save(r"C:\Users\DmitryD\Desktop\merged_cells.ods")
print("ODS file with merged cells created as 'merged_cells.ods'")
