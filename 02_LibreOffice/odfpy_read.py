
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

# Create a new OpenDocument Spreadsheet
doc = OpenDocumentSpreadsheet()

# Create a new table (sheet)
table = Table(name="Sheet1")

# Add rows and cells with data
data = [
    ["Name", "Age", "City"],
    ["Alice", 30, "New York"],
    ["Bob", 25, "Chicago"],
]

for row_data in data:
    row = TableRow()
    for cell_data in row_data:
        cell = TableCell()

        # Add text to the cell
        cell_content = P(text=str(cell_data))  # Convert data to string
        cell.addElement(cell_content)

        # Add the cell to the row
        row.addElement(cell)
    # Add the row to the table
    table.addElement(row)

# Add the table to the spreadsheet
doc.spreadsheet.addElement(table)

# Save the ODS file
doc.save(r"PATH_TO_FILE\output.ods")
print("Data written to 'output.ods'")

