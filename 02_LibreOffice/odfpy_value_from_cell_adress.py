from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P
import re

# Function to convert Excel-like cell address (e.g., "A1") to row and column indices
def address_to_indices(cell_address):
		match = re.match(r"([A-Z]+)(\d+)", cell_address)
		if not match:
				raise ValueError(f"Invalid cell address: {cell_address}")
		
		col_letters, row_number = match.groups()
		col_index = sum((ord(char) - ord('A') + 1) * (26 ** i) for i, char in enumerate(reversed(col_letters))) - 1
		row_index = int(row_number) - 1
		return row_index, col_index

# Load the ODS file
doc = load(r"C:\Users\DmitryD\Desktop\example.ods")

# Specify the table name
table_name = "Sheet1"

# Find the table by name
table = None
for t in doc.spreadsheet.getElementsByType(Table):
		if t.getAttribute("name") == table_name:
				table = t
				break

if table is None:
		raise ValueError(f"Table '{table_name}' not found!")

# Function to get cell value by address
def get_cell_value_by_address(cell_address):
		# Convert the cell address to row and column indices
		row_index, col_index = address_to_indices(cell_address)

		rows = table.getElementsByType(TableRow)
		if row_index >= len(rows):
				raise ValueError(f"Row index {row_index} is out of range.")

		row = rows[row_index]
		cells = row.getElementsByType(TableCell)

		if col_index >= len(cells):
				raise ValueError(f"Column index {col_index} is out of range.")

		# Get the cell content
		cell_content = cells[col_index].getElementsByType(P)
		if hasattr(cell_content[0], 'text'):
			return cell_content[0].text if cell_content else ""
		else:
			return str(cell_content[0]) if cell_content else ""

# Example usage
cell_address = "A1"
value = get_cell_value_by_address(cell_address)
print(f"Value at {cell_address}: {value}")
