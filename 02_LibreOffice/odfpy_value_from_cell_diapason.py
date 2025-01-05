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

# Function to parse range and get values
def get_values_from_range(range_address):
		# Split the range into start and end addresses
		start_address, end_address = range_address.split(":")
		
		# Convert addresses to row and column indices
		start_row, start_col = address_to_indices(start_address)
		end_row, end_col = address_to_indices(end_address)
		# print([end_address, start_row, start_col, end_row, end_col])
		
		rows = table.getElementsByType(TableRow)
		data = []

		# Iterate through the specified range
		for row_index in range(start_row, end_row+1):
				emptyRow = False
				if row_index >= len(rows):
					emptyRow = True
				else:
					row = rows[row_index]
					cells = row.getElementsByType(TableCell)
				row_data = []
				for col_index in range(start_col, end_col+1):
						if emptyRow:
							val = ''
						else:
							if col_index < len(cells):
									cell_content = cells[col_index].getElementsByType(P)
									if len(cell_content) > 0:
										if hasattr(cell_content[0], 'text'):
											val = cell_content[0].text if cell_content else ""
										else:
											val = str(cell_content[0]) if cell_content else ""
									else:
										val = ''
							else:
								val = ''
						row_data.append(val)  # Handle missing columns
						print([row_index, col_index, val])
				data.append(row_data)
		
		return data

# Example usage
start_end_address = "A1:D4"
values = get_values_from_range(start_end_address)
print(f"Values from range {start_end_address}:")
for row in values:
		print(row)
