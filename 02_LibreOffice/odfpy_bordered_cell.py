from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TableCellProperties

# Создаем новый файл ODS
doc = OpenDocumentSpreadsheet()

# Создаем стиль для ячеек с границами
border_style = Style(name="BorderStyle", family="table-cell")
border_properties = TableCellProperties(
    border="0.05pt solid #FF0000",  # Толщина, стиль и цвет границы
    padding="0.1cm"  # Внутренние отступы
)
border_style.addElement(border_properties)
doc.styles.addElement(border_style)

# Создаем таблицу
table = Table(name="Sheet1")

# Создаем строку
row = TableRow()
for text in ["Name", "Age", "City"]:
    cell = TableCell(stylename=border_style)  # Применяем стиль к ячейке
    cell.addElement(P(text=text))  # Добавляем текст в ячейку
    row.addElement(cell)
table.addElement(row)

# Добавляем таблицу в документ
doc.spreadsheet.addElement(table)

# Сохраняем документ
doc.save(r"C:\Users\DmitryD\Desktop\bordered_cells.ods")
print("ODS файл с контурами ячеек создан: 'bordered_cells.ods'")

