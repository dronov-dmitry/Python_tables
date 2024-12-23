```
def write_text(fPath, fName, data='msg', ext="txt", encd=''): # 
	completeName = os.path.join(fPath, fName + "." + ext)
	fPathNm = '{}\\{}'.format(fPath, fName)
	with open(completeName, "w") as text_file:
		text_file.write(data)
	return completeName

def my_macro():
	# ЗАПОЛНЯЕМ ЗНАЧЕНИЯ ДЛЯ СРЕЗА ЯЧЕЕК
	sheet.range('B1:B5').value = [1, 2, 3, 4, 5]
	# CREATE CHART
	chart = sheet.charts.add()
	chart.set_source_data(sheet.range('A1:B10'))
	chart.chart_type = 'xy_scatter_lines_no_markers'
	tps = '\n'.join(sorted(xw._xlwindows.chart_types_s2i.keys()))
	pth = write_text(r'C:\Users\DmitryD\Desktop\01_Excel', 'report', tps)
```
