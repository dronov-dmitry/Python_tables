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

```
# Работа с данными
sht.range('A1').options(transpose=True).value  # Транспонировать
sht.range('A1').options(numbers=int).value  # Указать тип данных
sht.range('A1').table.range  # Получить диапазон таблицы

# Форматирование
rng.api.NumberFormat = '0.00'  # Формат чисел
rng.column_width = 20  # Ширина столбца
rng.row_height = 15  # Высота строки
rng.api.Borders.Weight = 2  # Границы

# Работа с книгой
wb.macro('macro_name').run()  # Запуск макроса
app = xw.apps.active  # Активное приложение
app.visible = True/False  # Видимость Excel
wb.fullname  # Полный путь к файлу

# Специальные операции
sht.pictures.add('image.png')  # Добавить картинку
sht.charts.add()  # Добавить график
sht.range('A1').autofit()  # Автоподбор размера
sht.used_range  # Использованный диапазон

# Работа с выделением
rng = sht.range('A1').end('down')  # До последней ячейки вниз
rng = sht.range('A1').end('right')  # До последней ячейки вправо
rng = sht.range('A1').offset(rows=1, cols=1)  # Смещение

# Фильтры и сортировка
rng.api.AutoFilter()  # Включить автофильтр
rng.api.Sort(Key1=rng)  # Сортировка

# Формулы
rng.formula = '=SUM(A1:A10)'  # Записать формулу
rng.formula_array = '{=TRANSPOSE(A1:D1)}'  # Формула массива

# События
@xw.func  # Пользовательская функция
@xw.sub  # Макрос
@xw.arg()  # Аргументы функции

# Дополнительно
sht.cells.last_cell  # Последняя ячейка
sht.range('A1').copy()  # Копировать
sht.range('B1').paste()  # Вставить
sht.api.PageSetup  # Настройки страницы
wb.app.calculation = 'manual'  # Управление пересчетом

# Условное форматирование
rng.api.FormatConditions.Add(Type=3, Operator=1, Formula1="=A1<0")
rng.api.FormatConditions(1).Interior.Color = 255  # Красный цвет

# Работа с именованными диапазонами
wb.names.add('MyRange', '=Sheet1!A1:B10')
wb.names['MyRange'].refers_to
wb.names['MyRange'].delete()

# Защита
sht.api.Protect(Password='password')
sht.api.Unprotect('password')

# Дополнительные свойства ячеек
rng.wrap_text = True  # Перенос текста
rng.merge()  # Объединить ячейки
rng.api.Validation  # Проверка данных
rng.api.Comment  # Комментарии

# Печать
wb.api.PrintOut()  # Печать
sht.api.PageSetup.Orientation = 2  # Альбомная

# Работа с диаграммами
chart = sht.charts.add()
chart.set_source_data(sht.range('A1:D5'))
chart.chart_type = 'line'
chart.api.SetElement(2)  # Заголовок
chart.api.ChartTitle.Text = "График"

# Продвинутые операции с данными
rng.options(ndim=2).value  # Многомерные массивы
rng.options(empty='NA').value  # Обработка пустых ячеек
rng.options(dates=dt.date).value  # Даты

# VBA интеграция
wb.vba.modules['Module1']  # Доступ к модулям
wb.vba.references  # Ссылки VBA
wb.vba.collection  # Коллекции

# Управление приложением
app.screen_updating = False  # Отключить обновление экрана
app.display_alerts = False  # Отключить оповещения
app.quit()  # Закрыть Excel

# Работа с группировкой
sht.api.Outline.ShowLevels(RowLevels=2)
rng.api.Group()
rng.api.Ungroup()

# Стили ячеек
rng.api.Style = "Good"
rng.api.ApplyOutlineStyles()

# Сводные таблицы
pt = wb.api.PivotCaches().Create()
pt.CreatePivotTable(TableDestination="Sheet2!R1C1")

# Расширенное форматирование
rng.api.Borders(7).LineStyle = 1
rng.api.Borders(7).Weight = 2
rng.api.Interior.Pattern = -4124

# Гиперссылки
rng.add_hyperlink('http://example.com', 'Click here')
rng.hyperlink = None  # Удалить ссылку

# Проверка данных
rng.api.Validation.Add(Type=3, 
                     AlertStyle=1,
                     Operator=1,
                     Formula1="=Sheet1!$A$1:$A$10")

# Условные операторы
if rng.value is None:
   pass
if rng.color == (255, 0, 0):  # RGB
   pass

# Настройка окна
app.api.WindowState = -4137  # Развернуть
wb.api.Windows(1).Zoom = 85  # Масштаб
```
