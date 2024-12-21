В Python через библиотеку `pyuno` вы можете взаимодействовать с LibreOffice Calc, используя команды, которые вызываются через UNO API. Эти команды не являются встроенными функциями Python, а соответствуют интерфейсам и методам API LibreOffice.  

Вот основные команды и функции, которые можно использовать в LibreOffice Calc через Python:

---

### **1. Работа с документом Calc**
- **Создание нового документа**:
  ```python
  desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
  document = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
  ```

- **Открытие существующего файла**:
  ```python
  file_url = "file:///path/to/file.ods"
  document = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
  ```

- **Сохранение документа**:
  ```python
  document.store()
  ```

- **Сохранение как (с указанием нового пути)**:
  ```python
  new_url = "file:///path/to/newfile.ods"
  document.storeAsURL(new_url, ())
  ```

---

### **2. Работа с листами**
- **Получение всех листов**:
  ```python
  sheets = document.Sheets
  ```

- **Получение конкретного листа по имени**:
  ```python
  sheet = sheets.getByName("Sheet1")
  ```

- **Добавление нового листа**:
  ```python
  sheets.insertNewByName("NewSheet", 0)  # Добавить на позицию 0
  ```

- **Удаление листа**:
  ```python
  sheets.removeByName("Sheet1")
  ```

---

### **3. Работа с ячейками**
- **Доступ к конкретной ячейке**:
  ```python
  cell = sheet.getCellByPosition(0, 0)  # Первая ячейка A1
  ```

- **Изменение значения ячейки**:
  ```python
  cell.Value = 123  # Для чисел
  cell.String = "Привет"  # Для текста
  ```

- **Получение значения ячейки**:
  ```python
  value = cell.Value
  text = cell.String
  ```

- **Работа с диапазонами**:
  ```python
  range_cells = sheet.getCellRangeByPosition(0, 0, 2, 2)  # Диапазон от A1 до C3
  ```

- **Установка значений для диапазона**:
  ```python
  range_cells.setDataArray([[1, 2, 3], [4, 5, 6], [7, 8, 9]])
  ```

- **Чтение значений из диапазона**:
  ```python
  data = range_cells.getDataArray()
  ```

---

### **4. Форматирование**
- **Изменение шрифта ячейки**:
  ```python
  cell.CharHeight = 14  # Размер шрифта
  cell.CharWeight = com.sun.star.awt.FontWeight.BOLD  # Жирный текст
  ```

- **Изменение цвета ячейки**:
  ```python
  cell.CellBackColor = 0xFF0000  # Красный цвет
  ```

- **Выравнивание текста**:
  ```python
  cell.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER  # По центру
  cell.VertJustify = com.sun.star.table.CellVertJustify.CENTER
  ```

---

### **5. Формулы**
- **Установка формулы в ячейке**:
  ```python
  cell.Formula = "=SUM(A1:A10)"
  ```

- **Получение формулы из ячейки**:
  ```python
  formula = cell.Formula
  ```

---

### **6. Стили**
- **Применение стиля к ячейке**:
  ```python
  cell.CellStyle = "Good"  # Название стиля из LibreOffice
  ```

- **Создание нового стиля**:
  ```python
  styles = document.StyleFamilies.getByName("CellStyles")
  new_style = styles.createStyle("MyStyle")
  new_style.setPropertyValue("CellBackColor", 0xFFFF00)  # Желтый фон
  ```

---

### **7. Макросы и скрипты**
- **Выполнение макроса**:
  ```python
  macro = document.createInstance("com.sun.star.script.provider.ScriptProvider")
  macro.invoke(["vnd.sun.star.script:ModuleName.MacroName?language=Basic&location=document"], (), ())
  ```

---

### **8. Экспорт**
- **Экспорт документа в PDF**:
  ```python
  export_props = (
      {"FilterName": "calc_pdf_Export"},
  )
  document.storeToURL("file:///path/to/file.pdf", export_props)
  ```

---
# **9. com.sun.star.table**
В Python через `pyuno` и модуль `com.sun.star.table` доступны команды и интерфейсы, которые управляют элементами таблиц, ячеек и диапазонов LibreOffice Calc. Эти интерфейсы предоставляют методы для работы с ячейками, диапазонами, строками, столбцами и стилями.  

Вот список доступных классов, интерфейсов и их основных методов:

---

### **1. Основные интерфейсы и классы в `com.sun.star.table`**

#### **`com.sun.star.table.XCell`**
Работа с отдельной ячейкой.
- **Свойства**:
  - `Value` — числовое значение ячейки.
  - `String` — текстовое содержимое ячейки.
  - `Formula` — формула, установленная в ячейке.
- **Методы**:
  - `getError()` — возвращает код ошибки, если ячейка содержит ошибку.

---

#### **`com.sun.star.table.XCellRange`**
Работа с диапазонами ячеек.
- **Методы**:
  - `getCellByPosition(column, row)` — получить ячейку по координатам.
  - `getCellRangeByPosition(startCol, startRow, endCol, endRow)` — получить диапазон ячеек.
  - `getCellRangeByName(rangeName)` — получить диапазон по имени (например, "A1:B10").

---

#### **`com.sun.star.table.CellProperties`**
Настройка свойств ячеек.
- **Свойства**:
  - `CellBackColor` — цвет фона ячейки (в формате HEX, например, `0xFFFFFF`).
  - `IsTextWrapped` — перенос текста в ячейке (булево значение).
  - `HoriJustify` — горизонтальное выравнивание текста (значения из `com.sun.star.table.CellHoriJustify`).
  - `VertJustify` — вертикальное выравнивание текста (значения из `com.sun.star.table.CellVertJustify`).

---

#### **`com.sun.star.table.XCellRangeData`**
Работа с массивами данных для диапазонов.
- **Методы**:
  - `getDataArray()` — возвращает массив значений диапазона (двумерный массив).
  - `setDataArray(array)` — задаёт массив значений диапазона (двумерный массив).

---

#### **`com.sun.star.table.XTableRows` и `com.sun.star.table.XTableColumns`**
Работа со строками и столбцами.
- **Методы**:
  - `getByIndex(index)` — получить строку/столбец по индексу.
  - `insertByIndex(index, count)` — вставить строки/столбцы начиная с индекса.
  - `removeByIndex(index, count)` — удалить строки/столбцы начиная с индекса.

---

#### **`com.sun.star.table.TableBorder`**
Настройка границ таблицы или ячеек.
- **Свойства**:
  - `TopLine`, `BottomLine`, `LeftLine`, `RightLine` — свойства границ ячейки (объекты типа `com.sun.star.table.BorderLine`).
  - `HorizontalLine`, `VerticalLine` — свойства для внутренних границ.

---

#### **`com.sun.star.table.XMergeable`**
Объединение ячеек.
- **Методы**:
  - `merge(true/false)` — объединить или разделить ячейки.

---

#### **`com.sun.star.table.XSortable`**
Сортировка данных в таблице.
- **Методы**:
  - `sort(sortDescriptor)` — выполнить сортировку.  
    Пример дескриптора:
    ```python
    sortDescriptor = (
        {"Field": 0, "Ascending": True},  # Поле для сортировки
    )
    ```

---

#### **`com.sun.star.table.XAutoFormattable`**
Применение автоформатов.
- **Методы**:
  - `autoFormat(formatIndex)` — применить автоформат по индексу.

---

### **2. Константы из `com.sun.star.table`**

#### **Выравнивание текста**
Модуль `com.sun.star.table.CellHoriJustify`:
- `STANDARD` — стандартное.
- `LEFT` — по левому краю.
- `CENTER` — по центру.
- `RIGHT` — по правому краю.

Модуль `com.sun.star.table.CellVertJustify`:
- `STANDARD` — стандартное.
- `TOP` — по верхнему краю.
- `CENTER` — по центру.
- `BOTTOM` — по нижнему краю.

---

#### **Типы границ (`com.sun.star.table.BorderLine`)**
Свойства:
- `LineWidth` — толщина линии.
- `Color` — цвет линии (HEX).
- `LineStyle` — стиль линии (`SOLID`, `DASHED` и др.).

---

#### **Типы содержимого ячеек**
Модуль `com.sun.star.table.CellContentType`:
- `EMPTY` — пустая ячейка.
- `VALUE` — числовое значение.
- `TEXT` — текст.
- `FORMULA` — формула.

# **10. com.sun.star.sheet**

Модуль `com.sun.star.sheet` предоставляет интерфейсы и классы для работы с электронными таблицами LibreOffice Calc. Этот модуль дополняет функциональность `com.sun.star.table`, позволяя работать с формулами, функциями, фильтрами, диаграммами и другими элементами, специфичными для таблиц.

---

### **1. Основные интерфейсы и их методы**

#### **`com.sun.star.sheet.XSpreadsheet`**
Работа с отдельным листом.
- **Методы**:
  - `getCellByPosition(column, row)` — получить ячейку по координатам.
  - `getCellRangeByPosition(startCol, startRow, endCol, endRow)` — получить диапазон ячеек.
  - `getCellRangeByName(rangeName)` — получить диапазон по имени (например, "A1:B10").
  - `getRows()` — получить доступ к строкам (возвращает `XTableRows`).
  - `getColumns()` — получить доступ к столбцам (возвращает `XTableColumns`).

---

#### **`com.sun.star.sheet.XCellRangeFormula`**
Работа с формулами в диапазонах.
- **Методы**:
  - `getFormulaArray()` — возвращает массив формул в диапазоне.
  - `setFormulaArray(formulaArray)` — устанавливает массив формул в диапазоне.

---

#### **`com.sun.star.sheet.XSpreadsheetDocument`**
Работа с документом Calc.
- **Методы**:
  - `getSheets()` — получить коллекцию листов.
  - `createInstance()` — создать новый объект в документе (например, диаграмму).

---

#### **`com.sun.star.sheet.XFunctionAccess`**
Позволяет выполнять встроенные функции Calc.
- **Методы**:
  - `callFunction(functionName, arguments)` — вызывает функцию Calc.  
    Пример:
    ```python
    func_access = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.sheet.FunctionAccess", context
    )
    result = func_access.callFunction("SUM", [[1, 2, 3]])
    ```

---

#### **`com.sun.star.sheet.XCellRangeAddressable`**
Работа с адресами диапазонов.
- **Методы**:
  - `getRangeAddress()` — возвращает адрес диапазона (объект `com.sun.star.table.CellRangeAddress`).

---

#### **`com.sun.star.sheet.XSheetCellRange`**
Объединяет работу с ячейками, диапазонами и их форматированием.
- **Методы**:
  - `getSpreadsheet()` — возвращает лист, которому принадлежит диапазон.

---

#### **`com.sun.star.sheet.XSheetFilterable`**
Фильтрация данных.
- **Методы**:
  - `filter(filterDescriptor)` — применяет фильтр к диапазону.  
    Пример дескриптора:
    ```python
    filterDescriptor = range.createFilterDescriptor()
    filterDescriptor.ContainsHeader = True
    filterDescriptor.FilterFields = (
        {"Field": 0, "Operator": com.sun.star.sheet.FilterOperator.EQUAL, "Value": "Example"},
    )
    range.filter(filterDescriptor)
    ```

---

#### **`com.sun.star.sheet.XSheetOperation`**
Выполнение операций над диапазонами.
- **Методы**:
  - `computeFunction(function)` — вычисляет функцию для диапазона (например, сумму или среднее).  
    Пример:
    ```python
    result = range.computeFunction(com.sun.star.sheet.GeneralFunction.SUM)
    ```

---

#### **`com.sun.star.sheet.XNamedRanges`**
Работа с именованными диапазонами.
- **Методы**:
  - `addNewByName(name, range, content, type)` — создаёт именованный диапазон.
  - `getByName(name)` — возвращает диапазон по имени.

---

#### **`com.sun.star.sheet.XDataPilotTables`**
Работа с таблицами сводных данных (DataPilot).
- **Методы**:
  - `insertNewByName(name, position, sourceRange)` — создать новую сводную таблицу.
  - `removeByName(name)` — удалить сводную таблицу.

---

### **2. Константы из `com.sun.star.sheet`**

#### **Функции для вычислений (`GeneralFunction`)**
Используются с методами, такими как `computeFunction`:
- `SUM` — сумма.
- `AVERAGE` — среднее.
- `MAX` — максимум.
- `MIN` — минимум.
- `COUNT` — количество значений.

---

#### **Операторы фильтрации (`FilterOperator`)**
Используются для создания фильтров:
- `EQUAL` — равно.
- `NOT_EQUAL` — не равно.
- `GREATER` — больше.
- `LESS` — меньше.
- `BETWEEN` — между.
- `NOT_BETWEEN` — не между.

---

#### **Пример 2: Создание фильтра**
```python
def apply_filter(*args):
    # Получение текущего документа
    document = XSCRIPTCONTEXT.getDocument()
    sheet = document.Sheets.getByIndex(0)  # Получаем лист по индексу (0 для первого листа)

    # Определение диапазона ячеек для фильтрации (D1:E8)
    cell_range = sheet.getCellRangeByName("D1:E8")

    # Создание фильтра
    filter_descriptor = cell_range.createFilterDescriptor(True)

    # Установка параметров фильтра
    filter_descriptor.ContainsHeader = True

    # Создание поля фильтра
    filter_field = uno.createUnoStruct("com.sun.star.sheet.TableFilterField")
    filter_field.Field = 0  # Первый столбец (D)
    filter_field.Operator = 7  # 3 - equal, 2 - not equal, 4 - less and equal, 5 - more and equal, 6 - less, 7 - more
    filter_field.IsNumeric = True  # Указание, что это числовое значение
    filter_field.NumericValue = 2021 # datetime(2021, 7, 1)  # Устанавливаем значение даты (01/07/2021)

    # Применение фильтра
    filter_descriptor.FilterFields = (filter_field,)
    cell_range.filter(filter_descriptor)
```

---

#### **Пример 3: Работа с формулами**
```python
def set_formula():
    document = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
    sheet = document.Sheets.getByIndex(0)

    # Установка формулы
    cell = sheet.getCellByPosition(0, 0)  # A1
    cell.Formula = "=SUM(A2:A10)"

    # Чтение результата формулы
    print("Результат формулы:", cell.Value)
```

---
# **11. com.sun.star.frame**

Модуль `com.sun.star.frame` предоставляет интерфейсы и классы для управления рамками (фреймами) и окнами LibreOffice, а также для взаимодействия с документами, управления их жизненным циклом, макросами и действиями в пользовательском интерфейсе.

---

### **1. Основные интерфейсы и их методы**

#### **`com.sun.star.frame.XComponentLoader`**
Один из ключевых интерфейсов для загрузки документов.
- **Методы**:
  - `loadComponentFromURL(URL, targetFrameName, searchFlags, properties)` — загружает документ.
    - `URL` — путь к файлу, например, `"file:///path/to/file.odt"` или `"private:factory/scalc"`.
    - `targetFrameName` — имя целевого фрейма (например, `"_blank"` для нового окна).
    - `searchFlags` — обычно `0`.
    - `properties` — массив свойств загрузки.
    
    **Пример**:
    ```python
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    document = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
    ```

---

#### **`com.sun.star.frame.XDispatchHelper`**
Позволяет выполнять команды (например, из меню или панелей инструментов).
- **Методы**:
  - `executeDispatch(frame, URL, targetFrameName, searchFlags, arguments)` — выполняет действие, указанное в `URL`.

    **Пример**:
    ```python
    dispatch_helper = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", context
    )
    dispatch_helper.executeDispatch(document.getCurrentController().getFrame(), ".uno:Save", "", 0, ())
    ```

---

#### **`com.sun.star.frame.XFrame`**
Представляет собой окно или фрейм.
- **Свойства**:
  - `Component` — текущий компонент в фрейме (например, документ).
  - `Name` — имя фрейма.
- **Методы**:
  - `getController()` — получить контроллер текущего компонента.
  - `activate()` — активировать фрейм.

---

#### **`com.sun.star.frame.XModel`**
Представляет модель документа.
- **Свойства**:
  - `Title` — заголовок документа.
  - `URL` — путь к документу.
- **Методы**:
  - `getCurrentController()` — возвращает текущий контроллер документа.
  - `getArgs()` — получить параметры документа.

---

#### **`com.sun.star.frame.XController`**
Управляет отображением и состоянием документа.
- **Методы**:
  - `getViewData()` — получить данные о текущем состоянии вида.
  - `restoreViewData(data)` — восстановить состояние вида.
  - `getFrame()` — вернуть фрейм, связанный с контроллером.

---

#### **`com.sun.star.frame.XDesktop`**
Работа с окнами LibreOffice.
- **Методы**:
  - `getComponents()` — возвращает коллекцию всех открытых документов.
  - `getCurrentComponent()` — возвращает текущий активный документ.
  - `terminate()` — завершить работу LibreOffice.

    **Пример**:
    ```python
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    current_doc = desktop.getCurrentComponent()
    ```

---

#### **`com.sun.star.frame.XStorable`**
Интерфейс для сохранения документов.
- **Методы**:
  - `store()` — сохранить текущий документ.
  - `storeAsURL(URL, properties)` — сохранить как (с указанием нового пути).
  - `storeToURL(URL, properties)` — экспортировать документ.

    **Пример**:
    ```python
    document.store()
    document.storeAsURL("file:///path/to/newfile.ods", ())
    ```

---

#### **`com.sun.star.frame.XDispatchProvider`**
Позволяет получать доступ к командам для выполнения.
- **Методы**:
  - `queryDispatch(URL, targetFrameName, searchFlags)` — возвращает объект команды для выполнения.
  - `queryDispatches(descriptors)` — возвращает массив объектов команд.

---

#### **`com.sun.star.frame.XTerminateListener`**
Позволяет реагировать на завершение работы LibreOffice.
- **Методы**:
  - `queryTermination()` — проверка перед завершением работы.
  - `notifyTermination()` — уведомление о завершении.

---

### **2. Константы из `com.sun.star.frame`**

#### **Типы фреймов (`FrameSearchFlag`)**
Используются в методе `loadComponentFromURL`:
- `SELF` — текущий фрейм.
- `PARENT` — родительский фрейм.
- `CHILDREN` — дочерние фреймы.

---

#### **Команды (`Dispatch URLs`)**
Используются в `XDispatchHelper` или `XDispatchProvider`:
- `.uno:Save` — сохранить документ.
- `.uno:CloseDoc` — закрыть документ.
- `.uno:Print` — напечатать документ.
- `.uno:Quit` — выйти из LibreOffice.
- `.uno:Undo` — отменить действие.
- `.uno:Redo` — повторить действие.
- `.uno:Copy` — скопировать выделенное.
- `.uno:Paste` — вставить из буфера.

---

### **3. Примеры использования `com.sun.star.frame`**

#### **Пример 1: Открытие нового документа Calc**
```python
import uno

def open_calc():
    # Подключение к LibreOffice
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    
    # Открытие нового документа
    document = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
    print("Новый документ Calc открыт.")

if __name__ == "__main__":
    open_calc()
```

---

#### **Пример 2: Выполнение команды "Сохранить"**
```python
def save_document():
    document = desktop.getCurrentComponent()
    if document is not None:
        # Выполнение команды сохранения через DispatchHelper
        dispatch_helper = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", context
        )
        dispatch_helper.executeDispatch(
            document.getCurrentController().getFrame(),
            ".uno:Save",
            "",
            0,
            ()
        )
        print("Документ сохранён.")
```

---

#### **Пример 3: Завершение LibreOffice**
```python
def terminate_libreoffice():
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    desktop.terminate()
    print("LibreOffice завершён.")
```

---
