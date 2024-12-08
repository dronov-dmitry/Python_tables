# RU
Чтобы написать макрос на Python для Excel и привязать его к кнопке, вы можете использовать библиотеку `xlwings`. Вот пошаговая инструкция:

### 1. Установите библиотеку `xlwings`

```
pip install xlwings
```

### 2. Напишите код на Python

Создайте Python-скрипт для выполнения макроса. Например, пусть он будет называться `my_macro.py`.

Пример кода:

```
import xlwings as xw

def my_macro():
    wb = xw.Book.caller()  # Получаем ссылку на активную книгу Excel
    sheet = wb.sheets[0]  # Берём первый лист
    sheet["A1"].value = "Привет, Excel!"  # Изменяем значение ячейки A1

```

### **3. Включите надстройку `xlwings`**

Чтобы использовать `xlwings` в Excel 2007, необходимо включить соответствующую VBA-надстройку:

1. Откройте Excel 2007.
2. Перейдите в **Кнопка Office** (в верхнем левом углу) и выберите**Параметры Excel** .
3. Перейдите в раздел **Надстройки** .
4. В выпадающем меню "Управление" (внизу окна) выберите**Надстройки Excel** и нажмите **Перейти** .
5. Нажмите **Обзор** и найдите файл `xlwings.xlam`:
   * Файл `xlwings.xlam` обычно находится в директории Python, в папке `Lib\site-packages\xlwings\addin\`.
   * Пример пути: `C:\Python313\Lib\site-packages\xlwings\addin\xlwings.xlam`.
6. Выберите файл и нажмите **OK**.
7. Убедитесь, что надстройка `xlwings` отмечена в списке галочкой, и нажмите **OK**.
8. Откройте редактор VBA (**Alt + F11** ).
9. В меню VBA выберите**Tools → References** (Сервис → Ссылки).
10. Убедитесь, что в списке подключённых ссылок есть`xlwings`.
11. Включите если не была включена

### **4: Настройте `xlwings.conf`**

Создайте файл конфигурации `xlwings.conf` в папке с Excel-файлом. Добавьте в него следующий текст:

```plaintext
"INTERPRETER_WIN","C:\Python313\python.exe"
"PYTHONPATH","C:\Users\Admin\Desktop\01_Excel\Python"
```

Где `PYTHONPATH` - это полный путь к папке, где находится файл `macro.py`.

### 5. Создайте VBA-код для вызова Python-скрипта

Откройте редактор VBA (Alt + F11) и добавьте новый модуль. Вставьте следующий код:

```vba
Sub RunPythonMacro()
    RunPython "import macro ; macro.my_macro()"
End Sub
```

### 6. Добавьте кнопку в Excel

1. Перейдите во вкладку **Разработчик** (Developer).
2. Нажмите **Вставить -> Форма -> Кнопка (Button)** .
3. Нарисуйте кнопку на листе.
4. В появившемся окне "Назначить макрос" выберите `RunPythonMacro` и нажмите "ОК".

### 7. Запустите макрос

Теперь нажмите на кнопку в Excel, и ваш Python-скрипт выполнится.


# ENG

Here is the translation of the text:

To write a macro in Python for Excel and bind it to a button, you can use the `xlwings` library. Here is a step-by-step guide:

### 1. Install the `xlwings` library

```
pip install xlwings
```

### 2. Write the Python code

Create a Python script to perform the macro. For example, let it be called `my_macro.py`.

Example code:

```python
import xlwings as xw

def my_macro():
    wb = xw.Book.caller()  # Get a reference to the active Excel workbook
    sheet = wb.sheets[0]  # Get the first sheet
    sheet["A1"].value = "Hello, Excel!"  # Change the value of cell A1
```

### **3. Enable the `xlwings` add-in**

To use `xlwings` in Excel 2007, you need to enable the corresponding VBA add-in:

1. Open Excel 2007.
2. Go to the **Office Button** (in the top left corner) and select **Excel Options**.
3. Go to the **Add-ins** section.
4. In the "Manage" dropdown (at the bottom of the window), select **Excel Add-ins** and click **Go**.
5. Click **Browse** and find the file `xlwings.xlam`:
   * The `xlwings.xlam` file is usually located in the Python directory, in the `Lib\site-packages\xlwings\addin\` folder.
   * Example path: `C:\Python313\Lib\site-packages\xlwings\addin\xlwings.xlam`.
6. Select the file and click **OK**.
7. Make sure the `xlwings` add-in is checked in the list and click **OK**.
8. Open the VBA editor (**Alt + F11**).
9. In the VBA menu, select **Tools → References**.
10. Ensure that `xlwings` is listed among the references.
11. Enable it if it’s not already enabled.

### **4: Configure `xlwings.conf`**

Create a configuration file `xlwings.conf` in the folder with the Excel file. Add the following text to it:

```plaintext
"INTERPRETER_WIN","C:\Python313\python.exe"
"PYTHONPATH","C:\Users\Admin\Desktop\01_Excel\Python"
```

Where `PYTHONPATH` is the full path to the folder where the `macro.py` file is located.

### 5. Create VBA code to call the Python script

Open the VBA editor (Alt + F11) and add a new module. Insert the following code:

```vba
Sub RunPythonMacro()
    RunPython "import macro ; macro.my_macro()"
End Sub
```

### 6. Add a button in Excel

1. Go to the **Developer** tab.
2. Click **Insert → Form Controls → Button**.
3. Draw a button on the sheet.
4. In the "Assign Macro" window that appears, select `RunPythonMacro` and click "OK".

### 7. Run the macro

Now, click the button in Excel, and your Python script will execute.
