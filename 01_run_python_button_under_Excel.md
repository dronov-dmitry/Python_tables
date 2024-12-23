# UKR

Щоб написати макрос на Python для Excel і прив'язати його до кнопки, ви можете використати бібліотеку xlwings. Ось покрокова інструкція:

1. Встановіть бібліотеку xlwings
```
pip install xlwings
```

2. Напишіть код на Python
Створіть Python-скрипт для виконання макросу. Наприклад, нехай він буде називатися my_macro.py.

Приклад коду:

```python
import xlwings as xw

def my_macro():
    wb = xw.Book.caller()  # Отримуємо посилання на активну книгу Excel
    sheet = wb.sheets[0]  # Беремо перший лист
    sheet["A1"].value = "Привіт, Excel!"  # Змінюємо значення клітинки A1
```

3. Увімкніть надбудову xlwings
Щоб використовувати xlwings в Excel 2007, потрібно увімкнути відповідну VBA-надбудову:

- Відкрийте Excel 2007.
- Перейдіть у Кнопка Office (в верхньому лівому куті) та виберіть Параметри Excel.
- Перейдіть у розділ Надбудови.
- У випадаючому меню "Управління" (внизу вікна) виберіть Надбудови Excel та натисніть Перейти.
- Натисніть Обзор і знайдіть файл xlwings.xlam:
  Файл xlwings.xlam зазвичай знаходиться в директорії Python, у папці `Lib\site-packages\xlwings\addin\`.
  Приклад шляху: `C:\Python313\Lib\site-packages\xlwings\addin\xlwings.xlam`.
- Виберіть файл і натисніть OK.
- Переконайтесь, що надбудова xlwings відзначена в списку галочкою, та натисніть OK.
- Відкрийте редактор VBA (Alt + F11).
- У меню VBA виберіть Tools → References (Сервіс → Посилання).
- Переконайтесь, що в списку підключених посилань є xlwings. Увімкніть, якщо не була увімкнена.

4. Налаштуйте xlwings.conf
Створіть файл конфігурації xlwings.conf у папці з Excel-файлом. Додайте до нього наступний текст:

```
"INTERPRETER_WIN","C:\Python313\python.exe"
"PYTHONPATH","C:\Users\Admin\Desktop\01_Excel\Python"
```
де `PYTHONPATH` — це повний шлях до папки, де знаходиться файл `macro.py`.

5. Створіть VBA-код для виклику Python-скрипта
Відкрийте редактор VBA (Alt + F11) і додайте новий модуль. Вставте наступний код:

```vba
Sub RunPythonMacro()
    RunPython "import macro ; macro.my_macro()"
End Sub
```
or
```vba
Sub RunPythonScript()
    On Error GoTo ErrorHandler
    
    ' Ensure xlwings add-in is installed and active
    Application.Run "xlwings.xlam!RunPython", "import my_macro; my_macro.my_macro()"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
```

6. Додайте кнопку в Excel
- Перейдіть на вкладку Розробник (Developer).
- Натисніть Вставити -> Форма -> Кнопка (Button).
- Намалюйте кнопку на листі.
- У вікні "Призначити макрос" виберіть `RunPythonMacro` і натисніть "OK".

7. Запустіть макрос
Тепер натисніть на кнопку в Excel, і ваш Python-скрипт виконається.

---

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
or 
```vba
Sub RunPythonScript()
    On Error GoTo ErrorHandler
    
    ' Ensure xlwings add-in is installed and active
    Application.Run "xlwings.xlam!RunPython", "import my_macro; my_macro.my_macro()"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
```

### 6. Добавьте кнопку в Excel

1. Перейдите во вкладку **Разработчик** (Developer).
2. Нажмите **Вставить -> Форма -> Кнопка (Button)** .
3. Нарисуйте кнопку на листе.
4. В появившемся окне "Назначить макрос" выберите `RunPythonMacro` и нажмите "ОК".

### 7. Запустите макрос

Теперь нажмите на кнопку в Excel, и ваш Python-скрипт выполнится.

---

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
or 
```vba
Sub RunPythonScript()
    On Error GoTo ErrorHandler
    
    ' Ensure xlwings add-in is installed and active
    Application.Run "xlwings.xlam!RunPython", "import my_macro; my_macro.my_macro()"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
```

### 6. Add a button in Excel

1. Go to the **Developer** tab.
2. Click **Insert → Form Controls → Button**.
3. Draw a button on the sheet.
4. In the "Assign Macro" window that appears, select `RunPythonMacro` and click "OK".

### 7. Run the macro

Now, click the button in Excel, and your Python script will execute.

---
# DEU

Hier ist die Übersetzung ins Deutsche:

Um ein Makro in Python für Excel zu schreiben und es mit einer Schaltfläche zu verknüpfen, können Sie die Bibliothek xlwings verwenden. Hier ist eine Schritt-für-Schritt-Anleitung:

1. Installieren Sie die xlwings-Bibliothek
```
pip install xlwings
```

2. Schreiben Sie den Python-Code
Erstellen Sie ein Python-Skript für das Makro. Zum Beispiel nennen wir es `my_macro.py`.

Beispielcode:

```python
import xlwings as xw

def my_macro():
    wb = xw.Book.caller()  # Ruft das aktive Excel-Dokument ab
    sheet = wb.sheets[0]  # Wählt das erste Blatt
    sheet["A1"].value = "Hallo, Excel!"  # Ändert den Wert der Zelle A1
```

3. Aktivieren Sie das xlwings-Add-In
Um xlwings in Excel 2007 zu verwenden, müssen Sie das entsprechende VBA-Add-In aktivieren:

- Öffnen Sie Excel 2007.
- Gehen Sie zur Office-Schaltfläche (oben links) und wählen Sie Excel-Optionen.
- Gehen Sie zum Abschnitt Add-Ins.
- Wählen Sie im Dropdown-Menü "Verwalten" (unten im Fenster) Excel-Add-Ins und klicken Sie auf "Gehe zu".
- Klicken Sie auf Durchsuchen und suchen Sie die Datei `xlwings.xlam`:
  Die Datei `xlwings.xlam` befindet sich normalerweise im Python-Verzeichnis, im Ordner `Lib\site-packages\xlwings\addin\`.
  Beispielpfad: `C:\Python313\Lib\site-packages\xlwings\addin\xlwings.xlam`.
- Wählen Sie die Datei aus und klicken Sie auf OK.
- Stellen Sie sicher, dass das xlwings-Add-In im Listenfeld aktiviert ist, und klicken Sie auf OK.
- Öffnen Sie den VBA-Editor (Alt + F11).
- Wählen Sie im VBA-Menü Tools → References (Dienstprogramme → Verweise).
- Stellen Sie sicher, dass xlwings in der Liste der verbundenen Verweise aufgeführt ist. Aktivieren Sie es, falls es nicht aktiviert ist.

4. Konfigurieren Sie die xlwings.conf
Erstellen Sie eine Konfigurationsdatei `xlwings.conf` im Ordner mit der Excel-Datei. Fügen Sie den folgenden Text hinzu:

```
"INTERPRETER_WIN","C:\Python313\python.exe"
"PYTHONPATH","C:\Users\Admin\Desktop\01_Excel\Python"
```
wobei `PYTHONPATH` der vollständige Pfad zum Ordner ist, in dem sich die Datei `macro.py` befindet.

5. Erstellen Sie VBA-Code zum Aufrufen des Python-Skripts
Öffnen Sie den VBA-Editor (Alt + F11) und fügen Sie ein neues Modul hinzu. Fügen Sie den folgenden Code ein:

```vba
Sub RunPythonMacro()
    RunPython "import macro ; macro.my_macro()"
End Sub
```
or 
```vba
Sub RunPythonScript()
    On Error GoTo ErrorHandler
    
    ' Ensure xlwings add-in is installed and active
    Application.Run "xlwings.xlam!RunPython", "import my_macro; my_macro.my_macro()"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
```

6. Fügen Sie eine Schaltfläche in Excel hinzu
- Gehen Sie zur Registerkarte Entwickler (Developer).
- Klicken Sie auf Einfügen -> Formular -> Schaltfläche (Button).
- Zeichnen Sie die Schaltfläche auf dem Blatt.
- Wählen Sie im Fenster "Makro zuweisen" `RunPythonMacro` und klicken Sie auf "OK".

7. Starten Sie das Makro
Klicken Sie nun auf die Schaltfläche in Excel, und Ihr Python-Skript wird ausgeführt.
