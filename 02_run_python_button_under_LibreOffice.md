# UKR

Щоб створити кнопку в **LibreOffice**, яка буде запускати Python-скрипт, необхідно виконати кілька кроків. Я детально розповім, як це зробити.

### 1. **Налаштування Python-скрипта для LibreOffice**
Перед тим як створити кнопку в LibreOffice, переконайтеся, що Python та `pyuno` налаштовані, і ви можете запускати Python-скрипти з використанням UNO API.

#### Приклад Python-скрипта
Для початку створимо простий скрипт, який буде взаємодіяти з LibreOffice через UNO API. Наприклад, цей скрипт вставляє текст в документ Writer.

```
import uno

def fill_first_cell(*args):
    # Отримуємо компонентний контекст
    local_context = uno.getComponentContext()

    # Отримуємо службу для доступу до LibreOffice
    desktop = local_context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", local_context)

    # Отримуємо поточний документ (робочу книгу)
    document = desktop.getCurrentComponent()

    # Отримуємо доступ до першого листа (першої сторінки)
    sheet = document.Sheets[0]

    # Заповнюємо першу клітинку (A1)
    cell = sheet.getCellByPosition(0, 0)  # (0, 0) це клітинка A1
    cell.String = "Привіт, це Python!"
```

Збережіть цей скрипт у папку `Scripts/python/` у вашому профілі LibreOffice:
- На Windows це може бути: `C:\Users\<Ім'я_користувача>\AppData\Roaming\LibreOffice\4\user\Scripts\python\`.

### 2. **Створення кнопки в LibreOffice**

1. **Відкрийте LibreOffice Writer (або Calc).**
2. **Додайте панель інструментів з кнопками:**
   - Перейдіть у меню **Вигляд → Панелі інструментів → Елементи керування формами**.
3. **Створіть кнопку:**
   - Виберіть кнопку з панелі інструментів **Елементи керування формами** і намалюйте її в документі.
4. **Відкрийте властивості кнопки:**
   - Клацніть правою кнопкою миші на кнопці та виберіть **Властивості керування**.
5. **Прив'яжіть макрос:**
   - Перейдіть на вкладку **Події**.
   - У події **Натискання миші** натисніть `...` для вибору макроса.
   - У діалоговому вікні виберіть **Мої макроси** → **Python** → виберіть потрібний скрипт і функцію (наприклад, `insert_text`).
   - Натисніть **ОК**.

### 3. **Перевірка роботи кнопки**
1. Після того як ви прив'язали макрос, вийдіть з **Режиму дизайну**, щоб кнопка стала активною.
2. Натисніть на кнопку — відкриється новий документ LibreOffice Writer, і в ньому буде вставлений текст **"Привіт, це Python через UNO API!"**.

### 4. **Налаштування змінних середовища (якщо необхідно)**
Якщо LibreOffice не бачить ваш Python-скрипт або не може знайти потрібні бібліотеки, переконайтеся, що в змінній `PYTHONPATH` вказаний шлях до папки, де встановлені бібліотеки LibreOffice.

На Windows це можна зробити таким чином:
1. Перейдіть у **Властивості системи** → **Додаткові параметри системи** → **Змінні середовища**.
2. У розділі **Змінні користувача** додайте змінну `PYTHONPATH`, яка буде вказувати на папки з бібліотеками LibreOffice, наприклад:
   - Шлях: `C:\Program Files\LibreOffice\program\` (шлях до папки з `uno.py`).

---
# RUS

Чтобы создать кнопку в **LibreOffice**, которая будет запускать Python-скрипт, необходимо выполнить несколько шагов. Я подробно расскажу, как это сделать.

### 1. **Настройка Python-скрипта для LibreOffice**
Прежде чем создать кнопку в LibreOffice, убедитесь, что Python и `pyuno` настроены, и вы можете запускать Python-скрипты с использованием UNO API.

#### Пример Python-скрипта
Для начала создадим простой скрипт, который будет взаимодействовать с LibreOffice через UNO API. Например, этот скрипт вставляет текст в документ Writer.

```
import uno

def fill_first_cell(*args):
    # Получаем компонентный контекст
    local_context = uno.getComponentContext()

    # Получаем службу для доступа к LibreOffice
    desktop = local_context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", local_context)

    # Получаем текущий документ (рабочую книгу)
    document = desktop.getCurrentComponent()

    # Получаем доступ к первому листу (первой странице)
    sheet = document.Sheets[0]

    # Заполняем первую ячейку (A1)
    cell = sheet.getCellByPosition(0, 0)  # (0, 0) это ячейка A1
    cell.String = "Привет, это Python!"
```

Сохраните этот скрипт в папку `Scripts/python/` в вашем профиле LibreOffice:
- На Windows это может быть: `C:\Users\<Имя_пользователя>\AppData\Roaming\LibreOffice\4\user\Scripts\python\`.

### 2. **Создание кнопки в LibreOffice**

1. **Откройте LibreOffice Writer (или Calc).**
2. **Добавьте панель инструментов с кнопками:**
   - Перейдите в меню **Вид → Панели инструментов → Элементы управления формами**.
3. **Создайте кнопку:**
   - Выберите кнопку из панели инструментов **Элементы управления формами** и нарисуйте её в документе.
4. **Откройте свойства кнопки:**
   - Щелкните правой кнопкой мыши на кнопке и выберите **Свойства управления**.
5. **Привяжите макрос:**
   - Перейдите на вкладку **События**.
   - В событии **Нажатие мыши** нажмите `...` для выбора макроса.
   - В диалоговом окне выберите **Мои макросы** → **Python** → выберите нужный скрипт и функцию (например, `insert_text`).
   - Нажмите **ОК**.

### 3. **Проверка работы кнопки**
1. После того как вы привязали макрос, выйдите из **Режима дизайна**, чтобы кнопка стала активной.
2. Нажмите на кнопку — откроется новый документ LibreOffice Writer, и в нем будет вставлен текст **"Привет, это Python через UNO API!"**.

### 4. **Настройка переменных окружения (если необходимо)**
Если LibreOffice не видит ваш Python-скрипт или не может найти нужные библиотеки, убедитесь, что в переменной `PYTHONPATH` указан путь к папке, где установлены библиотеки LibreOffice.

На Windows это можно сделать следующим образом:
1. Перейдите в **Свойства системы** → **Дополнительные параметры системы** → **Переменные среды**.
2. В разделе **Переменные пользователя** добавьте переменную `PYTHONPATH`, которая будет указывать на папки с библиотеками LibreOffice, например:
   - Путь: `C:\Program Files\LibreOffice\program\` (путь к папке с `uno.py`).

---
# ENG

Here is the translated text in English:

To create a button in **LibreOffice** that will run a Python script, you need to follow several steps. I will explain how to do this in detail.

### 1. **Setting up the Python script for LibreOffice**
Before creating the button in LibreOffice, make sure that Python and `pyuno` are set up and that you can run Python scripts using the UNO API.

#### Example Python script
First, let's create a simple script that will interact with LibreOffice via the UNO API. For example, this script inserts text into a Writer document.

```
import uno

def fill_first_cell(*args):
    # Get the component context
    local_context = uno.getComponentContext()

    # Get the service to access LibreOffice
    desktop = local_context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", local_context)

    # Get the current document (spreadsheet)
    document = desktop.getCurrentComponent()

    # Get access to the first sheet (the first page)
    sheet = document.Sheets[0]

    # Fill the first cell (A1)
    cell = sheet.getCellByPosition(0, 0)  # (0, 0) is cell A1
    cell.String = "Hello, this is Python!"
```

Save this script in the `Scripts/python/` folder in your LibreOffice profile:
- On Windows, it could be: `C:\Users\<Username>\AppData\Roaming\LibreOffice\4\user\Scripts\python\`.

### 2. **Creating a button in LibreOffice**

1. **Open LibreOffice Writer (or Calc).**
2. **Add a toolbar with buttons:**
   - Go to the menu **View → Toolbars → Form Controls**.
3. **Create a button:**
   - Select the button from the **Form Controls** toolbar and draw it on the document.
4. **Open the button's properties:**
   - Right-click on the button and select **Control Properties**.
5. **Attach the macro:**
   - Go to the **Events** tab.
   - In the **Mouse button press** event, click `...` to select the macro.
   - In the dialog, select **My Macros** → **Python** → choose the desired script and function (e.g., `insert_text`).
   - Click **OK**.

### 3. **Testing the button**
1. After you have attached the macro, exit **Design Mode** to make the button active.
2. Click the button — a new LibreOffice Writer document will open, and the text **"Hello, this is Python via UNO API!"** will be inserted.

### 4. **Setting up environment variables (if necessary)**
If LibreOffice cannot find your Python script or the required libraries, make sure the `PYTHONPATH` environment variable points to the folder where LibreOffice libraries are installed.

On Windows, you can do this as follows:
1. Go to **System Properties** → **Advanced system settings** → **Environment Variables**.
2. In the **User variables** section, add a `PYTHONPATH` variable that points to the LibreOffice libraries folder, for example:
   - Path: `C:\Program Files\LibreOffice\program\` (the path to the folder containing `uno.py`).

---
# DEU

Hier ist der Text auf Deutsch übersetzt:

Um eine Schaltfläche in **LibreOffice** zu erstellen, die ein Python-Skript ausführt, müssen Sie mehrere Schritte ausführen. Ich erkläre Ihnen, wie das geht.

### 1. **Python-Skript für LibreOffice einrichten**
Bevor Sie die Schaltfläche in LibreOffice erstellen, stellen Sie sicher, dass Python und `pyuno` eingerichtet sind und Sie Python-Skripte mithilfe der UNO-API ausführen können.

#### Beispiel eines Python-Skripts
Erstellen wir zunächst ein einfaches Skript, das über die UNO-API mit LibreOffice interagiert. Dieses Skript fügt beispielsweise Text in ein Writer-Dokument ein.

```
import uno

def fill_first_cell(*args):
    # Komponenten-Kontext abrufen
    local_context = uno.getComponentContext()

    # Dienst für den Zugriff auf LibreOffice abrufen
    desktop = local_context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", local_context)

    # Aktuelles Dokument abrufen (Tabellenkalkulation)
    document = desktop.getCurrentComponent()

    # Zugriff auf das erste Tabellenblatt erhalten (erste Seite)
    sheet = document.Sheets[0]

    # Die erste Zelle (A1) ausfüllen
    cell = sheet.getCellByPosition(0, 0)  # (0, 0) ist die Zelle A1
    cell.String = "Hallo, das ist Python!"
```

Speichern Sie dieses Skript im Ordner `Scripts/python/` in Ihrem LibreOffice-Profil:
- Unter Windows könnte dies sein: `C:\Users\<Benutzername>\AppData\Roaming\LibreOffice\4\user\Scripts\python\`.

### 2. **Eine Schaltfläche in LibreOffice erstellen**

1. **Öffnen Sie LibreOffice Writer (oder Calc).**
2. **Fügen Sie eine Symbolleiste mit Schaltflächen hinzu:**
   - Gehen Sie zum Menü **Ansicht → Symbolleisten → Formular-Steuerelemente**.
3. **Erstellen Sie eine Schaltfläche:**
   - Wählen Sie die Schaltfläche aus der Symbolleiste **Formular-Steuerelemente** und zeichnen Sie sie im Dokument.
4. **Öffnen Sie die Eigenschaften der Schaltfläche:**
   - Klicken Sie mit der rechten Maustaste auf die Schaltfläche und wählen Sie **Steuerelementeigenschaften**.
5. **Makro zuweisen:**
   - Wechseln Sie zur Registerkarte **Ereignisse**.
   - Wählen Sie beim Ereignis **Mausklick** die Schaltfläche `...`, um das Makro auszuwählen.
   - Wählen Sie im Dialogfeld **Meine Makros** → **Python** → das gewünschte Skript und die Funktion (z. B. `insert_text`) aus.
   - Klicken Sie auf **OK**.

### 3. **Testen der Schaltfläche**
1. Nachdem Sie das Makro zugewiesen haben, verlassen Sie den **Entwurfsmodus**, um die Schaltfläche aktiv zu machen.
2. Klicken Sie auf die Schaltfläche – ein neues LibreOffice Writer-Dokument wird geöffnet, und der Text **"Hallo, das ist Python über die UNO-API!"** wird eingefügt.

### 4. **Umgebungsvariablen konfigurieren (falls erforderlich)**
Wenn LibreOffice Ihr Python-Skript oder die erforderlichen Bibliotheken nicht finden kann, stellen Sie sicher, dass die Umgebungsvariable `PYTHONPATH` auf den Ordner verweist, in dem die LibreOffice-Bibliotheken installiert sind.

Unter Windows können Sie dies wie folgt tun:
1. Gehen Sie zu **Systemeigenschaften** → **Erweiterte Systemeinstellungen** → **Umgebungsvariablen**.
2. Fügen Sie im Abschnitt **Benutzervariablen** eine Variable `PYTHONPATH` hinzu, die auf den Ordner mit den LibreOffice-Bibliotheken verweist, z. B.:
   - Pfad: `C:\Program Files\LibreOffice\program\` (der Pfad zum Ordner, der `uno.py` enthält).
