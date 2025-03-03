If you want to press winform button from csharp command




Spy++ действительно поставляется только с **Visual Studio Community, Professional или Enterprise** (не с Build Tools). Если у вас его нет, вот альтернативы:  

### **1. Установить Spy++ (если Visual Studio уже есть)**
#### **Шаги:**
1. Откройте `Visual Studio Installer`.  
2. Выберите установленную версию Visual Studio.  
3. Нажмите **"Изменить"**.  
4. В разделе "Отладка и тестирование" выберите Средства профилирования C++.
5. В разделе "Действия разработки" выберите Основные компоненты C++.

После этого Spy++ появится здесь:  
📂 `C:\Program Files (x86)\Microsoft Visual Studio\...\Common7\Tools\spyxx_amd64.exe`  


### **Как найти `WM_COMMAND ID` кнопки через Spy++ или Inspect.exe?**  

Чтобы программно нажать кнопку или пункт меню в Windows-приложении, нужно узнать его **ID команды (`WM_COMMAND ID`)**. Это можно сделать с помощью утилит **Spy++** (идет с Visual Studio) или **Inspect.exe** (из Windows SDK).  

---

## **Метод 1: Spy++ (для Win32-приложений)**
### **Шаги:**
1. **Запустите Spy++**  
   - Откройте Visual Studio, нажмите `Ctrl+Q`, введите `Spy++` и запустите.  
   - Или найдите его в `C:\Program Files (x86)\Microsoft Visual Studio\...\Common7\Tools\spyxx.exe`.  

2. **Выберите окно приложения**  
   - Нажмите `Find Window` (`Ctrl+F`).  
   - Перетащите **"прицел"** (`Finder Tool`) на нужное окно (например, Блокнот).  
   - Нажмите **ОК**.

3. **Отслеживание сообщений (WM_COMMAND)**  
   - Перейдите в `Spy++ → Log Messages` (`Ctrl+M`).  
   - Выберите свое окно и нажмите **ОК**.  
   - Включите фильтр:  
     - **Сообщения:** `WM_COMMAND`.  
     - **Классы:** можно оставить `All`.  
     - Нажмите **ОК**.

4. **Нажмите кнопку/меню в приложении**  
   - Например, в Блокноте откройте `Файл → Открыть...`.  
   - В **Spy++** появится строка `EN_KILLFOCUS wID:15`.  
   - Число `15` – это **ID команды** "Файл"!  
   - Число `2` – это **ID команды** "Открыть"!  

5. Впишите эти ID в программу на C#



```csharp
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

class Program
{
    // Импортируем функции из user32.dll
    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    // Константы Windows API
    const uint WM_COMMAND = 0x0111;

    static void Main()
    {
        // 1. Запускаем Блокнот
        Process process = Process.Start("notepad.exe");
        Thread.Sleep(2000); // Ждём, чтобы окно успело открыться

        // 2. Находим окно блокнота
        IntPtr hWnd = FindWindow("Notepad", null);
        if (hWnd == IntPtr.Zero)
        {
            Console.WriteLine("Окно не найдено!");
            return;
        }

        // 3. Отправляем команду "Файл"
        PostMessage(hWnd, WM_COMMAND, (IntPtr)15, IntPtr.Zero);
        Thread.Sleep(500);

        // 4. Отправляем команду "Открыть..."
        PostMessage(hWnd, WM_COMMAND, (IntPtr)2, IntPtr.Zero);

        Console.WriteLine("Открыто меню 'Файл' -> 'Открыть...'");
    }
}
```

```csharp
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

class Program
{
    [DllImport("user32.dll")]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    static void Main()
    {
        string filePath = @"C:\Users\DmitryD\Documents\RegisteredUpdaters.txt";

        // Запускаем Блокнот
        Process process = Process.Start("notepad.exe");
        Thread.Sleep(2000); // Ждём открытия окна

        // Находим окно Блокнота
        IntPtr hWnd = FindWindow("Notepad", null);
        if (hWnd == IntPtr.Zero)
        {
            Console.WriteLine("Окно не найдено!");
            return;
        }

        // Делаем окно активным
        SetForegroundWindow(hWnd);
        Thread.Sleep(500);

        // Открываем меню "Файл" → "Открыть..." (Alt + F, O)
        SendKeys.SendWait("%(FO)");
        Thread.Sleep(500);

        // Вводим путь к файлу и нажимаем Enter
        SendKeys.SendWait(filePath);
        SendKeys.SendWait("{ENTER}");

        Console.WriteLine("Файл открыт в Блокноте.");
    }
}
```
