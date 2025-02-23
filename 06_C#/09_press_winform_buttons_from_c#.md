Если нужно нажимать кнопки в чужом приложении, можно использовать WinAPI. Вот несколько способов:

### 1. **Использование `PostMessage` или `SendMessage`**  
Эти функции позволяют отправлять сообщения окнам в другом процессе.  

#### Найти окно и кнопку по `hWnd` и отправить `WM_COMMAND`:
```csharp
using System;
using System.Runtime.InteropServices;

class Program
{
    const int WM_COMMAND = 0x0111;

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll")]
    static extern int SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    static void Main()
    {
        // Найти главное окно по заголовку
        IntPtr hWndMain = FindWindow(null, "Название окна программы");

        if (hWndMain == IntPtr.Zero)
        {
            Console.WriteLine("Окно не найдено.");
            return;
        }

        // Найти кнопку внутри окна (обычно класс кнопки — "Button")
        IntPtr hWndButton = FindWindowEx(hWndMain, IntPtr.Zero, "Button", "Текст кнопки");

        if (hWndButton == IntPtr.Zero)
        {
            Console.WriteLine("Кнопка не найдена.");
            return;
        }

        // Отправить команду нажатия кнопки
        SendMessage(hWndButton, WM_COMMAND, IntPtr.Zero, IntPtr.Zero);
        Console.WriteLine("Кнопка нажата.");
    }
}
```

### 2. **Использование `PostMessage` (если кнопка реагирует на `WM_LBUTTONDOWN`)**
```csharp
const int WM_LBUTTONDOWN = 0x0201;
const int WM_LBUTTONUP = 0x0202;

PostMessage(hWndButton, WM_LBUTTONDOWN, IntPtr.Zero, IntPtr.Zero);
PostMessage(hWndButton, WM_LBUTTONUP, IntPtr.Zero, IntPtr.Zero);
```
Этот метод полезен, если кнопка не обрабатывает `WM_COMMAND`.

### 3. **Эмуляция нажатия с `SetCursorPos` и `mouse_event`**
Если кнопка не реагирует на `SendMessage`, можно эмулировать физический клик:
```csharp
using System.Runtime.InteropServices;

[DllImport("user32.dll")]
static extern void SetCursorPos(int X, int Y);

[DllImport("user32.dll")]
static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

const int MOUSEEVENTF_LEFTDOWN = 0x02;
const int MOUSEEVENTF_LEFTUP = 0x04;

SetCursorPos(x, y); // Координаты кнопки
mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
```
Координаты кнопки можно определить с помощью `GetWindowRect()`.

### Какой способ тебе больше подходит?
