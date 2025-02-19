To run a C# application from a VBA button in Excel, you have a few different approaches. The most common method is using VBA to call an external C# executable (`.exe`) or a C# DLL via COM Interop.

---

### **Method 1: Run a C# Executable from VBA**
This method runs a standalone C# application (`.exe`) from a button in Excel.

#### **Steps:**
1. **Install Dependencies:**
Make sure you have **Microsoft.Office.Interop.Excel** installed. If not, install it via NuGet:
```sh
Install-Package Microsoft.Office.Interop.Excel
```
2. **Create a C# Console Application**
   - Open **Visual Studio** and create a new **Console App (.NET Framework)**.
   - Write your C# logic inside `Main()`.

   ```csharp
   using System;
   using System.Runtime.InteropServices;
   using Excel = Microsoft.Office.Interop.Excel;
   class Program
   {
       static void Main()
       {
           Console.Write("Attach process for debug: Ctrl+Alt+P -> Track Window");
           string name = Console.ReadLine();
           Console.WriteLine($"Here is going program!");
           try
           {
               // Attach to the running Excel application
               Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
               if (excelApp == null || excelApp.Workbooks.Count == 0)
               {
                   Console.WriteLine("No open Excel workbooks found.");
                   return;
               }

               // Get the active workbook
               Excel.Workbook workbook = excelApp.ActiveWorkbook;

               Console.WriteLine("Workbook: " + workbook.Name);

               // Iterate through all sheets
               foreach (Excel.Worksheet sheet in workbook.Sheets)
               {
                   Console.WriteLine("Sheet: " + sheet.Name);

                   // Write a value to cell A1
                   sheet.Cells[1, 1] = "Hello, Excel!";
               }

               // Save workbook
               workbook.Save();
               Console.WriteLine("Data written successfully.");
           }
           catch (COMException)
           {
               Console.WriteLine("Excel is not running or no workbook is open.");
           }
           catch (Exception ex)
           {
               Console.WriteLine("Error: " + ex.Message);
           }
           Console.Write("End of programm");
           string endProgram = Console.ReadLine();
       }
   }
   ```

3. **Build the C# Project** to generate an `.exe` file (located in `bin\Debug` or `bin\Release`).

4. **Write VBA Code to Run the C# Executable**
   - Add a button to an Excel sheet.
   - Open **VBA Editor** (`Alt + F11`).
   - Create a new module and add this VBA code:

   ```vba
   Sub RunCSharpExe()
       Dim exePath As String
       exePath = "C:\path\to\your\CSharpApp.exe"
       Shell exePath, vbNormalFocus
   End Sub
   ```

5. **Assign the Macro** to the button.

6. **Run the Button**, and it should execute the C# program.

---

### **Method 2: Call a C# DLL via COM Interop**
If you need to return values to VBA, use a **C# Class Library (DLL)** registered as a COM object.

#### **Steps:**
1. **Create a C# Class Library**
   - In Visual Studio, create a **Class Library (.NET Framework)**.
   - Add `[ComVisible(true)]` and register it for COM.

   ```csharp
   using System;
   using System.Runtime.InteropServices;

   [ComVisible(true)]
   [Guid("D69BA7A4-1C6B-4E48-B0A3-7F6F5B038111")]
   [ClassInterface(ClassInterfaceType.AutoDual)]
   public class CSharpInterop
   {
       public string GetMessage()
       {
           return "Hello from C# DLL!";
       }
   }
   ```

2. **Register the DLL for COM Interop**
   - **Build the project**.
   - Open **Command Prompt as Administrator** and register the DLL:

     ```cmd
     regasm "C:\path\to\YourLibrary.dll" /codebase
     ```

3. **Use the DLL in VBA**
   - In VBA, go to **Tools > References** and find your DLL.
   - Use this VBA code:

   ```vba
   Sub CallCSharpDLL()
       Dim obj As Object
       Set obj = CreateObject("YourNamespace.CSharpInterop")
       MsgBox obj.GetMessage()
   End Sub
   ```

4. **Assign the macro** to a button.

5. **Click the button**, and it will call the C# function!

---

### **Which Method Should You Use?**
- **If your C# code is standalone**, Method 1 (EXE) is easiest.
- **If you need data exchange**, Method 2 (DLL via COM) is better.

Let me know if you need help with any step! 🚀
