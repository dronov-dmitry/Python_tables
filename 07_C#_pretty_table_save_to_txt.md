```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;

class Program
{
    static void Main()
    {
        List<string[]> data = new List<string[]>
        {
            new string[] { "ID", "Name", "Occupation" },
            new string[] { "1", "Alice", "Software Engineer" },
            new string[] { "2", "Bob", "Doctor" },
            new string[] { "3", "Charlie", "Architect" }
        };

        // Get the executable's current directory
        string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string filePath = Path.Combine(exeDirectory, "table.txt");

        // Calculate max column widths
        int[] columnWidths = new int[data[0].Length];
        for (int col = 0; col < columnWidths.Length; col++)
        {
            columnWidths[col] = data.Max(row => row[col].Length);
        }

        // Format table as string
        string table = FormatTable(data, columnWidths);

        // Save to file
        File.WriteAllText(filePath, table);

        Console.WriteLine($"Table saved to: {filePath}");
    }

    static string FormatTable(List<string[]> data, int[] columnWidths)
    {
        // Create horizontal line with '|' at start and end
        string borderLine = "|" + string.Join("+", columnWidths.Select(w => new string('-', w + 2))) + "|";
        string formattedTable = borderLine + "\n"; // Top border

        // Print header
        formattedTable += "| " + string.Join(" | ", data[0].Select((cell, i) => cell.PadRight(columnWidths[i]))) + " |\n";
        formattedTable += borderLine + "\n"; // Header separator

        // Print rows without additional separators
        for (int i = 1; i < data.Count; i++)
        {
            formattedTable += "| " + string.Join(" | ", data[i].Select((cell, j) => cell.PadRight(columnWidths[j]))) + " |\n";
        }

        formattedTable += borderLine + "\n"; // Bottom border

        return formattedTable;
    }
}
```

---

### **📝 Output in `table.txt`**
```
|----+----------+------------------|
| ID | Name     | Occupation        |
|----+----------+------------------|
| 1  | Alice    | Software Engineer |
| 2  | Bob      | Doctor            |
| 3  | Charlie  | Architect         |
|----+----------+------------------|
```
