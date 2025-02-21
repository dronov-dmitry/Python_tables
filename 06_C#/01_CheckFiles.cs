using System;
using System.IO;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        // Получаем путь к папке, где находится исполняемый файл
        string folderPath = AppDomain.CurrentDomain.BaseDirectory;

        // Запрашиваем количество дней
        Console.Write("Введите количество дней больше которых нужно указать файлы (по умолчанию 5): ");
        string input1 = Console.ReadLine();
        Console.Write("Введите количество дней меньше которых нужно указать файлы (по умолчанию 20): ");
        string input2 = Console.ReadLine();
        // Запрашиваем расширения файлов через запятую
        Console.Write("Введите расширения файлов через запятую, которые будут проверяться (по умолчанию nwc, nwd): ");
        string inputExtensions = Console.ReadLine();

        // Если расширения не введены, используем по умолчанию "nwc" и "nwd"
        string[] extensions = string.IsNullOrEmpty(inputExtensions)
            ? new[] { "nwc", "nwd" }
            : inputExtensions.Split(',').Select(ext => ext.Trim()).ToArray();


        // Если пользователь не ввел значение, используем по умолчанию 5 и 20
        int days = 5;
        int days_max = 20;
        if (!string.IsNullOrEmpty(input1) && int.TryParse(input1, out int inputDays))
        {
            days = inputDays;
        }
        if (!string.IsNullOrEmpty(input2) && int.TryParse(input2, out int inputDays2))
        {
            days_max = inputDays2;
        }


        // Получаем все файлы в папке
        string[] files = Directory.GetFiles(folderPath);

        // Получаем текущую дату
        DateTime currentDate = DateTime.Now;

        // Проверяем каждый файл
        foreach (string file in files)
        {
            // Получаем дату последней модификации файла
            DateTime lastModified = File.GetLastWriteTime(file);

            // Если файл был изменен позже, чем указанное количество дней
            double daysAgo = (currentDate - lastModified).TotalDays;
            string fileExtension = Path.GetExtension(file).TrimStart('.').ToLower(); // Получаем расширение файла

            // Проверяем, соответствует ли расширение файлу
            if (extensions.Contains(fileExtension) && daysAgo > days && daysAgo < days_max)
            {
                // Используем Path.GetFileName для получения только имени файла
                string fileName = Path.GetFileName(file);
                Console.WriteLine($"Файл: {fileName} был изменен {daysAgo:F0} дня(ей) назад (Дата изменения: {lastModified})");
            }
        }

        Console.Write("Проверка выполнена (нажмите ENTER) ");
        Console.ReadLine();
    }
}
