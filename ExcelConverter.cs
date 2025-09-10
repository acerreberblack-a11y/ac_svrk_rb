using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NLog;

namespace SpravkoBot_AsSapfir
{
public class ExcelConverter
{
    private static readonly Logger Log = LogManager.GetCurrentClassLogger();
    private const string DefaultPassword = "1234";
    private readonly string _excelPath;

    public ExcelConverter(string excelPath)
    {
        _excelPath = excelPath ?? throw new ArgumentNullException(nameof(excelPath));
        // Устанавливаем лицензию EPPlus (бесплатно для некоммерческого использования)
        ExcelPackage.License.SetNonCommercialOrganization("test");
    }

    public string ConvertToCsv()
    {
        Log.Info($"Начало конвертации файла Excel '{_excelPath}' в CSV.");

        if (string.IsNullOrEmpty(_excelPath))
        {
            throw new ArgumentException("Путь к файлу не может быть пустым", nameof(_excelPath));
        }

        if (!File.Exists(_excelPath))
        {
            throw new FileNotFoundException("Excel-файл не найден!", _excelPath);
        }

        string csvPath = Path.ChangeExtension(_excelPath, ".csv");

        try
        {
            using (var package = OpenExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets[0]; // Первый лист (индекс 0 в EPPlus)

                if (worksheet == null)
                {
                    throw new Exception("В файле нет доступных листов");
                }

                PrepareWorksheet(worksheet);
                SaveAsCsv(worksheet, csvPath);
            }

            Log.Info($"Файл успешно конвертирован в CSV: {csvPath}");
            // TryDeleteExcelFile();
            return csvPath;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Ошибка при конвертации Excel в CSV");
            throw new Exception($"Ошибка при конвертации Excel в CSV: {ex.Message}", ex);
        }
    }

    private ExcelPackage OpenExcelPackage()
    {
        try
        {
            var fileInfo = new FileInfo(_excelPath);
            return new ExcelPackage(fileInfo, DefaultPassword);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Ошибка при открытии Excel-файла");
            throw;
        }
    }

    private void PrepareWorksheet(ExcelWorksheet worksheet)
    {
        try
        {
            // EPPlus автоматически обрабатывает защищенные листы при наличии пароля
            // Удаляем автофильтры, если они есть
            if (worksheet.AutoFilter != null)
            {
                worksheet.AutoFilter.ClearAll();
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Ошибка при подготовке листа Excel");
            throw new Exception($"Ошибка при подготовке листа: {ex.Message}", ex);
        }
    }

    private void SaveAsCsv(ExcelWorksheet worksheet, string csvPath)
    {
        try
        {
            // Создаем StringBuilder для построения CSV
            var csvBuilder = new StringBuilder();
            int rowCount = worksheet.Dimension?.Rows ?? 0;
            int colCount = worksheet.Dimension?.Columns ?? 0;

            if (rowCount == 0 || colCount == 0)
            {
                throw new Exception("Лист пустой или не удалось определить размеры");
            }

            // Проходим по всем строкам и столбцам
            for (int row = 1; row <= rowCount; row++)
            {
                var rowData = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    // Экранируем кавычки и обрабатываем запятые
                    string escapedValue = $"\"{cellValue.Replace("\"", "\"\"")}\"";
                    rowData.Add(escapedValue);
                }
                csvBuilder.AppendLine(string.Join(",", rowData));
            }

            // Проверяем директорию и записываем файл
            string directory = Path.GetDirectoryName(csvPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (File.Exists(csvPath))
            {
                File.Delete(csvPath);
            }

            File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);
            Log.Info($"Сохранено как CSV: {csvPath}. Строк: {rowCount}, столбцов: {colCount}");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Ошибка при сохранении CSV-файла");
            throw new Exception($"Ошибка при сохранении CSV: {ex.Message}", ex);
        }
    }

    private void TryDeleteExcelFile()
    {
        try
        {
            if (File.Exists(_excelPath))
            {
                File.Delete(_excelPath);
                Log.Info($"Исходный файл {_excelPath} удален.");
            }
        }
        catch (Exception ex)
        {
            Log.Warn(ex, "Не удалось удалить Excel-файл");
        }
    }
}
}