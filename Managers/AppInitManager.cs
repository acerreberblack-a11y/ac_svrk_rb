using NLog;
using System;
using System.Collections.Generic;
using System.IO;

namespace SpravkoBot_AsSapfir
{
internal class AppInitManager
{
    private readonly string _baseAppPath;
    private readonly string _configFilePath;
    private readonly string _requestsPath;
    private readonly string[] _subdirectories = { "input", "error", "output", "logs", "temp" };
    private readonly LoggerManager _loggerManager;
    private static readonly Logger Log = LogManager.GetCurrentClassLogger();

    public AppInitManager()
    {
        _baseAppPath = AppDomain.CurrentDomain.BaseDirectory;
        _configFilePath = Path.Combine(_baseAppPath, "config.json");
        _requestsPath = Path.Combine(_baseAppPath, "data", "requests");
        _loggerManager = new LoggerManager(Path.Combine(_baseAppPath, _requestsPath, "logs"), // Путь до папки с логами
                                           30 // Количество дней хранения логов
        );
    }

    public void Init()
    {
        _loggerManager.ConfigureLogger();
        _loggerManager.ClearOldLogs();

        CreateDirectories();
        CreateConfigFile(_configFilePath);

        string tempFolder = Path.Combine(_requestsPath, "temp");
        ClearDirectory(tempFolder);
    }

    private void CreateDirectories()
    {
        try
        {
            if (!Directory.Exists(_requestsPath))
            {
                Directory.CreateDirectory(_requestsPath);
            }

            foreach (var subdirectory in _subdirectories)
            {
                string fullPath = Path.Combine(_requestsPath, subdirectory);

                if (!Directory.Exists(fullPath))
                {
                    Directory.CreateDirectory(fullPath);
                }
            }
        }
        catch (UnauthorizedAccessException ex)
        {
            Log.Error($"Ошибка доступа: {ex.Message}");
        }
        catch (IOException ex)
        {
            Log.Error($"Ошибка ввода-вывода: {ex.Message}");
        }
        catch (Exception ex)
        {
            Log.Error($"Неизвестная ошибка: {ex.Message}");
        }
    }

    private void CreateConfigFile(string path)
    {
        try
        {
            // Проверка существования файла и его содержимого
            if (!File.Exists(path) || new FileInfo(path).Length == 0)
            {
                File.WriteAllText(path, ConfigManager.GetTemplateConfig());
                Log.Info("Файл конфигурации config.json создан и заполнен данными.");
            }
        }
        catch (UnauthorizedAccessException ex)
        {
            Log.Error($"Ошибка доступа при создании config.json: {ex.Message}");
        }
        catch (IOException ex)
        {
            Log.Error($"Ошибка ввода-вывода при создании config.json: {ex.Message}");
        }
        catch (Exception ex)
        {
            Log.Error($"Неизвестная ошибка при создании config.json: {ex.Message}");
        }
    }

    public Dictionary<string, string> GetAppFolders()
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var subdir in _subdirectories)
        {
            dict[subdir] = Path.Combine(_requestsPath, subdir);
        }
        return dict;
    }

    public string GetPathConfigFile() => _configFilePath;

    public void ClearDirectory(string path)
    {
        if (!Directory.Exists(path))
        {
            Log.Error($"Не найдена папка {path}.");
            return;
        }

        try
        {
            foreach (var file in Directory.GetFiles(path))
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception e)
                {
                    Log.Error(e, $"Ошибка при удалении файла '{file}'");
                }
            }
        }
        catch (Exception e)
        {
            Log.Error(e, $"Ошибка при удалении файлов в папке {path}");
        }
    }
}
}
