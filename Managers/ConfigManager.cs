using NLog;
using System;
using System.Collections.Generic;

namespace SpravkoBot_AsSapfir
{
internal class ConfigManager
{
    private static readonly Logger Log = LogManager.GetCurrentClassLogger();
    private readonly string _configPath;
    public Config Config { get; private set; }
    private readonly JsonManager _jsonManager;

    public ConfigManager(string configPath)
    {
        _configPath = configPath;
        _jsonManager = new JsonManager(_configPath);
        LoadConfig();
    }

    public void LoadConfig()
    {
        try
        {
            Config = _jsonManager.GetValue<Config>("$");
            Log.Info("Конфигурация успешно загружена.");
        }
        catch (Exception ex)
        {
            Log.Error($"Ошибка загрузки конфигурации: {ex.Message}");
            throw new Exception($"Ошибка загрузки конфигурации: {ex.Message}");
        }
    }

    public static string GetTemplateConfig()
    {
        return @"{
  // Путь к SAP Logon (исполняемый файл)
  ""SapLogonPath"": ""C:\\Program Files (x86)\\SAP\\FrontEnd\\SapGui\\saplogon.exe"",
  // Путь к Excel файлу, который будет конвертирован в CSV
  ""ExcelPath"": ""C:\\Users\\RobinSapAC\\Desktop\\02-04-2025\\SpravkoBot_AsSapfir\\bin\\Debug\\KA\\Реестр по всем БЕ.xlsx"",
  // Название системы SAP для подключения
  ""SapStage"": ""1. Продуктивная среда САПФИР"",
  // Название тестовой системы SAP
  ""SapTestStage"": ""ER2 - среда тестирования САПФИР"",
  // Логин для тестовой среды SAP
  ""SapUser"": ""DIADOC_INT"",
  // Пароль для тестовой среды SAP
  ""SapPassword"": ""1Sdfghjkl12345^&2"",
  // Сопоставление названий БЕ и их кодов
  ""BeCodes"": {
    ""Бурейская ГЭС"": ""1030"",
    ""Исполнительный аппарат"": ""1010"",
    ""Волжская ГЭС"": ""1050"",
    ""Воткинская ГЭС"": ""1060"",
    ""Дагестанский филиал"": ""1080"",
    ""Жигулевская ГЭС"": ""1090"",
    ""Загорская ГАЭС"": ""1100"",
    ""Зейская ГЭС"": ""1140"",
    ""Кабардино-Балкарский филиал"": ""1170"",
    ""Камская ГЭС"": ""1180"",
    ""Каскад Верхневолжских ГЭС"": ""1200"",
    ""Каскад Кубанских ГЭС"": ""1210"",
    ""Нижегородская ГЭС"": ""1240"",
    ""Новосибирская ГЭС"": ""1260"",
    ""Саратовская ГЭС"": ""1300"",
    ""Северо-Осетинский филиал"": ""1320"",
    ""Чебоксарская ГЭС"": ""1350"",
    ""КорУнГ"": ""1410"",
    ""Хабаровский филиал"": ""1510"",
    ""Приморский филиал"": ""1520"",
    ""Якутский филиал"": ""1530"",
    ""Карачаево-Черкесский филиал"": ""1190"",
    ""Саяно-Шушенская ГЭС (СШГЭС им. П.С. Непорожнего)"": ""1310""
  }
}";
    }
}

public class Config
{
    public string SapLogonPath { get; set; }
    public string ExcelPath { get; set; }
    public string SapStage { get; set; }
    public string SapTestStage { get; set; }
    public string SapUser { get; set; }
    public string SapPassword { get; set; }
    public Dictionary<string, string> BeCodes { get; set; } = new Dictionary<string, string>();
}
}
