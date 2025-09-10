using FlaUI.UIA3;
using FlaUI.Core.AutomationElements;
using NLog;
using SAPFEWSELib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text; // для StringBuilder в CSV парсере
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpravkoBot_AsSapfir
{
    internal class Program
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        private static JsonManager _jsonManager;
        private static ConfigManager _configManager;
        private static Dictionary<string, string> _appFolders;
        private static Dictionary<string, string> _beCodes;
        public static GuiSession session;
        private static string _signatory1 = string.Empty;
        private static string _signatory2 = string.Empty;

        // ===== МОДЕЛИ CSV =====
        private sealed class CsvRow
        {
            public string Number { get; set; } = "";
            public string Branch { get; set; } = "";
            public string CompanyName { get; set; } = "";
            public string CompanyNumberSap { get; set; } = "";
            public string INN { get; set; } = "";
            public string KPP { get; set; } = "";
            public string Status { get; set; } = "";
            public string SignatoryLanDocs { get; set; } = "";
            public string PersonnelNumber { get; set; } = "";
            public string SignatoryLandocs { get; set; } = "";
            public string VGO { get; set; } = "";
        }

        private sealed class InventoryResult
        {
            public string SignatoryNumber { get; set; } = "";
            public string Status { get; set; } = "";
            public List<string> CounterpartyNumbers { get; set; } = new List<string>();
        }

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                AppInitManager initializer = new AppInitManager();
                initializer.Init();
                Log.Info("*////*");

                _appFolders = initializer.GetAppFolders(); // убедитесь, что метод реализован в AppInitManager
                Log.Info("Инициализация успешно завершена.");

                string configPath = initializer.GetPathConfigFile();
                _configManager = new ConfigManager(configPath);
                var config = _configManager.Config;
                string excelPath = config.ExcelPath;
                string sapLogonPath = config.SapLogonPath;
                string stage = config.SapStage;
                string testStage = config.SapTestStage;
                string sapUser = config.SapUser;
                string sapPassword = config.SapPassword;
                _beCodes = config.BeCodes;

                if (!_appFolders.TryGetValue("input", out string inputFolder) || string.IsNullOrWhiteSpace(inputFolder))
                {
                    Log.Error("Инициализация входной папки input произошла с ошибкой.");
                    return;
                }

                var files = Directory.EnumerateFiles(inputFolder, "SD*.txt")
                                     .Where(file => Path.GetFileName(file).Contains("+"))
                                     .ToList();

                if (!files.Any())
                {
                    Log.Info("Нет заявок для обработки. Робот завершает работу.");
                    return;
                }

                Log.Info($"В папке {inputFolder} найдено {files.Count} входящих файл(-а)(-ов) для обработки.");

                foreach (var file in files)
                {
                    var currentFile = MoveToOwnFolderInInput(file);
                    string fileName = Path.GetFileName(currentFile);
                    Log.Info($"Обработка заявки: {fileName}. Начинаю извлечение данных из заявки.");

                    SapfirManager sapfir = null;
                    try
                    {
                        _jsonManager = new JsonManager(currentFile, true);
                        Request request = RequestManager.FromJson(_jsonManager);
                        Log.Info($"Для заявки {request.UIID} создана сущность Request.");

                        Log.Info("Данные из заявки извлечены:");
                        Log.Info("********************************************");
                        Log.Info($"* Номер заявки => {request.UIID}");
                        Log.Info($"* Тип заявки => {request.Type}");
                        Log.Info($"* Услуга => {request.Service}");
                        Log.Info($"* Организация => {request.Organization}");
                        Log.Info($"*..........................................*");
                        Log.Info($"* БЕ => {request.Branch}");
                        Log.Info($"* ИНН => {request.INN}");
                        Log.Info($"* КПП => {request.KPP}");
                        Log.Info($"* Номер (-а) договора в системе => {request.RegNumbDoc}");
                        Log.Info($"* Начало периода => {request.DateStart:dd.MM.yyyy}");
                        Log.Info($"* Конец периода => {request.DateEnd:dd.MM.yyyy}");
                        Log.Info("**********************************************");

                        Log.Info($"Начинаю конвертирование excel файла в csv: {excelPath}");
                        var converter = new ExcelConverter(excelPath);
                        string csvPath = converter.ConvertToCsv();

                        if (!File.Exists(csvPath))
                        {
                            throw new FileNotFoundException("CSV-файл не найден!", csvPath);
                        }

                        // ===========================
                        //  Три сценария по request.Type
                        // ===========================
                        Log.Info(
                            $"Выполняю фильтрацию по параметрам из заявки. Type: {request.Type}; БЕ:{request.Branch}, ИНН:{request.INN}, КПП:{request.KPP}");

                        string searchType = request.Type?.ToString()?.Trim();

                        CsvRow singleRow = null;                // для сценариев 1 и 2
                        List<InventoryResult> inventory = null; // для сценария 3
                        var tasks = new List<SapTask>();        // формируемые задачи

                        switch (searchType)
                        {
                            case "По одному контрагенту по всем договорам":
                                {
                                    singleRow = SearchSingleByBranchInnKpp(
                                        csvPath,
                                        request.Branch.ToString(),
                                        request.INN?.ToString(),
                                        request.KPP?.ToString()
                                    );

                                    if (singleRow == null)
                                        throw new Exception("Не найдено ни одной строки по БЕ/ИНН/КПП.");

                                    var be = GetValueByName(request.Branch) ??
                                             throw new Exception("Не смог сопоставить БЕ из заявки с конфигом (BeCodes).");

                                    var signatoryArray = SplitSignatory(singleRow.PersonnelNumber);
                                    if (signatoryArray.Count == 0)
                                        throw new Exception("Ошибка: PersonnelNumber пустой для найденной строки.");

                                    var t = new SapTask
                                    {
                                        BeCode = be,
                                        INN = request.INN?.ToString(),
                                        KPP = request.KPP?.ToString(),
                                        ContractNumber = new List<string>(),
                                        CounterpartyNumbers = new List<string> { singleRow.CompanyNumberSap },
                                        SignatoryNumbers = signatoryArray,
                                        Status = singleRow.Status,
                                        DateStart = request.DateStart.Value,
                                        DateEnd = request.DateEnd.Value,
                                        vgo = false
                                    };
                                    tasks.Add(t);
                                    DumpTaskToLog(t);

                                    break;
                                }

                            case "По одному договору":
                                {
                                    singleRow = SearchSingleByBranchInnKpp(
                                        csvPath,
                                        request.Branch.ToString(),
                                        request.INN?.ToString(),
                                        request.KPP?.ToString()
                                    );

                                    if (singleRow == null)
                                        throw new Exception("Не найдено ни одной строки по ИНН/КПП. Проверьте корректность ИНН/КПП.");

                                    var be = GetValueByName(request.Branch) ??
                                             throw new Exception("Не смог сопоставить БЕ из заявки с конфигом (BeCodes).");

                                    var signatoryArray = SplitSignatory(singleRow.PersonnelNumber);
                                    if (signatoryArray.Count == 0)
                                        throw new Exception("Ошибка: PersonnelNumber пустой для найденной строки.");

                                    var agreements = GetAgreementNumbersFromRequest(request.RegNumbDoc);
                                    if (agreements.Count == 0)
                                        throw new Exception("Для типа 'По одному договору' не переданы номера договоров.");

                                    var t = new SapTask
                                    {
                                        BeCode = be,
                                        INN = request.INN?.ToString(),
                                        KPP = request.KPP?.ToString(),
                                        ContractNumber = agreements,
                                        CounterpartyNumbers = new List<string> { singleRow.CompanyNumberSap },
                                        SignatoryNumbers = signatoryArray,
                                        Status = singleRow.Status,
                                        DateStart = request.DateStart.Value,
                                        DateEnd = request.DateEnd.Value,
                                        vgo = false
                                    };
                                    tasks.Add(t);
                                    DumpTaskToLog(t);

                                    break;
                                }

                            case "Годовая инвентаризация (по всем контрагентам и договорам)":
                                {
                                    var rowsByBranch = FilterCsvRows(csvPath, new Dictionary<int, HashSet<string>>
                                    {
                                        { 1, new HashSet<string>(new[] { request.Branch?.ToString() ?? "" }, StringComparer.OrdinalIgnoreCase) }
                                    });

                                    Log.Info($"[Inventory] Начальный выбор по Branch='{request.Branch}': строк={rowsByBranch.Count}");

                                    var Signatories = rowsByBranch
                                        .Select(r => (r.PersonnelNumber ?? "").Trim())
                                        .Where(s => !string.IsNullOrWhiteSpace(s))
                                        .Distinct(StringComparer.OrdinalIgnoreCase)
                                        .ToList();

                                    var fixedStatuses = new List<string> { "ЭДО", "Не ЭДО" };

                                    foreach (var signRaw in Signatories)
                                    {
                                        Log.Info($"Формирую задачу для подписанта: {signRaw}");
                                        var signValue = (signRaw ?? "").Trim();
                                        if (string.IsNullOrWhiteSpace(signValue))
                                            continue;

                                        foreach (var status in fixedStatuses)
                                        {
                                            Log.Info($"  Формирую задачу для статуса: {status}");

                                            var filtered = rowsByBranch
                                                .Where(r =>
                                                    EqualsCI(r.PersonnelNumber, signValue) &&
                                                    EqualsCI(r.Status, status))
                                                .ToList();

                                            var companyNumbers = filtered
                                                .Select(r => (r.CompanyNumberSap ?? "").Trim())
                                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                                .ToList();

                                            if (companyNumbers.Count == 0)
                                            {
                                                Log.Info("Результатов нет — задачу не добавляем.");
                                                continue;
                                            }

                                            var be = GetValueByName(request.Branch) ??
                                                     throw new Exception("Не смог сопоставить БЕ из заявки с конфигом (BeCodes).");

                                            var signatoryArray = SplitSignatory(signValue);

                                            var t = new SapTask
                                            {
                                                BeCode = be,
                                                INN = null,
                                                KPP = null,
                                                ContractNumber = new List<string>(),
                                                CounterpartyNumbers = companyNumbers,
                                                SignatoryNumbers = signatoryArray,
                                                Status = status,
                                                DateStart = request.DateStart.Value,
                                                DateEnd = request.DateEnd.Value,
                                                vgo = false
                                            };
                                            tasks.Add(t);
                                            DumpTaskToLog(t, indent: "");
                                        }
                                    }

                                    if (tasks.Count == 0)
                                    {
                                        Log.Warn("Инвентаризация: задачи не сформированы (нет совпадений).");
                                    }
                                    else
                                    {
                                        Log.Info($"Инвентаризация: сформировано задач: {tasks.Count}");
                                    }

                                    _jsonManager.SetValue("status", "OK");
                                    _jsonManager.SetValue("message", $"Сформировано задач: {tasks.Count}");
                                    _jsonManager.SaveToFile(currentFile);

                                    break;
                                }

                            default:
                                throw new Exception($"Неизвестный тип заявки: '{searchType}'.");
                        }

                        // ======= ДАЛЕЕ — ВЫПОЛНЯЕМ SAP-БЛОК ДЛЯ КАЖДОЙ ЗАДАЧИ (кейсы 1 и 2) =======
                        if (tasks == null || tasks.Count == 0)
                            throw new Exception("Не удалось подготовить задачу для SAP.");

                        foreach (var taskToRun in tasks)
                        {
                            _signatory1 = taskToRun.SignatoryNumbers.ElementAtOrDefault(0) ?? "";
                            _signatory2 = taskToRun.SignatoryNumbers.ElementAtOrDefault(1) ?? _signatory1;

                            string companyNumberSap = taskToRun.CounterpartyNumbers.FirstOrDefault() ?? "";
                            if (string.IsNullOrWhiteSpace(companyNumberSap))
                                throw new Exception("Ошибка: CompanyNumberSap не может быть пустым или null.");

                            try
                            {
                                sapfir = new SapfirManager(sapLogonPath);
                                sapfir.LaunchSAP();
                                Log.Info("SAP Logon успешно запущен.");

                                GuiSession session = sapfir.GetSapSession(stage);
                                if (session == null)
                                {
                                    throw new Exception("Не удалось получить сессию SAP.");
                                }
                                Log.Info("Сессия SAP успешно получена.");

                                if (stage == testStage)
                                {
                                    sapfir.LoginToSAP(session, sapUser, sapPassword);
                                }
                                string statusBarValue = sapfir.GetStatusMessage();

                                if (statusBarValue.Contains("Этот мандант сейчас блокирован для регистрации в нём."))
                                {
                                    Log.Error("Возникла ошибка при попытке входа в САП. Ошибка: Этот мандант сейчас блокирован для регистрации в нём.");
                                    throw new Exception("Возникла ошибка при попытке входа в САП. Ошибка: Этот мандант сейчас блокирован для регистрации в нём.");
                                }

                                Log.Info("Успешно выполнен вход в SAP.");

                                Thread.Sleep(2000);

                                try
                                {
                                    session.StartTransaction("ZTSF_AKT_SVERKI");
                                    Log.Info("Транзакция ZTSF_AKT_SVERKI успешно запущена.");
                                }
                                catch (Exception ex)
                                {
                                    Log.Error($"Ошибка при запуске транзакции ZTSF_AKT_SVERKI: {ex.Message}");
                                    throw;
                                }

                                Log.Info("Успешно выполнен вход и запущена транзакция ZTSF_AKT_SVERKI");

                                string be = taskToRun.BeCode ??
                                            throw new Exception("Не удалось определить BeCode для задачи.");

                                sapfir.SetText("wnd[0]/usr/ctxtP_BUKRS", be);
                                Thread.Sleep(500);
                                sapfir.SetText("wnd[0]/usr/ctxtS_BUDAT-LOW", taskToRun.DateStart.ToString("dd.MM.yyyy"));
                                Thread.Sleep(500);
                                sapfir.SetText("wnd[0]/usr/ctxtS_BUDAT-HIGH", taskToRun.DateEnd.ToString("dd.MM.yyyy"));

                                var fixedAccounts = new List<string>();
                                fixedAccounts.Clear();

                                if (request.Type == "Годовая инвентаризация (по всем контрагентам и договорам)")
                                {
                                    Thread.Sleep(1000);
                                    sapfir.RadioButton("wnd[0]/usr/radP_PROCH", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRSLD", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRSPP", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRALL", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRHKT", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_WAERS", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_DETAIL", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_AGREE", true);
                                    Thread.Sleep(1000);
                                    sapfir.PressButton("wnd[0]/usr/btn%_S_PARTN_%_APP_%-VALU_PUSH");
                                    Thread.Sleep(1000);
                                    Clipboard.SetText(string.Join(Environment.NewLine, taskToRun.CounterpartyNumbers));

                                    sapfir.PressButton("wnd[1]/tbar[0]/btn[24]");
                                    sapfir.PressButton("wnd[1]/tbar[0]/btn[8]");
                                    fixedAccounts.AddRange(new[] { "60", "62", "76" });
                                }

                                if (request.Type == "По одному контрагенту по всем договорам")
                                {
                                    Thread.Sleep(1000);
                                    sapfir.RadioButton("wnd[0]/usr/radP_PROCH", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRSLD", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRSPP", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRALL", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRHKT", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_WAERS", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_DETAIL", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_AGREE", true);
                                    Thread.Sleep(500);
                                    sapfir.SetText("wnd[0]/usr/ctxtS_PARTN-LOW", taskToRun.CounterpartyNumbers[0]);
                                    fixedAccounts.AddRange(new[] { "*" });
                                }

                                if (request.Type == "По одному договору")
                                {
                                    Thread.Sleep(1000);
                                    sapfir.RadioButton("wnd[0]/usr/radP_PROCH", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRSLD", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRSPP", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRALL", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRHKT", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_WAERS", false);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_AGREE", true);
                                    Thread.Sleep(1000);
                                    sapfir.GuiCheckBox("wnd[0]/usr/chkP_NULL", false);
                                    Thread.Sleep(1000);
                                    sapfir.SetText("wnd[0]/usr/ctxtS_PARTN-LOW", taskToRun.CounterpartyNumbers[0]);
                                    Thread.Sleep(1000);
                                    sapfir.SetText("wnd[0]/usr/ctxtS_ZUONR-LOW", request.RegNumbDoc);
                                    fixedAccounts.AddRange(new[] { "*" });
                                }

                                foreach (var account in fixedAccounts)
                                {
                                    try
                                    {
                                        if (account == "*")
                                        {
                                            sapfir.PressButton("wnd[0]/usr/btn%_S_HKONT_%_APP_%-VALU_PUSH");
                                            Thread.Sleep(2000);
                                            sapfir.SetText("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/" +
                                                               "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]", "*");
                                            sapfir.SelectTab("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV");
                                            Thread.Sleep(1000);

                                            string[] accounts62 = new string[] { "6201010101", "6201010201", "6201010301", "6201010401",
                                                                     "6201010501", "6201010601", "6201020101", "6201030101" };

                                            for (int i = 0; i < accounts62.Length; i++)
                                            {
                                                sapfir.SetText(
                                                    $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                                    accounts62[i]);
                                            }

                                            sapfir.SetVerticalScrollPosition(
                                                "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E", 7);
                                            Thread.Sleep(1000);

                                            string[] accounts62_2 = new string[] { "6201030201", "6201030301", "6201030401", "6201030501",
                                                                       "6201030601", "6201030701", "6201040201" };

                                            for (int i = 1; i <= accounts62_2.Length; i++)
                                            {
                                                sapfir.SetText(
                                                    $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                                    accounts62_2[i - 1]);
                                            }

                                            sapfir.SetVerticalScrollPosition(
                                                "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                                                14);
                                            Thread.Sleep(1000);

                                            string[] accounts62_3 = new string[] { "6201040301", "6201110101", "6201130101", "6202010101",
                                                                       "6202010201", "6202010301", "6202010401" };

                                            for (int i = 1; i <= accounts62_3.Length; i++)
                                            {
                                                sapfir.SetText(
                                                    $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                                    accounts62_3[i - 1]);
                                            }

                                            sapfir.SetVerticalScrollPosition(
                                                "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                                                21);
                                            Thread.Sleep(1000);

                                            string[] accounts62_4 = new string[] { "6202010501", "6202010601", "6202020101", "6202030101",
                                                                       "6202030201", "6202030301", "6202030401" };

                                            for (int i = 1; i <= accounts62_4.Length; i++)
                                            {
                                                sapfir.SetText(
                                                    $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                                    accounts62_4[i - 1]);
                                            }

                                            sapfir.SetVerticalScrollPosition(
                                                "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                                                28);
                                            Thread.Sleep(1000);

                                            string[] accounts62_5 = new string[] { "6202030501", "6202030601", "6202030701", "6202040201",
                                                                       "6202040301", "6202110101", "6202120101" };

                                            for (int i = 1; i <= accounts62_5.Length; i++)
                                            {
                                                sapfir.SetText(
                                                    $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                                    accounts62_5[i - 1]);
                                            }

                                            sapfir.SetVerticalScrollPosition(
                                                "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                                                35);
                                            Thread.Sleep(1000);

                                            string[] accounts76_1 = new string[] { "6202130101", "7602020101", "7602040101", "7611010101",
                                                                       "7615020101", "7602010102", "7602020102" };

                                            for (int i = 1; i <= accounts76_1.Length; i++)
                                            {
                                                sapfir.SetText(
                                                    $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                                    accounts76_1[i - 1]);
                                            }

                                            sapfir.SetVerticalScrollPosition(
                                                "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                                                42);
                                            Thread.Sleep(1000);

                                            string[] accounts76_2 = new string[] { "7602030102", "7602010101", "7602040102", "7615020102", "760903*" };

                                            for (int i = 1; i <= accounts76_2.Length; i++)
                                            {
                                                sapfir.SetText(
                                                    $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                                    accounts76_2[i - 1]);
                                            }
                                            sapfir.PressButton("wnd[1]/tbar[0]/btn[8]");
                                            Thread.Sleep(1000);
                                            sapfir.PressButton("wnd[0]/tbar[1]/btn[8]");
                                            Thread.Sleep(2000);
                                        }

                                        if (account == "60")
                                        {
                                            Log.Info("Переходим в ветку по счету - 60");
                                            sapfir.PressButton("wnd[0]/tbar[1]/btn[8]");
                                            Thread.Sleep(3000);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[0,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[1,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[2,0]", true);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[3,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[4,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[5,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[6,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[7,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[8,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[9,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[10,0]", false);
                                            Thread.Sleep(500);
                                            sapfir.GuiCheckBox("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0160/sub:SAPLSPO5:0160/chkSPOPLI-SELFLAG[11,0]", false);
                                            Thread.Sleep(1000);
                                            sapfir.PressButton("wnd[1]/tbar[0]/btn[0]");
                                        }

                                        if (account == "62")
                                        {
                                            Log.Info("Переходим в ветку по счету - 62");
                                            sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRHKT", true);
                                            sapfir.SetText("wnd[0]/usr/ctxtS_HKONT-LOW", "62*");
                                            sapfir.PressButton("wnd[0]/usr/btn%_S_HKONT_%_APP_%-VALU_PUSH");
                                            sapfir.SelectTab("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV");

                                            string cellBase = "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E";
                                            void Put(int rowIndex, string value) => sapfir.SetText($"{cellBase}[1,{rowIndex}]", value);

                                            Put(0, "6201010101");
                                            Put(1, "6201010201");
                                            Put(2, "6201010301");
                                            Put(3, "6201010401");
                                            Put(4, "6201010501");
                                            Put(5, "6201010601");
                                            Put(6, "6201020101");
                                            Put(7, "6201030101");

                                            sapfir.SetVerticalScrollPosition("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E", 7);

                                            Put(1, "6201030201");
                                            Put(2, "6201030301");
                                            Put(3, "6201030401");
                                            Put(4, "6201030501");
                                            Put(5, "6201030601");
                                            Put(6, "6201030701");
                                            Put(7, "6201040201");

                                            sapfir.SetVerticalScrollPosition("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E", 14);

                                            Put(1, "6201040301");
                                            Put(2, "6201110101");
                                            Put(3, "6201130101");
                                            Put(4, "6202010101");
                                            Put(5, "6202010201");
                                            Put(6, "6202010301");
                                            Put(7, "6202010401");

                                            sapfir.SetVerticalScrollPosition("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E", 21);

                                            Put(1, "6202010501");
                                            Put(2, "6202010601");
                                            Put(3, "6202020101");
                                            Put(4, "6202030101");
                                            Put(5, "6202030201");
                                            Put(6, "6202030301");
                                            Put(7, "6202030401");

                                            sapfir.SetVerticalScrollPosition("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E", 28);

                                            Put(1, "6202030501");
                                            Put(2, "6202030601");
                                            Put(3, "6202030701");
                                            Put(4, "6202040201");
                                            Put(5, "6202040301");
                                            Put(6, "6202110101");
                                            Put(7, "6202120101");

                                            sapfir.SetVerticalScrollPosition("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E", 35);

                                            Put(1, "6202130101");

                                            sapfir.PressButton("wnd[1]/tbar[0]/btn[8]");
                                            sapfir.PressButton("wnd[0]/tbar[1]/btn[8]");
                                            Thread.Sleep(2000);
                                        }

                                        if (account == "76")
                                        {
                                            Log.Info("Переходим в ветку по счету - 76");
                                            sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRHKT", true);
                                            sapfir.SetText("wnd[0]/usr/ctxtS_HKONT-LOW", "76*");
                                            sapfir.PressButton("wnd[0]/usr/btn%_S_HKONT_%_APP_%-VALU_PUSH");
                                            sapfir.SelectTab("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV");

                                            string cellBase76 = "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E";
                                            void Put76(int rowIndex, string value) => sapfir.SetText($"{cellBase76}[1,{rowIndex}]", value);

                                            Put76(0, "7602020101");
                                            Put76(1, "7602040101");
                                            Put76(2, "7611010101");
                                            Put76(3, "7615020101");
                                            Put76(4, "7602010102");
                                            Put76(5, "7602020102");
                                            Put76(6, "7602030102");
                                            Put76(7, "7602010101");

                                            sapfir.SetVerticalScrollPosition("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E", 7);

                                            Put76(1, "7602040102");
                                            Put76(2, "7615020102");
                                            Put76(3, "760903*");

                                            sapfir.PressButton("wnd[1]/tbar[0]/btn[8]");
                                            Thread.Sleep(2000);
                                            sapfir.PressButton("wnd[0]/tbar[1]/btn[8]");
                                        }

                                        Thread.Sleep(3000);
                                        string windowActive = sapfir.GetFrameText();
                                        statusBarValue = sapfir.GetStatusMessage();

                                        if (statusBarValue.Contains("не найдены"))
                                            Log.Info($"Ожидаю появления окна \"Акт сверки расчетов с контрагентами: ALV отчет\", текущее окно: {windowActive}");

                                        if (windowActive != "Акт сверки расчетов с контрагентами: ALV отчет" &&
                                            !string.IsNullOrWhiteSpace(statusBarValue))
                                        {
                                            Log.Error($"Данные по контрагенту не найдены или допущены ошибки в заполнении. Статус окна: {statusBarValue}. Перехожу к следующей задаче.");
                                            throw new Exception();
                                        }

                                        GuiShell shell = session.FindById("wnd[0]/shellcont/shell") as GuiShell;
                                        GuiGridView grid = shell as GuiGridView;
                                        if (grid != null)
                                        {
                                            grid.SetCurrentCell(-1, "");
                                            grid.SelectAll();
                                        }
                                        else
                                        {
                                            throw new Exception("Не удалось привести shell к типу GuiGridView");
                                        }

                                        sapfir.PressButton("wnd[0]/tbar[1]/btn[16]");
                                        Thread.Sleep(2000);
                                        sapfir.PressButton("wnd[0]/tbar[1]/btn[13]");
                                        Thread.Sleep(5000);

                                        bool buttonGenerateDocuments = false;

                                        try
                                        {
                                            GuiButton button = session.FindById("wnd[1]/usr/btnBUTTON_1") as GuiButton;
                                            if (button != null)
                                            {
                                                sapfir.PressButton("wnd[1]/usr/btnBUTTON_1");
                                                buttonGenerateDocuments = true;
                                                Log.Warn("По данному контрагенту уже есть сформированные документы. Выполнено повторное формирование.");
                                                Thread.Sleep(2000);
                                            }
                                        }
                                        catch
                                        {
                                            buttonGenerateDocuments = false;
                                        }

                                        sapfir.ShowContextText("wnd[1]/usr/ctxtP_RUKOV1");
                                        sapfir.SendKey("wnd[1]", 4);
                                        sapfir.PressButton("wnd[2]/tbar[0]/btn[17]");
                                        Thread.Sleep(1000);

                                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                           "sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]", _signatory1);
                                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[1,24]", "");
                                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[5,24]", "");
                                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]", "");
                                        sapfir.ShowContextText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                               "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]");
                                        Thread.Sleep(1000);

                                        sapfir.PressButton("wnd[2]/tbar[0]/btn[0]");
                                        sapfir.PressButton("wnd[2]/tbar[0]/btn[0]");

                                        sapfir.ShowContextText("wnd[1]/usr/ctxtP_RUKOV1");

                                        sapfir.ShowContextText("wnd[1]/usr/ctxtP_BUGAL2");
                                        Thread.Sleep(1000);

                                        sapfir.SendKey("wnd[1]", 4);
                                        sapfir.PressButton("wnd[2]/tbar[0]/btn[17]");

                                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                           "sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]", _signatory2);
                                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[1,24]", "");
                                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[5,24]", "");
                                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]", "");
                                        sapfir.ShowContextText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                                               "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]");
                                        Thread.Sleep(1000);

                                        sapfir.PressButton("wnd[2]/tbar[0]/btn[0]");
                                        Thread.Sleep(1000);
                                        sapfir.PressButton("wnd[2]/tbar[0]/btn[0]");
                                        Thread.Sleep(1000);

                                        Log.Info(sapfir.GetTextFromField("wnd[1]/usr/ctxtP_BUGAL2"));

                                        sapfir.PressButton("wnd[1]/tbar[0]/btn[8]");
                                        Thread.Sleep(1000);
                                        sapfir.PressButton("wnd[1]/usr/btnBUTTON_1");
                                        Thread.Sleep(1000);

                                        statusBarValue = sapfir.GetStatusMessage();
                                        Log.Info($"Статус SAP: {statusBarValue}");

                                        sapfir.PressButton("wnd[0]/tbar[1]/btn[17]");

                                        shell = session.FindById("wnd[0]/shellcont/shell") as GuiShell;
                                        grid = shell as GuiGridView;
                                        if (grid != null)
                                        {
                                            grid.SetCurrentCell(-1, "");
                                            grid.SelectAll();
                                            grid.PressToolbarButton("&MB_FILTER");
                                            sapfir.PressButton("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH");
                                            sapfir.SelectTab("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV");
                                            sapfir.SetText("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/" +
                                                               "tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]", "@5b@");
                                            sapfir.PressButton("wnd[2]/tbar[0]/btn[8]");
                                            sapfir.PressButton("wnd[1]/tbar[0]/btn[0]");
                                            Thread.Sleep(5000);
                                            grid.SelectAll();
                                            sapfir.PressButton("wnd[0]/tbar[1]/btn[16]");
                                            grid.PressToolbarContextButton("&MB_EXPORT");
                                            grid.SelectContextMenuItem("&PC");
                                            sapfir.RadioButton("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/" +
                                                                   "radSPOPLI-SELFLAG[1,0]", true);
                                            sapfir.PressButton("wnd[1]/tbar[0]/btn[0]");

                                            _appFolders.TryGetValue("error", out string errorFolder);
                                            sapfir.SetText("wnd[1]/usr/ctxtDY_PATH", errorFolder);
                                            string filename_err = $"errorsAC_{taskToRun.BeCode}_{account}_{taskToRun.Status}.xls";
                                            sapfir.SetText("wnd[1]/usr/ctxtDY_FILENAME", filename_err);
                                            sapfir.PressButton("wnd[1]/tbar[0]/btn[0]");

                                            // === ЧТЕНИЕ errors-файла через Interop.Excel ===
                                            try
                                            {
                                                string errFullPath = Path.Combine(errorFolder ?? "", filename_err);
                                                // даём ОС договорить запись файла
                                                Thread.Sleep(1500);

                                                if (File.Exists(errFullPath))
                                                {
                                                    bool hasNoData = ExcelContainsText(errFullPath, "Список не содержит данных");
                                                    if (!hasNoData)
                                                    {
                                                        string msg = $"Есть несформированные АС. Файл с ошибками сохранён по пути: {errFullPath}. Статус: {taskToRun.Status}.";
                                                        Log.Warn(msg);
                                                        _jsonManager.SetValue("status", "error");
                                                        _jsonManager.SetValue("message", msg);
                                                        _jsonManager.SaveToFile(currentFile);
                                                    }
                                                    else
                                                    {
                                                        Log.Info($"Файл ошибок сохранён: {errFullPath}. Фраза 'Список не содержит данных' не найдена.");
                                                    }
                                                }
                                                else
                                                {
                                                    Log.Warn($"Ожидаемый файл ошибок не найден: {errFullPath}");
                                                }
                                            }
                                            catch (Exception readErrEx)
                                            {
                                                Log.Error(readErrEx, "Не удалось прочитать или проанализировать файл ошибок Excel.");
                                            }
                                            // === конец блока Interop.Excel ===

                                            Thread.Sleep(1000);

                                            grid.SetCurrentCell(-1, "");
                                            grid.SelectAll();
                                            grid.PressToolbarButton("&MB_FILTER");
                                            sapfir.SetText("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW", "");
                                            sapfir.PressButton("wnd[1]/tbar[0]/btn[14]");
                                            sapfir.SetText("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW", "@5b@");
                                            sapfir.PressButton("wnd[1]/tbar[0]/btn[0]");
                                            grid.SetCurrentCell(-1, "");
                                            grid.SelectAll();
                                            sapfir.PressButton("wnd[0]/tbar[1]/btn[16]");

                                            sapfir.PressButton("wnd[0]/tbar[1]/btn[20]");
                                            Log.Info(@"Нажали на кнопку ""Сформировать документы""");
                                            Thread.Sleep(1000);
                                            if (sapfir.ElementExists("wnd[1]/usr/btnBUTTON_1"))
                                            {
                                                // Получаем активное окно
                                                GuiFrameWindow wnd = sapfir.SapSession.ActiveWindow as GuiFrameWindow;
                                                if (wnd != null)
                                                {
                                                    
                                                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                                                    Log.Info("Элемент найден. Нажата клавиша Enter.");
                                                }
                                                else
                                                {
                                                    Log.Warn("Элемент найден, но активное окно не получено.");
                                                }
                                            }
                                            

                                            string savePath = @"C:\Users\RobinSapAC\Desktop";
                                            string fullPath = Path.Combine(savePath);

                                            Thread.Sleep(2000);

                                            const int MaxAttempts = 3;
                                            const int TimeoutSeconds = 30;

                                            using (var automation = new UIA3Automation())
                                            {
                                                bool mainWindowFound = false;

                                                for (int attempt = 1; attempt <= MaxAttempts; attempt++)
                                                {
                                                    try
                                                    {
                                                        var desktop = automation.GetDesktop();

                                                        var browseWindow =
                                                            desktop
                                                                .FindFirstChild(cf =>
                                                                                    cf.ByName("Browse for Files or Folders")
                                                                                        .Or(cf.ByControlType(
                                                                                            FlaUI.Core.Definitions.ControlType.Window)))
                                                                ?.AsWindow();

                                                        DateTime startTime = DateTime.Now;
                                                        while (browseWindow == null &&
                                                               (DateTime.Now - startTime).TotalSeconds < TimeoutSeconds)
                                                        {
                                                            Thread.Sleep(1000);
                                                            browseWindow =
                                                                desktop
                                                                    .FindFirstChild(
                                                                        cf => cf.ByName("Browse for Files or Folders")
                                                                                  .Or(cf.ByControlType(
                                                                                      FlaUI.Core.Definitions.ControlType.Window)))
                                                                    ?.AsWindow();
                                                        }

                                                        if (browseWindow == null)
                                                        {
                                                            Log.Info($"Попытка {attempt}: Окно 'Browse for Files or Folders' не найдено");
                                                            if (attempt == MaxAttempts)
                                                            {
                                                                throw new Exception("Окно сохранения не найдено после всех попыток.");
                                                            }
                                                            continue;
                                                        }

                                                        Log.Info("Окно 'Browse for Files or Folders' найдено");

                                                        var folderEdit =
                                                            browseWindow
                                                                .FindFirstDescendant(
                                                                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                                                                              .And(cf.ByName("Folder:")))
                                                                ?.AsTextBox();

                                                        if (folderEdit == null)
                                                        {
                                                            Log.Info($"Попытка {attempt}: Поле ввода 'Folder:' не найдено");
                                                            if (attempt == MaxAttempts)
                                                            {
                                                                throw new Exception("Поле ввода пути сохранения не найдено после всех попыток.");
                                                            }
                                                            continue;
                                                        }

                                                        Log.Info("Поле ввода 'Folder:' найдено");

                                                        folderEdit.Text = "";
                                                        folderEdit.Text = fullPath;
                                                        Log.Info($"Установлен путь сохранения: {fullPath}");

                                                        var okButton =
                                                            browseWindow
                                                                .FindFirstDescendant(cf => cf.ByName("OK").And(cf.ByControlType(
                                                                                     FlaUI.Core.Definitions.ControlType.Button)))
                                                                ?.AsButton();

                                                        if (okButton == null)
                                                        {
                                                            Log.Info($"Попытка {attempt}: Кнопка 'OK' не найдена");
                                                            if (attempt == MaxAttempts)
                                                            {
                                                                throw new Exception("Кнопка 'OK' не найдена после всех попыток.");
                                                            }
                                                            continue;
                                                        }

                                                        okButton.Click();
                                                        Log.Info("Акт(-ы) сверки успешно сохранены.");
                                                        break;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Log.Error($"Попытка {attempt}: Ошибка при сохранении файла - {ex.Message}");
                                                        if (attempt == MaxAttempts)
                                                        {
                                                            throw new Exception("Не удалось сохранить файл после всех попыток.", ex);
                                                        }
                                                        Thread.Sleep(1000);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            throw new Exception("Не удалось привести shell к типу GuiGridView для установки текущей ячейки.");
                                        }
                                        sapfir.PressButton("wnd[0]/tbar[0]/btn[3]");
                                    }
                                    catch (Exception accEx)
                                    {
                                        Log.Error(accEx, $"Ошибка при обработке счета '{account}'. Пропускаю этот счёт и перехожу к следующему.");
                                        sapfir.PressButton("wnd[0]/tbar[0]/btn[3]");
                                        continue;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Log.Error(ex, "Ошибка SAP (выполнение задачи)");
                                sapfir.CloseSAPWindowByNex(session);
                                continue;
                            }
                            finally
                            {
                                sapfir.CloseSAPWindowByNex(session);
                                sapfir?.KillSAP();
                            }
                        }

                        _jsonManager.SetValue("status", "OK");
                        _jsonManager.SetValue("message", "Успешно обработано.");
                        _jsonManager.SaveToFile(currentFile);
                    }
                    catch (Exception ex)
                    {
                        Log.Error($"Ошибка при обработке входного файла {fileName}. Переход к следующему файлу.");
                        sapfir?.KillSAP();

                        try
                        {
                            sapfir.CloseSAPWindowByNex(session);
                            _jsonManager?.SetValue("status", "error");
                            _jsonManager?.SetValue("message", ex.Message);
                            _jsonManager?.SaveToFile(currentFile);
                        }
                        catch (Exception jsonEx)
                        {
                            Log.Error(jsonEx, "Ошибка при записи статуса в JSON.");
                        }

                        continue;
                    }
                }

                Log.Info("Файлов для обработки нет. Робот завершает свою работу.");
            }
            catch (Exception ex)
            {
                Log.Fatal($"Ошибка при инициализации программы: {ex.Message}");
                throw;
            }
        }

        // ===== УНИВЕРСАЛЬНАЯ ФИЛЬТРАЦИЯ CSV =====
        private static List<CsvRow> FilterCsvRows(string csvPath, Dictionary<int, HashSet<string>> filters, int requiredColumns = 11)
        {
            var rows = new List<CsvRow>();

            foreach (var line in File.ReadLines(csvPath).Skip(1))
            {
                var columns = SplitCsvLine(line);

                if (columns.Length < requiredColumns)
                    continue;

                for (int i = 0; i < columns.Length; i++)
                    columns[i] = (columns[i] ?? string.Empty).Trim().Trim('\"');

                bool match = true;
                foreach (var kv in filters)
                {
                    int colIndex = kv.Key;
                    var allowed = kv.Value;

                    if (colIndex < 0 || colIndex >= columns.Length)
                    {
                        match = false;
                        break;
                    }

                    string value = columns[colIndex];
                    if (string.IsNullOrWhiteSpace(value) || !allowed.Contains(value))
                    {
                        match = false;
                        break;
                    }
                }

                if (!match) continue;

                rows.Add(new CsvRow
                {
                    Number = Get(columns, 0),
                    Branch = Get(columns, 1),
                    CompanyName = Get(columns, 2),
                    CompanyNumberSap = Get(columns, 3),
                    INN = Get(columns, 4),
                    KPP = Get(columns, 5),
                    Status = Get(columns, 6),
                    SignatoryLanDocs = Get(columns, 7),
                    PersonnelNumber = Get(columns, 8),
                    SignatoryLandocs = Get(columns, 9),
                    VGO = Get(columns, 10)
                });
            }

            return rows;

            string Get(string[] arr, int idx) =>
                (idx >= 0 && idx < arr.Length ? arr[idx] : "") ?? "";
        }

        // Разбор CSV-строки: поддержка кавычек и запятых, "" => "
        private static string[] SplitCsvLine(string line)
        {
            if (line == null)
                return Array.Empty<string>();

            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }

            result.Add(sb.ToString());
            return result.ToArray();
        }

        private static CsvRow SearchSingleByBranchInnKpp(string csvPath, string branch, string inn, string kpp)
        {
            var filters = new Dictionary<int, HashSet<string>>
            {
                { 1, new HashSet<string>(new[] { branch ?? ""    }, StringComparer.OrdinalIgnoreCase) },
                { 4, new HashSet<string>(new[] { inn ?? ""    }, StringComparer.OrdinalIgnoreCase) },
                { 5, new HashSet<string>(new[] { kpp ?? ""    }, StringComparer.OrdinalIgnoreCase) }
            };

            var rows = FilterCsvRows(csvPath, filters);

            Log.Info($"Найдено строк по INN/KPP: {rows.Count}.");

            if (rows.Count > 1)
            {
                const string msg = "При выборке по ИНН и КПП в реестре КА нашлось несколько результатов. Просьба проверить корректность данных";
                Log.Error(msg);
                foreach (var r in rows)
                {
                    Log.Error(string.Join(" | ", new[] { r.Number, r.Branch, r.CompanyName, r.CompanyNumberSap, r.INN, r.KPP, r.Status, r.PersonnelNumber }));
                }
                throw new Exception(msg);
            }

            return rows.FirstOrDefault();
        }

        private static string GetValueByName(string name)
        {
            return _beCodes != null && _beCodes.TryGetValue(name, out string value) ? value : null;
        }

        private static List<string> SplitSignatory(string raw)
        {
            var delims = new[] { ';', '\\', '/' };
            return (raw ?? "")
                   .Split(delims, StringSplitOptions.RemoveEmptyEntries)
                   .Select(s => s.Trim())
                   .Where(s => !string.IsNullOrWhiteSpace(s))
                   .Distinct(StringComparer.OrdinalIgnoreCase)
                   .ToList();
        }

        private static List<string> GetAgreementNumbersFromRequest(object contractField)
        {
            if (contractField == null) return new List<string>();

            if (contractField is IEnumerable<string> list)
            {
                return list.Where(s => !string.IsNullOrWhiteSpace(s))
                           .Select(s => s.Trim())
                           .Distinct(StringComparer.OrdinalIgnoreCase)
                           .ToList();
            }

            var sraw = contractField.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(sraw)) return new List<string>();

            var parts = sraw.Split(new[] { ';', ',', '/', '\\', '\n', '\r', '\t', ' ' },
                                   StringSplitOptions.RemoveEmptyEntries);
            return parts.Select(p => p.Trim())
                        .Where(p => !string.IsNullOrWhiteSpace(p))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();
        }

        private static bool EqualsCI(string a, string b) =>
            string.Equals(a?.Trim(), b?.Trim(), StringComparison.OrdinalIgnoreCase);

        private static void DumpTaskToLog(SapTask t, string indent = "")
        {
            Log.Info($"----Задача для SAP----");
            Log.Info($"BeCode: {t.BeCode}");
            Log.Info($"INN: {(t.INN ?? "null")}, KPP: {(t.KPP ?? "null")}");
            Log.Info($"ContractNumber: [{string.Join(", ", t.ContractNumber)}]");
            Log.Info($"CounterpartyNumbers: [{string.Join(", ", t.CounterpartyNumbers)}]");
            Log.Info($"SignatoryNumbers: [{string.Join(", ", t.SignatoryNumbers)}]");
            Log.Info($"Status: {t.Status}");
            Log.Info($"DateStart: {t.DateStart:dd.MM.yyyy}, DateEnd: {t.DateEnd:dd.MM.yyyy}");
            Log.Info($"vgo: {t.vgo}");
        }

        private static string MoveToOwnFolderInInput(string file)
        {
            try
            {
                var inputDir = Path.GetDirectoryName(file) ?? AppDomain.CurrentDomain.BaseDirectory;
                var fileName = Path.GetFileName(file);
                var baseName = Path.GetFileNameWithoutExtension(fileName);

                var targetDir = Path.Combine(inputDir, baseName);
                Directory.CreateDirectory(targetDir);

                var destPath = Path.Combine(targetDir, fileName);
                if (File.Exists(destPath))
                {
                    var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    destPath = Path.Combine(targetDir, $"{baseName}_{stamp}{Path.GetExtension(fileName)}");
                }

                File.Move(file, destPath);
                Log.Info($"Файл '{fileName}' перемещён в '{targetDir}'.");
                return destPath;
            }
            catch (Exception moveEx)
            {
                Log.Error(moveEx, "Не удалось переместить входной файл в подпапку внутри input.");
                return file;
            }
        }

        private static bool ExcelContainsText(string filePath, string needle)
        {
            if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(needle))
                return false;

            Excel.Application app = null;
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;

            try
            {
                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wbs = app.Workbooks;
                wb = wbs.Open(filePath, ReadOnly: true, Editable: false);

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    Excel.Range used = null;
                    try
                    {
                        used = ws.UsedRange;
                        object val = used?.Value2;

                        if (val == null)
                            continue;

                        // Value2 может быть:
                        // - string/num для одной ячейки
                        // - object[,] для диапазона
                        if (val is object[,] arr)
                        {
                            int rows = arr.GetLength(0);
                            int cols = arr.GetLength(1);
                            for (int r = 1; r <= rows; r++)
                            {
                                for (int c = 1; c <= cols; c++)
                                {
                                    var cellVal = arr[r, c];
                                    if (cellVal != null)
                                    {
                                        string s = cellVal.ToString();
                                        if (!string.IsNullOrEmpty(s) &&
                                            s.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            string s = val.ToString();
                            if (!string.IsNullOrEmpty(s) &&
                                s.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return true;
                            }
                        }
                    }
                    finally
                    {
                        if (used != null) Marshal.ReleaseComObject(used);
                        Marshal.ReleaseComObject(ws);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Log.Error(ex, $"Ошибка при чтении Excel '{filePath}' через Interop.");
                return false;
            }
            finally
            {
                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                    Marshal.ReleaseComObject(wb);
                }
                if (wbs != null) Marshal.ReleaseComObject(wbs);
                if (app != null)
                {
                    try { app.Quit(); } catch { }
                    Marshal.ReleaseComObject(app);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
