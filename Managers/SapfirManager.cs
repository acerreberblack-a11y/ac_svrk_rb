using FlaUI.UIA3;
using SAPFEWSELib;
using SapROTWr;
using System;
using System.Diagnostics;
using System.Threading;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using NLog;

namespace SpravkoBot_AsSapfir
{
    public class SapfirManager
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public GuiApplication SapApp { get; private set; }
        public GuiConnection SapConnection { get; private set; }
        public GuiSession SapSession { get; private set; }

        private readonly string _pathSAPLoginApp;

        public SapfirManager(string pathSAPLoginApp)
        {
            _pathSAPLoginApp = pathSAPLoginApp ?? throw new ArgumentNullException(nameof(pathSAPLoginApp),
                                                                                  "Путь к SAP Logon не может быть null.");
        }

        public void LaunchSAP()
        {
            using (var automation = new UIA3Automation())
            {
                /*                // Если процесс уже запущен, просто выходим
                                if (Process.GetProcessesByName("saplogon").Length > 0)
                                {
                                    return;
                                }*/

                Application app = null;

                try
                {
                    // Запуск через FlaUI
                    app = Application.Launch(_pathSAPLoginApp);

                    try
                    {
                        var desktop = automation.GetDesktop();
                        Window mainWindow = null;
                        var startTime = DateTime.Now;

                        // Ожидание окна до 20 секунд
                        while (mainWindow == null && (DateTime.Now - startTime).TotalSeconds < 20)
                        {
                            mainWindow =
                                desktop
                                    .FindFirstChild(cf => cf.ByName("SAP Logon 750")
                                                              .And(cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window)))
                                    ?.AsWindow();

                            if (mainWindow == null)
                            {
                                Thread.Sleep(1000); // Пауза перед следующей попыткой
                            }
                        }

                        if (mainWindow == null)
                        {
                            throw new TimeoutException("Не удалось запустить SAP Logon в течение заданного времени.");
                        }

                        // Дополнительное ожидание для стабилизации
                        Thread.Sleep(2000);
                    }
                    catch (Exception)
                    {
                        app?.Close(); // Закрываем приложение при ошибке
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Ошибка при запуске SAP Logon");
                    throw new Exception($"Не удалось запустить SAP Logon: {ex.Message}", ex);
                }
            }
        }

        public void KillSAP()
        {
            try
            {
                foreach (var process in Process.GetProcessesByName("saplogon"))
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit(5000); // Ожидание завершения процесса до 5 секунд
                    }
                    catch (Exception ex)
                    {
                        Log.Warn(ex, "Не удалось корректно завершить процесс saplogon");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при поиске процессов saplogon");
            }
        }

        public GuiSession GetSapSession(string sapSystemName)
        {
            if (string.IsNullOrEmpty(sapSystemName))
                throw new ArgumentNullException(nameof(sapSystemName), "Название системы SAP не может быть пустым.");

            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            object sapGuiRot = sapROTWrapper.GetROTEntry("SAPGUI");
            if (sapGuiRot == null)
                throw new Exception("Не удалось получить ROT Entry для SAPGUI.");

            object engine = sapGuiRot.GetType().InvokeMember(
                "GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, sapGuiRot, null);
            if (engine == null)
                throw new Exception("Не удалось получить Scripting Engine.");

            SapApp =
                engine as GuiApplication ?? throw new Exception("Не удалось привести Scripting Engine к GuiApplication.");
            SapConnection =
                SapApp.OpenConnection(sapSystemName, false, true); // false - не асинхронно, true - ждать завершения
            SapSession =
                SapConnection.Children.ElementAt(0) as GuiSession ?? throw new Exception("Не удалось получить сессию.");

            return SapSession;
        }

        public GuiSession TryGetCurrentWindowSession()
        {
            try
            {
                CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
                object sapGuiRot = sapROTWrapper.GetROTEntry("SAPGUI");
                if (sapGuiRot == null)
                    return null;

                object engine = sapGuiRot.GetType().InvokeMember(
                    "GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, sapGuiRot, null);
                if (engine == null)
                    return null;

                GuiApplication sapApp = engine as GuiApplication;
                if (sapApp?.Connections.Count == 0)
                    return null;

                GuiConnection connection = sapApp.Connections.ElementAt(0) as GuiConnection;
                if (connection?.Children.Count == 0)
                    return null;

                return connection.Children.ElementAt(0) as GuiSession;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при получении текущей сессии");
                return null;
            }
        }

        public void LoginToSAP(GuiSession session, string user, string password)
        {
            if (session == null)
                throw new ArgumentNullException(nameof(session), "Сессия не может быть null.");
            if (string.IsNullOrEmpty(user))
                throw new ArgumentNullException(nameof(user), "Имя пользователя не может быть пустым.");
            if (string.IsNullOrEmpty(password))
                throw new ArgumentNullException(nameof(password), "Пароль не может быть пустым.");

            Thread.Sleep(1000); // Ожидание для стабилизации интерфейса
            GuiMainWindow mainWindow =
                session.FindById("wnd[0]") as GuiMainWindow ?? throw new Exception("Главное окно wnd[0] не найдено.");
            GuiTextField userField = session.FindById("wnd[0]/usr/txtRSYST-BNAME") as GuiTextField ??
                                     throw new Exception("Поле логина не найдено.");
            GuiTextField pwdField = session.FindById("wnd[0]/usr/pwdRSYST-BCODE") as GuiTextField ??
                                    throw new Exception("Поле пароля не найдено.");

            userField.Text = user;
            pwdField.Text = password;
            mainWindow.Maximize();
            mainWindow.SendVKey(0); // Enter
        }

        public void CloseSAPWindowByNex(GuiSession session)
        {
            if (session == null)
            {
                Log.Warn("CloseSAPWindowByNex вызван без активной сессии.");
                return;
            }

            try
            {
                var mainWindow = session.FindById("wnd[0]") as GuiMainWindow;
                if (mainWindow == null)
                {
                    Log.Warn("Главное окно 'wnd[0]' не найдено при попытке закрытия.");
                    return;
                }

                // Выполняем команду /nex напрямую через HardCopy
                session.SendCommand("/nex");
            }
            catch (Exception ex)
            {
                Log.Error(ex, $"Ошибка при закрытии окна SAP: {ex.Message}");
            }
        }

        public void ToolBarButton(string id)
        {
            GuiToolbarControl toolbar = SapSession.FindById(id) as GuiToolbarControl ??
                                        throw new Exception($"Кнопка панели инструментов с ID {id} не найдена.");
            toolbar.SetFocus();
            toolbar.PressButton(id); // Исправлено на PressButton вместо SelectMenuItem
        }

        public void SendKey(string id, int numberKey)
        {
            GuiFrameWindow window =
                SapSession.FindById(id) as GuiFrameWindow ?? throw new Exception($"Окно с ID {id} не найдено.");
            window.SetFocus();
            window.SendVKey(numberKey);
        }

        public void SetText(string id, string text)
        {
            GuiComponent component =
                SapSession.FindById(id) as GuiComponent ?? throw new Exception($"Компонент с ID {id} не найден.");
            if (component.Type == "GuiTextField" || component.Type == "GuiCTextField")
            {
                dynamic textField = component; // Динамическое приведение для поддержки обоих типов
                textField.SetFocus();
                textField.Text = text ?? string.Empty;
            }
            else
            {
                throw new Exception($"Компонент с ID {id} не является текстовым полем (тип: {component.Type}).");
            }
        }

        public void GuiShellPress(string id, string context, bool raiseExceptionIfNotFound = true)
        {
            GuiShell shell = SapSession.FindById(id, raiseExceptionIfNotFound) as GuiShell;
            if (shell != null && shell.Type == "GuiGridView")
            {
                GuiGridView grid = shell as GuiGridView;
                grid.PressToolbarContextButton(context);
            }
            else if (raiseExceptionIfNotFound)
            {
                throw new Exception($"GuiShell с ID {id} не найден или не является GuiGridView.");
            }
        }

        public void PressButton(string id)
        {
            GuiButton button = SapSession.FindById(id) as GuiButton ?? throw new Exception($"Кнопка с ID {id} не найдена.");
            button.SetFocus();
            button.Press();
        }

        public void ContextMenu(string id, string key)
        {
            GuiGridView grid =
                SapSession.FindById(id) as GuiGridView ?? throw new Exception($"GridView с ID {id} не найден.");
            grid.ContextMenu();
            grid.SelectContextMenuItem(key);
        }

        public bool GuiCheckBox(string id, bool status)
        {
            GuiCheckBox checkbox =
                SapSession.FindById(id) as GuiCheckBox ?? throw new Exception($"Чекбокс с ID {id} не найден.");
            checkbox.SetFocus();
            checkbox.Selected = status;
            return true;
        }

        public bool RadioButton(string id, bool status)
        {
            GuiRadioButton radio =
                SapSession.FindById(id) as GuiRadioButton ?? throw new Exception($"Радиокнопка с ID {id} не найдена.");
            radio.SetFocus();
            radio.Selected = status;
            return true;
        }

        public string GetLabel(string id)
        {
            GuiLabel label = SapSession.FindById(id) as GuiLabel ?? throw new Exception($"Метка с ID {id} не найдена.");
            label.SetFocus();
            return label.Text ?? string.Empty;
        }

        public void SelectTab(string id)
        {
            GuiTab tab = SapSession.FindById(id) as GuiTab ?? throw new Exception($"Вкладка с ID {id} не найдена.");
            tab.SetFocus();
            tab.Select();
        }

        public string ShowContextText(string id)
        {
            GuiCTextField textField = SapSession.FindById(id) as GuiCTextField ??
                                      throw new Exception($"Поле с ID {id} не найдено или не является GuiCTextField.");
            textField.SetFocus();
            string text = textField.Text;
            textField.ShowContextMenu();
            return text;
        }

        public string GetComponentType(string id)
        {
            GuiComponent component = SapSession.FindById(id) as GuiComponent;
            return component?.Type ?? string.Empty;
        }

        public void SetComboBox(string id, string key)
        {
            GuiComboBox combo =
                SapSession.FindById(id) as GuiComboBox ?? throw new Exception($"Комбобокс с ID {id} не найден.");
            combo.SetFocus();
            combo.Key = key; // Используем string вместо int для большей гибкости
        }

        public void CloseFrameWindow(string id)
        {
            GuiFrameWindow window =
                SapSession.FindById(id) as GuiFrameWindow ?? throw new Exception($"Окно с ID {id} не найдено.");
            window.SetFocus();
            window.Close();
        }

        public string GetFrameText()
        {
            GuiFrameWindow window = SapSession.ActiveWindow as GuiFrameWindow;
            return window?.Text ?? string.Empty;
        }

        public void CloseFrame()
        {
            GuiFrameWindow window = SapSession.ActiveWindow as GuiFrameWindow;
            window?.Close();
        }

        public void SetVerticalScrollPosition(string tableId, int position)
        {
            GuiTableControl table = SapSession.FindById(tableId) as GuiTableControl ??
                                    throw new Exception($"Таблица с ID {tableId} не найдена.");
            table.VerticalScrollbar.Position = position;
        }

        public string GetStatusMessage()
        {
            GuiStatusbar statusbar = SapSession.FindById("wnd[0]/sbar") as GuiStatusbar;
            return statusbar?.Text ?? "Статус-бар не найден.";
        }

        public string CheckPopupMessage()
        {
            if (SapSession.Children.Count > 1)
            {
                GuiModalWindow popup = SapSession.FindById("wnd[1]") as GuiModalWindow;
                return popup?.Text ?? "Всплывающее окно не содержит текста.";
            }
            return "Нет всплывающих окон.";
        }

        public int GetWindowIndexByTitle(string windowTitle)
        {
            if (string.IsNullOrEmpty(windowTitle))
                throw new ArgumentNullException(nameof(windowTitle), "Заголовок окна не может быть пустым.");

            for (int i = 0; i < SapSession.Children.Count; i++)
            {
                GuiMainWindow window = SapSession.FindById($"wnd[{i}]") as GuiMainWindow;
                if (window != null && window.Text.Contains(windowTitle))
                {
                    return i;
                }
            }
            return -1;
        }

        public string GetTextFromField(string id)
        {
            GuiComponent component =
                SapSession.FindById(id) as GuiComponent ?? throw new Exception($"Компонент с ID {id} не найден.");
            if (component.Type == "GuiCTextField" || component.Type == "GuiTextField")
            {
                dynamic textField = component;
                return textField.Text ?? string.Empty;
            }
            throw new Exception($"Компонент с ID {id} не является текстовым полем (тип: {component.Type}).");
        }

        // Проверка существования любого SAP GUI элемента по Id
        public bool ElementExists(string id)
        {
            try
            {
                // В перегрузке FindById(string id, bool Raise) false = не бросать исключение
                var obj = SapSession?.FindById(id, false);
                return obj != null;
            }
            catch
            {
                return false;
            }
        }
    }
}