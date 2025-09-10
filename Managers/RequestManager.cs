using NLog;
using System;
using System.Globalization;

namespace SpravkoBot_AsSapfir
{
    internal static class RequestManager
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        public static Request FromJson(JsonManager jsonManager)
        {
            try
            {
                // Базовые строки
                string UUID = jsonManager.GetValue<string>("title"); // Номер заявки
                string Organization = jsonManager.GetValue<string>("organiz.title") ?? jsonManager.GetValue<string>("orgEmpl.title"); // Организация
                string Branch = jsonManager.GetValue<string>("orgfilial.title") ?? jsonManager.GetValue<string>("orgFil.title"); // Филиал
                string Type = jsonManager.GetValue<string>("formType.title") ?? jsonManager.GetValue<string>("formTypeInt.title"); // Вид формирования
                string docNumber = jsonManager.GetValue<string>("docNumber") ?? ""; // Номер договора на бумаге
                string RegNumbDoc = jsonManager.GetValue<string>("regNumbDoc") ?? ""; // Регистрационный номер в системе
                string Service = jsonManager.GetValue<string>("service.title") ?? ""; // Услуга
                string INN = jsonManager.GetValue<string>("innString") ?? ""; // ИНН 
                string KPP = jsonManager.GetValue<string>("kppString") ?? ""; // КПП

                // Даты: сначала пробуем как DateTime?, затем — как строку с парсингом
                DateTime? DateStart = jsonManager.GetValue<DateTime?>("startPeriod");
                if (!DateStart.HasValue)
                {
                    var s = jsonManager.GetValue<string>("startPeriod");
                    DateStart = ParseDateOrNull(s);
                }

                DateTime? DateEnd = jsonManager.GetValue<DateTime?>("endPeriod");
                if (!DateEnd.HasValue)
                {
                    var s = jsonManager.GetValue<string>("endPeriod");
                    DateEnd = ParseDateOrNull(s);
                }

                // Прочее, возможно пустое
                string status = jsonManager.GetValue<string>("status") ?? "";
                string message = jsonManager.GetValue<string>("message") ?? "";

                // Если пусто — нормализуем к null в JSON
                if (string.IsNullOrEmpty(status))
                    jsonManager.SetValue("status", null);

                if (string.IsNullOrEmpty(message))
                    jsonManager.SetValue("message", null);

                // Если у тебя в Request поля DateStart/DateEnd НЕ nullable (DateTime),
                // можно заменить на:
                // DateStart = DateStart ?? DateTime.MinValue;
                // DateEnd   = DateEnd   ?? DateTime.MinValue;

                return new Request
                {
                    UIID = UUID,
                    Type = Type,
                    Branch = Branch,
                    Organization = Organization,
                    DateStart = DateStart, // предполагаем DateTime? в модели Request
                    DateEnd = DateEnd,     // предполагаем DateTime? в модели Request
                    INN = INN,
                    KPP = KPP,
                    Service = Service,
                    DocNumber = docNumber,     // из docNumber
                    RegNumbDoc = RegNumbDoc, // из RegNumbDoc
                    status = status,
                    message = message
                };
            }
            catch (Exception ex)
            {
                Log.Error("Ошибка при извлечении данных из файла заявки: {0}", ex.Message);
                throw new Exception("Ошибка при извлечении данных из файла заявки: " + ex.Message, ex);
            }
        }

        // Поддержка частых форматов дат
        private static DateTime? ParseDateOrNull(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;

            var formats = new[]
            {
                "dd.MM.yyyy",
                "dd.MM.yyyy HH:mm:ss",
                "yyyy-MM-dd",
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-ddTHH:mm:ss",
                "yyyy-MM-ddTHH:mm:ss.fff",
                "yyyy-MM-ddTHH:mm:ssK",
                "yyyy-MM-ddTHH:mm:ss.fffK"
            };

            DateTime dt;
            if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out dt))
                return dt;

            // Fallback: парс с ru-RU и Invariant
            if (DateTime.TryParse(s, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.AssumeLocal, out dt)) return dt;
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out dt)) return dt;

            return null;
        }
    }
}
