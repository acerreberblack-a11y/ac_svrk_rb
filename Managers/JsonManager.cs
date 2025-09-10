using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.IO;

namespace SpravkoBot_AsSapfir
{
    internal class JsonManager
    {
        private readonly JObject _jsonObject;
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        public JsonManager(string filePath, bool isFile = true)
        {
            try
            {
                if (isFile)
                {
                    if (!File.Exists(filePath))
                    {
                        Log.Error("Файл не найден: {0}", filePath);
                        throw new FileNotFoundException("Файл не найден: " + filePath);
                    }

                    var json = File.ReadAllText(filePath);
                    _jsonObject = JObject.Parse(json, new JsonLoadSettings { CommentHandling = CommentHandling.Ignore });
                }
                else
                {
                    _jsonObject = JObject.Parse(filePath, new JsonLoadSettings { CommentHandling = CommentHandling.Ignore });
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при загрузке JSON");
                throw new JsonException("Ошибка при загрузке JSON", ex);
            }
        }

        private static string NormalizePath(string key)
        {
            return key == null ? null : key.Replace(':', '.');
        }

        /// <summary>
        /// Получить значение по ключу-пути. Если не найдено — вернёт default(T) и Warning в лог.
        /// </summary>
        public T GetValue<T>(string key)
        {
            var norm = NormalizePath(key);
            try
            {
                var token = _jsonObject.SelectToken(norm);
                if (token != null)
                    return token.ToObject<T>();

                Log.Warn("Ключ '{0}' не найден.", key);
                return default(T);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при получении значения по ключу '{0}'", key);
                throw new JsonException("Ошибка при получении значения по ключу '" + key + "'", ex);
            }
        }

        /// <summary>
        /// Пытается получить значение; без исключений при «не найдено».
        /// </summary>
        public bool TryGetValue<T>(string key, out T value, bool treatEmptyStringAsMissing = true)
        {
            value = default(T);
            var norm = NormalizePath(key);
            try
            {
                var token = _jsonObject.SelectToken(norm);
                if (token == null)
                    return false;

                var v = token.ToObject<T>();

                if (treatEmptyStringAsMissing && v is string)
                {
                    var s = v as string;
                    if (string.IsNullOrWhiteSpace(s))
                        return false;
                }

                value = v;
                return true;
            }
            catch (Exception ex)
            {
                // Это уже действительно ошибка (битый JSON и т.п.)
                Log.Error(ex, "TryGetValue: ошибка при доступе к ключу '{0}'", key);
                return false;
            }
        }

        /// <summary>
        /// Установить значение по пути, создавая недостающие узлы.
        /// Поддерживает только объектные пути вида a.b.c (без массивов).
        /// </summary>
        public void SetValue(string key, object value)
        {
            var norm = NormalizePath(key);
            try
            {
                var token = _jsonObject.SelectToken(norm);
                if (token != null)
                {
                    token.Replace(JToken.FromObject(value));
                    return;
                }

                // Создаём путь
                var parts = norm.Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 0)
                    throw new ArgumentException("Некорректный ключ", nameof(key));

                JToken current = _jsonObject;

                for (int i = 0; i < parts.Length; i++)
                {
                    var part = parts[i];

                    if (i == parts.Length - 1)
                    {
                        // Последний сегмент — устанавливаем значение
                        var obj = current as JObject;
                        if (obj == null)
                            throw new InvalidOperationException("Ожидался объект в пути для установки значения.");
                        obj[part] = JToken.FromObject(value);
                    }
                    else
                    {
                        var obj = current as JObject;
                        if (obj == null)
                            throw new InvalidOperationException("Ожидался объект в пути.");

                        if (obj[part] == null || obj[part].Type != JTokenType.Object)
                            obj[part] = new JObject();

                        current = obj[part];
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при записи значения по ключу '{0}'", key);
                throw new JsonException("Ошибка при записи значения по ключу '" + key + "'", ex);
            }
        }

        /// <summary>
        /// Добавляет ключ (как путь). Эквивалент SetValue, но оставлен для совместимости.
        /// </summary>
        public void AddKey(string key, object value)
        {
            // Для корректной вложенности используем SetValue
            SetValue(key, value);
        }

        /// <summary>
        /// Удаляет ключ/узел по пути. Если не найден — Warning и выход.
        /// </summary>
        public void RemoveKey(string key)
        {
            var norm = NormalizePath(key);
            try
            {
                var token = _jsonObject.SelectToken(norm);
                if (token != null && token.Parent != null)
                {
                    token.Parent.Remove();
                }
                else
                {
                Log.Warn("Ключ '{0}' не найден для удаления.", key);
                    // Без исключения — это допустимый сценарий.
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при удалении ключа '{0}'", key);
                throw new JsonException("Ошибка при удалении ключа '" + key + "'", ex);
            }
        }

        public void SaveToFile(string filePath)
        {
            try
            {
                string jsonString = _jsonObject.ToString();
                File.WriteAllText(filePath, jsonString);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при сохранении в файл");
                throw new JsonException("Ошибка при сохранении в файл", ex);
            }
        }

        public string ToJsonString()
        {
            return _jsonObject.ToString();
        }
    }
}
