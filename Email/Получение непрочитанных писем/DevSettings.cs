using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Класс настроек. Для отладки скрипта.
    /// </summary>
    internal class DevSettings
    {
        public DevSettings()
        {
        }

        public DevSettings(string configFilePath)
        {
            Read(configFilePath);
        }

        public DevSettings(DevSettings settings)
        {
            SetSettings(settings);
        }

        /// <summary>
        /// Хост сервера
        /// </summary>
        public string Host { get; set; }

        /// <summary>
        /// Порт
        /// </summary>
        public int Port { get; set; }

        /// <summary>
        /// Использовать SSL
        /// </summary>
        public bool UseSsl { get; set; }

        /// <summary>
        /// Имя пользователя
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// Пароль
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Тестовый текст для поиска в тексте темы
        /// </summary>
        public string SubjectText { get; set; }

        /// <summary>
        /// Тестовый текст для поиска отправителя
        /// </summary>
        public string SenderText { get; set; }

        /// <summary>
        /// Чтение/получение настроек из файла по указанному пути
        /// </summary>
        /// <param name="configFilePath"></param>
        public void Read(string configFilePath)
        {
            if (File.Exists(configFilePath))
            {
                var options = new JsonSerializerOptions
                {
                    // Использование "верблюжьего" стиля для всех имен свойств.
                    // Это выглядит так: camelCase, workFolder, fileName
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    // Расширенный (не в одну строку)
                    WriteIndented = true,
                    // Кодировка для Unicode: Basic Latin и Cyrillic
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic)
                };
                // Чтение json закодированной в utf8 байты
                byte[] jsonUtf8Bytes = File.ReadAllBytes(configFilePath);
                // Десериализация
                var settings = JsonSerializer.Deserialize<DevSettings>(jsonUtf8Bytes, options);
                SetSettings(settings);
            }
            else
            {
                Host = "imap.gmail.com";
                Port = 993;
                UseSsl = true;
                // Записываем настройки в файл
                Write(configFilePath);
            }
        }

        /// <summary>
        /// Установка настроек
        /// </summary>
        /// <param name="settings"></param>
        public void SetSettings(DevSettings settings)
        {
            Host = settings.Host;
            Port = settings.Port;
            UseSsl = settings.UseSsl;
            UserName = settings.UserName;
            Password = settings.Password;
            SubjectText = settings.SubjectText;
            SenderText = settings.SenderText;
        }

        /// <summary>
        /// Запись настроек в файл по указанному пути
        /// </summary>
        /// <param name="configFilePath"></param>
        public void Write(string configFilePath)
        {
            var options = new JsonSerializerOptions
            {
                // Использование "верблюжьего" стиля для всех имен свойств.
                // Это выглядит так: camelCase, workFolder, fileName
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                // Расширенный (не в одну строку)
                WriteIndented = true,
                // Кодировка для Unicode: Basic Latin и Cyrillic
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic)
            };
            // Получение строки json закодированной в utf8 байты
            byte[] jsonUtf8Bytes = JsonSerializer.SerializeToUtf8Bytes(this, options);
            // Запись в файл
            File.WriteAllBytes(configFilePath, jsonUtf8Bytes);
        }
    }
}
