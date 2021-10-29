using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Encodings.Web;
using System.Text.Unicode;
using System.Threading;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Экземпляр данного класса будет создан при выполнении скрипта.
    /// <summary>
    public class ScriptActivity
    {
        /// <summary>
        /// Данная функция является точкой входа.
        /// <summary>
        public void Execute(Context context)
        {
            // Получаем настройки
            ScriptSettings settings = ReadSettings(context.ConfigFilePath);
            // Меняем настройки
            settings.LastReadConfigDateTime = DateTime.Now;
            // Сохраняем настройки
            WriteSettings(context.ConfigFilePath, settings);

            // Сохраним нужные настройки в контекст робота,
            // чтобы потом можно было использовать это в процессе выполнения робота
            context.WorkFolder = settings.WorkFolder;
            context.FileNames = settings.FileNames;

            // Далее продолжим выполнять какие-то действия в этом сценарии,
            // как будто бы что-то нужное делаем с полученными данными
            DoSomething(context, settings);
        }

        /// <summary>
        /// Получение настроек из файла по указанному пути
        /// </summary>
        /// <param name="configFilePath"></param>
        /// <returns>Настройки</returns>
        private ScriptSettings ReadSettings(string configFilePath)
        {
            ScriptSettings settings = null;
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
                // Десериализации
                settings = JsonSerializer.Deserialize<ScriptSettings>(jsonUtf8Bytes, options);
            }
            else
            {
                // Если файл отсутствует, то генерируем для примера новый экземпляр настроек
                // и запишем туда какие-то "стандартные" значения
                string directory = Path.GetDirectoryName(configFilePath);
                settings = new ScriptSettings()
                {
                    LastReadConfigDateTime = DateTime.Now,
                    WaitTime = 1000,
                    WorkFolder = Path.Combine(directory, "Рабочая папка"),
                    FileNames = new List<string>
                    {
                        "Файл 1.txt",
                        "Файл 2.txt",
                        "Файл 3.txt"
                    }
                };
                // Записываем настройки в файл
                WriteSettings(configFilePath, settings);
            }
            return settings;
        }

        /// <summary>
        /// Запись настроек в файл по указанному пути
        /// </summary>
        /// <param name="configFilePath"></param>
        /// <param name="settings"></param>
        private void WriteSettings(string configFilePath, ScriptSettings settings)
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
            byte[] jsonUtf8Bytes = JsonSerializer.SerializeToUtf8Bytes(settings, options);
            // Запись в файл
            File.WriteAllBytes(configFilePath, jsonUtf8Bytes);
        }

        /// <summary>
        /// Какой-то метод, который что-то делает с полученными данными из настроек
        /// </summary>
        /// <param name="context"></param>
        /// <param name="scriptSettings"></param>
        private void DoSomething(Context context, ScriptSettings scriptSettings)
        {
            // В качестве примера пускай просто записывает текущее время
            // в файлы, используя имена файлов, рабочую папку и время для пауз
            if (context.FileNames == null || string.IsNullOrEmpty(context.WorkFolder) || context.FileNames.Count == 0)
            {
                return;
            }
            Directory.CreateDirectory(context.WorkFolder);
            foreach (var fileName in context.FileNames)
            {
                string text = "Текущее время: " + DateTime.Now.ToString();
                string filePath = Path.Combine(context.WorkFolder, fileName);
                byte[] bytes = Encoding.Default.GetBytes(text);
                string outputText = Encoding.UTF8.GetString(bytes);

                File.WriteAllText(filePath, outputText, Encoding.UTF8);

                Thread.Sleep(scriptSettings.WaitTime);
            }
        }
    }

    /// <summary>
    /// Класс настроек. Можно вынести в отдельный файл, 
    /// но т.к. кода мало, то для примера в этом же файле
    /// </summary>
    public class ScriptSettings
    {
        /// <summary>
        /// Некая рабочая папка, которую может использовать робот (не только данный скрипт),
        /// если сохранить ее в контекст робота.
        /// </summary>
        public string WorkFolder { get; set; }

        /// <summary>
        /// Время ожидания между обработками файлов, пусть в качестве примера будет.
        /// </summary>
        public int WaitTime { get; set; }

        /// <summary>
        /// Список файлов. Будем в качестве примера использовать в скрипте,
        /// а также сохранять в контекст робота.
        /// </summary>
        public List<string> FileNames { get; set; }

        /// <summary>
        /// Время чтения файла настроек. Представлено опять же исключительно
        /// в качестве примеря, будем записывать обратно в файл после чтения.
        /// </summary>
        public DateTime LastReadConfigDateTime { get; set; }
    }
}
