using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using HtmlAgilityPack;
using System.Text.Json;
using System.Text.Encodings.Web;
using System.Text.Unicode;

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
            if (!string.IsNullOrWhiteSpace(context.FilePath) 
                && File.Exists(context.FilePath))
            {
                HtmlDocument doc = new();
                doc.Load(context.FilePath);

                DataTable table1 = GetDataTable(doc, "Пример таблицы 1");
                string[][] table2 = GetDataTableAsStringArrays(doc);
                string table3 = GetDataTableAsJson(doc);
            }
        }

        /// <summary>
        /// Получить данные таблицы экземпляром типа <see cref="DataTable"/>.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private DataTable GetDataTable(HtmlDocument document, string tableName = "table")
        {
            // Создаем экземпляр класса DataTable.
            DataTable table = new(tableName);

            if (document == null)
            {
                // Если документ = null, то возвращаем table без данных строк и столбцов.
                // Как вариант можно вызывать эксепшин или возвращать null, если очень хочется.
                return table;
            }

            // Получаем ноды элементов tr в table, т.е. это прям строки которые.
            var nodes = document.DocumentNode.SelectNodes("//table/tr");

            if (nodes.Count == 0)
            {
                // Аналогично, что и в случае с документ = null.
                return table;
            }

            // Получаем значения хидера, т.е. заголовки, т.е. тексты элементов th.
            var headers = nodes[0]
                .Elements("th")
                .Select(th => th.InnerText.Trim());
            foreach (var header in headers)
            {
                // Добавляем столбцы.
                table.Columns.Add(header);
            }

            // Получаем значения ячеек тела таблицы, т.е. тексты элементов td.
            var rows = nodes
                .Skip(1)
                .Select(tr => tr
                    .Elements("td")
                    .Select(td => td.InnerText.Trim())
                .ToArray());
            foreach (var row in rows)
            {
                // Добавляем строки.
                table.Rows.Add(row);
            }

            return table;
        }

        /// <summary>
        /// Получить данные таблицы в виде массива массивов строк.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private string[][] GetDataTableAsStringArrays(HtmlDocument document)
        {
            // Почти как GetDataTable, поэтому тут не буду дублировать те же комментарии.

            if (document == null)
            {
                return Array.Empty<string[]>();
            }

            var nodes = document.DocumentNode.SelectNodes("//table/tr");

            if (nodes.Count == 0)
            {
                return Array.Empty<string[]>();
            }

            string[][] table = new string[nodes.Count][];

            // Тут сразу получаем массив и присваеваем в первый элемент table.
            var headers = nodes[0]
                .Elements("th")
                .Select(th => th.InnerText.Trim())
                .ToArray();
            table[0] = headers;

            var rows = nodes
                .Skip(1)
                .Select(tr => tr
                    .Elements("td")
                    .Select(td => td.InnerText.Trim())
                    .ToArray())
                .ToArray();

            for (int i = 0; i < rows.Length; i++)
            {
                // + 1, т.к. первый массив это хидер
                table[i + 1] = rows[i];
            }

            return table;
        }

        /// <summary>
        /// Получить данные таблицы в виде строки формата JSON.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private string GetDataTableAsJson(HtmlDocument document)
        {
            // Т.к. у нас уже есть метод получения данных в примитивном виде
            // массива массивов строк, то по сути нам остается только сериализовать (конвертировать в json).

            string[][] table = GetDataTableAsStringArrays(document);

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
            // Получение строки json
            return JsonSerializer.Serialize(table, options);
        }
    }
}
