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
            // Для начала можно ознакомиться с примером для извлечения одной таблицы.

            if (!string.IsNullOrWhiteSpace(context.FilePath) 
                && File.Exists(context.FilePath))
            {
                HtmlDocument doc = new();
                doc.Load(context.FilePath);

                List <DataTable> tables1 = GetMultiplyDataTables(doc);
                List <string[][]> table2 = GetMultiplyDataTablesAsStringArrays(doc);
                List <string> table3 = GetMultiplyDataTablesAsJsonArray(doc);
                string table4 = GetMultiplyDataTablesAsSingleJson(doc);
            }
        }

        /// <summary>
        /// Получить данные таблиц списком экземпляров типа <see cref="DataTable"/>.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private List<DataTable> GetMultiplyDataTables(HtmlDocument document)
        {
            List<DataTable> tables = new();

            if (document == null)
            {
                return tables;
            }

            // Получаем ноды элементов таблиц
            var tableNodes = document.DocumentNode.SelectNodes("//table");

            if (tableNodes.Count == 0)
            {
                return tables;
            }

            // Номер таблицы для наименования.
            int tableNumber = 0;
            foreach (var tableNode in tableNodes)
            {
                tableNumber++;
                DataTable table = GetDataTable(tableNode, $"table {tableNumber}");
                tables.Add(table);
            }

            return tables;
        }

        /// <summary>
        /// Получить данные таблиц списком массива массивов строк.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private List<string[][]> GetMultiplyDataTablesAsStringArrays(HtmlDocument document)
        {
            // Всё почти тоже самое, что и GetMultiplyDataTables,
            // только обработка ноды таблицы другим методом и возвращаемый объект друго типа.
            List<string[][]> tables = new();

            if (document == null)
            {
                return tables;
            }

            var tableNodes = document.DocumentNode.SelectNodes("//table");

            if (tableNodes.Count == 0)
            {
                return tables;
            }

            foreach (var tableNode in tableNodes)
            {
                string[][] table = GetDataTableAsStringArrays(tableNode);
                tables.Add(table);
            }

            return tables;
        }

        /// <summary>
        /// Получить данные таблиц списком строк формата JSON.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private List<string> GetMultiplyDataTablesAsJsonArray(HtmlDocument document)
        {
            List<string> tables = new();

            if (document == null)
            {
                return tables;
            }

            var tableNodes = document.DocumentNode.SelectNodes("//table");

            if (tableNodes.Count == 0)
            {
                return tables;
            }

            foreach (var tableNode in tableNodes)
            {
                string table = GetDataTableAsJson(tableNode);
                tables.Add(table);
            }

            return tables;
        }

        /// <summary>
        /// Получить данные таблиц в виде одной строки формата JSON.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private string GetMultiplyDataTablesAsSingleJson(HtmlDocument document)
        {
            List<string[][]> tables = GetMultiplyDataTablesAsStringArrays(document);

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic)
            };
            return JsonSerializer.Serialize(tables, options);
        }

        /// <summary>
        /// Получить данные таблицы экземпляром типа <see cref="DataTable"/>.
        /// </summary>
        /// <param name="tableNode"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private DataTable GetDataTable(HtmlNode tableNode, string tableName = "table")
        {
            // Этот метод является аналогом метода GetDataTable из примера, где извлекаем 1 таблицу.
            // Сейчас вместо HtmlDocument на входе HtmlNode.
            DataTable table = new(tableName);

            if (tableNode == null)
            {
                return table;
            }

            // Получаем ноды элементов tr в ноде таблицы.
            var nodes = tableNode.Elements("tr").ToArray();

            if (nodes.Length == 0)
            {
                return table;
            }

            var headers = nodes[0]
                .Elements("th")
                .Select(th => th.InnerText.Trim());
            foreach (var header in headers)
            {
                table.Columns.Add(header);
            }

            var rows = nodes
                .Skip(1)
                .Select(tr => tr
                    .Elements("td")
                    .Select(td => td.InnerText.Trim())
                .ToArray());
            foreach (var row in rows)
            {
                table.Rows.Add(row);
            }

            return table;
        }

        /// <summary>
        /// Получить данные таблицы в виде массива массивов строк.
        /// </summary>
        /// <param name="tableNode"></param>
        /// <returns></returns>
        private string[][] GetDataTableAsStringArrays(HtmlNode tableNode)
        {
            if (tableNode == null)
            {
                return Array.Empty<string[]>();
            }

            var nodes = tableNode.Elements("tr").ToArray();

            if (nodes.Length == 0)
            {
                return Array.Empty<string[]>();
            }

            string[][] table = new string[nodes.Length][];

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
                table[i + 1] = rows[i];
            }

            return table;
        }

        /// <summary>
        /// Получить данные таблицы в виде строки формата JSON.
        /// </summary>
        /// <param name="tableNode"></param>
        /// <returns></returns>
        private string GetDataTableAsJson(HtmlNode tableNode)
        {
            // Т.к. у нас уже есть метод получения данных в примитивном виде
            // массива массивов строк, то по сути нам остается только сериализовать (конвертировать в json).

            string[][] table = GetDataTableAsStringArrays(tableNode);

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic)
            };
            return JsonSerializer.Serialize(table, options);
        }
    }
}
