using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Примитивный класс управления конфигурации в excel файле.
    /// </summary>
    public class ExcelConfigurationManager : IDisposable
    {
        public ExcelConfigurationManager(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException($"'{nameof(filePath)}' не может быть null или пустым.", nameof(filePath));
            }

            FilePath = filePath;
            ExcelHelper = new ExcelHelper(filePath);
        }

        /// <summary>
        /// Помощник по работе с excel файлами.
        /// </summary>
        public ExcelHelper ExcelHelper { get; }

        /// <summary>
        /// Путь к файлу.
        /// </summary>
        public string FilePath { get; }

        /// <summary>
        /// Словарь параметров " лист, словарь " параметр, значение " "
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> Params { get; set; } = new();

        /// <summary>
        /// Прочитать документ. Получить словарь параметров.
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, Dictionary<string, string>> Read()
        {
            Params.Clear();

            var worksheets = ExcelHelper.GetWorksheets();

            foreach (var worksheet in worksheets)
            {
                Dictionary<string, string> subParams = new();
                var leftCells = ExcelHelper.GetNotEmptyCellsToEndColumn(worksheet, "A", 1);
                foreach (var leftCell in leftCells)
                {
                    var rigthCell = ExcelHelper.GetOrCreateRightCell(worksheet, leftCell);
                    subParams.Add(leftCell.CellValue.Text, rigthCell.CellValue.Text);
                }
                string worksheetName = ExcelHelper.GetWorksheetName(worksheet);
                Params.Add(worksheetName, subParams);
            }

            // Это просто для корректировки, на случай, если там будут пустые (не созданные ячейки в правой части).
            // Но по сути это не обязательно.
            ExcelHelper.Save();

            return Params;
        }

        /// <summary>
        /// Сохранить параметры.
        /// </summary>
        public void Save()
        {
            foreach (var sheetParams in Params)
            {
                if (sheetParams.Value.Any())
                {
                    var worksheet = ExcelHelper.GetOrCreateWorksheet(sheetParams.Key);
                    foreach (var param in sheetParams.Value)
                    {
                        SaveOrCreateParam(worksheet, param);
                    }
                }
            }

            ExcelHelper.Save();
        }

        /// <summary>
        /// Сохранить или создать (записать новый) параметр и значение в лист.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="keyValue"></param>
        private void SaveOrCreateParam(Worksheet worksheet, KeyValuePair<string, string> keyValue)
        {
            var cell = ExcelHelper.FindCellByText(worksheet, keyValue.Key);
            if (cell != null)
            {
                var rigthCell = ExcelHelper.GetOrCreateRightCell(worksheet, cell);
                ExcelHelper.SetCellValue(rigthCell, keyValue.Value);
            }
            else
            {
                CreateParam(worksheet, keyValue);
            }
        }

        /// <summary>
        /// Создать (записать новый) параметр и значение в лист.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="keyValue"></param>
        private void CreateParam(Worksheet worksheet, KeyValuePair<string, string> keyValue)
        {
            var leftCells = ExcelHelper.GetNotEmptyCellsToEndColumn(worksheet, "A", 1);
            uint leftCellsCount = (uint)leftCells.Count();
            uint nextRowIndex = leftCellsCount + 1;
            var cell = ExcelHelper.GetOrCreateCell(worksheet, "A", nextRowIndex);
            ExcelHelper.SetCellValue(cell, keyValue.Key);
            cell = ExcelHelper.GetOrCreateCell(worksheet, "B", nextRowIndex);
            ExcelHelper.SetCellValue(cell, keyValue.Value);
        }

        public void Dispose()
        {
            ExcelHelper.Dispose();
        }
    }
}
