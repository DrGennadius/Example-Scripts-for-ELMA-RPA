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
        /// <summary>
        /// Максимальное количество столбцов.
        /// </summary>
        private const int MaxNumberOfColumns = 16384;

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
        /// Словарь единичных параметров &lt; строка (наименование листа в excel), словарь &lt;&lt; параметр, значение &gt;&gt;.<para/>
        /// Т.е. по одному ключу имеем одно значение.
        /// Пример единичного параметра:
        /// <code>
        /// {
        ///     "filePath" : "./app.config"
        /// }
        /// </code>
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> SingleParams { get; set; } = new();

        /// <summary>
        /// Словарь множественных параметров &lt; строка (наименование листа в excel), словарь &lt;&lt; параметр, массив значений &gt;&gt;.<para/>
        /// Т.е. по одному ключу имеем много значений - массив. В excel это значит, что используется более 2х столбцов.<para/>
        /// Пример единичного параметра:
        /// <code>
        /// {
        ///     "docFiles" : [ "./1.png", "./2.png", "./3.png", "./пример изображения.jpg", "./фотка.jpg", "./ух.png" ]
        /// }
        /// </code>
        /// </summary>
        public Dictionary<string, Dictionary<string, List<string>>> MultipleParams { get; set; } = new();

        /// <summary>
        /// Прочитать документ.
        /// </summary>
        /// <returns></returns>
        public void Read()
        {
            SingleParams.Clear();
            MultipleParams.Clear();

            var worksheets = ExcelHelper.GetWorksheets();

            foreach (var worksheet in worksheets)
            {
                Dictionary<string, string> singleSubParams = new();
                Dictionary<string, List<string>> multipleSubParams = new();
                var leftCells = ExcelHelper.GetNotEmptyCellsToEndColumn(worksheet, "A", 1);
                foreach (var leftCell in leftCells)
                {
                    Cell nextRigthCell = ExcelHelper.GetOrCreateRightCell(worksheet, leftCell);
                    List<Cell> nextRigthCells = new()
                    {
                        nextRigthCell
                    };
                    int columnIndex = 2;
                    while (columnIndex < MaxNumberOfColumns)
                    {
                        nextRigthCell = ExcelHelper.GetRightCell(worksheet, nextRigthCell);
                        if (nextRigthCell != null)
                        {
                            nextRigthCells.Add(nextRigthCell);
                        }
                        else
                        {
                            break;
                        }
                        columnIndex++;
                    }
                    switch (nextRigthCells.Count)
                    {
                        case 0:
                            singleSubParams.Add(ExcelHelper.GetCellStringValue(leftCell), null);
                            break;
                        case 1:
                            singleSubParams.Add(ExcelHelper.GetCellStringValue(leftCell), ExcelHelper.GetCellStringValue(nextRigthCells[0]));
                            break;
                        default:
                            multipleSubParams.Add(ExcelHelper.GetCellStringValue(leftCell), nextRigthCells.Select(x => ExcelHelper.GetCellStringValue(x)).ToList());
                            break;
                    }
                }
                string worksheetName = ExcelHelper.GetWorksheetName(worksheet);
                if (singleSubParams.Any())
                {
                    SingleParams.Add(worksheetName, singleSubParams);
                }
                if (multipleSubParams.Any())
                {
                    MultipleParams.Add(worksheetName, multipleSubParams);
                }
            }

            // Это просто для корректировки, на случай, если там будут пустые (не созданные ячейки в правой части).
            // Но по сути это не обязательно.
            ExcelHelper.Save();
        }

        /// <summary>
        /// Сохранить параметры.
        /// </summary>
        public void Save()
        {
            // Единичные параметры.
            foreach (var sheetParams in SingleParams)
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

            // Множественные параметры.
            foreach (var sheetParams in MultipleParams)
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

            // Сохраняем файл.
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
        /// Сохранить или создать (записать новый) параметр и массив значений в лист.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="keyValue"></param>
        private void SaveOrCreateParam(Worksheet worksheet, KeyValuePair<string, List<string>> keyValue)
        {
            var cell = ExcelHelper.FindCellByText(worksheet, keyValue.Key);
            if (cell != null)
            {
                foreach (var cellValue in keyValue.Value)
                {
                    cell = ExcelHelper.GetOrCreateRightCell(worksheet, cell);
                    ExcelHelper.SetCellValue(cell, cellValue);
                }
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

        /// <summary>
        /// Создать (записать новый) параметр и массив значений в лист.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="keyValue"></param>
        private void CreateParam(Worksheet worksheet, KeyValuePair<string, List<string>> keyValue)
        {
            var leftCells = ExcelHelper.GetNotEmptyCellsToEndColumn(worksheet, "A", 1);
            uint leftCellsCount = (uint)leftCells.Count();
            uint nextRowIndex = leftCellsCount + 1;
            Cell cell = ExcelHelper.GetOrCreateCell(worksheet, "A", nextRowIndex);
            ExcelHelper.SetCellValue(cell, keyValue.Key);
            foreach (var cellValue in keyValue.Value)
            {
                cell = ExcelHelper.GetOrCreateRightCell(worksheet, cell);
                ExcelHelper.SetCellValue(cell, cellValue);
            }
        }

        public void Dispose()
        {
            ExcelHelper.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}
