using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Простой помощник по работе с excel файлам.
    /// </summary>
    public class ExcelHelper : IDisposable
    {
        public ExcelHelper(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException($"'{nameof(filePath)}' не может быть null или пустым.", nameof(filePath));
            }

            OpenOrCreate(filePath);
        }

        /// <summary>
        /// Документ.
        /// </summary>
        public SpreadsheetDocument Document { get; private set; }

        /// <summary>
        /// Книга документа.
        /// </summary>
        public Workbook Workbook { get; private set; }

        /// <summary>
        /// Путь к файлу.
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// Открыть (создать в случае отсутсвия) файл, указав путь к файлу.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public SpreadsheetDocument OpenOrCreate(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException($"'{nameof(filePath)}' не может быть null или пустым.", nameof(filePath));
            }

            FilePath = filePath;
            return OpenOrCreate();
        }

        /// <summary>
        /// Открыть (создать в случае отсутствия) файл.
        /// </summary>
        /// <returns></returns>
        public SpreadsheetDocument OpenOrCreate()
        {
            if (string.IsNullOrEmpty(FilePath))
            {
                throw new ArgumentException($"'{nameof(FilePath)}' не может быть null или пустым.", nameof(FilePath));
            }

            bool excelExists = File.Exists(FilePath);
            Document = excelExists
                ? SpreadsheetDocument.Open(FilePath, true)
                : SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook);
            
            // Еще в любом случае проверяем workbook сразу и создаем.
            WorkbookPart workbookPart = Document.WorkbookPart;
            if (workbookPart == null)
            {
                workbookPart = Document.AddWorkbookPart();
            }
            Workbook workbook = workbookPart.Workbook;
            if (workbook == null)
            {
                workbookPart.Workbook = new Workbook();
            }
            Workbook = workbookPart.Workbook;

            return Document;
        }

        /// <summary>
        /// Получить список листов.
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Worksheet> GetWorksheets()
        {
            if (Workbook is null)
            {
                throw new ArgumentNullException(nameof(Workbook));
            }

            return Workbook.WorkbookPart.GetPartsOfType<WorksheetPart>().Select(x => x.Worksheet);
        }

        /// <summary>
        /// Получить или создать лист с наименованием.
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        public Worksheet GetOrCreateWorksheet(string worksheetName)
        {
            if (string.IsNullOrEmpty(worksheetName))
            {
                throw new ArgumentException($"'{nameof(worksheetName)}' не может быть null или пустым.", nameof(worksheetName));
            }

            WorksheetPart worksheetPart = null;
            Sheet sheet = Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
            if (sheet == null)
            {
                worksheetPart = Workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                uint newSheetId = 1U;
                Sheets sheets = Workbook.GetFirstChild<Sheets>();
                if (sheets != null)
                {
                    // Нужно задать верный id для нового листа, после последнего
                    Sheet lastSheet = sheets.Elements<Sheet>().OrderBy(x => x.SheetId).LastOrDefault();
                    if (lastSheet != null)
                    {
                        newSheetId = lastSheet.SheetId + 1;
                    }
                }
                else
                {
                    // Если элемента с таким типом еще нет, то обязательно нужно создать:
                    sheets = Workbook.AppendChild(new Sheets());
                }

                sheet = new Sheet()
                {
                    Id = Workbook.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = newSheetId,
                    Name = worksheetName
                };
                sheets.Append(sheet);
            }
            else
            {
                worksheetPart = Workbook.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
            }

            return worksheetPart.Worksheet;
        }

        /// <summary>
        /// Получить наименование листа.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public string GetWorksheetName(Worksheet worksheet)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            string worksheetName = null;
            string relationshipId = Workbook.WorkbookPart.GetIdOfPart(worksheet.WorksheetPart);
            Sheets sheets = Workbook.GetFirstChild<Sheets>();
            Sheet sheet = sheets.Elements<Sheet>().FirstOrDefault(x => x.Id == relationshipId);
            if (sheet != null)
            {
                worksheetName = sheet.Name;
            }
            return worksheetName;
        }

        /// <summary>
        /// Получить строку.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            return sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        }

        /// <summary>
        /// Создать строку.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public Row CreateRow(Worksheet worksheet, uint rowIndex)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            Row nextRow = sheetData.Elements<Row>().OrderBy(r => r.RowIndex).FirstOrDefault(r => r.RowIndex > rowIndex);
            Row row = new Row()
            {
                RowIndex = rowIndex
            };
            sheetData.InsertBefore(row, nextRow);

            return row;
        }

        /// <summary>
        /// Получить или создать строку.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public Row GetOrCreateRow(Worksheet worksheet, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
            {
                row = CreateRow(worksheet, rowIndex);
            }

            return row;
        }
        
        /// <summary>
        /// Получить текстового значение ячейки.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public string GetCellStringValue(Cell cell)
        {
            if (cell is null)
            {
                return "";
            }

            string result = cell.CellValue.Text;
            if (cell.DataType.HasValue && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = Workbook.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                result = stringTable.SharedStringTable.ElementAt(int.Parse(result)).InnerText;
            }
            return result;
        }

        #region Get cell

        /// <summary>
        /// Получить ячейку.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="columnName"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public Cell GetCell(Worksheet worksheet, string columnName, Row row)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentException($"'{nameof(columnName)}' не может быть null или пустым.", nameof(columnName));
            }

            if (row is null)
            {
                throw new ArgumentNullException(nameof(row));
            }

            string cellReference = columnName + row.RowIndex;
            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);
        }

        /// <summary>
        /// Получить ячейку.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentException($"'{nameof(columnName)}' не может быть null или пустым.", nameof(columnName));
            }

            Cell cell = null;

            Row row = GetRow(worksheet, rowIndex);

            if (row != null)
            {
                cell = GetCell(worksheet, columnName, row);
            }

            return cell;
        }

        /// <summary>
        /// Получить ячейку.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellAddress"></param>
        /// <returns></returns>
        public Cell GetCell(Worksheet worksheet, string cellAddress)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (string.IsNullOrEmpty(cellAddress))
            {
                throw new ArgumentException($"'{nameof(cellAddress)}' не может быть null или пустым.", nameof(cellAddress));
            }

            string columnName = Regex.Match(cellAddress.ToUpper(), @"[A-Z]+").Value;
            uint rowIndex = uint.Parse(Regex.Match(cellAddress, @"[0-9]+").Value);
            return GetCell(worksheet, columnName, rowIndex);
        }

        /// <summary>
        /// Получить не пустые ячейки до конца столбца.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public IEnumerable<Cell> GetNotEmptyCellsToEndColumn(Worksheet worksheet, string columnName, uint rowIndex)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentException($"'{nameof(columnName)}' не может быть null или пустым.", nameof(columnName));
            }

            List<Cell> cells = new List<Cell>();
            Cell cell = null;
            do
            {
                cell = GetCell(worksheet, columnName, rowIndex);
                if (cell != null && (cell.CellValue == null || string.IsNullOrWhiteSpace(cell.CellValue.Text)))
                {
                    cell = null;
                }
                if (cell != null)
                {
                    cells.Add(cell);
                }
                rowIndex++;
            } while (cell != null);

            return cells;
        }

        #endregion

        #region Create cell

        /// <summary>
        /// Создать ячейку.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="columnName"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public Cell CreateCell(Worksheet worksheet, string columnName, Row row)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentException($"'{nameof(columnName)}' не может быть null или пустым.", nameof(columnName));
            }

            if (row is null)
            {
                throw new ArgumentNullException(nameof(row));
            }

            string cellReference = columnName + row.RowIndex;
            Cell refCell = null;
            foreach (Cell celli in row.Elements<Cell>())
            {
                if (celli.CellReference.Value.Length == cellReference.Length)
                {
                    if (string.Compare(celli.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = celli;
                        break;
                    }
                }
            }

            Cell cell = new Cell()
            {
                CellReference = cellReference
            };
            row.InsertBefore(cell, refCell);

            return cell;
        }

        #endregion

        #region Get or create cell

        /// <summary>
        /// Получить и создать ячейку.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="columnName"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public Cell GetOrCreateCell(Worksheet worksheet, string columnName, Row row)
        {
            Cell cell = GetCell(worksheet, columnName, row);

            if (cell == null)
            {
                cell = CreateCell(worksheet, columnName, row);
            }

            return cell;
        }

        /// <summary>
        /// Получить или создать ячейку
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public Cell GetOrCreateCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentException($"'{nameof(columnName)}' не может быть null или пустым.", nameof(columnName));
            }

            Row row = GetOrCreateRow(worksheet, rowIndex);
            return GetOrCreateCell(worksheet, columnName, row);
        }

        /// <summary>
        /// Получить или создать ячейку по адресу (например, B2).
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellAddress"></param>
        /// <returns></returns>
        public Cell GetOrCreateCell(Worksheet worksheet, string cellAddress)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (string.IsNullOrEmpty(cellAddress))
            {
                throw new ArgumentException($"'{nameof(cellAddress)}' не может быть null или пустым.", nameof(cellAddress));
            }

            string columnName = Regex.Match(cellAddress.ToUpper(), @"[A-Z]+").Value;
            uint rowIndex = uint.Parse(Regex.Match(cellAddress, @"[0-9]+").Value);
            return GetOrCreateCell(worksheet, columnName, rowIndex);
        }

        /// <summary>
        /// Получить или создать ячейку правее указанного местоположения.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public Cell GetOrCreateRightCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentException($"'{nameof(columnName)}' не может быть null или пустым.", nameof(columnName));
            }

            string rigthCellColumnName = IncrementColumnAddress(columnName);
            return GetOrCreateCell(worksheet, rigthCellColumnName, rowIndex);
        }

        /// <summary>
        /// Получить или создать ячейку правее указанного адреса ячейки.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellAddress"></param>
        /// <returns></returns>
        public Cell GetOrCreateRightCell(Worksheet worksheet, string cellAddress)
        {
            if (string.IsNullOrWhiteSpace(cellAddress))
            {
                throw new ArgumentException($"'{nameof(cellAddress)}' не может быть null или пустым.", nameof(cellAddress));
            }

            string columnName = Regex.Match(cellAddress.ToUpper(), @"[A-Z]+").Value;
            uint rowIndex = uint.Parse(Regex.Match(cellAddress, @"[0-9]+").Value);
            return GetOrCreateRightCell(worksheet, columnName, rowIndex);
        }

        /// <summary>
        /// Получить или создать ячейку правее указанной ячейки.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public Cell GetOrCreateRightCell(Worksheet worksheet, Cell cell)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            string currentCellAddress = cell.CellReference;
            string columnName = Regex.Match(currentCellAddress.ToUpper(), @"[A-Z]+").Value;
            uint rowIndex = uint.Parse(Regex.Match(currentCellAddress, @"[0-9]+").Value);
            string rigthCellColumnName = IncrementColumnAddress(columnName);
            return GetOrCreateCell(worksheet, rigthCellColumnName, rowIndex);
        }

        #endregion

        /// <summary>
        /// Найти ячейку по тексту.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="searchText"></param>
        /// <returns></returns>
        public Cell FindCellByText(Worksheet worksheet, string searchText)
        {
            if (worksheet is null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }

            if (string.IsNullOrEmpty(searchText))
            {
                throw new ArgumentException($"'{nameof(searchText)}' не может быть null или пустым.", nameof(searchText));
            }

            Cell cell = null;

            var rows = worksheet.Descendants<Row>().ToList();

            foreach (var row in rows)
            {
                var cells = row.Elements<Cell>().ToList();

                foreach (var celli in cells)
                {
                    if(GetCellStringValue(celli) == searchText)
                    {
                        cell = celli;
                        break;
                    }
                }
            }

            return cell;
        }

        #region Set cell value

        /// <summary>
        /// Установить текстовое значение в ячейку.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, string text, CellValues dataType = CellValues.String)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            cell.CellValue = new CellValue(text);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Установить значение даты/время в ячейку.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateTime"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, DateTime dateTime, CellValues dataType = CellValues.Date)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            cell.CellValue = new CellValue(dateTime);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Установить значение момента времени в ячейку.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateTimeOffset"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, DateTimeOffset dateTimeOffset, CellValues dataType = CellValues.Date)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            cell.CellValue = new CellValue(dateTimeOffset);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Установить булевое (да/нет, 1/0) значение в ячейку.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, bool value, CellValues dataType = CellValues.Boolean)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Установить число двойной точности с плавающей запятой (10.5) в ячейку.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, double value, CellValues dataType = CellValues.Number)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Установить целочисленное (10) значение в ячейку.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, int value, CellValues dataType = CellValues.Number)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Установить значение десятичного числа с плавающей запятой в ячейку.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, decimal value, CellValues dataType = CellValues.Number)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        #endregion

        /// <summary>
        /// Инкремент (увеличение на 1) адреса столбца (A -> B, Z -> AA).
        /// </summary>
        /// <param name="columnAddress"></param>
        /// <returns></returns>
        public static string IncrementColumnAddress(string columnAddress)
        {
            if (string.IsNullOrWhiteSpace(columnAddress))
            {
                throw new ArgumentException($"'{nameof(columnAddress)}' не может быть null или состоять из пробелов.", nameof(columnAddress));
            }

            string normalizedAddres = Regex.Match(columnAddress.ToUpper(), @"[A-Z]+").Value;
            int normalizedNumber = ColumnNameToNumber(normalizedAddres);
            return ColumnNumberToName(++normalizedNumber);
        }

        /// <summary>
        /// Адрес столбца в номер.
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private static int ColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentException($"'{nameof(columnName)}' не может быть null или состоять из пробелов.", nameof(columnName));
            }

            int result = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                result *= 26;
                char letter = columnName[i];

                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                result += letter - 'A' + 1;
            }
            return result;
        }
        
        /// <summary>
        /// Номер в адрес столбца.
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        private static string ColumnNumberToName(int columnNumber)
        {
            if (columnNumber < 1) return "A";

            string result = "";
            while (columnNumber > 0)
            {
                columnNumber--;
                int digit = columnNumber % 26;

                result = (char)(digit + 'A') + result;

                columnNumber /= 26;
            }

            return result;
        }

        /// <summary>
        /// Сохранить документ.
        /// </summary>
        public void Save()
        {
            Document.Save();
        }

        /// <summary>
        /// Сохранение и зыкрытие документа.
        /// </summary>
        public void Close()
        {
            if (Document != null)
            {
                Document.Close();
            }
        }

        public void Dispose()
        {
            Close();
        }
    }
}
