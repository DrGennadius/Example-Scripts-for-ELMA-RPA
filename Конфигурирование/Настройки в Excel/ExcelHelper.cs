using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
            FilePath = filePath;
            return OpenOrCreate();
        }

        /// <summary>
        /// Открыть (создать в случае отсутствия) файл.
        /// </summary>
        /// <returns></returns>
        public SpreadsheetDocument OpenOrCreate()
        {
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
        /// Получить или создать лист с наименованием.
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        public Worksheet GetOrCreateWorksheet(string worksheetName)
        {
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

        public Row GetOrCreateRow(Worksheet worksheet, uint rowIndex)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                Row nextRow = sheetData.Elements<Row>().OrderBy(r => r.RowIndex).FirstOrDefault(r => r.RowIndex > rowIndex);
                row = new Row() 
                {
                    RowIndex = rowIndex
                };
                sheetData.InsertBefore(row, nextRow);
            }
            return row;
        }

        public Cell GetOrCreateCell(Worksheet worksheet, string columnName, Row row)
        {
            string cellReference = columnName + row.RowIndex;
            Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);
            
            if (cell == null)
            {
                Cell refCell = row.Elements<Cell>().FirstOrDefault(x => string.Compare(cell.CellReference.Value, cellReference, true) > 0);
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

                cell = new Cell()
                {
                    CellReference = cellReference
                };
                row.InsertBefore(cell, refCell);

                // worksheet.Save();
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
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            Row row = GetOrCreateRow(worksheet, rowIndex);
            return GetOrCreateCell(worksheet, columnName, row);
        }

        public Cell GetOrCreateCell(Worksheet worksheet, string cellReference)
        {
            throw new NotImplementedException();
        }

        #region Set cell value

        /// <summary>
        /// Set cell text value.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, string text, CellValues dataType = CellValues.String)
        {
            cell.CellValue = new CellValue(text);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Set cell date time value.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateTime"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, DateTime dateTime, CellValues dataType = CellValues.Date)
        {
            cell.CellValue = new CellValue(dateTime);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Set cell date time offset value.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dateTimeOffset"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, DateTimeOffset dateTimeOffset, CellValues dataType = CellValues.Date)
        {
            cell.CellValue = new CellValue(dateTimeOffset);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Set cell boolean value.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, bool value, CellValues dataType = CellValues.Boolean)
        {
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Set cell double value.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, double value, CellValues dataType = CellValues.Number)
        {
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Set cell integer value.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, int value, CellValues dataType = CellValues.Number)
        {
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        /// <summary>
        /// Set cell decimal value.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="dataType"></param>
        public void SetCellValue(Cell cell, decimal value, CellValues dataType = CellValues.Number)
        {
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(dataType);
        }

        #endregion



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
