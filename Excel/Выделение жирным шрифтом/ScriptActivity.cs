using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

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
            GetData(context, out SpreadsheetDocument spreadsheetDocument, out WorksheetPart worksheetPart);

            // 0. Пример новой ячейки с простым текстом в A1
            Cell newCell1 = InsertCellInWorksheet("A", 1, worksheetPart);
            newCell1.CellValue = new CellValue("Текст");
            newCell1.DataType = new EnumValue<CellValues>(CellValues.String);

            // Чтобы выделить жирным текст, необходимо сначала создать и добавить стиль шрифта,
            // Потом нужно создать формат ячеек и тоже добавить в список.
            // А уже потом используем формат для нужных ячеек.

            // 1. Получаем/создаем объект типа WorkbookStylesPart
            WorkbookStylesPart stylesPart = spreadsheetDocument.WorkbookPart.WorkbookStylesPart;
            if (stylesPart == null)
            {
                stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            }
            if (stylesPart.Stylesheet == null)
            {
                stylesPart.Stylesheet = new Stylesheet();
            }
            if (stylesPart.Stylesheet.Fonts == null)
            {
                stylesPart.Stylesheet.Fonts = new Fonts();
            }
            if (stylesPart.Stylesheet.CellFormats == null)
            {
                stylesPart.Stylesheet.CellFormats = new CellFormats();
            }
            if (stylesPart.Stylesheet.Fills == null)
            {
                stylesPart.Stylesheet.Fills = new Fills();
            }
            if (stylesPart.Stylesheet.Borders == null)
            {
                stylesPart.Stylesheet.Borders = new Borders();
            }

            // 2. Проверяем шрифты. Предпологаем, что если нет, то создаем.
            if (stylesPart.Stylesheet.Fonts.ChildElements.Count == 0)
            {
                // Создаем некий "стандартный" шрифт, который пустой, и добавляем его в список
                Font defaultFont = new Font();
                stylesPart.Stylesheet.Fonts.Append(defaultFont);

                // Создаем "жирный" шрифт и добавляем
                Font boldFont = new Font(
                    // Тут определяем, что шрифт будет жирным
                    new Bold()
                );
                stylesPart.Stylesheet.Fonts.Append(boldFont);
            }

            // 3. Проверяем заполнения. Предпологаем, что если нет, то создаем хотя бы 1 "стандартный".
            if (stylesPart.Stylesheet.Fills.ChildElements.Count == 0)
            {
                Fill defaultFill = new Fill();
                stylesPart.Stylesheet.Fills.Append(defaultFill);
            }

            // 4. Проверяем рамки. Предпологаем, что если нет, то создаем хотя бы 1 "стандартные".
            if (stylesPart.Stylesheet.Borders.ChildElements.Count == 0)
            {
                Border defaultBorder = new Border();
                stylesPart.Stylesheet.Borders.Append(defaultBorder);
            }

            // 5. Проверяем форматы ячейки. Предпологаем, что если нет, то создаем.
            if (stylesPart.Stylesheet.CellFormats.ChildElements.Count == 0)
            {
                // Создаем "стандартный" формат ячейки и добавляем его в список
                CellFormat defaultCellformat = new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 };
                stylesPart.Stylesheet.CellFormats.Append(defaultCellformat);

                // Создаем формат "жирной" ячейки
                CellFormat boldCellFormat = new CellFormat()
                {
                    FontId = 1,
                    FillId = 0,
                    BorderId = 0,
                    ApplyFont = true
                };
                stylesPart.Stylesheet.CellFormats.Append(boldCellFormat);
            }

            // 6. Пример новой ячейки с жирным текстом в A2
            Cell newCell2 = InsertCellInWorksheet("A", 2, worksheetPart);
            newCell2.CellValue = new CellValue("Жирный текст");
            newCell2.DataType = new EnumValue<CellValues>(CellValues.String);
            // Используем id формата "жирной" ячейки
            UInt32Value boldFontId = 0U;
            UInt32Value boldCellFormatId = 0U;
            for (int i = 0; i < stylesPart.Stylesheet.Fonts.ChildElements.Count; i++)
            {
                if (stylesPart.Stylesheet.Fonts.ElementAt(i).GetFirstChild<Bold>() != null)
                {
                    boldFontId = Convert.ToUInt32(i);
                    break;
                }
            }
            for (int i = 0; i < stylesPart.Stylesheet.CellFormats.ChildElements.Count; i++)
            {
                if ((((CellFormat)stylesPart.Stylesheet.CellFormats.ElementAt(i))).FontId == boldFontId)
                {
                    boldCellFormatId = Convert.ToUInt32(i);
                    break;
                }
            }
            newCell2.StyleIndex = boldCellFormatId;

            // Не забываем закрывать, если не использовали блок using:
            spreadsheetDocument.Close();
        }

        /// <summary>
        /// Получаем данные: SpreadsheetDocument, worksheetPart и sheetData
        /// </summary>
        /// <param name="context"></param>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="worksheetPart"></param>
        /// <param name="sheetData"></param>
        private void GetData(Context context, out SpreadsheetDocument spreadsheetDocument, out WorksheetPart worksheetPart)
        {
            // В данном методе открывается/создается файл excel и получаем/создаем элементы структуры.
            // C комментариями можно разобрать тут:
            // https://github.com/DrGennadius/Example-Scripts-for-ELMA-RPA/blob/master/Excel/%D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5%20%D0%BD%D0%BE%D0%B2%D0%BE%D0%B3%D0%BE%20%D0%BB%D0%B8%D1%81%D1%82%D0%B0/ScriptActivity.cs
            spreadsheetDocument = null;
            bool excelExists = File.Exists(context.ExcelFilePath);
            if (excelExists)
            {
                spreadsheetDocument = SpreadsheetDocument.Open(context.ExcelFilePath, true);
            }
            else
            {
                spreadsheetDocument = SpreadsheetDocument.Create(context.ExcelFilePath, SpreadsheetDocumentType.Workbook);
            }
            WorkbookPart workbookPart = excelExists ? spreadsheetDocument.WorkbookPart : spreadsheetDocument.AddWorkbookPart();
            if (!excelExists || workbookPart == null)
            {
                workbookPart.Workbook = new Workbook();
            }
            Sheet sheet = null;
            worksheetPart = null;
            if (excelExists)
            {
                sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == context.SheetName);
                if (sheet == null)
                {
                    worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                }
                else
                {
                    worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                }
            }
            if (worksheetPart == null)
            {
                worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            }
            if (sheet == null)
            {
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                uint newSheetId = 1U;
                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                if (sheets != null)
                {
                    Sheet lastSheet = sheets.Elements<Sheet>().OrderBy(x => x.SheetId).LastOrDefault();
                    if (lastSheet != null)
                    {
                        newSheetId = lastSheet.SheetId + 1;
                    }
                }
                else
                {
                    sheets = workbookPart.Workbook.AppendChild(new Sheets());
                }

                sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = newSheetId,
                    Name = context.SheetName
                };
                sheets.Append(sheet);
            }
        }

        /// <summary>
        /// Вставить данные в ячейку листа
        /// </summary>
        /// <param name="columnName">Имя столбца в формате как в Excel [A-Z...]</param>
        /// <param name="rowIndex">Индекс строки</param>
        /// <param name="worksheetPart"></param>
        /// <param name="value">Значение</param>
        /// <param name="type">Тип значения</param>
        /// <returns></returns>
        private Cell InsertInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart, string value, CellValues type = CellValues.String)
        {
            Cell newCell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);
            newCell.CellValue = new CellValue(value);
            newCell.DataType = new EnumValue<CellValues>(type);
            return newCell;
        }

        /// <summary>
        /// Метод для получения или создания ячейки и вставки данных в эту ячейку
        /// </summary>
        /// <param name="columnName">Имя столбца в формате как в Excel [A-Z...]</param>
        /// <param name="rowIndex">Индекс строки</param>
        /// <param name="worksheetPart"></param>
        /// <returns>Ячейка</returns>
        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // Если рабочий лист не содержит строки с указанным индексом строки, то просто добавляем ее.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                // Обязательно необходимо соблюсти последовательность индексов
                Row nextRow = sheetData.Elements<Row>().OrderBy(r => r.RowIndex).FirstOrDefault(r => r.RowIndex > rowIndex);
                row = new Row() { RowIndex = rowIndex };
                sheetData.InsertBefore(row, nextRow);
            }

            // Если ячейки с указанным именем столбца нет, то добавляем.
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Ячейки должны располагаться в последовательном порядке согласно CellReference.
                // Необходимо определить, куда вставить новую ячейку.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
    }
}
