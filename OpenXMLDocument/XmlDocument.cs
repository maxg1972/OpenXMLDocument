using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLDocument
{
    public class XmlDocument
    {
        protected WorkbookPart WorkBookPart;
        protected WorksheetPart WorkSheetPart;

        protected XmlDocument()
        {
            WorkBookPart = null;
            WorkSheetPart = null;
        }
        
        /// <summary>
        /// Generate XML excel document
        /// </summary>
        /// <param name="dataset">Set of data (list of rows, where each row is a list of strings representing each cells value))</param>
        /// <param name="headers">List of string representing column headers</param>
        /// <param name="fileName">Full path nome of file will be create</param>
        /// <param name="sheetName">Nome of the sheet (default: Sheet1)</param>
        /// <param name="overwriteIfExists">Overwrite existing file (default: true)</param>
        public void CreateExcelDoc(IEnumerable<List<string>> dataset, IEnumerable<string> headers, string fileName, string sheetName = "Sheet1", bool overwriteIfExists = true)
        {
            if (overwriteIfExists && File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            
            using (var document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkBookPart = Init(document,sheetName);
         
                var sheetData = WorkSheetPart.Worksheet.AppendChild(new SheetData());
         
                // Row header
                var row = MakeHeaderRow(headers);
                sheetData.AppendChild(row);
         
                // Data rows
                var rowIndex = 2;
                foreach (var rowItem in dataset)
                {
                    row = new Row
                    {
                        RowIndex = (uint) rowIndex
                    };

                    foreach (var colItem in rowItem)
                    {
                        row.AppendChild(MakeCell(colItem, CellValues.String));
                    }

                    sheetData.AppendChild(row);

                    rowIndex++;
                }
         
                WorkSheetPart.Worksheet.Save();
            }
        }

        /// <summary>
        /// Initialize dcoument
        /// </summary>
        /// <param name="document">Shpreadsheet document pointer</param>
        /// <param name="sheetName">Name of the sheet</param>
        /// <returns>WorkbookPart object</returns>
        protected WorkbookPart Init(SpreadsheetDocument document, string sheetName)
        {
            WorkBookPart = document.AddWorkbookPart();
            WorkBookPart.Workbook = new Workbook();
         
            WorkSheetPart = WorkBookPart.AddNewPart<WorksheetPart>();
            WorkSheetPart.Worksheet = new Worksheet();
         
            var sheets = WorkBookPart.Workbook.AppendChild(new Sheets());
         
            var sheet = new Sheet
            {
                Id = WorkBookPart.GetIdOfPart(WorkSheetPart), 
                SheetId = 1, 
                Name = sheetName
            };
         
            sheets.AppendChild(sheet);
         
            WorkBookPart.Workbook.Save();

            return WorkBookPart;
        }
         
        /// <summary>
        /// Generate row object with headers
        /// </summary>
        /// <param name="headers">List of string headers</param>
        /// <param name="startIndex">Row index for headers (default: 1) </param>
        /// <returns>Row object</returns>
        protected Row MakeHeaderRow(IEnumerable<string> headers, int startIndex = 1)
        {
            var row = new Row
            {
                RowIndex = (uint) startIndex
            };

            foreach (var colItem in headers)
            {
                row.AppendChild(MakeCell(colItem, CellValues.String));
            }

            return row;
        }

        /// <summary>
        /// Generate cell object
        /// </summary>
        /// <param name="value">Value of the cell in string format</param>
        /// <param name="dataType">Type of the value</param>
        /// <returns>Cell object</returns>
        protected Cell MakeCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
    }
}