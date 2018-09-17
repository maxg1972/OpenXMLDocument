using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLDocument;

namespace TestExcel
{
    class ExcelDocs : XmlDocument
    {
        public void  CreateExcelDoc(IEnumerable<List<int>> dataset, IEnumerable<string> headers, string fileName,bool overwriteIfExists = true, string sheetName = "Foglio1")
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
                        row.AppendChild(MakeCell(colItem.ToString(), CellValues.Number));
                    }

                    sheetData.AppendChild(row);

                    rowIndex++;
                }
         
                WorkSheetPart.Worksheet.Save();
            }
        }
    }

    class TestProgram
    {
        static void Main(string[] args)
        {
            var excel = new ExcelDocs();
            
            var righe = new List<List<int>>();

            for (var rowIndex = 1; rowIndex <= 5; rowIndex++)
            {
                var colonne = new List<int>();
                for (var colIndex = 1; colIndex <= 10; colIndex++)
                {
                    colonne.Add(new Random().Next(1,1000 * rowIndex));    
                }
                righe.Add(colonne);
            }

            var intestazioni = new List<string>();
            for (var colIndex = 1; colIndex <= 10; colIndex++)
            {
                intestazioni.Add("Colonna " + colIndex);    
            }
            
            excel.CreateExcelDoc(righe,intestazioni,"./test1.xls");
            
            
            var righe2 = new  List<List<string>>();

            for (var rowIndex = 1; rowIndex <= 5; rowIndex++)
            {
                var colonne = new List<string>();
                for (var colIndex = 1; colIndex <= 10; colIndex++)
                {
                    colonne.Add(new Random().Next(1,1000 * rowIndex).ToString());    
                }
                righe2.Add(colonne);
            }

            intestazioni = new List<string>();
            for (var colIndex = 1; colIndex <= 10; colIndex++)
            {
                intestazioni.Add("Colonna " + colIndex);    
            }
            
            excel.CreateExcelDoc(righe2,intestazioni,"./test2.xls", "Foglio2");
        }
    }
}