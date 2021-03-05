using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using OpenXmlPocDocker;
using System;
using System.Collections.Generic;
using System.Data;

namespace testdocker
{
    class Program
    {


        static void Main(string[] args)
        {
            //var filename = $"/tmp/TestNewData_{GetTimestamp(DateTime.Now)}.xlsx";
            var filename = $"TestNewData_{GetTimestamp()}.xlsx";
            WriteExcelFile(filename);
        }

        static void WriteExcelFile(string filename)
        {
            // Lets converts our object data to Datatable for a simplified logic.
            // Datatable is most easy way to deal with complex datatypes for easy reading and formatting. 
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(UserDetailsData.Persons), (typeof(DataTable)));

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(column.ColumnName)
                    };

                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (String col in columns)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
                Console.WriteLine(filename);
            }
        }
        private static String GetTimestamp()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }
    }
}
