using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlPocDocker;
using System;
using System.IO;

namespace testdocker
{
    public static class ExcelGenerator
    {
        public static void CreateExcelFile(string OutPutFileDirectory)
        {
            TestModelList data = TestModelData.GetList();

            var fileFullname = Path.Combine(OutPutFileDirectory, $"output_{GetTimestamp()}.xlsx");

            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullname, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, data);

                //CalculationProperties p = new CalculationProperties
                //{
                //    ForceFullCalculation = true,
                //    FullCalculationOnLoad = true,
                //    CalculationMode = CalculateModeValues.Auto,
                //    CalculationCompleted = true,

                //};
                //package.WorkbookPart.Workbook.CalculationProperties = p;
            }
        }
        private static String GetTimestamp()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }
        private static void CreatePartsForExcel(SpreadsheetDocument document, TestModelList data)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(data);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>();
            GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("id1");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData);
        }

        private static void GenerateWorkbookPartContent(WorkbookPart workbookPart)
        {
            Workbook workbook1 = new Workbook();
            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "poc", SheetId = (UInt32Value)1U, Id = "id1" };
            sheets1.Append(sheet1);
            workbook1.Append(sheets1);
            workbookPart.Workbook = workbook1;
        }
        private static void GenerateWorksheetPartContent(WorksheetPart worksheetPart, SheetData sheetData)
        {
            Worksheet worksheet1 = new Worksheet();

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);

            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetData);
            worksheetPart.Worksheet = worksheet1;
        }
        private static void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet1 = new Stylesheet();

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 18D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            Underline ul = new Underline();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Rgb = "FFFF0000" };
            FontName fontName2 = new FontName() { Val = "Times New Roman" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(ul);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Solid };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thick };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thick };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);

            workbookStylesPart.Stylesheet = stylesheet1;
        }
        private static SheetData GenerateSheetdataForDetails(TestModelList data)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateHeaderRowForExcel());

            foreach (TestModel testmodel in data.testData)
            {
                Row partsRows = GenerateRowForChildPartDetail(testmodel);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }
        private static Row CreateHeaderRowForExcel()
        {
            Row workRow = new Row();
            workRow.Append(CreateCell("Test Id", 1U));
            workRow.Append(CreateCell("Test Name", 2U));
            workRow.Append(CreateCell("Test Description", 1U));

            workRow.Append(CreateCell("TestDesc1", 2U));
            workRow.Append(CreateCell("TestDesc2", 2U));
            workRow.Append(CreateCell("TestDesc3", 2U));
            workRow.Append(CreateCell("TestDesc4", 2U));
            workRow.Append(CreateCell("TestDesc5", 2U));
            workRow.Append(CreateCell("TestDesc6", 2U));
            workRow.Append(CreateCell("TestDesc7", 2U));
            workRow.Append(CreateCell("TestDesc8", 2U));
            workRow.Append(CreateCell("TestDesc9", 2U));
            workRow.Append(CreateCell("TestDesc10", 2U));
            workRow.Append(CreateCell("TestDesc11", 2U));
            workRow.Append(CreateCell("TestDesc12", 2U));
            workRow.Append(CreateCell("TestDesc13", 2U));
            workRow.Append(CreateCell("TestDesc14", 2U));
            workRow.Append(CreateCell("TestDesc15", 2U));
            workRow.Append(CreateCell("TestDesc16", 2U));
            workRow.Append(CreateCell("TestDesc17", 2U));
            workRow.Append(CreateCell("TestDesc18", 2U));
            workRow.Append(CreateCell("TestDesc19", 2U));
            workRow.Append(CreateCell("TestDesc20", 2U));


            workRow.Append(CreateCell("Test Date", 2U));
            return workRow;
        }
        private static Row GenerateRowForChildPartDetail(TestModel testmodel)
        {
            Row tRow = new Row();
            tRow.Append(CreateCell(testmodel.TestId.ToString()));
            tRow.Append(CreateCell(testmodel.TestName));
            tRow.Append(CreateCell(testmodel.TestDesc));

            tRow.Append(CreateCell(testmodel.TestDesc1));
            tRow.Append(CreateCell(testmodel.TestDesc2));
            tRow.Append(CreateCell(testmodel.TestDesc3));
            tRow.Append(CreateCellFormula(testmodel.TestDesc4));
            tRow.Append(CreateCell(testmodel.TestDesc5));
            tRow.Append(CreateCell(testmodel.TestDesc6));
            tRow.Append(CreateCell(testmodel.TestDesc7));
            tRow.Append(CreateCell(testmodel.TestDesc8));
            tRow.Append(CreateCell(testmodel.TestDesc9));
            tRow.Append(CreateCell(testmodel.TestDesc10));
            tRow.Append(CreateCell(testmodel.TestDesc11));
            tRow.Append(CreateCell(testmodel.TestDesc12));
            tRow.Append(CreateCell(testmodel.TestDesc13));
            tRow.Append(CreateCell(testmodel.TestDesc14));
            tRow.Append(CreateCell(testmodel.TestDesc15));
            tRow.Append(CreateCell(testmodel.TestDesc16));
            tRow.Append(CreateCell(testmodel.TestDesc17));
            tRow.Append(CreateCell(testmodel.TestDesc18));
            tRow.Append(CreateCell(testmodel.TestDesc19));
            tRow.Append(CreateCell(testmodel.TestDesc20));


            tRow.Append(CreateCell(testmodel.TestDate.ToShortDateString()));

            return tRow;
        }
        private static Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;

        }

        private static Cell CreateCell(int text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.CellValue = new CellValue(text);
            return cell;

        }
        private static Cell CreateCellFormula(string text)
        {
            Cell cell = new Cell
            {
                StyleIndex = 1U,
                DataType = new EnumValue<CellValues>(CellValues.Number),
            };

            CellFormula forumula = new CellFormula("SUM(D4:F4)")
            {
                CalculateCell = true
            };

            cell.CellFormula = forumula;
            cell.CellFormula.CalculateCell = true;
            return cell;

        }

        private static Cell CreateCell(string text, uint styleIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private static EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
    }
}
