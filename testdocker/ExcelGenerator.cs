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
            Sheet sheet1 = new Sheet() { Name = "poc", SheetId = 1, Id = "id1" };
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
            var styleSheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            styleSheet.AddNamespaceDeclaration("mc", "http: //schemas.openxmlformats.org/markup-compatibility/2016");
            styleSheet.AddNamespaceDeclaration("x14ac", "http: //schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            #region fonts

            Fonts fontList = new Fonts { Count = 3 };

            Font f1 = new Font();
            Color color1 = new Color() { Rgb = HexBinaryValue.FromString("FF000000") };
            f1.Append(color1);

            Font f2 = new Font();
            FontSize fontSize2 = new FontSize() { Val = 12D };
            Color color2 = new Color() { Rgb = HexBinaryValue.FromString("FF808080") };
            Bold bold2 = new Bold();
            Italic it2 = new Italic();
            FontName fontName2 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 1 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            f2.Append(fontSize2);
            f2.Append(color2);
            f2.Append(fontName2);
            f2.Append(bold2);
            f2.Append(it2);
            f2.Append(fontFamilyNumbering2);
            f2.Append(fontScheme2);

            Font f3 = new Font();
            FontSize fontSize3 = new FontSize() { Val = 16D };
            Color color3 = new Color() { Rgb = HexBinaryValue.FromString("FF0000FF") };
            Underline ud3 = new Underline();
            FontName fontName3 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 1 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            f3.Append(fontSize3);
            f3.Append(color3);
            f3.Append(ud3);
            f3.Append(fontName3);
            f3.Append(fontFamilyNumbering3);
            f3.Append(fontScheme3);

            fontList.Append(f1);
            fontList.Append(f2);
            fontList.Append(f3);
            #endregion

            #region Fills

            Fills fillList = new Fills();

            // create a solid red fill
            var solidRed = new PatternFill() { PatternType = PatternValues.Solid };
            solidRed.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") }; // red fill
            solidRed.BackgroundColor = new BackgroundColor { Indexed = 64 };

            fillList.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            fillList.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            fillList.AppendChild(new Fill { PatternFill = solidRed });
            fillList.Count = 3;
            #endregion


            #region Borders
            Borders bordersList = new Borders
            {
                Count = 1
            };

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

            bordersList.Append(border1);


            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thick };
            Color colorborder = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(colorborder);

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

            bordersList.Append(border2);

            #endregion

            #region Cellformats



            // blank cell format list
            CellStyleFormats blankcellStyleFormatList = new CellStyleFormats
            {
                Count = 1
            };
            blankcellStyleFormatList.AppendChild(new CellFormat());

            // cell format list
            CellFormats cellStyleFormatList = new CellFormats();

            // empty one for index 0, seems to be required
            cellStyleFormatList.AppendChild(new CellFormat());

            // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
            cellStyleFormatList.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 1, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
            cellStyleFormatList.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 1, FillId = 0, ApplyFill = true });
            cellStyleFormatList.AppendChild(new CellFormat { FormatId = 0, FontId = 2, BorderId = 1, FillId = 0, ApplyFill = true });
            cellStyleFormatList.Count = 4;

            #endregion

            styleSheet.Append(fontList);
            styleSheet.Append(fillList);
            styleSheet.Append(bordersList);
            styleSheet.Append(blankcellStyleFormatList);
            styleSheet.Append(cellStyleFormatList);

            workbookStylesPart.Stylesheet = styleSheet;
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
            workRow.Append(CreateCell("Test Description", 0U));

            workRow.Append(CreateCell("TestDesc1", 0U));
            workRow.Append(CreateCell("TestDesc2", 0U));
            workRow.Append(CreateCell("TestDesc3", 0U));
            workRow.Append(CreateCell("TestDesc4", 0U));
            workRow.Append(CreateCell("TestDesc5", 0U));
            workRow.Append(CreateCell("TestDesc6", 0U));
            workRow.Append(CreateCell("TestDesc7", 0U));
            workRow.Append(CreateCell("TestDesc8", 0U));
            workRow.Append(CreateCell("TestDesc9", 0U));
            workRow.Append(CreateCell("TestDesc10", 0U));
            workRow.Append(CreateCell("TestDesc11", 0U));
            workRow.Append(CreateCell("TestDesc12", 0U));
            workRow.Append(CreateCell("TestDesc13", 0U));
            workRow.Append(CreateCell("TestDesc14", 0U));
            workRow.Append(CreateCell("TestDesc15", 0U));
            workRow.Append(CreateCell("TestDesc16", 0U));
            workRow.Append(CreateCell("TestDesc17", 0U));
            workRow.Append(CreateCell("TestDesc18", 0U));
            workRow.Append(CreateCell("TestDesc19", 0U));
            workRow.Append(CreateCell("TestDesc20", 0U));


            workRow.Append(CreateCell("Test Date", 0U));
            return workRow;
        }
        private static Row GenerateRowForChildPartDetail(TestModel testmodel)
        {
            Row tRow = new Row();
            tRow.Append(CreateCell(testmodel.TestId.ToString()));
            tRow.Append(CreateCell(testmodel.TestName, 1));
            tRow.Append(CreateCell(testmodel.TestDesc, 2));

            tRow.Append(CreateCell(testmodel.TestDesc1));
            tRow.Append(CreateCell(testmodel.TestDesc2));
            tRow.Append(CreateCell(testmodel.TestDesc3));
            tRow.Append(CreateCellFormula(testmodel.TestDesc4));
            tRow.Append(CreateCellStringFormula(testmodel.TestDesc5));

            tRow.Append(BuildHyperlinkCell("lien hypertexte", "https://www.rapidtables.com/web/color/RGB_Color.html"));
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
            cell.StyleIndex = 0U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;

        }

        private static Cell CreateCell(int text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 0U;
            cell.CellValue = new CellValue(text);
            return cell;

        }
        private static Cell CreateCellFormula(string text)
        {
            Cell cell = new Cell
            {
                StyleIndex = 0U,
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

        private static Cell CreateCellStringFormula(string text)
        {
            Cell cell = new Cell
            {
                StyleIndex = 0U,
                DataType = new EnumValue<CellValues>(CellValues.String),
            };

            CellFormula forumula = new CellFormula("=CONCAT(J4:L4)")
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

        private static Cell BuildHyperlinkCell(string val, string url) =>
    new Cell
    {
        DataType = new EnumValue<CellValues>(CellValues.SharedString),
        CellFormula = new CellFormula($"HyperLink(\"{url}\",\"{val}\")"),
        StyleIndex = 3
    };

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
