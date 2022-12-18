using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel.SpreadSheet.OpenXml
{
    public class ExcelWorkbook
    {
        private readonly WorkbookPart workbookPart;
        private WorksheetPart worksheetPart;
        private Sheets sheets;
        private List<ExcelWorksheet> worksheets;

        internal ExcelWorkbook(WorkbookPart workbookPart)
        {
            worksheets = new List<ExcelWorksheet>();
            this.workbookPart = workbookPart;

            worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            sheets = workbookPart.Workbook.AppendChild(new Sheets());
        }

        public ExcelWorksheet AddWorksheet()
        {
            return AddWorksheet($"Sheet {worksheets.Count + 1}");
        }

        public ExcelWorksheet AddWorksheet(string sheetName)
        {
            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)worksheets.Count + 1,
                Name = sheetName
            };

            sheets.Append(sheet);

            var worksheet = new ExcelWorksheet(worksheetPart, sheet);
            worksheets.Add(worksheet);
            return worksheet;
        }

        public ExcelWorksheet Find(int index)
        {
            return worksheets.ElementAt(index); ;
        }

        public void AddStyle()
        {
            var styleSheet = new Stylesheet();
            // Create "fonts" node.
            var fonts = new Fonts();
            fonts.Append(new Font()
            {
                FontName = new FontName() { Val = "Arial" },
                FontSize = new FontSize() { Val = 11 },
                FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 }
            });

            fonts.Count = (uint)fonts.ChildElements.Count;

            // Create "fills" node.
            var fills = new Fills();
            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } });
            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Solid, BackgroundColor = new BackgroundColor { Rgb = new HexBinaryValue { Value = "#2E9D00" } } } });
            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Solid, BackgroundColor = new BackgroundColor { Rgb = new HexBinaryValue { Value = "#2E9DFF" } } } });

            fills.Count = (uint)fills.ChildElements.Count;

            // Create "borders" node.
            var borders = new Borders();
            borders.Append(new Border()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });

            borders.Count = (uint)borders.ChildElements.Count;

            // Create "cellStyleXfs" node.
            var cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.Append(new CellFormat()
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            });

            cellStyleFormats.Count = (uint)cellStyleFormats.ChildElements.Count;

            // Create "cellXfs" node.
            var cellFormats = new CellFormats();

            // StyleIndex = 0, A default style that works for most things (But not strings? )
            cellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            });

            // StyleIndex = 1, A style that works for DateTime (just the date)
            cellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 14, //Date
                FormatId = 0,
                ApplyNumberFormat = true
            });

            // StyleIndex = 2, A style that works for DateTime (Date and Time)
            cellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 22, //Date Time
                FormatId = 0,
                ApplyNumberFormat = true
            });
            //StyleIndex = 3
            cellFormats.Append(new CellFormat { Alignment = new Alignment { Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Center), Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Top) } });
            //StyleIndex = 4
            cellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 2U,
                FontId = 0,
                ApplyFill = true,
                ApplyBorder = true,
            });

            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            // Create "cellStyles" node.
            var cellStyles = new CellStyles();
            cellStyles.Append(new CellStyle()
            {
                Name = "Normal",
                FormatId = 0,
                BuiltinId = 0,
            });
            cellStyles.Count = (uint)cellStyles.ChildElements.Count;

            // Append all nodes in order.
            styleSheet.Append(fonts);
            styleSheet.Append(fills);
            styleSheet.Append(borders);
            styleSheet.Append(cellStyleFormats);
            styleSheet.Append(cellFormats);
            styleSheet.Append(cellStyles);

            var WorkbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            WorkbookStylesPart.Stylesheet = styleSheet;
            WorkbookStylesPart.Stylesheet.Save();
        }
    }
}