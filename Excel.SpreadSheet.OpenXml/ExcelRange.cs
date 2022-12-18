using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace Excel.SpreadSheet.OpenXml
{
    public sealed class ExcelRange
    {
        private readonly Worksheet worksheet;
        private readonly string cell1;
        private readonly string cell2;

        internal ExcelRange(Worksheet worksheet, string cell1, string cell2)
        {
            this.worksheet = worksheet;
            this.cell1 = cell1;
            this.cell2 = cell2;
            Style = new ExcelStyle();
        }

        public ExcelStyle Style { get; set; }

        public int RowHeight { get; set; }

        public int ColumnWidth { get; set; }

        public void Merge()
        {
            MergeCells mergeCells;
            if (worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
            }

            var mergeCell = new MergeCell { Reference = new StringValue(ToString()) };
            mergeCells.Append(mergeCell);

            worksheet.Save();
        }

        public ExcelRange Value(string text)
        {
            string cellReference = cell2 + cell1;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return this;

            Row? row = sheetData.Elements<Row>().Where(r => r.RowIndex == cell1).FirstOrDefault();
            if (row == null)
            {
                row = new Row() { RowIndex = Convert.ToUInt32(cell1) };
                sheetData.Append(row);
            }

            Cell? newCell = row.Elements<Cell>().Where(c => c.CellReference?.Value == cellReference).FirstOrDefault();
            if (newCell == null)
            {
                Cell? refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
            }

            if (newCell != null)
            {
                newCell.CellValue = new CellValue(text);
                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                newCell.StyleIndex = 4;
            }

            worksheet.Save();

            return this;
        }

        private void GetCellStyle()
        {
            var spreadSheet = worksheet?.WorksheetPart?.OpenXmlPackage as SpreadsheetDocument;
            if (spreadSheet != null)
            {
                var stylesPart = spreadSheet?.WorkbookPart?.GetPartsOfType<WorkbookStylesPart>() as WorkbookStylesPart;
                if (stylesPart != null)
                {
                    var fills = stylesPart.Stylesheet.Elements<Fills>().First();
                    if (fills != null)
                    {
                        foreach (Fill fill in fills.Elements<Fill>())
                        {
                            if (fill.PatternFill?.BackgroundColor?.Rgb?.Value == Style.BackgroundColor)
                            {

                            }
                        }
                    }
                }

            }

        }

        public override string ToString()
        {
            return $"{cell1}:{cell2}";
        }
    }
}