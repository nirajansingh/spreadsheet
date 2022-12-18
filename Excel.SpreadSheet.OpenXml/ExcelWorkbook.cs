using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel.SpreadSheet.OpenXml;

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
}

