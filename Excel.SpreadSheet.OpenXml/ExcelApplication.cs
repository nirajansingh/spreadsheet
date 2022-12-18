using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel.SpreadSheet.OpenXml;

public class ExcelApplication: IDisposable
{
    private SpreadsheetDocument spreadsheet;
    private WorkbookPart workbookpart;
    private ExcelWorkbook? workbook;

    public ExcelApplication(string fileName = "New Excel.xlsx")
    {
        spreadsheet = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
        workbookpart = spreadsheet.AddWorkbookPart();
    }

    public ExcelWorkbook? Workbook { get => workbook; }

    public ExcelWorkbook AddWorkbook()
    {
        workbookpart.Workbook = new Workbook();
        return workbook = new ExcelWorkbook(workbookpart);
    }

    public void Save()
    {
        spreadsheet.Save();
    }

    public void Close()
    {
        spreadsheet.Close();
    }

    public void SaveAndClose()
    {
        Save();
        Close();
    }

    public void Dispose()
    {
        SaveAndClose();
    }
}

