﻿using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel.SpreadSheet.OpenXml;

public class ExcelWorksheet
{
    private readonly WorksheetPart worksheetPart;
    private readonly Sheet sheet;
    private readonly List<ExcelRange> ranges;

    internal ExcelWorksheet(WorksheetPart worksheetPart, Sheet sheet)
    {
        this.worksheetPart = worksheetPart;
        this.sheet = sheet;
        ranges = new List<ExcelRange>();
    }

    public void Rename(string sheetName)
    {
        sheet.Name = sheetName;
    }

    public ExcelRange Range(string cell1, string cell2)
    {
        return new ExcelRange(worksheetPart.Worksheet, cell1, cell2);
    }

    public ExcelRange Columns(string columnIndex)
    {
        if (string.IsNullOrWhiteSpace(columnIndex))
        {
            throw new ArgumentException(nameof(columnIndex));
        }

        string[] cells = new string[2];
        if (columnIndex.Contains(':'))
        {
            cells = columnIndex.Split(':');
        }
        else
        {
            cells[0] = columnIndex;
            cells[1] = columnIndex;
        }

        return new ExcelRange(worksheetPart.Worksheet, cells[0], cells[1]);
    }

    public ExcelRange Cells(int rowIndex, string columnIndex)
    {
        return new ExcelRange(worksheetPart.Worksheet, Convert.ToString(rowIndex), columnIndex);
    }

    public void AddImage(string imageFileName)
    {
        DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
        { Id = worksheetPart.GetIdOfPart(drawingsPart) });

        
        var imagePart = worksheetPart.AddImagePart(ImagePartType.Bmp);

        
        worksheetPart.Worksheet.Save();
    }
}
