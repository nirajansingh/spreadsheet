using System;

namespace Excel.SpreadSheet.OpenXml;

public struct ExcelStyle
{
    public ExcelStyle()
    {
        Font = new ExcelFont();
        BackgroundColor = "";
    }

	public ExcelFont Font { get; set; }
    public HorizontalAlignment HorizontalAlignment { get; set; }
    public VerticalAlignment VerticalAlignment { get; set; }
    public string BackgroundColor { get; set; }
}

public struct ExcelFont
{
    public string Family { get; set; }
    public int Size { get; set; }
    public string Color { get; set; }
    public int Weight { get; set; }
}

