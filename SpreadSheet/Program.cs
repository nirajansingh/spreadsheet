using Excel.SpreadSheet.OpenXml;

var app = new ExcelApplication();
var workbook = app.AddWorkbook();
workbook.AddWorksheet();
workbook.AddWorksheet();

var sheet1 = workbook.Find(0);
sheet1.Rename("Summary");

var sheet2 = workbook.Find(1);
sheet2.Rename("Data");

sheet1.Range("A1", "F1").Merge();
sheet1.Cells(1, "A").Value("Hello Excel!");

app.SaveAndClose();


