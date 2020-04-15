using Excel = Microsoft.Office.Interop.Excel; 

private static Excel.Workbook MyBook = null; 
private static Excel.Application MyApp = null;
private static Excel.Worksheet MySheet = null;

MyApp = new Excel.Application();
MyApp.Visible = false;
MyBook = MyApp.Workbooks.Open('./data.xlsx');
MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; 