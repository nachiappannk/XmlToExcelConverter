using System;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace XmlToExcel
{
    public class ExcelWriter : IDisposable
    {
        private readonly ExcelPackage _package;
        private readonly ExcelWorksheet _workSheet;

        private int _rowToWrite = 1;
        public ExcelWriter(string fileName, string sheetName)
        {
            FileInfo file = new FileInfo(fileName);
            _package = new ExcelPackage(file);
            _workSheet = CreateWorkSheet(sheetName, _package);
        }

        private ExcelWorksheet CreateWorkSheet(string sheetName, ExcelPackage excelPackage)
        {
            if (!IsSheetPresent(sheetName, excelPackage))
            {
                return excelPackage.Workbook.Worksheets.Add(sheetName);
            }
            for(int i = 0; ;i++)
            {
                var newSheetName = sheetName + i;
                if (!IsSheetPresent(newSheetName, excelPackage))
                {
                    return excelPackage.Workbook.Worksheets.Add(newSheetName);
                }
            }
        }

        private static bool IsSheetPresent(string sheetName, ExcelPackage excelPackage)
        {
            return excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName) != null;
        }

        public void Write(params string[] strs)
        {
            for (int i = 1; i <= strs.Length; i++)
            {
                _workSheet.Cells[_rowToWrite, i].Value = strs[i - 1];
            }
            _rowToWrite++;
        }


        public void WriteLineGreen(params string[] strs)
        {
            using (var range = _workSheet.Cells[_rowToWrite, 1, _rowToWrite, strs.Length])
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            }
            Write(strs);
        }

        public void WriteLineBlue(params string[] strs)
        {
            using (var range = _workSheet.Cells[_rowToWrite, 1, _rowToWrite, strs.Length])
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            }
            Write(strs);
        }

        public void WriteHeading(params string[] strs)
        {
            Write(strs);
            using (var range = _workSheet.Cells[1, 1, 1, strs.Length])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                range.Style.Font.Color.SetColor(Color.White);
                range.AutoFilter = true;
            }
        }

        public void Dispose()
        {
            _package.Save();
            _workSheet.Dispose();
            _package.Dispose();
        }
    }
}