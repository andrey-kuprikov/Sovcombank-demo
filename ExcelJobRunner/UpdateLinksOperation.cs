using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

public static class UpdateLinksOperation
{
    public static UpdateLinksResult Run(UpdateLinksParams parameters)
    {
        if (!File.Exists(parameters.InputFile))
            throw new FileNotFoundException($"Input file not found: {parameters.InputFile}");

        var excel = new Excel.Application { DisplayAlerts = false, Visible = false };
        Excel.Workbook? workbook = null;
        var updated = new List<UpdatedCell>();

        try
        {
            workbook = excel.Workbooks.Open(parameters.InputFile, ReadOnly: false);

            var cellsInfo = new List<(string sheet, string addr, string newValue, string oldValue)>();

            foreach (var cell in parameters.Cells)
            {
                var (sheetName, cellAddr) = SplitAddress(cell.Address);
                Excel.Worksheet sheet;
                try
                {
                    sheet = (Excel.Worksheet)workbook.Worksheets[sheetName];
                }
                catch
                {
                    throw new Exception($"Адрес ячейки не найден: {cell.Address}");
                }

                try
                {
                    Excel.Range range = sheet.Range[cellAddr];
                    string oldVal = Convert.ToString(range.Value2 ?? range.Value) ?? string.Empty;
                    cellsInfo.Add((sheetName, cellAddr, cell.NewValue, oldVal));
                    Marshal.ReleaseComObject(range);
                }
                catch
                {
                    throw new Exception($"Адрес ячейки не найден: {cell.Address}");
                }
                finally
                {
                    Marshal.ReleaseComObject(sheet);
                }
            }

            foreach (var info in cellsInfo)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets[info.sheet];
                Excel.Range range = sheet.Range[info.addr];
                range.Value2 = info.newValue;
                updated.Add(new UpdatedCell
                {
                    Address = $"{info.sheet}!{info.addr}",
                    OldValue = info.oldValue,
                    NewValue = info.newValue
                });
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(sheet);
            }

            workbook.SaveCopyAs(parameters.OutputFile);
            return new UpdateLinksResult { Status = "OK", UpdatedCells = updated };
        }
        finally
        {
            workbook?.Close(false);
            excel.Quit();
            if (workbook != null) Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excel);
        }
    }

    private static (string sheet, string addr) SplitAddress(string fullAddress)
    {
        var idx = fullAddress.IndexOf('!');
        if (idx <= 0 || idx >= fullAddress.Length - 1)
            throw new Exception($"Некорректный адрес ячейки: {fullAddress}");
        var sheet = fullAddress.Substring(0, idx);
        var addr = fullAddress.Substring(idx + 1);
        return (sheet, addr);
    }
}
