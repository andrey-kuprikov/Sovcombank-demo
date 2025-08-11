using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ExcelJobRunner.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJobRunner;

public static class UpdateLinksJob
{
    public static UpdateLinksResult Run(UpdateLinksParams p)
    {
        if (p.InputFile == null || p.OutputFile == null || p.Cells == null)
        {
            throw new ArgumentException("Invalid parameters");
        }
        if (!File.Exists(p.InputFile))
        {
            throw new FileNotFoundException("Input file not found", p.InputFile);
        }

        Excel.Application? app = null;
        Excel.Workbook? wb = null;
        var updated = new List<UpdatedCellResult>();
        try
        {
            app = new Excel.Application { DisplayAlerts = false, Visible = false };
            wb = app.Workbooks.Open(p.InputFile);

            var operations = new List<(Excel.Range range, string newValue, string address, string oldValue)>();

            foreach (var cell in p.Cells)
            {
                if (string.IsNullOrWhiteSpace(cell.Address))
                {
                    throw new Exception("Cell address is empty");
                }
                var split = cell.Address.Split('!');
                if (split.Length != 2)
                {
                    throw new Exception($"Invalid cell address: {cell.Address}");
                }
                var sheetName = split[0];
                var cellRef = split[1];

                Excel.Worksheet? ws;
                try
                {
                    ws = (Excel.Worksheet)wb.Worksheets[sheetName];
                }
                catch
                {
                    throw new Exception($"Sheet not found: {sheetName}");
                }

                Excel.Range range;
                try
                {
                    range = ws.Range[cellRef];
                }
                catch
                {
                    throw new Exception($"Invalid cell address: {cell.Address}");
                }

                var oldVal = Convert.ToString(range.Value2) ?? string.Empty;
                operations.Add((range, cell.NewValue ?? string.Empty, cell.Address, oldVal));
            }

            foreach (var op in operations)
            {
                op.range.Value2 = op.newValue;
                updated.Add(new UpdatedCellResult(op.address, op.oldValue, op.newValue));
            }

            var dir = Path.GetDirectoryName(p.OutputFile);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (string.Equals(p.InputFile, p.OutputFile, StringComparison.OrdinalIgnoreCase))
            {
                wb.Save();
            }
            else
            {
                wb.SaveCopyAs(p.OutputFile);
            }

            return new UpdateLinksResult("OK", updated);
        }
        finally
        {
            if (wb != null)
            {
                wb.Close(false);
                Marshal.ReleaseComObject(wb);
            }
            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}

public record UpdateLinksResult(string Status, List<UpdatedCellResult> UpdatedCells);

public record UpdatedCellResult(string Address, string OldValue, string NewValue);
