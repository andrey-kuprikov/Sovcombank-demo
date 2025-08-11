using System;
using System.IO;
using System.Runtime.InteropServices;
using ExcelJobRunner.Models;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelJobRunner;

public static class RecalculateJob
{
    public static JobResult Run(RecalculateParams p)
    {
        if (p.InputFile == null || p.OutputFile == null)
        {
            throw new ArgumentException("Invalid parameters");
        }
        if (!File.Exists(p.InputFile))
        {
            throw new FileNotFoundException("Input file not found", p.InputFile);
        }

        Excel.Application? app = null;
        Excel.Workbook? wb = null;
        try
        {
            app = new Excel.Application { DisplayAlerts = false, Visible = false };
            app.GetType().InvokeMember("AutomationSecurity", BindingFlags.SetProperty, null, app, new object[] { 3 });
            wb = app.Workbooks.Open(p.InputFile);

            wb.ForceFullCalculation = true;
            app.CalculateFullRebuild();

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

            var message = $"Recalculation completed. File saved as {Path.GetFileName(p.OutputFile)}";
            return new JobResult("OK", message);
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
