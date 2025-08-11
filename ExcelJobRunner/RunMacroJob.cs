using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using ExcelJobRunner.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJobRunner;

public static class RunMacroJob
{
    public static JobResult Run(RunMacroParams p)
    {
        if (p.InputFile == null || p.MacroName == null)
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
            app.GetType().InvokeMember("AutomationSecurity", BindingFlags.SetProperty, null, app, new object[] { 1 });
            wb = app.Workbooks.Open(p.InputFile);

            try
            {
                // app.Run("Module1.MacroDemo");
                app.Run(p.MacroName);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error executing macro '{p.MacroName}': {ex.Message}", ex);
            }

            var dir = Path.GetDirectoryName(p.OutputFile);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (p.OutputFile == null || string.Equals(p.InputFile, p.OutputFile, StringComparison.OrdinalIgnoreCase))
            {
                wb.Save();
            }
            else
            {
                wb.SaveCopyAs(p.OutputFile);
            }


            var message = $"Macro '{p.MacroName}' executed and file saved as {Path.GetFileName(p.OutputFile)}";
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
