using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ExcelJobRunner.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJobRunner;

public static class CopyColumnsJob
{
    public static CopyColumnsResult Run(CopyColumnsParams p)
    {
        if (p.SourceFile == null || p.TargetFile == null || p.Mappings == null)
        {
            throw new ArgumentException("Invalid parameters");
        }
        if (!File.Exists(p.SourceFile))
        {
            throw new FileNotFoundException("Source file not found", p.SourceFile);
        }
        if (!File.Exists(p.TargetFile))
        {
            throw new FileNotFoundException("Target file not found", p.TargetFile);
        }

        Excel.Application? app = null;
        Excel.Workbook? sourceWb = null;
        Excel.Workbook? targetWb = null;
        var copied = new List<CopiedColumnResult>();
        var fills = new List<FillFormulaResult>();

        try
        {
            app = new Excel.Application { DisplayAlerts = false, Visible = false };
            sourceWb = app.Workbooks.Open(p.SourceFile);
            targetWb = string.Equals(p.SourceFile, p.TargetFile, StringComparison.OrdinalIgnoreCase)
                ? sourceWb
                : app.Workbooks.Open(p.TargetFile);

            var ops = new List<OperationInfo>();
            foreach (var m in p.Mappings)
            {
                if (m.Source == null || m.Target == null)
                {
                    throw new Exception("Mapping must have source and target");
                }
                if (string.IsNullOrWhiteSpace(m.Source.Sheet) || string.IsNullOrWhiteSpace(m.Source.Column))
                {
                    throw new Exception("Source sheet or column is empty");
                }
                if (string.IsNullOrWhiteSpace(m.Target.Sheet) || string.IsNullOrWhiteSpace(m.Target.Column))
                {
                    throw new Exception("Target sheet or column is empty");
                }

                Excel.Worksheet wsSrc;
                Excel.Worksheet wsTgt;
                try
                {
                    wsSrc = (Excel.Worksheet)sourceWb.Worksheets[m.Source.Sheet];
                }
                catch
                {
                    throw new Exception($"Sheet {m.Source.Sheet} not found in source file");
                }
                try
                {
                    wsTgt = (Excel.Worksheet)targetWb.Worksheets[m.Target.Sheet];
                }
                catch
                {
                    throw new Exception($"Sheet {m.Target.Sheet} not found in target file");
                }

                var srcCol = m.Source.Column;
                int srcStartRow = m.Source.StartRow;
                int srcLastRow;
                try
                {
                    srcLastRow = ((Excel.Range)wsSrc.Cells[wsSrc.Rows.Count, srcCol])
                        .End[Excel.XlDirection.xlUp].Row;
                }
                catch
                {
                    throw new Exception($"Invalid source column {srcCol}");
                }
                if (srcLastRow < srcStartRow)
                {
                    throw new Exception($"Source range {m.Source.Sheet}!{srcCol}{srcStartRow} is empty");
                }
                int rowsCount = srcLastRow - srcStartRow + 1;

                var tgtCol = m.Target.Column;
                int tgtStartRow;
                if (string.Equals(m.Target.Mode, "append", StringComparison.OrdinalIgnoreCase))
                {
                    tgtStartRow = ((Excel.Range)wsTgt.Cells[wsTgt.Rows.Count, tgtCol])
                        .End[Excel.XlDirection.xlUp].Row + 1;
                }
                else
                {
                    tgtStartRow = m.Target.StartRow ?? srcStartRow;
                }
                int tgtEndRow = tgtStartRow + rowsCount - 1;

                string srcAddress = $"{srcCol}{srcStartRow}:{srcCol}{srcLastRow}";
                string tgtAddress = $"{tgtCol}{tgtStartRow}:{tgtCol}{tgtEndRow}";

                ops.Add(new OperationInfo(m.Source.Sheet!, srcAddress, m.Target.Sheet!, tgtAddress, rowsCount, tgtStartRow, tgtEndRow, m.FillFormulaColumns));

                Marshal.ReleaseComObject(wsSrc);
                Marshal.ReleaseComObject(wsTgt);
            }

            foreach (var op in ops)
            {
                var wsSrc = (Excel.Worksheet)sourceWb.Worksheets[op.SourceSheet];
                var wsTgt = (Excel.Worksheet)targetWb.Worksheets[op.TargetSheet];
                var srcRange = wsSrc.Range[op.SourceAddress];
                var tgtRange = wsTgt.Range[op.TargetAddress];
                srcRange.Copy(tgtRange);
                copied.Add(new CopiedColumnResult($"{op.SourceSheet}!{op.SourceAddress}", $"{op.TargetSheet}!{op.TargetAddress}", op.Count));

                if (op.FillColumns != null)
                {
                    foreach (var col in op.FillColumns)
                    {
                        var cell = wsTgt.Range[$"{col}{op.TargetStartRow}"];
                        var formula = Convert.ToString(cell.Formula);
                        if (!string.IsNullOrEmpty(formula))
                        {
                            var fillRange = wsTgt.Range[$"{col}{op.TargetStartRow}:{col}{op.TargetEndRow}"];
                            cell.AutoFill(fillRange, Excel.XlAutoFillType.xlFillDefault);
                            fills.Add(new FillFormulaResult(wsTgt.Name, col, op.TargetStartRow, op.TargetEndRow, $"{col}{op.TargetStartRow}"));
                            Marshal.ReleaseComObject(fillRange);
                        }
                        Marshal.ReleaseComObject(cell);
                    }
                }

                Marshal.ReleaseComObject(srcRange);
                Marshal.ReleaseComObject(tgtRange);
                Marshal.ReleaseComObject(wsSrc);
                Marshal.ReleaseComObject(wsTgt);
            }

            targetWb.Save();
            app.CalculateFull();

            return new CopyColumnsResult("OK", copied, fills.Count > 0 ? fills : null);
        }
        finally
        {
            if (targetWb != null && targetWb != sourceWb)
            {
                targetWb.Close(false);
                Marshal.ReleaseComObject(targetWb);
            }
            if (sourceWb != null)
            {
                sourceWb.Close(false);
                Marshal.ReleaseComObject(sourceWb);
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

    private record OperationInfo(string SourceSheet, string SourceAddress, string TargetSheet, string TargetAddress, int Count, int TargetStartRow, int TargetEndRow, List<string>? FillColumns);
}

public record CopyColumnsResult(string Status, List<CopiedColumnResult> CopiedColumns, List<FillFormulaResult>? FillFormulas);

public record CopiedColumnResult(string Source, string Target, int Count);

public record FillFormulaResult(string Sheet, string Column, int FromRow, int ToRow, string FormulaSource);
