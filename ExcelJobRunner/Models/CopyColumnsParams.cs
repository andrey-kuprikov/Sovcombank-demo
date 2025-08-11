using System.Collections.Generic;

namespace ExcelJobRunner.Models;

public class CopyColumnsParams
{
    public string? SourceFile { get; set; }
    public string? TargetFile { get; set; }
    public List<ColumnMapping>? Mappings { get; set; }
}

public class ColumnMapping
{
    public ColumnRef? Source { get; set; }
    public ColumnTarget? Target { get; set; }
    public List<string>? FillFormulaColumns { get; set; }
}

public class ColumnRef
{
    public string? Sheet { get; set; }
    public string? Column { get; set; }
    public int StartRow { get; set; }
}

public class ColumnTarget
{
    public string? Sheet { get; set; }
    public string? Column { get; set; }
    public int? StartRow { get; set; }
    public string? Mode { get; set; }
}
