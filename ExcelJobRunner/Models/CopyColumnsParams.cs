using System.Collections.Generic;

namespace ExcelJobRunner.Models;

public class CopyColumnsParams
{
    public string? InputFile { get; set; }
    public List<ColumnCopy>? Columns { get; set; }
    public string? OutputFile { get; set; }
}

public class ColumnCopy
{
    public string? From { get; set; }
    public string? To { get; set; }
}
