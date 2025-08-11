using System.Collections.Generic;

namespace ExcelJobRunner.Models;

public class RecalculateParams
{
    public string? InputFile { get; set; }
    public List<string>? Sheets { get; set; }
    public string? OutputFile { get; set; }
}
