using System.Collections.Generic;

namespace ExcelJobRunner.Models;

public class UpdateLinksParams
{
    public string? InputFile { get; set; }
    public List<CellUpdate>? Cells { get; set; }
    public string? OutputFile { get; set; }
}

public class CellUpdate
{
    public string? Address { get; set; }
    public string? NewValue { get; set; }
}
