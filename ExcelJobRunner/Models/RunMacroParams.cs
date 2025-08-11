namespace ExcelJobRunner.Models;

public class RunMacroParams
{
    public string? InputFile { get; set; }
    public string? MacroName { get; set; }
    public string? OutputFile { get; set; }
}
