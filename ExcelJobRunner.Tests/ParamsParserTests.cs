using System.IO;
using ExcelJobRunner;
using ExcelJobRunner.Models;
using Xunit;

namespace ExcelJobRunner.Tests;

public class ParamsParserTests
{
    [Fact]
    public void UpdateLinksParams_CanBeReadFromJson()
    {
        var json = """
{
  "inputFile": "file.xlsx",
  "cells": [{ "address": "Sheet1!A1", "newValue": "test" }],
  "outputFile": "out.xlsx"
}
""";
        var path = Path.GetTempFileName();
        File.WriteAllText(path, json);

        var text = File.ReadAllText(path);
        var result = ParamsParser.Parse("updateLinks", text);

        var parsed = Assert.IsType<UpdateLinksParams>(result);
        Assert.Equal("file.xlsx", parsed.InputFile);
        Assert.Single(parsed.Cells!);
        Assert.Equal("Sheet1!A1", parsed.Cells![0].Address);
    }

    [Fact]
    public void CopyColumnsParams_CanBeReadFromJson()
    {
        var json = """
{
  "sourceFile": "C:\\files\\donor.xlsx",
  "targetFile": "C:\\files\\acceptor.xlsx",
  "mappings": [
    {
      "source": { "sheet": "Sheet1", "column": "E", "startRow": 2 },
      "target": { "sheet": "Sheet1", "column": "K", "mode": "append" },
      "fillFormulaColumns": [ "L", "M" ]
    },
    {
      "source": { "sheet": "Данные", "column": "B", "startRow": 5 },
      "target": { "sheet": "Отчет", "column": "C", "startRow": 12 }
    }
  ]
}
""";

        var result = ParamsParser.Parse("copyColumns", json);

        var parsed = Assert.IsType<CopyColumnsParams>(result);
        Assert.Equal("C:\\files\\donor.xlsx", parsed.SourceFile);
        Assert.Equal("C:\\files\\acceptor.xlsx", parsed.TargetFile);
        Assert.Equal(2, parsed.Mappings!.Count);
        Assert.Equal("Sheet1", parsed.Mappings![0].Source!.Sheet);
        Assert.Equal("M", parsed.Mappings![0].FillFormulaColumns![1]);
    }

    [Fact]
    public void RecalculateParams_CanBeReadFromJson()
    {
        var json = """
{
  "inputFile": "file.xlsb",
  "outputFile": "out.xlsb",
  "sheets": ["Sheet1", "Sheet2"]
}
""";

        var result = ParamsParser.Parse("recalculate", json);

        var parsed = Assert.IsType<RecalculateParams>(result);
        Assert.Equal("file.xlsb", parsed.InputFile);
        Assert.Equal("out.xlsb", parsed.OutputFile);
        Assert.Equal(2, parsed.Sheets!.Count);
        Assert.Equal("Sheet2", parsed.Sheets![1]);
    }

    [Fact]
    public void RunMacroParams_CanBeReadFromJson()
    {
        var json = """
{
  "inputFile": "C:\\files\\macrosource.xlsm",
  "macroName": "Module1.MySpecialMacro",
  "outputFile": "C:\\files\\macrosource_after_macro.xlsm"
}
""";

        var result = ParamsParser.Parse("runMacro", json);

        var parsed = Assert.IsType<RunMacroParams>(result);
        Assert.Equal("C:\\files\\macrosource.xlsm", parsed.InputFile);
        Assert.Equal("Module1.MySpecialMacro", parsed.MacroName);
        Assert.Equal("C:\\files\\macrosource_after_macro.xlsm", parsed.OutputFile);
    }
}
