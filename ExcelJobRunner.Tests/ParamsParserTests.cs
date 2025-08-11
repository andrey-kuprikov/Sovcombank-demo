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
}
