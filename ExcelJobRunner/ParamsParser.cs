using System;
using System.Text.Json;
using ExcelJobRunner.Models;

namespace ExcelJobRunner;

public static class ParamsParser
{
    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web)
    {
        PropertyNameCaseInsensitive = true
    };

    public static object Parse(string action, string json)
    {
        return action switch
        {
            "updateLinks" => JsonSerializer.Deserialize<UpdateLinksParams>(json, Options)
                               ?? throw new JsonException("Invalid updateLinks params"),
            "copyColumns" => JsonSerializer.Deserialize<CopyColumnsParams>(json, Options)
                               ?? throw new JsonException("Invalid copyColumns params"),
            "recalculate" => JsonSerializer.Deserialize<RecalculateParams>(json, Options)
                               ?? throw new JsonException("Invalid recalculate params"),
            "findErrors" => JsonSerializer.Deserialize<FindErrorsParams>(json, Options)
                               ?? throw new JsonException("Invalid findErrors params"),
            _ => throw new ArgumentException($"Unknown action: {action}")
        };
    }
}
