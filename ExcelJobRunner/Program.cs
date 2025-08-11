using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelJobRunner.Models;

namespace ExcelJobRunner;

public class Program
{
    private static readonly HashSet<string> AllowedActions = new(new[] { "updateLinks", "copyColumns", "recalculate", "findErrors", "runMacro" });

    public static int Main(string[] args)
    {
        if (!TryParseArgs(args, out var action, out var paramsPath, out var resultPath, out var error))
        {
            Console.Write("Invalid arguments");
            WriteResult(resultPath, new JobResult("Fail", error ?? "Invalid arguments"));
            return 1;
        }

        object result;
        try
        {
            var json = File.ReadAllText(paramsPath);
            switch (action)
            {
                case "updateLinks":
                    result = UpdateLinksJob.Run((UpdateLinksParams)ParamsParser.Parse(action, json));
                    break;
                case "copyColumns":
                    result = CopyColumnsJob.Run((CopyColumnsParams)ParamsParser.Parse(action, json));
                    break;
                case "recalculate":
                    result = RecalculateJob.Run((RecalculateParams)ParamsParser.Parse(action, json));
                    break;
                case "runMacro":
                    result = RunMacroJob.Run((RunMacroParams)ParamsParser.Parse(action, json));
                    break;
                default:
                    _ = ParamsParser.Parse(action, json);
                    result = new JobResult("OK", $"{action} parsed");
                    break;
            }
        }
        catch (Exception ex)
        {
            result = new JobResult("Fail", ex.Message);
        }

        WriteResult(resultPath, result);
        var status = GetStatus(result);
        return string.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) ? 0 : 1;
    }

    private static bool TryParseArgs(string[] args, out string action, out string paramsPath, out string resultPath, out string? error)
    {
        action = paramsPath = resultPath = string.Empty;
        error = null;

        var dict = args.Select(a => a.Split('=', 2))
                       .Where(kv => kv.Length == 2)
                       .ToDictionary(kv => kv[0], kv => kv[1], StringComparer.OrdinalIgnoreCase);

        if (!dict.TryGetValue("action", out action))
        {
            error = "Missing required argument 'action'";
            return false;
        }
        if (!AllowedActions.Contains(action))
        {
            error = $"Unknown action '{action}'";
            return false;
        }
        if (!dict.TryGetValue("params", out paramsPath))
        {
            error = "Missing required argument 'params'";
            return false;
        }
        if (!dict.TryGetValue("result", out resultPath))
        {
            error = "Missing required argument 'result'";
            return false;
        }
        if (!File.Exists(paramsPath))
        {
            error = "params.json not found";
            return false;
        }
        return true;
    }

    private static void WriteResult(string path, object result)
    {
        try
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }
            var options = new JsonSerializerOptions(JsonSerializerDefaults.Web)
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };
            var json = JsonSerializer.Serialize(result, options);
            File.WriteAllText(path, json);
        }
        catch
        {
            // suppress exceptions when writing result
        }
    }

    private static string GetStatus(object result) =>
        result.GetType().GetProperty("Status")?.GetValue(result)?.ToString() ?? "Fail";
}
