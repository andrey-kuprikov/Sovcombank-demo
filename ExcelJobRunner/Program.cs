using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

class Program
{
    static int Main(string[] args)
    {
        var arguments = ParseArgs(args);
        if (!arguments.TryGetValue("action", out var action))
        {
            Console.Error.WriteLine("Missing action parameter");
            return 1;
        }
        if (!arguments.TryGetValue("params", out var paramsPath))
        {
            Console.Error.WriteLine("Missing params parameter");
            return 1;
        }
        var resultPath = arguments.ContainsKey("result") ? arguments["result"] : "result.json";

        try
        {
            if (string.Equals(action, "updateLinks", StringComparison.OrdinalIgnoreCase))
            {
                var json = File.ReadAllText(paramsPath);
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var parameters = JsonSerializer.Deserialize<UpdateLinksParams>(json, options) ?? throw new Exception("Params parsing failed");
                var result = UpdateLinksOperation.Run(parameters);
                File.WriteAllText(resultPath, JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
                return result.Status == "OK" ? 0 : 1;
            }
            else
            {
                var result = new UpdateLinksResult { Status = "Fail", Message = $"Unknown action: {action}" };
                File.WriteAllText(resultPath, JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
                return 1;
            }
        }
        catch (Exception ex)
        {
            var result = new UpdateLinksResult { Status = "Fail", Message = ex.Message };
            File.WriteAllText(resultPath, JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }

    static Dictionary<string, string> ParseArgs(string[] args)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var arg in args)
        {
            var parts = arg.Split('=', 2);
            if (parts.Length == 2)
                dict[parts[0]] = parts[1];
        }
        return dict;
    }
}
