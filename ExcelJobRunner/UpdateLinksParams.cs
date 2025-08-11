using System.Collections.Generic;
using System.Text.Json.Serialization;

public class UpdateLinksParams
{
    [JsonPropertyName("inputFile")]
    public string InputFile { get; set; } = string.Empty;

    [JsonPropertyName("cells")]
    public List<CellUpdate> Cells { get; set; } = new();

    [JsonPropertyName("outputFile")]
    public string OutputFile { get; set; } = string.Empty;
}

public class CellUpdate
{
    [JsonPropertyName("address")]
    public string Address { get; set; } = string.Empty;

    [JsonPropertyName("newValue")]
    public string NewValue { get; set; } = string.Empty;
}
