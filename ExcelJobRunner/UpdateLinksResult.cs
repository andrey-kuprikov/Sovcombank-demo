using System.Collections.Generic;
using System.Text.Json.Serialization;

public class UpdateLinksResult
{
    [JsonPropertyName("status")]
    public string Status { get; set; } = string.Empty;

    [JsonPropertyName("updatedCells")]
    public List<UpdatedCell> UpdatedCells { get; set; } = new();

    [JsonPropertyName("message")]
    public string? Message { get; set; }
}

public class UpdatedCell
{
    [JsonPropertyName("address")]
    public string Address { get; set; } = string.Empty;

    [JsonPropertyName("oldValue")]
    public string OldValue { get; set; } = string.Empty;

    [JsonPropertyName("newValue")]
    public string NewValue { get; set; } = string.Empty;
}
