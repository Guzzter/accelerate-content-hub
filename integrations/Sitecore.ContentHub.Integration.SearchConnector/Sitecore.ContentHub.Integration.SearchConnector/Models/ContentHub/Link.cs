using System.Text.Json.Serialization;

namespace Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub
{
    public class Link
    {
        [JsonPropertyName("href")]
        public required string Uri { get; set; }

        public string? Title { get; set; }
    }
}
