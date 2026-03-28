using System.Text.Json.Serialization;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Models.ContentHub
{
    public class UpdateMessage
    {
        [JsonPropertyName("invocation_id")]
        public required string InvocationId { get; set; }

        public required IEnumerable<EntityUpdate> Updates { get; set; }

        public bool Continues { get; set; }
    }
}
