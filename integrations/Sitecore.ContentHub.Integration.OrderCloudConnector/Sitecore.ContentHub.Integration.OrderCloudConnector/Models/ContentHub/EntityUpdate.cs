using System.Text.Json.Serialization;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Models.ContentHub
{
    public class EntityUpdate
    {
        public required string Identifier { get; set; }

        [JsonPropertyName("entity_definition")]
        public string? EntityDefinition { get; set; }

        [JsonConverter(typeof(JsonStringEnumConverter))]
        public required EntityUpdateOperation Operation { get; set; }
    }

    public enum EntityUpdateOperation
    {
        Update,
        Delete
    }
}
