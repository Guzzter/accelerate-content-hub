using System.Text.Json.Serialization;

namespace Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub
{
    public class Relation
    {
        public required string Name { get; set; }

        public RelationRole Role { get; set; } = RelationRole.Parent;
    }

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public enum RelationRole
    {
        Parent,
        Child
    }
}
