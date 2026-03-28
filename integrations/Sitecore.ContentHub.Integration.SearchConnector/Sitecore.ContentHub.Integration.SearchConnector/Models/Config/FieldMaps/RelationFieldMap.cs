using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;

namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Config.FieldMaps
{
    public class RelationFieldMap : FieldMap
    {
        public required IEnumerable<Relation> ContentHubRelations { get; set; }

        public required string ContentHubRelatedPropertyName { get; set; }
    }
}
