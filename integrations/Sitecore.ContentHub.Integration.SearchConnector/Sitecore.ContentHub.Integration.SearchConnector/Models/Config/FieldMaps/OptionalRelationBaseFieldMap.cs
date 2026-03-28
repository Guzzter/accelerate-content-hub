using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;

namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Config.FieldMaps
{
    public class OptionalRelationBaseFieldMap : FieldMap
    {
        public IEnumerable<Relation>? ContentHubRelations { get; set; }
    }
}
