namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Config.FieldMaps
{
    public class PublicLinkFieldMap : OptionalRelationBaseFieldMap
    {
        public required string ContentHubResourceName { get; set; }

        public bool CreateLinkIfNotExists { get; set; } = false;
    }
}
