namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Config
{
    public class DefinitionMap
    {
        public required string ContentHubEntityDefinition { get; set; }

        public required string SearchEntity { get; set; }

        public required IEnumerable<FieldMap> FieldMaps { get; set; }
    }
}
