namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Config
{
    public class MappingConfig
    {
        public required IEnumerable<CultureMap> CultureMaps { get; set; }

        public required IEnumerable<DefinitionMap> DefinitionMaps { get; set; }
    }
}
