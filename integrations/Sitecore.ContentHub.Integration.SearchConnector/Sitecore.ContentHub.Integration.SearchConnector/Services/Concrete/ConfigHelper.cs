using Sitecore.ContentHub.Integration.SearchConnector.Constants;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Config;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using System.Data;
using System.Text.Json;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    public class ConfigHelper : IConfigHelper
    {
        private readonly MappingConfig _mappingConfig  = JsonSerializer.Deserialize<MappingConfig>(File.ReadAllText(ApplicationConstants.MappingConfigFileName)) ?? throw new DataException("Could not parse config.json");

        public IEnumerable<CultureMap> CultureMaps => _mappingConfig.CultureMaps;

        public IEnumerable<DefinitionMap> DefinitionMaps => _mappingConfig.DefinitionMaps;

        public DefinitionMap GetDefinitionMap(string contentHubEntityDefinition)
        {
            return DefinitionMaps.SingleOrDefault(m => m.ContentHubEntityDefinition == contentHubEntityDefinition) ?? throw new InvalidDataException($"Definition map for {contentHubEntityDefinition} not found.");
        }
    }
}
