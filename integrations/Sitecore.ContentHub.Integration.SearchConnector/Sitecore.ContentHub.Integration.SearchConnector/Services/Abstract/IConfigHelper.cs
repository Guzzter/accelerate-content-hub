using Sitecore.ContentHub.Integration.SearchConnector.Models.Config;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface IConfigHelper
    {
        public IEnumerable<CultureMap> CultureMaps { get; }
        public IEnumerable<DefinitionMap> DefinitionMaps { get; }

        public DefinitionMap GetDefinitionMap(string contentHubEntityDefinition);
    }
}
