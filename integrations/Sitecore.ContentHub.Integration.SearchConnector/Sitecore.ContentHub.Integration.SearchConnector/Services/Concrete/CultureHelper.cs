using Sitecore.ContentHub.Integration.SearchConnector.Models.Config;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using System.Globalization;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    class CultureHelper(IConfigHelper configHelper, IContentHubClientHelper contentHubClientHelper) : ICultureHelper
    {
        public async Task<CultureInfo> GetContentHubDefaultCulture()
        {
            return await contentHubClientHelper.Execute(c => c.Cultures.GetDefaultCultureCachedAsync());
        }

        public IEnumerable<CultureInfo> GetMappedContentHubCultures()
        {
            return configHelper.CultureMaps.Select(m => new CultureInfo(m.ContentHubCulture));
        }

        public IEnumerable<CultureMap> GetCultureMaps(IEnumerable<CultureInfo> contentHubCultures)
        {
            var contentHubCultureStrings = contentHubCultures.Select(c => c.Name).ToList();
            return configHelper.CultureMaps.Where(m => contentHubCultureStrings.Contains(m.ContentHubCulture));
        }
    }
}
