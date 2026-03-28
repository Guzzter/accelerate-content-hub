using Sitecore.ContentHub.Integration.SearchConnector.Models.Config;
using System.Globalization;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface ICultureHelper
    {
        Task<CultureInfo> GetContentHubDefaultCulture();

        IEnumerable<CultureInfo> GetMappedContentHubCultures();

        IEnumerable<CultureMap> GetCultureMaps(IEnumerable<CultureInfo> contentHubCultures);
    }
}
