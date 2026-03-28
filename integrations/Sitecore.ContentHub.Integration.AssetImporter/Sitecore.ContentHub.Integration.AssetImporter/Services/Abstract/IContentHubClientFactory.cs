using Stylelabs.M.Sdk.WebClient;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract
{
    interface IContentHubClientFactory
    {
        IWebMClient CreateClient();
    }
}
