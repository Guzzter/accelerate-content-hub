using Stylelabs.M.Sdk.WebClient;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    interface IContentHubClientFactory
    {
        IWebMClient CreateClient();
    }
}
