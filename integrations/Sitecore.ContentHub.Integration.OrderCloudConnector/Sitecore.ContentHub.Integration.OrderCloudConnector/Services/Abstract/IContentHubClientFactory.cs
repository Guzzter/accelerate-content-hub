using Stylelabs.M.Sdk.WebClient;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract
{
    interface IContentHubClientFactory
    {
        IWebMClient CreateClient();
    }
}
