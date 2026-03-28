using OrderCloud.SDK;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract
{
    interface IOrderCloudClientFactory
    {
        OrderCloudClient CreateClient();
    }
}
