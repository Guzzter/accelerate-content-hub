using Microsoft.Extensions.Options;
using OrderCloud.SDK;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Models.Options;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Concrete
{
    class OrderCloudClientFactory(IOptions<OrderCloudOptions> orderCloudOptions) : IOrderCloudClientFactory
    {
        private OrderCloudClient? client;

        public OrderCloudClient CreateClient()
        {
            if(client == null)
            {
                var config = new OrderCloudClientConfig
                {
                    ApiUrl = orderCloudOptions.Value.ApiUrl,
                    AuthUrl = orderCloudOptions.Value.AuthUrl ?? orderCloudOptions.Value.ApiUrl,
                    ClientId = orderCloudOptions.Value.ClientId,
                    ClientSecret = orderCloudOptions.Value.ClientSecret,
                    Roles = [ApiRole.FullAccess], // todo: review this
                };

                client = new OrderCloudClient(config);
            }
            return client;
        }
    }
}
